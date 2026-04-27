// Excel Engine MCP Connector — Power Platform Custom Connector Script
// C# transformation script for request/response processing.
//
// Responsibilities:
//   1. Ensure JSON-RPC 2.0 compliance on every outbound request
//   2. Inject a request ID when the caller omits one
//   3. Convert SSE (text/event-stream) responses back to plain JSON
//      so Copilot Studio can parse the result normally

public class Script : ScriptBase
{
    public override async Task<HttpResponseMessage> ExecuteAsync()
    {
        if (this.Context.OperationId == "McpEndpoint")
        {
            return await HandleMcpRequestAsync().ConfigureAwait(false);
        }

        // Pass-through for all other operations (e.g. HealthCheck)
        return await this.Context.SendAsync(this.Context.Request, this.CancellationToken)
            .ConfigureAwait(false);
    }

    // ── MCP request handler ──────────────────────────────────────────

    private async Task<HttpResponseMessage> HandleMcpRequestAsync()
    {
        // Read and validate request body
        var rawBody = await this.Context.Request.Content
            .ReadAsStringAsync().ConfigureAwait(false);

        JObject requestBody;
        try
        {
            requestBody = JObject.Parse(rawBody);
        }
        catch (JsonException)
        {
            return BuildErrorResponse(
                null, -32700, "Parse error: request body is not valid JSON",
                System.Net.HttpStatusCode.BadRequest);
        }

        // Enforce JSON-RPC 2.0
        if (requestBody["jsonrpc"] == null || requestBody["jsonrpc"].ToString() != "2.0")
        {
            requestBody["jsonrpc"] = "2.0";
        }

        // Generate an ID when the caller omits one (notifications still work without IDs)
        if (requestBody["id"] == null)
        {
            requestBody["id"] = Guid.NewGuid().ToString("N");
        }

        // Replace request content with the normalised body
        this.Context.Request.Content = CreateJsonContent(requestBody.ToString(Formatting.None));

        // Ask for plain JSON — the script will handle SSE if the server streams
        this.Context.Request.Headers.Accept.Clear();
        this.Context.Request.Headers.Accept.Add(
            new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));

        var response = await this.Context.SendAsync(this.Context.Request, this.CancellationToken)
            .ConfigureAwait(false);

        // If the server replied with SSE, extract the last data frame
        var mediaType = response.Content.Headers.ContentType?.MediaType ?? string.Empty;
        if (string.Equals(mediaType, "text/event-stream", StringComparison.OrdinalIgnoreCase))
        {
            var sseBody = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
            var jsonPayload = ExtractLastSseData(sseBody);
            response.Content = CreateJsonContent(jsonPayload);
            response.Content.Headers.ContentType =
                new System.Net.Http.Headers.MediaTypeHeaderValue("application/json");
        }

        return response;
    }

    // ── SSE helpers ──────────────────────────────────────────────────

    /// <summary>
    /// Parse an SSE stream and return the JSON payload of the last
    /// non-terminal data frame (i.e. not "[DONE]").
    /// </summary>
    private static string ExtractLastSseData(string sseContent)
    {
        string lastData = null;

        foreach (var rawLine in sseContent.Split('\n'))
        {
            var line = rawLine.TrimEnd('\r');
            if (line.StartsWith("data: ", StringComparison.Ordinal))
            {
                var payload = line.Substring(6).Trim();
                if (!string.Equals(payload, "[DONE]", StringComparison.Ordinal))
                {
                    lastData = payload;
                }
            }
        }

        // Return the last seen frame, or a minimal JSON-RPC error if nothing arrived
        return lastData ?? JsonConvert.SerializeObject(new
        {
            jsonrpc = "2.0",
            id = (string)null,
            error = new { code = -32603, message = "Empty SSE stream from MCP server" }
        });
    }

    // ── Error response builder ───────────────────────────────────────

    private static HttpResponseMessage BuildErrorResponse(
        string id, int code, string message,
        System.Net.HttpStatusCode statusCode = System.Net.HttpStatusCode.OK)
    {
        var body = JsonConvert.SerializeObject(new
        {
            jsonrpc = "2.0",
            id,
            error = new { code, message }
        });

        var response = new HttpResponseMessage(statusCode);
        response.Content = new StringContent(
            body,
            System.Text.Encoding.UTF8,
            "application/json");
        return response;
    }
}
