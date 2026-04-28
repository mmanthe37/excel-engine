import Foundation
import SwiftUI

@MainActor
class EngineService: ObservableObject {
    @Published var workbookURL: URL?
    @Published var instructionURL: URL?
    @Published var instructionText: String = ""
    @Published var resourceURLs: [URL] = []
    @Published var isProcessing = false
    @Published var progress: Double = 0
    @Published var statusMessage: String = ""
    @Published var result: EngineResult?
    @Published var outputURL: URL?

    @AppStorage("serverURL") var serverURL = "http://localhost:8080"
    @AppStorage("apiKey") var apiKey = ""
    @AppStorage("maxLayer") var maxLayer = 6
    @AppStorage("dryRun") var dryRun = false
    @AppStorage("recalcFormulas") var recalcFormulas = false

    var canRun: Bool {
        workbookURL != nil && (instructionURL != nil || !instructionText.isEmpty) && !isProcessing
    }

    func run() async {
        guard canRun else { return }

        isProcessing = true
        progress = 0
        statusMessage = "Uploading files..."
        result = nil
        outputURL = nil

        do {
            progress = 0.1
            statusMessage = "Connecting to server..."

            let workbookData = try Data(contentsOf: workbookURL!)
            let workbookName = workbookURL!.lastPathComponent

            var instructionContent = instructionText
            if let instrURL = instructionURL {
                instructionContent = try String(contentsOf: instrURL, encoding: .utf8)
            }

            progress = 0.3
            statusMessage = "Running engine..."

            let engineResult = try await callEngine(
                workbookName: workbookName,
                workbookData: workbookData,
                instructions: instructionContent
            )

            progress = 1.0
            statusMessage = engineResult.success ? "Complete!" : "Finished with issues"
            result = engineResult

            // Save output to temp
            let tempDir = FileManager.default.temporaryDirectory
            let outputPath = tempDir.appendingPathComponent("result_\(workbookName)")
            try workbookData.write(to: outputPath)
            outputURL = outputPath

        } catch {
            statusMessage = "Error: \(error.localizedDescription)"
            result = EngineResult(success: false, summary: error.localizedDescription, errors: [error.localizedDescription])
        }

        isProcessing = false
    }

    private func callEngine(workbookName: String, workbookData: Data, instructions: String) async throws -> EngineResult {
        let base = serverURL.trimmingCharacters(in: .whitespacesAndNewlines)
        guard let url = URL(string: "\(base)/mcp") else {
            throw URLError(.badURL)
        }

        let request = MCPRequest(
            method: "tools/call",
            params: MCPParams(
                name: "complete_assignment",
                arguments: [
                    "workbook_path": workbookName,
                    "instruction_text": instructions,
                    "max_layer": String(maxLayer),
                    "dry_run": String(dryRun),
                    "recalculate": String(recalcFormulas),
                ]
            )
        )

        var urlRequest = URLRequest(url: url)
        urlRequest.httpMethod = "POST"
        urlRequest.setValue("application/json", forHTTPHeaderField: "Content-Type")
        if !apiKey.isEmpty {
            urlRequest.setValue(apiKey, forHTTPHeaderField: "X-API-Key")
        }
        urlRequest.httpBody = try JSONEncoder().encode(request)

        let (data, response) = try await URLSession.shared.data(for: urlRequest)

        guard let httpResponse = response as? HTTPURLResponse,
              (200...299).contains(httpResponse.statusCode) else {
            let statusCode = (response as? HTTPURLResponse)?.statusCode ?? 0
            throw URLError(.badServerResponse, userInfo: [
                NSLocalizedDescriptionKey: "Server returned status \(statusCode)"
            ])
        }

        let mcpResponse = try JSONDecoder().decode(MCPResponse.self, from: data)

        if let error = mcpResponse.error {
            return EngineResult(success: false, summary: error.message, errors: [error.message])
        }

        if let content = mcpResponse.result?.content?.first?.text {
            if let resultData = content.data(using: .utf8),
               let parsed = try? JSONDecoder().decode(EngineResult.self, from: resultData) {
                return parsed
            }
            return EngineResult(success: true, summary: content)
        }

        return EngineResult(success: true, summary: "Engine completed successfully")
    }
}
