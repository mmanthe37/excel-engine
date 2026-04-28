import Foundation

struct EngineResult: Codable {
    let success: Bool
    let tasksCompleted: Int
    let tasksTotal: Int
    let summary: String
    let errors: [String]

    enum CodingKeys: String, CodingKey {
        case success
        case tasksCompleted = "tasks_completed"
        case tasksTotal = "tasks_total"
        case summary
        case errors
    }

    init(success: Bool = false, tasksCompleted: Int = 0, tasksTotal: Int = 0, summary: String = "", errors: [String] = []) {
        self.success = success
        self.tasksCompleted = tasksCompleted
        self.tasksTotal = tasksTotal
        self.summary = summary
        self.errors = errors
    }
}

struct MCPRequest: Codable {
    let jsonrpc: String
    let id: Int
    let method: String
    let params: MCPParams

    init(method: String, params: MCPParams) {
        self.jsonrpc = "2.0"
        self.id = Int.random(in: 1...999999)
        self.method = method
        self.params = params
    }
}

struct MCPParams: Codable {
    let name: String
    let arguments: [String: String]
}

struct MCPResponse: Codable {
    let jsonrpc: String
    let id: Int?
    let result: MCPResult?
    let error: MCPError?
}

struct MCPResult: Codable {
    let content: [MCPContent]?
}

struct MCPContent: Codable {
    let type: String
    let text: String?
}

struct MCPError: Codable {
    let code: Int
    let message: String
}
