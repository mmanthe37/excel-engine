import SwiftUI

@main
struct ExcelEngineApp: App {
    @StateObject private var engineService = EngineService()

    var body: some Scene {
        WindowGroup {
            ContentView()
                .environmentObject(engineService)
        }
    }
}
