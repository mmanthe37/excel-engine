import SwiftUI

struct SettingsView: View {
    @EnvironmentObject var engine: EngineService
    @Environment(\.dismiss) private var dismiss

    var body: some View {
        NavigationStack {
            Form {
                Section("Server") {
                    TextField("Server URL", text: $engine.serverURL)
                        .textContentType(.URL)
                        .autocapitalization(.none)
                        .keyboardType(.URL)

                    SecureField("API Key (optional)", text: $engine.apiKey)
                }

                Section("Engine Options") {
                    Picker("Max Layer", selection: $engine.maxLayer) {
                        Text("Layer 1 — openpyxl").tag(1)
                        Text("Layer 2 — xlwings").tag(2)
                        Text("Layer 3 — AppleScript").tag(3)
                        Text("Layer 4 — System Events").tag(4)
                        Text("Layer 5 — VBA").tag(5)
                        Text("Layer 6 — PyAutoGUI").tag(6)
                    }

                    Toggle("Dry Run", isOn: $engine.dryRun)
                    Toggle("Recalculate Formulas", isOn: $engine.recalcFormulas)
                }

                Section("About") {
                    LabeledContent("Version", value: "1.2.0")
                    LabeledContent("Engine", value: "6-Layer Cascading")
                    Link("GitHub", destination: URL(string: "https://github.com/mmanthe37/excel-engine")!)
                }
            }
            .navigationTitle("Settings")
            .navigationBarTitleDisplayMode(.inline)
            .toolbar {
                ToolbarItem(placement: .topBarTrailing) {
                    Button("Done") { dismiss() }
                }
            }
        }
    }
}
