import SwiftUI

struct ContentView: View {
    @EnvironmentObject var engine: EngineService
    @State private var showSettings = false

    var body: some View {
        NavigationStack {
            ScrollView {
                VStack(spacing: 24) {
                    headerSection
                    filePickerSection
                    instructionSection
                    if engine.isProcessing {
                        progressSection
                    }
                    if let result = engine.result {
                        resultSection(result)
                    }
                    runButton
                }
                .padding()
            }
            .navigationTitle("Excel Engine")
            .toolbar {
                ToolbarItem(placement: .topBarTrailing) {
                    Button { showSettings = true } label: {
                        Image(systemName: "gear")
                    }
                }
            }
            .sheet(isPresented: $showSettings) {
                SettingsView()
                    .environmentObject(engine)
            }
        }
    }

    private var headerSection: some View {
        VStack(spacing: 8) {
            Image(systemName: "tablecells.badge.ellipsis")
                .font(.system(size: 48))
                .foregroundStyle(.blue)
            Text("Autonomous Excel Automation")
                .font(.subheadline)
                .foregroundStyle(.secondary)
        }
        .padding(.top)
    }

    private var filePickerSection: some View {
        VStack(alignment: .leading, spacing: 12) {
            Label("Workbook", systemImage: "doc.badge.plus")
                .font(.headline)

            FilePickerButton(
                title: engine.workbookURL?.lastPathComponent ?? "Select .xlsx file",
                types: [.spreadsheet, .data],
                onPick: { engine.workbookURL = $0 }
            )

            if !engine.resourceURLs.isEmpty {
                ForEach(engine.resourceURLs, id: \.self) { url in
                    HStack {
                        Image(systemName: "paperclip")
                        Text(url.lastPathComponent)
                            .lineLimit(1)
                        Spacer()
                        Button { engine.resourceURLs.removeAll { $0 == url } } label: {
                            Image(systemName: "xmark.circle.fill")
                                .foregroundStyle(.secondary)
                        }
                    }
                    .font(.caption)
                }
            }

            FilePickerButton(
                title: "Add Resource Files",
                types: [.spreadsheet, .pdf, .plainText, .data, .image],
                onPick: { engine.resourceURLs.append($0) },
                isSecondary: true
            )
        }
        .padding()
        .background(.ultraThinMaterial, in: RoundedRectangle(cornerRadius: 12))
    }

    private var instructionSection: some View {
        VStack(alignment: .leading, spacing: 12) {
            Label("Instructions", systemImage: "text.document")
                .font(.headline)

            FilePickerButton(
                title: engine.instructionURL?.lastPathComponent ?? "Select instructions file",
                types: [.pdf, .plainText, .data],
                onPick: { engine.instructionURL = $0 }
            )

            Text("Or type instructions directly:")
                .font(.caption)
                .foregroundStyle(.secondary)

            TextEditor(text: $engine.instructionText)
                .frame(minHeight: 80)
                .overlay(
                    RoundedRectangle(cornerRadius: 8)
                        .stroke(.quaternary, lineWidth: 1)
                )
        }
        .padding()
        .background(.ultraThinMaterial, in: RoundedRectangle(cornerRadius: 12))
    }

    private var progressSection: some View {
        VStack(spacing: 12) {
            ProgressView(value: engine.progress)
                .tint(.blue)
            Text(engine.statusMessage)
                .font(.caption)
                .foregroundStyle(.secondary)
        }
        .padding()
        .background(.ultraThinMaterial, in: RoundedRectangle(cornerRadius: 12))
    }

    private func resultSection(_ result: EngineResult) -> some View {
        VStack(alignment: .leading, spacing: 12) {
            HStack {
                Image(systemName: result.success ? "checkmark.circle.fill" : "xmark.circle.fill")
                    .foregroundStyle(result.success ? .green : .red)
                Text(result.success ? "Assignment Complete" : "Completed with Issues")
                    .font(.headline)
            }

            Text("Tasks: \(result.tasksCompleted)/\(result.tasksTotal)")
                .font(.subheadline)

            if !result.summary.isEmpty {
                Text(result.summary)
                    .font(.caption)
                    .foregroundStyle(.secondary)
            }

            if let outputURL = engine.outputURL {
                ShareLink(item: outputURL) {
                    Label("Share Result", systemImage: "square.and.arrow.up")
                }
                .buttonStyle(.bordered)
            }
        }
        .padding()
        .background(
            (result.success ? Color.green : Color.red).opacity(0.1),
            in: RoundedRectangle(cornerRadius: 12)
        )
    }

    private var runButton: some View {
        Button {
            Task { await engine.run() }
        } label: {
            Label("Run Engine", systemImage: "play.fill")
                .frame(maxWidth: .infinity)
                .padding(.vertical, 4)
        }
        .buttonStyle(.borderedProminent)
        .disabled(!engine.canRun)
        .padding(.bottom)
    }
}
