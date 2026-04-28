import SwiftUI
import UniformTypeIdentifiers

struct FilePickerButton: View {
    let title: String
    let types: [UTType]
    let onPick: (URL) -> Void
    var isSecondary: Bool = false

    @State private var showPicker = false

    var body: some View {
        Button { showPicker = true } label: {
            HStack {
                Image(systemName: isSecondary ? "plus.circle" : "folder")
                Text(title)
                    .lineLimit(1)
                Spacer()
            }
            .padding(.horizontal, 12)
            .padding(.vertical, 10)
            .background(isSecondary ? .clear : Color(.systemGray6))
            .cornerRadius(8)
            .overlay(
                RoundedRectangle(cornerRadius: 8)
                    .stroke(isSecondary ? Color.accentColor : .clear, lineWidth: 1)
            )
        }
        .buttonStyle(.plain)
        .foregroundStyle(isSecondary ? .blue : .primary)
        .fileImporter(
            isPresented: $showPicker,
            allowedContentTypes: types,
            allowsMultipleSelection: false
        ) { result in
            if case .success(let urls) = result, let url = urls.first {
                _ = url.startAccessingSecurityScopedResource()
                onPick(url)
            }
        }
    }
}
