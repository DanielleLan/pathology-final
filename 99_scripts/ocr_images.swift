import Foundation
import Vision
import AppKit

let imagePaths = Array(CommandLine.arguments.dropFirst())

if imagePaths.isEmpty {
    fputs("Usage: swift ocr_images.swift <image>...\n", stderr)
    exit(1)
}

func cgImage(from path: String) -> CGImage? {
    guard let image = NSImage(contentsOfFile: path) else { return nil }
    var rect = CGRect(origin: .zero, size: image.size)
    return image.cgImage(forProposedRect: &rect, context: nil, hints: nil)
}

for path in imagePaths {
    print("\n===== \(URL(fileURLWithPath: path).lastPathComponent) =====")

    guard let image = cgImage(from: path) else {
        print("[OCR failed: cannot load image]")
        continue
    }

    let request = VNRecognizeTextRequest()
    request.recognitionLevel = .accurate
    request.usesLanguageCorrection = true
    request.recognitionLanguages = ["zh-Hans", "en-US"]

    let handler = VNImageRequestHandler(cgImage: image, options: [:])
    do {
        try handler.perform([request])
        let lines = (request.results ?? [])
            .compactMap { $0.topCandidates(1).first?.string.trimmingCharacters(in: .whitespacesAndNewlines) }
            .filter { !$0.isEmpty }
        print(lines.joined(separator: "\n"))
    } catch {
        print("[OCR failed: \(error.localizedDescription)]")
    }
}
