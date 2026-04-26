import Vision
import Foundation
import AppKit
import CoreGraphics

guard CommandLine.arguments.count >= 3 else {
    FileHandle.standardError.write("Usage: ocr_pdf <input.pdf> <output.txt>\n".data(using: .utf8)!)
    exit(1)
}

let pdfPath = CommandLine.arguments[1]
let outPath = CommandLine.arguments[2]

let url = URL(fileURLWithPath: pdfPath)
guard let doc = CGPDFDocument(url as CFURL) else {
    FileHandle.standardError.write("Failed to open PDF\n".data(using: .utf8)!)
    exit(1)
}

var allText = ""

for pageNum in 1...doc.numberOfPages {
    guard let page = doc.page(at: pageNum) else { continue }

    let mediaBox = page.getBoxRect(.mediaBox)
    let scale: CGFloat = 3.0
    let width = Int(mediaBox.width * scale)
    let height = Int(mediaBox.height * scale)

    guard let context = CGContext(
        data: nil,
        width: width,
        height: height,
        bitsPerComponent: 8,
        bytesPerRow: 0,
        space: CGColorSpaceCreateDeviceRGB(),
        bitmapInfo: CGImageAlphaInfo.premultipliedLast.rawValue
    ) else { continue }

    context.setFillColor(CGColor.white)
    context.fill(CGRect(x: 0, y: 0, width: width, height: height))
    context.scaleBy(x: scale, y: scale)
    context.drawPDFPage(page)

    guard let cgImg = context.makeImage() else { continue }

    let request = VNRecognizeTextRequest()
    request.recognitionLevel = .accurate
    request.recognitionLanguages = ["ko-KR", "zh-Hant", "en-US"]
    request.usesLanguageCorrection = false

    let handler = VNImageRequestHandler(cgImage: cgImg, options: [:])
    do {
        try handler.perform([request])
        if let observations = request.results {
            for obs in observations {
                if let top = obs.topCandidates(1).first {
                    allText += top.string + "\n"
                }
            }
        }
    } catch {
        FileHandle.standardError.write("Error on page \(pageNum): \(error)\n".data(using: .utf8)!)
    }
    allText += "\n"
}

try allText.write(toFile: outPath, atomically: true, encoding: .utf8)
print("Wrote \(allText.count) chars to \(outPath)")
