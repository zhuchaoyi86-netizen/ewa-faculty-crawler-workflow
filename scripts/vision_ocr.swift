#!/usr/bin/env swift

import Foundation
import Vision
import ImageIO

guard CommandLine.arguments.count >= 2 else {
    fputs("usage: vision_ocr.swift <image_path>\n", stderr)
    exit(1)
}

let imagePath = CommandLine.arguments[1]
let imageURL = URL(fileURLWithPath: imagePath)

guard
    let source = CGImageSourceCreateWithURL(imageURL as CFURL, nil),
    let cgImage = CGImageSourceCreateImageAtIndex(source, 0, nil)
else {
    fputs("failed to load image\n", stderr)
    exit(2)
}

var recognizedLines: [String] = []
let request = VNRecognizeTextRequest { request, error in
    if let error {
        fputs("vision error: \(error.localizedDescription)\n", stderr)
        exit(3)
    }

    let observations = request.results as? [VNRecognizedTextObservation] ?? []
    for observation in observations {
        if let candidate = observation.topCandidates(1).first {
            recognizedLines.append(candidate.string)
        }
    }
}

request.recognitionLevel = .accurate
request.usesLanguageCorrection = false
request.recognitionLanguages = ["zh-Hans", "en-US"]

do {
    try VNImageRequestHandler(cgImage: cgImage, options: [:]).perform([request])
    print(recognizedLines.joined(separator: "\n"))
} catch {
    fputs("request error: \(error.localizedDescription)\n", stderr)
    exit(4)
}
