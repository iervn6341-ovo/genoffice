// print-pdf: shows the standard macOS print panel (with its page preview) for a PDF and prints it.
//
// Electron's webContents.print() opens Chromium's own system panel, which has no preview. Driving
// NSPrintOperation through PDFKit gives the real panel: preview pane, "PDF" menu, presets, AirPrint.
//
// usage: print-pdf <pdf-path> [job-title]
// exit codes: 0 printed / saved from the panel, 1 canceled, 2 bad arguments or unreadable PDF,
//             3 the print operation could not be created
import Cocoa
import PDFKit

let arguments = CommandLine.arguments
guard arguments.count >= 2 else {
  FileHandle.standardError.write(Data("usage: print-pdf <pdf-path> [job-title]\n".utf8))
  exit(2)
}
guard let document = PDFDocument(url: URL(fileURLWithPath: arguments[1])) else {
  FileHandle.standardError.write(Data("cannot open the PDF\n".utf8))
  exit(2)
}

let app = NSApplication.shared
// no Dock icon / menu bar of its own, but windows (the print panel) are allowed
app.setActivationPolicy(.accessory)
app.activate(ignoringOtherApps: true)

let info = NSPrintInfo()
info.topMargin = 0
info.bottomMargin = 0
info.leftMargin = 0
info.rightMargin = 0
info.isHorizontallyCentered = true
info.isVerticallyCentered = true
info.scalingFactor = 1

// start the panel on the PDF's own paper and orientation (the document's page setup is baked into
// its pages), so Paper Size and Orientation open on the right values
if let first = document.page(at: 0) {
  let box = first.bounds(for: .mediaBox)
  let rotated = first.rotation % 180 != 0
  let width = rotated ? box.height : box.width
  let height = rotated ? box.width : box.height
  info.paperSize = NSSize(width: min(width, height), height: max(width, height))
  info.orientation = width > height ? .landscape : .portrait
}

// A page that does not fit the paper the user picks is shrunk to fit (never cropped), and the pages
// are not auto-rotated, so Orientation in the panel decides how the sheet is fed
guard
  let operation = document.printOperation(
    for: info, scalingMode: .pageScaleDownToFit, autoRotate: false)
else {
  exit(3)
}
operation.jobTitle = arguments.count >= 3 ? arguments[2] : "GenOffice"
operation.showsPrintPanel = true
operation.showsProgressPanel = true
operation.printPanel.options.formUnion([.showsPreview, .showsPaperSize, .showsOrientation, .showsScaling])

exit(operation.run() ? 0 : 1)
