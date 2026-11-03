import AppKit
import Foundation
// 讀 jobs.json：[{id, lines:[str], font, size, minSize, maxW, color:[r,g,b], red:[r,g,b], bold}]
// 每行輸出 <id>_<n>.png（透明底），並寫 out.json：每行的寬高與基線
struct Job: Codable { let id: String; let lines: [String]; let font: String; let size: Double; let minSize: Double; let maxW: Double; let color: [Double]; let red: [Double]; let bold: Bool }
let args = CommandLine.arguments
let jobs = try! JSONDecoder().decode([Job].self, from: Data(contentsOf: URL(fileURLWithPath: args[1])))
let outDir = args[2]
func mkFont(_ name: String, _ size: Double, _ bold: Bool) -> NSFont {
  if name == "rounded" {
    let base = NSFont.systemFont(ofSize: size, weight: bold ? .bold : .medium)
    if let d = base.fontDescriptor.withDesign(.rounded) { return NSFont(descriptor: d, size: size)! }
    return base
  }
  return NSFont(name: name, size: size)!
}
func col(_ c: [Double]) -> NSColor { NSColor(srgbRed: c[0]/255, green: c[1]/255, blue: c[2]/255, alpha: 1) }
func attr(_ line: String, _ f: NSFont, _ c: NSColor, _ r: NSColor) -> NSAttributedString {
  let s = NSMutableAttributedString(); var red = false; var buf = ""
  func flush() { if !buf.isEmpty { s.append(NSAttributedString(string: buf, attributes: [.font: f, .foregroundColor: red ? r : c])); buf = "" } }
  for ch in line { if ch == "[" { flush(); red = true } else if ch == "]" { flush(); red = false } else { buf.append(ch) } }
  flush(); return s
}
var meta: [[String: Any]] = []
for j in jobs {
  var size = j.size
  var strs: [NSAttributedString] = []
  while true {
    let f = mkFont(j.font, size, j.bold)
    strs = j.lines.map { attr($0, f, col(j.color), col(j.red)) }
    let w = strs.map { $0.size().width }.max() ?? 0
    if w <= j.maxW || size <= j.minSize { break }
    size -= 0.5
  }
  let f = mkFont(j.font, size, j.bold)
  for (n, s) in strs.enumerated() {
    let line = CTLineCreateWithAttributedString(s)
    var asc: CGFloat = 0, desc: CGFloat = 0, lead: CGFloat = 0
    let w = CGFloat(CTLineGetTypographicBounds(line, &asc, &desc, &lead))
    let pad: CGFloat = 4
    let W = Int(ceil(w + pad * 2)), H = Int(ceil(asc + desc + pad * 2))
    let rep = NSBitmapImageRep(bitmapDataPlanes: nil, pixelsWide: W, pixelsHigh: H, bitsPerSample: 8, samplesPerPixel: 4, hasAlpha: true, isPlanar: false, colorSpaceName: .deviceRGB, bytesPerRow: 0, bitsPerPixel: 0)!
    NSGraphicsContext.saveGraphicsState()
    let ctx = NSGraphicsContext(bitmapImageRep: rep)!
    NSGraphicsContext.current = ctx
    let cg = ctx.cgContext
    cg.setShouldAntialias(true); cg.setShouldSmoothFonts(false)
    cg.textPosition = CGPoint(x: pad, y: pad + desc)
    CTLineDraw(line, cg)
    NSGraphicsContext.restoreGraphicsState()
    let p = "\(outDir)/\(j.id)_\(n).png"
    try! rep.representation(using: .png, properties: [:])!.write(to: URL(fileURLWithPath: p))
    meta.append(["id": j.id, "n": n, "w": Double(w), "W": W, "H": H, "asc": Double(asc), "desc": Double(desc), "pad": Double(pad), "size": size, "font": f.fontName])
  }
}
try! JSONSerialization.data(withJSONObject: meta, options: [.prettyPrinted]).write(to: URL(fileURLWithPath: "\(outDir)/out.json"))
print("ok", meta.count)
