import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import { convertEpubToDocx } from "./convertEpubToDocx";

async function createEpub() {
  const zip = new JSZip();
  zip.file("META-INF/container.xml", `<?xml version="1.0"?><container><rootfiles><rootfile full-path="OPS/book.opf"/></rootfiles></container>`);
  zip.file("OPS/book.opf", `<?xml version="1.0"?><package><manifest>
    <item id="second" href="second.xhtml" media-type="application/xhtml+xml"/>
    <item id="first" href="first.xhtml" media-type="application/xhtml+xml"/>
  </manifest><spine><itemref idref="first"/><itemref idref="second"/></spine></package>`);
  zip.file("OPS/first.xhtml", `<html><body><h1>Primeiro</h1><div><p>Uma <em>frase</em> com <strong>espaços</strong>.</p><p>Outro parágrafo.</p></div><p>Linha A<br/>Linha B</p><p style="page-break-before: always">Nova página</p></body></html>`);
  zip.file("OPS/second.xhtml", `<html><body><h1>Segundo</h1><ul><li>Item um</li><li>Item dois</li></ul></body></html>`);
  return zip.generateAsync({ type: "arraybuffer" });
}

describe("convertEpubToDocx", () => {
  it("respeita o spine e preserva estrutura básica do XHTML", async () => {
    const output = await convertEpubToDocx(await createEpub(), "livro.epub");
    const docx = await JSZip.loadAsync(output);
    const xml = await docx.file("word/document.xml")!.async("text");

    expect(xml.indexOf("Primeiro")).toBeLessThan(xml.indexOf("Segundo"));
    expect(xml.match(/>Primeiro</g)).toHaveLength(2); // sumário + título do capítulo
    expect(xml).toContain("Uma ");
    expect(xml).toContain("frase");
    expect(xml).toContain("espaços");
    expect(xml).toContain("Outro parágrafo.");
    expect(xml).toContain("Linha A");
    expect(xml).toContain("Linha B");
    expect(xml).toContain("w:br");
    expect(xml).toContain("w:br w:type=\"page\"");
    expect(xml).toContain("w:numPr");
  });
});
