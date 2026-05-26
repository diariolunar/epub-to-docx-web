import JSZip from "jszip";
import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  HeadingLevel,
} from "docx";

function stripHtml(html: string) {
  return html
    .replace(/<script[\s\S]*?<\/script>/gi, "")
    .replace(/<style[\s\S]*?<\/style>/gi, "")
    .replace(/<h1[^>]*>(.*?)<\/h1>/gi, "\n# $1\n")
    .replace(/<h2[^>]*>(.*?)<\/h2>/gi, "\n## $1\n")
    .replace(/<h3[^>]*>(.*?)<\/h3>/gi, "\n### $1\n")
    .replace(/<p[^>]*>/gi, "\n")
    .replace(/<\/p>/gi, "\n")
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<[^>]+>/g, "")
    .replace(/&nbsp;/g, " ")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, "\"")
    .replace(/&#39;/g, "'")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

export async function convertEpubToDocx(
  arrayBuffer: ArrayBuffer,
  fileName: string
) {
  const zip = await JSZip.loadAsync(arrayBuffer);

  const htmlFiles = Object.keys(zip.files)
    .filter((name) =>
      name.match(/\.(xhtml|html|htm)$/i)
    )
    .sort();

  const children: Paragraph[] = [];

  children.push(
    new Paragraph({
      text: fileName.replace(/\.epub$/i, ""),
      heading: HeadingLevel.TITLE,
    })
  );

  for (const htmlFile of htmlFiles) {
    const rawHtml = await zip.files[htmlFile].async("text");
    const text = stripHtml(rawHtml);

    if (!text) continue;

    const lines = text
      .split("\n")
      .map((line) => line.trim())
      .filter(Boolean);

    for (const line of lines) {
      if (line.startsWith("# ")) {
        children.push(
          new Paragraph({
            text: line.replace("# ", ""),
            heading: HeadingLevel.HEADING_1,
          })
        );
      } else if (line.startsWith("## ")) {
        children.push(
          new Paragraph({
            text: line.replace("## ", ""),
            heading: HeadingLevel.HEADING_2,
          })
        );
      } else if (line.startsWith("### ")) {
        children.push(
          new Paragraph({
            text: line.replace("### ", ""),
            heading: HeadingLevel.HEADING_3,
          })
        );
      } else {
        children.push(
          new Paragraph({
            children: [new TextRun(line)],
          })
        );
      }
    }
  }

  const doc = new Document({
    sections: [
      {
        children,
      },
    ],
  });

  return await Packer.toBuffer(doc);
}