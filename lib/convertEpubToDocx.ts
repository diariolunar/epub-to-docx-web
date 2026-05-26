import JSZip from "jszip";
import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  HeadingLevel,
  AlignmentType,
  PageBreak,
} from "docx";

function decodeHtmlEntities(text: string) {
  return text
    .replace(/&nbsp;/g, " ")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, "\"")
    .replace(/&#39;/g, "'");
}

function stripHtml(html: string) {
  return decodeHtmlEntities(
    html
      .replace(/<script[\s\S]*?<\/script>/gi, "")
      .replace(/<style[\s\S]*?<\/style>/gi, "")
      .replace(/<nav[\s\S]*?<\/nav>/gi, "")
      .replace(/<h1[^>]*>(.*?)<\/h1>/gi, "\n# $1\n")
      .replace(/<h2[^>]*>(.*?)<\/h2>/gi, "\n# $1\n")
      .replace(/<h3[^>]*>(.*?)<\/h3>/gi, "\n## $1\n")
      .replace(/<p[^>]*>/gi, "\n")
      .replace(/<\/p>/gi, "\n")
      .replace(/<br\s*\/?>/gi, "\n")
      .replace(/<[^>]+>/g, "")
      .replace(/\n{3,}/g, "\n\n")
      .trim()
  );
}

function getTitleFromHtml(html: string, fallback: string) {
  const h1 = html.match(/<h1[^>]*>(.*?)<\/h1>/i);
  const h2 = html.match(/<h2[^>]*>(.*?)<\/h2>/i);
  const title = h1?.[1] || h2?.[1];

  if (!title) return fallback;

  return stripHtml(title).replace(/^#+\s*/, "").trim() || fallback;
}

function getCleanFileName(path: string) {
  return path.split("/").pop()?.replace(/\.(xhtml|html|htm)$/i, "") || path;
}

export async function convertEpubToDocx(
  arrayBuffer: ArrayBuffer,
  fileName: string
) {
  const zip = await JSZip.loadAsync(arrayBuffer);

  const htmlFiles = Object.keys(zip.files)
    .filter((name) => name.match(/\.(xhtml|html|htm)$/i))
    .filter((name) => !name.toLowerCase().includes("toc"))
    .filter((name) => !name.toLowerCase().includes("nav"))
    .sort();

  const chapterData: {
    title: string;
    lines: string[];
  }[] = [];

  for (const htmlFile of htmlFiles) {
    const rawHtml = await zip.files[htmlFile].async("text");
    const text = stripHtml(rawHtml);

    if (!text) continue;

    const lines = text
      .split("\n")
      .map((line) => line.trim())
      .filter(Boolean);

    if (lines.length === 0) continue;

    const title = getTitleFromHtml(rawHtml, getCleanFileName(htmlFile));

    chapterData.push({
      title,
      lines,
    });
  }

  const children: Paragraph[] = [];

  children.push(
    new Paragraph({
      text: fileName.replace(/\.epub$/i, ""),
      heading: HeadingLevel.TITLE,
      alignment: AlignmentType.CENTER,
      spacing: { after: 400 },
    })
  );

  children.push(
    new Paragraph({
      text: "Table of Contents",
      heading: HeadingLevel.HEADING_1,
      alignment: AlignmentType.CENTER,
      spacing: { before: 400, after: 300 },
    })
  );

  for (const chapter of chapterData) {
    children.push(
      new Paragraph({
        text: chapter.title,
        spacing: { after: 120 },
      })
    );
  }

  for (const chapter of chapterData) {
    children.push(
      new Paragraph({
        children: [new PageBreak()],
      })
    );

    children.push(
      new Paragraph({
        text: chapter.title,
        heading: HeadingLevel.HEADING_1,
        alignment: AlignmentType.CENTER,
        spacing: { after: 300 },
      })
    );

    for (const line of chapter.lines) {
      const cleanLine = line.replace(/^#+\s*/, "").trim();

      if (!cleanLine || cleanLine === chapter.title) continue;

      children.push(
        new Paragraph({
          children: [
            new TextRun({
              text: cleanLine,
              size: 24,
            }),
          ],
          spacing: {
            after: 120,
          },
        })
      );
    }
  }

  const doc = new Document({
    sections: [
      {
        properties: {
          page: {
            margin: {
              top: 1134,
              right: 1134,
              bottom: 1134,
              left: 1134,
            },
          },
        },
        children,
      },
    ],
  });

  return await Packer.toBuffer(doc);
}