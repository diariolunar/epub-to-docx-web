import JSZip from "jszip";
import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  HeadingLevel,
  AlignmentType,
  PageBreak,
  Bookmark,
  InternalHyperlink,
  ImageRun,
} from "docx";

type ChapterBlock =
  | { type: "text"; text: string }
  | { type: "image"; src: string };

function decodeHtmlEntities(text: string) {
  return text
    .replace(/&nbsp;/g, " ")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, "\"")
    .replace(/&#39;/g, "'");
}

function cleanText(html: string) {
  return decodeHtmlEntities(
    html
      .replace(/<script[\s\S]*?<\/script>/gi, "")
      .replace(/<style[\s\S]*?<\/style>/gi, "")
      .replace(/<[^>]+>/g, "")
      .replace(/\s+/g, " ")
      .trim()
  );
}

function getTitleFromHtml(html: string, fallback: string) {
  const h1 = html.match(/<h1[^>]*>(.*?)<\/h1>/i);
  const h2 = html.match(/<h2[^>]*>(.*?)<\/h2>/i);
  const title = h1?.[1] || h2?.[1];

  if (!title) return fallback;

  return cleanText(title) || fallback;
}

function getCleanFileName(path: string) {
  return path.split("/").pop()?.replace(/\.(xhtml|html|htm)$/i, "") || path;
}

function shouldIgnoreHtmlFile(name: string) {
  const lower = name.toLowerCase();

  return (
    lower.includes("toc") ||
    lower.includes("nav") ||
    lower.includes("cover") ||
    lower.includes("titlepage") ||
    lower.includes("title-page") ||
    lower.includes("copyright")
  );
}

function makeBookmarkId(index: number) {
  return `chapter_${index + 1}`;
}

function getImageSrc(tag: string) {
  const src =
    tag.match(/\ssrc=["']([^"']+)["']/i)?.[1] ||
    tag.match(/\shref=["']([^"']+)["']/i)?.[1] ||
    tag.match(/\sxlink:href=["']([^"']+)["']/i)?.[1];

  return src ? decodeHtmlEntities(src.trim()) : "";
}

function parseHtmlBlocks(html: string): ChapterBlock[] {
  const blocks: ChapterBlock[] = [];

  const body =
    html.match(/<body[^>]*>([\s\S]*?)<\/body>/i)?.[1] || html;

  const regex =
    /<(h1|h2|h3|p|div|img|image)[^>]*>[\s\S]*?<\/\1>|<(img|image)[^>]*\/?>/gi;

  const matches = body.match(regex) || [];

  for (const item of matches) {
    if (/^<(img|image)/i.test(item)) {
      const src = getImageSrc(item);
      if (src) blocks.push({ type: "image", src });
      continue;
    }

    const imgTags = item.match(/<(img|image)[^>]*\/?>/gi);
    if (imgTags) {
      for (const imgTag of imgTags) {
        const src = getImageSrc(imgTag);
        if (src) blocks.push({ type: "image", src });
      }
    }

    const text = cleanText(item);
    if (text && !/^https?:\/\/img\.wattpad\.com/i.test(text)) {
      blocks.push({ type: "text", text });
    }
  }

  return blocks;
}

function resolveZipPath(currentHtmlFile: string, src: string) {
  if (src.startsWith("http://") || src.startsWith("https://")) return src;

  const currentDir = currentHtmlFile.split("/").slice(0, -1).join("/");
  const combined = currentDir ? `${currentDir}/${src}` : src;

  const parts: string[] = [];

  for (const part of combined.split("/")) {
    if (!part || part === ".") continue;
    if (part === "..") parts.pop();
    else parts.push(part);
  }

  return parts.join("/");
}

async function getImageData(zip: JSZip, currentHtmlFile: string, src: string) {
  try {
    const resolved = resolveZipPath(currentHtmlFile, src);

    if (resolved.startsWith("http://") || resolved.startsWith("https://")) {
      const response = await fetch(resolved);
      if (!response.ok) return null;

      const contentType = response.headers.get("content-type") || "";
      if (!contentType.startsWith("image/")) return null;

      const buffer = await response.arrayBuffer();

      return {
        data: new Uint8Array(buffer),
        type: contentType.includes("png") ? "png" : "jpg",
      } as const;
    }

    const file = zip.file(resolved);
    if (!file) return null;

    const buffer = await file.async("uint8array");
    const lower = resolved.toLowerCase();

    return {
      data: buffer,
      type: lower.endsWith(".png") ? "png" : "jpg",
    } as const;
  } catch {
    return null;
  }
}

export async function convertEpubToDocx(
  arrayBuffer: ArrayBuffer,
  fileName: string
) {
  const zip = await JSZip.loadAsync(arrayBuffer);

  const htmlFiles = Object.keys(zip.files)
    .filter((name) => name.match(/\.(xhtml|html|htm)$/i))
    .filter((name) => !shouldIgnoreHtmlFile(name))
    .sort();

  const chapterData: {
    title: string;
    blocks: ChapterBlock[];
    htmlFile: string;
    bookmarkId: string;
  }[] = [];

  for (const htmlFile of htmlFiles) {
    const rawHtml = await zip.files[htmlFile].async("text");
    const blocks = parseHtmlBlocks(rawHtml);

    if (blocks.length === 0) continue;

    const textBlocks = blocks.filter((block) => block.type === "text");
    if (textBlocks.length === 0) continue;

    const title = getTitleFromHtml(rawHtml, getCleanFileName(htmlFile));
    if (title.toLowerCase() === "cover") continue;

    chapterData.push({
      title,
      blocks,
      htmlFile,
      bookmarkId: makeBookmarkId(chapterData.length),
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
        children: [
          new InternalHyperlink({
            anchor: chapter.bookmarkId,
            children: [
              new TextRun({
                text: chapter.title,
                style: "Hyperlink",
              }),
            ],
          }),
        ],
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
        children: [
          new Bookmark({
            id: chapter.bookmarkId,
            children: [
              new TextRun({
                text: chapter.title,
                bold: true,
                size: 32,
              }),
            ],
          }),
        ],
        heading: HeadingLevel.HEADING_1,
        alignment: AlignmentType.CENTER,
        spacing: { after: 300 },
      })
    );

    for (const block of chapter.blocks) {
      if (block.type === "text") {
        const cleanLine = block.text.replace(/^#+\s*/, "").trim();

        if (!cleanLine || cleanLine === chapter.title) continue;
        if (/^https?:\/\/img\.wattpad\.com/i.test(cleanLine)) continue;

        children.push(
          new Paragraph({
            children: [
              new TextRun({
                text: cleanLine,
                size: 24,
              }),
            ],
            spacing: { after: 120 },
          })
        );
      }

      if (block.type === "image") {
        const image = await getImageData(zip, chapter.htmlFile, block.src);

        if (!image) continue;

        children.push(
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 200, after: 200 },
            children: [
              new ImageRun({
                data: image.data,
                transformation: {
                  width: 420,
                  height: 260,
                },
                type: image.type,
              }),
            ],
          })
        );
      }
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