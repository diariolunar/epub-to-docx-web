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

type ImageData = {
  data: Uint8Array;
  type: "jpg" | "png" | "gif" | "bmp";
  width: number;
  height: number;
};

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

  return title ? cleanText(title) || fallback : fallback;
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
  const body = html.match(/<body[^>]*>([\s\S]*?)<\/body>/i)?.[1] || html;

  const regex =
    /<(h1|h2|h3|p|div|section|article|img|image|svg)[^>]*>[\s\S]*?<\/\1>|<(img|image)[^>]*\/?>/gi;

  const matches = body.match(regex) || [];

  for (const item of matches) {
    const imgTags = item.match(/<(img|image)[^>]*\/?>/gi);

    if (imgTags) {
      for (const imgTag of imgTags) {
        const src = getImageSrc(imgTag);
        if (src) blocks.push({ type: "image", src });
      }
    }

    const text = cleanText(item);

    if (text && /^https?:\/\/img\.wattpad\.com/i.test(text)) {
      blocks.push({ type: "image", src: text });
      continue;
    }

    if (text) blocks.push({ type: "text", text });
  }

  return blocks;
}

function resolveZipPath(currentHtmlFile: string, src: string) {
  if (src.startsWith("http://") || src.startsWith("https://")) return src;

  const cleanSrc = src.split("#")[0].split("?")[0];
  const currentDir = currentHtmlFile.split("/").slice(0, -1).join("/");
  const combined = currentDir ? `${currentDir}/${cleanSrc}` : cleanSrc;

  const parts: string[] = [];

  for (const part of combined.split("/")) {
    if (!part || part === ".") continue;
    if (part === "..") parts.pop();
    else parts.push(part);
  }

  return parts.join("/");
}

function imageTypeFromName(name: string): ImageData["type"] {
  const lower = name.toLowerCase();

  if (lower.endsWith(".png")) return "png";
  if (lower.endsWith(".gif")) return "gif";
  if (lower.endsWith(".bmp")) return "bmp";

  return "jpg";
}

async function getImageSize(data: Uint8Array) {
  try {
    const blob = new Blob([data]);
    const bitmap = await createImageBitmap(blob);

    return {
      width: bitmap.width,
      height: bitmap.height,
    };
  } catch {
    return {
      width: 800,
      height: 600,
    };
  }
}

function fitImageSize(
  originalWidth: number,
  originalHeight: number,
  maxWidth: number,
  maxHeight: number
) {
  const ratio = Math.min(maxWidth / originalWidth, maxHeight / originalHeight);

  return {
    width: Math.round(originalWidth * ratio),
    height: Math.round(originalHeight * ratio),
  };
}

async function getImageData(
  zip: JSZip,
  currentHtmlFile: string,
  src: string
): Promise<ImageData | null> {
  try {
    const resolved = resolveZipPath(currentHtmlFile, src);

    if (resolved.startsWith("http://") || resolved.startsWith("https://")) {
      const response = await fetch(resolved);

      if (!response.ok) return null;

      const contentType = response.headers.get("content-type") || "";
      if (!contentType.startsWith("image/")) return null;

      const buffer = await response.arrayBuffer();
      const data = new Uint8Array(buffer);
      const size = await getImageSize(data);

      return {
        data,
        width: size.width,
        height: size.height,
        type: contentType.includes("png")
          ? "png"
          : contentType.includes("gif")
          ? "gif"
          : contentType.includes("bmp")
          ? "bmp"
          : "jpg",
      };
    }

    const file = zip.file(resolved);
    if (!file) return null;

    const data = await file.async("uint8array");
    const size = await getImageSize(data);

    return {
      data,
      width: size.width,
      height: size.height,
      type: imageTypeFromName(resolved),
    };
  } catch {
    return null;
  }
}

function findCoverImagePath(zip: JSZip) {
  const files = Object.keys(zip.files);

  const likelyCover = files.find((name) => {
    const lower = name.toLowerCase();

    return (
      lower.includes("cover") && lower.match(/\.(jpg|jpeg|png|gif|bmp)$/i)
    );
  });

  if (likelyCover) return likelyCover;

  return (
    files.find((name) => name.match(/\.(jpg|jpeg|png|gif|bmp)$/i)) || null
  );
}

async function getCoverImage(zip: JSZip): Promise<ImageData | null> {
  const coverPath = findCoverImagePath(zip);

  if (!coverPath) return null;

  const file = zip.file(coverPath);
  if (!file) return null;

  const data = await file.async("uint8array");
  const size = await getImageSize(data);

  return {
    data,
    width: size.width,
    height: size.height,
    type: imageTypeFromName(coverPath),
  };
}

export async function convertEpubToDocx(
  arrayBuffer: ArrayBuffer,
  fileName: string
) {
  const zip = await JSZip.loadAsync(arrayBuffer);
  const coverImage = await getCoverImage(zip);

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

  if (chapterData.length === 0) {
    throw new Error("Nenhum capítulo de texto foi encontrado neste EPUB.");
  }

  const children: Paragraph[] = [];

  if (coverImage) {
    children.push(
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 300 },
        children: [
          new ImageRun({
            data: coverImage.data,
            type: coverImage.type,
            transformation: {
              width: 595,
              height: 842,
            },
          }),
        ],
      })
    );

    children.push(
      new Paragraph({
        children: [new PageBreak()],
      })
    );
  } else {
    children.push(
      new Paragraph({
        text: fileName.replace(/\.epub$/i, ""),
        heading: HeadingLevel.TITLE,
        alignment: AlignmentType.CENTER,
        spacing: { after: 400 },
      })
    );
  }

  children.push(
    new Paragraph({
      text: "Sumário",
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

        const size = fitImageSize(image.width, image.height, 420, 520);

        children.push(
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 200, after: 200 },
            children: [
              new ImageRun({
                data: image.data,
                type: image.type,
                transformation: size,
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