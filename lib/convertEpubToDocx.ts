import JSZip from "jszip";
import { parseDocument } from "htmlparser2";
import {
  AlignmentType,
  Bookmark,
  Document,
  HeadingLevel,
  ImageRun,
  InternalHyperlink,
  Packer,
  PageBreak,
  Paragraph,
  TextRun,
} from "docx";

type DomNode = {
  type: string;
  name?: string;
  data?: string;
  attribs?: Record<string, string>;
  children?: DomNode[];
};

type RunStyle = { bold?: boolean; italics?: boolean; underline?: {}; superScript?: boolean; subScript?: boolean };
type Heading = (typeof HeadingLevel)[keyof typeof HeadingLevel];
type TextBlock = { type: "text"; runs: Array<{ text?: string; break?: number; style: RunStyle }>; heading?: Heading; bullet?: boolean; pageBreakBefore?: boolean };
type ImageBlock = { type: "image"; src: string; pageBreakBefore?: boolean };
type ChapterBlock = TextBlock | ImageBlock;

type ImageData = {
  data: Uint8Array;
  type: "jpg" | "png" | "gif" | "bmp";
  width: number;
  height: number;
};

const BLOCK_TAGS = new Set(["p", "div", "section", "article", "aside", "blockquote", "li", "pre", "address", "figure", "figcaption", "dt", "dd"]);
const PAGE_BREAK_PATTERN = /(?:page-break-(?:before|after)\s*:\s*(?:always|page)|break-(?:before|after)\s*:\s*page)/i;

function childrenOf(node: DomNode) {
  return node.children || [];
}

function isElement(node: DomNode) {
  return node.type === "tag" || node.type === "script" || node.type === "style";
}

function getText(node: DomNode): string {
  if (node.type === "text") return node.data || "";
  return childrenOf(node).map(getText).join("");
}

function normalizePath(path: string) {
  const parts: string[] = [];
  for (const part of path.replace(/\\/g, "/").split("/")) {
    if (!part || part === ".") continue;
    if (part === "..") parts.pop();
    else parts.push(part);
  }
  return parts.join("/");
}

function resolveZipPath(currentFile: string, target: string) {
  const cleanTarget = target.split("#")[0].split("?")[0];
  if (/^https?:\/\//i.test(cleanTarget)) return cleanTarget;
  const directory = currentFile.split("/").slice(0, -1).join("/");
  return normalizePath(directory ? `${directory}/${cleanTarget}` : cleanTarget);
}

function cleanFileName(path: string) {
  return path.split("/").pop()?.replace(/\.(xhtml|html|htm)$/i, "") || path;
}

function imageTypeFromName(name: string): ImageData["type"] {
  if (/\.png$/i.test(name)) return "png";
  if (/\.gif$/i.test(name)) return "gif";
  if (/\.bmp$/i.test(name)) return "bmp";
  return "jpg";
}

function fitImageSize(width: number, height: number, maxWidth: number, maxHeight: number) {
  const ratio = Math.min(maxWidth / width, maxHeight / height, 1);
  return { width: Math.round(width * ratio), height: Math.round(height * ratio) };
}

async function getImageSize(data: Uint8Array) {
  // createImageBitmap is available in browsers; the server falls back to a safe size.
  if (typeof createImageBitmap !== "function") return { width: 800, height: 600 };
  try {
    const bitmap = await createImageBitmap(new Blob([data.buffer.slice(data.byteOffset, data.byteOffset + data.byteLength) as ArrayBuffer]));
    return { width: bitmap.width, height: bitmap.height };
  } catch {
    return { width: 800, height: 600 };
  }
}

async function getImageData(zip: JSZip, htmlFile: string, src: string): Promise<ImageData | null> {
  const path = resolveZipPath(htmlFile, src);
  // Do not fetch remote URLs from an uploaded EPUB: this keeps the API free of SSRF.
  if (/^https?:\/\//i.test(path)) return null;
  const file = zip.file(path);
  if (!file) return null;
  const data = await file.async("uint8array");
  const size = await getImageSize(data);
  return { data, type: imageTypeFromName(path), ...size };
}

function findFirst(node: DomNode, predicate: (item: DomNode) => boolean): DomNode | undefined {
  if (predicate(node)) return node;
  for (const child of childrenOf(node)) {
    const found = findFirst(child, predicate);
    if (found) return found;
  }
}

function findAll(node: DomNode, predicate: (item: DomNode) => boolean, result: DomNode[] = []) {
  if (predicate(node)) result.push(node);
  for (const child of childrenOf(node)) findAll(child, predicate, result);
  return result;
}

async function getSpineHtmlFiles(zip: JSZip) {
  const container = zip.file("META-INF/container.xml");
  if (container) {
    const containerDom = parseDocument(await container.async("text"), { xmlMode: true, decodeEntities: true }) as unknown as DomNode;
    const rootfile = findFirst(containerDom, (node) => node.name?.toLowerCase() === "rootfile");
    const opfPath = rootfile?.attribs?.["full-path"];
    const opf = opfPath && zip.file(opfPath);
    if (opf && opfPath) {
      const opfDom = parseDocument(await opf.async("text"), { xmlMode: true, decodeEntities: true }) as unknown as DomNode;
      const manifest = new Map<string, string>();
      for (const item of findAll(opfDom, (node) => node.name?.toLowerCase() === "item")) {
        const id = item.attribs?.id;
        const href = item.attribs?.href;
        const type = item.attribs?.["media-type"];
        if (id && href && /(?:xhtml|html)/i.test(type || href)) manifest.set(id, resolveZipPath(opfPath, href));
      }
      const spine = findAll(opfDom, (node) => node.name?.toLowerCase() === "itemref")
        .map((item) => manifest.get(item.attribs?.idref || ""))
        .filter((path): path is string => Boolean(path && zip.file(path)));
      if (spine.length) return spine;
    }
  }
  return Object.keys(zip.files).filter((name) => /\.(xhtml|html|htm)$/i.test(name)).sort();
}

function headingFor(tag: string): Heading | undefined {
  const headings: Record<string, Heading> = {
    h1: HeadingLevel.HEADING_1, h2: HeadingLevel.HEADING_2, h3: HeadingLevel.HEADING_3,
    h4: HeadingLevel.HEADING_4, h5: HeadingLevel.HEADING_5, h6: HeadingLevel.HEADING_6,
  };
  return headings[tag];
}

function hasPageBreak(node: DomNode) {
  const attrs = node.attribs || {};
  return PAGE_BREAK_PATTERN.test(attrs.style || "") || /pagebreak|page-break/i.test(`${attrs.class || ""} ${attrs["epub:type"] || ""}`);
}

function inlineRuns(nodes: DomNode[], style: RunStyle = {}): TextBlock["runs"] {
  const runs: TextBlock["runs"] = [];
  const addText = (text: string) => {
    const normalized = text.replace(/[\t\r\n ]+/g, " ");
    if (normalized) runs.push({ text: normalized, style });
  };
  for (const node of nodes) {
    if (node.type === "text") { addText(node.data || ""); continue; }
    if (!isElement(node)) continue;
    const tag = node.name?.toLowerCase() || "";
    if (tag === "br") { runs.push({ break: 1, style }); continue; }
    if (tag === "script" || tag === "style" || tag === "noscript") continue;
    const nextStyle: RunStyle = {
      ...style,
      bold: style.bold || tag === "strong" || tag === "b",
      italics: style.italics || tag === "em" || tag === "i",
      underline: style.underline || (tag === "u" ? {} : undefined),
      superScript: style.superScript || tag === "sup",
      subScript: style.subScript || tag === "sub",
    };
    runs.push(...inlineRuns(childrenOf(node), nextStyle));
  }
  return runs;
}

function parseHtmlBlocks(html: string): ChapterBlock[] {
  const document = parseDocument(html, { xmlMode: false, decodeEntities: true }) as unknown as DomNode;
  const body = findFirst(document, (node) => node.name?.toLowerCase() === "body") || document;
  const blocks: ChapterBlock[] = [];

  const visit = (nodes: DomNode[]) => {
    for (const node of nodes) {
      if (node.type === "text" && node.data?.trim()) {
        const runs = inlineRuns([node]);
        if (runs.length) blocks.push({ type: "text", runs });
        continue;
      }
      if (!isElement(node)) continue;
      const tag = node.name?.toLowerCase() || "";
      if (tag === "script" || tag === "style" || tag === "nav") continue;
      if (tag === "img" || tag === "image") {
        const src = node.attribs?.src || node.attribs?.href || node.attribs?.["xlink:href"];
        if (src) blocks.push({ type: "image", src, pageBreakBefore: hasPageBreak(node) });
        continue;
      }
      if (tag === "br") continue;
      if (BLOCK_TAGS.has(tag) || headingFor(tag)) {
        const containsNestedBlocks = childrenOf(node).some((child) => {
          const childTag = child.name?.toLowerCase() || "";
          return BLOCK_TAGS.has(childTag) || Boolean(headingFor(childTag));
        });
        // EPUBs commonly use div/section solely as wrappers around paragraphs.
        // Keeping their children separate prevents text from adjacent paragraphs joining.
        if (tag !== "p" && tag !== "li" && containsNestedBlocks) {
          if (hasPageBreak(node)) blocks.push({ type: "text", runs: [], pageBreakBefore: true });
          visit(childrenOf(node));
          continue;
        }
        const runs = inlineRuns(childrenOf(node));
        if (runs.some((run) => run.text?.trim() || run.break)) {
          blocks.push({ type: "text", runs, heading: headingFor(tag), bullet: tag === "li", pageBreakBefore: hasPageBreak(node) });
        }
        continue;
      }
      if (hasPageBreak(node)) blocks.push({ type: "text", runs: [], pageBreakBefore: true });
      visit(childrenOf(node));
    }
  };
  visit(childrenOf(body));
  return blocks;
}

function titleFromDocument(html: string, fallback: string) {
  const document = parseDocument(html, { xmlMode: false, decodeEntities: true }) as unknown as DomNode;
  const heading = findFirst(document, (node) => ["h1", "h2"].includes(node.name?.toLowerCase() || ""));
  return getText(heading || document).replace(/\s+/g, " ").trim() || fallback;
}

function findCoverImagePath(zip: JSZip) {
  return Object.keys(zip.files).find((name) => /cover/i.test(name) && /\.(jpg|jpeg|png|gif|bmp)$/i.test(name)) || null;
}

async function getCoverImage(zip: JSZip) {
  const path = findCoverImagePath(zip);
  return path ? getImageData(zip, "", path) : null;
}

export async function convertEpubToDocx(arrayBuffer: ArrayBuffer, fileName: string) {
  const zip = await JSZip.loadAsync(arrayBuffer);
  const htmlFiles = await getSpineHtmlFiles(zip);
  const chapterData: Array<{ title: string; blocks: ChapterBlock[]; htmlFile: string; bookmarkId: string }> = [];

  for (const htmlFile of htmlFiles) {
    const file = zip.file(htmlFile);
    if (!file) continue;
    const html = await file.async("text");
    const blocks = parseHtmlBlocks(html);
    if (!blocks.some((block) => block.type === "text" && block.runs.some((run) => run.text?.trim()))) continue;
    const title = titleFromDocument(html, cleanFileName(htmlFile));
    chapterData.push({ title, blocks, htmlFile, bookmarkId: `chapter_${chapterData.length + 1}` });
  }
  if (!chapterData.length) throw new Error("Nenhum capítulo de texto foi encontrado neste EPUB.");

  const children: Paragraph[] = [];
  const cover = await getCoverImage(zip);
  if (cover) {
    const size = fitImageSize(cover.width, cover.height, 595, 842);
    children.push(new Paragraph({ alignment: AlignmentType.CENTER, children: [new ImageRun({ data: cover.data, type: cover.type, transformation: size })] }));
    children.push(new Paragraph({ children: [new PageBreak()] }));
  } else {
    children.push(new Paragraph({ text: fileName.replace(/\.epub$/i, ""), heading: HeadingLevel.TITLE, alignment: AlignmentType.CENTER }));
  }
  children.push(new Paragraph({ text: "Sumário", heading: HeadingLevel.HEADING_1, alignment: AlignmentType.CENTER }));
  for (const chapter of chapterData) {
    children.push(new Paragraph({ children: [new InternalHyperlink({ anchor: chapter.bookmarkId, children: [new TextRun({ text: chapter.title, style: "Hyperlink" })] })] }));
  }
  for (const chapter of chapterData) {
    children.push(new Paragraph({ children: [new PageBreak()] }));
    children.push(new Paragraph({ heading: HeadingLevel.HEADING_1, alignment: AlignmentType.CENTER, children: [new Bookmark({ id: chapter.bookmarkId, children: [new TextRun({ text: chapter.title, bold: true, size: 32 })] })] }));
    for (const block of chapter.blocks) {
      if (block.pageBreakBefore) children.push(new Paragraph({ children: [new PageBreak()] }));
      if (block.type === "image") {
        const image = await getImageData(zip, chapter.htmlFile, block.src);
        if (!image) continue;
        children.push(new Paragraph({ alignment: AlignmentType.CENTER, children: [new ImageRun({ data: image.data, type: image.type, transformation: fitImageSize(image.width, image.height, 420, 520) })] }));
      } else if (block.runs.length) {
        children.push(new Paragraph({ heading: block.heading, bullet: block.bullet ? { level: 0 } : undefined, spacing: { after: 120 }, children: block.runs.map((run) => new TextRun({ text: run.text, break: run.break, size: 24, ...run.style })) }));
      }
    }
  }
  return Packer.toBuffer(new Document({ sections: [{ properties: { page: { margin: { top: 1134, right: 1134, bottom: 1134, left: 1134 } } }, children }] }));
}
