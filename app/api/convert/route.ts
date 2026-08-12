import { NextRequest, NextResponse } from "next/server";
import { convertEpubToDocx } from "@/lib/convertEpubToDocx";

export const runtime = "nodejs";
const MAX_FILE_SIZE_BYTES = 25 * 1024 * 1024;

export async function POST(request: NextRequest) {
  try {
    const formData = await request.formData();
    const file = formData.get("file");

    if (!file || !(file instanceof File)) {
      return NextResponse.json(
        { error: "Nenhum arquivo EPUB enviado." },
        { status: 400 }
      );
    }

    if (!file.name.toLowerCase().endsWith(".epub")) {
      return NextResponse.json(
        { error: "O arquivo precisa ser .epub." },
        { status: 400 }
      );
    }

    if (file.size > MAX_FILE_SIZE_BYTES) {
      return NextResponse.json(
        { error: "O arquivo excede o limite de 25 MB." },
        { status: 413 }
      );
    }

    const arrayBuffer = await file.arrayBuffer();

    const docxBuffer = await convertEpubToDocx(
      arrayBuffer,
      file.name
    );

    const outputName = file.name.replace(/\.epub$/i, ".docx");

    return new NextResponse(new Uint8Array(docxBuffer), {
      status: 200,
      headers: {
        "Content-Type":
          "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        "Content-Disposition": `attachment; filename*=UTF-8''${encodeURIComponent(outputName)}`,
      },
    });
  } catch (error) {
    console.error("Erro detalhado na conversão:", error);

    const message =
      error instanceof Error
        ? error.message
        : "Erro desconhecido durante a conversão.";

    return NextResponse.json(
      {
        error: `Falha ao converter: ${message}`,
      },
      { status: 500 }
    );
  }
}
