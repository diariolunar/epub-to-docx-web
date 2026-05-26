import { NextRequest, NextResponse } from "next/server";
import { convertEpubToDocx } from "@/lib/convertEpubToDocx";

export const runtime = "nodejs";

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

    const arrayBuffer = await file.arrayBuffer();
    const docxBuffer = await convertEpubToDocx(arrayBuffer, file.name);

    const outputName = file.name.replace(/\.epub$/i, ".docx");

    return new NextResponse(docxBuffer, {
      status: 200,
      headers: {
        "Content-Type":
          "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        "Content-Disposition": `attachment; filename="${outputName}"`,
      },
    });
  } catch (error) {
    console.error(error);

    return NextResponse.json(
      { error: "Erro ao converter o arquivo." },
      { status: 500 }
    );
  }
}