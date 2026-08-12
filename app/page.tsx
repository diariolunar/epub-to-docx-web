"use client";

import { useState } from "react";
import JSZip from "jszip";
import { convertEpubToDocx } from "@/lib/convertEpubToDocx";

const MAX_FILES = 30;
const MAX_FILE_SIZE_MB = 25;
const MAX_FILE_SIZE_BYTES = MAX_FILE_SIZE_MB * 1024 * 1024;

type FileStatus = "waiting" | "converting" | "done" | "error";

type FileItem = {
  file: File;
  status: FileStatus;
  progress: number;
  downloadUrl?: string;
  docxBlob?: Blob;
  errorMessage?: string;
};

function formatBytes(bytes: number) {
  return `${(bytes / 1024 / 1024).toFixed(2)} MB`;
}

function statusLabel(status: FileStatus) {
  if (status === "waiting") return "Aguardando";
  if (status === "converting") return "Convertendo";
  if (status === "done") return "Concluído";
  return "Erro";
}

export default function Home() {
  const [files, setFiles] = useState<FileItem[]>([]);
  const [isConverting, setIsConverting] = useState(false);
  const [generalMessage, setGeneralMessage] = useState("");

  function handleFiles(selected: FileList | null) {
    if (!selected) return;

    const allFiles = Array.from(selected);
    const epubFiles = allFiles.filter((file) =>
      file.name.toLowerCase().endsWith(".epub")
    );

    const limitedFiles = epubFiles.slice(0, MAX_FILES);

    setFiles(
      limitedFiles.map((file) => {
        const isTooLarge = file.size > MAX_FILE_SIZE_BYTES;

        return {
          file,
          status: isTooLarge ? "error" : "waiting",
          progress: 0,
          errorMessage: isTooLarge
            ? `Arquivo maior que ${MAX_FILE_SIZE_MB} MB. Tamanho: ${formatBytes(
                file.size
              )}.`
            : undefined,
        };
      })
    );

    let message = "";

    if (allFiles.length !== epubFiles.length) {
      message += "Alguns arquivos foram ignorados porque não eram EPUB. ";
    }

    if (epubFiles.length > MAX_FILES) {
      message += `Foram selecionados ${epubFiles.length} EPUBs, mas o limite é ${MAX_FILES}.`;
    }

    setGeneralMessage(message.trim());
  }

  async function convertAll() {
    setIsConverting(true);
    setGeneralMessage("");

    const updated = [...files];

    for (let i = 0; i < updated.length; i++) {
      if (updated[i].status === "error") continue;

      try {
        updated[i].status = "converting";
        updated[i].progress = 15;
        updated[i].errorMessage = undefined;
        setFiles([...updated]);

        const arrayBuffer = await updated[i].file.arrayBuffer();

        updated[i].progress = 45;
        setFiles([...updated]);

        const docxBuffer = await convertEpubToDocx(
          arrayBuffer,
          updated[i].file.name
        );

        updated[i].progress = 85;
        setFiles([...updated]);

        const blob = new Blob([new Uint8Array(docxBuffer)], {
          type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        });

        updated[i].docxBlob = blob;
        updated[i].downloadUrl = URL.createObjectURL(blob);
        updated[i].status = "done";
        updated[i].progress = 100;

        setFiles([...updated]);
      } catch (error) {
        updated[i].status = "error";
        updated[i].progress = 0;
        updated[i].errorMessage =
          error instanceof Error
            ? error.message
            : "Erro inesperado durante a conversão.";
        setFiles([...updated]);
      }
    }

    setIsConverting(false);
  }

  async function downloadAllZip() {
    const zip = new JSZip();

    files.forEach((item) => {
      if (item.docxBlob) {
        zip.file(item.file.name.replace(/\.epub$/i, ".docx"), item.docxBlob);
      }
    });

    const zipBlob = await zip.generateAsync({ type: "blob" });
    const url = URL.createObjectURL(zipBlob);

    const a = document.createElement("a");
    a.href = url;
    a.download = "arquivos-convertidos.zip";
    a.click();

    URL.revokeObjectURL(url);
  }

  function clearList() {
    files.forEach((item) => {
      if (item.downloadUrl) URL.revokeObjectURL(item.downloadUrl);
    });

    setFiles([]);
    setGeneralMessage("");
    setIsConverting(false);
  }

  const doneCount = files.filter((item) => item.status === "done").length;
  const errorCount = files.filter((item) => item.status === "error").length;
  const waitingCount = files.filter((item) => item.status === "waiting").length;
  const canDownloadZip = doneCount > 0;

  return (
    <main className="min-h-screen bg-gradient-to-br from-slate-950 via-purple-950 to-slate-900 px-4 py-8 text-white">
      <div className="mx-auto max-w-5xl">
        <section className="mb-8 rounded-3xl border border-white/10 bg-white/10 p-8 shadow-2xl backdrop-blur">
          <p className="mb-2 text-sm font-semibold uppercase tracking-[0.25em] text-purple-200">
            Conversor pessoal
          </p>

          <h1 className="text-4xl font-bold md:text-5xl">EPUB para DOCX</h1>

          <p className="mt-4 max-w-2xl text-slate-200">
            Envie até {MAX_FILES} arquivos EPUB, converta um por vez, baixe
            individualmente ou gere um ZIP com todos os DOCX.
          </p>

          <label className="mt-8 flex cursor-pointer flex-col items-center justify-center rounded-2xl border-2 border-dashed border-purple-300/50 bg-black/20 px-6 py-10 text-center transition hover:border-purple-200 hover:bg-white/10">
            <span className="text-lg font-semibold">
              Clique para selecionar seus EPUBs
            </span>

            <span className="mt-2 text-sm text-slate-300">
              Limite: {MAX_FILES} arquivos • Máximo: {MAX_FILE_SIZE_MB} MB por
              arquivo
            </span>

            <input
              className="hidden"
              type="file"
              accept=".epub"
              multiple
              onChange={(e) => {
                handleFiles(e.target.files);
                // Allow selecting the same EPUB again after clearing or retrying.
                e.currentTarget.value = "";
              }}
            />
          </label>

          {generalMessage && (
            <div className="mt-4 rounded-xl border border-yellow-300/30 bg-yellow-500/10 p-4 text-sm text-yellow-100">
              {generalMessage}
            </div>
          )}

          <div className="mt-6 grid gap-3 md:grid-cols-4">
            <div className="rounded-2xl bg-white/10 p-4">
              <p className="text-sm text-slate-300">Arquivos</p>
              <p className="text-2xl font-bold">{files.length}</p>
            </div>

            <div className="rounded-2xl bg-white/10 p-4">
              <p className="text-sm text-slate-300">Aguardando</p>
              <p className="text-2xl font-bold">{waitingCount}</p>
            </div>

            <div className="rounded-2xl bg-white/10 p-4">
              <p className="text-sm text-slate-300">Concluídos</p>
              <p className="text-2xl font-bold">{doneCount}</p>
            </div>

            <div className="rounded-2xl bg-white/10 p-4">
              <p className="text-sm text-slate-300">Erros</p>
              <p className="text-2xl font-bold">{errorCount}</p>
            </div>
          </div>

          <div className="mt-6 flex flex-wrap gap-3">
            <button
              onClick={convertAll}
              disabled={files.length === 0 || isConverting}
              className="rounded-xl bg-purple-500 px-5 py-3 font-semibold text-white shadow-lg shadow-purple-950/40 transition hover:bg-purple-400 disabled:cursor-not-allowed disabled:opacity-50"
            >
              {isConverting ? "Convertendo..." : "Converter todos"}
            </button>

            <button
              onClick={downloadAllZip}
              disabled={!canDownloadZip || isConverting}
              className="rounded-xl bg-emerald-500 px-5 py-3 font-semibold text-white shadow-lg shadow-emerald-950/40 transition hover:bg-emerald-400 disabled:cursor-not-allowed disabled:opacity-50"
            >
              Baixar todos (.zip)
            </button>

            <button
              onClick={clearList}
              disabled={files.length === 0 || isConverting}
              className="rounded-xl bg-white/10 px-5 py-3 font-semibold text-white transition hover:bg-white/20 disabled:cursor-not-allowed disabled:opacity-50"
            >
              Limpar lista
            </button>
          </div>
        </section>

        <section className="space-y-4">
          {files.length === 0 && (
            <div className="rounded-3xl border border-white/10 bg-white/10 p-8 text-center text-slate-300">
              Nenhum arquivo selecionado ainda.
            </div>
          )}

          {files.map((item, index) => (
            <div
              key={`${item.file.name}-${index}`}
              className="rounded-3xl border border-white/10 bg-white/10 p-5 shadow-xl backdrop-blur"
            >
              <div className="flex flex-col gap-4 md:flex-row md:items-center md:justify-between">
                <div>
                  <h2 className="break-all text-lg font-bold">
                    {item.file.name}
                  </h2>

                  <p className="mt-1 text-sm text-slate-300">
                    {formatBytes(item.file.size)} • {statusLabel(item.status)}
                  </p>
                </div>

                {item.downloadUrl && (
                  <a
                    href={item.downloadUrl}
                    download={item.file.name.replace(/\.epub$/i, ".docx")}
                    className="rounded-xl bg-emerald-500 px-4 py-2 text-sm font-semibold text-white transition hover:bg-emerald-400"
                  >
                    Baixar
                  </a>
                )}
              </div>

              <div className="mt-4 h-3 overflow-hidden rounded-full bg-black/30">
                <div
                  className={[
                    "h-full rounded-full transition-all duration-300",
                    item.status === "done"
                      ? "bg-emerald-400"
                      : item.status === "error"
                      ? "bg-red-400"
                      : "bg-purple-400",
                  ].join(" ")}
                  style={{ width: `${item.progress}%` }}
                />
              </div>

              {item.errorMessage && (
                <div className="mt-4 rounded-xl border border-red-300/30 bg-red-500/10 p-3 text-sm text-red-100">
                  {item.errorMessage}
                </div>
              )}
            </div>
          ))}
        </section>
      </div>
    </main>
  );
}
