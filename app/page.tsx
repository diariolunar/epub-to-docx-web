"use client";

import { useState } from "react";
import JSZip from "jszip";

type FileItem = {
  file: File;
  status: "waiting" | "converting" | "done" | "error";
  progress: number;
  downloadUrl?: string;
  docxBlob?: Blob;
};

export default function Home() {
  const [files, setFiles] = useState<FileItem[]>([]);
  const [isConverting, setIsConverting] = useState(false);

  function handleFiles(selected: FileList | null) {
    if (!selected) return;

    const chosen = Array.from(selected).slice(0, 30);

    setFiles(
      chosen.map((file) => ({
        file,
        status: "waiting",
        progress: 0,
      }))
    );
  }

  async function convertAll() {
    setIsConverting(true);
    const updated = [...files];

    for (let i = 0; i < updated.length; i++) {
      updated[i].status = "converting";
      updated[i].progress = 15;
      setFiles([...updated]);

      try {
        const formData = new FormData();
        formData.append("file", updated[i].file);

        updated[i].progress = 45;
        setFiles([...updated]);

        const response = await fetch("/api/convert", {
          method: "POST",
          body: formData,
        });

        if (!response.ok) throw new Error("Erro na conversão");

        updated[i].progress = 80;
        setFiles([...updated]);

        const blob = await response.blob();
        const url = URL.createObjectURL(blob);

        updated[i].docxBlob = blob;
        updated[i].downloadUrl = url;
        updated[i].status = "done";
        updated[i].progress = 100;

        setFiles([...updated]);
      } catch {
        updated[i].status = "error";
        updated[i].progress = 0;
        setFiles([...updated]);
      }
    }

    setIsConverting(false);
  }

  async function downloadAllZip() {
    const zip = new JSZip();

    files.forEach((item) => {
      if (item.docxBlob) {
        const name = item.file.name.replace(/\.epub$/i, ".docx");
        zip.file(name, item.docxBlob);
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

  const doneCount = files.filter((item) => item.status === "done").length;
  const canDownloadZip = doneCount > 0;

  return (
    <main
      style={{
        maxWidth: 900,
        margin: "40px auto",
        padding: 20,
        fontFamily: "Arial",
      }}
    >
      <h1>EPUB → DOCX</h1>

      <p>Selecione até 30 arquivos EPUB.</p>

      <input
        type="file"
        accept=".epub"
        multiple
        onChange={(e) => handleFiles(e.target.files)}
      />

      <div style={{ marginTop: 20, display: "flex", gap: 10 }}>
        <button
          onClick={convertAll}
          disabled={files.length === 0 || isConverting}
        >
          {isConverting ? "Convertendo..." : "Converter Todos"}
        </button>

        <button
          onClick={downloadAllZip}
          disabled={!canDownloadZip || isConverting}
        >
          Baixar Todos (.zip)
        </button>
      </div>

      <p style={{ marginTop: 15 }}>
        Concluídos: {doneCount}/{files.length}
      </p>

      <div style={{ marginTop: 30 }}>
        {files.map((item, index) => (
          <div
            key={index}
            style={{
              border: "1px solid #ccc",
              borderRadius: 10,
              padding: 15,
              marginBottom: 15,
            }}
          >
            <strong>{item.file.name}</strong>

            <div
              style={{
                width: "100%",
                height: 12,
                background: "#eee",
                borderRadius: 999,
                overflow: "hidden",
                marginTop: 10,
              }}
            >
              <div
                style={{
                  width: `${item.progress}%`,
                  height: "100%",
                  background:
                    item.status === "done"
                      ? "green"
                      : item.status === "error"
                      ? "red"
                      : "#0070f3",
                  transition: "0.3s",
                }}
              />
            </div>

            <div style={{ marginTop: 10 }}>
              Status: {item.status}
            </div>

            {item.downloadUrl && (
              <a
                href={item.downloadUrl}
                download={item.file.name.replace(/\.epub$/i, ".docx")}
              >
                <button style={{ marginTop: 10 }}>Baixar</button>
              </a>
            )}
          </div>
        ))}
      </div>
    </main>
  );
}