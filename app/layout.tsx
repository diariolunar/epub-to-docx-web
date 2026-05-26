import "./globals.css";

export const metadata = {
  title: "EPUB para DOCX",
  description: "Conversor pessoal de EPUB para DOCX",
};

export default function RootLayout({
  children,
}: {
  children: React.ReactNode;
}) {
  return (
    <html lang="pt-BR">
      <body>{children}</body>
    </html>
  );
}