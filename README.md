# EPUB para DOCX

Conversor pessoal de EPUB para DOCX feito com Next.js. A interface realiza a conversão no navegador; a rota `POST /api/convert` está disponível para integrações.

## Requisitos e uso

```bash
npm install
npm run dev
```

Abra `http://localhost:3000`, selecione até 30 arquivos EPUB (máximo de 25 MB cada) e faça o download individual ou em ZIP.

## Fidelidade da conversão

O conversor usa a ordem do `spine` do EPUB e preserva parágrafos, listas, títulos, quebras de linha, estilos básicos (negrito, itálico, sublinhado, sobrescrito e subscrito) e quebras de página declaradas no XHTML. Imagens internas do EPUB são incorporadas; imagens remotas são propositalmente ignoradas para impedir requisições externas a partir de arquivos enviados.

EPUB é um formato de publicação web, e recursos como tabelas complexas, fontes incorporadas, CSS avançado e scripts não possuem equivalência perfeita em DOCX.

## Validação

```bash
npm run typecheck
npm test
npm run build
```
