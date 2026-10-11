// Render slides with an explicit CJK font fallback (the deck's "Inter" font is
// not installed on this machine, so LibreOffice otherwise picks a garbled face).
// usage: node tools_render_fallback.mjs <input.pptx> <outDir> <pages>
import { createConverter } from "file:///D:/deepseek-harnmes/resources/app.asar.unpacked/dsh/node_modules/@deepseek-ai/libreoffice-kit/lib/index.js";

const [input, outDir, pages] = process.argv.slice(2);
const converter = await createConverter({
  timeoutMs: 300000,
  fontFallbacks: [
    ["Inter", "Microsoft YaHei"],
    ["Calibri", "Microsoft YaHei"],
  ],
});
try {
  const result = await converter.renderImages({
    inputPath: input,
    outputDir: outDir,
    pages: pages ? pages.split(",").map(Number) : undefined,
    dpi: 110,
  });
  console.log(JSON.stringify({
    backend: result.backend,
    missingFonts: result.missingFonts,
    images: result.images.map((i) => ({ page: i.page, path: i.path })),
  }));
} finally {
  await converter.dispose();
}
