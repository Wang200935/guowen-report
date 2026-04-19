const fs = await import("node:fs/promises");
const path = await import("node:path");
const { Presentation, PresentationFile } = await import("@oai/artifact-tool");

const W = 1280;
const H = 720;
const ROOT = path.resolve("/Users/wang/Documents/CTF/guowen-report/final");
const PREVIEW_DIR = path.join(ROOT, "readme-slides");

async function readImageBlob(imagePath) {
  const bytes = await fs.readFile(imagePath);
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
}

const presentation = Presentation.create({ slideSize: { width: W, height: H } });

for (let i = 1; i <= 9; i += 1) {
  const slide = presentation.slides.add();
  const imagePath = path.join(PREVIEW_DIR, `slide-${String(i).padStart(2, "0")}.png`);
  const image = slide.images.add({
    blob: await readImageBlob(imagePath),
    fit: "cover",
    alt: `Canva slide ${i}`,
  });
  image.position = { left: 0, top: 0, width: W, height: H };
}

const pptx = await PresentationFile.exportPptx(presentation);
await pptx.save(path.join(ROOT, "臺南城市變遷歷史簡報.pptx"));
console.log(path.join(ROOT, "臺南城市變遷歷史簡報.pptx"));
