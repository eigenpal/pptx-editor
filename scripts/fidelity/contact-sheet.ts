// Dev aid: tile renders into one PNG for a quick visual pass.
//   bun scripts/fidelity/contact-sheet.ts <out.png> <png...>
import { readFileSync, writeFileSync } from 'node:fs';
import { PNG } from 'pngjs';

const [out, ...files] = process.argv.slice(2);
const cols = 4;
const tw = 480;
const th = 270;
const sheet = new PNG({ width: cols * tw, height: Math.ceil(files.length / cols) * th });
sheet.data.fill(255);
files.forEach((f, i) => {
  const src = PNG.sync.read(readFileSync(f));
  const ox = (i % cols) * tw;
  const oy = Math.floor(i / cols) * th;
  for (let y = 0; y < th - 2; y += 1)
    for (let x = 0; x < tw - 2; x += 1) {
      const sx = Math.floor((x * src.width) / tw);
      const sy = Math.floor((y * src.width) / tw);
      if (sy >= src.height) continue;
      const si = (sy * src.width + sx) * 4;
      const di = ((oy + y) * sheet.width + ox + x) * 4;
      src.data.copy(sheet.data, di, si, si + 4);
    }
});
writeFileSync(out, PNG.sync.write(sheet));
