import { defineConfig } from 'tsup';

export default defineConfig({
  entry: { index: 'src/index.ts', ja: 'src/ja.ts', ko: 'src/ko.ts', 'zh-CN': 'src/zh-CN.ts' },
  format: ['esm'],
  dts: true,
  clean: true,
  metafile: true,
  sourcemap: false,
});
