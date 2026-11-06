// The save evidence of the feature matrix (`pptx-features.ts`, "Evidence rule"). It stays out of
// the matrix file, which the docs site syncs as site data.

/**
 * Fixtures whose saved copies open with no repair prompt. Both copies from
 * `bun run fidelity:roundtrip -- --saved <dir>` count: `.saved.pptx` (open, save) and
 * `.edited.pptx` (a scripted edit session, then save). Add a name only after you do that check
 * (CONTRIBUTING.md, "Fidelity harness"). Last checked on 6 November 2026: all 44 fixtures,
 * 88 copies written at commit 79a5d22, no prompt.
 */
export const SAVE_VERIFIED_FIXTURES: readonly string[] = [
  'alignment-anchor',
  'arrange-flip-hidden',
  'autofit',
  'bullets-gallery',
  'bullets-numbering',
  'chart-sizes',
  'charts',
  'colors',
  'custom-geometry',
  'diagrams',
  'east-asian-line-breaking',
  'east-asian-theme-fonts',
  'east-asian-typography',
  'effects',
  'fill-fidelity',
  'fills-lines',
  'follow-ups',
  'font-effects',
  'gradients',
  'groups-transforms',
  'line-ends-dashes',
  'line-model',
  'lines-arrows',
  'open-font-families',
  'paragraph-format',
  'picture-editing',
  'pictures',
  'placeholders-background',
  'preserve-unsupported',
  'preset-shapes',
  'rtl-text',
  'shape-effects-presets',
  'shape-styles',
  'slide-layouts',
  'slide-size',
  'slide-size-rounding',
  'slide-structure',
  'spacing',
  'table-editing',
  'tables',
  'text-direction',
  'text-runs',
  'vertical-columns',
  'vertical-east-asian',
];

/**
 * Feature rows whose own edit (`FEATURE_EDITS` in `scripts/fidelity/feature-edits.ts`) writes a
 * copy that opens with no repair prompt, with the fingerprint of that copy (`editFingerprint`).
 * `bun scripts/fidelity/feature-edits.ts --out <dir>` writes one copy per row. Add a row, or
 * change its fingerprint, only after you open its copy with no prompt. `feature-edits.test.ts`
 * fails when a recipe or an engine change gives a copy with another fingerprint: then open the new
 * copy, and record its fingerprint. Last checked on 6 November 2026: 45 copies, no prompt.
 */
export const SAVE_VERIFIED_FEATURE_EDITS: Readonly<Record<string, string>> = {
  'effects.glow-reflection':
    'sha256:655b261b22d2d3b6cbebb2e15a360430a2c835772c9ffdf4de71d34c4d73353a',
  'effects.outer-shadow': 'sha256:8651c930e6b8cede989b44bfdf051e2c292f11213605aa9771b3675ff27dfc08',
  'fills-lines.lines': 'sha256:787dad09a0f940e0bd143f799958b0ce31bdcdb4807828e94266611b2396a103',
  'fills-lines.solid-gradient':
    'sha256:52fc8eebacb29a7b1086908ac2c5ee7c99490f84f6b2117935627241a3ad3566',
  'fills-lines.theme-colors':
    'sha256:14e9e99471e58c36bd5ab4481fb6ff38f33bfd5de7f7559caeb513b3373900ec',
  'lists.bullets-numbering':
    'sha256:82ab4c653f2a81526c327d61dc6fbb98846ff7d7693f6d5e1074b1c89258962e',
  'paragraphs.alignment': 'sha256:a2b952bd4f79849bd99ce77bb38ea05ad35dd158c07896385625ae868bb324f5',
  'paragraphs.autofit': 'sha256:9be678219413b8be4dadb5ad35eed1c80cee89ea3e4ec535198face6fac28dbf',
  'paragraphs.indent-tabs':
    'sha256:c3bab3b1d66b1af5ed1a0d8b92261c6ad77bacdd8d9bdfe8c39bd61bc4b52736',
  'paragraphs.spacing': 'sha256:6358ec2eab1217db508ea5a61b2e728ba06de9c27a8c30e42d5ddae7f9078e94',
  'paragraphs.text-direction-columns':
    'sha256:85357d91657c57161e624dec1c86d06643ee39f9a9cde4017a2e055b2a2e37ee',
  'pictures.adjustments': 'sha256:596e6467e01c5f0286e6ceb90fd66cd95dd55a37ec52d3d950abb57b03706bf9',
  'pictures.alt-text': 'sha256:37947c4e162743f5798cf9111fc9cc72b9d155d827590a06814df41e13d8f554',
  'pictures.crop': 'sha256:a18b3ec0e363b332cb402aa892c730ec75a9dfd1a10b2e07724d1f39aea18777',
  'pictures.raster-svg': 'sha256:249272729a3eb511481f25dd73afff6d447800d42229287e801c751f81ddb21c',
  'pictures.styles-border':
    'sha256:3b6fd20b71fdab3424e195e60d7ead6eb2d3ab976b5c9b0bdc3442629a066f3b',
  'shapes.arrange': 'sha256:b6d005ab323fcd0bc0da39d4b8779d5573edb0b3e7d6931b7586a645e9af8a32',
  'shapes.connectors': 'sha256:e83b762304c223bade52fada639b48f16b17fbbd4127add4e48838f0761ec8d1',
  'shapes.flip-rotate': 'sha256:fb2698993db1441daf928c0c46416f1a6faab63d078ca79bc8aef098a7b40ffa',
  'shapes.groups': 'sha256:c6ececd8c9dda3c326e1960ba7df3c5fda2cb64c2ba2d9210b41dcc4b2f7bfec',
  'shapes.preset-geometry':
    'sha256:9bd321bc3ac2f9c69e4d690e56582ca65ae40445dde321fc9e75bf3618cc9c75',
  'shapes.selection-pane':
    'sha256:c9e83030dc9272a4e47d5552dd4fb118ab555399f7394874d82d27e220f856fc',
  'shapes.styles-gallery':
    'sha256:f42df5f28dc5e386b402908a8ec95df12c11bd78376e4fd296619c832afd6698',
  'slides.hidden': 'sha256:f6cb6d24689ab65129ebec4e3dc1aa14de665aad95385517e9ee874d50ea8b7a',
  'slides.notes': 'sha256:31904b1462eb6f74f45ef7d94789ef3f3a2a6cf4fde15778267089d2a5075cc6',
  'slides.placeholder-prompts':
    'sha256:943c07e54088874955d84c22563da8876268aac42c21fefb710f40182fe9fee5',
  'slides.sections': 'sha256:9119fd3502a5d0a70245d585aac314135b5c91681f728f4d4324d7a945d8d0d5',
  'slides.structure': 'sha256:93506297d38221820fc0b172144f3850b1ab553898fc5f368002f9b827317c54',
  'tables.cell-format': 'sha256:a94bf7763c9f6827d3ebcc3db8a828e4bdba817dd6f2f81643ae91073abab0ae',
  'tables.cell-text': 'sha256:440730390c55baa61dd1d377d9c2634d66b7b632f2d427b3dd5bac4f62c6e5d7',
  'tables.merge-split': 'sha256:c77c79314a24c597f9097164d51949c49e30b7282805a599570af7ec780c3091',
  'tables.render': 'sha256:42fe54846c40ad46fe020648f8d06fa8cdf363425c6affe8d10b2dc6e5aa9c8f',
  'tables.sizes': 'sha256:3cee788dd2a0c01710466907397ef70e996c22b1dd175ecb98dd5c794cf14734',
  'tables.structure': 'sha256:76c7616db70899b583002877c13ed8728fc56e286dc22e69579c52842b08aba3',
  'tables.styles': 'sha256:ee73f04a23be3496825087d9e5b16a49fd018ebe511806f04852f1a1aaf44cb3',
  'text.basic-formatting':
    'sha256:21d398a34f42eda867476ff920ca7c063084f03119faf9cf945a02704cdb1488',
  'text.caps-baseline': 'sha256:f66b27ef0923ee5e0d3be66280b3cfb8209ab4989f4c095ae0c09c19a50c017a',
  'text.change-case': 'sha256:14a940552b0478a26a6b73ad79b667b651e363e2a643a155d7bcca72e5aa70a4',
  'text.character-spacing':
    'sha256:0378eab2a1dc5e470606b220707e63721b1c345e3598ab96445527165afce17b',
  'text.color-highlight': 'sha256:17aaecca43044d0f5419915b0057b76d19bb6efb75f2d95324c648d321f9c5dd',
  'text.complex-scripts': 'sha256:2b121a41e8d8ba37ad157f716d9678b588bf8142504772ba10ee0f94b12a4ae2',
  'text.font-family-size':
    'sha256:27f17f5a54b3bd06a4963dcc58ac124743f5b13d2e3a0a93ed298e20da2be546',
  'text.input-method': 'sha256:f5c6ed12a66e6a63442fa0535329c5d069a13bd2078d5c909803c64a6db2f3ca',
  'themes.backgrounds': 'sha256:1a52919c4717146b3a0cdfa495f6f9bec240023cd4262504154965e131c9df35',
  'themes.change-layout-theme':
    'sha256:0a9a7a463347339f54cb53a10cace58e9f189bfac41f2b65cddeee16adf988d1',
};
