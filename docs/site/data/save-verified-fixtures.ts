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
 * copy, and record its fingerprint. Last checked on 6 November 2026: 45 copies written at
 * commit 74a4186, no prompt.
 */
export const SAVE_VERIFIED_FEATURE_EDITS: Readonly<Record<string, string>> = {
  'effects.glow-reflection':
    'sha256:331acbe87981e8c099e725961c34f216079391deeb2a91734ae00e0ea5b2da72',
  'effects.outer-shadow': 'sha256:1ae1086598cd4f1320936df7d43135bba0e27815f23793af204ca6600e10af8e',
  'fills-lines.lines': 'sha256:fb26a1338e4636ad2e9ecb0c7178e877fc1c2ec0a62f10cdd71a145ad089ceda',
  'fills-lines.solid-gradient':
    'sha256:84ead1d10f7274ca776ab5dbf8995b0fae2a88bc4fbd1a5acff4dbabc8bd0aa2',
  'fills-lines.theme-colors':
    'sha256:937f0feb9baa7944af5c51556bd358d06ad1bf770773841306d7a516c5c3b909',
  'lists.bullets-numbering':
    'sha256:69286419928fadc77b9942ac7b3e57b8cdd782c0c09437ef7b55565ac5d944f3',
  'paragraphs.alignment': 'sha256:c0670263a946c3b5caa6e64dc55753b44a30ad25a0b535766ec3b3d86a813360',
  'paragraphs.autofit': 'sha256:b0817fe90eb17090756d5109eaf2860919aced976a527b515411e72e9fe54b64',
  'paragraphs.indent-tabs':
    'sha256:d089f943a4298e5b7c7a86d7494f15da7e8115532f31b3e18fbbc12326818f2c',
  'paragraphs.spacing': 'sha256:b8efc95910eacf7d2a73468df0b0dd532d4cf553cde1dcd745260d2617238204',
  'paragraphs.text-direction-columns':
    'sha256:69312ada59aa7f6baa42d6a5c87700963a38f4a0c3b49b7d9c8d51bd5bdfac31',
  'pictures.adjustments': 'sha256:e53fe2ca52eb226ff0bfbea3e088063e03298857805bc50859eff19ea382eb2b',
  'pictures.alt-text': 'sha256:0a7c5b5893063a0064b613c44b95813a1f520270c6920decdc414803d45aa6b5',
  'pictures.crop': 'sha256:534571051b0a2130475c373fdbabb78236d32f5558a5c8106fdc616c44c2bf36',
  'pictures.raster-svg': 'sha256:83f5cb08254b0d81af8626c9a4fa8bdba6cea6d63e2d35d390284ca6047dfadb',
  'pictures.styles-border':
    'sha256:043e5036d5b63116d9ff5d107948c943d23bb309adca4f8a9eaee82137daf636',
  'shapes.arrange': 'sha256:4701cd3ccd6930d3b7f250b7bb632460e869c08904ce4f132c7272c5764ad5f3',
  'shapes.connectors': 'sha256:e734d6ef3ed3218480c862dff622121acdee1f26ca896d8f5c6b4a8895dfbe23',
  'shapes.flip-rotate': 'sha256:26b871a092abb1e4eeae76b4fb83a42b043d7da8336f81cd177cd919dda79223',
  'shapes.groups': 'sha256:f7ccb22ab864eb29c2ee0bb82a4331b9420c98d5bc3e3615986af6b6e50da807',
  'shapes.preset-geometry':
    'sha256:8438cc2581381fe31dbd2a587efad3f00e9bd67bc11cc32e741a6a8e37bde9d3',
  'shapes.selection-pane':
    'sha256:9e99e23ce91f4fb75343c48c460eb9d545bcc703621c1fee23a533b37af91388',
  'shapes.styles-gallery':
    'sha256:56f55da0a23919445ff463e4400ec23c3bdb33e046dc2606c42c23c9195e7825',
  'slides.hidden': 'sha256:f1c99bf87bd4bfe96988737f06ffc2412e9a260f6b4eebb867f75e5e19cd602c',
  'slides.notes': 'sha256:ec160fc621cbd580462d1e80f07a520364a5a7baa1819b68b84ad291cb2ddfdd',
  'slides.placeholder-prompts':
    'sha256:bc5500d4770b869cdc03744456a34dba357ea17d0207a46550f6e97c8f9a72ef',
  'slides.sections': 'sha256:e708cc676d8f24ca4ee0156233256430664540625fa2dcf43f90df180188cfd3',
  'slides.structure': 'sha256:4c94fe29b0f69573b04eb4a5cee860216e25d65ca035e5b963dac749828ad02d',
  'tables.cell-format': 'sha256:7d0f4cf1b2a0be903a54f7aa0cb2a9652540899a0c74ca077e913606af3c9133',
  'tables.cell-text': 'sha256:ac550e68680e930832758db43b0de7b44479aa8f518ee84928d769d8c4df06a2',
  'tables.merge-split': 'sha256:c2666f07264d1dda79fda8fb69a5d1f4a423d2030a3e0b5e53d4bcb306487af0',
  'tables.render': 'sha256:63586b485ff47bdc3ab6e4d5570230ac3a62519d93aacb534bff79b1d9d53894',
  'tables.sizes': 'sha256:7ea5cf467e47bb36b99ae768abdf86723364d1bd2e416f53e0a91784d3a5ed89',
  'tables.structure': 'sha256:198657d9a06f331066fa103e15feb9ee8617baa9c40cbfe8392180528ba7928b',
  'tables.styles': 'sha256:facf8b18bf0bc5e14d83dd0bcd514e1d9e8f45bc4bcdc746d56c414bd47c5156',
  'text.basic-formatting':
    'sha256:3fa607e2f582845931d672982176c2a38a68c2469236846bb4681959aadfe411',
  'text.caps-baseline': 'sha256:79a6ad7c035e0b90b97f21f908c49f79f5e4046a0bd11094e50ae7c73073c5c9',
  'text.change-case': 'sha256:73f0bc49d5856e5bfdbae4ae90815e40f613dfe8f53c50bb19b964d84441208b',
  'text.character-spacing':
    'sha256:702aff38ffdf6a441aabc4c8407b18d5a9cb62f058a56523101ef47614d5cf20',
  'text.color-highlight': 'sha256:cb0c2ab5b21e2b6cef6f2a0204a43a12eda971165cb76470054e883714050632',
  'text.complex-scripts': 'sha256:e402ea570ff05a3119aaa7c0fdea9c60f4ea3f0eea149ca037ea9f1cd02634f1',
  'text.font-family-size':
    'sha256:4a5c2bc26a3b131d3f39dbb42b4f282c87ad97ebb1b97f7733c509d72ee49acf',
  'text.input-method': 'sha256:e2dccdc16125bdba88d31c8633ed7bef4f81273109149e0efa8fd728ec01c359',
  'themes.backgrounds': 'sha256:111b73c89fba4487451c7ede83e331f4c7409dab545f65e32c2feabe8230016f',
  'themes.change-layout-theme':
    'sha256:87ab840601bce4efe3da5dc5c4ac7dedfa22ff14d3c240df0d8bf7169978d4ac',
};
