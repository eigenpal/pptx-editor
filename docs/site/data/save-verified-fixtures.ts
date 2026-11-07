// The save evidence of the feature matrix (`pptx-features.ts`, "Evidence rule"). It stays out of
// the matrix file, which the docs site syncs as site data.

/**
 * Fixtures whose saved copies open with no repair prompt. Both copies from
 * `bun run fidelity:roundtrip -- --saved <dir>` count: `.saved.pptx` (open, save) and
 * `.edited.pptx` (a scripted edit session, then save). Add a name only after you do that check
 * (CONTRIBUTING.md, "Fidelity harness"). Last checked on 6 November 2026: all 44 fixtures,
 * 88 copies written at commit 79a5d22, no prompt. On 7 November 2026: `header-footer`,
 * `hyperlinks`, `preserved-content`, and `svg-tab-stops`, 8 copies, no prompt; and again
 * `header-footer`, with its date and footer placeholders, 2 copies, no prompt.
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
  'header-footer',
  'hyperlinks',
  'line-ends-dashes',
  'line-model',
  'lines-arrows',
  'open-font-families',
  'paragraph-format',
  'picture-editing',
  'pictures',
  'placeholders-background',
  'preserve-unsupported',
  'preserved-content',
  'preset-shapes',
  'rtl-text',
  'shape-effects-presets',
  'shape-styles',
  'slide-layouts',
  'slide-size',
  'slide-size-rounding',
  'slide-structure',
  'spacing',
  'svg-tab-stops',
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
 * copy, and record its fingerprint. Last checked on 7 November 2026: all 66 copies, written by
 * the recipes and fixtures as they are now, no prompt; and the `text.clipboard` copy, no prompt.
 */
export const SAVE_VERIFIED_FEATURE_EDITS: Readonly<Record<string, string>> = {
  'collaboration.automation-api':
    'sha256:dd9192d0997886b39bf62e39c1b3cabeb08decee94e95d6e5cdda95715ecb911',
  'collaboration.realtime':
    'sha256:47e811007d57bbb5b9257d1727c1396f4add81546e85417cc8dd61361aa54189',
  'effects.3d-and-other': 'sha256:7f521bf4e87dd4507133e5a1afa183f8f38797d7575cb807ffc7973f6b462465',
  'effects.fill-overlay': 'sha256:88e9c3cc500e47ff17f120022a184f682c8c08dcbc07ab7d0d5abda814ba2fbd',
  'effects.glow-reflection':
    'sha256:331acbe87981e8c099e725961c34f216079391deeb2a91734ae00e0ea5b2da72',
  'effects.outer-shadow': 'sha256:1ae1086598cd4f1320936df7d43135bba0e27815f23793af204ca6600e10af8e',
  'fills-lines.lines': 'sha256:fb26a1338e4636ad2e9ecb0c7178e877fc1c2ec0a62f10cdd71a145ad089ceda',
  'fills-lines.pattern-picture':
    'sha256:d5d149d4d33512845a4c3b50f801ca1d1f4101d34013960bf3864af22a70de78',
  'fills-lines.solid-gradient':
    'sha256:84ead1d10f7274ca776ab5dbf8995b0fae2a88bc4fbd1a5acff4dbabc8bd0aa2',
  'fills-lines.theme-colors':
    'sha256:937f0feb9baa7944af5c51556bd358d06ad1bf770773841306d7a516c5c3b909',
  'lists.bullets-numbering':
    'sha256:69286419928fadc77b9942ac7b3e57b8cdd782c0c09437ef7b55565ac5d944f3',
  'objects.charts': 'sha256:42f18731591c3a7d930e44bfbcc6ec53ac87e470e4733f21ccdfbc0df1bdfca1',
  'objects.diagrams': 'sha256:aeb9619902390e6835fb90efb2d8f3d942add7d2ef02639a7e1bfc133351f751',
  'objects.equations': 'sha256:24251ac84a6747362dda9cdf121e956733aaef472fa71667b7fe384ea1efa122',
  'objects.media-ole': 'sha256:1bbe774aba3c5a603055c432f37718cd2acbfbd8e24fa7fe2beff9a04b63d5b2',
  'paragraphs.alignment': 'sha256:c0670263a946c3b5caa6e64dc55753b44a30ad25a0b535766ec3b3d86a813360',
  'paragraphs.autofit': 'sha256:b0817fe90eb17090756d5109eaf2860919aced976a527b515411e72e9fe54b64',
  'paragraphs.indent-tabs':
    'sha256:c152515c603aeb89c28ccf35ea4f44a5724a705df5eede526b6a3e99cd9444bc',
  'paragraphs.spacing': 'sha256:b8efc95910eacf7d2a73468df0b0dd532d4cf553cde1dcd745260d2617238204',
  'paragraphs.text-direction-columns':
    'sha256:69312ada59aa7f6baa42d6a5c87700963a38f4a0c3b49b7d9c8d51bd5bdfac31',
  'paragraphs.text-warp': 'sha256:86747641df24f181e08533ae23665665b0558b5d8b05a15a7119540c03edc5a8',
  'pictures.adjustments': 'sha256:e53fe2ca52eb226ff0bfbea3e088063e03298857805bc50859eff19ea382eb2b',
  'pictures.alt-text': 'sha256:0a7c5b5893063a0064b613c44b95813a1f520270c6920decdc414803d45aa6b5',
  'pictures.crop': 'sha256:534571051b0a2130475c373fdbabb78236d32f5558a5c8106fdc616c44c2bf36',
  'pictures.emf-wmf': 'sha256:6508977092fad7f38c6edecf549e0ccb07d922eaa16cffbe358a747e0027798a',
  'pictures.raster-svg': 'sha256:0e894497a8c6909bdc7b4515bd57ffa90ad2e00f5582ca7c78eefee7c3905b5b',
  'pictures.styles-border':
    'sha256:043e5036d5b63116d9ff5d107948c943d23bb309adca4f8a9eaee82137daf636',
  'presenting.animations-transitions':
    'sha256:18faeb28c7810413c87f81817bda47e140f789eadc4563833062ce0981c1cd27',
  'review.comments': 'sha256:f67938b2e1691f8c4005435cbaaaa5d30685c21acf3d82dd96d8418fa651bae1',
  'shapes.arrange': 'sha256:4701cd3ccd6930d3b7f250b7bb632460e869c08904ce4f132c7272c5764ad5f3',
  'shapes.connectors': 'sha256:e734d6ef3ed3218480c862dff622121acdee1f26ca896d8f5c6b4a8895dfbe23',
  'shapes.custom-geometry':
    'sha256:92f1dde2f588f6ae5ed166bf172650e33d7743573d26d47ffe8f2ed4da668081',
  'shapes.flip-rotate': 'sha256:26b871a092abb1e4eeae76b4fb83a42b043d7da8336f81cd177cd919dda79223',
  'shapes.groups': 'sha256:f7ccb22ab864eb29c2ee0bb82a4331b9420c98d5bc3e3615986af6b6e50da807',
  'shapes.preset-geometry':
    'sha256:8438cc2581381fe31dbd2a587efad3f00e9bd67bc11cc32e741a6a8e37bde9d3',
  'shapes.selection-pane':
    'sha256:9e99e23ce91f4fb75343c48c460eb9d545bcc703621c1fee23a533b37af91388',
  'shapes.styles-gallery':
    'sha256:56f55da0a23919445ff463e4400ec23c3bdb33e046dc2606c42c23c9195e7825',
  'slides.header-footer': 'sha256:71adad7795149d7b22c3438b8af207d75a3454e1959e965c9d9d17b93dc0ebbb',
  'slides.hidden': 'sha256:f1c99bf87bd4bfe96988737f06ffc2412e9a260f6b4eebb867f75e5e19cd602c',
  'slides.notes': 'sha256:ec160fc621cbd580462d1e80f07a520364a5a7baa1819b68b84ad291cb2ddfdd',
  'slides.placeholder-prompts':
    'sha256:bc5500d4770b869cdc03744456a34dba357ea17d0207a46550f6e97c8f9a72ef',
  'slides.sections': 'sha256:e708cc676d8f24ca4ee0156233256430664540625fa2dcf43f90df180188cfd3',
  'slides.size': 'sha256:cc819a1b4ad52d1d80284fa75ef418f773bf33f1fe4f6321b002241c4873addc',
  'slides.sorter': 'sha256:82573bad16e9f1b281f31d9feebcd029c3f850b5ab0d266fe65569827f05e567',
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
  'text.clipboard': 'sha256:11f9ad15e9d57168de2ccd5c357bfc34782b4737f93fb1fc7009b496a344efa4',
  'text.color-highlight': 'sha256:cb0c2ab5b21e2b6cef6f2a0204a43a12eda971165cb76470054e883714050632',
  'text.complex-scripts': 'sha256:e402ea570ff05a3119aaa7c0fdea9c60f4ea3f0eea149ca037ea9f1cd02634f1',
  'text.find-replace': 'sha256:9667d4f2087fb86efe2571042ea7c7e1d371e91cb66189b57a0b44750c317c74',
  'text.font-family-size':
    'sha256:4a5c2bc26a3b131d3f39dbb42b4f282c87ad97ebb1b97f7733c509d72ee49acf',
  'text.format-painter': 'sha256:bfc0643f53ce620501692fa93a67daeadae2e18bb2b623651f571329dc85eaf9',
  'text.hyperlinks': 'sha256:dea8f85db1ad583f8388d3d8b6d65008fa5050334226c20f2f77ce36ee18b062',
  'text.input-method': 'sha256:e2dccdc16125bdba88d31c8633ed7bef4f81273109149e0efa8fd728ec01c359',
  'themes.backgrounds': 'sha256:111b73c89fba4487451c7ede83e331f4c7409dab545f65e32c2feabe8230016f',
  'themes.change-layout-theme':
    'sha256:e7525077ce9a9a8b1601314a8a3b53d88c270294211a6d102bf8a50efc68844a',
  'themes.placeholder-inheritance':
    'sha256:3d08c7e8f4ca1472bc6496a8f634a11e1855636683ece38413cf371775e20be4',
};
