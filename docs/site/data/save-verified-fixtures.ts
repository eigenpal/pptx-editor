// The save evidence of the feature matrix (`pptx-features.ts`, "Evidence rule"). It stays out of
// the matrix file, which the docs site syncs as site data.

/**
 * Fixtures whose saved copies open with no repair prompt. Both copies from
 * `bun run fidelity:roundtrip -- --saved <dir>` count: `.saved.pptx` (open, save) and
 * `.edited.pptx` (a scripted edit session, then save). Add a name only after you do that check
 * (CONTRIBUTING.md, "Fidelity harness"). Last checked on 6 November 2026: all 44 fixtures,
 * 88 copies written at commit 79a5d22, no prompt. On 7 November 2026: `header-footer`,
 * `hyperlinks`, `preserved-content`, and `svg-tab-stops`, 8 copies, no prompt; and again
 * `header-footer`, with its date and footer placeholders, 2 copies, no prompt;
 * `picture-bullets`, 2 copies, no prompt; `text-warp`, 2 copies, no prompt; and
 * `chart-data`, 2 copies, no prompt; `three-d-effects`, 2 copies, no prompt;
 * `diagram-editing`, 2 copies, no prompt; `equations`, 2 copies, no prompt; and
 * `three-d-turned-effects`, 2 copies and a copy with bevels set and removed on theme-styled
 * shapes, no prompt. On 8 November 2026: `equation-lines`, 2 copies, no prompt.
 */
export const SAVE_VERIFIED_FIXTURES: readonly string[] = [
  'alignment-anchor',
  'arrange-flip-hidden',
  'autofit',
  'bullets-gallery',
  'bullets-numbering',
  'chart-data',
  'chart-sizes',
  'charts',
  'colors',
  'custom-geometry',
  'diagram-editing',
  'diagrams',
  'east-asian-line-breaking',
  'east-asian-theme-fonts',
  'east-asian-typography',
  'effects',
  'equation-lines',
  'equations',
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
  'picture-bullets',
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
  'text-warp',
  'three-d-effects',
  'three-d-turned-effects',
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
 * the recipes and fixtures as they are now, no prompt; the `text.clipboard` copy, no prompt; and again
 * `presenting.animations-transitions`, with its new effects and transition,
 * `lists.bullets-numbering` and `text.hyperlinks`, with picture bullets and custom show and file
 * links, and the new `objects.charts` copy, which edits chart data, no prompt; and the new
 * `effects.3d-and-other` copy, which sets and clears bevels, cameras, and blur, the
 * `objects.diagrams` copy, which types in, adds, and deletes diagram shapes, and
 * `review.comments` and `slides.structure`, with comment parts saved as XML trees, no prompt.
 * On 8 November 2026: the new `objects.equations` copy, whose inserted equation box is shorter
 * after the change to equation line heights, no prompt. On 13 November 2026: the
 * `objects.equations` copy after the change to equation line spacing, no prompt.
 */
export const SAVE_VERIFIED_FEATURE_EDITS: Readonly<Record<string, string>> = {
  'collaboration.automation-api':
    'sha256:dd9192d0997886b39bf62e39c1b3cabeb08decee94e95d6e5cdda95715ecb911',
  'collaboration.realtime':
    'sha256:47e811007d57bbb5b9257d1727c1396f4add81546e85417cc8dd61361aa54189',
  'effects.3d-and-other': 'sha256:e97f4075d4485c849ef3dba4419a7593ebda6ddeaefcdc1dabb093eb21703f6a',
  'effects.fill-overlay': 'sha256:1ca1da39ba290ed51260330eea20da74df48a1fa2f57d38db01ed46e6ee2a740',
  'effects.glow-reflection':
    'sha256:331acbe87981e8c099e725961c34f216079391deeb2a91734ae00e0ea5b2da72',
  'effects.outer-shadow': 'sha256:1ae1086598cd4f1320936df7d43135bba0e27815f23793af204ca6600e10af8e',
  'fills-lines.lines': 'sha256:fb26a1338e4636ad2e9ecb0c7178e877fc1c2ec0a62f10cdd71a145ad089ceda',
  'fills-lines.pattern-picture':
    'sha256:d27f45aa49306bd4a5df38e239b65bff5cd8e83dec4bf5c1b80fa51bc136ae2b',
  'fills-lines.solid-gradient':
    'sha256:6fccb555a3f2253b2eae428c3b63462a6d9f51ab67c08d2c675e2b28e9e9e06f',
  'fills-lines.theme-colors':
    'sha256:937f0feb9baa7944af5c51556bd358d06ad1bf770773841306d7a516c5c3b909',
  'lists.bullets-numbering':
    'sha256:37e710090974cae1ffd0a0fb047f4ee7ffeffd845c72ffd4acd6bdd4b49be27f',
  'objects.charts': 'sha256:82b50397e0ff8613b6198d351393aa7e6953924f39a5064b18e6a909b412e1d8',
  'objects.diagrams': 'sha256:376fb84db5305fe825baeb1e64f7dc9cd7b0418fce156bda3238cf55498b3e8b',
  'objects.equations': 'sha256:41fb0b674839ee40af07cf38090533cec20e011c043ae95430db039ae952d773',
  'objects.media-ole': 'sha256:1bbe774aba3c5a603055c432f37718cd2acbfbd8e24fa7fe2beff9a04b63d5b2',
  'paragraphs.alignment': 'sha256:c0670263a946c3b5caa6e64dc55753b44a30ad25a0b535766ec3b3d86a813360',
  'paragraphs.autofit': 'sha256:b0817fe90eb17090756d5109eaf2860919aced976a527b515411e72e9fe54b64',
  'paragraphs.indent-tabs':
    'sha256:c152515c603aeb89c28ccf35ea4f44a5724a705df5eede526b6a3e99cd9444bc',
  'paragraphs.spacing': 'sha256:b8efc95910eacf7d2a73468df0b0dd532d4cf553cde1dcd745260d2617238204',
  'paragraphs.text-direction-columns':
    'sha256:69312ada59aa7f6baa42d6a5c87700963a38f4a0c3b49b7d9c8d51bd5bdfac31',
  'paragraphs.text-warp': 'sha256:f2e48a000435c153a7e2e154a5e71e30e54bcc06bfc475cd78bf71e2c2b5ddb2',
  'pictures.adjustments': 'sha256:e53fe2ca52eb226ff0bfbea3e088063e03298857805bc50859eff19ea382eb2b',
  'pictures.alt-text': 'sha256:0a7c5b5893063a0064b613c44b95813a1f520270c6920decdc414803d45aa6b5',
  'pictures.crop': 'sha256:534571051b0a2130475c373fdbabb78236d32f5558a5c8106fdc616c44c2bf36',
  'pictures.emf-wmf': 'sha256:6508977092fad7f38c6edecf549e0ccb07d922eaa16cffbe358a747e0027798a',
  'pictures.raster-svg': 'sha256:0e894497a8c6909bdc7b4515bd57ffa90ad2e00f5582ca7c78eefee7c3905b5b',
  'pictures.styles-border':
    'sha256:043e5036d5b63116d9ff5d107948c943d23bb309adca4f8a9eaee82137daf636',
  'presenting.animations-transitions':
    'sha256:11fee431d75edb3255ee2e5fe61cfae95b23496c194e581df535033fe80c243d',
  'review.comments': 'sha256:2823b6a3277894e25e5d9ae6519d032e276acc6416ed89f2188ba1387ce7629a',
  'shapes.arrange': 'sha256:4701cd3ccd6930d3b7f250b7bb632460e869c08904ce4f132c7272c5764ad5f3',
  'shapes.connectors': 'sha256:570a05eba2d398edf54a8d0dae99babf4bed618847e96fe349c8876c63e619a8',
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
  'slides.structure': 'sha256:a87a205b99be3115c5d45bb1e0f9c1249e96392c6561d7ac126cd966cc2a1e98',
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
  'text.clipboard': 'sha256:088d7127d31eabe7de43db0c01347ead554fb6bf655a56f812f944c3b9e68f1b',
  'text.color-highlight': 'sha256:cb0c2ab5b21e2b6cef6f2a0204a43a12eda971165cb76470054e883714050632',
  'text.complex-scripts': 'sha256:e402ea570ff05a3119aaa7c0fdea9c60f4ea3f0eea149ca037ea9f1cd02634f1',
  'text.find-replace': 'sha256:9667d4f2087fb86efe2571042ea7c7e1d371e91cb66189b57a0b44750c317c74',
  'text.font-family-size':
    'sha256:6a8534838c1ae38d136c6d80539f8a0b5f73f8df876318d41e4b265205c6656a',
  'text.format-painter': 'sha256:bfc0643f53ce620501692fa93a67daeadae2e18bb2b623651f571329dc85eaf9',
  'text.hyperlinks': 'sha256:3a55fd374282395cdefeb2b5f3f0b58f41fcedf64fc53657d000b5112663702c',
  'text.input-method': 'sha256:e2dccdc16125bdba88d31c8633ed7bef4f81273109149e0efa8fd728ec01c359',
  'themes.backgrounds': 'sha256:111b73c89fba4487451c7ede83e331f4c7409dab545f65e32c2feabe8230016f',
  'themes.change-layout-theme':
    'sha256:e7525077ce9a9a8b1601314a8a3b53d88c270294211a6d102bf8a50efc68844a',
  'themes.placeholder-inheritance':
    'sha256:3d08c7e8f4ca1472bc6496a8f634a11e1855636683ece38413cf371775e20be4',
};

/**
 * Structural edits of kept content (`KEPT_CONTENT_EDITS` in
 * `scripts/fidelity/kept-content-edits.ts`) whose copy opens with no repair prompt, with the
 * fingerprint of that copy (`editFingerprint`). `bun scripts/fidelity/kept-content-edits.ts --out
 * <dir>` writes one copy per recipe. The rules of `SAVE_VERIFIED_FEATURE_EDITS` apply. Last
 * checked on 7 November 2026: all 37 copies, no prompt. The animation and comment rows now
 * prove a full round trip elsewhere, so their duplicate-slide copies are no longer listed.
 */
export const SAVE_VERIFIED_KEPT_CONTENT_EDITS: Readonly<Record<string, string>> = {
  'shapes.custom-geometry: duplicate slide':
    'sha256:02ce251d16a85cfc638440adbf1bbb8cec92c582bf2b575f5b3e4a3a250b9712',
  'shapes.custom-geometry: paste on a new slide':
    'sha256:ed71ac210577fbece88d29bd8b17f96c2ee0115c13f26f6050f03dcdf0a6b9ed',
  'shapes.custom-geometry: paste into another deck':
    'sha256:ed71ac210577fbece88d29bd8b17f96c2ee0115c13f26f6050f03dcdf0a6b9ed',
  'effects.fill-overlay: duplicate slide':
    'sha256:4114727dae3d4d6acee03c2907a301816f66295fe175105d830e21917be7007b',
  'effects.fill-overlay: paste on a new slide':
    'sha256:20b1b3134391d24065667068ac056f9ed874d88bbf0211c9b5c1cdf46d7e74f5',
  'effects.fill-overlay: paste into another deck':
    'sha256:20b1b3134391d24065667068ac056f9ed874d88bbf0211c9b5c1cdf46d7e74f5',
  'effects.fill-overlay: group':
    'sha256:1412f5d06e757bcc1a282133b90a102b69f7b7b649ac828fdf0d54d7fecb4111',
  'effects.3d-and-other: duplicate slide':
    'sha256:c94c33372ad368532f3e7b50ee56133ff4b414094edfcd7772adfafb31bb2bae',
  'effects.3d-and-other: paste on a new slide':
    'sha256:a2880f93ecc8d212720b17382fbd32302aeab5707e0d478aa9a7e4392eaba416',
  'effects.3d-and-other: paste into another deck':
    'sha256:a2880f93ecc8d212720b17382fbd32302aeab5707e0d478aa9a7e4392eaba416',
  'effects.3d-and-other: group':
    'sha256:d26375a04dfa05343b2a9841ffe5922492958405fccec095150e09e11df5f240',
  'pictures.emf-wmf: duplicate slide':
    'sha256:c94c33372ad368532f3e7b50ee56133ff4b414094edfcd7772adfafb31bb2bae',
  'pictures.emf-wmf: paste on a new slide':
    'sha256:6c3733d72b712fd5ec4684d56ded87f97b93cab8d3ea8f6c816a20333b53bc24',
  'pictures.emf-wmf: paste into another deck':
    'sha256:6c3733d72b712fd5ec4684d56ded87f97b93cab8d3ea8f6c816a20333b53bc24',
  'pictures.emf-wmf: group':
    'sha256:5ec05cd5e84f0601e2507e92f273ddd1ea1b948cbb49c88ef8b00e8b3ed0b70a',
  'objects.charts: duplicate slide':
    'sha256:977244faafd42525ffcb576a4d92ca819994adb9b342dab99246949a3811ec33',
  'objects.charts: paste on a new slide':
    'sha256:34165a3d642dac6ab3a35045c8a50326e7971276839e22d2df9a700fc95a293e',
  'objects.charts: paste into another deck':
    'sha256:34165a3d642dac6ab3a35045c8a50326e7971276839e22d2df9a700fc95a293e',
  'objects.charts: group':
    'sha256:a2b8b462f8e242ec8082cf70a2c3a6a416fef831b6d9327061baa37523355a6c',
  'objects.diagrams: duplicate slide':
    'sha256:ebba58655e8431d3d979e488f181abfcc76b90be01c40fe7a74089899fc064fd',
  'objects.diagrams: paste on a new slide':
    'sha256:6bef0632f1f598ad61b46412c4c3644734ffcd9b2882baa70ee41c47c9142adb',
  'objects.diagrams: paste into another deck':
    'sha256:6bef0632f1f598ad61b46412c4c3644734ffcd9b2882baa70ee41c47c9142adb',
  'objects.diagrams: group':
    'sha256:24b86ac67a2769552684b1f291df64d9ad956706c76d9d6f0d0790b05d92deae',
  'objects.media-ole: duplicate slide':
    'sha256:c94c33372ad368532f3e7b50ee56133ff4b414094edfcd7772adfafb31bb2bae',
  'objects.media-ole: paste on a new slide':
    'sha256:911bba4af59888e1c9693fb835de131740cce0965943fa7ac1389510c59c35c8',
  'objects.media-ole: paste into another deck':
    'sha256:911bba4af59888e1c9693fb835de131740cce0965943fa7ac1389510c59c35c8',
  'objects.media-ole: group':
    'sha256:2be7e4da1df23208d258a7625354883a6f9b00bb48e3acb424c3cd5cab35247a',
  'objects.equations: duplicate slide':
    'sha256:c94c33372ad368532f3e7b50ee56133ff4b414094edfcd7772adfafb31bb2bae',
  'objects.equations: paste on a new slide':
    'sha256:fe28abbe16211ae3cd2448e3b86bbe1f4b26bd45e99f927aa20cccc48c2f9d29',
  'objects.equations: paste into another deck':
    'sha256:fe28abbe16211ae3cd2448e3b86bbe1f4b26bd45e99f927aa20cccc48c2f9d29',
  'themes.placeholder-inheritance: duplicate slide':
    'sha256:31dd35989a1ea1c0899c64cfdb075b51ce09424c235932cd82e630ae75d4a8d3',
};

/**
 * Comment edits (`COMMENT_EDITS` in `scripts/fidelity/comment-edits.ts`) whose copy opens with no
 * repair prompt, with the fingerprint of that copy (`commentFingerprint`). `bun
 * scripts/fidelity/comment-edits.ts --out <dir>` writes one copy per recipe. The rules of
 * `SAVE_VERIFIED_FEATURE_EDITS` apply. Last checked on 7 November 2026: all 4 copies, and the
 * `slide-structure` copies of `SAVE_VERIFIED_FIXTURES`, no prompt.
 */
export const SAVE_VERIFIED_COMMENT_EDITS: Readonly<Record<string, string>> = {
  'review.comments: legacy':
    'sha256:64183420e8f55b8733283cd7539dd9d667929833e5637fbdec2e62451ee6347d',
  'review.comments: duplicate slide':
    'sha256:38b70874f013aa43ce7cce68fc5ee1fe1abb2fed76076a574d26c3d468f8825a',
  'review.comments: duplicate threads':
    'sha256:536faea983a8f5292a63f709167171463ecf7486d38eaf01772618babaf39a2b',
  'review.comments: collaboration':
    'sha256:bb86cb89a50181a666fac29b3c8f5cafc5c5d14d752b6a04537a8c11c1c70551',
};

/**
 * Chart data edits (`CHART_DATA_EDITS` in `scripts/fidelity/chart-data-edits.ts`) whose copy opens
 * with no repair prompt, with the fingerprint of that copy (`editFingerprint`). `bun
 * scripts/fidelity/chart-data-edits.ts --out <dir>` writes one copy per recipe. The rules of
 * `SAVE_VERIFIED_FEATURE_EDITS` apply. Each embedded workbook of these copies also opens in a
 * spreadsheet application with no repair prompt. Last checked on 7 November 2026: all 7 copies
 * and their 14 workbooks, no prompt.
 */
export const SAVE_VERIFIED_CHART_DATA_EDITS: Readonly<Record<string, string>> = {
  'objects.charts: edit values':
    'sha256:b74a14578f542418e6603bab49d5285c2d3282eef95477c72ec3bcc15e3de982',
  'objects.charts: rename a series':
    'sha256:1f443d27e53a53c66140d755f44af7049068929e5ce465c534f9c6f4a04af0d9',
  'objects.charts: add a series':
    'sha256:e752034e4404ed11ff5a367d00ff0f007d44f957f10bee49726f1ad03ba64f05',
  'objects.charts: remove a series':
    'sha256:81b98b243aabccfd58ccb5196f4ec8e5c9d6a73d58b4727720382619e03aa5bb',
  'objects.charts: add a category':
    'sha256:b721ccc1c35c34a0f605b33ee8a0aab620d44abe03b79a57ff525f539e7a7dd7',
  'objects.charts: remove a category':
    'sha256:2145c94758bf47dbaec941625f3e369d5defb336b6a2a0f4b7145d8bd73bec8e',
  'objects.charts: edit a scatter chart':
    'sha256:249994d1f46f9e9b5220abcd107011c0cd891a25e5eb1935d8ab74a66105ee10',
};

/**
 * Diagram edits (`DIAGRAM_EDITS` in `scripts/fidelity/diagram-edits.ts`) whose copy opens with
 * no repair prompt, with the fingerprint of that copy (`editFingerprint`). `bun
 * scripts/fidelity/diagram-edits.ts --out <dir>` writes one copy per recipe. The rules of
 * `SAVE_VERIFIED_FEATURE_EDITS` apply. The presentation application lays each diagram out again
 * from its data model, and its PDF export of each copy shows the edit. Last checked on 7 November
 * 2026: all 5 copies, no prompt.
 */
export const SAVE_VERIFIED_DIAGRAM_EDITS: Readonly<Record<string, string>> = {
  'objects.diagrams: edit text':
    'sha256:d8a4710823dda85160e1f901f237254c32eb69b5386b28be7a8f2fb77a317168',
  'objects.diagrams: add a shape after':
    'sha256:0eee1174b5af2ebe8c5e7d6896878aa0be405a37fc45c4a5b0a2df3423f8a02a',
  'objects.diagrams: add a shape before':
    'sha256:88ebaa63a57b7d22c58e99a80b16eff98c3971afb723e8281901e9e3f582e592',
  'objects.diagrams: add a shape below':
    'sha256:c6de8f98b490ff05f5ab03e1b4f54e24273b28ce507ac6284742b0d693d72183',
  'objects.diagrams: delete a shape':
    'sha256:4ef5c76978e775d4a1d8b8e0ca5b7ae0b13d423d415905242e6f1bebc1c66526',
};
