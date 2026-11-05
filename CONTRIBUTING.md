# Contributing to pptx-editor.dev

This guide covers setup, checks, and pull requests for the monorepo.

## Prerequisites

- [Bun](https://bun.sh/) 1.4.3, the CI version.
- [Node.js](https://nodejs.org/) 22 or later.
- Playwright Chromium for the fidelity harness: `bunx playwright install chromium`.
- Poppler (`pdftoppm`, `pdftotext`) for the comparison with the reference PDFs.

## Development setup

```bash
git clone https://github.com/eigenpal/pptx-editor.git
cd pptx-editor
bun install
bun run dev
# Open http://localhost:5274
```

The demo builds against workspace source, so it needs no package build.

## Make changes

1. Create a branch from `main`.
2. Read the [engine lanes](docs/architecture/engine-lanes.md) and [AGENTS.md](AGENTS.md).
3. Keep each change focused, with tests and a fixture for any rendering feature.
4. Verify the change:
   ```bash
   bun run typecheck
   bun run lint
   bun run format:check
   bun run test
   bun run build:packages && bun run api:check
   bun run i18n:validate
   bun run check:security
   bun run check:css-scope
   bun run fidelity
   ```
   `bun run test` runs one `bun test` process per test file over a pool of workers, slowest file first, as DOCX Editor does (`scripts/test/run-parallel.mjs`). Only the React and Vue adapter tests preload happy-dom. The engine tests run with no global DOM. Other test commands:
   - `bun run test:changed` runs only the test files that your change can affect. Add `--list` to print them.
   - `bun run test:serial` runs the suite in two `bun test` processes. Use it to find a test that depends on file order.
   - `bun run test --shard 1/3 --balance` runs one CI shard. `--only FILE`, `--exclude PREFIX`, `--jobs N`, `--fail-fast`, and `--retry-failed` are also available. Arguments after `--` go to each `bun test`.
5. Add a changeset for code changes with `bun changeset`. Documentation, test, and CI-only changes do not need one.
6. Submit a pull request against `main` with a conventional-commit title.

## Add a collaboration fixture

Every editing command and every kind of text input must converge when two people edit one deck. The gate, `packages/pro/src/__tests__/collaboration-coverage.test.ts`, runs each edit on peer A of two editors on linked Y.Docs, flushes, and compares both replicas and their saved decks. When you add a command:

1. Add the command to `EditorCommand` in `packages/core/src/contracts/index.ts`. `bun run typecheck` then reports `Property '<name>' is missing` in `packages/core/src/editor/command-types.ts`. Add `<name>: true` to the right group there.
2. Add a fixture to `CONVERGENCE_FIXTURES` in `packages/pro/src/__tests__/collaboration-coverage-fixtures.ts`. Its `arrange` prepares peer A (a selection, a slide, or earlier edits), and its `command` builds the command. The coverage deck is `coverageDeck()` in `convergence-harness.ts`; pass `deck` when the command needs other content. The command must change the deck: a fixture where it is a no-op fails.
3. Only when the command writes nothing to the deck (a view setting or session state), add its reason to `COLLABORATION_UNCOVERED` in `packages/pro/src/collaboration/coverage-contract.ts` and a probe to `NO_WRITE_PROBES`. The probe proves that the command changes neither replica.
4. Run `bun test packages/pro/src/__tests__/collaboration-coverage.test.ts`. Each failure message names the file to change and, when replicas differ, the parts that differ.

A new `TextInputKind` follows the same steps with `TEXT_INPUT_KINDS` and `TEXT_INPUT_FIXTURES`. The gate covers one edit on one peer. Test concurrent edits, undo, and deletion during an edit in the feature's own `yjs-*.test.ts` file.

## Add an editing API scenario

The editing API (`packages/editor-api`) has one catalog of add-in scenarios in `packages/editor-api/src/__tests__/scenarios/`. Each scenario is an Office.js-style script, `async (context) => { ... await context.sync(); }`, that answers JSON. Two suites run every scenario:

- `bun run test:automation` runs it on a server runtime over its fixture deck. A scenario that writes is saved, checked with `scripts/fixtures/package-integrity.ts`, reopened, and read back. The outcomes are pinned in `__snapshots__/scenarios.test.ts.snap`.
- The same command runs it in bun on a browser runtime over a live editor (`createPptxEditor` with `createBrowserAutomationHost`), and compares the outcome with the server outcome. The editor's save must be whole and reopen on a server runtime with the same read-back. Each sync that writes must be one undo step, even when it splits. Pictures answer the same on both hosts: the server draws them with a stub rasterizer, and the live editor with the real paint on a stub canvas (`stub-drawing.ts`). The checks are in `scenario-hosts.ts`.

When you add a member to the API:

1. Add a scenario to the module of its area. Import only types at run time, and write enum values as string literals, so the same script runs in both hosts.
2. Give a scenario that writes a `readBack`, which reads what the script wrote.
3. Set `host: 'selection'` when the scenario uses the selection, which a server does not have. Set `requires` (for example `['Shape.newMember']`) for a member that is not built yet; both suites then report the scenario as to do.
4. Run `bun run test:automation`, and check the new snapshot by eye.

## Fuzz tests

The collaboration fuzzers make seeded random edits by three editors on linked Y.Docs. A relay holds each update back and delivers parts of it in a random order, sometimes over several rounds. Each seed is deterministic, so a failure names a seed that you can run again on any machine.

| File | Edits | Prefix | Name for `-t` |
| --- | --- | --- | --- |
| `packages/pro/src/__tests__/yjs-text-fuzz.test.ts` | Typing, typing over a selection, deletes in one paragraph and across two, Enter, joins, undo, and redo | `TEXT` | `text fuzz` |
| `packages/pro/src/__tests__/yjs-text-fuzz.test.ts` | The same, with bold, italic, and color | `FORMAT` | `format fuzz` |
| `packages/pro/src/__tests__/yjs-table-fuzz.test.ts` | Row and column inserts and deletes, merges, splits, cell text, undo, and redo | `TABLE` | `table fuzz` |
| `packages/pro/src/__tests__/yjs-table-fuzz.test.ts` | The same, in a table whose center cell has no text body | `TABLE_BODILESS` | `table fuzz, cell without a text body` |

The checks:

- Each edit changes only what it may on its own replica. For example, a delete removes exactly its selection, and an undo removes only text that its own peer typed.
- After each flush, every session is `ready`, and all replicas hold the same package.
- The text keeps each character exactly once (`packages/pro/src/__tests__/fuzz-ledger.ts`). A character shows unless a peer removed it, and stays removed unless a peer brought it back. The only exemption is a character placed at the same time as a peer's removal of its paragraph, row, or column. The summary line of each run counts the exemptions.
- The table fuzzer also compares the grid that the canvas draws with the grid of the saved deck.
- At the end, every replica's save reopens to the same deck and passes the part checks.

The shared harness is `packages/pro/src/__tests__/fuzz-support.ts`.

`bun run test` runs the default seeds (24 for each text mode, 16 for tables), fixed histories, and the known failures. Pull request CI skips the two fuzz files, because they are the longest in the suite; every push to `main` runs them. To run more seeds, use `bun run fuzz:collaboration`, which runs 400 seeds of each fuzzer. As in DOCX Editor, every failing seed fails the run, so the command fails until the open issues are fixed. Read the summary at the end of each fuzzer: it groups the failing seeds by kind of problem.

Set these variables to control a run. Replace `<PREFIX>` with the prefix from the table:

| Variable | Effect |
| --- | --- |
| `<PREFIX>_FUZZ_SEEDS=<n>` | Run `n` seeds. |
| `<PREFIX>_FUZZ_FROM=<seed>` | Start at this seed. The default is 1. |
| `<PREFIX>_FUZZ_TRACE=1` | Print every action with the state of each replica. |
| `<PREFIX>_FUZZ_SHRINK=<seed>` | Reduce the failure of this seed to the fewest actions, and print them. |

Add `-t '<name>'` with the name from the table to run one fuzzer of a file. A failing seed prints its problem, its last actions, the known failures of the same kind, and the commands to trace and shrink it, for example:

```bash
FORMAT_FUZZ_FROM=4 FORMAT_FUZZ_SEEDS=1 FORMAT_FUZZ_TRACE=1 bun test packages/pro/src/__tests__/yjs-text-fuzz.test.ts -t 'format fuzz.* seed 4( \(known|$)'
FORMAT_FUZZ_SHRINK=4 bun test packages/pro/src/__tests__/yjs-text-fuzz.test.ts -t 'format fuzz'
```

### Known failures

A failure that an open issue explains is pinned, not skipped. Its test passes only while it fails with exactly the problem it pins. The test fails when the problem changes, and when the history passes. So a new defect cannot hide in a known failure, and a fix shows at once.

- `KNOWN_FAILURES` in each fuzzer file holds shrunk histories, each with its issue and its `problem`.
- `KNOWN_SEEDS` holds the default seeds that fail today, each with its issue and its `problem`.
- `REGRESSIONS` (and `yjs-text-fuzz-regressions.ts`) hold fixed histories that must pass.

To turn a new failure into a test:

1. Shrink it, for example `TABLE_FUZZ_SHRINK=17 bun test packages/pro/src/__tests__/yjs-table-fuzz.test.ts`. The output is a `KNOWN_FAILURES` entry with the minimal actions and their problem.
2. Look for an open issue with the same steps. If there is none, file one with the actions and the steps they show.
3. Add the entry to `KNOWN_FAILURES` with the issue number. If the seed is a default seed, also add the `KNOWN_SEEDS` line that the shrink prints.
4. When you fix the issue, its tests fail with "passes now". Move each `KNOWN_FAILURES` entry to `REGRESSIONS` (drop `issue` and `problem`), and remove its seeds from `KNOWN_SEEDS`.

When a change alters a pinned problem (for example, a fix changes the text that a known failure shows), the test prints the new problem. Update the pin if the change explains it. Otherwise, file an issue.

## Fidelity harness

Every rendering feature lands with a synthetic fixture in `fixtures/` (`bun run fixtures:build`) and its harness result. Add its reference PDF next to it, with the same base name and its `fixtures/references.json` entry. Never commit content from the private corpus in `demo-presentations/`.

### Commands

| Command | What it does |
| --- | --- |
| `bun run fidelity` | Renders every fixture slide and compares it with the fixture's reference PDF. This is the CI gate. |
| `bun run fidelity -- --print-gaps` | Prints the `FIXTURE_KNOWN_GAPS` entries that this run's values record. |
| `bun run fidelity:check-gaps -- --base origin/main` | Checks that the known-gap list only tightens against the base branch. |
| `bun run fidelity:missing` | Lists missing and stale references. `--check` exits 1 when one is missing or stale. |
| `bun run fidelity:roundtrip -- --list <jobs> --out <dir> [--saved <dir>]` | Opens, saves, reopens, and edits each deck of a job list, with no browser. `--saved` keeps the saved copies for the save check. `--help` lists every flag. |
| `bun run fidelity:corpus`, `bun run fidelity:census`, `bun run fidelity:calibrate --smoke` | Local checks over private decks. Never run them in CI. |
| `bun scripts/fidelity/run.ts --list <jobs> --out <dir>` | Corpus mode: a JSON or JSON Lines job list of `{ "id", "deck", "pdf" }` in, one result line per deck out. `--shard i/n`, `--resume`, `--fresh`, and `--images` are available. |
| `bun run validate:xsd` | Validates parts against the ECMA-376 Transitional schemas before and after save and after an edit session. |

The harness starts the demo's Vite server, opens `examples/vite/harness.html` in Playwright Chromium, and renders each slide through the editor's own `openPresentation`, `layoutSlide`, and `paintSlide`. It needs Poppler (`pdftoppm`, `pdftotext`) and Playwright Chromium.

### Environment variables

Bun reads them from the git-ignored `.env.local`.

| Variable | Use |
| --- | --- |
| `PPTX_HARNESS_FONT_DIRS` | The font folders that the harness loads fonts from, separated by `:`, in priority order. The first folders hold the reference fonts. A font collection (`.ttc`, `.otc`) gives the page each of its faces as a single font. Without it, the harness uses the metric-compatible substitutes. Never commit these fonts. |
| `PPTX_FIDELITY_JOBS` | The number of harness lanes: child processes that score decks in parallel, each deck in a new browser context, so no score depends on the decks before it. The default is one for each core, at most 4. The report is the same for any number. |
| `PPTX_FIDELITY_RASTER_CACHE` | The folder that keeps the fixture rasters between runs. The default is `.cache/fidelity-rasters`. |
| `PPTX_EVAL_CORPUS` | The corpus folder for `bun run fidelity:calibrate --smoke`. |
| `PPTX_EVAL_GENERATION` | The default reference generation for `bun run fidelity:calibrate --smoke`. `--generation` overrides it. |

### Checks

Thresholds and tolerances live in one file, `scripts/fidelity/thresholds.ts`. For each slide the harness checks the following.

| Check | Rule |
| --- | --- |
| Slide size | The render and the raster are exactly their expected sizes. A reference page side is a whole number of points, so the raster is cut where the page rounds a side down (`PAGE_SIZE_TOLERANCE_PT`). |
| Pixels | At most 0.5 % of the pixels differ, with anti-aliasing tolerance judged at the device scale. |
| Line breaks | Every laid-out line appears as a PDF line with the same characters. Whitespace is ignored, and a bullet or number may lead. |
| Baselines | Each baseline is within 0.5 pt of the glyph origin in the PDF. |
| Save consistency | Save, reopen, and render again give identical pixels, compared at the device scale (see Supersampled rasters, below). |

A reference is current while the deck and the PDF are the bytes that `fixtures/references.json` records. A stale or missing reference is a gap, never a pass. Hidden slides are not in the reference PDFs, so PDF pages map to the visible slides in order.

Text in table cells, rotated text, and vertical text are checked by pixels only. Geometry is checked by pixels only.

### Supersampled rasters

The two rasterizers snap glyphs to the pixel grid in different ways. Chromium rounds each glyph's y to a whole device pixel. Poppler's Splash rasterizer floors it. At the comparison size of 2 px per point, an exactly placed line is then 1 px low in our image whenever the PDF baseline has a pixel fraction of 0.5 or more.

So the harness works at `THRESHOLDS.rasterScale` (4) device pixels for each comparison pixel:

1. Chromium paints the slide at device scale 4, and `pdftoppm` rasterizes the page at 4 times the resolution.
2. `scripts/fidelity/supersample.ts` box-filters each image down to 1920 px wide with the same filter. The snapping now moves a glyph by at most 0.25 px.
3. The pixel check finds the pixels that differ on the filtered images. It decides which differences are anti-aliasing on the device-scale images (`deviceAntialiased`): a differing pixel counts when its 4 × 4 block of device pixels has more real differences than anti-aliased ones. The filter turns every edge into a ramp, so pixelmatch's anti-aliasing test on the filtered images would excuse a stroke moved by 1 px.
4. The save-consistency check compares the two device-scale renders, before the filter.

`pdftoppm` writes PPM, which holds the same pixels as its PNG without the encoding. The fixture set keeps each device raster in `.cache/fidelity-rasters/`, compressed, under a key of the PDF's SHA-256, the Poppler version, and the `pdftoppm` options, so a later run reads it and does not run `pdftoppm`. A run over every fixture removes the entries that it did not use. CI restores this folder, and pushes to `main` save it. The save-consistency check compares the PNG bytes first: the same bytes are the same pixels.

The thresholds are the same as at scale 1. These measurements on every fixture slide (macOS, reference fonts) support the method:

- At scale 1, lines whose PDF baseline has a pixel fraction of 0.5 or more are 0.85 px low (median of 180 lines), and the other lines are 0.14 px high (142 lines). At scale 4, the two groups are 0.02 px high and 0.12 px low.
- A text box moved down by 0.25, 0.5, 0.75, and 1 px scores higher at each step (`line-model/1`: 0.45 %, 0.55 %, 0.73 %, 0.95 %, 1.27 %). At scale 1 the score jumps at 0.5 px (0.29 % to 0.51 %).
- Errors put into fixture decks fail or score worse, as at scale 1: a 1 pt move of the text on `line-model/1`, `line-model/2`, `text-runs/1`, `spacing/1`, and `alignment-anchor/1`; a 1 px move of a shape with a 1.5 pt outline; a 1 px move of a 4.5 pt double line; and a fill color that is 24 levels off in one channel. A fill 16 levels off passes at both scales (`THRESHOLDS.pixelThreshold`).
- 39 slides pass, at scale 1 and at scale 4. 4 slides that fail on lines (`autofit/1`, `fills-lines/1`, `paragraph-format/1`, `vertical-columns/1`) fail on pixels too, and most pixel scores are higher, because a 1 px move of a stroke now counts.
- Scale 8 breaks a pattern fill (`fill-fidelity/5`, 0.00 % to 1.56 %) and takes about 8 times longer than scale 1. Scale 4 takes about 2.5 times longer: about 5 minutes for the fixtures on a laptop.

Each known-gap list records the `rasterScale` and the `antialiasing` mode that it was measured with. Each changed once, and every list was measured again in that change: the scale from 1 to 4, then the anti-aliasing test from the compared images (`images`) to the device-scale images (`device`). `bun run fidelity:check-gaps` rejects any other change of `THRESHOLDS.rasterScale` or `THRESHOLDS.antialiasing` (`RASTER_SCALE_MOVE` and `ANTIALIASING_MOVE` in `scripts/fidelity/check-known-gaps.ts`), so neither can re-record the lists again.

The baseline of a PDF line is the glyph origin that carries most of the line's own text, so a raised or lowered run does not stand in for its line.

### Known gaps

`FIXTURE_KNOWN_GAPS` in `scripts/fidelity/thresholds.ts` lists, for each font environment, the slides that fail today, each threshold they miss, and the measured value. `bun run fidelity` exits 1 when a slide fails in a way its entry does not list, when a listed slide scores worse or better than its entry, when a reference is missing or stale, or when Poppler is not installed. When a slide gets better, record the new value in the same change.

The list only tightens. A new entry, a new kind, or a looser value is allowed only for a new fixture, a changed deck with its new reference PDF, or a reference PDF from another macOS version or font set. A list that is measured again with other Poppler or Chromium versions records them in the list and may take new values. A renamed fixture keeps the entries of its old name. One looser value or smaller line count has an owner exception: `OWNER_GAP_EXCEPTION` in `scripts/fidelity/check-known-gaps.ts` names its exact slide, environment, kind, and values, and gives the evidence.

The report is in `fidelity-report/<set>/index.html`. `gate.json` holds the gate's verdict, and `summary.json` holds each slide's scores. The CI artifact gives the `linux` entries.

### Save check

Schema validation cannot show a repair prompt. `bun run fidelity:roundtrip -- --list <fixture jobs> --out <dir> --saved <dir>` writes `<id>.saved.pptx` (open, save) and `<id>.edited.pptx` (a scripted edit session, then save) for each fixture. Open each copy and note any repair prompt. Add a fixture to `SAVE_VERIFIED_FIXTURES` in `docs/site/data/save-verified-fixtures.ts` only after both copies open with no prompt and with every slide. Repeat the check after a change to the store, the writer, or an op.

### Approximations

Name every approximation in `packages/core/src/layout/approximations.ts` and in the known approximations of `docs/site/content/pptx-fidelity.mdx`. A test keeps the two lists equal. Then name its id in `approximations` of the evidence entry of each feature row whose behavior it changes (`docs/site/data/pptx-feature-evidence.ts`). A test checks that every approximation that layout finds on a fixture is named by a row that lists such a fixture.

## Write documentation

Follow the [Google developer documentation style guide](https://developers.google.com/style/highlights). Use the second person, active voice, present tense, sentence-case headings, and descriptive link text.

## Releases and canary builds

The Release workflow (`.github/workflows/release.yml`) publishes from `main` with npm trusted publishing. No npm token is stored.

- **Stable release.** A push to `main` with pending changesets opens or updates the `chore: release` PR. Merge that PR to publish every package on the `latest` tag.
- **Canary.** Each night at 03:23 UTC, the workflow publishes the newest `main` commit that passed CI on the `canary` tag, for example `0.1.0-canary.20261106032300.sha-b20ba9f`. To start one manually, run the Release workflow on `main` with **channel** set to `canary`.

Install a canary with `npm install @pptx-editor.dev/react@canary`. A caret range such as `^0.1.0` never installs a canary.

A canary run publishes nothing, and does not fail, in these cases:

- No pending changeset releases a published package.
- Every package already has a canary from that commit.
- A package has no `latest` version yet. npm would point `latest` at the canary.
- A stable release is due or running: `main` has no pending changesets, or another Release run that is not scheduled is active. A queued manual canary counts as such a run.
- A registry lookup or the stable-release check fails.

A canary run fails in these cases: it runs on a branch other than `main`, no `main` commit has a passing CI run, GitHub cannot list the CI runs, or a build, check, or publish step fails.

The stable-release check runs again right before the canary publishes, but a release PR merged in the minutes after that check can still collide with the canary. If the stable publish then fails, run the Release workflow again with **channel** set to `release`. `changeset publish` skips the versions that are already on npm.

`changeset version --snapshot canary` changes the versions on the runner only. A canary makes no commit, git tag, changelog entry, GitHub Release, or Slack message, and `latest` does not move. It pins the internal peer and runtime ranges to the exact canary version. The snapshot settings are under `snapshot` in `.changeset/config.json`. The decision logic is in `scripts/release-mode.mjs`.

A published package must not list a `workspace:` range in `dependencies`, because `npm publish` does not rewrite it. Use a real range, such as `^0.0.0`. Changesets updates it on each version.

## Contributor license agreement

You must sign the [Contributor License Agreement](CLA.md).

## License

Contributions to `core`, `react`, `vue`, `i18n`, and `ooxml-core` use [Apache 2.0](LICENSE). Contributions to `pro`, `pptx-to-pdf`, and `editor-api` use the EigenPal Pro License, as will the future `telemetry` package. Third-party material keeps its own terms.

## Report bugs

Use a [GitHub issue](https://github.com/eigenpal/pptx-editor/issues) for reproduction steps, expected behavior, and actual behavior. Attach only synthetic or sanitized decks. Follow [SECURITY.md](SECURITY.md) for vulnerabilities.

## Public terminology

Use PPTX, presentation, or reference PDF in product prose, and follow "Naming" in [Agent instructions](AGENTS.md). Keep technical identifiers, source attribution, and legal notices exact.

Run `node scripts/check-public-terminology.mjs` after you edit public wording. The check covers root consumer docs, package READMEs, site pages, locale catalogs, and demo text sources. It does not cover engine code, generated API snapshots, legal notices, or third-party references, so review those by hand.
