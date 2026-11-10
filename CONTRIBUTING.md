# Contributing to pptx-editor.dev

This guide covers setup, checks, and pull requests for the monorepo.

## Prerequisites

- [Bun](https://bun.sh/) 1.4.3, the CI version.
- [Node.js](https://nodejs.org/) 22 or later.
- Playwright Chromium for the fidelity harness: `bunx playwright install chromium`.

## Development setup

```bash
git clone https://github.com/eigenpal/pptx-editor.git
cd pptx-editor
bun install
bun run dev
# Open http://localhost:5274
```

The demo builds against workspace source, so it needs no package build.

## TypeScript tooling

`bun run typecheck` uses TypeScript 6 through the `tsc` command. The Next.js example uses the standard compiler API integration. Declaration generation, source analysis, and API compatibility checks use the same TypeScript dependency.

The separate DOM-free lane guard retains `bun check` to detect ambient DOM type leakage.

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
   - `bun run test:dom` runs the React and Vue adapter tests in one process, so they share one set of DOM classes. A test that patches a DOM prototype must restore the saved property descriptor, or delete an own property that was not there. An assignment such as `HTMLElement.prototype.getBoundingClientRect = saved` adds a copy that hides later patches of `Element.prototype`. `spyOn(HTMLElement.prototype, 'getBoundingClientRect')` followed by `mockRestore()` also leaves such a copy, because the method is inherited. Save and restore the descriptor instead:

     ```ts
     const own = Object.getOwnPropertyDescriptor(HTMLElement.prototype, 'getBoundingClientRect');
     HTMLElement.prototype.getBoundingClientRect = fake;
     // ... the test ...
     if (own) Object.defineProperty(HTMLElement.prototype, 'getBoundingClientRect', own);
     else delete (HTMLElement.prototype as Partial<HTMLElement>).getBoundingClientRect;
     ```

     The preload fails a test that leaves such a copy, and fails the run when a prototype differs at the end.

   - `bun run test --shard 1/3 --balance` runs one CI shard. `--only FILE`, `--exclude PREFIX`, `--jobs N`, `--fail-fast`, and `--retry-failed` are also available. Arguments after `--` go to each `bun test`.
   - Copy DOCX Editor pro code only into `packages/pro`, never into an Apache-2.0 package.
5. Add a changeset for code changes with `bun changeset`. Write one short paragraph of at most 400 characters, as [the changeset guide](.changeset/README.md) says. Documentation, test, and CI-only changes do not need one.
6. Submit a pull request against `main` with a conventional-commit title.

## Browser responsiveness checks

Run these checks against the source demos in this checkout:

```bash
bun run test:hang:browser
bun run test:stress:interactions
bun run test:fixtures:browser
```

The hang corpus covers large presentations, extreme text, geometry, tables, pictures, and font waits. Each presentation opens through the demo's file input in a fresh browser context. The check records open time, main-thread gaps, and frame gaps, then checks recovery with another presentation. A crash, an uncaught error, a request outside the demo, an open over 20 seconds, or a main-thread gap over 5 seconds fails it. Set `HANG_STRICT=1` to also fail on a gap over 1 second. In Chromium, the check also records long tasks, heap size, and DOM size. `bun run test:hang:browser` also runs the bullet and font wait checks (`bullet-hang.stress.ts` and `font-hang.stress.ts`). `bun run bench:hang:browser` runs only the hang corpus (`hang.stress.ts`). Rows go to `test-results/hang-<browser>/rows.jsonl`. The interaction suite runs the same navigation, editing, paste, and replacement sequences in React and Vue. The fixture suite opens each synthetic fixture, navigates its slides, and saves and reopens it. These checks test responsiveness. They do not replace PDF fidelity comparisons.

Set `HANG_BROWSER=firefox` or `HANG_BROWSER=webkit` to use another browser engine. Install the engines first with `bunx playwright install chromium firefox webkit`. Set `HANG_ONLY=slides-3000,tabs-many-stops` to select hang corpus cases. Each suite starts its own servers and refuses occupied ports.

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
- The text keeps each character exactly once (`packages/pro/src/__tests__/fuzz-ledger.ts`). A character shows unless a peer removed it, and stays removed unless a peer brought it back. The summary line reports exemptions and blind checks. Both counts must stay at zero.
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

The harness prints each fixture in `fixtures/` to PDF with the editor's own paint, and compares that PDF with the fixture's reference PDF. pdf.js draws and reads both PDFs in the same way. It converts ICC-based colors through their profiles with its WebAssembly color engine, as a color-managed viewer does. The check uses no system tools and no system fonts. Run it on your own machine; no container or second system is needed. Each slide must pass these checks (`scripts/fidelity/thresholds.ts`): at most 0.5 % of the pixels differ, every reference text line is in our page, our page has no other text lines, and each line is within 0.5 pt of the reference line, vertically (baseline) and horizontally (its nearer edge). A line pairs with a line of the same text in the other PDF only within 4 pt (`LINE_PAIR_LIMIT_PT` in `scripts/fidelity/compare.ts`, 8 times the line tolerance), nearest pair first. pdf.js lines carry no shape, so the limit keeps repeated text in one shape from pairing with the same text in another shape. After that, lines of the same text that are left pair at any distance, nearest first: a line that moved fails by its distance, and is not reported as missing and extra. Both PDFs follow the same rules, so the pairs do not depend on which PDF is the reference. A line is checked when its text, without whitespace and a leading bullet, has at least two letters or digits, so numbers, dates, percentages, and chart and table values are checked as words are. A currency sign, a math sign, or a dash with no space after it is part of the number (`$1,200`, `−5`), not a bullet. A line of one letter or digit, such as a superscript or a chart tick, is checked by pixels only, and so are rotated, vertical, and right-to-left text and glyphs with no advance (marks that sit on another glyph). Sheared text, such as a synthetic oblique face, is checked. Save, reopen, and paint again must give the same SVG markup.

Our PDF must draw only with fonts that the harness loaded: the fixture fonts, the engine's fonts, and two open fallback faces. A system font in our PDF fails the gate. A fixture's reference PDF must embed only the open fonts that its deck names. No fixture is exempt.

### Run

```bash
bun run fidelity                 # every fixture, both lanes: the gate
bun run fidelity tables          # one fixture
bun run fidelity --lane export   # one lane: editor or export
```

Each deck prints one line for each lane. The gate summary of each lane says what to do next. The report is `fidelity-report/fixtures/index.html`: for each slide, our page, the reference page, the diff, and the scores. Both PDFs are in `fidelity-report/fixtures/<deck>/`.

### The export lane

The export lane gates the PDF export (`packages/pptx-to-pdf`) on the same fixtures. It exports each fixture with `exportPdf` from source, with the settings of the evaluation protocol (`packages/pptx-to-pdf/scripts/evaluation/settings.ts`) and the font faces that the editor lane gets, plus the packaged substitutes of the export. It compares that PDF with the same reference PDF through the same code (`rasterPdf`, `comparePixels`, and `compareLines`) and the same thresholds. It has no save check. `exportPdf` names each font that it embeds `PptxFont`, so the lane checks only that every font is embedded. A system font cannot reach the exported PDF: the exporter reads no font file other than the optional CJK package, and `scripts/fidelity/export-fonts.test.ts` keeps it so. The pages of both lanes are drawn with pdf.js in Chromium, so a change of the pinned Chromium records both lists again. Its report is `fidelity-report/fixtures-export/index.html`, and the report of each lane links to the other.

### Update the known gaps

`FIXTURE_KNOWN_GAPS` in `scripts/fidelity/thresholds.ts` lists, for each slide that fails today, the checks it fails. `FIXTURE_EXPORT_KNOWN_GAPS` is the list of the export lane, with the same rules. It records no values, because rasters differ slightly between machines. The gate fails when a slide fails a check that is not listed, or passes a listed check. When your change fixes a check, remove it:

```bash
bun run fidelity --update-gaps
```

The command updates the list of each lane that ran.

A new entry or kind is written only with `--accept-worse`. In CI, `scripts/fidelity/check-known-gaps.ts` accepts it only for a new fixture, in a pull request that changes no engine or harness source, or in a pull request whose source changes are only the comparison files (`COMPARISON_FILES` in `scripts/fidelity/thresholds.ts`), harness tests, and documentation. Such a change prints the same pages and checks them in another way, so the failures it finds were there before. It may not check fewer lines, and it changes no engine code. The check reads `COMPARISON_FILES` from the base branch. A limit can go away only with its reason in `RETIRED_LIMITS`.

### Add a fixture

1. Add the deck to `scripts/fixtures/build.ts` and run `bun run fixtures:build`. Use only the open fonts of the engine (Carlito, Liberation Sans, Serif, and Mono, Noto Sans CJK JP) and of `fixtures/fonts/` (Noto Sans Arabic, Noto Sans Symbols 2, STIX Two Math). Name a font for every script in the deck, also in the theme, so the presentation app uses no fallback font.
2. Install those fonts where the presentation app runs, open the deck, and export `fixtures/<name>.pdf`. Any machine can export a reference.
3. Run `bun run fidelity:record-reference <name>`. It records the deck and PDF hashes, the page count, and the page size in `fixtures/references.json`. A missing or stale reference is a gap, never a pass.
4. Run `bun run fidelity <name> --update-gaps --accept-worse`, and commit the deck, the PDF, `references.json`, and `thresholds.ts`.

### Other commands

| Command | Use |
| --- | --- |
| `bun run fidelity:corpus` | The private decks in `demo-presentations/` (local only). `PPTX_HARNESS_FONT_DIRS` names the font folders of their reference fonts. Never commit these fonts. |
| `bun scripts/fidelity/run.ts --list <jobs> --out <dir>` | Corpus mode for `eigenpal/pptx-eval`: a job list of `{ "id", "deck", "pdf" }` in, one result line per deck out. `--help` lists every flag. |
| `bun run fidelity:roundtrip`, `fidelity:census`, `fidelity:calibrate` | Save checks with no browser, feature counts, and line model calibration over private decks. `--help` lists the flags. |

`PPTX_FIDELITY_JOBS` sets the number of parallel lanes (default: one for each core, at most 4). Every fixture deck gets the same fonts, so each lane opens the fixture decks in one harness page. The page fetches, hashes, and loads each font file once (Noto Sans CJK JP is 16 MB), and the next decks use the same faces. A corpus deck gets the fonts that it names, so each corpus deck opens in a new browser context.

Schema validation cannot show a repair prompt. `bun run fidelity:roundtrip -- --list <fixture jobs> --out <dir> --saved <dir>` writes a saved and an edited copy of each fixture. Add a fixture to `SAVE_VERIFIED_FIXTURES` (`docs/site/data/save-verified-fixtures.ts`) only after both copies open with no prompt.

A feature row claims full rendering only when every fixture of the row passes the harness, the row owns no approximation but the line model, glyph advances, and missing glyphs, and its notes name no rendering gap.

A feature row claims a full round trip only with its own edit. `FEATURE_EDITS` in `scripts/fidelity/feature-edits.ts` holds one recipe per row: it edits a fixture of the row through the public editor API, saves, reopens, and finds the edit again. The check of a recipe must fail on the unedited deck, a second save must give the same package, and the edit must change only the parts it owns. `bun test scripts/fidelity/feature-edits.test.ts` runs every recipe. `bun scripts/fidelity/feature-edits.ts --out <dir>` writes one edited copy per row. A row that claims `preserved` also needs a recipe: it edits next to the kept content or its box, and its check finds the content as it was. A recipe with `collaborate` edits in two editors on linked replicas, each with its own collaboration session. Add the row and the fingerprint of its copy to `SAVE_VERIFIED_FEATURE_EDITS` only after the copy opens with no prompt. When a change gives a copy with another fingerprint, the test fails: open the new copy, and record its fingerprint.

A row that claims `preserved` also keeps its content through structural edits. `KEPT_CONTENT_EDITS` in `scripts/fidelity/kept-content-edits.ts` duplicates the slide of each such row, pastes its kept objects on a new slide and into another deck, and groups them. Its check finds each kept object with the same XML and the same parts, and every relationship names a part that exists. `bun test scripts/fidelity/kept-content-edits.test.ts` runs every recipe, and `bun scripts/fidelity/kept-content-edits.ts --out <dir>` writes one copy per recipe. Record each copy in `SAVE_VERIFIED_KEPT_CONTENT_EDITS` only after it opens with no prompt.

`COMMENT_EDITS` in `scripts/fidelity/comment-edits.ts` edits comments in both formats, on duplicated slides, and in two collaborating editors. `bun test scripts/fidelity/comment-edits.test.ts` runs every recipe, and `bun scripts/fidelity/comment-edits.ts --out <dir>` writes one copy per recipe. Record each copy in `SAVE_VERIFIED_COMMENT_EDITS` only after it opens with no prompt. The fingerprint of a collaboration recipe leaves out the numbers of the session's id stripe, because they depend on the sessions that ran before in the process.

A repair prompt holds the open of a deck until you answer it. To check many copies, open each one as a file manager does, and wait until the presentation application lists it as open. First, prove the check with a deck that must prompt, such as a copy with a slide part removed.

Name every approximation in `packages/core/src/layout/approximations.ts`, in `docs/site/content/pptx-fidelity.mdx`, and in the evidence of each feature row that it changes (`docs/site/data/pptx-feature-evidence.ts`). Tests keep the lists equal.

## Write documentation

Follow the [developer documentation rules](AGENTS.md#developer-documentation) and [site source conventions](docs/site/README.md). Lead with the developer task and a useful code example. Use concise prose for prerequisites, defaults, lifecycle, and limits. Keep shared SDK tasks in the established page structure.

Verify samples against public exports and types. Register pages in both navigation files. Run `bun test docs`, `node scripts/check-public-terminology.mjs`, and Prettier on edited documentation. Source checks do not replace a website build or browser review.

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

Contributions to `core`, `react`, `vue`, `i18n`, and `ooxml-core` use [Apache 2.0](LICENSE). Contributions to `pro`, `pptx-to-pdf`, and `editor-api` use the EigenPal Pro License. Third-party material keeps its own terms.

## Report bugs

Use a [GitHub issue](https://github.com/eigenpal/pptx-editor/issues) for reproduction steps, expected behavior, and actual behavior. Attach only synthetic or sanitized decks. Follow [SECURITY.md](SECURITY.md) for vulnerabilities.

## Public terminology

Use PPTX, presentation, or reference PDF in product prose, and follow "Naming" in [Agent instructions](AGENTS.md). Keep technical identifiers, source attribution, and legal notices exact.

Run `node scripts/check-public-terminology.mjs` after you edit public wording. The check covers root consumer docs, package READMEs, site pages, locale catalogs, and demo text sources. It does not cover engine code, generated API snapshots, legal notices, or third-party references, so review those by hand.

## Collaboration release compatibility

Record compatibility decisions for Core and Pro implementation changes. Follow [the compatibility policy](docs/architecture/collaboration-compatibility.md). Run `bun run collaboration:check --base origin/main` and, after building packages and notices, `bun run collaboration:test`. Stable releases capture published installations and saved rooms through the post-release baseline workflow. Merge the capture pull request before the next stable publication.
