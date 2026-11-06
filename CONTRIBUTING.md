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

The harness prints each fixture in `fixtures/` to PDF with the editor's own paint, and compares that PDF with the fixture's reference PDF. pdf.js draws and reads both PDFs in the same way. The check uses no system tools and no system fonts. Run it on your own machine; no container or second system is needed. Each slide must pass these checks (`scripts/fidelity/thresholds.ts`): at most 0.5 % of the pixels differ, every reference text line is in our page, our page has no other text lines, and each line is within 0.5 pt of the reference line, vertically (baseline) and horizontally (its nearer edge). Rotated, vertical, and right-to-left text is checked by pixels only. Save, reopen, and paint again must give the same SVG markup.

Our PDF must draw only with fonts that the harness loaded: the fixture fonts, the engine's fonts, and two open fallback faces. A system font in our PDF fails the gate. A fixture's reference PDF must embed only the open fonts that its deck names. `NON_OPEN_FONT_DECKS` in `scripts/fidelity/local-fonts.ts` lists the exceptions (`font-substitution`).

### Run

```bash
bun run fidelity          # every fixture: the gate
bun run fidelity tables   # one fixture
```

Each deck prints one line. The gate summary says what to do next. The report is `fidelity-report/fixtures/index.html`: for each slide, our page, the reference page, the diff, and the scores. Both PDFs are in `fidelity-report/fixtures/<deck>/`.

### Update the known gaps

`FIXTURE_KNOWN_GAPS` in `scripts/fidelity/thresholds.ts` lists, for each slide that fails today, the checks it fails. It records no values, because rasters differ slightly between machines. The gate fails when a slide fails a check that is not listed, or passes a listed check. When your change fixes a check, remove it:

```bash
bun run fidelity --update-gaps
```

A new entry or kind is written only with `--accept-worse`. In CI, `scripts/fidelity/check-known-gaps.ts` accepts it only for a new fixture, or in a pull request that changes no engine or harness source. A limit can go away only with its reason in `RETIRED_LIMITS`.

### Add a fixture

1. Add the deck to `scripts/fixtures/build.ts` and run `bun run fixtures:build`. Use only the open fonts of the engine (Carlito, Liberation Sans, Serif, and Mono, Noto Sans CJK JP) and of `fixtures/fonts/` (Noto Sans Arabic, Noto Sans Symbols 2). Name a font for every script in the deck, also in the theme, so the presentation app uses no fallback font.
2. Install those fonts where the presentation app runs, open the deck, and export `fixtures/<name>.pdf`. Any machine can export a reference.
3. Run `bun run fidelity:record-reference <name>`. It records the deck and PDF hashes, the page count, and the page size in `fixtures/references.json`. A missing or stale reference is a gap, never a pass.
4. Run `bun run fidelity <name> --update-gaps --accept-worse`, and commit the deck, the PDF, `references.json`, and `thresholds.ts`.

### Other commands

| Command | Use |
| --- | --- |
| `bun run fidelity:corpus` | The private decks in `demo-presentations/` (local only). `PPTX_HARNESS_FONT_DIRS` names the font folders of their reference fonts. Never commit these fonts. |
| `bun scripts/fidelity/run.ts --list <jobs> --out <dir>` | Corpus mode for `eigenpal/pptx-eval`: a job list of `{ "id", "deck", "pdf" }` in, one result line per deck out. `--help` lists every flag. |
| `bun run fidelity:roundtrip`, `fidelity:census`, `fidelity:calibrate` | Save checks with no browser, feature counts, and line model calibration over private decks. `--help` lists the flags. |

`PPTX_FIDELITY_JOBS` sets the number of parallel lanes (default: one for each core, at most 4).

Schema validation cannot show a repair prompt. `bun run fidelity:roundtrip -- --list <fixture jobs> --out <dir> --saved <dir>` writes a saved and an edited copy of each fixture. Add a fixture to `SAVE_VERIFIED_FIXTURES` (`docs/site/data/save-verified-fixtures.ts`) only after both copies open with no prompt.

Name every approximation in `packages/core/src/layout/approximations.ts`, in `docs/site/content/pptx-fidelity.mdx`, and in the evidence of each feature row that it changes (`docs/site/data/pptx-feature-evidence.ts`). Tests keep the lists equal.

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
