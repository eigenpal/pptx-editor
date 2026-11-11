# Collaboration compatibility and release baselines

The room format has four independent fields: protocol, shared schema, repair, and canonical model. The constants live in `packages/pro/src/collaboration/yjs/format.ts`. The legacy `schema` guard remains equal to the shared schema field. Sessions check all fields before attachment and after metadata changes. Hosts must check admission before exchanging updates.

Increment each field whose incompatible interpretation changes. Do not increment fields for fixes that released readers interpret identically. Never decrease or reuse published field values. Treat the public format string as opaque outside compatibility tooling.

## Record a decision

Add `.collaboration/changes/<name>.json` for each change to Core or Pro implementation. Use `impact: no-impact`, `compatible`, or `migration-required`. Include `before`, `after`, `fields`, `reason`, and executable evidence paths in `tests`. A format change also requires `changeset` and `migration` paths. Use a minor or major Pro changeset with a "Breaking collaboration upgrade" notice. Merged records remain immutable when Changesets consumes their release notes.

Run `bun run collaboration:check --base origin/main` before merging. The check protects existing records, catalog entries, and fixture files. It checks decisions against the actual base and candidate formats and the fields that change. It validates evidence paths and requires migration documentation for changed formats. Reviewers must check that the selected fields match the behavior change.

## Verify release compatibility

Build packages and notices before running `bun run collaboration:test`. The matrix packs the built candidate and installs it outside the workspace. Each release uses its recorded `package-lock.json` with `npm ci --ignore-scripts`. The harness requires exactly one Core and one Yjs installation per process.

The catalog starts at stable release `0.3.1`. Every stable Pro release from that baseline must have an entry. The verifier checks registry integrity, source commits, installation locks, and fixture hashes. It checks all baselines, even when formats differ. A registry failure or missing fixture fails the check.

The matrix restores each historical room with its published package. It checks content against the captured canonical fingerprint and paragraphs. It inspects saved state without mutation and refuses incompatible admission. It exports incompatible rooms through their original package and seeds a new candidate room. The new room must have empty undo history and accept further edits. Equal formats run three-peer mixed-client scenarios in both seed directions. The candidate also runs against itself. The matrix checks exact text preservation, undo and redo results, and reordered incremental updates. Migration must preserve the captured canonical package fingerprint, including non-text content. These scenarios do not replace the text, formatting, and table fuzz suites.

The stable release job runs this matrix against its downloaded publish artifacts immediately before publication. It must not rebuild packages between this check and publication. Canary publications do not create stable baselines.

## Capture a published release

After publication, the post-release workflow calls the baseline capture workflow. It creates a pull request with the release-pal App token so normal CI runs. Review and merge that pull request before the next stable release. The workflow does not republish packages or automatically merge evidence.

For local capture, run `bun run collaboration:catalog --capture X.Y.Z`. Capture releases in chronological order from actual npm packages. The command checks package `gitHead` against tag `vX.Y.Z`. It records registry integrity, source lock hash, exact installation lock, and a synthetic saved room. The synthetic fixture contains reversible editing history. The baseline worker never imports workspace packages. The initial release predates format inspection exports, so its worker reads its legacy schema metadata.

Post-release capture waits for all three published package records before it creates a branch. All package probes share a ten-minute deadline. Requests have a maximum thirty-second timeout. Retries use increasing delays up to thirty seconds and respect `Retry-After` without extending the deadline. Temporary HTTP errors, connection failures, and interrupted response bodies retry. Authentication failures and invalid metadata fail immediately. Registry reads prefer online metadata to avoid stale cache entries.

Captured files cannot be overwritten. Retries reuse the existing capture branch and validate existing fixture hashes. Use the **Capture collaboration baseline** manual workflow to retry a failed post-release capture. A capture failure occurs after publication; do not republish the release.

The design follows DOCX Editor's collaboration compatibility policy and release harness. Source revision: `13f08556889ef7c192023f92aa865cc9cc8e3fe4`.
