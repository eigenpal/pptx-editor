# Changesets

Every code change to a published package needs a changeset. Run `bun changeset`, or write `.changeset/<name>.md` by hand. Published packages are one fixed group: declare one bump and the others follow.

Write one short paragraph of at most 400 characters. Say what changed for the developer, and name new or changed API identifiers and breaking changes. Put details in the documentation, not the changeset. The release PR shows every changelog only while its body stays under 60,000 characters, and `scripts/release-plan.test.ts` checks both limits.
