# Agent instructions

**Commits:** All commits in this repo must follow [Conventional Commits](https://www.conventionalcommits.org/): use a type and optional scope (e.g. `feat(calendar): add X`, `fix(roster): correct Y`, `chore(deps): Z`), and add a short body when it helps.

For Apps Script–specific rules (version bump, deployment), see [coach-sheet-apps-script/AGENTS.md](coach-sheet-apps-script/AGENTS.md).

**Pre-commit hook:** `.githooks/pre-commit` runs the coach sheet regression harness (`coach-sheet-apps-script/test/harness.js`) against the staged content whenever a `.gs` or `test/` file under `coach-sheet-apps-script/` is staged, and blocks the commit on failure. Hooks are not versioned by git itself, so activate it once per clone with `git config core.hooksPath .githooks`. Do not bypass it with `--no-verify`; fix the code or the assertion.
