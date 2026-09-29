# Repository Instructions

This repository is managed with Engineering Fleet. These instructions extend
the global Fleet doctrine; they do not replace it.

## Architecture

- Codex is the architectural decision gate and final technical reviewer.
- Workers operate only within explicitly assigned scope.
- Record important decisions and their evidence in version control.

## Verification

- Run `fleet verify` before proposing or promoting changes.
- A worker claim is not proof. Treat only recorded deterministic verification
  or identified production evidence as proof.

## Repository-specific constraints

- Autoship is a Google Apps Script project. `Code.js` is loaded by the Apps
  Script V8 runtime; it is not a standalone Node entry point.
- Treat `Code.js` as a high-risk monolith because it can mutate Sheets, Drive,
  Calendar, Canvas-backed intake data, and Script Properties.
- Never run `clasp push`, deploy, install triggers, call live Google/Canvas/
  Gemini services, or mutate credentials without explicit Codex and human
  approval. Field-test validation must use local mocks only.
- Never print or commit values from `.clasp.json`, Script Properties, tokens,
  API keys, document IDs, calendar IDs, or fetched user data.
- Preserve `appsscript.json` services and scopes unless a task explicitly
  authorizes an architecture-sensitive permission change.
- Keep worker changes inside the Fleet task's named files and prohibited-area
  boundary. No broad refactor of `Code.js` during a bounded repair.
- GitHub issues are the canonical Fleet task records. Every verification,
  worker result, PR, CI result, and Codex decision must bind to an exact SHA.
- Do not merge PRs or deploy from this repository as part of the field test.

## Local commands

- `npm run lint` checks JavaScript and inline sidebar syntax.
- `npm test` runs dependency-free deterministic tests with mocked boundaries.
- `npm run build` validates the Apps Script manifest.
- `node /Users/sadius/engineering-fleet/dist/cli.js verify` runs the canonical
  Fleet validation sequence used by CI.
