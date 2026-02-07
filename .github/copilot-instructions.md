# Copilot / AI agent instructions for this repository

Purpose
- Help AI coding agents quickly become productive in this repo by listing concrete, discoverable patterns and workflows.

Quick context (what to look for first)
- Check for a top-level entry file: `Code.js` (seen at c:/Users/kevin/autoship/Code.js).
- If present, run it with `node Code.js` to see runtime behavior when no `package.json` exists.
- If `package.json` exists, prefer `npm install` then `npm test` / `npm start`.

Architecture & big-picture hints (how to discover)
- Identify major components by searching for folders named `src`, `lib`, `routes`, `components`, or `services`.
- Determine module format by scanning for `module.exports` / `require()` (CommonJS) or `export` / `import` (ESM).
- Look for an HTTP server (express/koa) by searching `require('express')` or `from 'express'` — this signals network boundaries.

Patterns & conventions to detect (repo-specific guidance)
- Prefer discovering concrete patterns rather than assuming them. Examples to extract automatically:
  - If files use `module.exports`, implement changes using CommonJS style to remain consistent.
  - If you find `async/await` heavily used, follow that async style for new functions and tests.
  - If a `config` or `.env` file is present, read configuration via that pattern; avoid hardcoded secrets.

Tests, build, and debug workflows
- If `package.json` exists: run `npm run test` (or inspect `scripts` for the test/build scripts). Use `npm run <script>` as defined.
- If no `package.json`, run relevant single-file commands such as `node Code.js` or `node ./src/index.js`.
- To debug locally, prefer `node --inspect-brk` on the main entrypoint and attach a debugger.

Integration points & external dependencies
- Search for `require(` or `import(` to list external npm packages; treat these as external dependencies to preserve in edits.
- Look for HTTP clients, DB clients (e.g., `pg`, `mongoose`), or SDK usage to understand external integrations.

How to generate changes safely
- Always run the repo's tests (if present) before proposing changes.
- When adding or updating code, follow the repository's module style (CommonJS vs ESM) and async style (`callbacks`, `Promises`, or `async/await`).
- Prefer minimal, focused edits: change only the files necessary to implement the requested behavior.

What to include in PR descriptions
- One-line summary of intent.
- Files changed and why (high-level).
- Any manual steps to validate (commands and expected results).

Merging with an existing `.github/copilot-instructions.md`
- Preserve any custom troubleshooting steps or credential instructions already present.
- Merge by keeping repo-specific commands and replacing generic guidance with the concrete steps above.

Examples from this repo (concrete snippets to look for)
- Entrypoint example: `Code.js` — run with `node Code.js` if no package.json exists.
- Module detection: search for `module.exports` vs `export default` to decide code style.

When you are missing information
- Add a short TODO block in the PR describing what you couldn't determine (e.g., test command, CI hooks, DB credentials) and how to verify the change.
- Ask the maintainer for the missing runtime command or CI details.

Last step
- After writing or updating files, run tests (if available) and include the results in the PR body.

If anything here is unclear or you want more repo-specific rules, tell me what files to inspect and I'll update this document.