# Autoship Fleet field-test baseline

Assessment date: 2026-09-23
Baseline commit: `b2513c1d2efc4681f41838fb1cc0e2a8a68a44c7`

## Repository health

Autoship is a paused Google Apps Script scholarship workflow. Its production
implementation is concentrated in a roughly 2,550-line `Code.js`, with one
HTML sidebar and an Apps Script manifest. The script integrates with Google
Sheets, Drive, Calendar, Canvas, Gemini, and Apps Script triggers. The default
branch was clean at assessment time.

The repository had no package manifest, automated test harness, or GitHub
Actions workflow. `node --check Code.js` passed, but syntax validation alone
cannot detect behavioral defects in Apps Script flows. No existing workflow
can conflict with Fleet's generated workflow.

## Confirmed and suspected problem areas

- The Canvas scanner accumulates matching rows but never calls its existing
  persistence helper, so a successful-looking scan discards its discoveries.
- `basicExtractFromText_` is declared twice, so the later declaration silently
  overrides the earlier implementation.
- `pickCourses_` contains a duplicated, unreachable return statement.
- The checked-in Copilot guidance describes `Code.js` as a Node entry point,
  although it is an Apps Script global file.
- Several integrations have real external side effects and cannot be exercised
  safely without mocks and explicit approval.

These are candidate findings, not proof that every affected user workflow is
broken. Each selected repair must obtain deterministic evidence on its own
task branch.

## Safety boundary

The field test will not deploy, run `clasp push`, install triggers, call live
Google/Canvas/Gemini services, mutate production data, or expose identifiers
and credentials. `.clasp.json`, `appsscript.json`, CI configuration, and the
monolithic production script are risk-sensitive. All product work will use
dedicated branches and PRs; `main` remains unchanged.

## Fleet acceptance evidence

The repeated A/B commits on `fleet/field-test-bootstrap` exercise only Fleet's
exact-SHA verification and review-staleness behavior. This repository has no
configured Phase 4 provider route, so those commits are not evidence for
provider resolution, dispatch, context hashing, or artifact provenance. Those
paths are covered by Engineering Fleet's separate deterministic architecture
acceptance fixture.
