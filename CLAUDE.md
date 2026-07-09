# CLAUDE.md

Claude-specific supplement. All agents should read `AGENTS.md` first; that file
is the canonical workflow for this repository. This file only records
Claude/Claude Code conventions and overrides.

## Worktree Startup Checklist

After reading this file:

1. Run `pwd` and identify the worktree role from the path.
2. Read the package-level `CLAUDE.md` for any package you will edit.
3. Follow the ownership limits below unless the user explicitly asks for a
   cross-package change.

## Claude Worktree Roles

Some Claude sessions run in package-scoped worktrees:

- `.claude/worktrees/pptx` or `ooxml-pptx` → PPTX session; edit
  `packages/pptx` only unless the task requires shared code.
- DOCX-focused worktrees should primarily edit `packages/docx`.
- XLSX-focused worktrees should primarily edit `packages/xlsx`.

Reading other packages is fine. Editing other packages should happen only for a
coherent cross-package fix described in `AGENTS.md`.

## Parallel Session Safety

Multiple Claude/Codex sessions may run at once.

- Treat unexpected dirty files as user or peer-agent work.
- Do not revert or overwrite unrelated changes.
- Preserve dirty work before switching branches or pulling.
- Private samples, local sample stories, WASM outputs, VRT screenshots, and
  reference images remain local/generated unless the user explicitly asks to
  commit them.

## Claude Autonomous Work

Between 01:00 and 09:00 local time, Claude sessions may proceed without asking
for confirmation for non-destructive work:

- code changes,
- tests,
- WASM builds,
- feature-branch commits and pushes,
- package scripts.

Still ask before destructive operations, reference image updates, direct changes
to `main`, or anything that conflicts with `AGENTS.md`.

## Commit Attribution

When Claude/Fable authors a commit, include:

```text
Co-Authored-By: Claude Fable 5 <noreply@anthropic.com>
```

If Claude starts or directs Codex for the work, include both the Claude/Fable
and Codex co-author trailers, as described in `AGENTS.md`.
