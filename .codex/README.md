# Codex Agent Workspace

This directory centralizes project-level agent coordination artifacts.

## Layout

- `project/` — source-of-truth project goal, constraints, roadmap, and execution context.
- `agents/` — role definitions and operating checklists for recommended agents.

## Recommended agents

1. **Implementation Agent** — applies minimal, correct code changes.
2. **QA Agent** — validates behavior, regression risks, and runbook checks.
3. **Security Agent** — checks secrets handling, scopes, auth, and data exposure risk.
4. **Release Agent** — manages deployment readiness and rollback notes.
5. **Ops/Automation Agent** — maintains CI/CD, workflows, and environment setup docs.

## Usage

1. Update `project/goal.md` first when project intent changes.
2. Keep `project/context.md` aligned with current architecture and constraints.
3. Track pending work in `project/roadmap.md`.
4. Use corresponding `agents/*.md` checklist before finishing a task.
