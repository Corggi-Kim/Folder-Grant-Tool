# Active Job Cancellation Implementation Plan

> **For agentic workers:** Use superpowers:executing-plans to implement task by task.

**Goal:** Stop the currently executing PowerShell and its descendants after an authenticated Job cancellation, preserve partial results, and never complete BUS for a stopped operation.

**Architecture:** Poll the persisted cancel flag while a worker-owned process runs. Windows Job Objects contain descendants before releasing a stdin startup gate; POSIX process groups provide equivalent test containment. Cancellation is terminal only after process termination is confirmed; failure to confirm is an executor failure, never success.

**Tech Stack:** Python subprocess/ctypes, SQLAlchemy, FastAPI, existing PyQt5 client.

**Spec:** User approved active process-tree cancellation and immediate merge in this conversation on 2026-11-02. Already applied changes are not rolled back. The parent agent implements inline; one reviewer checks before merge.

## Global Constraints

- No process-name-wide termination, new dependencies, DB schema change, or Office regeneration.
- Retain existing read-only validation behavior and normal UTF-8/timeout handling.
- Running access, project and group-preparation Jobs share cancellation semantics.
- API polling and UI stop controls must remain responsive and avoid duplicate requests/messages.

## Review Focus

- Descendants survive parent exit: process-tree containment must remain until cleanup.
- Windows containment fails: command stays behind its startup gate and must not mutate.
- Stop arrives during the last step: no subsequent success/BUS result.
- Completed steps remain succeeded; interrupted and queued steps are cancelled.
- A cancellation or timeout cannot kill another concurrent Job.

### Task 1: Controlled process execution

Files: API/infrastructure/powershell.py, new process_tree.py, tests/unit/test_active_cancellation.py (paths under src/folder_grant_api for production).

- [x] Add failing tests for active stop, descendant cleanup, timeout cleanup, pre-start stop and independent concurrent processes.
- [x] Add optional cancel_check constructor argument to PowerShellRunner; preserve run(command) interface. Use Popen with a stdin startup gate, 0.5-second communicate/persisted-cancel polling, bounded cleanup and dedicated cancellation exception.
- [x] Add Windows Job Object management with kill-on-close and fail-closed assignment; POSIX uses isolated process groups.
- [x] Verify process tests and existing runner tests.

### Task 2: Persisted Job cancellation

Files: workers/powershell_worker.py, tests/api/test_active_cancellation.py.

- [x] Write failing tests using queued registration, real worker thread and authenticated cancel endpoint, plus final-step cancellation race.
- [x] Query cancellation through fresh DB sessions; handle dedicated interruption separately from failure. Preserve succeeded steps and audit cancellation; stop before final success if cancellation is pending.
- [x] Verify access/project behavior, partial result retention and absence of later steps.

### Task 3: Client feedback and release

Files: Client .pyw/tests, API package metadata, current Markdown documentation.

- [x] Test duplicate Stop calls log once and disable Stop while server shutdown is pending; stopped results never complete BUS.
- [x] Implement stop-in-progress state and clear it after worker completion; preserve final state polling.
- [x] Release Client 1.1.3 Build:20261102 / API 0.4.4; update current help/README and release notes.
- [x] Run full Client and API suites, inspect Office diff, request review, resolve findings, commit/push, create and merge PR as authorized.

Verification: Client 152 passed; API 143 passed, native Windows smoke 1 skipped, 4 existing warnings. Reviewer findings on EOF gate and final-success race were fixed and regression-tested. Native Windows acceptance remains explicitly recorded in the open checklist.
