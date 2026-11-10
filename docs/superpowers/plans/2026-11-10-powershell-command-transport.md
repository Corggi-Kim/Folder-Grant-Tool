# PowerShell command transport implementation plan

> **For agentic workers:** Use superpowers:executing-plans to implement the tasks below with a failing regression test before production changes.

**Goal:** Run every validation and permission script without putting the full script on the Windows command line, and distinguish a launch-length failure from a missing executable.

**Architecture:** Use one short, fixed encoded bootstrap for both PowerShellRunner paths. Send the existing UTF-8-wrapped script as base64 ASCII on stdin after the existing FGT_START gate and an exact payload-length header; reject missing or incomplete input before script execution. The cancellable path continues to assign the process to its Job Object before its collector sends any input; cancellation monitoring and confirmed tree shutdown stay independent of pipe collection.

**Tech stack:** Python subprocess/threading, Windows PowerShell 5.1 / PowerShell 7; no new runtime dependencies.

**Spec:** The reported Study / STAT / Manager validation command exceeds Windows' 32,767-character limit. The user requests a common fix for every folder and Role, rather than a per-folder exception.

## Constraints and review focus

- Preserve exact script Unicode, quoting and output/error behavior.
- Keep the start gate closed on stdin EOF and never execute before process containment.
- Preserve cancellation, timeout, single pipe ownership and child-process cleanup.
- Do not change folder policies, preflight scope, BUS completion, job retry or server PowerShell selection.
- Preserve Windows error information; do not label WinError 206 as a missing executable.

## Tasks

- [x] Add a Windows-length-limited process fixture that rejects command lines >=32,767 characters; reproduce the real project 26066 / Study / STAT / Manager validation and large Unicode scripts in normal and cancellable execution.
- [x] Add launch-error regressions for WinError 206 versus missing files, plus real-interpreter UTF-8, failure, timeout, cancellation and gate EOF coverage.
- [x] Run the new regressions and record their expected failures.
- [x] Replace source-bearing command arguments with the common bootstrap and stdin payload, retaining the existing containment/monitor lifecycle.
- [x] Run focused regressions, the full API suite and real PowerShell interpreter tests; obtain a final review of process ownership and transport behavior.
- [x] Update API version to 0.4.7 and current text release/operating guidance. Keep Client 1.1.6 and existing Office files unchanged.
- [x] Commit and create a reviewable PR with exact validation results and Windows field-test limits.

## Evidence and decisions

- Actual reported Study/STAT/Manager validation produces a 33,132-character Windows command line. CPython maps WinError 206 to FileNotFoundError, which the old runner mislabeled as a missing executable.
- The HTTP regression matrix initially produced 4 expected failures / 11 passes, including long custom share and template paths.
- Reject incomplete stdin using the declared base64 payload length, preserving fail-closed startup even at a valid base64 truncation boundary.
- Use ASCII base64 stdin rather than temporary scripts to avoid new disk permissions, source files and cleanup requirements.

## Verification

- New unit regressions: 21 expected failures / 7 passes before implementation, including large-source launch, partial stdin and launch-error classification. Added two invalid-name error cases also failed before implementation.
- New focused suite: 45 passed after implementation.
- Full API suite: 297 passed, 1 native Windows-only test skipped, 5 existing deprecation warnings. Full Client suite: 175 passed.
- Independent final review found no blocking or important issues. Native Windows and Windows PowerShell 5.1 execution remain field-test requirements; no production AD/SMB operations were executed here.

- Final API rerun after version/documentation changes: 297 passed, 1 skipped, 5 warnings in 35.27s. FastAPI TestClient needs working local socket wakeups; use the network-enabled test environment.
