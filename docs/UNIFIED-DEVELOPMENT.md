# One OMNIX development and release line

`main` is the canonical implementation. Further work must branch from current `main`, use a reviewed pull request and pass its existing exact-source gates. Do not publish new installers from `rebuild/native-v4` or replace main with that branch. The historical branch, tags and recovery archive remain available for recovery; PR #147 is superseded by this integration.

The v4.0.0-preview.1–3 artifacts were built from the smaller native rebuild. The v4.0.0-preview.4 artifact came from the independently developed main implementation. Those binaries are not interchangeable builds of one source tree. Subsequent unified previews use main, its complete-source ZIP, embedded payload identity, manifest and CI provenance. Version numbers alone are not evidence of lineage or acceptance.

## Reconciliation

| Rebuild behavior | Canonical implementation |
| --- | --- |
| Automatic window workspace | Ported into all three existing task-pane services, with deferred initial startup and explicit window ownership. Existing hidden panes stay hidden. |
| Idempotent Open Workspace | Ported into all three hosts. Use the native pane close control to hide it. |
| Failed pane creation cleanup | Controller and host control are disposed on failed attachment. Existing post-close lifecycle cleanup stays in place. |
| Request cancellation and stale document isolation | Preserve main's request-scope checks, controller disposal and compiled tool-isolation tests. |
| Separate custom undo stack | Do not copy it: main deliberately uses Office-native undo through its adapters. |
| Separate named-pipe gateway | Do not mix incompatible protocol/storage trees: keep main's provider gateway and budget/continuity engine. |
| Encrypted chat history | Preserve main's DPAPI storage, migration and corruption checks. The rebuild's different history schema is retained in its original data directory. No silent migration or deletion. |
| Runtime checks | Recognize v4 and v4R; bundle Microsoft's current published redistributable with signature verification; compile and execute the production restart decision in CI. An ambiguous exit-0 result gets one recovery restart, then a repair/diagnostic failure instead of an endless restart loop. |
| Actual Office test | Preserve main's bound E2E harness and add visible automatic-workspace proof before any Open Workspace click. |

Keep the original install directories until a verified migration is performed. The two generations use different installer AppIds. Do not run their uninstallers during an active Office session; unregistration ownership must be reviewed before removing an older generation.

No installation, Office UI, reboot, live-provider, consumer-security or trusted-signing PASS has been produced by this reconciliation alone. Production approval remains blocked by the existing real-evidence gates.
