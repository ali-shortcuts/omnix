# Office write access diagnosis

Based on main 784b07846990801f82df67f09ada61d51a286354; preserves the user's expanded Office capability engine.

A model's statement that write access is unavailable is not an Office permission result. The screenshot does not identify the original COM/provider error. This update:

- Inspects live host document/read-only/protection state and includes it in each model turn. Excel distinguishes workbook structure protection from active-sheet protection and reports whether table creation passes its preflight.
- Exposes read_office_access through the same document/cancellation boundary; reports presence of the write-confirmation handler without modifying protection.
- Makes at most one corrective provider turn when an unsupported access denial appears before any write attempt. It obtains measured state first; cancelled or failed writes are never automatically retried by this recovery path.
- Logs tool dispatch/result metadata without arguments, document contents or API keys.
- Uses the final gateway answer for the saved/displayed assistant turn; previously streamed planning or a denial could replace a later authoritative answer. Late streaming callbacks are ignored after completion.

The existing capability registry remains available to every configured provider through OMNIX's protocol. This is not unrestricted access to every Office feature, nor proof that every catalog entry works on every Office version. Native protection, preview approval, active-document scope and cancellation remain enforced. Office integration must still be tested in installed Excel/Word/PowerPoint; hosted build tests cannot prove COM behavior on the user's machine.
