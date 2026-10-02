# Backlog operating standard

[Open backlog](https://github.com/lagrafx/watermark-python/issues?q=is%3Aissue%20is%3Aopen%20repo%3Alagrafx%2Fwatermark-python) · [All issues](https://github.com/lagrafx/watermark-python/issues?q=is%3Aissue%20repo%3Alagrafx%2Fwatermark-python) · [Propose an item](https://github.com/lagrafx/watermark-python/issues/new?template=backlog-item.md)

GitHub Issues are authoritative. This document is process guidance and a filtered entry point, not a duplicate issue ledger. Use the backlog-item issue template for new work. The permanent ID is the repository-qualified GitHub issue number and canonical URL; retain both when linking to another system.

## Required properties

Every item records type (bug, enhancement, technical debt or investigation); provisional planning level (SAFe Feature or Story) and parent if applicable; user need/business benefit; scope and exclusions; acceptance criteria/tests; evidence/reproduction/confidence; suggested priority, dependencies, risks and owner; human approval status/approver/date/evidence; delivery status/claim/branch/PR; and ServiceNow ID/link/mapping status.

Stories include user-story wording and team-estimated points. Features include a benefit hypothesis and measurable outcome, with child Stories linked only when agreed. Planning level is provisional until refinement. Do not invent a Feature parent, owner, estimate, approval, or ServiceNow mapping.

## Refinement and approval

New items start with Approval **Needs review**, Delivery **Proposed**, owner/claim/estimate/approver/date unset. Suggested priorities are recommendations, not approvals. A human reviews scope, tests, risk and dependencies, confirms Feature/Story classification, and records the approval decision with approver, date and evidence. Approval and delivery status are separate. Creating an issue, assigning a suggested priority, or approving backlog administration does not authorize implementation.

Suggested approval states: Needs review, Approved, Changes requested, Rejected.
Suggested delivery states: Proposed, Ready, In progress, Blocked, In review, Done, Cancelled.

## Work selection gate

An item is menu-eligible only when explicitly human-approved, Delivery Ready, unclaimed, dependencies satisfied, and free of overlapping active work/PRs. Re-read live issue properties, approval evidence, linked dependencies and current branches/PRs immediately before offering, claiming or starting it. Missing or stale information makes the item ineligible. Record a claim before work begins; resolve competing claims with the human owner. Labels, Markdown fields and templates are conventions, not access controls or enforced approval mechanisms.

Completion requires the agreed acceptance criteria, verification evidence and linked delivery PR/commit. Merge, deployment and other consequential operations still need their own applicable authorization.

## ServiceNow

Keep ServiceNow ID and canonical link unset until a real corresponding record exists. Exact field and workflow mapping remains pending the user's screenshots and agreement. Do not invent synchronization, workflow enforcement, credentials, or platform IDs.

## Review integrity

Source-review findings distinguish verified control-flow behavior from hypotheses and hardening opportunities. Link evidence, state which tests were actually run, and check existing issues/PRs before filing. Never include credentials or sensitive operational data. Retain rejected/completed decisions in their original issue history.
