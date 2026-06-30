Use the uploaded workbook `vendor_onboarding_hold_queue.xlsx`.

Populate the `HoldQueue` sheet with final literal values only. Save the completed workbook as `/workspace/vendor_onboarding_hold_queue_completed.xlsx`.

Write results to `HoldQueue!A2:I11` with columns:

`Application ID`, `Canonical Vendor`, `Office Code`, `Matched Root`, `Matched Search Volume`, `Win Count`, `Spend Status`, `Trigger Type`, `Trigger Key`.

Do not write formulas in the output range. Preserve all source sheets and protected notes.

Rules:

1. Collapse duplicate Application ID rows first, keeping only the row with the latest Submitted Date.
2. Include rows whose Status is exactly `Ready` or exactly `Ready - resubmitted`; exclude all other statuses.
3. Canonical Vendor is the trimmed text before the first semicolon in Vendor Alias.
4. Suppress an exact active existing vendor only when ExistingVendors has the same Canonical Vendor and Country with Status exactly `Active`, except SAN risk code overrides suppression.
5. Determine the trigger after suppression:
   - If any manual risk-code rule matches, choose the matching code whose RiskRules row appears earliest.
   - Otherwise use a matching country/category rule.
   - For a resubmitted row with a risk-code trigger, Trigger Type must be `risk_code_resubmitted`; otherwise use `risk_code` or `country_category`.
   - Trigger Key is only the code, such as `EXP`, or the country/category key, such as `ZA|Mining`.
6. Office Code comes from OfficeMap using both Location and Country.
7. Matched Root comes from the `RootWords` sheet using an exact match between `Applications.Category` and `RootWords.Applies To Category`.
   - In `RootWords`, the root token is in column `Root`.
   - In `RootWords`, the category key is in column `Applies To Category`.
   - For the matched category, use the corresponding `Root` value as the root token.
   - Then search `SearchTerms.Keyword` for that root as a complete word, case-insensitive.
   - Sum all matching `Search Volume` values for `Matched Search Volume`.
8. Win Count comes from Scoreboard for Canonical Vendor: count how many game scores equal the maximum score in that game column.
9. Spend Status:
   - `UNKNOWN_SPEND` for blank, missing, or nonnumeric text.
   - `ZERO_SPEND` for numeric zero or text that parses to zero after trimming whitespace and commas.
   - `KNOWN_SPEND` for any other numeric spend.
10. Sort final rows by Application ID ascending.
