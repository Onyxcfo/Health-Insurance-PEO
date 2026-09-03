# Onyx Accounting Group — 2026 benefits vendor evaluation

Medical effective 10/1/2026. Vendors under evaluation: ADP TotalSource, TriNet,
Insperity (PEO) and Angle Health / CapFi (direct level-funded).

## deliverables/

| File | What it is |
|---|---|
| `onyx-medical-placement.html` | The report for Steven: verdict and rankings, employer-only cost, renewal exposure, what each plan costs each employee, employee-by-employee notes, added value, 401(k), method and open items. |
| `Onyx_2026_Medical_Plans_ADP_TriNet_Angle.xlsx` | Every plan offered by ADP, TriNet and Angle — 30 plans — with design specs, premiums by tier, Onyx's contribution, the employee's payroll deduction and the worst-case annual exposure, each in its own column. |
| `trinet-questions.html` | Twenty questions for the TriNet meeting, ordered by dollar impact, with what we already hold against each and space to write the answer. |
| `Onyx_TriNet_Questions.xlsx` | The same diligence list as a working spreadsheet — 26 rows, fill-in columns for the answer, whether it came in writing, and status; a Summary tab tallies progress and separates the kinds of dollar exposure. |

## src/

| File | What it is |
|---|---|
| `plans.py` | The plan dataset, transcribed from the vendors' own documents. Coinsurance normalized to the share the plan pays. |
| `build_wb.py` | Builds the plan-comparison workbook from `plans.py`. Formula-driven — it recalculates from the blue input cells. |
| `qsheet.py` | The TriNet diligence questions, with the context and dollar exposure attached to each. |
| `build_q.py` | Builds the question workbook from `qsheet.py`. |

## Contribution basis

Fixed employer dollars, medical only: $300 employee-only, $600 employee + spouse,
$1,000 family. At four enrollees that is $2,900/month, $34,800/year. Dental, vision,
life and disability are priced separately and are not in the workbook.

## Enrollment

Steven Nikolov (family), Lisa Danforth (employee only), Jessica B. (family),
Christine Johnson (employee + spouse). Josephine Mack is waiving medical, dental
and vision. Elena Nikolov is not on payroll.
