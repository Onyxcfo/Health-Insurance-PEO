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
| `adp-questions.html` | Five questions for the ADP meeting, the three quick confirms, and what ADP has already settled. |
| `angle-hcra-forms/` | The three New York DOH Health Care Reform Act forms for the Angle level-funded plan — DOH-4399 Payor Election, DOH-4264 Electronic Filing User ID, DOH-4403 TPA/ASO Status Change — filled for Onyx. Signature and date left blank. |
| `pdf/` | Print-ready PDFs of the three HTML documents. The report is landscape so its wide tables are not clipped. Regenerate with `src/topdf.mjs` after any edit. |
| `Onyx_Angle_Questco_Cost_Model.xlsx` | The live model for the chosen arrangement. Seven tabs: Contributions, Angle Plans, Cost by Tier, By Employee, Questco, 401(k), Combined. Change a contribution or flip an employee to Yes and every figure recalculates. |

## src/

| File | What it is |
|---|---|
| `plans.py` | The plan dataset, transcribed from the vendors' own documents. Coinsurance normalized to the share the plan pays. |
| `build_wb.py` | Builds the plan-comparison workbook from `plans.py`. Formula-driven — it recalculates from the blue input cells. |
| `qsheet.py` | The TriNet diligence questions, with the context and dollar exposure attached to each. |
| `build_q.py` | Builds the question workbook from `qsheet.py`. |
| `fill_forms.py` | Fills the three NY DOH HCRA forms from the values at the top of the file. Re-run it if any of them change. |
| `topdf.mjs` | Renders the three HTML documents to print-ready PDFs. Run it again after editing any of them. |
| `angle_model.py` | Builds the Angle + Questco cost model. |

## Contribution basis

Fixed employer dollars, medical only: $300 employee-only, $600 employee + spouse,
$1,000 family. At four enrollees that is $2,900/month, $34,800/year. Dental, vision,
life and disability are priced separately and are not in the workbook.

## Enrollment

Steven Nikolov (family), Lisa Danforth (employee only), Jessica B. (family),
Christine Johnson (employee + spouse). Josephine Mack is waiving medical, dental
and vision. Elena Nikolov is not on payroll.

## Outcome

Medical placed with **Angle Health** through CapFi, who administer the plan.
PEO services other than health go to **Questco**. The PEO comparison below the
line is the record of how that decision was reached.

Because Angle is level funded, Onyx is the sponsor of a self-funded plan. That
is what brings the HCRA forms, the PCORI filing on Form 720, and Forms
1094-B/1095-B — obligations a PEO medical plan would have absorbed.

### Open on the HCRA forms

- Section III of DOH-4403 still carries Angle's mark on "Previous TPA/ASO will
  continue to process claims," which cannot be true on an initial election.
  Have CapFi reissue with the third box, or strike and initial.
- The pre-printed TPA phone on DOH-4403 is nine digits.
- Confirm 10/01/2026 is the right election date, and whether Adrem submits the
  forms or Onyx mails them.
