# EURO 2028 – two-page collection form + Power BI (v2)

Forces fill in **one workbook with two pages**; every calculation happens in Power BI.

| File (`output/`) | For |
|---|---|
| `EURO_2028_Resource_Collection_Form.xlsx` | The form sent to forces: **MD Collection** and **WT Collection** pages |
| `EURO_2028_Reference_Data.xlsx` | Kept by the national team: rates, mutual aid PSU package, WT days, fixture list |
| `EURO_2028_Power_BI_Setup.txt` | Step-by-step set-up, every Power Query query, every DAX measure |

`samples/` holds three **synthetic** test returns, a reference file with test fixtures and
`EXPECTED_Power_BI_results_SAMPLE.xlsx` – the figures Power BI should show for them.

## The form
* Each page is one Excel table (`tblMatchDay`, `tblWiderTournament`; force details in `tblDetails`)
  with one header row, so Power BI reads it directly – no merged cells inside the data.
* Every line is described by **Section | Resource | Supplied by | Role / Skill | Rank**.
  Rank is always a rank; skills (ARV, CBRN Search, PO Bronze Commander …) and PSU unit types
  (Level 2 PSU / Serial …) are in Role / Skill.
* Numbers: **Officers**, **Normal hrs** and **Extra hrs** (per officer) for each risk level
  (Low / Medium / High / Super High) and, on the Match Day page, each day MD-2 … MD+1.
* Colour key: white = type a number, yellow = choose from the list, grey = not needed (locked).
  Drop-downs live in the cells (no hidden sheet); numbers and hours are checked on entry.
* PSUs / Serials: a "Number of PSUs / Serials" line (units) plus Inspector / Sergeant / Constable
  lines with the total officers in those units; mutual aid PSUs give the number of units only.
* Non-matchday PSU lines lock the Match Day columns; matchday-only lines lock the other days.

## Power BI calculations (original collection workbook rules)
* Overtime cost = Officers × Extra hrs × hourly rate (Supplied by + Rank; resource overrides win).
* Force-supplied Inspector and above: not costed. In Stadium PSU Inspectors: mutual aid Inspector rate.
* Mutual aid PSU / Serial = units × package per unit per day.
* Normal duty cost reported separately ("incl Normal Duty" measures).
* Fixture calendar: MD-2 … MD+1 around each fixture; one day per date and risk level
  (MD > MD-1 > MD-2 > MD+1; non-matchday PSU lines MD-1 > MD-2 > MD+1).
* Wider Tournament total = typical day × days at each risk level (7 / 11 / 11 / 3).

## Rebuilding and checks
```
python3 make_form.py                      # blank form
python3 make_reference.py                 # reference data
python3 make_powerbi_pack.py              # set-up text file
python3 reference_calc.py samples/returns samples/EURO_2028_Reference_Data_SAMPLE.xlsx samples/EXPECTED_Power_BI_results_SAMPLE.xlsx
```
`reference_calc.py` repeats the Power BI steps in Python to give the expected figures. The Power
Query files were syntax-checked with Microsoft's Power Query parser (`@microsoft/powerquery-parser`).

## Needs confirmation
* Whether force-supplied Inspectors and above should stay at £0 (original rule).
* Whether the mutual aid package applies to Level 3 BDUs / Serials as well as Level 2 PSUs (original workbook did).
* Police staff: no staff rate exists, so the form offers police ranks only.
