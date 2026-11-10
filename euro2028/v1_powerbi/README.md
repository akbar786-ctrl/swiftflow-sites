# EURO 2028 – original-style template → Power BI

Forces fill **Match Day Resource** and **Wider Tournament Resource** on the familiar template
(risk level on row 3, day/measure on row 4, Super High added). Power BI does every calculation.

| `output/` | |
|---|---|
| `EURO_2028_Resource_Collection_Template.xlsx` | Sent to forces. BLANK v1 template + header swap + Super High + hidden named ranges `MD_Data`, `WT_Data`. No calculation pages. |
| `EURO_2028_Power_BI_Reference.xlsx` | Row maps, column maps, rates, PSU package, WT days, fixtures. |
| `EURO_2028_Power_BI_Step_by_Step.txt` | The guide + every Power Query query + every DAX measure. |
| `EURO_2028_Power_BI_TEST_PACK.zip` | Three synthetic returns, sample reference, expected Power BI figures. |

How it works: `fnReadSheet` reads each named range, the row map says what each Excel row is,
`fnBuildFact` builds the costed master tables (`MD_Fact`, `WT_Fact`), `Calendar_Fact` places
Match Day on fixture dates, `DataQuality` lists what to chase.

Proof: `pbi_mirror.py` is the same logic in Python. `validate_against_original.py` runs it on the
original collection workbook's figures entered in this template and compares with the original's
calculation results: officers and cost on every date and the Wider Tournament total – **0
differences** (also with the figures moved to Super High). Power Query files pass Microsoft's
Power Query parser.

Rebuild: `make_forces_template.py`, `make_reference.py`, `make_samples.py`, `make_guide.py`.
