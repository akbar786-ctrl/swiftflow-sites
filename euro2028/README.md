# EURO 2028 – Resource Data Collection and Power BI Model

A simplified force collection template and a Power BI model that replaces the
background calculation sheets of the legacy *Base Assumptions MPSTH – Uniform*
workbook.

> Classification: the template and outputs are designed for **OFFICIAL** use.
> No force data is stored in this repository. `output/samples/` contains a
> **synthetic** submission for testing only.

## What is here

| Path | Contents |
|---|---|
| `output/E28_Collection_Template_v1.0.0.xlsx` | Force collection template (Cover, Guidance, MD Collection, WT Collection, optional Command Team / Other Resources). |
| `output/E28_Reference_Tables.xlsx` | Every reference table the model loads: resources, ranks, rates, uplifts, parameters, old-to-new mapping, decisions, open items, and empty tables for fixtures, capacity, equipment and scenarios. |
| `output/samples/SYNTHETIC_Submission_TEST_FORCE.xlsx` | Synthetic completed return with deliberate data-quality errors. |
| `powerbi/*.pq` | Power Query M, one file per query, numbered in dependency order. |
| `powerbi/measures.dax` | DAX measures. |
| `engine/e28_engine.py` | Python reference implementation of the same rules (used for testing and reconciliation). |
| `tests/` | Unit tests (`python3 -m unittest discover -s euro2028/tests -v`). |
| `build/` | **Source of truth.** `catalogue.py` (resources), `reference_data.py` (rates, parameters, decisions), `legacy_mapping.py`, and the builders. |
| `docs/` | Audit, calculation specification, data model, validation plan. |

## Rebuilding

```bash
python3 euro2028/build/build_all.py        # template + reference workbook
python3 euro2028/build/build_sample.py     # synthetic submission
python3 -m unittest discover -s euro2028/tests -v
```

To add, rename or retire a resource, edit `build/catalogue.py` (and
`build/legacy_mapping.py`), then rebuild. Resource IDs are stable keys – never
reuse an ID for a different resource.

## Documents

1. [Stage 1 – Audit of the legacy workbook](docs/01_audit.md)
2. [Stage 2/3 – Calculation specification](docs/02_calculation_spec.md)
3. [Stage 3 – Power BI data model and build steps](docs/03_data_model.md)
4. [Stage 4 – Validation strategy and results](docs/04_validation.md)

## Status

Decisions D-01 to D-11 and open items OI-01 to OI-08 are listed in
`docs/02_calculation_spec.md` and in the reference workbook (`LogDecision`,
`LogOpenItem`). The model runs with documented defaults for every open item;
each default is a single parameter or reference-table change.
