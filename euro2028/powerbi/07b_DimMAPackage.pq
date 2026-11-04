// ============================================================================
// DimMAPackage (load: yes) - mutual-aid PSU package rate at RateEffectiveDate
// (base x MAPSU uplift steps). Null BasePackage (serials) stays null -> DQ14.
// ============================================================================
let
    Params  = Record.FromList(RefParameter[Value], RefParameter[Parameter]),
    EffDate = Date.From(Params[RateEffectiveDate]),
    F       = List.Product(List.Combine({{1}, Table.SelectRows(RefUpliftStep,
                  each [RateTable] = "MAPSU" and Date.From([EffectiveFrom]) <= EffDate)[Factor]})),
    Result  = Table.AddColumn(RefMAPackage, "PackageRate",
                  each if [BasePackage] = null then null else [BasePackage] * F, type nullable number)
in
    Result
