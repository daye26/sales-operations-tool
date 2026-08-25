# SQL portfolio examples

These examples are derived from SQL workflows used to inform parts of the application logic. They are intentionally anonymized: generic schemas, identifiers, and fictional data replace operational names, local paths, data, and company-specific exceptions. They do not connect to, query, or reproduce an operational environment.

The example queries are read-only `SELECT` statements written for DuckDB. The setup script creates only fictional demonstration tables in the active database.

## Running the examples

1. Run `00_demo_schema_and_data.sql` in an empty DuckDB database.
2. Run any of the numbered query files.

The setup script can be run again to reset the fictional sample data.

## Examples

- `00_demo_schema_and_data.sql`: fictional schema and sample data used by the examples.
- `01_normalize_material_codes.sql`: basic data cleaning and validation with CTEs and `CASE` expressions.
- `02_validate_registrations.sql`: medium-complexity record validation and deduplication with a window function.
- `03_rank_allocation_candidates.sql`: multi-stage allocation candidate ranking with CTEs, joins, scoring rules, and window functions.
- `04_data_quality_audit.sql`: independent validation checks combined with `UNION ALL`.
- `05_monthly_operations_summary.sql`: monthly operational KPIs built from several source tables.
- `06_allocation_exceptions.sql`: orders and vehicles that require manual review after allocation checks.
- `07_material_code_substitution_candidates.sql`: compatible alternative material codes ranked by product attributes and availability.
- `08_leads_funnel_analysis.sql`: daily lead funnel by reporting model, derived from `leads.sql`.
- `09_free_cars_reconciliation.sql`: consistency checks for free-car records, derived from `not_allocated.sql`.

The examples are designed as portfolio material: table names, identifiers, values, and rules are anonymized. They preserve the analytical objective of the original workflows without exposing operational data, local paths, confidential identifiers, or company-specific exceptions.

The allocation-ranking example intentionally resolves only first-choice conflicts. Cases that need a fallback allocation are surfaced by `06_allocation_exceptions.sql` for manual review rather than hidden by the example.
