-- DuckDB-compatible example.
-- Expected source table:
-- raw_material_codes(material_code, product_series, trim_level, interior_color, exterior_color)

WITH normalized_values AS (
    SELECT
        NULLIF(UPPER(TRIM(material_code)), '') AS material_code,
        NULLIF(UPPER(TRIM(product_series)), '') AS product_series,
        NULLIF(UPPER(TRIM(trim_level)), '') AS trim_level,
        NULLIF(UPPER(TRIM(interior_color)), '') AS interior_color,
        NULLIF(UPPER(TRIM(exterior_color)), '') AS exterior_color
    FROM raw_material_codes
),
standardized_colors AS (
    SELECT
        material_code,
        product_series,
        trim_level,
        CASE
            WHEN interior_color IN ('DARK', 'BLACK') THEN 'DARK'
            WHEN interior_color IN ('LIGHT', 'BEIGE') THEN 'LIGHT'
            ELSE interior_color
        END AS interior_color,
        CASE
            WHEN exterior_color IN ('GREY', 'GRAY', 'SILVER') THEN 'SILVER'
            WHEN exterior_color IN ('WHITE', 'PEARL WHITE') THEN 'WHITE'
            ELSE exterior_color
        END AS exterior_color
    FROM normalized_values
),
validated_codes AS (
    SELECT
        *,
        CASE
            WHEN material_code IS NULL THEN 'MISSING_MATERIAL_CODE'
            WHEN LENGTH(material_code) < 4 THEN 'INVALID_MATERIAL_CODE'
            WHEN product_series IS NULL OR trim_level IS NULL THEN 'INCOMPLETE_PRODUCT_DATA'
            ELSE 'VALID'
        END AS validation_status
    FROM standardized_colors
)
SELECT
    material_code,
    product_series,
    trim_level,
    interior_color,
    exterior_color,
    validation_status
FROM validated_codes
ORDER BY validation_status, material_code;
