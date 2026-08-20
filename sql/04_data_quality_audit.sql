-- DuckDB-compatible example.
-- Produces one row per detected issue so multiple errors remain visible.

WITH material_issues AS (
    SELECT
        'MATERIAL_CODE' AS source_area,
        COALESCE(NULLIF(UPPER(TRIM(material_code)), ''), '<MISSING>') AS record_id,
        'MISSING_MATERIAL_CODE' AS issue_type
    FROM raw_material_codes
    WHERE NULLIF(TRIM(material_code), '') IS NULL

    UNION ALL

    SELECT
        'MATERIAL_CODE' AS source_area,
        UPPER(TRIM(material_code)) AS record_id,
        'INVALID_MATERIAL_CODE' AS issue_type
    FROM raw_material_codes
    WHERE NULLIF(TRIM(material_code), '') IS NOT NULL
        AND LENGTH(TRIM(material_code)) < 4

    UNION ALL

    SELECT
        'MATERIAL_CODE' AS source_area,
        UPPER(TRIM(material_code)) AS record_id,
        'INCOMPLETE_PRODUCT_DATA' AS issue_type
    FROM raw_material_codes
    WHERE NULLIF(TRIM(material_code), '') IS NOT NULL
        AND (
            NULLIF(TRIM(product_series), '') IS NULL
            OR NULLIF(TRIM(trim_level), '') IS NULL
        )
),
registration_issues AS (
    SELECT
        'REGISTRATION' AS source_area,
        COALESCE(NULLIF(TRIM(registration_id), ''), '<MISSING>') AS record_id,
        'MISSING_VEHICLE_ID' AS issue_type
    FROM raw_registrations
    WHERE NULLIF(TRIM(vehicle_id), '') IS NULL

    UNION ALL

    SELECT
        'REGISTRATION' AS source_area,
        COALESCE(NULLIF(TRIM(registration_id), ''), '<MISSING>') AS record_id,
        'UNKNOWN_VEHICLE' AS issue_type
    FROM raw_registrations AS registration
    LEFT JOIN vehicle_catalog AS catalog
        ON catalog.vehicle_id = UPPER(TRIM(registration.vehicle_id))
    WHERE NULLIF(TRIM(registration.vehicle_id), '') IS NOT NULL
        AND catalog.vehicle_id IS NULL

    UNION ALL

    SELECT
        'REGISTRATION' AS source_area,
        COALESCE(NULLIF(TRIM(registration.registration_id), ''), '<MISSING>') AS record_id,
        'UNKNOWN_DEALER' AS issue_type
    FROM raw_registrations AS registration
    LEFT JOIN dealer_directory AS directory
        ON directory.dealer_code = UPPER(TRIM(registration.dealer_code))
    WHERE NULLIF(TRIM(registration.dealer_code), '') IS NOT NULL
        AND directory.dealer_code IS NULL

    UNION ALL

    SELECT
        'REGISTRATION' AS source_area,
        COALESCE(NULLIF(TRIM(registration_id), ''), '<MISSING>') AS record_id,
        'INVALID_REGISTRATION_DATE' AS issue_type
    FROM raw_registrations
    WHERE NULLIF(TRIM(registration_date), '') IS NOT NULL
        AND TRY_CAST(registration_date AS DATE) IS NULL
)
SELECT
    source_area,
    record_id,
    issue_type
FROM material_issues

UNION ALL

SELECT
    source_area,
    record_id,
    issue_type
FROM registration_issues
ORDER BY source_area, record_id, issue_type;
