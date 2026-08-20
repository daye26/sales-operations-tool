-- DuckDB-compatible example.
-- Expected source tables:
-- raw_registrations(registration_id, vehicle_id, dealer_code, registration_date, ingestion_timestamp)
-- vehicle_catalog(vehicle_id)
-- dealer_directory(dealer_code, is_active)

WITH parsed_registrations AS (
    SELECT
        NULLIF(TRIM(registration_id), '') AS registration_id,
        NULLIF(UPPER(TRIM(vehicle_id)), '') AS vehicle_id,
        NULLIF(UPPER(TRIM(dealer_code)), '') AS dealer_code,
        TRY_CAST(registration_date AS DATE) AS registration_date,
        ingestion_timestamp
    FROM raw_registrations
),
deduplicated_registrations AS (
    SELECT
        *,
        ROW_NUMBER() OVER (
            PARTITION BY COALESCE(vehicle_id, registration_id)
            ORDER BY ingestion_timestamp DESC NULLS LAST, registration_id DESC
        ) AS recency_rank
    FROM parsed_registrations
),
validated_registrations AS (
    SELECT
        registration_id,
        registration.vehicle_id,
        registration.dealer_code,
        registration.registration_date,
        CASE
            WHEN registration.vehicle_id IS NULL THEN 'MISSING_VEHICLE_ID'
            WHEN catalog.vehicle_id IS NULL THEN 'UNKNOWN_VEHICLE'
            WHEN registration.dealer_code IS NULL THEN 'MISSING_DEALER_CODE'
            WHEN directory.dealer_code IS NULL THEN 'UNKNOWN_DEALER'
            WHEN COALESCE(directory.is_active, FALSE) = FALSE THEN 'INACTIVE_DEALER'
            WHEN registration.registration_date IS NULL THEN 'INVALID_REGISTRATION_DATE'
            ELSE 'VALID'
        END AS validation_status
    FROM deduplicated_registrations AS registration
    LEFT JOIN vehicle_catalog AS catalog
        ON catalog.vehicle_id = registration.vehicle_id
    LEFT JOIN dealer_directory AS directory
        ON directory.dealer_code = registration.dealer_code
    WHERE registration.recency_rank = 1
)
SELECT
    registration_id,
    vehicle_id,
    dealer_code,
    registration_date,
    validation_status
FROM validated_registrations
ORDER BY validation_status, registration_date, vehicle_id;
