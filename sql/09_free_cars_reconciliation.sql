-- DuckDB-compatible example.
-- Derived and anonymized from the consistency checks in not_allocated.sql.
-- This is a read-only diagnostic query: it does not update or remove free-car records.
-- Expected source tables:
-- free_car_records(vehicle_id, note, deleted_at)
-- vehicle_tracking_status(vehicle_id, vehicle_status, delivery_destination,
--                         reserved_order_id, related_order_id, tag, vessel_name)
-- vehicle_reservations(vehicle_id)

WITH active_free_cars AS (
    SELECT
        UPPER(TRIM(vehicle_id)) AS vehicle_id,
        NULLIF(TRIM(note), '') AS note
    FROM free_car_records
    WHERE NULLIF(TRIM(vehicle_id), '') IS NOT NULL
        AND deleted_at IS NULL
),
assigned_or_linked AS (
    SELECT
        free_car.vehicle_id,
        'ASSIGNED_OR_LINKED_IN_TRACKING' AS issue_type
    FROM active_free_cars AS free_car
    INNER JOIN vehicle_tracking_status AS tracking
        ON tracking.vehicle_id = free_car.vehicle_id
    WHERE NULLIF(TRIM(tracking.delivery_destination), '') IS NOT NULL
        OR NULLIF(TRIM(tracking.reserved_order_id), '') IS NOT NULL
        OR NULLIF(TRIM(tracking.related_order_id), '') IS NOT NULL
),
note_tag_mismatches AS (
    SELECT
        free_car.vehicle_id,
        'FREE_CAR_NOTE_TAG_MISMATCH' AS issue_type
    FROM active_free_cars AS free_car
    INNER JOIN vehicle_tracking_status AS tracking
        ON tracking.vehicle_id = free_car.vehicle_id
    WHERE UPPER(COALESCE(TRIM(tracking.vehicle_status), '')) <> 'OFFLINE'
        AND (
            (free_car.note IS NULL AND NULLIF(TRIM(tracking.tag), '') IS NOT NULL)
            OR (free_car.note IS NOT NULL AND NULLIF(TRIM(tracking.tag), '') IS NULL)
        )
),
reserved_free_cars AS (
    SELECT
        free_car.vehicle_id,
        'ALSO_IN_RESERVATIONS' AS issue_type
    FROM active_free_cars AS free_car
    INNER JOIN vehicle_reservations AS reservation
        ON reservation.vehicle_id = free_car.vehicle_id
),
offline_or_unshipped AS (
    SELECT
        free_car.vehicle_id,
        'OFFLINE_OR_MISSING_VESSEL' AS issue_type
    FROM active_free_cars AS free_car
    INNER JOIN vehicle_tracking_status AS tracking
        ON tracking.vehicle_id = free_car.vehicle_id
    WHERE UPPER(COALESCE(TRIM(tracking.vehicle_status), '')) = 'OFFLINE'
        OR NULLIF(TRIM(tracking.vessel_name), '') IS NULL
),
missing_from_tracking AS (
    SELECT
        free_car.vehicle_id,
        'MISSING_FROM_VEHICLE_TRACKING' AS issue_type
    FROM active_free_cars AS free_car
    WHERE NOT EXISTS (
        SELECT 1
        FROM vehicle_tracking_status AS tracking
        WHERE tracking.vehicle_id = free_car.vehicle_id
    )
)
SELECT vehicle_id, issue_type FROM assigned_or_linked
UNION ALL
SELECT vehicle_id, issue_type FROM note_tag_mismatches
UNION ALL
SELECT vehicle_id, issue_type FROM reserved_free_cars
UNION ALL
SELECT vehicle_id, issue_type FROM offline_or_unshipped
UNION ALL
SELECT vehicle_id, issue_type FROM missing_from_tracking
ORDER BY vehicle_id, issue_type;
