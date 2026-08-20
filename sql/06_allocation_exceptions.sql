-- DuckDB-compatible example.
-- Identifies cases that need manual review after the same eligibility checks
-- used by the allocation candidate-ranking example.

WITH active_orders AS (
    SELECT
        order_record.order_id,
        order_record.dealer_id,
        UPPER(TRIM(order_record.material_group)) AS material_group,
        UPPER(TRIM(order_record.country_code)) AS country_code,
        TRY_CAST(order_record.order_date AS DATE) AS order_date,
        CASE UPPER(TRIM(order_record.priority))
            WHEN 'HIGH' THEN 1
            WHEN 'MEDIUM' THEN 2
            ELSE 3
        END AS priority_rank
    FROM sales_orders AS order_record
    INNER JOIN dealer_directory AS dealer
        ON dealer.dealer_id = order_record.dealer_id
    WHERE UPPER(TRIM(order_record.order_status)) = 'OPEN'
        AND dealer.is_active
),
available_vehicles AS (
    SELECT
        vehicle_id,
        UPPER(TRIM(material_group)) AS material_group,
        UPPER(TRIM(country_code)) AS country_code,
        TRY_CAST(availability_date AS DATE) AS availability_date
    FROM vehicle_inventory
    WHERE UPPER(TRIM(vehicle_status)) = 'AVAILABLE'
        AND TRY_CAST(availability_date AS DATE) <= CURRENT_DATE + INTERVAL 30 DAY
),
eligible_candidates AS (
    SELECT
        orders.order_id,
        orders.order_date,
        orders.priority_rank,
        vehicles.vehicle_id,
        CASE WHEN orders.country_code = vehicles.country_code THEN 100 ELSE 0 END AS candidate_score
    FROM active_orders AS orders
    INNER JOIN available_vehicles AS vehicles
        ON vehicles.material_group = orders.material_group
),
ranked_candidates AS (
    SELECT
        *,
        ROW_NUMBER() OVER (
            PARTITION BY order_id
            ORDER BY candidate_score DESC, vehicle_id
        ) AS candidate_rank
    FROM eligible_candidates
),
best_candidates AS (
    SELECT
        *
    FROM ranked_candidates
    WHERE candidate_rank = 1
),
vehicle_claims AS (
    SELECT
        *,
        ROW_NUMBER() OVER (
            PARTITION BY vehicle_id
            ORDER BY priority_rank, order_date, order_id
        ) AS vehicle_claim_rank
    FROM best_candidates
),
orders_without_candidates AS (
    SELECT
        orders.order_id AS record_id,
        'ORDER_WITHOUT_ELIGIBLE_VEHICLE' AS issue_type
    FROM active_orders AS orders
    LEFT JOIN eligible_candidates AS candidates
        ON candidates.order_id = orders.order_id
    WHERE candidates.order_id IS NULL
),
orders_with_conflicts AS (
    SELECT
        order_id AS record_id,
        'ORDER_REQUIRES_REVIEW_AFTER_VEHICLE_CONFLICT' AS issue_type
    FROM vehicle_claims
    WHERE vehicle_claim_rank > 1
),
vehicles_outside_window AS (
    SELECT
        vehicle_id AS record_id,
        'VEHICLE_OUTSIDE_AVAILABILITY_WINDOW' AS issue_type
    FROM vehicle_inventory
    WHERE UPPER(TRIM(vehicle_status)) = 'AVAILABLE'
        AND TRY_CAST(availability_date AS DATE) > CURRENT_DATE + INTERVAL 30 DAY
)
SELECT issue_type, record_id FROM orders_without_candidates
UNION ALL
SELECT issue_type, record_id FROM orders_with_conflicts
UNION ALL
SELECT issue_type, record_id FROM vehicles_outside_window
ORDER BY issue_type, record_id;
