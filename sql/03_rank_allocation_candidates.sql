-- DuckDB-compatible example.
-- Expected source tables:
-- sales_orders(order_id, dealer_id, material_group, country_code, order_date, priority, order_status)
-- vehicle_inventory(vehicle_id, material_group, country_code, availability_date, vehicle_status)
-- dealer_directory(dealer_id, dealer_group, is_active)

WITH clean_orders AS (
    SELECT
        order_record.order_id,
        order_record.dealer_id,
        dealer.dealer_group,
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
        AND order_record.order_id IS NOT NULL
        AND order_record.material_group IS NOT NULL
),
available_vehicles AS (
    SELECT
        vehicle_id,
        UPPER(TRIM(material_group)) AS material_group,
        UPPER(TRIM(country_code)) AS country_code,
        TRY_CAST(availability_date AS DATE) AS availability_date
    FROM vehicle_inventory
    WHERE UPPER(TRIM(vehicle_status)) = 'AVAILABLE'
        AND vehicle_id IS NOT NULL
        AND material_group IS NOT NULL
),
eligible_candidates AS (
    SELECT
        orders.order_id,
        orders.dealer_id,
        orders.dealer_group,
        orders.order_date,
        orders.priority_rank,
        vehicles.vehicle_id,
        vehicles.availability_date,
        CASE
            WHEN orders.country_code = vehicles.country_code THEN 100
            WHEN orders.country_code IS NULL OR vehicles.country_code IS NULL THEN 50
            ELSE 0
        END
        + CASE WHEN orders.dealer_group = 'PRIORITY' THEN 20 ELSE 0 END
        AS candidate_score
    FROM clean_orders AS orders
    INNER JOIN available_vehicles AS vehicles
        ON vehicles.material_group = orders.material_group
    WHERE vehicles.availability_date <= CURRENT_DATE + INTERVAL 30 DAY
),
ranked_per_order AS (
    SELECT
        *,
        ROW_NUMBER() OVER (
            PARTITION BY order_id
            ORDER BY candidate_score DESC, availability_date, vehicle_id
        ) AS candidate_rank
    FROM eligible_candidates
),
best_candidate_per_order AS (
    SELECT
        *
    FROM ranked_per_order
    WHERE candidate_rank = 1
),
resolved_vehicle_conflicts AS (
    SELECT
        *,
        ROW_NUMBER() OVER (
            PARTITION BY vehicle_id
            ORDER BY priority_rank, order_date, order_id
        ) AS vehicle_claim_rank
    FROM best_candidate_per_order
)
SELECT
    order_id,
    dealer_id,
    vehicle_id,
    availability_date,
    candidate_score,
    priority_rank
FROM resolved_vehicle_conflicts
WHERE vehicle_claim_rank = 1
ORDER BY priority_rank, order_date, order_id;
