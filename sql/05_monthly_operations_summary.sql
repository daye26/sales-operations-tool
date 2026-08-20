-- DuckDB-compatible example.
-- Builds a compact monthly KPI report from orders, registrations, and inventory.

WITH order_metrics AS (
    SELECT
        DATE_TRUNC('month', TRY_CAST(order_date AS DATE)) AS report_month,
        COUNT(*) AS open_order_count,
        SUM(CASE WHEN UPPER(TRIM(priority)) = 'HIGH' THEN 1 ELSE 0 END) AS high_priority_order_count
    FROM sales_orders
    WHERE UPPER(TRIM(order_status)) = 'OPEN'
        AND TRY_CAST(order_date AS DATE) IS NOT NULL
    GROUP BY 1
),
registration_metrics AS (
    SELECT
        DATE_TRUNC('month', CAST(ingestion_timestamp AS DATE)) AS report_month,
        COUNT(*) AS registration_count,
        SUM(
            CASE
                WHEN TRY_CAST(registration_date AS DATE) IS NOT NULL THEN 1
                ELSE 0
            END
        ) AS parseable_registration_count
    FROM raw_registrations
    WHERE ingestion_timestamp IS NOT NULL
    GROUP BY 1
),
inventory_metrics AS (
    SELECT
        DATE_TRUNC('month', TRY_CAST(availability_date AS DATE)) AS report_month,
        COUNT(*) AS available_vehicle_count
    FROM vehicle_inventory
    WHERE UPPER(TRIM(vehicle_status)) = 'AVAILABLE'
        AND TRY_CAST(availability_date AS DATE) IS NOT NULL
    GROUP BY 1
),
report_months AS (
    SELECT report_month FROM order_metrics
    UNION
    SELECT report_month FROM registration_metrics
    UNION
    SELECT report_month FROM inventory_metrics
)
SELECT
    months.report_month,
    COALESCE(orders.open_order_count, 0) AS open_order_count,
    COALESCE(orders.high_priority_order_count, 0) AS high_priority_order_count,
    COALESCE(registrations.registration_count, 0) AS registration_count,
    COALESCE(registrations.parseable_registration_count, 0) AS parseable_registration_count,
    COALESCE(inventory.available_vehicle_count, 0) AS available_vehicle_count
FROM report_months AS months
LEFT JOIN order_metrics AS orders
    ON orders.report_month = months.report_month
LEFT JOIN registration_metrics AS registrations
    ON registrations.report_month = months.report_month
LEFT JOIN inventory_metrics AS inventory
    ON inventory.report_month = months.report_month
ORDER BY months.report_month;
