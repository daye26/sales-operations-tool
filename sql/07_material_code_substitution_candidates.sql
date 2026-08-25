-- DuckDB-compatible example.
-- Recommends a compatible alternative material code for each request.
-- This is a read-only diagnostic query: it never updates an order or inventory record.

WITH clean_requests AS (
    SELECT
        request_id,
        UPPER(TRIM(requested_material_code)) AS requested_material_code,
        UPPER(TRIM(product_family)) AS product_family,
        UPPER(TRIM(trim_level)) AS trim_level,
        UPPER(TRIM(interior_color)) AS interior_color,
        UPPER(TRIM(exterior_color)) AS exterior_color,
        UPPER(TRIM(model_year)) AS model_year,
        TRY_CAST(request_date AS DATE) AS request_date
    FROM material_change_requests
    WHERE request_id IS NOT NULL
        AND requested_material_code IS NOT NULL
),
available_options AS (
    SELECT
        vehicle_id,
        UPPER(TRIM(material_code)) AS material_code,
        UPPER(TRIM(product_family)) AS product_family,
        UPPER(TRIM(trim_level)) AS trim_level,
        UPPER(TRIM(interior_color)) AS interior_color,
        UPPER(TRIM(exterior_color)) AS exterior_color,
        UPPER(TRIM(model_year)) AS model_year,
        TRY_CAST(availability_date AS DATE) AS availability_date,
        COUNT(*) OVER (PARTITION BY UPPER(TRIM(material_code))) AS available_vehicle_count
    FROM available_material_options
    WHERE UPPER(TRIM(vehicle_status)) = 'AVAILABLE'
        AND vehicle_id IS NOT NULL
        AND material_code IS NOT NULL
),
compatible_candidates AS (
    SELECT
        requests.request_id,
        requests.requested_material_code,
        options.material_code AS proposed_material_code,
        options.vehicle_id,
        options.availability_date,
        options.available_vehicle_count,
        CASE WHEN requests.interior_color = options.interior_color THEN 1 ELSE 0 END AS interior_color_match,
        CASE WHEN requests.exterior_color = options.exterior_color THEN 1 ELSE 0 END AS exterior_color_match,
        CASE WHEN requests.model_year = options.model_year THEN 1 ELSE 0 END AS model_year_match,
        100
            + CASE WHEN requests.interior_color = options.interior_color THEN 20 ELSE 0 END
            + CASE WHEN requests.exterior_color = options.exterior_color THEN 20 ELSE 0 END
            + CASE WHEN requests.model_year = options.model_year THEN 10 ELSE 0 END
            + LEAST(options.available_vehicle_count, 10) AS compatibility_score
    FROM clean_requests AS requests
    INNER JOIN available_options AS options
        ON options.product_family = requests.product_family
        AND options.trim_level = requests.trim_level
    WHERE options.material_code <> requests.requested_material_code
        AND options.availability_date <= CURRENT_DATE + INTERVAL 30 DAY
),
ranked_candidates AS (
    SELECT
        *,
        ROW_NUMBER() OVER (
            PARTITION BY request_id
            ORDER BY compatibility_score DESC, availability_date, proposed_material_code, vehicle_id
        ) AS candidate_rank
    FROM compatible_candidates
)
SELECT
    request_id,
    requested_material_code,
    proposed_material_code,
    vehicle_id,
    availability_date,
    available_vehicle_count,
    interior_color_match,
    exterior_color_match,
    model_year_match,
    compatibility_score
FROM ranked_candidates
WHERE candidate_rank = 1
ORDER BY request_id;
