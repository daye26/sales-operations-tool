-- DuckDB-compatible example.
-- Derived and anonymized from the leads aggregation workflow in leads.sql.
-- Expected source tables:
-- lead_events(lead_id, created_date, request_type, model_key)
-- model_equivalence(model_key, reporting_model)

WITH normalized_leads AS (
    SELECT
        TRY_CAST(created_date AS DATE) AS created_date,
        UPPER(COALESCE(NULLIF(TRIM(request_type), ''), 'DEALER LEADS')) AS request_type,
        UPPER(NULLIF(TRIM(model_key), '')) AS model_key
    FROM lead_events
),
daily_metrics AS (
    SELECT
        model_map.reporting_model,
        leads.created_date,
        NULLIF(COUNT(CASE WHEN leads.request_type = 'OFFER REQUEST' THEN 1 END), 0) AS offer_request,
        NULLIF(COUNT(CASE WHEN leads.request_type = 'TEST DRIVE REQUEST' THEN 1 END), 0) AS test_drive_request,
        NULLIF(COUNT(CASE WHEN leads.request_type = 'DEALER LEADS' THEN 1 END), 0) AS dealer_leads,
        NULLIF(
            ROUND(
                COUNT(CASE WHEN leads.request_type = 'TEST DRIVE REQUEST' THEN 1 END)
                + COUNT(CASE WHEN leads.request_type = 'DEALER LEADS' THEN 1 END) * 0.3
            ),
            0
        ) AS test_drive
    FROM normalized_leads AS leads
    LEFT JOIN model_equivalence AS model_map
        ON model_map.model_key = leads.model_key
    WHERE leads.created_date IS NOT NULL
        AND leads.model_key IS NOT NULL
    GROUP BY model_map.reporting_model, leads.created_date
),
long_metrics AS (
    SELECT reporting_model, created_date, 'offer' AS metric, offer_request AS value FROM daily_metrics
    UNION ALL
    SELECT reporting_model, created_date, 'test drive request', test_drive_request FROM daily_metrics
    UNION ALL
    SELECT reporting_model, created_date, 'test drive', test_drive FROM daily_metrics
    UNION ALL
    SELECT reporting_model, created_date, 'dealer leads', dealer_leads FROM daily_metrics
)
PIVOT long_metrics
ON created_date
USING MAX(value)
GROUP BY reporting_model, metric
ORDER BY
    reporting_model,
    CASE metric
        WHEN 'offer' THEN 1
        WHEN 'test drive request' THEN 2
        WHEN 'test drive' THEN 3
        WHEN 'dealer leads' THEN 4
        ELSE 99
    END;
