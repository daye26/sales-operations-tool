-- DuckDB-compatible fictional data used to run the portfolio examples.
-- This script only creates and resets demonstration tables.

DROP TABLE IF EXISTS raw_material_codes;
DROP TABLE IF EXISTS raw_registrations;
DROP TABLE IF EXISTS vehicle_catalog;
DROP TABLE IF EXISTS dealer_directory;
DROP TABLE IF EXISTS sales_orders;
DROP TABLE IF EXISTS vehicle_inventory;
DROP TABLE IF EXISTS material_change_requests;
DROP TABLE IF EXISTS available_material_options;

CREATE TABLE raw_material_codes (
    material_code VARCHAR,
    product_series VARCHAR,
    trim_level VARCHAR,
    interior_color VARCHAR,
    exterior_color VARCHAR
);

INSERT INTO raw_material_codes VALUES
    (' mc-001 ', 'model a', 'standard', 'black', 'grey'),
    ('mc-002', 'model b', 'premium', 'beige', 'pearl white'),
    ('', 'model c', 'standard', 'dark', 'silver'),
    ('x1', 'model d', NULL, 'light', 'white');

CREATE TABLE raw_registrations (
    registration_id VARCHAR,
    vehicle_id VARCHAR,
    dealer_code VARCHAR,
    registration_date VARCHAR,
    ingestion_timestamp TIMESTAMP
);

CREATE TABLE vehicle_catalog (
    vehicle_id VARCHAR
);

CREATE TABLE dealer_directory (
    dealer_code VARCHAR,
    dealer_id VARCHAR,
    dealer_group VARCHAR,
    is_active BOOLEAN
);

INSERT INTO raw_registrations VALUES
    ('REG-001', 'VEH-001', 'DLR-001', '2026-01-10', '2026-01-10 08:00:00'),
    ('REG-002', 'VEH-001', 'DLR-001', '2026-01-10', '2026-01-11 08:00:00'),
    ('REG-003', 'VEH-999', 'DLR-001', '2026-01-12', '2026-01-12 08:00:00'),
    ('REG-004', 'VEH-002', 'DLR-999', 'invalid date', '2026-01-13 08:00:00'),
    ('REG-005', NULL, 'DLR-001', '2026-01-14', '2026-01-14 08:00:00');

INSERT INTO vehicle_catalog VALUES
    ('VEH-001'),
    ('VEH-002');

INSERT INTO dealer_directory VALUES
    ('DLR-001', 'DLR-001', 'PRIORITY', TRUE),
    ('DLR-002', 'DLR-002', 'STANDARD', TRUE),
    ('DLR-003', 'DLR-003', 'STANDARD', FALSE);

CREATE TABLE sales_orders (
    order_id VARCHAR,
    dealer_id VARCHAR,
    material_group VARCHAR,
    country_code VARCHAR,
    order_date VARCHAR,
    priority VARCHAR,
    order_status VARCHAR
);

CREATE TABLE vehicle_inventory (
    vehicle_id VARCHAR,
    material_group VARCHAR,
    country_code VARCHAR,
    availability_date VARCHAR,
    vehicle_status VARCHAR
);

INSERT INTO sales_orders VALUES
    ('ORD-001', 'DLR-001', 'GROUP-A', 'AA', '2026-01-01', 'HIGH', 'OPEN'),
    ('ORD-002', 'DLR-002', 'GROUP-A', 'BB', '2026-01-02', 'MEDIUM', 'OPEN'),
    ('ORD-003', 'DLR-001', 'GROUP-B', 'AA', '2026-01-03', 'LOW', 'CLOSED'),
    ('ORD-004', 'DLR-999', 'GROUP-A', 'AA', '2026-01-04', 'HIGH', 'OPEN'),
    ('ORD-005', 'DLR-002', 'GROUP-A', 'AA', '2026-01-05', 'MEDIUM', 'OPEN'),
    ('ORD-006', 'DLR-002', 'GROUP-C', 'AA', '2026-01-06', 'LOW', 'OPEN');

INSERT INTO vehicle_inventory VALUES
    ('VEH-101', 'GROUP-A', 'AA', CAST(CURRENT_DATE AS VARCHAR), 'AVAILABLE'),
    ('VEH-102', 'GROUP-A', 'BB', CAST(CURRENT_DATE AS VARCHAR), 'AVAILABLE'),
    ('VEH-103', 'GROUP-A', 'AA', CAST(CURRENT_DATE + INTERVAL 45 DAY AS VARCHAR), 'AVAILABLE'),
    ('VEH-104', 'GROUP-B', 'AA', CAST(CURRENT_DATE AS VARCHAR), 'RESERVED');

CREATE TABLE material_change_requests (
    request_id VARCHAR,
    requested_material_code VARCHAR,
    product_family VARCHAR,
    trim_level VARCHAR,
    interior_color VARCHAR,
    exterior_color VARCHAR,
    model_year VARCHAR,
    request_date VARCHAR
);

CREATE TABLE available_material_options (
    vehicle_id VARCHAR,
    material_code VARCHAR,
    product_family VARCHAR,
    trim_level VARCHAR,
    interior_color VARCHAR,
    exterior_color VARCHAR,
    model_year VARCHAR,
    availability_date VARCHAR,
    vehicle_status VARCHAR
);

INSERT INTO material_change_requests VALUES
    ('REQ-001', 'MC-A1', 'FAMILY-A', 'STANDARD', 'DARK', 'SILVER', 'MY25', '2026-01-10'),
    ('REQ-002', 'MC-B1', 'FAMILY-B', 'PREMIUM', 'LIGHT', 'WHITE', 'MY25', '2026-01-11'),
    ('REQ-003', 'MC-C1', 'FAMILY-C', 'STANDARD', 'DARK', 'WHITE', 'MY25', '2026-01-12');

INSERT INTO available_material_options VALUES
    ('VEH-201', 'MC-A2', 'FAMILY-A', 'STANDARD', 'DARK', 'SILVER', 'MY25', CAST(CURRENT_DATE AS VARCHAR), 'AVAILABLE'),
    ('VEH-202', 'MC-A3', 'FAMILY-A', 'STANDARD', 'DARK', 'WHITE', 'MY25', CAST(CURRENT_DATE + INTERVAL 1 DAY AS VARCHAR), 'AVAILABLE'),
    ('VEH-203', 'MC-A2', 'FAMILY-A', 'STANDARD', 'DARK', 'SILVER', 'MY24', CAST(CURRENT_DATE + INTERVAL 2 DAY AS VARCHAR), 'AVAILABLE'),
    ('VEH-204', 'MC-B2', 'FAMILY-B', 'PREMIUM', 'LIGHT', 'WHITE', 'MY26', CAST(CURRENT_DATE AS VARCHAR), 'AVAILABLE'),
    ('VEH-205', 'MC-C2', 'FAMILY-C', 'STANDARD', 'LIGHT', 'WHITE', 'MY25', CAST(CURRENT_DATE AS VARCHAR), 'AVAILABLE'),
    ('VEH-206', 'MC-B1', 'FAMILY-B', 'PREMIUM', 'LIGHT', 'WHITE', 'MY25', CAST(CURRENT_DATE AS VARCHAR), 'AVAILABLE');
