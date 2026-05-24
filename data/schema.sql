-- S&OP Hub — DuckDB schema (SPEC §8.1)
-- Composite PKs on fact tables include data_version so prior uploads
-- remain queryable as 'superseded' rows for audit.

CREATE TABLE IF NOT EXISTS dim_sku (
    sku_code        VARCHAR PRIMARY KEY,
    sku_name        VARCHAR,
    brand           VARCHAR,
    category        VARCHAR,
    subcategory     VARCHAR,
    uom             VARCHAR,
    uom_to_cases    FLOAT,
    is_active       BOOLEAN DEFAULT TRUE,
    loaded_at       TIMESTAMP
);

CREATE TABLE IF NOT EXISTS dim_channel (
    channel_code    VARCHAR PRIMARY KEY,
    channel_name    VARCHAR,
    channel_type    VARCHAR,
    loaded_at       TIMESTAMP
);

CREATE TABLE IF NOT EXISTS dim_region (
    region_code     VARCHAR PRIMARY KEY,
    region_name     VARCHAR,
    country         VARCHAR,
    cluster         VARCHAR,
    loaded_at       TIMESTAMP
);

CREATE TABLE IF NOT EXISTS dim_plant (
    plant_code      VARCHAR PRIMARY KEY,
    plant_name      VARCHAR,
    country         VARCHAR,
    loaded_at       TIMESTAMP
);

CREATE TABLE IF NOT EXISTS fact_demand (
    period_date         DATE,
    sku_code            VARCHAR,
    channel_code        VARCHAR,
    region_code         VARCHAR,
    statistical_fcst    FLOAT,
    consensus_fcst      FLOAT,
    actuals             FLOAT,
    data_version        INTEGER,
    loaded_at           TIMESTAMP,
    PRIMARY KEY (period_date, sku_code, channel_code, region_code, data_version)
);

-- DOS is NOT stored — computed at query time joining to fact_demand.
CREATE TABLE IF NOT EXISTS fact_supply (
    period_date         DATE,
    sku_code            VARCHAR,
    plant_code          VARCHAR,
    inventory_qty       FLOAT,
    production_plan     FLOAT,
    production_actual   FLOAT,
    capacity_plan       FLOAT,
    orders_requested    FLOAT,
    orders_delivered    FLOAT,
    data_version        INTEGER,
    loaded_at           TIMESTAMP,
    PRIMARY KEY (period_date, sku_code, plant_code, data_version)
);

CREATE TABLE IF NOT EXISTS fact_financial (
    period_date         DATE,
    sku_code            VARCHAR,
    channel_code        VARCHAR,
    revenue_actual      FLOAT,
    revenue_budget      FLOAT,
    revenue_le          FLOAT,
    gm_actual           FLOAT,
    gm_budget           FLOAT,
    promo_spend_actual  FLOAT,
    promo_spend_budget  FLOAT,
    currency_code       VARCHAR DEFAULT 'EUR',
    data_version        INTEGER,
    loaded_at           TIMESTAMP,
    PRIMARY KEY (period_date, sku_code, channel_code, data_version)
);

CREATE SEQUENCE IF NOT EXISTS upload_log_seq START 1;

CREATE TABLE IF NOT EXISTS upload_log (
    id              INTEGER PRIMARY KEY DEFAULT nextval('upload_log_seq'),
    filename        VARCHAR,
    domain          VARCHAR,
    uploaded_at     TIMESTAMP,
    row_count       INTEGER,
    data_version    INTEGER,
    status          VARCHAR,
    error_msg       VARCHAR
);
