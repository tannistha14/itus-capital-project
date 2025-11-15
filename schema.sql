-- schema.sql

-- Define the main financial table structure
CREATE TABLE IF NOT EXISTS daily_metrics (
    accord_code INTEGER NOT NULL,
    company_name TEXT,
    sector TEXT,
    mcap_category TEXT,
    date TEXT NOT NULL,
    mcap REAL,
    -- Add other fields if present in your source data
    PRIMARY KEY (accord_code, date)
);

-- Create the required composite index for fast lookups and series queries
CREATE INDEX IF NOT EXISTS idx_accord_date
ON daily_metrics (accord_code, date);