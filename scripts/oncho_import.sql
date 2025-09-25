-- 1. Copy CSV to the server
--    scp -C oncho_dataset.csv ocs@107.20.181.165:~
-- 2. Copy CSV into the postgres
--    docker cp oncho_dataset.csv espen-sql-api-postgres:/tmp/
-- 3. Run the script below to re-create the table and import the data
-- Create the table schema
DROP TABLE IF EXISTS oncho_data;
CREATE TABLE oncho_data (
    iu_code_name VARCHAR(50) NOT NULL,
    scenario VARCHAR(50) NOT NULL,
    scenario_label VARCHAR(100) NOT NULL,
    year INTEGER NOT NULL,
    percentile_50 DECIMAL(10,8),
    admin0iso3 VARCHAR(3) NOT NULL,
    iu_name VARCHAR(100) NOT NULL,
    iu_id INTEGER NOT NULL,
    country_name VARCHAR(100) NOT NULL,
    admin1 VARCHAR(100),
    pop_tot INTEGER,
    mda INTEGER,
    drug VARCHAR(128),
    coverage_of_eligibles DECIMAL(3,2),
    fitz_cost DECIMAL(12,2)
);

-- Add indexes for better query performance
CREATE INDEX idx_oncho_data_iu_id ON oncho_data(iu_id);
CREATE INDEX idx_oncho_data_year ON oncho_data(year);
CREATE INDEX idx_oncho_data_country ON oncho_data(country_name);
CREATE INDEX idx_oncho_data_scenario ON oncho_data(scenario_label);

-- Import the CSV data
COPY oncho_data(
    iu_code_name, scenario, scenario_label, year, percentile_50,
    admin0iso3, iu_name, iu_id, country_name, admin1, pop_tot,
    mda, drug, coverage_of_eligibles, fitz_cost
)
FROM '/tmp/oncho_dataset.csv'
WITH (
    FORMAT CSV,
    HEADER true,
    DELIMITER ','
);
