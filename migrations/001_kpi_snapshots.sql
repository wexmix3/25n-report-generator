-- Engine 3: Historical KPI Snapshot Store
-- Eliminates the P.1 template dependency by persisting one row per month per location.
-- Applied via: Supabase Management API → POST /v1/projects/{ref}/database/query

CREATE TABLE IF NOT EXISTS kpi_snapshots (
  id                    BIGSERIAL PRIMARY KEY,

  -- Identity
  month_code            TEXT        NOT NULL,           -- "Aug-24"
  month_serial          INTEGER     NOT NULL,           -- Excel date serial (1st of month)
  location              TEXT        NOT NULL DEFAULT '', -- "Waco", "Naperville", etc.

  -- Occupancy KPIs (fractions 0.0–1.0)
  occupancy_pct_spaces  NUMERIC(6,4) NOT NULL DEFAULT 0,
  occupancy_pct_sqft    NUMERIC(6,4) NOT NULL DEFAULT 0,

  -- Membership counts
  contracts             INTEGER     NOT NULL DEFAULT 0,
  flex_members          INTEGER     NOT NULL DEFAULT 0,
  mail_members          INTEGER     NOT NULL DEFAULT 0,

  -- Meeting room metrics
  meeting_hours         NUMERIC(10,2) NOT NULL DEFAULT 0,
  meeting_revenue_dollar NUMERIC(12,2) NOT NULL DEFAULT 0,

  -- Housekeeping
  updated_at            TIMESTAMPTZ NOT NULL DEFAULT NOW(),

  -- One row per month per location; upsert uses this constraint
  CONSTRAINT uq_snapshot_month_location UNIQUE (month_code, location)
);

-- Index for ordered retrieval in the API route
CREATE INDEX IF NOT EXISTS idx_kpi_snapshots_month_serial
  ON kpi_snapshots (month_serial ASC);

-- Enable RLS (service role key bypasses it; auth is handled at the API layer)
ALTER TABLE kpi_snapshots ENABLE ROW LEVEL SECURITY;

-- Service role has full access; no public read needed
CREATE POLICY "service_role_all" ON kpi_snapshots
  USING (true)
  WITH CHECK (true);
