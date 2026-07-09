-- Engine 4: Budget Line Item Store
-- Persists budget values extracted from P.4 so future generations never require
-- a historical P.4 source tab.  One row per label per location.
-- Applied via: Supabase Management API → POST /v1/projects/{ref}/database/query

CREATE TABLE IF NOT EXISTS budget_line_items (
  id                   BIGSERIAL PRIMARY KEY,

  -- Identity
  label                TEXT        NOT NULL,           -- original label (with leading spaces)
  normalized_label     TEXT        NOT NULL,           -- trimmed + synonym-normalized
  location             TEXT        NOT NULL DEFAULT '', -- matches kpi_snapshots.location

  -- Budget values
  budget_value         NUMERIC(14,2) NOT NULL DEFAULT 0,
  is_section_header    BOOLEAN     NOT NULL DEFAULT FALSE,
  is_total             BOOLEAN     NOT NULL DEFAULT FALSE,
  sort_order           INTEGER     NOT NULL DEFAULT 0, -- row position in source P.4

  -- Housekeeping
  updated_at           TIMESTAMPTZ NOT NULL DEFAULT NOW(),

  -- One row per label per location
  CONSTRAINT uq_budget_label_location UNIQUE (normalized_label, location)
);

CREATE INDEX IF NOT EXISTS idx_budget_line_items_location_order
  ON budget_line_items (location, sort_order ASC);

ALTER TABLE budget_line_items ENABLE ROW LEVEL SECURITY;

CREATE POLICY "service_role_all" ON budget_line_items
  USING (true)
  WITH CHECK (true);
