-- Dashboard download history
-- One row per generated Excel dashboard, mirroring the `reports` table for PDFs.

CREATE TABLE IF NOT EXISTS excel_downloads (
  id          BIGSERIAL    PRIMARY KEY,
  filename    TEXT         NOT NULL,
  report_date DATE         NOT NULL,
  excel_path  TEXT         NOT NULL,
  created_at  TIMESTAMPTZ  NOT NULL DEFAULT NOW()
);

CREATE INDEX IF NOT EXISTS idx_excel_downloads_created_at
  ON excel_downloads (created_at DESC);

ALTER TABLE excel_downloads ENABLE ROW LEVEL SECURITY;

DO $$ BEGIN
  IF NOT EXISTS (
    SELECT 1 FROM pg_policies
    WHERE tablename = 'excel_downloads' AND policyname = 'service_role_all'
  ) THEN
    CREATE POLICY "service_role_all" ON excel_downloads
      USING (true)
      WITH CHECK (true);
  END IF;
END $$;
