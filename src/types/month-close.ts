export type CheckStatus =
  | "confirmed"
  | "variance"
  | "missing"
  | "over_limit"
  | "cannot_compute"
  | "review"
  | "na";

export interface MonthCloseItem {
  id: string;
  description: string;
  account_codes: string[];
  check_type: string;
  status: CheckStatus;
  actual: number | null;
  expected: number | null;
  note: string;
  requires_external: boolean;
}

export interface MonthCloseSection {
  id: string;
  title: string;
  note?: string;
  items: MonthCloseItem[];
}

export interface MonthCloseCounts {
  confirmed: number;
  attention: number;
  review: number;
  na: number;
  total: number;
}

export interface MonthCloseData {
  location: string;
  property_code: string;
  month: string;
  budget_aligned: boolean;
  budget_mtd_month: string;
  generated_at: string;
  counts: MonthCloseCounts;
  sections: MonthCloseSection[];
  attention_items: MonthCloseItem[];
  external_items: MonthCloseItem[];
}
