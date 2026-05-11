export interface SpaceTypeStats {
  label: string;
  totalSpaces: number;
  occupiedSpaces: number;
  occupancyPct: number;
  totalSqFt: number;
  occupiedSqFt: number;
  sqFtPct: number;
}

export interface WeeklySnapshot {
  date: string;
  occupancyPct: number;
  sqFtPct: number;
  contracts: number;
  flexMembers: number;
  mailMembers: number;
  meetingHours: number;
  meetingRevenue: number;
}

export interface DashboardData {
  reportDate: string;
  location: string;
  currentMonthRevenue: number;
  spaceTypes: SpaceTypeStats[];
  totals: SpaceTypeStats;
  weeklySnapshots: WeeklySnapshot[];
}

export interface RevenueLineItem {
  label: string;
  isTotal: boolean;
  isSectionHeader: boolean;
  values: (number | null)[];
}

export interface RevenueData {
  months: string[];
  lineItems: RevenueLineItem[];
}

export interface OccupancyRow {
  spaceId: string;
  spaceName: string;
  spaceType: string;
  status: "Occupied" | "Vacant";
  capacity: number;
  sqFt: number;
  contractName: string;
  companyName: string;
  moveIn: string;
  moveOut: string;
  term: string;
}

export interface OccupancyData {
  rows: OccupancyRow[];
  totalSpaces: number;
  occupiedCount: number;
  vacantCount: number;
}

export interface BudgetLineItem {
  label: string;
  isSectionHeader: boolean;
  isTotal: boolean;
  actual: number | null;
  budget: number | null;
  varianceDollar: number | null;
  variancePct: number | null;
}

export interface BudgetData {
  lineItems: BudgetLineItem[];
  totalRevenue: number;
  totalCogs: number;
  grossProfit: number;
}

export interface MeetingRoomTransaction {
  company: string;
  description: string;
  chargeCode: string;
  month: string;
  amount: number;
  unit: "$" | "hrs";   // detected from Excel cell format: "$420" → "$", "48" → "hrs"
}

export interface MeetingRoomTypeRow {
  label: string;
  actual: number;
  budget: number;
  varianceDollar: number;
  variancePct: number | null;
  isTotal: boolean;
}

export interface MeetingMonthly {
  month: string;
  revenue: number;
}

export interface MeetingData {
  mtdRevenue: number;
  mtdBudget: number;
  varianceDollar: number;
  variancePct: number | null;
  recurringTransactions: MeetingRoomTransaction[];
  roomTypes: MeetingRoomTypeRow[];
  monthlyTrend: MeetingMonthly[];
}

export interface ValidationCheck {
  label: string;
  passed: boolean;
  expected: string;
  actual: string;
  note: string;
}

export interface ValidationResult {
  checks: ValidationCheck[];
  passCount: number;
  failCount: number;
}

export interface ReportData {
  reportDate: string;
  dashboard: DashboardData;
  revenue: RevenueData;
  occupancy: OccupancyData;
  budget: BudgetData;
  meeting: MeetingData;
  validation: ValidationResult;
}
