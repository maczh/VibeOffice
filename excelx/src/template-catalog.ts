export interface ExcelxWorkbookTemplate {
  readonly id: string;
  readonly name: string;
}

/** Workbook creation choices rendered by the Excelx React task pane. */
export const EXCELX_WORKBOOK_TEMPLATES: readonly ExcelxWorkbookTemplate[] = [
  { id: 'loan', name: 'Loan Amortization' },
  { id: 'expense', name: 'Expense Statement' },
  { id: 'invoice', name: 'Sales Invoice' },
  { id: 'timecard', name: 'Timecard' },
  { id: 'budget', name: 'Personal Monthly Budget' },
  { id: 'tasks', name: 'Task Tracker' },
  { id: 'shopping', name: 'Shopping List' },
  { id: 'schedule', name: 'Weekly Schedule' },
];
