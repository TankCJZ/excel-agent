export type WorkbookHistoryEntry = {
  id: string
  workbookId: string | null
  taskId: string | null
  rootId: string
  parentWorkbookId: string | null
  fileName: string
  displayName: string
  version: number
  isLatest: boolean
  createdAt: string
  description: string
  kind: 'upload' | 'spreadsheet_generation' | 'workbook_mutation' | 'workbook_export'
  rowCount: number | null
  sheetCount: number | null
  calculationPending: boolean
}

export type WorkbookHistoryPage = { entries: WorkbookHistoryEntry[]; nextOffset: number | null }
