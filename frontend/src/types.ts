export type ReportMode = 'detailed' | 'clinical'

export type GroupingMode =
  | 'subject'
  | 'subject_timepoint'
  | 'subject_visit'
  | 'subject_visit_timepoint'

export type ReviewReason = 'multi_entry' | 'pair_alert' | 'multi_entry_pair_alert'

export interface ImportState {
  reportMode: ReportMode
  groupingMode: GroupingMode
  files: string[]
  outputPath: string
  greenMax: number
  yellowMax: number
  pairAlertThreshold: number
}

export interface ReviewPatient {
  id: string
  reason: ReviewReason
  selectedCount: number
  approved: boolean
  selectedFiles: string[]
  currentPairDiffs: {
    sys: number
    dia: number
    map: number
    aorticSys: number
    aorticDia: number
  }
  rows: ReviewRow[]
}

export interface ReviewRow {
  keep: boolean
  sys: number
  dia: number
  map: number
  aorticSys: number
  aorticDia: number
  sourceFile: string
  pairMethod: 'Auto' | 'Manual'
}

export interface AppSnapshot {
  importState: ImportState
  reviewPatients: ReviewPatient[]
  summary: {
    processedFiles: number
    reviewPatients: number
    averagedPatients: number
    specialRows: number
  }
}

export interface BackendBridge {
  getSnapshot(): Promise<AppSnapshot>
  processFiles(input: ImportState): Promise<AppSnapshot>
}
