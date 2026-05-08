import { mockSnapshot } from './mockData'
import type { AppSnapshot, BackendBridge, ImportState } from './types'

const delay = (ms: number) => new Promise((resolve) => setTimeout(resolve, ms))

const cloneSnapshot = (): AppSnapshot =>
  JSON.parse(JSON.stringify(mockSnapshot)) as AppSnapshot

export const bridge: BackendBridge = {
  async getSnapshot() {
    await delay(120)
    return cloneSnapshot()
  },
  async processFiles(input: ImportState) {
    await delay(700)
    const snapshot = cloneSnapshot()
    snapshot.importState = {
      ...snapshot.importState,
      ...input,
    }
    snapshot.summary.processedFiles = Math.max(input.files.length, 0)
    return snapshot
  },
}
