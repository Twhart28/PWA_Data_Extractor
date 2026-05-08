import { useEffect, useMemo, useState } from 'react'
import { bridge } from './backendBridge'
import './styles.css'
import type { AppSnapshot, GroupingMode, ImportState, ReportMode, ReviewPatient } from './types'

const reportModeOptions: Array<{ value: ReportMode; label: string }> = [
  { value: 'detailed', label: 'Detailed reports' },
  { value: 'clinical', label: 'Clinical reports' },
]

const groupingModeOptions: Array<{ value: GroupingMode; label: string }> = [
  { value: 'subject', label: 'Subject only' },
  { value: 'subject_timepoint', label: 'Subject + timepoint' },
  { value: 'subject_visit', label: 'Subject + visit' },
  { value: 'subject_visit_timepoint', label: 'Subject + visit + timepoint' },
]

const reasonLabel: Record<ReviewPatient['reason'], string> = {
  multi_entry: 'Multi-entry review',
  pair_alert: 'Pair alert approval',
  multi_entry_pair_alert: 'Multi-entry + pair alert',
}

function App() {
  const [snapshot, setSnapshot] = useState<AppSnapshot | null>(null)
  const [importState, setImportState] = useState<ImportState | null>(null)
  const [workflowScreen, setWorkflowScreen] = useState<'import' | 'review'>('import')
  const [selectedReviewId, setSelectedReviewId] = useState<string | null>(null)
  const [processing, setProcessing] = useState(false)

  useEffect(() => {
    bridge.getSnapshot().then((next) => {
      setSnapshot(next)
      setImportState(next.importState)
      setSelectedReviewId(next.reviewPatients[0]?.id ?? null)
    })
  }, [])

  const selectedPatient = useMemo(
    () => snapshot?.reviewPatients.find((patient) => patient.id === selectedReviewId) ?? null,
    [selectedReviewId, snapshot],
  )

  if (!snapshot || !importState) {
    return <div className="loading-shell">Loading migration preview…</div>
  }

  const handleProcess = async () => {
    setProcessing(true)
    const next = await bridge.processFiles(importState)
    setSnapshot(next)
    setSelectedReviewId(next.reviewPatients[0]?.id ?? null)
    setWorkflowScreen('review')
    setProcessing(false)
  }

  return (
    <div className="app-shell">
      <header className="app-header">
        <div>
          <p className="eyebrow">Desktop migration preview</p>
          <h1>PWA Data Extractor</h1>
          <p className="lede">
            Proposed Tauri + React workspace for a cleaner two-step flow while keeping the Python
            parsing and export engine.
          </p>
        </div>
        <div className="header-chip-stack">
          <span className="chip chip-muted">Backend preserved in Python</span>
          <span className="chip chip-accent">Frontend rebuilt for workflow clarity</span>
        </div>
      </header>

      <nav className="workflow-switcher" aria-label="Workflow screens">
        <button
          className={workflowScreen === 'import' ? 'active' : ''}
          onClick={() => setWorkflowScreen('import')}
          type="button"
        >
          1. Import / Process
        </button>
        <button
          className={workflowScreen === 'review' ? 'active' : ''}
          onClick={() => setWorkflowScreen('review')}
          type="button"
        >
          2. Review / Export
        </button>
      </nav>

      {workflowScreen === 'import' ? (
        <section className="screen-grid screen-grid-import">
          <aside className="setup-sidebar">
            <div className="hero-card">
              <p className="eyebrow">Step 1</p>
              <h2>Set the extraction context once</h2>
              <p>
                Instead of a narrow sidebar, this screen treats setup as its own workspace. The goal
                is to make report mode, grouping, file selection, and output decisions obvious
                before the user starts processing.
              </p>
            </div>

            <div className="info-card">
              <h3>Run summary</h3>
              <dl className="definition-list">
                <div>
                  <dt>Input mode</dt>
                  <dd>{reportModeOptions.find((option) => option.value === importState.reportMode)?.label}</dd>
                </div>
                <div>
                  <dt>Grouping</dt>
                  <dd>{groupingModeOptions.find((option) => option.value === importState.groupingMode)?.label}</dd>
                </div>
                <div>
                  <dt>Files selected</dt>
                  <dd>{importState.files.length}</dd>
                </div>
                <div>
                  <dt>Workbook</dt>
                  <dd>{importState.outputPath.split('\\').slice(-1)[0]}</dd>
                </div>
              </dl>
            </div>

            <div className="info-card">
              <h3>Grouping guide</h3>
              <ul className="plain-list">
                <li>Subject only: use when all repeated measures belong under one subject number.</li>
                <li>Subject + timepoint: best fit for names like <code>S01_000a</code>.</li>
                <li>Subject + visit: use only when visit is explicitly encoded.</li>
                <li>Subject + visit + timepoint: use when both dimensions are present and matter.</li>
              </ul>
            </div>
          </aside>

          <main className="setup-workspace">
            <div className="workspace-header">
              <div>
                <p className="eyebrow">Import and process</p>
                <h2>Build the dataset for review</h2>
                <p>
                  Choose the report source, confirm filename grouping, then process the PDFs into a
                  review-ready dataset.
                </p>
              </div>
              <button className="primary" disabled={processing} onClick={handleProcess} type="button">
                {processing ? 'Processing…' : 'Process reports'}
              </button>
            </div>

            <div className="workspace-grid">
              <section className="workspace-card workspace-card-wide">
                <header>
                  <h3>Report setup</h3>
                  <p>These settings control how the Python backend interprets the files you select.</p>
                </header>
                <div className="form-grid">
                  <label>
                    <span>Input mode</span>
                    <select
                      value={importState.reportMode}
                      onChange={(event) =>
                        setImportState({ ...importState, reportMode: event.target.value as ReportMode })
                      }
                    >
                      {reportModeOptions.map((option) => (
                        <option key={option.value} value={option.value}>
                          {option.label}
                        </option>
                      ))}
                    </select>
                  </label>
                  <label>
                    <span>Grouping mode</span>
                    <select
                      value={importState.groupingMode}
                      onChange={(event) =>
                        setImportState({ ...importState, groupingMode: event.target.value as GroupingMode })
                      }
                    >
                      {groupingModeOptions.map((option) => (
                        <option key={option.value} value={option.value}>
                          {option.label}
                        </option>
                      ))}
                    </select>
                  </label>
                </div>
              </section>

              <section className="workspace-card workspace-card-wide">
                <header>
                  <h3>Selected files</h3>
                  <p>In the desktop build, this becomes a native file picker plus drag-and-drop target.</p>
                </header>
                <div className="file-well">
                  {importState.files.map((file) => (
                    <div key={file} className="file-pill">
                      {file}
                    </div>
                  ))}
                </div>
                <div className="button-row">
                  <button type="button">Add PDFs</button>
                  <button type="button" className="muted">
                    Remove selected
                  </button>
                  <button type="button" className="muted">
                    Clear
                  </button>
                </div>
              </section>

              <section className="workspace-card">
                <header>
                  <h3>Workbook output</h3>
                  <p>Keep the output decision near processing instead of hiding it in a sidebar.</p>
                </header>
                <label className="stack-field">
                  <span>Excel export path</span>
                  <input
                    type="text"
                    value={importState.outputPath}
                    onChange={(event) =>
                      setImportState({ ...importState, outputPath: event.target.value })
                    }
                  />
                </label>
              </section>

              <section className="workspace-card">
                <header>
                  <h3>Thresholds</h3>
                  <p>These control pair-difference highlighting and alert gating in review.</p>
                </header>
                <div className="threshold-grid">
                  <label>
                    <span>Green up to</span>
                    <input
                      type="number"
                      value={importState.greenMax}
                      onChange={(event) =>
                        setImportState({ ...importState, greenMax: Number(event.target.value) })
                      }
                    />
                  </label>
                  <label>
                    <span>Yellow up to</span>
                    <input
                      type="number"
                      value={importState.yellowMax}
                      onChange={(event) =>
                        setImportState({ ...importState, yellowMax: Number(event.target.value) })
                      }
                    />
                  </label>
                  <label>
                    <span>Pair alert above</span>
                    <input
                      type="number"
                      value={importState.pairAlertThreshold}
                      onChange={(event) =>
                        setImportState({ ...importState, pairAlertThreshold: Number(event.target.value) })
                      }
                    />
                  </label>
                </div>
              </section>
            </div>
          </main>
        </section>
      ) : (
        <section className="screen-grid screen-grid-review">
          <aside className="review-sidebar">
            <div className="review-sidebar-header">
              <div>
                <p className="eyebrow">Review queue</p>
                <h2>Flagged patients</h2>
              </div>
              <span className="count-pill">{snapshot.reviewPatients.length}</span>
            </div>
            <div className="queue-list">
              {snapshot.reviewPatients.map((patient) => (
                <button
                  key={patient.id}
                  type="button"
                  className={`queue-item ${patient.id === selectedReviewId ? 'selected' : ''}`}
                  onClick={() => setSelectedReviewId(patient.id)}
                >
                  <div>
                    <strong>{patient.id}</strong>
                    <span>{reasonLabel[patient.reason]}</span>
                  </div>
                  <small>{patient.approved ? 'Approved' : `${patient.selectedCount}/2 selected`}</small>
                </button>
              ))}
            </div>
          </aside>

          <main className="review-workspace">
            <div className="workspace-header">
              <div>
                <p className="eyebrow">Review and export</p>
                <h2>{selectedPatient?.id ?? 'No patient selected'}</h2>
                <p>
                  Use the full-width review workspace for pair selection, alert approval, and export.
                </p>
              </div>
              <div className="header-actions">
                <button type="button" className="muted">
                  Reset to auto pair
                </button>
                <button type="button" className="muted">
                  View selected PDF
                </button>
                <button type="button" className="primary">
                  Export workbook
                </button>
              </div>
            </div>

            {selectedPatient ? (
              <>
                <div className="review-context-grid">
                  <div className="context-card">
                    <span className="eyebrow">Reason</span>
                    <strong>{reasonLabel[selectedPatient.reason]}</strong>
                  </div>
                  <div className="context-card">
                    <span className="eyebrow">Approval</span>
                    <strong>{selectedPatient.approved ? 'Approved' : 'Needs approval'}</strong>
                  </div>
                  <div className="context-card">
                    <span className="eyebrow">Selected files</span>
                    <strong>{selectedPatient.selectedFiles.join(' | ')}</strong>
                  </div>
                </div>

                <section className="workspace-card workspace-card-wide">
                  <header>
                    <h3>Selected measurements</h3>
                    <p>This table keeps the same core metrics as the current desktop workflow.</p>
                  </header>
                  <div className="table-shell">
                    <table>
                      <thead>
                        <tr>
                          <th>Keep</th>
                          <th>SYS</th>
                          <th>DIA</th>
                          <th>MAP</th>
                          <th>Aortic SYS</th>
                          <th>Aortic DIA</th>
                          <th>Source File</th>
                          <th>Pair method</th>
                        </tr>
                      </thead>
                      <tbody>
                        {selectedPatient.rows.map((row) => (
                          <tr key={row.sourceFile} className={row.keep ? 'kept' : ''}>
                            <td>{row.keep ? '✓' : ''}</td>
                            <td>{row.sys}</td>
                            <td>{row.dia}</td>
                            <td>{row.map}</td>
                            <td>{row.aorticSys}</td>
                            <td>{row.aorticDia}</td>
                            <td className="left">{row.sourceFile}</td>
                            <td>{row.pairMethod}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </section>

                <section className="workspace-card workspace-card-wide">
                  <header>
                    <h3>Selected pair differences</h3>
                    <p>Use this section to justify approval before export.</p>
                  </header>
                  <div className="diff-grid">
                    {Object.entries(selectedPatient.currentPairDiffs).map(([key, value]) => (
                      <div key={key} className={`diff-card diff-${value <= 3 ? 'good' : value <= 6 ? 'warn' : 'bad'}`}>
                        <span>{key}</span>
                        <strong>{value}</strong>
                      </div>
                    ))}
                  </div>
                </section>
              </>
            ) : null}
          </main>
        </section>
      )}
    </div>
  )
}

export default App
