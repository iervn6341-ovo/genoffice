/** Orders Save and workbook Undo/Redo clicks without blocking unrelated UI work. */
export class HistorySaveBarrier {
  private saving: Promise<unknown> | null = null
  private history: Promise<void> = Promise.resolve()

  runHistory(action: () => Promise<unknown>, waitForCarry: () => Promise<boolean>): Promise<void> {
    const precedingSave = this.saving
    const operation = this.history.then(async () => {
      await precedingSave?.catch(() => undefined)
      if (!(await waitForCarry())) throw new Error('Workbook history is still loading.')
      await action()
    })
    this.history = operation.catch(() => undefined)
    return operation
  }

  runSave<T>(action: () => Promise<T>, waitForCarry?: () => Promise<boolean>): Promise<T> {
    const precedingSave = this.saving
    const precedingHistory = this.history
    const operation = (async () => {
      await precedingSave?.catch(() => undefined)
      await precedingHistory
      if (waitForCarry && !(await waitForCarry())) {
        throw new Error('Workbook history is still loading.')
      }
      return action()
    })()
    this.saving = operation
    void operation
      .finally(() => {
        if (this.saving === operation) this.saving = null
      })
      .catch(() => undefined)
    return operation
  }
}
