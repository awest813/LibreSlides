/**
 * Undo/redo as JSON snapshots of the deck's editable state (slides, theme,
 * size). Pictures live outside the snapshot (deck.media), so snapshots stay
 * small. Changes with the same `coalesce` key in quick succession (nudging
 * with the arrow keys, dragging a slider) undo as one step.
 */
const LIMIT = 100;
const COALESCE_MS = 1200;

export class History {
  constructor() {
    this.undoStack = [];
    this.redoStack = [];
    this.lastKey = null;
    this.lastTime = 0;
  }

  /** Remember `snapshot` (the state before a change). */
  push(snapshot, coalesce = null) {
    const now = Date.now();
    if (coalesce && coalesce === this.lastKey && now - this.lastTime < COALESCE_MS) {
      this.lastTime = now;
      return;
    }
    this.undoStack.push(snapshot);
    if (this.undoStack.length > LIMIT) this.undoStack.shift();
    this.redoStack.length = 0;
    this.lastKey = coalesce;
    this.lastTime = now;
  }

  /** Break coalescing, so the next change starts a new undo step. */
  seal() {
    this.lastKey = null;
  }

  undo(current) {
    if (!this.undoStack.length) return null;
    this.redoStack.push(current);
    this.lastKey = null;
    return this.undoStack.pop();
  }

  redo(current) {
    if (!this.redoStack.length) return null;
    this.undoStack.push(current);
    this.lastKey = null;
    return this.redoStack.pop();
  }

  get canUndo() { return this.undoStack.length > 0; }
  get canRedo() { return this.redoStack.length > 0; }
}
