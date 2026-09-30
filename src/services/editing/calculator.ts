import { stringAt } from '../../models/stringPool';
import { CellKind, type WorkbookModel, type WorksheetModel } from '../../models/workbook';
import { collectRefs, type FormulaNode, type RefNode } from '../formula/ast';
import { evaluateFormula, type EvaluationContext } from '../formula/evaluator';
import { parseFormula } from '../formula/parser';
import { formulaError, type Matrix, type Scalar } from '../formula/values';
import { lowerBoundColumn } from '../workbook/sheetService';
import { globalKey, globalLocal, globalSheet, keyCol, keyRow } from './cells';

/** Ranges up to this many cells are indexed cell by cell; bigger ones by column. */
const SMALL_RANGE = 64;
/** Columns a big range may span and still be indexed per column. */
const MAX_BUCKET_COLUMNS = 26;

interface Dependency {
  sheet: number;
  r0: number;
  c0: number;
  r1: number;
  c1: number;
}

interface CompiledFormula {
  ast: FormulaNode | null;
  deps: Dependency[];
}

interface RangeWatcher extends Dependency {
  formula: number;
}

/** Marker for formulas we couldn't evaluate: they keep the value Excel saved. */
const KEEP = Symbol('keep');

export interface RecalcResult {
  /** New values of recalculated formula cells, by global key. */
  values: Map<number, Scalar>;
  /** Formulas (global keys) left at their saved value because they use unsupported features. */
  unsupported: number[];
  circular: boolean;
}

function storedValue(sheet: WorksheetModel, i: number): Scalar {
  const { kinds, numbers, texts } = sheet.cells;
  switch (kinds[i]) {
    case CellKind.Number:
    case CellKind.Date:
      return numbers[i];
    case CellKind.Boolean:
      return numbers[i] === 1;
    case CellKind.String:
      return stringAt(sheet.strings, texts[i]);
    case CellKind.Error:
      return formulaError(stringAt(sheet.strings, texts[i]) as Parameters<typeof formulaError>[0]);
    default:
      return null;
  }
}

/**
 * Keeps a dependency index of every formula in the workbook and recalculates
 * the formulas affected by an edit. Evaluation is pull-based: a dirty formula
 * computes its dirty precedents first, so no explicit topological sort is needed
 * and cycles are detected while evaluating.
 */
export class Calculator {
  private readonly getWorkbook: () => WorkbookModel;
  private built = false;
  private readonly formulas = new Map<number, CompiledFormula>();
  private readonly cellWatchers = new Map<number, Set<number>>();
  /** Big ranges, bucketed by sheet and column. */
  private readonly columnWatchers = new Map<number, RangeWatcher[]>();
  private wideWatchers: RangeWatcher[] = [];

  constructor(getWorkbook: () => WorkbookModel) {
    this.getWorkbook = getWorkbook;
  }

  private sheetIndex(name: string): number {
    const lower = name.toLowerCase();
    return this.getWorkbook().sheets.findIndex((s) => s.name.toLowerCase() === lower);
  }

  private ensureBuilt(): void {
    if (this.built) return;
    this.built = true;
    this.getWorkbook().sheets.forEach((sheet, s) => {
      const { rowStart, cols } = sheet.cells;
      const { cells, texts } = sheet.formulas;
      let fp = 0;
      for (let r = 0; r < sheet.rowCount && fp < cells.length; r++) {
        const end = rowStart[r + 1];
        while (fp < cells.length && cells[fp] < end) {
          this.register(globalKey(s, r, cols[cells[fp]]), s, texts[fp]);
          fp++;
        }
      }
    });
  }

  private register(key: number, sheet: number, formula: string): void {
    let ast: FormulaNode | null = null;
    let refs: RefNode[] = [];
    try {
      ast = parseFormula(formula);
      refs = collectRefs(ast);
    } catch {
      ast = null;
    }
    const deps: Dependency[] = [];
    for (const ref of refs) {
      const target = ref.sheet === undefined ? sheet : this.sheetIndex(ref.sheet);
      if (target < 0) continue;
      const dep = { sheet: target, r0: ref.r0, c0: ref.c0, r1: ref.r1, c1: ref.c1 };
      deps.push(dep);
      const size = (dep.r1 - dep.r0 + 1) * (dep.c1 - dep.c0 + 1);
      if (size <= SMALL_RANGE) {
        for (let r = dep.r0; r <= dep.r1; r++) {
          for (let c = dep.c0; c <= dep.c1; c++) {
            const cell = globalKey(target, r, c);
            let set = this.cellWatchers.get(cell);
            if (!set) this.cellWatchers.set(cell, (set = new Set()));
            set.add(key);
          }
        }
      } else if (dep.c1 - dep.c0 < MAX_BUCKET_COLUMNS) {
        for (let c = dep.c0; c <= dep.c1; c++) {
          const bucket = target * 16_384 + c;
          let list = this.columnWatchers.get(bucket);
          if (!list) this.columnWatchers.set(bucket, (list = []));
          list.push({ ...dep, formula: key });
        }
      } else {
        this.wideWatchers.push({ ...dep, formula: key });
      }
    }
    this.formulas.set(key, { ast, deps });
  }

  private unregister(key: number): void {
    const compiled = this.formulas.get(key);
    if (!compiled) return;
    this.formulas.delete(key);
    for (const dep of compiled.deps) {
      const size = (dep.r1 - dep.r0 + 1) * (dep.c1 - dep.c0 + 1);
      if (size <= SMALL_RANGE) {
        for (let r = dep.r0; r <= dep.r1; r++) {
          for (let c = dep.c0; c <= dep.c1; c++) this.cellWatchers.get(globalKey(dep.sheet, r, c))?.delete(key);
        }
      } else if (dep.c1 - dep.c0 < MAX_BUCKET_COLUMNS) {
        for (let c = dep.c0; c <= dep.c1; c++) {
          const bucket = dep.sheet * 16_384 + c;
          const list = this.columnWatchers.get(bucket);
          if (list) this.columnWatchers.set(bucket, list.filter((w) => w.formula !== key));
        }
      } else {
        this.wideWatchers = this.wideWatchers.filter((w) => w.formula !== key);
      }
    }
  }

  /** Keeps the index in sync after a cell's formula was added, changed or removed. */
  formulaChanged(sheet: number, row: number, col: number, formula: string | undefined): void {
    if (!this.built) return; // indexed lazily on first recalculation
    const key = globalKey(sheet, row, col);
    this.unregister(key);
    if (formula) this.register(key, sheet, formula);
  }

  private dependentsOf(key: number, visit: (formula: number) => void): void {
    this.cellWatchers.get(key)?.forEach(visit);
    const sheet = globalSheet(key);
    const local = globalLocal(key);
    const row = keyRow(local);
    const col = keyCol(local);
    const bucket = this.columnWatchers.get(sheet * 16_384 + col);
    if (bucket) for (const w of bucket) if (row >= w.r0 && row <= w.r1) visit(w.formula);
    for (const w of this.wideWatchers) {
      if (w.sheet === sheet && row >= w.r0 && row <= w.r1 && col >= w.c0 && col <= w.c1) visit(w.formula);
    }
  }

  /**
   * Recalculates everything that depends on `changed`, plus the formulas in
   * `formulas` (cells that were just given a formula).
   */
  recalculate(changed: Iterable<number>, formulas: Iterable<number> = []): RecalcResult {
    this.ensureBuilt();
    const dirty = new Set<number>();
    const queue: number[] = [];
    const mark = (formula: number) => {
      if (!dirty.has(formula) && this.formulas.has(formula)) {
        dirty.add(formula);
        queue.push(formula);
      }
    };
    for (const key of formulas) mark(key);
    for (const key of changed) this.dependentsOf(key, mark);
    while (queue.length > 0) this.dependentsOf(queue.pop()!, mark);

    const workbook = this.getWorkbook();
    const results = new Map<number, Scalar | typeof KEEP>();
    const computing = new Set<number>();
    const unsupported: number[] = [];
    let circular = false;

    const valueAt = (sheetIndex: number, row: number, col: number): Scalar => {
      const key = globalKey(sheetIndex, row, col);
      if (dirty.has(key)) {
        const result = results.has(key) ? results.get(key)! : compute(key);
        if (result !== KEEP) return result;
      }
      const sheet = workbook.sheets[sheetIndex];
      if (!sheet || row >= sheet.rowCount) return null;
      const { rowStart, cols } = sheet.cells;
      const i = lowerBoundColumn(cols, rowStart[row], rowStart[row + 1], col);
      return i < rowStart[row + 1] && cols[i] === col ? storedValue(sheet, i) : null;
    };

    const context: EvaluationContext = {
      sheetIndex: (name) => this.sheetIndex(name),
      cell: valueAt,
      range: (sheetIndex, ref) => {
        const sheet = workbook.sheets[sheetIndex];
        if (!sheet) return [[formulaError('#REF!')]];
        const r1 = Math.min(ref.r1, Math.max(ref.r0, sheet.rowCount - 1));
        const c1 = Math.min(ref.c1, Math.max(ref.c0, sheet.columnCount - 1));
        const matrix: Matrix = [];
        const { rowStart, cols } = sheet.cells;
        for (let r = ref.r0; r <= r1; r++) {
          const row: Scalar[] = new Array(c1 - ref.c0 + 1).fill(null);
          if (r < sheet.rowCount) {
            const end = rowStart[r + 1];
            for (let i = lowerBoundColumn(cols, rowStart[r], end, ref.c0); i < end && cols[i] <= c1; i++) {
              row[cols[i] - ref.c0] = valueAt(sheetIndex, r, cols[i]);
            }
          }
          matrix.push(row);
        }
        return matrix;
      },
    };

    const compute = (key: number): Scalar | typeof KEEP => {
      if (computing.has(key)) {
        circular = true;
        return 0; // Excel shows 0 for circular references
      }
      const compiled = this.formulas.get(key);
      let result: Scalar | typeof KEEP = KEEP;
      if (compiled?.ast) {
        computing.add(key);
        try {
          const local = globalLocal(key);
          result = evaluateFormula(compiled.ast, { sheet: globalSheet(key), row: keyRow(local), col: keyCol(local) }, context);
        } catch {
          result = KEEP;
        }
        computing.delete(key);
      }
      if (result === KEEP) unsupported.push(key);
      results.set(key, result);
      return result;
    };

    for (const key of dirty) if (!results.has(key)) compute(key);

    const values = new Map<number, Scalar>();
    results.forEach((value, key) => {
      if (value !== KEEP) values.set(key, value);
    });
    return { values, unsupported, circular };
  }

}
