/** Commands shared by keyboard shortcuts, the native menu and toolbar buttons. */
export const COMMAND_IDS = [
  'app.quit',
  'file.new',
  'file.open',
  'file.close',
  'file.save',
  'file.saveAs',
  'edit.undo',
  'edit.redo',
  'edit.cut',
  'edit.paste',
  'edit.pasteValues',
  'edit.clear',
  'edit.copy',
  'edit.selectAll',
  'edit.fillDown',
  'edit.fillRight',
  'edit.find',
  'edit.replace',
  'edit.findNext',
  'edit.findPrevious',
  'edit.goTo',
  'insert.cells',
  'insert.rowsAbove',
  'insert.rowsBelow',
  'insert.columnsLeft',
  'insert.columnsRight',
  'insert.sheet',
  'insert.date',
  'insert.time',
  'delete.cells',
  'delete.rows',
  'delete.columns',
  'delete.sheet',
  'format.bold',
  'format.italic',
  'format.underline',
  'format.strikethrough',
  'format.increaseFont',
  'format.decreaseFont',
  'format.alignLeft',
  'format.alignCenter',
  'format.alignRight',
  'format.wrap',
  'format.mergeCenter',
  'format.unmerge',
  'format.numberGeneral',
  'format.numberNumber',
  'format.numberCurrency',
  'format.numberPercent',
  'format.numberDate',
  'format.increaseDecimals',
  'format.decreaseDecimals',
  'format.clear',
  'format.hideRows',
  'format.unhideRows',
  'format.hideColumns',
  'format.unhideColumns',
  'format.renameSheet',
  'data.sortAscending',
  'data.sortDescending',
  'view.zoomIn',
  'view.zoomOut',
  'view.zoomReset',
  'view.freezeTopRow',
  'view.freezeFirstColumn',
  'view.freezeAtSelection',
  'view.unfreeze',
  'view.toggleGridlines',
  'view.nextSheet',
  'view.previousSheet',
  'view.toggleInfo',
  'recent.clear',
] as const;

export type CommandId = (typeof COMMAND_IDS)[number];

export function isCommandId(value: string): value is CommandId {
  return (COMMAND_IDS as readonly string[]).includes(value);
}

/** Commands that text fields handle natively (undo typing) when they have focus. */
export const TEXT_FIELD_COMMANDS: ReadonlySet<CommandId> = new Set(['edit.undo', 'edit.redo']);

/** Commands that act on the selected cells; ignored while typing in the search or Name Box fields. */
export function isCellCommand(id: CommandId): boolean {
  return /^(format|insert|delete|data)\./.test(id) || id === 'edit.fillDown' || id === 'edit.fillRight' || id === 'edit.pasteValues';
}

type KeyInput = Pick<KeyboardEvent, 'key' | 'metaKey' | 'ctrlKey' | 'altKey' | 'shiftKey'> & { code?: string };

/**
 * Global shortcuts. Copy, cut, paste and Select All are deliberately absent:
 * the grid handles them itself so text fields keep their native behavior.
 */
export function commandForKeyEvent(event: KeyInput, isMac: boolean): CommandId | null {
  const primary = isMac ? event.metaKey : event.ctrlKey;
  const key = event.key;
  const code = event.code ?? '';
  if (primary && event.altKey) {
    // ⌥ changes the character on macOS, so these go by physical key.
    if (code === 'Equal' || key === '=' || key === '+') return 'insert.cells';
    if (code === 'Minus' || key === '-') return 'delete.cells';
    if (code === 'KeyI') return 'view.toggleInfo';
    return null;
  }
  if (primary && event.shiftKey) {
    switch (code) {
      case 'Digit1':
        return 'format.numberNumber';
      case 'Digit3':
        return 'format.numberDate';
      case 'Digit4':
        return 'format.numberCurrency';
      case 'Digit5':
        return 'format.numberPercent';
      case 'Backquote':
        return 'format.numberGeneral';
      case 'Period':
        return 'format.increaseFont';
      case 'Comma':
        return 'format.decreaseFont';
      case 'Semicolon':
        return 'insert.time';
      default:
        break;
    }
    switch (key.toLowerCase()) {
      case 's':
        return 'file.saveAs';
      case 'z':
        return 'edit.redo';
      case 'v':
        return 'edit.pasteValues';
      case 'x':
        return isMac ? 'format.strikethrough' : null;
      case 'h':
        return isMac ? 'edit.replace' : null;
      case 'l':
        return 'format.alignLeft';
      case 'e':
        return 'format.alignCenter';
      case 'r':
        return 'format.alignRight';
      case 'f':
        return 'edit.find';
      case '=':
      case '+':
        return 'view.zoomIn';
      case '_':
      case '-':
        return 'view.zoomOut';
      default:
        return null;
    }
  }
  if (primary) {
    if (code === 'Semicolon' || key === ';') return 'insert.date';
    if (code === 'Backslash' || key === '\\') return 'format.clear';
    switch (key.toLowerCase()) {
      case 'o':
        return 'file.open';
      case 'n':
        return 'file.new';
      case 's':
        return 'file.save';
      case 'z':
        return 'edit.undo';
      case 'y':
        return isMac ? null : 'edit.redo';
      case 'q':
        return isMac ? 'app.quit' : null;
      case 'w':
        return 'file.close';
      case 'f':
        return 'edit.find';
      case 'h':
        return isMac ? null : 'edit.replace';
      case 'g':
        return 'edit.goTo';
      case 'b':
        return 'format.bold';
      case 'i':
        return 'format.italic';
      case 'u':
        return 'format.underline';
      case '5':
        return isMac ? null : 'format.strikethrough';
      case 'd':
        return 'edit.fillDown';
      case 'r':
        return 'edit.fillRight';
      case '=':
      case '+':
        return 'view.zoomIn';
      case '-':
      case '_':
        return 'view.zoomOut';
      case '0':
        return 'view.zoomReset';
      default:
        break;
    }
  }
  if (event.ctrlKey && !event.altKey && key === 'PageDown') return 'view.nextSheet';
  if (event.ctrlKey && !event.altKey && key === 'PageUp') return 'view.previousSheet';
  if (key === 'F3') return event.shiftKey ? 'edit.findPrevious' : 'edit.findNext';
  if (key === 'F5' && !primary && !event.shiftKey) return 'edit.goTo';
  if (key === 'F11' && event.shiftKey) return 'insert.sheet';
  return null;
}
