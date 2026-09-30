import { useEffect } from 'react';
import { commandForKeyEvent, isCellCommand, TEXT_FIELD_COMMANDS } from '../state/commands';
import type { ViewerController } from '../state/controller';

function isTextField(target: EventTarget | null): boolean {
  return target instanceof HTMLElement && (target.isContentEditable || /^(INPUT|TEXTAREA|SELECT)$/.test(target.tagName));
}

/** The cell editor and the formula bar edit a cell, so cell commands (Bold…) still apply there. */
function isCellField(target: EventTarget | null): boolean {
  return target instanceof HTMLElement && target.closest('[data-cell-editor],[data-formula-bar]') !== null;
}

/** Routes global shortcuts (Cmd/Ctrl+O, F, G, +, −, 0 …) to commands. */
export function useKeyboardShortcuts(controller: ViewerController, isMac: boolean): void {
  useEffect(() => {
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.defaultPrevented || event.isComposing) return;
      if (event.key === 'Escape') {
        if (controller.handleEscape()) event.preventDefault();
        return;
      }
      const command = commandForKeyEvent(event, isMac);
      if (!command) return;
      // Undo/redo inside a text field undoes the typing there.
      if (TEXT_FIELD_COMMANDS.has(command) && isTextField(event.target)) return;
      // Formatting and row/column commands don't apply while typing a search or a sheet name.
      if (isCellCommand(command) && isTextField(event.target) && !isCellField(event.target)) return;
      event.preventDefault();
      controller.runCommand(command, 'keyboard');
    };
    window.addEventListener('keydown', onKeyDown);
    return () => window.removeEventListener('keydown', onKeyDown);
  }, [controller, isMac]);
}
