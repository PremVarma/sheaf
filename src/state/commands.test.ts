import { describe, expect, it } from 'vitest';
import { commandForKeyEvent } from './commands';

const key = (k: string, mods: Partial<Record<'metaKey' | 'ctrlKey' | 'altKey' | 'shiftKey', boolean>> = {}) => ({
  key: k,
  metaKey: false,
  ctrlKey: false,
  altKey: false,
  shiftKey: false,
  ...mods,
});

describe('commandForKeyEvent', () => {
  it('maps the documented shortcuts on macOS (Cmd)', () => {
    expect(commandForKeyEvent(key('o', { metaKey: true }), true)).toBe('file.open');
    expect(commandForKeyEvent(key('f', { metaKey: true }), true)).toBe('edit.find');
    expect(commandForKeyEvent(key('g', { metaKey: true }), true)).toBe('edit.goTo');
    expect(commandForKeyEvent(key('=', { metaKey: true }), true)).toBe('view.zoomIn');
    expect(commandForKeyEvent(key('+', { metaKey: true, shiftKey: true }), true)).toBe('view.zoomIn');
    expect(commandForKeyEvent(key('-', { metaKey: true }), true)).toBe('view.zoomOut');
    expect(commandForKeyEvent(key('0', { metaKey: true }), true)).toBe('view.zoomReset');
  });

  it('uses Ctrl on Windows and ignores Cmd-only combinations there', () => {
    expect(commandForKeyEvent(key('o', { ctrlKey: true }), false)).toBe('file.open');
    expect(commandForKeyEvent(key('o', { metaKey: true }), false)).toBeNull();
    expect(commandForKeyEvent(key('o', { ctrlKey: true }), true)).toBeNull();
  });

  it('leaves copy, select all and plain typing to the focused element', () => {
    expect(commandForKeyEvent(key('c', { metaKey: true }), true)).toBeNull();
    expect(commandForKeyEvent(key('a', { metaKey: true }), true)).toBeNull();
    expect(commandForKeyEvent(key('o'), true)).toBeNull();
  });

  it('maps sheet navigation and find next', () => {
    expect(commandForKeyEvent(key('PageDown', { ctrlKey: true }), true)).toBe('view.nextSheet');
    expect(commandForKeyEvent(key('PageUp', { ctrlKey: true }), false)).toBe('view.previousSheet');
    expect(commandForKeyEvent(key('F3'), true)).toBe('edit.findNext');
    expect(commandForKeyEvent(key('F3', { shiftKey: true }), true)).toBe('edit.findPrevious');
  });

  it('maps formatting and editing shortcuts', () => {
    expect(commandForKeyEvent(key('b', { metaKey: true }), true)).toBe('format.bold');
    expect(commandForKeyEvent(key('i', { metaKey: true }), true)).toBe('format.italic');
    expect(commandForKeyEvent(key('x', { metaKey: true, shiftKey: true }), true)).toBe('format.strikethrough');
    expect(commandForKeyEvent(key('5', { ctrlKey: true }), false)).toBe('format.strikethrough');
    expect(commandForKeyEvent({ ...key('%', { metaKey: true, shiftKey: true }), code: 'Digit5' }, true)).toBe('format.numberPercent');
    expect(commandForKeyEvent(key('d', { metaKey: true }), true)).toBe('edit.fillDown');
    expect(commandForKeyEvent(key('v', { metaKey: true, shiftKey: true }), true)).toBe('edit.pasteValues');
    expect(commandForKeyEvent(key('h', { ctrlKey: true }), false)).toBe('edit.replace');
    expect(commandForKeyEvent(key('h', { metaKey: true, shiftKey: true }), true)).toBe('edit.replace');
  });

  it('uses physical keys for Option shortcuts on macOS', () => {
    expect(commandForKeyEvent({ ...key('≠', { metaKey: true, altKey: true }), code: 'Equal' }, true)).toBe('insert.cells');
    expect(commandForKeyEvent({ ...key('–', { metaKey: true, altKey: true }), code: 'Minus' }, true)).toBe('delete.cells');
    expect(commandForKeyEvent({ ...key('ˆ', { metaKey: true, altKey: true }), code: 'KeyI' }, true)).toBe('view.toggleInfo');
  });
});
