import { act, fireEvent, render, screen, waitFor } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import * as XLSX from 'xlsx';
import { describe, expect, it } from 'vitest';
import App from './App';
import { parseWorkbookBytes } from './services/workbook/parser';
import { getCell } from './services/workbook/sheetService';
import { createAppServices } from './state/AppContext';
import { buildWorkbookFile, createMockPlatform, sourceOf } from './test/fixtures';

const INVOICE = buildWorkbookFile({
  Invoice: [
    ['Item', 'Qty', 'Rate', 'Amount'],
    ['Pens', 10, 5, { t: 'n', v: 50, f: 'B2*C2' } as XLSX.CellObject],
    ['Total', null, null, { t: 'n', v: 50, f: 'SUM(D2:D2)' } as XLSX.CellObject],
  ],
});

async function openInvoice() {
  const platform = createMockPlatform();
  platform.nextFile = sourceOf('invoice.xlsx', INVOICE, '/Users/me/invoice.xlsx');
  const services = createAppServices(platform);
  const user = userEvent.setup();
  render(<App services={services} />);
  await user.click(screen.getByRole('button', { name: 'Open File' }));
  const grid = await screen.findByRole('grid');
  return { platform, services, user, grid };
}

const nameBox = () => screen.getByLabelText('Name Box: go to a cell') as HTMLInputElement;
const formulaBar = () => screen.getByTestId('formula-bar') as HTMLTextAreaElement;
const goTo = async (user: ReturnType<typeof userEvent.setup>, ref: string) => {
  await user.click(nameBox());
  await user.keyboard(`${ref}{Enter}`);
};
const cellValue = (services: ReturnType<typeof createAppServices>, row: number, col: number) =>
  getCell(services.store.getState().workbook!.sheets[0], row, col)?.value;

describe('editing', () => {
  it('types into a cell, recalculates formulas and moves down on Enter', async () => {
    const { services, user, platform } = await openInvoice();
    await goTo(user, 'B2');
    await user.keyboard('12');
    expect(screen.getByLabelText('Edit cell')).toHaveValue('12');
    expect(formulaBar()).toHaveValue('12');
    await user.keyboard('{Enter}');

    await waitFor(() => expect(cellValue(services, 1, 1)).toBe(12));
    expect(cellValue(services, 1, 3)).toBe(60);
    expect(cellValue(services, 2, 3)).toBe(60);
    expect(nameBox().value).toBe('B3');
    expect(screen.getByTestId('document-state')).toHaveTextContent('Unsaved changes');
    expect(platform.title).toBe('invoice.xlsx — Edited — Sheaf');
    expect(platform.edited).toBe(true);
  });

  it('enters formulas, edits with F2 and cancels with Esc', async () => {
    const { services, user } = await openInvoice();
    await goTo(user, 'E2');
    await user.keyboard('=D2*2{Enter}');
    await waitFor(() => expect(cellValue(services, 1, 4)).toBe(100));

    await goTo(user, 'E2');
    expect(formulaBar()).toHaveValue('=D2*2');
    await user.keyboard('{F2}');
    expect(screen.getByLabelText('Edit cell')).toHaveValue('=D2*2');
    await user.keyboard('0{Escape}');
    expect(screen.queryByLabelText('Edit cell')).toBeNull();
    expect(cellValue(services, 1, 4)).toBe(100);
  });

  it('undoes and redoes with the keyboard', async () => {
    const { services, user } = await openInvoice();
    await goTo(user, 'C2');
    await user.keyboard('7{Enter}');
    await waitFor(() => expect(cellValue(services, 1, 3)).toBe(70));

    fireEvent.keyDown(window, { key: 'z', metaKey: true });
    await waitFor(() => expect(cellValue(services, 1, 3)).toBe(50));
    expect(screen.getByTestId('document-state')).toHaveTextContent('Saved');
    fireEvent.keyDown(window, { key: 'z', metaKey: true, shiftKey: true });
    await waitFor(() => expect(cellValue(services, 1, 3)).toBe(70));
  });

  it('clears the selection with Delete and pastes tab-separated text', async () => {
    const { services, user, platform, grid } = await openInvoice();
    await goTo(user, 'B2');
    fireEvent.keyDown(grid, { key: 'Delete' });
    await waitFor(() => expect(cellValue(services, 1, 3)).toBe(0));

    platform.clipboardText = 'Ink\t3\t20\nTape\t1\t15\n';
    await goTo(user, 'A5');
    fireEvent.keyDown(grid, { key: 'v', metaKey: true });
    await waitFor(() => expect(cellValue(services, 5, 2)).toBe(15));
    expect(cellValue(services, 4, 0)).toBe('Ink');
    expect(nameBox().value).toBe('A5:C6');
  });

  it('saves into the original file, keeping it a valid workbook', async () => {
    const { services, user, platform } = await openInvoice();
    await goTo(user, 'B2');
    await user.keyboard('3{Enter}');
    await waitFor(() => expect(cellValue(services, 1, 3)).toBe(15));

    fireEvent.keyDown(window, { key: 's', metaKey: true });
    await waitFor(() => expect(platform.saved).toHaveLength(1));
    expect(platform.saved[0].path).toBe('/Users/me/invoice.xlsx');
    const reopened = parseWorkbookBytes(platform.saved[0].bytes, 'invoice.xlsx');
    expect(getCell(reopened.sheets[0], 1, 3)).toMatchObject({ value: 15, formula: '=B2*C2' });
    expect(getCell(reopened.sheets[0], 2, 3)?.value).toBe(15);
    await waitFor(() => expect(screen.getByTestId('document-state')).toHaveTextContent('Saved'));
    expect(platform.edited).toBe(false);
    expect(platform.title).toBe('invoice.xlsx — Sheaf');
  });

  it('asks before discarding unsaved changes', async () => {
    const { services, user, platform } = await openInvoice();
    await goTo(user, 'B2');
    await user.keyboard('9{Enter}');
    await waitFor(() => expect(cellValue(services, 1, 1)).toBe(9));

    platform.unsavedChoice = 'cancel';
    fireEvent.keyDown(window, { key: 'n', metaKey: true });
    await waitFor(() => expect(platform.unsavedPrompts).toBe(1));
    expect(services.store.getState().workbook?.fileName).toBe('invoice.xlsx');

    platform.unsavedChoice = 'discard';
    fireEvent.keyDown(window, { key: 'n', metaKey: true });
    await waitFor(() => expect(services.store.getState().workbook?.fileName).toBe('Untitled.xlsx'));
    expect(platform.unsavedPrompts).toBe(2);
  });

  it('creates a new workbook and saves it with Save As', async () => {
    const platform = createMockPlatform();
    const services = createAppServices(platform);
    const user = userEvent.setup();
    render(<App services={services} />);
    await user.click(screen.getByRole('button', { name: 'New workbook' }));
    await screen.findByRole('grid', { name: 'Sheet Sheet1' });
    await user.keyboard('42{Tab}=A1/2{Enter}');
    await waitFor(() => expect(cellValue(services, 0, 1)).toBe(21));

    platform.nextSavePath = '/Users/me/answers.xlsx';
    await act(async () => {
      fireEvent.keyDown(window, { key: 's', metaKey: true });
    });
    await waitFor(() => expect(platform.saved).toHaveLength(1));
    const reopened = parseWorkbookBytes(platform.saved[0].bytes, 'answers.xlsx');
    expect(getCell(reopened.sheets[0], 0, 1)).toMatchObject({ value: 21, formula: '=A1/2' });
    await waitFor(() => expect(platform.title).toBe('answers.xlsx — Sheaf'));
  });
});

describe('formatting and structure', () => {
  const styleOf = (services: ReturnType<typeof createAppServices>, row: number, col: number) => {
    const workbook = services.store.getState().workbook!;
    return workbook.styles[getCell(workbook.sheets[0], row, col)?.styleId ?? 0];
  };

  it('formats the selection from the toolbar and with shortcuts', async () => {
    const { services, user } = await openInvoice();
    await goTo(user, 'A1:D1');
    await user.click(screen.getByRole('button', { name: 'Bold' }));
    await waitFor(() => expect(styleOf(services, 0, 3)?.bold).toBe(true));
    expect(screen.getByRole('button', { name: 'Bold' })).toHaveAttribute('aria-pressed', 'true');
    fireEvent.keyDown(window, { key: 'b', metaKey: true });
    await waitFor(() => expect(styleOf(services, 0, 3)?.bold).toBeUndefined());
    fireEvent.keyDown(window, { key: 'i', metaKey: true });
    await waitFor(() => expect(styleOf(services, 0, 0)?.italic).toBe(true));
    await user.click(screen.getByRole('button', { name: 'Center' }));
    await waitFor(() => expect(styleOf(services, 0, 1)?.hAlign).toBe('center'));
  });

  it('inserts and deletes rows from the row header menu, keeping formulas right', async () => {
    const { services, grid } = await openInvoice();
    const header = Array.from(grid.querySelectorAll('.xv-rowhdr')).find((el) => el.textContent === '2') as HTMLElement;
    fireEvent.contextMenu(header, { clientX: 10, clientY: 50 });
    fireEvent.click(await screen.findByRole('menuitem', { name: 'Insert Row Above' }));
    await waitFor(() => expect(cellValue(services, 2, 0)).toBe('Pens'));
    expect(getCell(services.store.getState().workbook!.sheets[0], 3, 3)?.formula).toBe('=SUM(D3:D3)');

    fireEvent.contextMenu(header, { clientX: 10, clientY: 50 });
    fireEvent.click(await screen.findByRole('menuitem', { name: 'Delete Row' }));
    await waitFor(() => expect(cellValue(services, 1, 0)).toBe('Pens'));
    expect(getCell(services.store.getState().workbook!.sheets[0], 2, 3)?.formula).toBe('=SUM(D2:D2)');
  });

  it('adds, renames and deletes sheets from the tabs', async () => {
    const { services, user } = await openInvoice();
    await user.click(screen.getByRole('button', { name: 'New sheet' }));
    await waitFor(() => expect(services.store.getState().workbook!.sheets.map((s) => s.name)).toEqual(['Invoice', 'Sheet2']));
    expect(screen.getByRole('tab', { name: 'Sheet2' })).toHaveAttribute('aria-selected', 'true');

    await user.dblClick(screen.getByRole('tab', { name: 'Sheet2' }));
    const field = screen.getByLabelText('Sheet name');
    await user.clear(field);
    await user.type(field, 'Invoice{Enter}');
    expect(await screen.findByRole('alert')).toHaveTextContent('already a sheet named');
    await user.clear(field);
    await user.type(field, 'Notes{Enter}');
    await waitFor(() => expect(screen.getByRole('tab', { name: 'Notes' })).toBeInTheDocument());

    fireEvent.contextMenu(screen.getByRole('tab', { name: 'Notes' }), { clientX: 100, clientY: 700 });
    fireEvent.click(await screen.findByRole('menuitem', { name: 'Delete' }));
    await waitFor(() => expect(services.store.getState().workbook!.sheets.map((s) => s.name)).toEqual(['Invoice']));
  });

  it('asks before merging cells that would lose values', async () => {
    const { services, user } = await openInvoice();
    await goTo(user, 'A2:B2');
    await user.click(screen.getByRole('button', { name: 'Merge & Center' }));
    expect(await screen.findByRole('alertdialog')).toHaveTextContent('Merging keeps only the upper-left value');
    await user.click(screen.getByRole('button', { name: 'Merge' }));
    await waitFor(() => expect(services.store.getState().workbook!.sheets[0].merges).toEqual([{ r0: 1, c0: 0, r1: 1, c1: 1 }]));
    expect(cellValue(services, 1, 1)).toBeUndefined();
    expect(cellValue(services, 1, 3)).toBe(0);
  });

  it('copies and pastes inside the app with formulas moved', async () => {
    const { services, user, platform, grid } = await openInvoice();
    await goTo(user, 'D2');
    fireEvent.keyDown(grid, { key: 'c', metaKey: true });
    await waitFor(() => expect(platform.clipboard).toHaveLength(1));
    platform.clipboardText = platform.clipboard[0];
    await goTo(user, 'E3');
    fireEvent.keyDown(grid, { key: 'v', metaKey: true });
    await waitFor(() => expect(getCell(services.store.getState().workbook!.sheets[0], 2, 4)?.formula).toBe('=C3*D3'));
  });
});
