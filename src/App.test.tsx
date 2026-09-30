import { fireEvent, render, screen, waitFor, within } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { describe, expect, it } from 'vitest';
import App from './App';
import { createAppServices } from './state/AppContext';
import { buildWorkbookFile, createMockPlatform, sourceOf } from './test/fixtures';

const REPORT = buildWorkbookFile({
  Summary: [
    ['City', 'Sales'],
    ['Mumbai', 120],
    ['Delhi', 80],
    ['Mumbai', 45],
  ],
  Data: [
    ['Name', 'City'],
    ['Asha', 'Pune'],
    ['Ravi', 'Mumbai'],
  ],
});

async function openApp(bytes = REPORT, name = 'report.xlsx') {
  const platform = createMockPlatform();
  platform.nextFile = sourceOf(name, bytes, `/Users/me/${name}`);
  const services = createAppServices(platform);
  const user = userEvent.setup();
  render(<App services={services} />);
  await user.click(screen.getByRole('button', { name: 'Open File' }));
  return { platform, services, user };
}

const nameBox = () => screen.getByLabelText('Name Box: go to a cell') as HTMLInputElement;
const formulaBar = () => screen.getByTestId('formula-bar');

describe('App', () => {
  it('opens a workbook and shows its sheets', async () => {
    const { platform } = await openApp();
    expect(await screen.findByRole('grid', { name: 'Sheet Summary' })).toBeInTheDocument();
    const tabs = within(screen.getByRole('tablist', { name: 'Worksheets' })).getAllByRole('tab');
    expect(tabs.map((tab) => tab.textContent)).toEqual(['Summary', 'Data']);
    expect(tabs[0]).toHaveAttribute('aria-selected', 'true');
    expect(platform.title).toBe('report.xlsx — Sheaf');
    expect(platform.recent).toEqual(['/Users/me/report.xlsx']);
    expect(screen.getByTestId('sheet-dimensions')).toHaveTextContent('4 rows × 2 columns');
    expect(nameBox().value).toBe('A1');
    expect(formulaBar()).toHaveTextContent('City');
  });

  it('switches worksheets', async () => {
    const { user } = await openApp();
    await screen.findByRole('grid', { name: 'Sheet Summary' });
    await user.click(screen.getByRole('tab', { name: 'Data' }));
    expect(await screen.findByRole('grid', { name: 'Sheet Data' })).toBeInTheDocument();
    expect(screen.getByRole('tab', { name: 'Data' })).toHaveAttribute('aria-selected', 'true');
    expect(screen.getByTestId('sheet-dimensions')).toHaveTextContent('3 rows × 2 columns');
  });

  it('searches, counts, highlights and steps through matches', async () => {
    const { user } = await openApp();
    await screen.findByRole('grid');
    const search = screen.getByRole('textbox', { name: 'Search' });
    await user.type(search, 'mumbai');
    await waitFor(() => expect(screen.getByTestId('search-status')).toHaveTextContent('2 matches'));
    expect(document.querySelectorAll('.xv-match')).toHaveLength(2);

    await user.keyboard('{Enter}');
    expect(screen.getByTestId('search-status')).toHaveTextContent('1 of 2');
    expect(nameBox().value).toBe('A2');
    await user.keyboard('{Enter}');
    expect(nameBox().value).toBe('A4');
    await user.keyboard('{Shift>}{Enter}{/Shift}');
    expect(nameBox().value).toBe('A2');

    await user.click(screen.getByRole('button', { name: 'Workbook' }));
    await waitFor(() => expect(screen.getByTestId('search-status')).toHaveTextContent('3 matches'));

    await user.click(search);
    await user.keyboard('{Escape}');
    expect(search).toHaveValue('');
    expect(document.querySelectorAll('.xv-match')).toHaveLength(0);

    // Esc also closes a search while the grid has focus.
    await user.type(search, 'delhi');
    await waitFor(() => expect(screen.getByTestId('search-status')).toHaveTextContent('1 match'));
    screen.getByRole('grid').focus();
    await user.keyboard('{Escape}');
    expect(search).toHaveValue('');
    expect(document.activeElement).toHaveClass('xv-scroller');
  });

  it('jumps to a cell typed in the Name Box', async () => {
    const { user } = await openApp();
    await screen.findByRole('grid');
    await user.click(nameBox());
    await user.keyboard('B3{Enter}');
    expect(nameBox().value).toBe('B3');
    expect(formulaBar()).toHaveTextContent('80');
    expect(document.activeElement).toHaveClass('xv-scroller');

    await user.click(nameBox());
    await user.keyboard('nonsense{Enter}');
    expect(screen.getByRole('alert')).toHaveTextContent('Enter a cell reference');
  });

  it('zooms with the toolbar and shortcuts', async () => {
    const { user } = await openApp();
    await screen.findByRole('grid');
    await user.click(screen.getByRole('button', { name: 'Zoom in' }));
    expect(screen.getByRole('button', { name: 'Zoom 125%' })).toBeInTheDocument();
    fireEvent.keyDown(window, { key: '-', metaKey: true });
    fireEvent.keyDown(window, { key: '-', metaKey: true });
    expect(screen.getByRole('button', { name: 'Zoom 75%' })).toBeInTheDocument();
    fireEvent.keyDown(window, { key: '0', metaKey: true });
    expect(screen.getByRole('button', { name: 'Zoom 100%' })).toBeInTheDocument();
  });

  it('copies the selected range as tab-separated text', async () => {
    const { platform } = await openApp();
    const grid = await screen.findByRole('grid');
    fireEvent.keyDown(grid, { key: 'ArrowDown', shiftKey: true });
    fireEvent.keyDown(grid, { key: 'ArrowRight', shiftKey: true });
    fireEvent.keyDown(grid, { key: 'c', metaKey: true });
    await waitFor(() => expect(platform.clipboard).toEqual(['City\tSales\nMumbai\t120']));
    await waitFor(() => expect(screen.getByTestId('status-text')).toHaveTextContent('Copied 4 cells'));
    expect(screen.getByTestId('selection-stats')).toHaveTextContent('Count: 4');
  });

  it('explains files that cannot be opened, without technical details', async () => {
    const garbage = Uint8Array.from({ length: 2048 }, (_, i) => (i * 31 + 7) % 256);
    await openApp(garbage, 'broken.xlsx');
    const dialog = await screen.findByRole('alertdialog');
    expect(dialog).toHaveTextContent('Unable to open “broken.xlsx”');
    expect(dialog).toHaveTextContent('doesn’t appear to be a valid spreadsheet');
    expect(dialog.textContent).not.toMatch(/Error:|at |\.ts|\.js/);
    fireEvent.click(within(dialog).getByRole('button', { name: 'OK' }));
    expect(screen.queryByRole('alertdialog')).toBeNull();
    expect(screen.getByRole('button', { name: 'Open File' })).toBeInTheDocument();
  });

  it('shows workbook information', async () => {
    await openApp();
    await screen.findByRole('grid');
    fireEvent.keyDown(window, { key: 'ˆ', code: 'KeyI', metaKey: true, altKey: true });
    const panel = screen.getByRole('complementary', { name: 'Workbook information' });
    expect(panel).toHaveTextContent('report.xlsx');
    expect(panel).toHaveTextContent('Excel Workbook (.xlsx)');
    expect(panel).toHaveTextContent('Sheets2');
    expect(panel).toHaveTextContent('Active sheetSummary');
  });
});
