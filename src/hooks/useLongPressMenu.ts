import { useEffect } from 'react';

/**
 * iOS WebKit doesn't fire `contextmenu` on a long press (Android does). Turns a touch held
 * still for half a second into one, so right-click menus also open by touch.
 */
export function useLongPressMenu(enabled: boolean): void {
  useEffect(() => {
    if (!enabled) return;
    let timer = 0;
    let press: { x: number; y: number; target: Element } | null = null;
    let suppressClickUntil = 0;

    const cancel = () => {
      window.clearTimeout(timer);
      press = null;
    };
    const onDown = (event: PointerEvent) => {
      if (event.pointerType !== 'touch' || !event.isPrimary) return;
      const target = event.target instanceof Element ? event.target : null;
      if (!target || target.closest('input, textarea, select, [contenteditable="true"]')) return;
      cancel();
      press = { x: event.clientX, y: event.clientY, target };
      timer = window.setTimeout(() => {
        if (!press) return;
        const { x, y, target: element } = press;
        press = null;
        suppressClickUntil = Date.now() + 800;
        element.dispatchEvent(new MouseEvent('contextmenu', { bubbles: true, cancelable: true, clientX: x, clientY: y, button: 2, buttons: 2 }));
      }, 500);
    };
    const onMove = (event: PointerEvent) => {
      if (press && Math.hypot(event.clientX - press.x, event.clientY - press.y) > 10) cancel();
    };
    // The finger lifts over the menu that just opened; that isn't a click on it.
    const onClick = (event: MouseEvent) => {
      if (Date.now() >= suppressClickUntil) return;
      suppressClickUntil = 0;
      event.preventDefault();
      event.stopPropagation();
    };

    window.addEventListener('pointerdown', onDown, true);
    window.addEventListener('pointermove', onMove, true);
    window.addEventListener('pointerup', cancel, true);
    window.addEventListener('pointercancel', cancel, true);
    window.addEventListener('scroll', cancel, true);
    window.addEventListener('click', onClick, true);
    return () => {
      cancel();
      window.removeEventListener('pointerdown', onDown, true);
      window.removeEventListener('pointermove', onMove, true);
      window.removeEventListener('pointerup', cancel, true);
      window.removeEventListener('pointercancel', cancel, true);
      window.removeEventListener('scroll', cancel, true);
      window.removeEventListener('click', onClick, true);
    };
  }, [enabled]);
}
