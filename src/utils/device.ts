/** Device checks for the few places where phones and tablets behave differently. */
const userAgent = typeof navigator === 'undefined' ? '' : navigator.userAgent;

export const isAndroid = /Android/.test(userAgent);
// iPadOS reports a Mac user agent; touch support tells them apart.
export const isIOS = /iPhone|iPad|iPod/.test(userAgent) || (/Mac/.test(userAgent) && typeof navigator !== 'undefined' && navigator.maxTouchPoints > 1);
