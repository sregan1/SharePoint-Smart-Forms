/** Replaces `{token}`-style placeholders in a loc string with the given values. */
export function formatString(template: string, params: { [key: string]: string | number }): string {
  return template.replace(/\{(\w+)\}/g, (match, key) => {
    const value = params[key];
    return value === undefined ? match : String(value);
  });
}

/** The SharePoint page's UI culture, for locale-aware date/number formatting. */
export function getEffectiveLocale(
  context: { pageContext?: { cultureInfo?: { currentUICultureName?: string } } } | undefined
): string {
  return context?.pageContext?.cultureInfo?.currentUICultureName || 'en-US';
}

const RTL_LANGUAGES = ['ar', 'he', 'fa', 'ur'];

/** True when the culture (e.g. 'ar-SA', 'he-IL') is written right-to-left. */
export function isRtlLocale(locale: string | undefined): boolean {
  const lang = (locale || '').split('-')[0].toLowerCase();
  return RTL_LANGUAGES.indexOf(lang) !== -1;
}
