/**
 * Allowlist HTML sanitizer.
 *
 * Rich text answers are stored in a SharePoint `RichTextMode="FullHtml"` column,
 * which means the raw markup is later rendered by the out-of-the-box list view,
 * Microsoft Lists, search previews and anything else reading the column. Smart
 * Forms strips tags for its own display, so an unsanitized value is harmless
 * *here* and dangerous everywhere else — sanitize on the way in instead.
 *
 * The implementation walks a detached DOM built with DOMParser, so nothing is
 * ever attached to the live document and no network or script side effect can
 * fire during parsing.
 */

/** Tags kept in sanitized output. Everything else is unwrapped or dropped. */
export const ALLOWED_TAGS: string[] = [
  'B',
  'STRONG',
  'I',
  'EM',
  'U',
  'S',
  'STRIKE',
  'P',
  'BR',
  'DIV',
  'SPAN',
  'UL',
  'OL',
  'LI',
  'A',
  'H3',
  'H4',
  'BLOCKQUOTE',
  'CODE',
  'PRE'
];

/** Tags whose entire subtree is discarded rather than unwrapped. */
export const DROP_SUBTREE_TAGS: string[] = [
  'SCRIPT',
  'STYLE',
  'IFRAME',
  'OBJECT',
  'EMBED',
  'FORM',
  'INPUT',
  'BUTTON',
  'SELECT',
  'TEXTAREA',
  'LINK',
  'META',
  'BASE',
  'SVG',
  'MATH',
  'TEMPLATE',
  'NOSCRIPT'
];

/** Per-tag attribute allowlist. Any attribute not listed is removed. */
const ALLOWED_ATTRIBUTES: { [tag: string]: string[] } = {
  A: ['href', 'title', 'target', 'rel']
};

/** URL schemes permitted in href values. */
const SAFE_URL = /^(https?:\/\/|mailto:|tel:|\/|#)/i;

/**
 * Whitespace plus C0/C7F control characters.
 *
 * Matching control characters is the entire point here: `java\x0Ascript:alert(1)`
 * is a working URL in some parsers, so they have to be stripped before the scheme
 * is tested. The lint rule that objects to control characters in a regex is
 * exactly backwards for a sanitizer.
 */
// eslint-disable-next-line no-control-regex
const INVISIBLE = /[\s\x00-\x1F\x7F]/g;

/**
 * Whether an href may survive sanitization.
 *
 * Exported because it is the security-critical decision in this module and the
 * only part testable without a DOM: the tree walk below needs DOMParser, which
 * exists in the browser this actually runs in but not under the test runner.
 */
export const isSafeUrl = (value: string): boolean => {
  const cleaned = (value || '').replace(INVISIBLE, '');
  if (cleaned.length === 0) {
    return false;
  }
  // reject javascript:, data:, vbscript: and any other unlisted scheme
  return SAFE_URL.test(cleaned);
};

const stripAttributes = (element: Element): void => {
  const allowed = ALLOWED_ATTRIBUTES[element.tagName] || [];
  // copy the list first — removing while iterating a live NamedNodeMap skips entries
  const names: string[] = [];
  for (let i = 0; i < element.attributes.length; i++) {
    names.push(element.attributes[i].name);
  }
  names.forEach((name) => {
    const lower = name.toLowerCase();
    if (allowed.indexOf(lower) === -1) {
      element.removeAttribute(name);
      return;
    }
    if (lower === 'href' && !isSafeUrl(element.getAttribute('href') || '')) {
      element.removeAttribute(name);
    }
  });
  if (element.tagName === 'A' && element.getAttribute('href')) {
    // links from untrusted content open safely
    element.setAttribute('rel', 'noopener noreferrer');
  }
};

/** Replace an element with its children, preserving the text content. */
const unwrap = (element: Element): void => {
  const parent = element.parentNode;
  if (!parent) {
    return;
  }
  while (element.firstChild) {
    parent.insertBefore(element.firstChild, element);
  }
  parent.removeChild(element);
};

const NODE_ELEMENT = 1;
const NODE_COMMENT = 8;

const cleanNode = (node: Node): void => {
  // snapshot children — the walk mutates the tree
  const children: Node[] = [];
  for (let i = 0; i < node.childNodes.length; i++) {
    children.push(node.childNodes[i]);
  }

  children.forEach((child) => {
    if (child.nodeType === NODE_COMMENT) {
      if (child.parentNode) {
        child.parentNode.removeChild(child);
      }
      return;
    }
    if (child.nodeType !== NODE_ELEMENT) {
      return;
    }
    const element = child as Element;
    const tag = element.tagName;

    if (DROP_SUBTREE_TAGS.indexOf(tag) !== -1) {
      if (element.parentNode) {
        element.parentNode.removeChild(element);
      }
      return;
    }

    // clean descendants before deciding what to do with this element, so
    // unwrapping never re-introduces unsanitized markup
    cleanNode(element);

    if (ALLOWED_TAGS.indexOf(tag) === -1) {
      unwrap(element);
      return;
    }
    stripAttributes(element);
  });
};

/** Plain-text projection of HTML, used for tables, CSV and search. */
export const stripHtml = (html: string): string =>
  String(html || '')
    .replace(/<(script|style)[\s\S]*?<\/\1>/gi, ' ')
    .replace(/<[^>]*>/g, ' ')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/\s+/g, ' ')
    .trim();

/**
 * Return `html` reduced to the allowlisted subset. Text content is preserved
 * for disallowed tags; script-bearing subtrees are removed outright.
 */
export const sanitizeHtml = (html: string): string => {
  const input = html || '';
  if (input.indexOf('<') === -1) {
    // no markup at all — nothing to strip
    return input;
  }
  if (typeof DOMParser === 'undefined') {
    // non-browser host (tests): fall back to removing every tag
    return stripHtml(input);
  }
  const parsed = new DOMParser().parseFromString('<body>' + input + '</body>', 'text/html');
  const body = parsed.body;
  if (!body) {
    return stripHtml(input);
  }
  cleanNode(body);
  return body.innerHTML;
};

/**
 * Sanitize a clipboard payload before it lands in a contenteditable. Falls back
 * to the plain-text flavor when no HTML flavor is present.
 */
export const sanitizePastedData = (data: DataTransfer | undefined): string => {
  if (!data) {
    return '';
  }
  const html = data.getData('text/html');
  if (html) {
    return sanitizeHtml(html);
  }
  const text = data.getData('text/plain') || '';
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/\r?\n/g, '<br />');
};
