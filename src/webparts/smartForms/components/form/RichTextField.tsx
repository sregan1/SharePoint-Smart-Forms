import * as React from 'react';
import { IconButton } from '@fluentui/react';
import styles from './FormRenderer.module.scss';
import { sanitizeHtml, sanitizePastedData } from '../../utils/sanitizeHtml';
import * as strings from 'SmartFormsWebPartStrings';

export interface IRichTextFieldProps {
  value: string;
  placeholder?: string;
  rows?: number;
  disabled?: boolean;
  ariaLabel?: string;
  ariaDescribedBy?: string;
  invalid?: boolean;
  onChange: (html: string) => void;
}

interface ICommand {
  icon: string;
  title: string;
  command: string;
}

const COMMANDS: ICommand[] = [
  { icon: 'Bold', title: strings.Form_RichText_Bold, command: 'bold' },
  { icon: 'Italic', title: strings.Form_RichText_Italic, command: 'italic' },
  { icon: 'Underline', title: strings.Form_RichText_Underline, command: 'underline' },
  { icon: 'BulletedList', title: strings.Form_RichText_BulletedList, command: 'insertUnorderedList' },
  { icon: 'NumberedList', title: strings.Form_RichText_NumberedList, command: 'insertOrderedList' },
  { icon: 'RemoveFormat', title: strings.Form_RichText_ClearFormatting, command: 'removeFormat' }
];

/**
 * Lightweight rich text editor storing HTML, matching what a SharePoint
 * enhanced rich text column expects.
 *
 * Everything leaving this component is run through the allowlist sanitizer. The
 * value ends up in a `RichTextMode="FullHtml"` column that the out-of-the-box
 * list view renders verbatim, so a pasted `<img onerror>` or `<script>` would be
 * a stored-XSS payload for every other reader of the list even though this UI
 * strips tags for its own display. Paste is intercepted for the same reason —
 * waiting until submit would leave hostile markup live in the editor's DOM.
 */
export const RichTextField: React.FunctionComponent<IRichTextFieldProps> = (props) => {
  const editorRef = React.useRef<HTMLDivElement>(null);
  // what we last pushed out, so an external value change can be told apart from
  // our own echo and not reset the caret mid-typing
  const lastEmitted = React.useRef<string>(props.value || '');

  React.useEffect(() => {
    const editor = editorRef.current;
    if (!editor) {
      return;
    }
    const incoming = props.value || '';
    if (incoming !== lastEmitted.current && incoming !== editor.innerHTML) {
      editor.innerHTML = sanitizeHtml(incoming);
      lastEmitted.current = incoming;
    }
  }, [props.value]);

  const emitChange = (): void => {
    const editor = editorRef.current;
    if (!editor) {
      return;
    }
    const clean = sanitizeHtml(editor.innerHTML);
    if (clean !== editor.innerHTML) {
      // markup was rejected — replace it in place so what's on screen is what
      // will be stored
      editor.innerHTML = clean;
    }
    lastEmitted.current = clean;
    props.onChange(clean);
  };

  const exec = (command: string): void => {
    if (props.disabled) {
      return;
    }
    if (editorRef.current) {
      editorRef.current.focus();
    }
    // execCommand is deprecated but remains the only cross-browser way to apply
    // inline formatting to a contenteditable selection without shipping an editor
    document.execCommand(command, false, undefined);
    emitChange();
  };

  const handlePaste = (event: React.ClipboardEvent<HTMLDivElement>): void => {
    event.preventDefault();
    const clean = sanitizePastedData(event.clipboardData);
    if (clean) {
      document.execCommand('insertHTML', false, clean);
    }
    emitChange();
  };

  const minHeight = props.rows ? Math.max(48, props.rows * 22) : undefined;

  return (
    <div className={styles.richText}>
      {!props.disabled && (
        <div className={styles.richTextToolbar} role="toolbar" aria-label={strings.Form_RichText_ToolbarAria}>
          {COMMANDS.map((c) => (
            <IconButton
              key={c.command}
              iconProps={{ iconName: c.icon }}
              title={c.title}
              ariaLabel={c.title}
              disabled={props.disabled}
              onMouseDown={(event) => {
                // keep the selection: focusing the button would collapse it
                event.preventDefault();
                exec(c.command);
              }}
            />
          ))}
        </div>
      )}
      <div
        ref={editorRef}
        className={styles.richTextEditor}
        style={minHeight ? { minHeight: minHeight } : undefined}
        contentEditable={!props.disabled}
        role="textbox"
        aria-multiline="true"
        aria-label={props.ariaLabel}
        aria-describedby={props.ariaDescribedBy}
        aria-invalid={props.invalid ? true : undefined}
        aria-disabled={props.disabled ? true : undefined}
        data-placeholder={props.placeholder || ''}
        suppressContentEditableWarning={true}
        onInput={emitChange}
        onBlur={emitChange}
        onPaste={handlePaste}
      />
    </div>
  );
};

export interface IRichTextViewProps {
  html: string;
}

/** Read-only rendering of a stored rich text answer. */
export const RichTextView: React.FunctionComponent<IRichTextViewProps> = (props) => {
  const clean = React.useMemo(() => sanitizeHtml(props.html || ''), [props.html]);
  if (!clean) {
    return null;
  }
  return (
    <div
      className={styles.richTextView}
      // sanitized immediately above, on every render
      dangerouslySetInnerHTML={{ __html: clean }}
    />
  );
};
