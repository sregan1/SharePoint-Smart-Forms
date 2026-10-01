import * as React from 'react';
import { Icon, IconButton, MessageBar, MessageBarType } from '@fluentui/react';
import styles from './FormRenderer.module.scss';
import { IFormFile } from '../../models';
import { fileExtension } from '../../utils/formUtils';
import * as strings from 'SmartFormsWebPartStrings';
import { formatString } from '../../utils/localeUtils';

export interface IFileUploadFieldProps {
  value: IFormFile[];
  maxFiles: number;
  maxFileSizeMb: number;
  allowedExtensions: string[];
  disabled?: boolean;
  ariaLabel?: string;
  ariaDescribedBy?: string;
  invalid?: boolean;
  onChange: (files: IFormFile[]) => void;
}

export const formatBytes = (bytes: number): string => {
  if (!bytes || bytes <= 0) {
    return '';
  }
  if (bytes < 1024) {
    return formatString(strings.Form_File_SizeBytes, { size: bytes });
  }
  if (bytes < 1024 * 1024) {
    return formatString(strings.Form_File_SizeKb, { size: Math.round(bytes / 1024) });
  }
  return formatString(strings.Form_File_SizeMb, { size: (bytes / (1024 * 1024)).toFixed(bytes < 10 * 1024 * 1024 ? 1 : 0) });
};

const ICON_FOR_EXTENSION: { [extension: string]: string } = {
  pdf: 'PDF',
  doc: 'WordDocument',
  docx: 'WordDocument',
  xls: 'ExcelDocument',
  xlsx: 'ExcelDocument',
  csv: 'ExcelDocument',
  ppt: 'PowerPointDocument',
  pptx: 'PowerPointDocument',
  png: 'FileImage',
  jpg: 'FileImage',
  jpeg: 'FileImage',
  gif: 'FileImage',
  svg: 'FileImage',
  zip: 'ZipFolder',
  txt: 'TextDocument',
  msg: 'Mail'
};

/** Read a File into the base64 payload the attachment API needs. */
const readAsBase64 = (file: File): Promise<string> =>
  new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      const result = String(reader.result || '');
      // strip the "data:<mime>;base64," prefix
      const comma = result.indexOf(',');
      resolve(comma >= 0 ? result.slice(comma + 1) : result);
    };
    reader.onerror = () => reject(new Error('read failed'));
    reader.readAsDataURL(file);
  });

/**
 * Drag-and-drop file picker.
 *
 * Files are held in form state as base64 and only become list item attachments
 * after the item itself is created — SharePoint has no way to attach to an item
 * that doesn't exist yet, so the submit path is create-then-attach.
 */
export const FileUploadField: React.FunctionComponent<IFileUploadFieldProps> = (props) => {
  const files = props.value || [];
  // always the latest committed list, so a slow read can't clobber files added meanwhile
  const filesRef = React.useRef<IFormFile[]>(files);
  filesRef.current = files;
  const readingRef = React.useRef<boolean>(false);
  const inputRef = React.useRef<HTMLInputElement>(null);
  const [dragActive, setDragActive] = React.useState<boolean>(false);
  const [rejected, setRejected] = React.useState<string>('');
  const [reading, setReading] = React.useState<boolean>(false);

  const allowed = (props.allowedExtensions || []).map((e) => e.toLowerCase().replace(/^\./, ''));
  const maxBytes = (props.maxFileSizeMb || 10) * 1024 * 1024;
  const remaining = Math.max(0, (props.maxFiles || 3) - files.length);

  const accept = allowed.length > 0 ? allowed.map((e) => '.' + e).join(',') : undefined;

  const addFiles = async (incoming: FileList | undefined): Promise<void> => {
    if (!incoming || incoming.length === 0 || readingRef.current || props.disabled) {
      return;
    }
    const problems: string[] = [];
    const accepted: File[] = [];

    for (let i = 0; i < incoming.length; i++) {
      const file = incoming[i];
      if (accepted.length >= remaining) {
        problems.push(formatString(strings.Form_File_SkippedLimit, { name: file.name, limit: props.maxFiles || 3 }));
        continue;
      }
      if (allowed.length > 0 && allowed.indexOf(fileExtension(file.name)) === -1) {
        problems.push(formatString(strings.Form_File_TypeNotAccepted, { name: file.name }));
        continue;
      }
      if (file.size > maxBytes) {
        problems.push(formatString(strings.Form_File_TooLarge, { name: file.name, size: props.maxFileSizeMb || 10 }));
        continue;
      }
      const lower = file.name.toLowerCase();
      if (
        files.some((existing) => existing.name.toLowerCase() === lower) ||
        accepted.some((a) => a.name.toLowerCase() === lower)
      ) {
        problems.push(formatString(strings.Form_File_AlreadyAttached, { name: file.name }));
        continue;
      }
      accepted.push(file);
    }

    setRejected(problems.join(' '));
    if (accepted.length === 0) {
      return;
    }

    readingRef.current = true;
    setReading(true);
    try {
      const read: IFormFile[] = [];
      for (const file of accepted) {
        const content = await readAsBase64(file);
        read.push({ name: file.name, size: file.size, content });
      }
      // merge into the latest list rather than the snapshot taken before the read
      const latest = filesRef.current;
      const known = latest.map((f) => f.name.toLowerCase());
      const room = Math.max(0, (props.maxFiles || 3) - latest.length);
      const fresh = read.filter((f) => known.indexOf(f.name.toLowerCase()) === -1).slice(0, room);
      if (fresh.length > 0) {
        const merged = latest.concat(fresh);
        filesRef.current = merged;
        props.onChange(merged);
      }
    } catch {
      setRejected(strings.Form_File_ReadError);
    }
    readingRef.current = false;
    setReading(false);
  };

  const removeAt = (index: number): void => {
    const next = filesRef.current.filter((_f, i) => i !== index);
    filesRef.current = next;
    props.onChange(next);
    setRejected('');
  };

  const openPicker = (): void => {
    if (inputRef.current) {
      inputRef.current.click();
    }
  };

  const hintParts: string[] = [];
  if (allowed.length > 0) {
    hintParts.push(allowed.join(', '));
  }
  hintParts.push(formatString(strings.Form_File_HintMaxSize, { size: props.maxFileSizeMb || 10 }));
  hintParts.push((props.maxFiles || 3) === 1
      ? strings.Form_File_HintOneFile
      : formatString(strings.Form_File_HintMaxFiles, { count: props.maxFiles || 3 }));

  return (
    <div>
      {remaining > 0 && !props.disabled && (
        <div
          className={dragActive ? styles.dropZoneActive : styles.dropZone}
          role="button"
          tabIndex={reading ? -1 : 0}
          aria-busy={reading ? true : undefined}
          aria-disabled={reading ? true : undefined}
          aria-label={props.ariaLabel}
          aria-describedby={props.ariaDescribedBy}
          aria-invalid={props.invalid ? true : undefined}
          onClick={() => {
            if (!reading) {
              openPicker();
            }
          }}
          onKeyDown={(event) => {
            if (!reading && (event.key === 'Enter' || event.key === ' ')) {
              event.preventDefault();
              openPicker();
            }
          }}
          onDragOver={(event) => {
            event.preventDefault();
            setDragActive(!reading);
          }}
          onDragLeave={() => setDragActive(false)}
          onDrop={(event) => {
            event.preventDefault();
            setDragActive(false);
            void addFiles(event.dataTransfer ? event.dataTransfer.files : undefined);
          }}
        >
          <Icon iconName={reading ? 'Sync' : 'CloudUpload'} className={styles.dropZoneIcon} />
          <span className={styles.dropZoneText}>
            {reading ? strings.Form_File_Reading : strings.Form_File_DropPrompt}
          </span>
          <span className={styles.dropZoneHint}>{hintParts.join(' · ')}</span>
        </div>
      )}

      <input
        ref={inputRef}
        type="file"
        className={styles.hiddenInput}
        multiple={(props.maxFiles || 3) > 1}
        accept={accept}
        tabIndex={-1}
        onChange={(event) => {
          void addFiles(event.target.files || undefined);
          // reset so picking the same file twice still fires a change
          event.target.value = '';
        }}
      />

      {rejected && (
        <MessageBar
          messageBarType={MessageBarType.warning}
          onDismiss={() => setRejected('')}
          styles={{ root: { marginTop: 8 } }}
        >
          {rejected}
        </MessageBar>
      )}

      {files.length > 0 && (
        <div className={styles.fileList}>
          {files.map((file, index) => (
            <div key={file.name + '-' + index} className={styles.fileRow}>
              <Icon
                className={styles.fileIcon}
                iconName={ICON_FOR_EXTENSION[fileExtension(file.name)] || 'Page'}
              />
              <span className={styles.fileName} title={file.name}>
                {file.name}
              </span>
              {file.size > 0 && <span className={styles.fileSize}>{formatBytes(file.size)}</span>}
              {!props.disabled && (
                <IconButton
                  iconProps={{ iconName: 'Delete' }}
                  title={formatString(strings.Form_File_Remove, { name: file.name })}
                  ariaLabel={formatString(strings.Form_File_Remove, { name: file.name })}
                  onClick={() => removeAt(index)}
                />
              )}
            </div>
          ))}
        </div>
      )}
    </div>
  );
};
