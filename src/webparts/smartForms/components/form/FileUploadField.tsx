import * as React from 'react';
import { Icon, IconButton, MessageBar, MessageBarType } from '@fluentui/react';
import styles from './FormRenderer.module.scss';
import { IFormFile } from '../../models';
import { fileExtension } from '../../utils/formUtils';

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
    return bytes + ' B';
  }
  if (bytes < 1024 * 1024) {
    return Math.round(bytes / 1024) + ' KB';
  }
  return (bytes / (1024 * 1024)).toFixed(bytes < 10 * 1024 * 1024 ? 1 : 0) + ' MB';
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
  const inputRef = React.useRef<HTMLInputElement>(null);
  const [dragActive, setDragActive] = React.useState<boolean>(false);
  const [rejected, setRejected] = React.useState<string>('');
  const [reading, setReading] = React.useState<boolean>(false);

  const allowed = (props.allowedExtensions || []).map((e) => e.toLowerCase().replace(/^\./, ''));
  const maxBytes = (props.maxFileSizeMb || 10) * 1024 * 1024;
  const remaining = Math.max(0, (props.maxFiles || 3) - files.length);

  const accept = allowed.length > 0 ? allowed.map((e) => '.' + e).join(',') : undefined;

  const addFiles = async (incoming: FileList | undefined): Promise<void> => {
    if (!incoming || incoming.length === 0) {
      return;
    }
    const problems: string[] = [];
    const accepted: File[] = [];

    for (let i = 0; i < incoming.length; i++) {
      const file = incoming[i];
      if (accepted.length >= remaining) {
        problems.push('"' + file.name + '" was skipped — the limit is ' + (props.maxFiles || 3) + '.');
        continue;
      }
      if (allowed.length > 0 && allowed.indexOf(fileExtension(file.name)) === -1) {
        problems.push('"' + file.name + '" is not an accepted type.');
        continue;
      }
      if (file.size > maxBytes) {
        problems.push('"' + file.name + '" is larger than ' + (props.maxFileSizeMb || 10) + ' MB.');
        continue;
      }
      if (files.filter((existing) => existing.name === file.name).length > 0) {
        problems.push('"' + file.name + '" is already attached.');
        continue;
      }
      accepted.push(file);
    }

    setRejected(problems.join(' '));
    if (accepted.length === 0) {
      return;
    }

    setReading(true);
    try {
      const read: IFormFile[] = [];
      for (const file of accepted) {
        const content = await readAsBase64(file);
        read.push({ name: file.name, size: file.size, content });
      }
      props.onChange(files.concat(read));
    } catch {
      setRejected('One of the files could not be read. Try attaching it again.');
    }
    setReading(false);
  };

  const removeAt = (index: number): void => {
    props.onChange(files.filter((_f, i) => i !== index));
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
  hintParts.push('up to ' + (props.maxFileSizeMb || 10) + ' MB each');
  hintParts.push((props.maxFiles || 3) === 1 ? '1 file' : 'max ' + (props.maxFiles || 3) + ' files');

  return (
    <div>
      {remaining > 0 && !props.disabled && (
        <div
          className={dragActive ? styles.dropZoneActive : styles.dropZone}
          role="button"
          tabIndex={0}
          aria-label={props.ariaLabel}
          aria-describedby={props.ariaDescribedBy}
          aria-invalid={props.invalid ? true : undefined}
          onClick={openPicker}
          onKeyDown={(event) => {
            if (event.key === 'Enter' || event.key === ' ') {
              event.preventDefault();
              openPicker();
            }
          }}
          onDragOver={(event) => {
            event.preventDefault();
            setDragActive(true);
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
            {reading ? 'Reading files…' : 'Drop files here, or click to browse'}
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
                  title={'Remove ' + file.name}
                  ariaLabel={'Remove ' + file.name}
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
