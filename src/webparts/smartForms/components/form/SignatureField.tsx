import * as React from 'react';
import { DefaultButton, TextField } from '@fluentui/react';
import styles from './FormRenderer.module.scss';
import * as strings from 'SmartFormsWebPartStrings';
import { formatString } from '../../utils/localeUtils';

export interface ISignatureFieldProps {
  /** data URL of the captured signature, or '' when unsigned */
  value: string;
  disabled?: boolean;
  ariaLabel?: string;
  ariaDescribedBy?: string;
  invalid?: boolean;
  /** ink color; defaults to the theme's body text color */
  inkColor?: string;
  onChange: (dataUrl: string) => void;
}

const WIDTH = 420;
const HEIGHT = 140;

interface IPoint {
  x: number;
  y: number;
}

/**
 * Freehand signature capture.
 *
 * The canvas is sized at a fixed logical resolution and scaled by
 * devicePixelRatio so a signature drawn on a high-DPI screen isn't stored
 * blurry. Pointer events cover mouse, pen and touch in one code path;
 * `touch-action: none` on the canvas (see the stylesheet) stops a finger drag
 * from scrolling the page instead of drawing.
 *
 * The signature is exported as a PNG data URL and becomes a list item
 * attachment on submit — a base64 PNG is far too large for a text column.
 */
export const SignatureField: React.FunctionComponent<ISignatureFieldProps> = (props) => {
  const canvasRef = React.useRef<HTMLCanvasElement>(null);
  const drawing = React.useRef<boolean>(false);
  const lastPoint = React.useRef<IPoint | undefined>(undefined);
  const [hasInk, setHasInk] = React.useState<boolean>(!!props.value);
  const [typing, setTyping] = React.useState<boolean>(false);
  const [typedName, setTypedName] = React.useState<string>('');

  const ink = props.inkColor || '#201f1e';

  const context = (): CanvasRenderingContext2D | undefined => {
    const canvas = canvasRef.current;
    if (!canvas) {
      return undefined;
    }
    const ctx = canvas.getContext('2d');
    return ctx || undefined;
  };

  // size the backing store once, then restore any existing signature into it
  React.useEffect(() => {
    const canvas = canvasRef.current;
    const ctx = context();
    if (!canvas || !ctx) {
      return;
    }
    const ratio = Math.max(1, Math.min(window.devicePixelRatio || 1, 3));
    canvas.width = WIDTH * ratio;
    canvas.height = HEIGHT * ratio;
    canvas.style.width = WIDTH + 'px';
    canvas.style.height = HEIGHT + 'px';
    ctx.scale(ratio, ratio);
    ctx.lineWidth = 2;
    ctx.lineCap = 'round';
    ctx.lineJoin = 'round';
    ctx.strokeStyle = ink;

    if (props.value) {
      const image = new Image();
      image.onload = () => {
        ctx.drawImage(image, 0, 0, WIDTH, HEIGHT);
      };
      image.src = props.value;
      setHasInk(true);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // keep the stroke color in step with a theme change
  React.useEffect(() => {
    const ctx = context();
    if (ctx) {
      ctx.strokeStyle = ink;
    }
  }, [ink]);

  const pointFrom = (event: React.PointerEvent<HTMLCanvasElement>): IPoint => {
    const canvas = canvasRef.current;
    if (!canvas) {
      return { x: 0, y: 0 };
    }
    const bounds = canvas.getBoundingClientRect();
    return {
      x: ((event.clientX - bounds.left) / bounds.width) * WIDTH,
      y: ((event.clientY - bounds.top) / bounds.height) * HEIGHT
    };
  };

  const start = (event: React.PointerEvent<HTMLCanvasElement>): void => {
    if (props.disabled) {
      return;
    }
    const canvas = canvasRef.current;
    drawing.current = true;
    lastPoint.current = pointFrom(event);
    // keep receiving moves even if the pointer leaves the canvas mid-stroke
    if (canvas && canvas.setPointerCapture) {
      try {
        canvas.setPointerCapture(event.pointerId);
      } catch {
        // capture is best-effort
      }
    }
  };

  const draw = (event: React.PointerEvent<HTMLCanvasElement>): void => {
    if (!drawing.current || props.disabled) {
      return;
    }
    const ctx = context();
    const from = lastPoint.current;
    if (!ctx || !from) {
      return;
    }
    const to = pointFrom(event);
    ctx.beginPath();
    ctx.moveTo(from.x, from.y);
    ctx.lineTo(to.x, to.y);
    ctx.stroke();
    lastPoint.current = to;
    if (!hasInk) {
      setHasInk(true);
    }
  };

  const end = (): void => {
    if (!drawing.current) {
      return;
    }
    drawing.current = false;
    lastPoint.current = undefined;
    const canvas = canvasRef.current;
    if (canvas && hasInk) {
      props.onChange(canvas.toDataURL('image/png'));
    }
  };

  const clear = (): void => {
    const canvas = canvasRef.current;
    const ctx = context();
    if (canvas && ctx) {
      ctx.clearRect(0, 0, WIDTH, HEIGHT);
    }
    setHasInk(false);
    setTypedName('');
    props.onChange('');
  };

  /** Render a typed name onto the canvas so the stored value is still a PNG data URL. */
  const renderTyped = (text: string): void => {
    const canvas = canvasRef.current;
    const ctx = context();
    setTypedName(text);
    if (!canvas || !ctx) {
      return;
    }
    ctx.clearRect(0, 0, WIDTH, HEIGHT);
    const name = text.trim();
    if (name.length === 0) {
      setHasInk(false);
      props.onChange('');
      return;
    }
    ctx.save();
    ctx.fillStyle = ink;
    ctx.textBaseline = 'middle';
    ctx.textAlign = 'center';
    let size = 48;
    const family = '"Segoe Script", "Brush Script MT", cursive, serif';
    ctx.font = 'italic ' + size + 'px ' + family;
    while (size > 14 && ctx.measureText(name).width > WIDTH - 24) {
      size -= 2;
      ctx.font = 'italic ' + size + 'px ' + family;
    }
    ctx.fillText(name, WIDTH / 2, HEIGHT / 2, WIDTH - 16);
    ctx.restore();
    setHasInk(true);
    props.onChange(canvas.toDataURL('image/png'));
  };

  const switchMode = (): void => {
    // the two modes share one canvas, so switching starts fresh
    clear();
    setTyping(!typing);
  };

  if (props.disabled && props.value) {
    return <img className={styles.signatureImage} src={props.value} alt={strings.Form_Signature_Alt} />;
  }

  return (
    <div className={styles.signatureWrap}>
      {typing && (
        <TextField
          label={strings.Form_Signature_TypeLabel}
          value={typedName}
          maxLength={80}
          autoComplete="name"
          ariaLabel={strings.Form_Signature_TypeLabel}
          onChange={(_e, v) => renderTyped(v || '')}
          styles={{ root: { maxWidth: WIDTH, marginBottom: 8 } }}
        />
      )}
      <canvas
        ref={canvasRef}
        className={hasInk ? styles.signaturePad + ' ' + styles.signaturePadFilled : styles.signaturePad}
        role="img"
        aria-label={formatString(hasInk ? strings.Form_Signature_AriaSigned : strings.Form_Signature_AriaEmpty, {
          label: props.ariaLabel || strings.Form_Signature_Alt
        })}
        aria-describedby={props.ariaDescribedBy}
        aria-invalid={props.invalid ? true : undefined}
        onPointerDown={typing ? undefined : start}
        onPointerMove={typing ? undefined : draw}
        onPointerUp={typing ? undefined : end}
        onPointerCancel={typing ? undefined : end}
        onPointerLeave={typing ? undefined : end}
      />
      <div className={styles.signatureActions}>
        <DefaultButton
          iconProps={{ iconName: 'EraseTool' }}
          text={strings.Form_Signature_Clear}
          disabled={!hasInk || props.disabled}
          onClick={clear}
        />
        <DefaultButton
          iconProps={{ iconName: typing ? 'InkingTool' : 'Keyboard' }}
          text={typing ? strings.Form_Signature_DrawInstead : strings.Form_Signature_TypeInstead}
          disabled={props.disabled}
          onClick={switchMode}
        />
        <span className={styles.signatureHint}>
          {hasInk ? strings.Form_Signature_Signed : strings.Form_Signature_Prompt}
        </span>
      </div>
    </div>
  );
};
