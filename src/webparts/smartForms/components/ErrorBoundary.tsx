import * as React from 'react';
import * as strings from 'SmartFormsWebPartStrings';
import { logError } from '../utils/debug';

export interface IErrorBoundaryProps {
  children: React.ReactNode;
}

interface IErrorBoundaryState {
  error: Error | undefined;
}

/**
 * Catches a render-time exception anywhere below it and shows a readable panel
 * instead of leaving the web part blank.
 *
 * Matches the pattern used in SharePointSmartPermissionsWebPart's App.tsx: the
 * panel is styled with inline styles rather than a CSS module, since if the
 * app is broken enough to hit this boundary, its own stylesheet may not have
 * loaded either. `logError` puts the real stack in the console — the previous
 * failure mode here was a web part that rendered nothing with no clue why.
 */
export class ErrorBoundary extends React.Component<IErrorBoundaryProps, IErrorBoundaryState> {
  constructor(props: IErrorBoundaryProps) {
    super(props);
    this.state = { error: undefined };
  }

  public static getDerivedStateFromError(error: Error): IErrorBoundaryState {
    return { error };
  }

  public componentDidCatch(error: Error, info: React.ErrorInfo): void {
    logError('Render', error);
    // eslint-disable-next-line no-console
    console.error('[SmartForms] Component stack:', info.componentStack);
  }

  public render(): React.ReactNode {
    const { error } = this.state;
    if (!error) {
      return this.props.children;
    }
    return (
      <div
        style={{
          padding: 16,
          fontFamily: 'Consolas, monospace',
          fontSize: 13,
          background: '#fff3f3',
          border: '1px solid #c00',
          borderRadius: 4,
          margin: 8
        }}
      >
        <strong style={{ color: '#c00', fontSize: 14 }}>{strings.App_ErrorBoundary_Title}</strong>
        <br />
        <br />
        <strong>{strings.App_ErrorBoundary_Message}</strong> {error.message || String(error)}
        <br />
        <br />
        <strong>{strings.App_ErrorBoundary_Stack}</strong>
        <pre
          style={{
            fontSize: 11,
            whiteSpace: 'pre-wrap',
            overflowWrap: 'anywhere',
            background: '#f5f5f5',
            padding: 8,
            margin: '4px 0',
            borderRadius: 2
          }}
        >
          {error.stack || strings.App_ErrorBoundary_NoStack}
        </pre>
        <span style={{ fontSize: 11, color: '#666' }}>
          {strings.App_ErrorBoundary_ConsoleNote}
        </span>
      </div>
    );
  }
}
