import * as React from 'react';
import styles from './FormRenderer.module.scss';
import { ContentStyle } from '../../models';
import { sanitizeHtml } from '../../utils/sanitizeHtml';

export interface IContentBlockProps {
  html?: string;
  imageUrl?: string;
  style?: ContentStyle;
}

const STYLE_CLASS: { [key: string]: string } = {
  text: styles.contentPlain,
  info: styles.contentInfo,
  success: styles.contentSuccess,
  warning: styles.contentWarning,
  divider: styles.contentDivider
};

/**
 * Static instructions, an image, or a divider between questions. Collects no
 * answer and provisions no column — previously the only way to add guidance was
 * to invent a section just to use its description.
 */
export const ContentBlock: React.FunctionComponent<IContentBlockProps> = (props) => {
  const style = props.style || 'info';
  const html = React.useMemo(() => sanitizeHtml(props.html || ''), [props.html]);

  if (style === 'divider') {
    return <hr className={styles.contentDivider} />;
  }

  const className = styles.contentBlock + ' ' + (STYLE_CLASS[style] || styles.contentInfo);

  return (
    <div className={className}>
      {props.imageUrl && (
        <img className={styles.contentImage} src={props.imageUrl} alt="" loading="lazy" />
      )}
      {html && (
        // sanitized above; authored by the form owner, read-only for respondents
        <div dangerouslySetInnerHTML={{ __html: html }} />
      )}
    </div>
  );
};
