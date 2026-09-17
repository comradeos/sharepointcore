import * as React from 'react';
import styles from './SiteInfo.module.scss';
import type { ISiteInfo, ISiteInfoSource } from '../../services/site-info/ISiteInfoSource';

const NOT_PROVIDED_TEXT = 'Not provided';
const JSON_INDENTATION = 2;
const DATE_FORMAT_OPTIONS: Intl.DateTimeFormatOptions = {
  timeZone: 'UTC',
  timeZoneName: 'short'
};

interface ISiteDetail {
  label: string;
  value: string;
}

/** Форматує дату сайту англійською мовою із зазначенням часового поясу. */
function formatSiteDate(value: string | undefined): string {
  if (!value) {
    return NOT_PROVIDED_TEXT;
  }

  const date = new Date(value);

  return Number.isNaN(date.getTime())
    ? NOT_PROVIDED_TEXT
    : date.toLocaleString('en-US', DATE_FORMAT_OPTIONS);
}

interface ISiteInfoProps {
  source: ISiteInfoSource;
}

/** Перетворює відповідь SharePoint API на рядки для відображення у секції. */
function getSiteDetails(site: ISiteInfo): ISiteDetail[] {
  return [
    { label: 'Display name', value: site.displayName },
    { label: 'Name', value: site.name },
    { label: 'Description', value: site.description || NOT_PROVIDED_TEXT },
    { label: 'Site ID', value: site.id },
    { label: 'Web URL', value: site.webUrl },
    { label: 'Hostname', value: site.hostname || NOT_PROVIDED_TEXT },
    { label: 'Created', value: formatSiteDate(site.createdDateTime) },
    { label: 'Last modified', value: formatSiteDate(site.lastModifiedDateTime) }
  ];
}

/** Відображає відомості про поточний SharePoint-сайт, отримані через PnPjs. */
export default function SiteInfo({ source }: ISiteInfoProps): React.ReactElement {
  const [site, setSite] = React.useState<ISiteInfo>();
  const [loading, setLoading] = React.useState(false);
  const [error, setError] = React.useState('');

  /** Завантажує відомості сайту після монтування компонента або зміни джерела. */
  React.useEffect(() => {
    let isCurrent = true;

    /** Оновлює стан лише якщо компонент ще відображається на сторінці. */
    const loadSiteInfo = async (): Promise<void> => {
      setLoading(true);
      setError('');

      try {
        const result = await source.getSiteInfo();

        if (isCurrent) {
          setSite(result);
        }
      } catch {
        if (isCurrent) {
          setSite(undefined);
          setError('Unable to load SharePoint site information. Check your permissions and try again.');
        }
      } finally {
        if (isCurrent) {
          setLoading(false);
        }
      }
    };

    loadSiteInfo().catch(() => undefined);

    /** Позначає результат запиту неактуальним після демонтування компонента. */
    return () => {
      isCurrent = false;
    };
  }, [source]);

  return (
    <section className={styles.section} aria-label="SharePoint site information" lang="en">
      <h2 className={styles.heading}>SharePoint site</h2>

      {loading && (
        <p className={styles.feedback} role="status">
          Loading site information…
        </p>
      )}

      {error && (
        <p className={styles.error} role="alert">
          {error}
        </p>
      )}

      {site && (
        <div className={styles.siteResult}>
          <p className={styles.feedback} role="status">
            Site information loaded.
          </p>
          <a className={styles.siteLink} href={site.webUrl} target="_blank" rel="noreferrer">
            {site.webUrl}
          </a>
          <dl className={styles.siteDetails}>
            {getSiteDetails(site).map(detail => (
              <React.Fragment key={detail.label}>
                <dt>{detail.label}</dt>
                <dd>{detail.value}</dd>
              </React.Fragment>
            ))}
          </dl>
          <details className={styles.siteJson}>
            <summary>JSON response</summary>
            <pre>{JSON.stringify(site, undefined, JSON_INDENTATION)}</pre>
          </details>
        </div>
      )}
    </section>
  );
}
