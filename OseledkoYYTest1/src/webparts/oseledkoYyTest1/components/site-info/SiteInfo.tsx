import * as React from 'react';
import styles from './SiteInfo.module.scss';
import { getSiteInfo, ISiteInfo, SITE_WEB_URL, SiteInfoError } from '../../services/site-info/GraphSiteService';

const REQUEST_TIMEOUT_SECONDS = 15;
const REQUEST_TIMEOUT_MS = REQUEST_TIMEOUT_SECONDS * 1_000;
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

/** Перетворює відповідь Microsoft Graph на рядки для відображення у формі. */
function getSiteDetails(site: ISiteInfo): ISiteDetail[] {
  return [
    { label: 'Display name', value: site.displayName },
    { label: 'Name', value: site.name },
    { label: 'Description', value: site.description || NOT_PROVIDED_TEXT },
    { label: 'Site ID', value: site.id },
    { label: 'Web URL', value: site.webUrl },
    { label: 'Hostname', value: site.siteCollection?.hostname || NOT_PROVIDED_TEXT },
    { label: 'Created', value: formatSiteDate(site.createdDateTime) },
    { label: 'Last modified', value: formatSiteDate(site.lastModifiedDateTime) }
  ];
}

/** Створює безпечне повідомлення про помилку виконання запиту до Microsoft Graph. */
function getRequestErrorMessage(reason: unknown, timedOut: boolean): string {
  if (timedOut) {
    return `The request timed out after ${REQUEST_TIMEOUT_SECONDS} seconds. Please try again.`;
  }

  if (reason instanceof SiteInfoError) {
    return reason.message;
  }

  return 'Unable to connect to Microsoft Graph. Check your connection and try again.';
}

/** Відображає форму токена та отримані з Microsoft Graph відомості про сайт. */
export default function SiteInfo(): React.ReactElement {
  const [token, setToken] = React.useState('');
  const [site, setSite] = React.useState<ISiteInfo>();
  const [loading, setLoading] = React.useState(false);
  const [error, setError] = React.useState('');
  const requestControllerRef = React.useRef<AbortController>();
  const timeoutIdRef = React.useRef<number>();

  /** Реєструє очищення запиту та таймера під час демонтування форми. */
  React.useEffect(() => {
    /** Скасовує активний запит і звільняє таймер після видалення компонента. */
    return () => {
      requestControllerRef.current?.abort();
      requestControllerRef.current = undefined;
      window.clearTimeout(timeoutIdRef.current);
    };
  }, []);

  /** Зберігає введений токен лише в оперативному стані компонента. */
  const changeToken = (event: React.ChangeEvent<HTMLInputElement>): void => {
    setToken(event.target.value);
    setError('');
  };

  /** Виконує запит за натисканням кнопки та оновлює результат або помилку. */
  const submit = async (event: React.FormEvent<HTMLFormElement>): Promise<void> => {
    event.preventDefault();

    requestControllerRef.current?.abort();
    window.clearTimeout(timeoutIdRef.current);

    const requestController = new AbortController();
    requestControllerRef.current = requestController;
    let timedOut = false;

    setError('');
    setSite(undefined);
    setLoading(true);

    /** Зупиняє запит після перевищення часу очікування відповіді. */
    const abortOnTimeout = (): void => {
      timedOut = true;
      requestController.abort();
    };

    timeoutIdRef.current = window.setTimeout(abortOnTimeout, REQUEST_TIMEOUT_MS);

    try {
      const result = await getSiteInfo(token, requestController.signal);

      if (requestControllerRef.current === requestController) {
        setSite(result);
      }
    } catch (reason) {
      if (requestControllerRef.current === requestController) {
        setError(getRequestErrorMessage(reason, timedOut));
      }
    } finally {
      if (requestControllerRef.current === requestController) {
        window.clearTimeout(timeoutIdRef.current);
        requestControllerRef.current = undefined;
        setLoading(false);
      }
    }
  };

  return (
    <section className={styles.section} aria-label="SharePoint site information" lang="en">
      <h2 className={styles.heading}>SharePoint site</h2>

      <a className={styles.siteLink} href={SITE_WEB_URL} target="_blank" rel="noreferrer">
        {SITE_WEB_URL}
      </a>

      <form className={styles.siteForm} onSubmit={submit} noValidate>
        <label className={styles.tokenLabel}>
          Access token
          <input
            className={styles.tokenInput}
            type="password"
            value={token}
            onChange={changeToken}
            placeholder="Paste your Microsoft Graph access token"
            autoComplete="off"
            spellCheck={false}
            autoCapitalize="none"
            disabled={loading}
            aria-required="true"
          />
        </label>

        <button className={styles.button} type="submit" disabled={loading}>
          Get site info
        </button>
      </form>

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
