import * as React from 'react';
import styles from './OseledkoYyTest1.module.scss';
import { getSiteInfo, ISiteInfo, SITE_WEB_URL, SiteInfoError } from '../services/GraphSiteService';

/** Форматує дату сайту англійською мовою із зазначенням часового поясу. */
function formatSiteDate(value: string | undefined): string {
  if (!value) return 'Not provided';
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? 'Not provided' : date.toLocaleString('en-US', { timeZone: 'UTC', timeZoneName: 'short' });
}

/** Відображає форму токена та отримані з Microsoft Graph відомості про сайт. */
export default function SiteInfo(): React.ReactElement {
  const [token, setToken] = React.useState('');
  const [site, setSite] = React.useState<ISiteInfo>();
  const [loading, setLoading] = React.useState(false);
  const [error, setError] = React.useState('');
  const request = React.useRef<AbortController>();
  const timer = React.useRef<number>();

  /** Реєструє очищення запиту та таймера під час демонтування форми. */
  React.useEffect(() => {
    /** Скасовує активний запит і звільняє таймер після видалення компонента. */
    return () => {
      request.current?.abort();
      request.current = undefined;
      window.clearTimeout(timer.current);
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
    request.current?.abort();
    window.clearTimeout(timer.current);
    const controller = new AbortController();
    request.current = controller;
    let timedOut = false;
    setError('');
    setSite(undefined);
    setLoading(true);

    /** Зупиняє запит після перевищення часу очікування відповіді. */
    timer.current = window.setTimeout(() => {
      timedOut = true;
      controller.abort();
    }, 15000);

    try {
      const result = await getSiteInfo(token, controller.signal);
      if (request.current === controller) setSite(result);
    } catch (reason) {
      if (request.current === controller) {
        setError(timedOut ? 'The request timed out after 15 seconds. Please try again.' :
          reason instanceof SiteInfoError ? reason.message :
            'Unable to connect to Microsoft Graph. Check your connection and try again.');
      }
    } finally {
      if (request.current === controller) {
        window.clearTimeout(timer.current);
        request.current = undefined;
        setLoading(false);
      }
    }
  };

  return (
    <>
      <h2 className={styles.heading}>SharePoint site</h2>
      <a className={styles.siteLink} href={SITE_WEB_URL} target="_blank" rel="noreferrer">{SITE_WEB_URL}</a>
      <form className={styles.siteForm} onSubmit={submit} noValidate>
        <label className={styles.tokenLabel}>
          Access token
          <input className={styles.tokenInput} type="password" value={token} onChange={changeToken}
            placeholder="Paste your Microsoft Graph access token" autoComplete="off" spellCheck={false}
            autoCapitalize="none" disabled={loading} aria-required="true" />
        </label>
        <button className={styles.button} type="submit" disabled={loading}>Get site info</button>
      </form>
      {loading && <p className={styles.feedback} role="status">Loading site information…</p>}
      {error && <p className={styles.error} role="alert">{error}</p>}
      {site && <div className={styles.siteResult}>
        <p className={styles.feedback} role="status">Site information loaded.</p>
        <dl className={styles.siteDetails}>
          <dt>Display name</dt><dd>{site.displayName}</dd>
          <dt>Name</dt><dd>{site.name}</dd>
          <dt>Description</dt><dd>{site.description || 'Not provided'}</dd>
          <dt>Site ID</dt><dd>{site.id}</dd>
          <dt>Web URL</dt><dd>{site.webUrl}</dd>
          <dt>Hostname</dt><dd>{site.siteCollection?.hostname || 'Not provided'}</dd>
          <dt>Created</dt><dd>{formatSiteDate(site.createdDateTime)}</dd>
          <dt>Last modified</dt><dd>{formatSiteDate(site.lastModifiedDateTime)}</dd>
        </dl>
        <details className={styles.siteJson}>
          <summary>JSON response</summary>
          <pre>{JSON.stringify(site, undefined, 2)}</pre>
        </details>
      </div>}
    </>
  );
}
