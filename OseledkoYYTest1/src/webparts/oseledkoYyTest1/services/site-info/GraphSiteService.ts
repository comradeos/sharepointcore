export const SITE_API_URL = 'https://graph.microsoft.com/v1.0/sites/uaenergy0.sharepoint.com:/sites/oseledko-yy-test';
export const SITE_WEB_URL = 'https://uaenergy0.sharepoint.com/sites/oseledko-yy-test';

export interface ISiteInfo {
  id: string;
  name: string;
  displayName: string;
  webUrl: string;
  description?: string;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  siteCollection?: { hostname?: string };
}

/** Позначає помилку, повідомлення якої можна безпечно показати у формі. */
export class SiteInfoError extends Error {
  /** Створює помилку з англомовним повідомленням без вмісту токена. */
  public constructor(message: string) {
    super(message);
    this.name = 'SiteInfoError';
    Object.setPrototypeOf(this, SiteInfoError.prototype);
  }
}

/** Перевіряє структуру отриманих даних сайту перед відображенням. */
function isSiteInfo(value: unknown): value is ISiteInfo {
  if (!value || typeof value !== 'object') return false;
  const site = value as Partial<ISiteInfo>;
  return typeof site.id === 'string' && typeof site.name === 'string' &&
    typeof site.displayName === 'string' && typeof site.webUrl === 'string' &&
    (site.description === undefined || typeof site.description === 'string') &&
    (site.createdDateTime === undefined || typeof site.createdDateTime === 'string') &&
    (site.lastModifiedDateTime === undefined || typeof site.lastModifiedDateTime === 'string') &&
    (site.siteCollection === undefined || (!!site.siteCollection && typeof site.siteCollection === 'object' &&
      (site.siteCollection.hostname === undefined || typeof site.siteCollection.hostname === 'string')));
}

/** Отримує відомості про заданий сайт, передаючи токен лише в заголовку Microsoft Graph. */
export async function getSiteInfo(token: string, signal: AbortSignal): Promise<ISiteInfo> {
  const accessToken = token.trim().replace(/^Bearer(?:\s+|$)/i, '');
  if (!accessToken) throw new SiteInfoError('Enter an access token.');
  if (/\s/.test(accessToken)) throw new SiteInfoError('The access token must not contain spaces or line breaks.');

  const response = await fetch(SITE_API_URL, {
    method: 'GET',
    headers: { Authorization: `Bearer ${accessToken}`, Accept: 'application/json' },
    credentials: 'omit',
    cache: 'no-store',
    redirect: 'error',
    signal
  });
  if (!response.ok) {
    // Використовує власні повідомлення, щоб не відображати службові дані відповіді.
    const messages: { [status: number]: string } = {
      401: 'The access token is invalid or expired. Paste a valid Microsoft Graph token and try again.',
      403: 'Access denied. Check that the token has permission to read this SharePoint site.',
      404: 'The SharePoint site was not found.',
      429: 'Microsoft Graph is receiving too many requests. Please try again later.'
    };
    throw new SiteInfoError(messages[response.status] || `Microsoft Graph returned HTTP ${response.status}. Please try again.`);
  }
  let data: unknown;
  try {
    data = await response.json();
  } catch {
    throw new SiteInfoError('Microsoft Graph returned an invalid JSON response.');
  }
  if (!isSiteInfo(data)) throw new SiteInfoError('Microsoft Graph returned invalid site information.');
  return data;
}
