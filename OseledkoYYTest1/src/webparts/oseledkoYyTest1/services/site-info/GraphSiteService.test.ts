import { getSiteInfo, SITE_API_URL } from './GraphSiteService';

const fetchMock = jest.fn();
const originalFetch = globalThis.fetch;
const site = { id: 'tenant,collection,web', name: 'example-site', displayName: 'Example site', webUrl: 'https://example.sharepoint.com/sites/example-site', description: 'Example description' };

/** Готує ізольовану підміну мережевого запиту для кожного тесту. */
beforeEach(() => { fetchMock.mockReset(); globalThis.fetch = fetchMock; });
/** Відновлює мережеву функцію після завершення тестів. */
afterAll(() => { globalThis.fetch = originalFetch; });

/** Перевіряє адресу запиту, заголовок авторизації та повернення відомостей сайту. */
it('sends the token only in the authorization header to the fixed Graph endpoint', async () => {
  fetchMock.mockResolvedValue({ ok: true, json: /** Повертає тестові відомості сайту. */ async () => site });
  const signal = new AbortController().signal;
  await expect(getSiteInfo('  Bearer test-token  ', signal)).resolves.toEqual(site);
  expect(fetchMock).toHaveBeenCalledWith(SITE_API_URL, {
    method: 'GET', headers: { Authorization: 'Bearer test-token', Accept: 'application/json' },
    credentials: 'omit', cache: 'no-store', redirect: 'error', signal
  });
});

/** Перевіряє локальне відхилення порожнього або пошкодженого токена без запиту. */
it.each(['', '   ', 'Bearer ', 'bad\ntoken'])('rejects an empty or malformed token: %j', async token => {
  await expect(getSiteInfo(token, new AbortController().signal)).rejects.toThrow();
  expect(fetchMock).not.toHaveBeenCalled();
});

/** Перевіряє зрозумілі повідомлення для основних помилок Microsoft Graph. */
it.each([[401, 'invalid or expired'], [403, 'Access denied'], [404, 'not found'], [429, 'too many requests'], [500, 'HTTP 500']])
('handles HTTP %s', async (status, message) => {
  fetchMock.mockResolvedValue({ ok: false, status });
  await expect(getSiteInfo('test-token', new AbortController().signal)).rejects.toThrow(String(message));
});

/** Перевіряє відхилення неочікуваної структури успішної відповіді. */
it('rejects malformed site data', async () => {
  fetchMock.mockResolvedValue({ ok: true, json: /** Повертає некоректні тестові дані. */ async () => ({ id: 'only-id' }) });
  await expect(getSiteInfo('test-token', new AbortController().signal)).rejects.toThrow('invalid site information');
});

/** Перевіряє обробку відповіді, яку неможливо прочитати як JSON. */
it('handles invalid JSON', async () => {
  fetchMock.mockResolvedValue({ ok: true, json: /** Імітує помилку читання JSON. */ async () => { throw new SyntaxError('Invalid JSON'); } });
  await expect(getSiteInfo('test-token', new AbortController().signal)).rejects.toThrow('invalid JSON response');
});

/** Перевіряє передавання скасування запиту до компонента. */
it('propagates cancellation', async () => {
  fetchMock.mockRejectedValue(new DOMException('Aborted', 'AbortError'));
  await expect(getSiteInfo('test-token', new AbortController().signal)).rejects.toMatchObject({ name: 'AbortError' });
});
