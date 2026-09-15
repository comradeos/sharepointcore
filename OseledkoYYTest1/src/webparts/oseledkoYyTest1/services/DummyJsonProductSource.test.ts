import { DummyJsonProductSource } from './DummyJsonProductSource';

const product = { id: 1, title: 'R&D', description: 'Example', category: 'beauty', price: 9.99, thumbnail: 'https://example.com/image.png' };
const fetchMock = jest.fn();
const originalFetch = globalThis.fetch;

/** Готує ізольовану підміну мережевого запиту перед кожним тестом. */
beforeEach(() => {
  fetchMock.mockReset();
  globalThis.fetch = fetchMock;
});
/** Відновлює початкову функцію fetch після завершення тестів. */
afterAll(() => { globalThis.fetch = originalFetch; });

/** Перевіряє пагінацію, передавання сигналу скасування та відсутність облікових даних. */
it('loads the requested page and forwards cancellation without credentials', async () => {
  fetchMock.mockResolvedValue({ ok: true, json: /** Повертає підготовлені дані відповіді для тесту. */ async () => ({ products: [product], total: 21 }) });
  const signal = new AbortController().signal;
  const result = await new DummyJsonProductSource().getProducts(20, 20, signal);
  expect(result).toEqual({ products: [product], total: 21 });
  expect(fetchMock).toHaveBeenCalledWith(
    'https://dummyjson.com/products?limit=20&skip=20&select=title,description,category,price,thumbnail',
    { signal, credentials: 'omit' }
  );
});

/** Перевіряє обробку порожнього каталогу. */
it('accepts an empty catalog', async () => {
  fetchMock.mockResolvedValue({ ok: true, json: /** Повертає підготовлені дані відповіді для тесту. */ async () => ({ products: [], total: 0 }) });
  await expect(new DummyJsonProductSource().getProducts(0, 20, new AbortController().signal))
    .resolves.toEqual({ products: [], total: 0 });
});

/** Перевіряє повідомлення про помилку HTTP. */
it('rejects failed HTTP responses', async () => {
  fetchMock.mockResolvedValue({ ok: false, status: 503 });
  await expect(new DummyJsonProductSource().getProducts(0, 20, new AbortController().signal))
    .rejects.toThrow('HTTP error 503');
});

/** Перевіряє відхилення відповідей із некоректною структурою даних. */
it.each([
  null,
  { products: {}, total: 1 },
  { products: [product], total: -1 },
  { products: [{ ...product, price: '9.99' }], total: 1 },
  { products: [{ ...product, title: null }], total: 1 }
])('rejects malformed API data: %j', async data => {
  fetchMock.mockResolvedValue({ ok: true, json: /** Повертає підготовлені дані відповіді для тесту. */ async () => data });
  await expect(new DummyJsonProductSource().getProducts(0, 20, new AbortController().signal))
    .rejects.toThrow('invalid');
});

/** Перевіряє передавання мережевої помилки виклику компонента. */
it('propagates a network failure', async () => {
  fetchMock.mockRejectedValue(new TypeError('Failed to fetch'));
  await expect(new DummyJsonProductSource().getProducts(0, 20, new AbortController().signal))
    .rejects.toThrow('Failed to fetch');
});

/** Перевіряє передавання помилки скасованого запиту. */
it('propagates cancellation', async () => {
  const controller = new AbortController();
  controller.abort();
  fetchMock.mockRejectedValue(new DOMException('Aborted', 'AbortError'));
  await expect(new DummyJsonProductSource().getProducts(0, 20, controller.signal))
    .rejects.toMatchObject({ name: 'AbortError' });
});
