import type { IProduct, IProductPage, IProductSource } from './IProductSource';

/** Перевіряє обов’язкові поля товару у відповіді зовнішнього API. */
function isProduct(value: unknown): value is IProduct {
  if (!value || typeof value !== 'object') return false;
  const product = value as Partial<IProduct>;
  return typeof product.id === 'number' && Number.isInteger(product.id) &&
    typeof product.title === 'string' && typeof product.description === 'string' &&
    typeof product.category === 'string' && typeof product.thumbnail === 'string' &&
    typeof product.price === 'number' && Number.isFinite(product.price) && product.price >= 0;
}

/** Завантажує сторінки каталогу DummyJSON у спільному форматі товарів. */
export class DummyJsonProductSource implements IProductSource {
  public readonly id: string = 'dummyjson';
  public readonly name: string = 'DummyJSON';

  /** Отримує сторінку товарів, перевіряючи HTTP-статус і структуру даних. */
  public async getProducts(skip: number, limit: number, signal: AbortSignal): Promise<IProductPage> {
    const url = `https://dummyjson.com/products?limit=${limit}&skip=${skip}&select=title,description,category,price,thumbnail`;
    const response = await fetch(url, { signal, credentials: 'omit' });
    if (!response.ok) {
      throw new Error(`The source returned HTTP error ${response.status}.`);
    }
    const data: unknown = await response.json();
    if (!data || typeof data !== 'object') {
      throw new Error('The source returned an invalid response.');
    }
    const page = data as Partial<IProductPage>;
    if (!Array.isArray(page.products) || !page.products.every(isProduct) ||
        typeof page.total !== 'number' || !Number.isInteger(page.total) || page.total < 0) {
      throw new Error('The source returned invalid product data.');
    }
    return { products: page.products, total: page.total };
  }
}
