import * as React from 'react';
import type { IProduct, IProductPage } from '../../services/products/IProductSource';
import styles from './Products.module.scss';
import type { IProductsProps } from './IProductsProps';

const PAGE_SIZE = 20;
const REQUEST_TIMEOUT_SECONDS = 15;
const REQUEST_TIMEOUT_MS = REQUEST_TIMEOUT_SECONDS * 1_000;
const DEFAULT_TITLE = 'Products';
const LEGACY_DEFAULT_TITLE = 'Товары';
const EMPTY_PRODUCT_PAGE: IProductPage = { products: [], total: 0 };
const priceFormatter = new Intl.NumberFormat('en-US', {
  maximumFractionDigits: 2,
  minimumFractionDigits: 2
});

/** Повертає заголовок каталогу та замінює попереднє значення за замовчуванням. */
function getCatalogTitle(description: string): string {
  return !description || description === LEGACY_DEFAULT_TITLE ? DEFAULT_TITLE : description;
}

/** Створює повідомлення про помилку завантаження без приховування причини від джерела. */
function getLoadErrorMessage(reason: unknown, timedOut: boolean): string {
  if (timedOut) {
    return `The source did not respond within ${REQUEST_TIMEOUT_SECONDS} seconds. Please try again.`;
  }

  const detail = reason instanceof Error ? ` ${reason.message}` : '';
  return `Unable to load products. Check your connection and try again.${detail}`;
}

/** Приховує недоступне зображення, зберігаючи розміри рядка товару. */
function hideBrokenImage(event: React.SyntheticEvent<HTMLImageElement>): void {
  event.currentTarget.style.visibility = 'hidden';
}

interface IProductItemProps {
  product: IProduct;
  sourceId: string;
}

/** Відображає один товар у списку каталогу. */
function ProductItem({ product, sourceId }: IProductItemProps): React.ReactElement {
  return (
    <li className={styles.product}>
      <div className={styles.imageBox}>
        <img
          src={product.thumbnail}
          alt=""
          loading="lazy"
          className={styles.image}
          onError={hideBrokenImage}
        />
      </div>
      <div className={styles.details}>
        <p className={styles.category}>{product.category}</p>
        <h3 className={styles.title}>{product.title}</h3>
        <p className={styles.description}>{product.description}</p>
      </div>
      <div className={styles.price}>
        <span className={styles.priceLabel}>Price</span>
        {priceFormatter.format(product.price)}
      </div>
    </li>
  );
}

/** Відображає каталог товарів і керує завантаженням сторінок джерела. */
export default function Products({ description, source }: IProductsProps): React.ReactElement {
  const [skip, setSkip] = React.useState(0);
  const [revision, setRevision] = React.useState(0);
  const [page, setPage] = React.useState<IProductPage>(EMPTY_PRODUCT_PAGE);
  const [loading, setLoading] = React.useState(true);
  const [error, setError] = React.useState('');
  const title = getCatalogTitle(description);
  const isPreviousPageDisabled = loading || skip === 0;
  const isNextPageDisabled = loading || Boolean(error) || skip + PAGE_SIZE >= page.total;

  /** Завантажує поточну сторінку та скасовує застарілі запити. */
  React.useEffect(() => {
    const controller = new AbortController();
    let isRequestCurrent = true;
    let timedOut = false;

    setLoading(true);
    setError('');

    /** Перериває запит, якщо джерело не відповідає протягом 15 секунд. */
    const abortOnTimeout = (): void => {
      timedOut = true;
      controller.abort();
    };

    const timeoutId = window.setTimeout(abortOnTimeout, REQUEST_TIMEOUT_MS);

    /** Отримує товари та оновлює дані або повідомлення про помилку. */
    const load = async (): Promise<void> => {
      try {
        const result = await source.getProducts(skip, PAGE_SIZE, controller.signal);
        if (isRequestCurrent) {
          setPage(result);
        }
      } catch (reason) {
        if (isRequestCurrent) {
          setError(getLoadErrorMessage(reason, timedOut));
        }
      } finally {
        window.clearTimeout(timeoutId);
        if (isRequestCurrent) {
          setLoading(false);
        }
      }
    };

    // Функція load обробляє власні помилки та не передає їх у React.
    load().catch(() => undefined);

    /** Прибирає таймер і скасовує запит під час зміни сторінки або демонтування. */
    return () => {
      isRequestCurrent = false;
      window.clearTimeout(timeoutId);
      controller.abort();
    };
  }, [source, skip, revision]);

  /** Повторно завантажує поточну сторінку. */
  const retry = (): void => {
    setRevision(value => value + 1);
  };

  /** Переходить на попередню сторінку без від'ємного зміщення. */
  const previousPage = (): void => {
    setSkip(value => Math.max(0, value - PAGE_SIZE));
  };

  /** Переходить на наступну сторінку товарів. */
  const nextPage = (): void => {
    setSkip(value => value + PAGE_SIZE);
  };

  return (
    <section className={styles.section} aria-label="Product list" lang="en">
      <header className={styles.header}>
        <h2 className={styles.heading}>{title}</h2>
        <button type="button" className={styles.button} onClick={retry} disabled={loading}>
          Refresh
        </button>
      </header>

      {loading && (
        <p role="status" className={styles.feedback}>
          Loading products…
        </p>
      )}
      {!loading && !error && page.products.length === 0 &&
        <p className={styles.feedback}>No products found.</p>
      }
      {error && (
        <div role="alert" className={styles.error}>
          <p>{error}</p>
          <button type="button" className={styles.button} onClick={retry}>
            Retry
          </button>
        </div>
      )}
      {!loading && !error && (
        <ul className={styles.list}>
          {page.products.map(product => (
            <ProductItem key={`${source.id}:${product.id}`} product={product} sourceId={source.id} />
          ))}
        </ul>
      )}
      <nav className={styles.pagination} aria-label="Product pages">
        <button
          type="button"
          className={styles.button}
          disabled={isPreviousPageDisabled}
          onClick={previousPage}
        >
          Previous
        </button>
        <button
          type="button"
          className={styles.button}
          disabled={isNextPageDisabled}
          onClick={nextPage}
        >
          Next
        </button>
      </nav>
    </section>
  );
}
