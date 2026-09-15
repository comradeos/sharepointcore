import * as React from 'react';
import SiteInfo from './SiteInfo';
import styles from './OseledkoYyTest1.module.scss';
import type { IOseledkoYyTest1Props } from './IOseledkoYyTest1Props';
import type { IProduct, IProductPage } from '../services/IProductSource';

const PAGE_SIZE = 20;
const priceFormat = new Intl.NumberFormat('en-US', { maximumFractionDigits: 2, minimumFractionDigits: 2 });

/** Відображає каталог товарів і керує завантаженням сторінок джерела. */
export default function OseledkoYyTest1({ description, source }: IOseledkoYyTest1Props): React.ReactElement {
  const [skip, setSkip] = React.useState(0);
  const [revision, setRevision] = React.useState(0);
  const [page, setPage] = React.useState<IProductPage>({ products: [], total: 0 });
  const [loading, setLoading] = React.useState(true);
  const [error, setError] = React.useState('');
  // Замінює попередній стандартний заголовок у вже збережених вебчастинах.
  const title = !description || description === 'Товары' ? 'Products' : description;

  /** Завантажує поточну сторінку та скасовує застарілі запити. */
  React.useEffect(() => {
    const controller = new AbortController();
    let active = true;
    let timedOut = false;
    setLoading(true);
    setError('');

    /** Перериває запит, якщо джерело не відповідає протягом 15 секунд. */
    const abortOnTimeout = (): void => {
      timedOut = true;
      controller.abort();
    };
    const timeout = window.setTimeout(abortOnTimeout, 15000);

    /** Отримує товари та оновлює дані або повідомлення про помилку. */
    const load = async (): Promise<void> => {
      try {
        const result = await source.getProducts(skip, PAGE_SIZE, controller.signal);
        if (active) setPage(result);
      } catch (reason) {
        if (active) {
          const detail = reason instanceof Error ? reason.message : '';
          setError(timedOut
            ? 'The source did not respond within 15 seconds. Please try again.'
            : `Unable to load products. Check your connection and try again. ${detail}`);
        }
      } finally {
        window.clearTimeout(timeout);
        if (active) setLoading(false);
      }
    };
    // Обробляє завершення промісу; помилки запиту вже відображає функція load.
    load().catch(() => undefined);

    /** Прибирає таймер і скасовує запит під час зміни сторінки або демонтування. */
    return () => {
      active = false;
      window.clearTimeout(timeout);
      controller.abort();
    };
  }, [source, skip, revision]);

  /** Повторно завантажує поточну сторінку. */
  const retry = (): void => {
    // Збільшує лічильник повторних запитів на основі актуального стану.
    setRevision(value => value + 1);
  };

  /** Переходить на попередню сторінку без від'ємного зміщення. */
  const previousPage = (): void => {
    // Обчислює зміщення попередньої сторінки на основі актуального стану.
    setSkip(value => Math.max(0, value - PAGE_SIZE));
  };

  /** Переходить на наступну сторінку товарів. */
  const nextPage = (): void => {
    // Обчислює зміщення наступної сторінки на основі актуального стану.
    setSkip(value => value + PAGE_SIZE);
  };

  /** Приховує недоступне зображення, зберігаючи розміри рядка. */
  const hideBrokenImage = (event: React.SyntheticEvent<HTMLImageElement>): void => {
    event.currentTarget.style.visibility = 'hidden';
  };

  /** Формує рядок списку з даними одного товару. */
  const renderProduct = (product: IProduct): React.ReactElement => (
    <li key={`${source.id}:${product.id}`} className={styles.product}>
      <div className={styles.imageBox}>
        <img src={product.thumbnail} alt="" loading="lazy" className={styles.image} onError={hideBrokenImage} />
      </div>
      <div className={styles.details}>
        <p className={styles.category}>{product.category}</p>
        <h3 className={styles.title}>{product.title}</h3>
        <p className={styles.description}>{product.description}</p>
      </div>
      <div className={styles.price}>
        <span className={styles.priceLabel}>Price</span>
        {priceFormat.format(product.price)}
      </div>
    </li>
  );

  return (
    <>
      <section className={styles.oseledkoYyTest1} aria-label="Product list" lang="en">
        <header className={styles.header}>
          <h2 className={styles.heading}>{title}</h2>
          <button type="button" className={styles.button} onClick={retry} disabled={loading}>Refresh</button>
        </header>

        {loading && <p role="status" className={styles.feedback}>Loading products…</p>}
        {!loading && !error && page.products.length === 0 &&
          <p className={styles.feedback}>No products found.</p>}
        {error && <div role="alert" className={styles.error}>
          <p>{error}</p>
          <button type="button" className={styles.button} onClick={retry}>Retry</button>
        </div>}
        {!loading && !error && <ul className={styles.list}>{page.products.map(renderProduct)}</ul>}
        <nav className={styles.pagination} aria-label="Product pages">
          <button type="button" className={styles.button} disabled={loading || skip === 0}
            onClick={previousPage}>Previous</button>
          <button type="button" className={styles.button}
            disabled={loading || !!error || skip + PAGE_SIZE >= page.total}
            onClick={nextPage}>Next</button>
        </nav>
      </section>
      <section className={`${styles.oseledkoYyTest1} ${styles.siteSection}`} aria-label="SharePoint site information" lang="en">
        <SiteInfo />
      </section>
    </>
  );
}
