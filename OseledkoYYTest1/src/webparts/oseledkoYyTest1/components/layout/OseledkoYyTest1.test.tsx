/** Замінює стилі контейнера у тестовому DOM. */
jest.mock('./OseledkoYyTest1.module.scss', () => ({}), { virtual: true });
/** Замінює незалежні стилі каталогу у тестовому DOM. */
jest.mock('../products/Products.module.scss', () => ({}), { virtual: true });
/** Замінює незалежні стилі форми сайту у тестовому DOM. */
jest.mock('../site-info/SiteInfo.module.scss', () => ({}), { virtual: true });

import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { act, Simulate } from 'react-dom/test-utils';
import OseledkoYyTest1 from './OseledkoYyTest1';

const PRODUCT_LIST_SELECTOR = '[aria-label="Product list"]';
const SITE_INFO_SELECTOR = '[aria-label="SharePoint site information"]';
const DRAFT_ACCESS_TOKEN = 'draft-test-token';
const NEXT_PAGE_SKIP = 20;

const firstProduct = {
  id: 1,
  title: 'First product',
  description: 'Example',
  category: 'test',
  price: 1,
  thumbnail: ''
};

/** Повертає обов’язковий елемент тестового DOM або завершує тест зрозумілою помилкою. */
function getRequiredElement<T extends Element>(container: ParentNode, selector: string): T {
  const element = container.querySelector(selector);

  if (!element) {
    throw new Error(`The element "${selector}" was not found.`);
  }

  return element as T;
}

/** Перевіряє, що пагінація та помилка товарів не змінюють стан незалежної форми сайту. */
it('preserves the site token when product pagination fails and cancels the old product request', async () => {
  // Arrange: перша сторінка завантажується успішно, а наступна повертає помилку.
  const getProducts = jest.fn()
    .mockResolvedValueOnce({ products: [firstProduct], total: 40 })
    .mockRejectedValueOnce(new Error('Product source unavailable'));

  const source = { id: 'test', name: 'Test source', getProducts };
  const container = document.createElement('div');
  document.body.appendChild(container);

  try {
    // Act: монтує контейнер і очікує початкового завантаження каталогу.
    await act(async () => {
      ReactDOM.render(<OseledkoYyTest1 description="Products" source={source} />, container);
    });

    const productSection = getRequiredElement<HTMLElement>(container, PRODUCT_LIST_SELECTOR);
    const siteSection = getRequiredElement<HTMLElement>(container, SITE_INFO_SELECTOR);

    // Assert: обидві секції відображаються в одному контейнері.
    expect(productSection.parentElement).toBe(siteSection.parentElement);
    expect(productSection.textContent).toContain('First product');

    const tokenInput = getRequiredElement<HTMLInputElement>(siteSection, 'input');

    // Act: вводить токен у незалежну форму відомостей сайту.
    act(() => {
      tokenInput.value = DRAFT_ACCESS_TOKEN;
      Simulate.change(tokenInput);
    });

    const firstSignal = getProducts.mock.calls[0][2] as AbortSignal;
    const nextButton = productSection.querySelectorAll('nav button')[1];

    if (!nextButton) {
      throw new Error('The next page button was not found.');
    }

    // Act: переходить на наступну сторінку з імітованою помилкою джерела товарів.
    await act(async () => {
      Simulate.click(nextButton);
    });

    // Assert: каталог обробив помилку, а форма сайту зберегла власний стан.
    expect(getProducts.mock.calls[1][0]).toBe(NEXT_PAGE_SKIP);
    expect(firstSignal.aborted).toBe(true);
    expect(getRequiredElement<HTMLElement>(productSection, '[role="alert"]').textContent)
      .toContain('Product source unavailable');
    expect(tokenInput.value).toBe(DRAFT_ACCESS_TOKEN);
    expect(getRequiredElement<HTMLButtonElement>(siteSection, 'button').disabled).toBe(false);
    expect(siteSection.querySelector('[role="alert"]')).toBeNull();
  } finally {
    // Очищує DOM і запускає скасування активних запитів компонентів.
    act(() => {
      ReactDOM.unmountComponentAtNode(container);
    });
    container.remove();
  }
});
