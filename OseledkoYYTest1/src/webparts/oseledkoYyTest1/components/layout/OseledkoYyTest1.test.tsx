/** Замінює стилі контейнера у тестовому DOM. */
jest.mock('./OseledkoYyTest1.module.scss', () => ({}), { virtual: true });
/** Замінює незалежні стилі каталогу у тестовому DOM. */
jest.mock('../products/Products.module.scss', () => ({}), { virtual: true });
/** Замінює незалежні стилі форми сайту у тестовому DOM. */
jest.mock('../site-info/SiteInfo.module.scss', () => ({}), { virtual: true });
/** Замінює незалежні стилі списків сайту у тестовому DOM. */
jest.mock('../site-lists/SiteLists.module.scss', () => ({}), { virtual: true });
/** Замінює незалежні стилі пошуку Users у тестовому DOM. */
jest.mock('../user-lookup/UserLookup.module.scss', () => ({}), { virtual: true });
/** Замінює незалежні стилі створення Users у тестовому DOM. */
jest.mock('../user-create/UserCreate.module.scss', () => ({}), { virtual: true });

import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { act, Simulate } from 'react-dom/test-utils';
import OseledkoYyTest1 from './OseledkoYyTest1';

const PRODUCT_LIST_SELECTOR = '[aria-label="Product list"]';
const SITE_INFO_SELECTOR = '[aria-label="SharePoint site information"]';
const SITE_LISTS_SELECTOR = '[aria-label="SharePoint lists"]';
const USER_LOOKUP_SELECTOR = '[aria-label="User lookup"]';
const USER_CREATE_SELECTOR = '[aria-label="Create user"]';
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

/** Перевіряє, що помилка товарів не очищує незалежно завантажені відомості про сайт. */
it('keeps site information when product pagination fails and cancels the old product request', async () => {
  // Arrange: перша сторінка завантажується успішно, а наступна повертає помилку.
  const getProducts = jest.fn()
    .mockResolvedValueOnce({ products: [firstProduct], total: 40 })
    .mockRejectedValueOnce(new Error('Product source unavailable'));

  const source = { id: 'test', name: 'Test source', getProducts };
  const getSiteInfo = jest.fn().mockResolvedValue({
    id: 'site-id',
    name: 'Test site',
    displayName: 'Test site',
    webUrl: 'https://example.sharepoint.com/sites/test'
  });
  const siteInfoSource = { getSiteInfo };
  const getSiteLists = jest.fn().mockResolvedValue([]);
  const siteListSource = { getSiteLists };
  const getUserById = jest.fn();
  const userLookupSource = { getUserById };
  const createUser = jest.fn();
  const userCreateSource = { createUser };
  const container = document.createElement('div');
  document.body.appendChild(container);

  try {
    // Act: монтує контейнер і очікує початкового завантаження каталогу.
    await act(async () => {
      ReactDOM.render(
        <OseledkoYyTest1
          description="Products"
          source={source}
          siteInfoSource={siteInfoSource}
          siteListSource={siteListSource}
          userLookupSource={userLookupSource}
          userCreateSource={userCreateSource}
        />,
        container
      );
      await Promise.resolve();
    });

    const productSection = getRequiredElement<HTMLElement>(container, PRODUCT_LIST_SELECTOR);
    const siteSection = getRequiredElement<HTMLElement>(container, SITE_INFO_SELECTOR);
    const siteListsSection = getRequiredElement<HTMLElement>(container, SITE_LISTS_SELECTOR);
    const userLookupSection = getRequiredElement<HTMLElement>(container, USER_LOOKUP_SELECTOR);
    const userCreateSection = getRequiredElement<HTMLElement>(container, USER_CREATE_SELECTOR);

    // Assert: обидві секції відображаються в одному контейнері.
    expect(productSection.parentElement).toBe(siteSection.parentElement);
    expect(siteListsSection.parentElement).toBe(siteSection.parentElement);
    expect(userLookupSection.parentElement).toBe(siteSection.parentElement);
    expect(userCreateSection.parentElement).toBe(siteSection.parentElement);
    expect(productSection.textContent).toContain('First product');
    expect(siteSection.textContent).toContain('Test site');
    expect(getSiteInfo).toHaveBeenCalledTimes(1);
    expect(getSiteLists).toHaveBeenCalledTimes(1);
    expect(getUserById).not.toHaveBeenCalled();
    expect(createUser).not.toHaveBeenCalled();

    const firstSignal = getProducts.mock.calls[0][2] as AbortSignal;
    const nextButton = productSection.querySelectorAll('nav button')[1];

    if (!nextButton) {
      throw new Error('The next page button was not found.');
    }

    // Act: переходить на наступну сторінку з імітованою помилкою джерела товарів.
    await act(async () => {
      Simulate.click(nextButton);
    });

    // Assert: каталог обробив помилку, а секція сайту зберегла власні дані.
    expect(getProducts.mock.calls[1][0]).toBe(NEXT_PAGE_SKIP);
    expect(firstSignal.aborted).toBe(true);
    expect(getRequiredElement<HTMLElement>(productSection, '[role="alert"]').textContent)
      .toContain('Product source unavailable');
    expect(siteSection.textContent).toContain('Test site');
    expect(siteSection.querySelector('input')).toBeNull();
    expect(siteSection.querySelector('[role="alert"]')).toBeNull();
  } finally {
    // Очищує DOM і запускає скасування активних запитів компонентів.
    act(() => {
      ReactDOM.unmountComponentAtNode(container);
    });
    container.remove();
  }
});
