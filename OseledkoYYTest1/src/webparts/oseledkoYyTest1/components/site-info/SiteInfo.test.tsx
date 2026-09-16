/** Замінює стилі для перевірки поведінки компонента у тестовому DOM. */
jest.mock('./SiteInfo.module.scss', () => ({}), { virtual: true });

import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { act, Simulate } from 'react-dom/test-utils';
import SiteInfo from './SiteInfo';

const SITE_FORM_SELECTOR = 'form';
const TOKEN_INPUT_SELECTOR = 'input';
const SITE_DETAILS_SELECTOR = 'dl';
const SITE_RESPONSE_SELECTOR = 'pre';
const ACCESS_TOKEN = 'test-token';
const SUCCESS_MESSAGE = 'Site information loaded.';
const AUTHORIZATION_ERROR_TEXT = 'invalid or expired';
const TEST_SITE = {
  id: 'test-site-id',
  name: 'example-site',
  displayName: 'Example site',
  webUrl: 'https://example.sharepoint.com',
  description: 'Example description'
};

const originalFetch = globalThis.fetch;

/** Повертає обов’язковий елемент тестового DOM або завершує тест зрозумілою помилкою. */
function getRequiredElement<T extends Element>(container: ParentNode, selector: string): T {
  const element = container.querySelector(selector);

  if (!element) {
    throw new Error(`Expected element matching "${selector}" to exist.`);
  }

  return element as T;
}

/** Відновлює мережеву функцію після завершення тесту. */
afterEach(() => {
  globalThis.fetch = originalFetch;
});

/** Перевіряє ручне завантаження, відображення відповіді та очищення старих даних після помилки. */
it('loads site metadata on submit and clears the previous result when authorization fails', async () => {
  // Arrange
  const fetchMock = jest.fn();
  globalThis.fetch = fetchMock;
  
  const container = document.createElement('div');
  document.body.appendChild(container);

  try {
    // Act
    act(() => {
      ReactDOM.render(<SiteInfo />, container);
    });

    // Assert
    expect(fetchMock).not.toHaveBeenCalled();
    const tokenInput = getRequiredElement<HTMLInputElement>(container, TOKEN_INPUT_SELECTOR);
    const form = getRequiredElement<HTMLFormElement>(container, SITE_FORM_SELECTOR);
    expect(tokenInput.type).toBe('password');

    // Act
    act(() => {
      tokenInput.value = ACCESS_TOKEN;
      Simulate.change(tokenInput);
    });
    fetchMock.mockResolvedValueOnce({
      ok: true,
      json: async () => TEST_SITE
    });
    await act(async () => {
      Simulate.submit(form);
    });

    // Assert
    expect(container.textContent).toContain(SUCCESS_MESSAGE);
    expect(container.querySelector(SITE_DETAILS_SELECTOR)?.textContent).toContain(TEST_SITE.displayName);
    expect(container.querySelector(SITE_RESPONSE_SELECTOR)?.textContent).toContain(TEST_SITE.id);
    expect(container.textContent).not.toContain(ACCESS_TOKEN);

    // Act
    fetchMock.mockResolvedValueOnce({ ok: false, status: 401 });
    await act(async () => {
      Simulate.submit(form);
    });

    // Assert
    expect(container.querySelector('[role="alert"]')?.textContent).toContain(AUTHORIZATION_ERROR_TEXT);
    expect(container.querySelector(SITE_DETAILS_SELECTOR)).toBeNull();
  } finally {
    /** Демонтує форму та очищує її таймери після тесту. */
    act(() => {
      ReactDOM.unmountComponentAtNode(container);
    });
    container.remove();
  }
});
