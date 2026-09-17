/** Замінює стилі для перевірки поведінки компонента у тестовому DOM. */
jest.mock('./SiteInfo.module.scss', () => ({}), { virtual: true });

import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { act } from 'react-dom/test-utils';
import SiteInfo from './SiteInfo';

const SITE_DETAILS_SELECTOR = 'dl';
const SITE_RESPONSE_SELECTOR = 'pre';
const SUCCESS_MESSAGE = 'Site information loaded.';
const TEST_SITE = {
  id: 'test-site-id',
  name: 'example-site',
  displayName: 'Example site',
  webUrl: 'https://example.sharepoint.com',
  description: 'Example description'
};

/** Перевіряє автоматичне завантаження та відображення відомостей поточного сайту. */
it('loads current site metadata automatically through its source', async () => {
  // Arrange
  const getSiteInfo = jest.fn().mockResolvedValue(TEST_SITE);
  const source = { getSiteInfo };
  
  const container = document.createElement('div');
  document.body.appendChild(container);

  try {
    // Act
    await act(async () => {
      ReactDOM.render(<SiteInfo source={source} />, container);
      await Promise.resolve();
    });

    // Assert
    expect(getSiteInfo).toHaveBeenCalledTimes(1);
    expect(container.textContent).toContain(SUCCESS_MESSAGE);
    expect(container.querySelector(SITE_DETAILS_SELECTOR)?.textContent).toContain(TEST_SITE.displayName);
    expect(container.querySelector(SITE_RESPONSE_SELECTOR)?.textContent).toContain(TEST_SITE.id);
    expect(container.querySelector('input')).toBeNull();
    expect(container.querySelector('button')).toBeNull();
  } finally {
    /** Демонтує форму та очищує її таймери після тесту. */
    act(() => {
      ReactDOM.unmountComponentAtNode(container);
    });
    container.remove();
  }
});
