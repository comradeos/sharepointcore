/** Замінює стилі для перевірки поведінки компонента у тестовому DOM. */
jest.mock('./SiteLists.module.scss', () => ({}), { virtual: true });

import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { act } from 'react-dom/test-utils';
import SiteLists from './SiteLists';

const LIST_TITLE = 'Project tasks';
const LIST_TEMPLATE_TEXT = 'List';
const LIST_ITEM_COUNT = '8';
const TEST_LISTS = [
  {
    id: 'tasks-list-id',
    title: LIST_TITLE,
    itemCount: 8,
    template: 100,
    webUrl: 'https://example.sharepoint.com/sites/test/Lists/Tasks'
  }
];

/** Перевіряє автоматичне завантаження та відображення видимих списків сайту. */
it('loads visible SharePoint lists automatically', async () => {
  // Arrange
  const getSiteLists = jest.fn().mockResolvedValue(TEST_LISTS);
  const source = { getSiteLists };
  const container = document.createElement('div');
  document.body.appendChild(container);

  try {
    // Act
    await act(async () => {
      ReactDOM.render(<SiteLists source={source} />, container);
      await Promise.resolve();
    });

    // Assert
    expect(getSiteLists).toHaveBeenCalledTimes(1);
    expect(container.textContent).toContain(LIST_TITLE);
    expect(container.textContent).toContain(LIST_TEMPLATE_TEXT);
    expect(container.textContent).toContain(LIST_ITEM_COUNT);
    expect(container.querySelector('a')?.getAttribute('href')).toBe(TEST_LISTS[0].webUrl);
  } finally {
    /** Демонтує компонент після завершення тесту. */
    act(() => {
      ReactDOM.unmountComponentAtNode(container);
    });
    container.remove();
  }
});
