/** Замінює стилі для перевірки поведінки компонента у тестовому DOM. */
jest.mock('./OseledkoYyTest1.module.scss', () => ({}), { virtual: true });

import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { act, Simulate } from 'react-dom/test-utils';
import SiteInfo from './SiteInfo';


const originalFetch = globalThis.fetch;
/** Відновлює мережеву функцію після завершення тесту. */
afterEach(() => { globalThis.fetch = originalFetch; });

/** Перевіряє ручне завантаження, відображення відповіді та очищення старих даних після помилки. */
it('loads site metadata on submit and clears the previous result when authorization fails', async () => {
  const fetchMock = jest.fn();
  globalThis.fetch = fetchMock;
  const container = document.createElement('div');
  document.body.appendChild(container);
  const site = { id: 'test-site-id', name: 'example-site', displayName: 'Example site', webUrl: 'https://example.sharepoint.com', description: 'Example description' };
  try {
    /** Монтує форму без автоматичного мережевого запиту. */
    act(() => { ReactDOM.render(<SiteInfo />, container); });
    expect(fetchMock).not.toHaveBeenCalled();
    const input = container.querySelector('input') as HTMLInputElement;
    expect(input.type).toBe('password');
    /** Імітує введення тестового токена користувачем. */
    act(() => { input.value = 'test-token'; Simulate.change(input); });
    fetchMock.mockResolvedValueOnce({ ok: true, json: /** Повертає успішну тестову відповідь Graph. */ async () => site });
    /** Відправляє форму та очікує завершення асинхронного оновлення React. */
    await act(async () => { Simulate.submit(container.querySelector('form') as HTMLFormElement); });
    expect(container.textContent).toContain('Site information loaded.');
    expect(container.querySelector('dl')?.textContent).toContain('Example site');
    expect(container.querySelector('pre')?.textContent).toContain('test-site-id');
    expect(container.textContent).not.toContain('test-token');
    fetchMock.mockResolvedValueOnce({ ok: false, status: 401 });
    /** Повторює запит із відмовою авторизації для перевірки очищення результату. */
    await act(async () => { Simulate.submit(container.querySelector('form') as HTMLFormElement); });
    expect(container.querySelector('[role="alert"]')?.textContent).toContain('invalid or expired');
    expect(container.querySelector('dl')).toBeNull();
  } finally {
    /** Демонтує форму та очищує її таймери після тесту. */
    act(() => { ReactDOM.unmountComponentAtNode(container); });
    container.remove();
  }
});
