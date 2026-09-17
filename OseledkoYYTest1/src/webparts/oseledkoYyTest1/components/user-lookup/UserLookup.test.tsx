/** Замінює стилі для перевірки поведінки компонента у тестовому DOM. */
jest.mock('./UserLookup.module.scss', () => ({}), { virtual: true });

import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { act, Simulate } from 'react-dom/test-utils';
import UserLookup from './UserLookup';

const FORM_SELECTOR = 'form';
const USER_ID_INPUT_SELECTOR = 'input';
const FOUND_MESSAGE = 'User record loaded.';
const NOT_FOUND_MESSAGE = 'No user record was found for this ID.';
const TEST_USER = {
  id: 25,
  fields: {
    Id: 25,
    Title: 'Ada Lovelace',
    Department: 'Engineering'
  }
};

/** Повертає обов’язковий елемент тестового DOM або завершує тест зрозумілою помилкою. */
function getRequiredElement<T extends Element>(container: ParentNode, selector: string): T {
  const element = container.querySelector(selector);

  if (!element) {
    throw new Error(`The element "${selector}" was not found.`);
  }

  return element as T;
}

/** Перевіряє очищення нецифрових символів і відображення знайденого запису Users. */
it('keeps only digits in the user ID and displays a found user record', async () => {
  // Arrange
  const getUserById = jest.fn().mockResolvedValue(TEST_USER);
  const source = { getUserById };
  const container = document.createElement('div');
  document.body.appendChild(container);

  try {
    // Act
    act(() => {
      ReactDOM.render(<UserLookup source={source} />, container);
    });

    const input = getRequiredElement<HTMLInputElement>(container, USER_ID_INPUT_SELECTOR);
    const form = getRequiredElement<HTMLFormElement>(container, FORM_SELECTOR);
    act(() => {
      input.value = '2a5!';
      Simulate.change(input);
    });
    await act(async () => {
      Simulate.submit(form);
    });

    // Assert
    expect(input.value).toBe('25');
    expect(getUserById).toHaveBeenCalledWith(TEST_USER.id);
    expect(container.textContent).toContain(FOUND_MESSAGE);
    expect(container.textContent).toContain(TEST_USER.fields.Title);
    expect(container.querySelector('pre')?.textContent).toContain(TEST_USER.fields.Department);
  } finally {
    /** Демонтує компонент після завершення тесту. */
    act(() => {
      ReactDOM.unmountComponentAtNode(container);
    });
    container.remove();
  }
});

/** Перевіряє повідомлення для ID, за яким запис Users не знайдено. */
it('shows a not found message when the Users list has no matching ID', async () => {
  // Arrange
  const getUserById = jest.fn().mockResolvedValue(undefined);
  const source = { getUserById };
  const container = document.createElement('div');
  document.body.appendChild(container);

  try {
    // Act
    act(() => {
      ReactDOM.render(<UserLookup source={source} />, container);
    });

    const input = getRequiredElement<HTMLInputElement>(container, USER_ID_INPUT_SELECTOR);
    const form = getRequiredElement<HTMLFormElement>(container, FORM_SELECTOR);
    act(() => {
      input.value = '404';
      Simulate.change(input);
    });
    await act(async () => {
      Simulate.submit(form);
    });

    // Assert
    expect(getUserById).toHaveBeenCalledWith(404);
    expect(container.textContent).toContain(NOT_FOUND_MESSAGE);
    expect(container.querySelector('pre')).toBeNull();
  } finally {
    /** Демонтує компонент після завершення тесту. */
    act(() => {
      ReactDOM.unmountComponentAtNode(container);
    });
    container.remove();
  }
});
