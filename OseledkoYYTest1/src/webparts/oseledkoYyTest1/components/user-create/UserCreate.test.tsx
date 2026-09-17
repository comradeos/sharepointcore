/** Замінює стилі для перевірки поведінки компонента у тестовому DOM. */
jest.mock('./UserCreate.module.scss', () => ({}), { virtual: true });

import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { act, Simulate } from 'react-dom/test-utils';
import UserCreate from './UserCreate';

const FORM_SELECTOR = 'form';
const USER_NAME_INPUT_SELECTOR = 'input';
const USER_NAME = 'Ada Lovelace';
const CREATED_USER_ID = 42;
const CREATED_MESSAGE = `User record created. New ID: ${CREATED_USER_ID}.`;
const EMPTY_NAME_MESSAGE = 'Enter a user name.';

/** Повертає обов’язковий елемент тестового DOM або завершує тест зрозумілою помилкою. */
function getRequiredElement<T extends Element>(container: ParentNode, selector: string): T {
  const element = container.querySelector(selector);

  if (!element) {
    throw new Error(`The element "${selector}" was not found.`);
  }

  return element as T;
}

/** Перевіряє створення запису Users і відображення ID, поверненого джерелом. */
it('creates a user record and displays the generated ID', async () => {
  // Arrange
  const createUser = jest.fn().mockResolvedValue(CREATED_USER_ID);
  const source = { createUser };
  const container = document.createElement('div');
  document.body.appendChild(container);

  try {
    // Act
    act(() => {
      ReactDOM.render(<UserCreate source={source} />, container);
    });

    const input = getRequiredElement<HTMLInputElement>(container, USER_NAME_INPUT_SELECTOR);
    const form = getRequiredElement<HTMLFormElement>(container, FORM_SELECTOR);
    act(() => {
      input.value = `  ${USER_NAME}  `;
      Simulate.change(input);
    });
    await act(async () => {
      Simulate.submit(form);
    });

    // Assert
    expect(createUser).toHaveBeenCalledWith(USER_NAME);
    expect(input.value).toBe('');
    expect(container.textContent).toContain(CREATED_MESSAGE);
  } finally {
    /** Демонтує компонент після завершення тесту. */
    act(() => {
      ReactDOM.unmountComponentAtNode(container);
    });
    container.remove();
  }
});

/** Перевіряє локальне відхилення порожньої назви без створення запису. */
it('rejects an empty user name without creating a record', async () => {
  // Arrange
  const createUser = jest.fn();
  const source = { createUser };
  const container = document.createElement('div');
  document.body.appendChild(container);

  try {
    // Act
    act(() => {
      ReactDOM.render(<UserCreate source={source} />, container);
    });
    const form = getRequiredElement<HTMLFormElement>(container, FORM_SELECTOR);
    act(() => {
      Simulate.submit(form);
    });

    // Assert
    expect(createUser).not.toHaveBeenCalled();
    expect(container.textContent).toContain(EMPTY_NAME_MESSAGE);
  } finally {
    /** Демонтує компонент після завершення тесту. */
    act(() => {
      ReactDOM.unmountComponentAtNode(container);
    });
    container.remove();
  }
});
