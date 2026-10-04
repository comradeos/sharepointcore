import { IPersonaProps } from '@fluentui/react';
import { IPeoplePickerPersona } from './ServiceRequestForm.types';

const allowedEmailDomain = '@ua.energy';
const millisecondsPerMinute = 60 * 1000;

// залишає у результатах пошуку користувачів з дозволеною поштою
export function filterUaEnergyUsers(results: IPersonaProps[]): IPersonaProps[] {
  const filteredResults: IPersonaProps[] = [];

  for (const result of results) {
    const email = result.secondaryText?.trim().toLowerCase() ?? '';

    if (email.endsWith(allowedEmailDomain)) {
      filteredResults.push(result);
    }
  }

  return filteredResults;
}

// перетворює дату sharepoint на локальне значення поля дати
export function toDateTimeLocal(value?: string): string {
  if (!value) {
    return '';
  }

  const date = new Date(value);
  if (Number.isNaN(date.getTime())) {
    return '';
  }

  const localDate = new Date(
    date.getTime() - date.getTimezoneOffset() * millisecondsPerMinute
  );

  return localDate.toISOString().slice(0, 16);
}

// повертає поточну локальну дату та час без секунд
export function getMinimumDateTimeValue(): string {
  const currentDate = new Date();
  currentDate.setSeconds(0, 0);
  const localDate = new Date(
    currentDate.getTime() - currentDate.getTimezoneOffset() * millisecondsPerMinute
  );
  return localDate.toISOString().slice(0, 16);
}

// повертає адресу або логін вибраного користувача
export function getPersonIdentity(people: IPersonaProps[], fallback = ''): string {
  const person = people[0] as IPeoplePickerPersona | undefined;

  if (!person) {
    return fallback;
  }

  return person.loginName || person.secondaryText || String(person.id || fallback);
}

// перевіряє поштовий домен вибраного користувача
export function hasAllowedEmailDomain(people: IPersonaProps[]): boolean {
  const identity = getPersonIdentity(people).trim().toLowerCase();
  return identity.endsWith(allowedEmailDomain);
}
