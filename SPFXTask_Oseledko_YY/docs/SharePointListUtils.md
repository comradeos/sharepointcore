# SharePointListUtils

`SharePointListUtils.ts` — універсальна утиліта для SPFx, яка працює з одним списком SharePoint через REST API. Вона зосереджує в одному місці типові операції зі списком:

- читання однієї сторінки або всіх елементів
- створення, оновлення та видалення елементів
- захист оновлення через ETag
- фонове посторінкове завантаження з відміною
- тимчасове блокування елемента на час редагування
- автоматичне продовження блокування
- безпечне оновлення локального масиву в React

Файл утиліти:

```text
src/webparts/spfxTaskOseledkoYy/utils/SharePointListUtils.ts
```

---

## 1. Імпорт

```ts
import SharePointListUtils, {
  addListItems,
  ISharePointListEditSession,
  removeListItem,
  replaceListItem,
  SharePointBackgroundPageLoader,
  SharePointEditLockRenewal,
  SharePointListEditLockUtils
} from '../utils/SharePointListUtils';
```

Шлях імпорту змінюється залежно від папки файлу, який використовує утиліту.

---

## 2. Основні поняття

### Внутрішня назва поля

У `selectFields`, `filter`, `orderBy` та payload потрібно використовувати **внутрішні назви** полів SharePoint, а не назви, які бачить користувач у формі списку.

Наприклад, поле з відображуваною назвою `Власник блокування` може мати внутрішню назву `EditLockOwner`.

### ETag

ETag — версія конкретного елемента SharePoint. Після читання елемента утиліта отримує ETag із відповіді сервера. Якщо інший користувач встиг змінити елемент, старий ETag перестає бути актуальним, а SharePoint повертає HTTP 412.

Це запобігає непомітному перезапису чужих змін.

### Сторінка

SharePoint повертає елементи частинами. `ISharePointListPage<TItem>` містить:

```ts
interface ISharePointListPage<TItem> {
  items: TItem[];
  nextPageUrl?: string;
}
```

- `items` — елементи поточної частини
- `nextPageUrl` — адреса наступної частини
- відсутній `nextPageUrl` — це остання сторінка

---

## 3. Опис типів

### `ISharePointListOptions`

Налаштування екземпляра утиліти для одного списку.

| Поле           | Обов’язкове | Призначення                                                       |
|----------------|-------------|-------------------------------------------------------------------|
| `listTitle`    | так         | точна назва списку SharePoint                                     |
| `selectFields` | так         | внутрішні назви полів, які треба отримати                         |
| `expandFields` | ні          | поля типу Lookup або Person, які треба розгорнути через `$expand` |
| `pageSize`     | ні          | типовий розмір сторінки для цього списку                          |

Якщо `pageSize` не задано, утиліта передає `$top=5000`.

### `ISharePointListQueryOptions`

Параметри одного читання або читання всього списку.

| Поле       | Призначення                                                                     |
|------------|---------------------------------------------------------------------------------|
| `filter`   | OData-умова для `$filter`                                                       |
| `orderBy`  | сортування для `$orderby`, наприклад `Id desc`                                  |
| `pageSize` | розмір сторінки для конкретного запиту; має пріоритет над `pageSize` екземпляра |

### `ISharePointListItemWithETag<TItem>`

Результат `getByIdWithETag`.

```ts
interface ISharePointListItemWithETag<TItem> {
  item: TItem;
  eTag: string;
}
```

Значення `eTag` слід передати в `update`, якщо потрібне безпечне оновлення.

---

## 4. Створення екземпляра для списку

Приклад для списку `ServiceRequests`.

```ts
import { SPHttpClient } from '@microsoft/sp-http';
import SharePointListUtils from '../utils/SharePointListUtils';

interface IServiceRequest {
  Id: number;
  Title: string;
  Description?: string;
  Created: string;
  Category?: {
    Id: number;
    Title: string;
  };
  CategoryId?: number;
}

interface IServiceRequestPayload {
  Title: string;
  Description?: string;
  CategoryId?: number;
}

const requestsList = new SharePointListUtils<
  IServiceRequest,
  IServiceRequestPayload
>(
  this.context.spHttpClient,
  this.context.pageContext.web.absoluteUrl,
  {
    listTitle: 'ServiceRequests',
    selectFields: [
      'Id',
      'Title',
      'Description',
      'Created',
      'CategoryId',
      'Category/Id',
      'Category/Title'
    ],
    expandFields: ['Category'],
    pageSize: 1000
  }
);
```

Перший generic `TItem` описує відповідь SharePoint. Другий generic `TPayload` описує дані, які дозволено створити або оновити.

Для Person або Lookup-поля слід:

1. додати пов’язане поле в `selectFields`, наприклад `Category/Title`
2. додати саме поле в `expandFields`, наприклад `Category`
3. за потреби додати ідентифікатор, наприклад `CategoryId`

---

## 5. Читання даних

### `getPage` — одна сторінка

Використовуйте цей метод для великих списків, прогресу завантаження та власної пагінації.

```ts
const firstPage = await requestsList.getPage({
  filter: "Created ge datetime'2026-10-01T00:00:00Z' and Created lt datetime'2026-10-02T00:00:00Z'",
  orderBy: 'Id desc',
  pageSize: 1000
});

console.log(firstPage.items);
console.log(firstPage.nextPageUrl);
```

Щоб отримати наступну сторінку, передайте `nextPageUrl` другим аргументом:

```ts
const secondPage = await requestsList.getPage(
  {},
  firstPage.nextPageUrl
);
```

`nextPageUrl` уже містить параметри вихідного запиту, тому для наступної сторінки не потрібно повторювати `filter`, `orderBy` і `pageSize`.

### `getAll` — усі сторінки послідовно

```ts
const allActiveRequests = await requestsList.getAll({
  filter: 'IsActive eq 1',
  orderBy: 'Id desc',
  pageSize: 1000
});
```

Метод послідовно читає сторінки, доки SharePoint не перестане повертати `nextPageUrl`, після чого повертає один масив.

`getAll` доречний для невеликих довідників. Для великих робочих списків краще використовувати `getPage` або `SharePointBackgroundPageLoader`, щоб не чекати завершення всього завантаження до показу інтерфейсу.

### `getById` — елемент за ID

```ts
const request = await requestsList.getById(125);
```

`itemId` має бути додатним цілим числом.

### `getByIdWithETag` — елемент разом із версією

```ts
const { item, eTag } = await requestsList.getByIdWithETag(125);
```

Цей метод потрібен перед редагуванням, якщо потрібно перевіряти, що елемент не змінився між відкриттям форми й збереженням.

---

## 6. Фільтрування та сортування

Утиліта передає рядок `filter` у `$filter` та кодує його для URL. Синтаксис умови має відповідати OData для SharePoint REST API.

```ts
const page = await requestsList.getPage({
  filter: "Status eq 'New' and Created ge datetime'2026-10-01T00:00:00Z'",
  orderBy: 'Created desc',
  pageSize: 500
});
```

Приклади:

| Завдання         | `filter`                                    |
|------------------|---------------------------------------------|
| елемент за ID    | `Id eq 125`                                 |
| булеве поле      | `IsActive eq 1`                             |
| значення Lookup  | `CategoryId eq 3`                           |
| дата від включно | `Created ge datetime'2026-10-01T00:00:00Z'` |
| дата до виключно | `Created lt datetime'2026-10-02T00:00:00Z'` |

Для великого списку фільтри мають спиратися на індексовані поля. `$top=5000` не скасовує List View Threshold: SharePoint може повернути HTTP 500, якщо серверу треба переглянути понад 5000 елементів для виконання неіндексованого запиту.

---

## 7. Створення, оновлення та видалення

### `create`

```ts
const createdRequest = await requestsList.create({
  Title: 'Не працює принтер',
  Description: 'Принтер у кабінеті 204 не друкує',
  CategoryId: 3
});
```

Метод створює елемент, отримує його ID з відповіді SharePoint і потім повертає актуальний елемент через `getById`.

### `update` із ETag

```ts
const { item, eTag } = await requestsList.getByIdWithETag(125);

const updatedRequest = await requestsList.update(
  item.Id,
  {
    Title: item.Title,
    Description: 'Перевірено кабель живлення',
    CategoryId: item.CategoryId
  },
  eTag
);
```

Якщо елемент було змінено після `getByIdWithETag`, SharePoint поверне HTTP 412. Перевірити таку помилку можна через `isSharePointListVersionConflictError`.

```ts
import { isSharePointListVersionConflictError } from '../utils/SharePointListUtils';

try {
  await requestsList.update(125, payload, eTag);
} catch (error) {
  if (isSharePointListVersionConflictError(error)) {
    // покажіть користувачу повідомлення та перечитайте елемент
  }

  throw error;
}
```

Якщо третій аргумент `eTag` не передати, `update` використовує `'*'`. Це означає оновлення без перевірки версії. Такий режим придатний лише для даних, де паралельні зміни не мають значення.

### `delete`

```ts
await requestsList.delete(125);
```

Якщо другий аргумент `eTag` не передати, `delete` використовує `IF-MATCH: '*'` і видаляє елемент без перевірки версії.

```ts
const { eTag } = await requestsList.getByIdWithETag(125);
await requestsList.delete(125, eTag);
```

Для списку з блокуванням редагування не використовуйте `delete` напряму. Викликайте `deleteUnlocked`, описаний нижче.

---

## 8. Фонове посторінкове завантаження

`SharePointBackgroundPageLoader` виконує один запит, передає результат у `onPageLoaded`, а потім планує наступний запит лише за наявності курсора. Між сторінками є задана затримка.

### Приклад із `SharePointListUtils`

```ts
import {
  SharePointBackgroundPageLoader
} from '../utils/SharePointListUtils';

type RequestsCursor = string | undefined;

let requests: IServiceRequest[] = [];

const loader = new SharePointBackgroundPageLoader<
  IServiceRequest,
  RequestsCursor
>({
  delayMilliseconds: 300,
  pageSize: 1000,
  loadPage: async (nextPageUrl, pageSize) => {
    const page = await requestsList.getPage(
      {
        filter: "Created ge datetime'2026-10-01T00:00:00Z'",
        orderBy: 'Id desc',
        pageSize
      },
      nextPageUrl
    );

    return {
      items: page.items,
      cursor: page.nextPageUrl
    };
  },
  onPageLoaded: page => {
    requests = [...requests, ...page.items];
    // оновіть state React, прогрес або таблицю
  },
  onError: error => {
    console.error('Не вдалося завантажити наступну сторінку', error);
  },
  onCompleted: () => {
    console.log('Завантаження завершено');
  }
});

loader.start(undefined);
```

### Методи loader

| Метод           | Призначення                                                                              |
|-----------------|------------------------------------------------------------------------------------------|
| `start(cursor)` | скасовує попередній запуск і починає новий із вказаного курсора; повертає версію запуску |
| `cancel()`      | прибирає запланований таймер і забороняє застосовувати відповіді попереднього запуску    |
| `getVersion()`  | повертає поточну версію запуску для перевірки пов’язаних асинхронних операцій            |

`cancel()` не може фізично перервати HTTP-запит, який уже надіслано. Утиліта ігнорує його відповідь, якщо за час очікування було скасовано або перезапущено завантаження.

### Рекомендації

- показуйте першу сторінку одразу, а наступні завантажуйте через loader
- перед новим пошуком, зміною фільтра або демонтуванням компонента викликайте `loader.cancel()`
- робіть `pageSize` достатнім для інтерфейсу, але не більшим за 5000
- не запускайте кілька loader для однієї таблиці одночасно
- не вважайте кількість отриманих елементів загальною кількістю без окремого підрахунку

---

## 9. Тимчасове блокування редагування

### Призначення

`SharePointListEditLockUtils` захищає елемент, коли його редагує конкретна вкладка браузера. Під час захоплення створюється випадковий token. Інша вкладка не може зберегти або продовжити чуже блокування.

Блокування має строк дії. Якщо користувач закрив вкладку без скасування, продовження припиняється, і після завершення строку інший користувач зможе відкрити елемент.

### Поля, які треба створити у списку

Для списку `ServiceRequests` створіть три поля:

| Відображувана назва            | Рекомендована внутрішня назва | Тип                              |
|--------------------------------|-------------------------------|----------------------------------|
| Власник блокування редагування | `EditLockOwner`               | Особа або група, один користувач |
| Блокування дійсне до           | `EditLockExpiresAt`           | Дата й час                       |
| Ключ блокування                | `EditLockToken`               | Однорядковий текст               |

Поле `EditLockOwner` у REST-моделі також має технічне значення `EditLockOwnerId`. Для відображення імені власника необхідно отримати `EditLockOwner/Title` і виконати `$expand=EditLockOwner`.

### Типи даних елемента та payload

```ts
interface IServiceRequest {
  Id: number;
  Title: string;
  EditLockOwnerId?: number | null;
  EditLockOwner?: {
    Id: number;
    Title: string;
  };
  EditLockExpiresAt?: string | null;
  EditLockToken?: string | null;
}

interface IServiceRequestPayload {
  Title?: string;
  Description?: string | null;
  EditLockOwnerId?: number | null;
  EditLockExpiresAt?: string | null;
  EditLockToken?: string | null;
}
```

### Підготовка `SharePointListUtils`

Поля блокування мають бути в `selectFields`.

```ts
const requestsList = new SharePointListUtils<
  IServiceRequest,
  IServiceRequestPayload
>(client, webUrl, {
  listTitle: 'ServiceRequests',
  selectFields: [
    'Id',
    'Title',
    'EditLockOwnerId',
    'EditLockOwner/Title',
    'EditLockExpiresAt',
    'EditLockToken'
  ],
  expandFields: ['EditLockOwner']
});
```

### Створення менеджера блокувань

```ts
const editLock = new SharePointListEditLockUtils(requestsList, {
  createPayload: values => {
    const isLockBeingCleared = values.token === undefined;

    return {
      EditLockOwnerId: isLockBeingCleared ? null : values.ownerId,
      EditLockExpiresAt: isLockBeingCleared ? null : values.expiresAt,
      EditLockToken: isLockBeingCleared ? null : values.token
    };
  },
  getOwnerId: request => request.EditLockOwnerId ?? undefined,
  getOwnerName: request => request.EditLockOwner?.Title ?? undefined,
  getExpiresAt: request => request.EditLockExpiresAt ?? undefined,
  getToken: request => request.EditLockToken ?? undefined,
  itemName: 'заявку',
  itemNameGenitive: 'заявки',
  durationMilliseconds: 5 * 60 * 1000
});
```

`durationMilliseconds` необов’язковий. Якщо його не передати, строк блокування становить 5 хвилин.

### Життєвий цикл редагування

### `acquire` — захоплення блокування

```ts
let editSession: ISharePointListEditSession<IServiceRequest> | undefined;

try {
  editSession = await editLock.acquire(requestId, currentUserId);
  // відкрийте форму з editSession.item
} catch (error) {
  // покажіть текст error користувачу
  // наприклад: "Заявку редагує Іван Петренко до 14:30"
}
```

`acquire` читає елемент з ETag, перевіряє строк чинного блокування і встановлює новий token. Якщо двоє користувачів натиснуть редагування одночасно, ETag не дозволить обом успішно захопити блокування. Після конфлікту версії утиліта один раз перечитує елемент і повторює перевірку.

### `renew` — продовження строку

```ts
await editLock.renew(editSession.item.Id, editSession.token);
```

Викликайте цей метод регулярно, доки форма відкрита. Він перевіряє, що token належить поточній вкладці, і подовжує строк блокування.

### `updateLocked` — збереження форми

```ts
try {
  if (!editSession) {
    return;
  }

  const updatedRequest = await editLock.updateLocked(
    editSession.item.Id,
    editSession.token,
    {
      Title: formValues.title,
      Description: formValues.description
    }
  );

  // updateLocked одночасно зберігає дані й очищує поля блокування
  editSession = undefined;
} catch (error) {
  // блокування могло завершитися або дані змінив інший користувач
}
```

`updateLocked` перевіряє token і строк дії, використовує ETag та очищує поля блокування в тому самому запиті, що й зберігає дані форми.

### `release` — скасування або закриття форми

```ts
await editLock.release(editSession.item.Id, editSession.token);
```

Викликайте `release` після натискання `Скасувати`, закриття модального вікна та у `componentWillUnmount`. Якщо блокування вже належить іншій вкладці, `release` нічого не змінить.

Якщо вкладку або браузер закрито аварійно, `release` може не виконатися. Це штатна ситуація: блокування стане нечинним після `durationMilliseconds`.

### `deleteUnlocked` — видалення без чинного блокування

```ts
await editLock.deleteUnlocked(requestId);
```

Метод читає актуальний елемент разом з ETag. Якщо будь-яка вкладка має чинне блокування, він повертає повідомлення з ім’ям користувача та не виконує видалення.

Якщо блокування з’явилося після перевірки, SharePoint поверне HTTP 412, бо ETag уже змінився. Утиліта перечитає елемент, визначить нове блокування і також відхилить видалення. Тому між перевіркою та видаленням не виникає вікно, у якому можна видалити заявку, що вже відкрили для редагування.

---

## 10. Автоматичне продовження блокування

`SharePointEditLockRenewal` викликає `renew` через заданий інтервал і припиняє роботу після помилки.

```ts
import { SharePointEditLockRenewal } from '../utils/SharePointListUtils';

const renewal = new SharePointEditLockRenewal({
  intervalMilliseconds: 60 * 1000,
  renew: () => editLock.renew(editSession.item.Id, editSession.token),
  onRenewed: () => {
    // за потреби оновіть стан форми
  },
  onError: error => {
    // заблокуйте збереження та покажіть помилку
    console.error('Блокування втрачено', error);
  }
});

renewal.start();
```

Інтервал повинен бути меншим за строк блокування. Для блокування на 5 хвилин підходить інтервал 1 хвилина.

| Метод                      | Призначення                                                   |
|----------------------------|---------------------------------------------------------------|
| `start()`                  | запускає періодичне продовження                               |
| `stop()`                   | зупиняє таймер і відкидає результат запиту, що ще виконується |
| `getIsRenewalInProgress()` | повертає `true`, поки виконується запит `renew`               |

Перед закриттям форми викликайте `renewal.stop()`, а потім `editLock.release(...)`.

---

## 11. Допоміжні функції для локального стану

Ці функції не виконують HTTP-запитів. Вони повертають новий масив, тому зручні для React state.

### Додати елементи на початок

```ts
const nextRequests = addListItems(currentRequests, [createdRequest]);
```

Результат: `createdRequest` буде першим у масиві.

### Замінити елемент

```ts
const nextRequests = replaceListItem(
  currentRequests,
  updatedRequest,
  'Id'
);
```

### Видалити елемент

```ts
const nextRequests = removeListItem(
  currentRequests,
  requestId,
  'Id'
);
```

---

## 12. Обробка помилок

Утиліта повертає помилки у вигляді `Error` з дією, назвою списку, HTTP-статусом і текстом помилки SharePoint, якщо сервер його надав.

```ts
try {
  const page = await requestsList.getPage();
  // використайте page.items
} catch (error) {
  const message = error instanceof Error
    ? error.message
    : 'Сталася невідома помилка';

  this.setState({ error: message });
}
```

Типові причини:

| Код або ситуація                | Дія                                                                                               |
|---------------------------------|---------------------------------------------------------------------------------------------------|
| HTTP 412                        | перечитати елемент; користувач або інша вкладка вже змінили його                                  |
| HTTP 500 із List View Threshold | перевірити індекси полів і серверний фільтр; не намагатися обійти поріг лише збільшенням `$top`   |
| 403                             | перевірити права користувача на список                                                            |
| 404                             | перевірити назву списку та адресу сайту                                                           |
| немає ETag                      | переконатися, що елемент читається через `getByIdWithETag` і SharePoint повертає заголовок `ETag` |

---

## 13. Мінімальний сценарій у React

```ts
private readonly requestsList = new SharePointListUtils<
  IServiceRequest,
  IServiceRequestPayload
>(
  this.props.spHttpClient,
  this.props.webUrl,
  {
    listTitle: 'ServiceRequests',
    selectFields: ['Id', 'Title', 'Created'],
    pageSize: 1000
  }
);

private readonly loadRequests = async (): Promise<void> => {
  this.setState({ isLoading: true, error: undefined });

  try {
    const page = await this.requestsList.getPage({
      orderBy: 'Id desc'
    });

    this.setState({
      requests: page.items,
      isLoading: false
    });
  } catch (error) {
    this.setState({
      isLoading: false,
      error: error instanceof Error
        ? error.message
        : 'Не вдалося завантажити заявки'
    });
  }
};
```

Для списку, який може містити багато елементів, після першого `getPage` створіть `SharePointBackgroundPageLoader` і передайте в нього курсор `page.nextPageUrl`.

---

## 14. Контрольний список інтеграції

1. Створити TypeScript-інтерфейси для елемента списку та payload.
2. Вказати точну назву списку в `listTitle`.
3. Додати потрібні внутрішні назви до `selectFields`.
4. Для Lookup і Person додати поля у `expandFields`.
5. Для великого списку індексувати поля, за якими виконується серверний фільтр.
6. Використовувати `getPage` або фоновий loader для великих вибірок.
7. Використовувати `getByIdWithETag` та `update(..., eTag)` для критичних змін.
8. Для блокування редагування створити три поля блокування у списку та додати їх до `selectFields`.
9. Після `acquire` запускати `SharePointEditLockRenewal`.
10. Після збереження, скасування або демонтування компонента зупиняти renewal і викликати `release`.
