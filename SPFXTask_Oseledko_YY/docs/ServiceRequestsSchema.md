# Схема списків Service Desk

Сайт: https://uaenergy0.sharepoint.com.mcas.ms/sites/oseledko-yy-test
Призначення: тестова SPFx вебчастина для CRUD заявок, каскадних довідників і фільтрів ag-Grid.
Статус: три списки та їхні стовпці створені користувачем і перевірені за скриншотами. Читання та створення через REST API перевірені у вебчастині. Перед завершенням потрібна ручна перевірка редагування, видалення, фільтрів і адаптивного відображення.

## 1. RequestCategories

| Internal name | Підпис | Тип SharePoint | Обов'язкове | Примітка |
| --- | --- | --- | --- | --- |
| Title | Категорія | Single line of text | Так | Унікальна назва серед активних категорій. |
| IsActive | Активна | Yes/No | Так | За замовчуванням Так. Історичні записи не видаляти. |

## 2. RequestSubcategories

| Internal name | Підпис | Тип SharePoint | Обов'язкове | Примітка |
| --- | --- | --- | --- | --- |
| Title | Підкатегорія | Single line of text | Так | Назва може повторюватися в різних категоріях. |
| Category | Категорія | Lookup (single) | Так | Посилання на RequestCategories.Title. |
| IsActive | Активна | Yes/No | Так | За замовчуванням Так. |

## 3. ServiceRequests

| Internal name | Підпис | Тип SharePoint | Обов'язкове | Правило |
| --- | --- | --- | --- | --- |
| Title | Назва заявки | Single line of text | Так | 3–255 символів після обрізання пробілів. |
| Description | Опис | Multiple lines of text (plain text) | Так | Не порожній після обрізання пробілів. |
| Category | Категорія | Lookup (single) | Так | Посилання на RequestCategories.Title. |
| Subcategory | Підкатегорія | Lookup (single) | Так | Посилання на RequestSubcategories.Title; належить обраній категорії. |
| Status | Статус | Choice (single) | Так | Нова / В роботі / Вирішена / Закрита; за замовчуванням Нова. |
| Priority | Пріоритет | Choice (single) | Так | Низький / Середній / Високий; за замовчуванням Середній. |
| Requester | Заявник | Person or Group (person, single) | Так | За замовчуванням поточний користувач. |
| Assignee | Виконавець | Person or Group (person, single) | Умовно | Обов'язковий для статусів Вирішена та Закрита. |
| PlannedStart | Плановий початок | Date and Time | Ні | Якщо заданий, не пізніше DueDate. |
| DueDate | Кінцевий термін | Date and Time | Так | Не раніше PlannedStart, якщо той заданий. |
| EstimatedHours | Оцінка часу (год) | Number, 1 decimal | Ні | Якщо задана, більше 0 (мінімум 0,1 за точності 1 десятковий знак); верхню межу не задавати. |
| ContactEmail | Контактний email | Single line of text | Ні | Якщо заданий, перевірка формату email. |
| RequiresOnsiteVisit | Потрібен виїзд | Yes/No | Так | За замовчуванням Ні. |

Власних полів у ServiceRequests: 13. Типів: Text, Note, Lookup, Choice, Person, DateTime, Number, Boolean (8). Службові ID, Created, Modified, Author не враховані.

## Зв'язки та каскад

- RequestCategories (1) -> (N) RequestSubcategories через RequestSubcategories.Category.
- RequestCategories (1) -> (N) ServiceRequests через ServiceRequests.Category.
- RequestSubcategories (1) -> (N) ServiceRequests через ServiceRequests.Subcategory.
- У формі Create/Edit підкатегорії фільтруються за Category ID. При зміні категорії вибір підкатегорії скидається.
- Перед збереженням перевіряється, що Subcategory.Category ID дорівнює обраному ServiceRequests.Category ID.
- Неактивні значення не пропонуються для нових заявок. Уже збережені значення залишаються видимими в View/Edit.

## Початкові дані довідників

| Категорія | Підкатегорії |
| --- | --- |
| Обладнання | Ноутбук, Монітор |
| Програмне забезпечення | Office, CRM |
| Мережа | Wi-Fi, VPN |

## Відповідність фільтрам ag-Grid

| Тип фільтра | Поле |
| --- | --- |
| Текстовий | Title, Description |
| Одиночний вибір | Category або Priority |
| Множинний вибір | Status |
| Дата | DueDate |
| Числовий | EstimatedHours |

Додатково: загальний пошук по видимих текстових значеннях, сортування колонок і кнопка оновлення без перезавантаження сторінки.

## Порядок створення

1. Створити RequestCategories та додати початкові дані.
2. Створити RequestSubcategories з Lookup Category та додати початкові дані.
3. Створити ServiceRequests з двома Lookup полями та рештою колонок.
4. Перевірити внутрішні імена, типи полів, значення Choice і права на CRUD.
5. Створити кілька тестових заявок, перевірити каскад і всі операції через SPFx вебчастину.
