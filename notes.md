# SharePoint Framework (SPFx): базовый конспект

## 1. Что такое SharePoint

**SharePoint** — платформа Microsoft 365 для создания корпоративных сайтов, хранения документов и организации совместной работы.

В SharePoint есть:

- сайты и страницы;
- списки — структурированные данные, похожие на таблицы;
- библиотеки документов;
- пользователи, группы и права доступа;
- стандартные веб-части — готовые блоки страницы.

Если стандартных возможностей недостаточно, SharePoint можно расширять с помощью **SharePoint Framework**.

## 2. Что такое SharePoint Framework (SPFx)

**SharePoint Framework**, сокращённо **SPFx**, — официальная клиентская модель разработки Microsoft для создания собственных компонентов и расширений SharePoint.

Проще говоря:

> SPFx позволяет написать веб-приложение или элемент интерфейса на TypeScript/JavaScript и встроить его в SharePoint как полноценную часть страницы.

SPFx также используется для расширения Microsoft Teams и Viva Connections.

### Что можно создавать

- **Web Part (веб-часть)** — визуальный блок, который пользователь добавляет на страницу: таблица, форма, карточки сотрудников, новости, отчёт и т. п.
- **Extension (расширение)** — изменение поведения или интерфейса SharePoint: верхняя/нижняя панель, пользовательское действие, собственное отображение полей или форм списка.
- **Adaptive Card Extension (ACE)** — компактная интерактивная карточка, прежде всего для Viva Connections.
- **Library Component** — общий код, который могут использовать несколько SPFx-компонентов.

## 3. Главный принцип работы

SPFx-компонент выполняется **в браузере пользователя**, внутри страницы SharePoint и в контексте уже вошедшего пользователя.

Упрощённая схема:

```text
Пользователь открывает страницу SharePoint
                  ↓
SharePoint загружает установленный SPFx-компонент
                  ↓
JavaScript/TypeScript-код выполняется в браузере
                  ↓
Компонент получает контекст сайта и пользователя
                  ↓
Через SharePoint REST API или Microsoft Graph получает данные
                  ↓
Отрисовывает интерфейс в DOM страницы
```

В отличие от старых подходов, компонент обычно не помещается в отдельный `iframe`. Он встраивается прямо в DOM страницы, поэтому выглядит и ведёт себя как часть SharePoint.

## 4. Из чего состоит SPFx-решение

### Компонент

Это веб-часть, расширение или ACE, которые реализуют конкретную функцию.

### Solution (решение)

Проект может содержать один или несколько компонентов. При публикации они собираются в пакет решения с расширением **`.sppkg`**.

### App Catalog

**App Catalog** — специальный каталог приложений SharePoint. Администратор загружает туда `.sppkg`, проверяет и развёртывает решение.

### Tenant

**Tenant (тенант)** — отдельная организация и её среда Microsoft 365. Внутри тенанта находятся пользователи, сайты SharePoint, Teams и другие сервисы.

### Context

**Контекст SPFx** содержит сведения о текущем сайте, странице, пользователе и доступных сервисах. Благодаря ему компонент понимает, где и от чьего имени он запущен.

## 5. Доступ к данным и безопасность

SPFx-код работает от имени текущего пользователя. Он не получает автоматически больше прав, чем есть у этого пользователя.

Для данных обычно используются:

- **SharePoint REST API** — работа с сайтами, списками, библиотеками и элементами SharePoint;
- **Microsoft Graph** — единый API для данных Microsoft 365: пользователей, групп, Teams, почты, календарей и т. д.;
- `SPHttpClient`, `HttpClient` и `MSGraphClient` — предоставляемые SPFx клиенты для HTTP-запросов и аутентификации.

Важно: запуск кода в контексте пользователя не означает, что любой внешний API уже разрешён. Для некоторых разрешений Microsoft Graph или собственных API требуется отдельное одобрение администратора.

## 6. Основные технологии

- **TypeScript** — основной язык разработки; компилируется в JavaScript.
- **JavaScript** — выполняется браузером.
- **React** — популярный, но не обязательный способ создания интерфейса.
- **Node.js** — запускает инструменты сборки на компьютере разработчика; сам SPFx-компонент работает не в Node.js, а в браузере.
- **npm** — устанавливает библиотеки и инструменты проекта.
- **Heft или Gulp** — инструменты сборки; выбор зависит от версии SPFx. Начиная с SPFx 1.22 основной toolchain перешёл с Gulp на Heft.
- **VS Code** — популярный редактор для разработки.

SPFx не требует React: технически можно использовать и другие JavaScript-фреймворки или обычный TypeScript. Но React часто встречается в документации и реальных проектах.

## 7. Жизненный цикл разработки

1. Разработчик создаёт SPFx-проект и выбирает тип компонента.
2. Пишет TypeScript-код и пользовательский интерфейс.
3. Локально запускает и проверяет компонент в Workbench или на тестовом сайте.
4. Создаёт production-сборку.
5. Упаковывает решение в файл `.sppkg`.
6. Загружает пакет в App Catalog.
7. Администратор развёртывает решение, после чего оно становится доступным на разрешённых сайтах.
8. Пользователь добавляет веб-часть на страницу либо расширение подключается согласно его настройке.

## 8. SPFx и SPFx Toolkit — не одно и то же

**SPFx** — платформа и модель выполнения компонентов.

**SharePoint Framework Toolkit** — расширение для VS Code от сообщества Microsoft 365 & Power Platform PnP. Оно помогает:

- подготовить окружение под нужную версию SPFx;
- проверить совместимость Node.js и зависимостей;
- создать, собрать, упаковать и опубликовать решение;
- обновить старый SPFx-проект;
- управлять App Catalog, приложениями и некоторыми объектами SharePoint;
- использовать CLI for Microsoft 365 и некоторые AI-возможности из VS Code.

Toolkit повышает удобство работы, но не заменяет знания SPFx, TypeScript, SharePoint API и процесса развёртывания.

## 9. Минимальная ментальная модель

Запомнить можно так:

```text
SPFx = правила + API + инструменты для расширения SharePoint
Web Part/Extension = написанный нами компонент
.sppkg = пакет компонента для установки
App Catalog = место публикации пакета
SharePoint page = место выполнения компонента
Текущий пользователь = его права используются при доступе к данным
SPFx Toolkit = помощник разработчика в VS Code
```

## 10. Что изучать дальше

1. Базовые сущности SharePoint: site, page, list, library, item, permissions.
2. TypeScript и основы React.
3. Структура SPFx-проекта и жизненный цикл веб-части.
4. Чтение данных списка через SharePoint REST API.
5. Работа с Microsoft Graph.
6. Сборка `.sppkg`, App Catalog и развёртывание.
7. Extensions, ACE и более сложные сценарии.

## Источники

- [Обзор SharePoint Framework — Microsoft Learn](https://learn.microsoft.com/ru-ru/sharepoint/dev/spfx/sharepoint-framework-overview)
- [SharePoint Framework Toolkit 2025 wrap up — Microsoft 365 & Power Platform Community](https://pnp.github.io/blog/post/spfx-toolkit-vscode-2025-wrap-up/)

## 11. SharePoint Server

**SharePoint Server** — серверная версия SharePoint, которую организация устанавливает и обслуживает самостоятельно. Её также называют **SharePoint On-Premises**, **SharePoint On-Prem** или локальным SharePoint.

В отличие от SPFx, SharePoint Server является полноценной серверной системой: у него есть серверные приложения, службы, базы данных, фоновые процессы, система прав, поиск и административные инструменты.

### SharePoint Server и SharePoint Online

| SharePoint Online | SharePoint Server |
|---|---|
| Работает в облаке Microsoft 365 | Устанавливается на серверы организации |
| Инфраструктуру обслуживает Microsoft | Инфраструктуру обслуживает организация |
| Microsoft управляет обновлениями платформы | Администраторы организации устанавливают обновления |
| SPFx поддерживается наиболее полно | Поддерживаются только определённые версии SPFx |
| Не нужно самостоятельно развёртывать SharePoint и SQL Server | Нужны серверы, сеть, лицензии, настройка и резервное копирование |

### Базовая архитектура

SharePoint Server устанавливается на **Windows Server**, а для постоянного хранения данных обычно использует **Microsoft SQL Server**.

```text
Браузер пользователя
          ↓
SharePoint Server
          ↓
SQL Server
```

Для небольшой тестовой среды SharePoint и SQL Server могут находиться на одной машине. В production SQL Server обычно разворачивается отдельно.

Совокупность связанных серверов SharePoint называется **фермой SharePoint (SharePoint Farm)**.

Пример более крупной фермы:

```text
                   ┌─ SharePoint Front-end 1
Пользователи → балансировщик нагрузки
                   └─ SharePoint Front-end 2
                              ↓
                    Application Servers
                              ↓
                    SQL Server Cluster
```

### Основные части SharePoint Server

- сайты и страницы;
- списки и библиотеки документов;
- пользователи, группы и права доступа;
- REST API и CSOM;
- поиск и индексация;
- серверные службы и Service Applications;
- фоновые задания — Timer Jobs;
- Central Administration — панель управления фермой;
- конфигурационные базы и базы содержимого в SQL Server.

### Где хранятся данные

В базах SQL Server хранятся:

- содержимое и настройки сайтов;
- страницы;
- элементы списков;
- документы и их метаданные;
- конфигурация фермы;
- данные некоторых серверных служб.

```text
Сайт SharePoint
      ↓
Content Database в SQL Server
      ↓
Страницы, списки, элементы и документы
```

Разработчик не должен напрямую читать или изменять внутренние таблицы SharePoint в SQL Server. Для работы с данными нужно использовать поддерживаемые интерфейсы: SharePoint REST API, CSOM, SPFx API и другие официальные API. Прямое изменение базы может повредить данные и сделать конфигурацию неподдерживаемой Microsoft.

### Роли серверов

В ферме разные серверы могут выполнять разные функции:

- **Front-end** — принимает пользовательские HTTP-запросы и возвращает страницы;
- **Application** — выполняет серверные службы и фоновые процессы;
- **Search** — индексирует содержимое и обрабатывает поисковые запросы;
- **Distributed Cache** — хранит временные данные в памяти;
- **SQL Server** — хранит постоянные данные.

В SharePoint Server используется модель **MinRole**: администратор назначает серверу роль, после чего SharePoint запускает подходящий набор служб.

### Связь SharePoint Server и SPFx

В локальной среде SharePoint Server выступает backend-платформой, а SPFx остаётся frontend-кодом, выполняемым в браузере:

```text
Браузер пользователя
        ↓ загружает страницу и JavaScript
SPFx-компонент
        ↓ вызывает REST API
SharePoint Server
        ↓ читает или изменяет данные
SQL Server
```

Таким образом:

- SharePoint Server хранит данные и выполняет серверную логику;
- SPFx отображает интерфейс и отправляет запросы к API;
- запросы выполняются с учётом прав текущего пользователя.

### Ограничения версий SPFx

Локальные версии SharePoint Server не поддерживают все современные версии SPFx:

- SharePoint Server 2016 — SPFx до версии `1.1.0`;
- SharePoint Server 2019 — SPFx до версии `1.4.1`;
- SharePoint Server Subscription Edition — та же поддержка SPFx, что у Server 2019;
- SharePoint Online — наиболее новая и полная поддержка SPFx.

Поэтому SPFx-проект для локального SharePoint необходимо создавать под версию, которую поддерживает целевой SharePoint Server.

### SharePoint Server Subscription Edition

Современная локальная редакция называется **SharePoint Server Subscription Edition**. Microsoft выпускает для неё ежемесячные обновления и отдельные Feature Updates, но устанавливать и контролировать эти обновления должна сама организация.

> **Кратко:** SharePoint Server — самостоятельно развёрнутый backend и веб-портал SharePoint с серверными службами и базами SQL Server. SPFx — frontend, который загружается с SharePoint Server, выполняется в браузере и обращается к его API.

### Дополнительные источники

- [Обзор ролей MinRole — Microsoft Learn](https://learn.microsoft.com/en-us/sharepoint/install/overview-of-minrole-server-roles-in-sharepoint-server)
- [Поддерживаемые платформы расширяемости SPFx — Microsoft Learn](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/supported-extensibility-platforms-overview)
- [Политика обслуживания SharePoint Server Subscription Edition — Microsoft Learn](https://learn.microsoft.com/en-us/sharepoint/product-servicing-policy/faq/faq-subscription-edition)

## 12. Быстрый старт разработки SPFx

Этот раздел описывает подготовку Windows-компьютера для современной разработки под **SharePoint Online** с использованием SPFx `1.22+` и системы сборки **Heft**.

### 12.1. Необходимые инструменты

- доступ к тестовому SharePoint Online tenant;
- Node.js поддерживаемой LTS-версии;
- npm;
- Visual Studio Code;
- Heft — инструмент сборки;
- Yeoman и генератор SharePoint Framework.

Для SPFx `1.22.x` и `1.23.x` используется **Node.js 22 LTS**. Node.js 24 для этих версий SPFx не поддерживается.

Версию Node.js всегда нужно сверять с таблицей совместимости выбранной версии SPFx:

- [SPFx Platform & Toolchain Compatibility Reference](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/compatibility)

### 12.2. Проверка Node.js и npm

В CMD:

```cmd
node --version
npm.cmd --version
where node
```

Ожидаемый пример:

```text
v22.x.x
C:\Program Files\nodejs\node.exe
```

Если отображается Node.js 24, необходимо удалить его, установить Node.js 22 LTS x64 и открыть новое окно терминала.

Если после переустановки показывается старая версия, команда `where node` поможет обнаружить несколько установок Node.js в `PATH`.

### 12.3. Установка инструментов SPFx

Установка глобальных инструментов:

```cmd
npm.cmd install --global @rushstack/heft yo @microsoft/generator-sharepoint@latest
```

Здесь устанавливаются:

- `@rushstack/heft` — запуск сборки, локальной разработки и упаковки;
- `yo` — Yeoman, инструмент генерации проектов;
- `@microsoft/generator-sharepoint` — шаблоны и структура SPFx-проекта.

### 12.4. Ошибка PowerShell: running scripts is disabled

PowerShell может блокировать файл `npm.ps1`:

```text
npm.ps1 cannot be loaded because running scripts is disabled on this system
```

Самое простое решение — использовать исполняемый файл `npm.cmd`, не изменяя системную политику:

```powershell
npm.cmd install --global @rushstack/heft yo @microsoft/generator-sharepoint@latest
```

Аналогично при необходимости можно использовать:

```powershell
yo.cmd
heft.cmd
npx.cmd
```

Альтернативный вариант — разрешить локальные PowerShell-скрипты только для текущего пользователя:

```powershell
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
```

Проверка действующих политик:

```powershell
Get-ExecutionPolicy -List
```

Изменять Execution Policy необязательно: для начала достаточно команд с окончанием `.cmd`.

### 12.5. Проверка установленных инструментов

```cmd
node --version
npm.cmd --version
yo.cmd --version
npm.cmd list --global @rushstack/heft --depth=0
npm.cmd list --global @microsoft/generator-sharepoint --depth=0
```

Запуск `heft.cmd --version` вне проекта может завершиться сообщением:

```text
No package.json file found. Are you in a project folder?
```

Это не означает, что Heft установлен неправильно. Он пытается найти конфигурацию проекта и ожидает файл `package.json`. Глобальную установку Heft лучше проверять через `npm.cmd list`, а команды сборки выполнять из корня созданного проекта.

### 12.6. Создание первого проекта

Не следует создавать проект в `C:\Windows\System32`. Лучше использовать отдельную рабочую папку:

```cmd
cd /d C:\Workspace\sharepointcore
mkdir hello-spfx
cd hello-spfx
```

Запуск генератора:

```cmd
yo.cmd @microsoft/sharepoint
```

Для первого учебного проекта можно выбрать:

```text
Solution name: hello-spfx
Target: SharePoint Online only
Place files: current folder
Tenant-wide deployment: No
Component type: WebPart
Web part name: HelloWorld
Template: React
```

После генерации в папке проекта появятся `package.json`, `src`, `config`, `sharepoint` и другие файлы.

### 12.7. Сертификат локальной разработки

Из корня проекта:

```cmd
heft.cmd trust-dev-cert
```

Сертификат нужен, чтобы SharePoint мог безопасно загружать локальные JavaScript- и CSS-файлы через HTTPS.

### 12.8. Dev-режим

Запуск локального сервера разработки:

```cmd
heft.cmd start
```

В dev-режиме Heft:

- компилирует TypeScript и SCSS;
- создаёт debug-bundle;
- запускает локальный HTTPS-сервер;
- отслеживает изменения исходных файлов;
- повторно собирает проект после сохранения изменений.

SharePoint остаётся в Microsoft 365, а JavaScript и CSS временно загружаются с компьютера разработчика.

Остановить сервер можно сочетанием `Ctrl+C`.

Проверочная development-сборка:

```cmd
heft.cmd build
```

Чистая development-сборка:

```cmd
heft.cmd build --clean
```

### 12.9. Production-сборка

Оптимизированная сборка:

```cmd
heft.cmd build --clean --production
```

Создание SharePoint-пакета:

```cmd
heft.cmd package-solution --production
```

Результат создаётся в каталоге:

```text
sharepoint/solution/*.sppkg
```

Файл `.sppkg` загружается в SharePoint **App Catalog**, после чего решение можно развернуть и добавить на сайт.

```text
TypeScript/React
      ↓
heft build --clean --production
      ↓
Оптимизированные JavaScript и CSS
      ↓
heft package-solution --production
      ↓
Пакет .sppkg
      ↓
SharePoint App Catalog
```

Для production всегда следует использовать сборку с флагом `--production`, а не debug-сборку.

### 12.10. Два значения DEV и PROD

**Режим сборки:**

```text
DEV  → heft start / heft build
PROD → heft build --production / heft package-solution --production
```

**Среда SharePoint:**

```text
DEV site или tenant  → разработка
TEST site или tenant → тестирование
PROD tenant          → реальные пользователи
```

SPFx предоставляет debug- и production-сборку, но не создаёт DEV, TEST и PROD tenant автоматически. Раздельные среды организует сама компания.

### 12.11. Рекомендуемый первый учебный проект

Полезный вариант — Web Part «Мои задачи»:

1. Создать простую HelloWorld Web Part.
2. Изменить React-разметку и SCSS-стили.
3. Добавить свойства через Property Pane.
4. Создать SharePoint List `Tasks`.
5. Прочитать элементы через `SPHttpClient`.
6. Показать загрузку и обработку ошибок.
7. Добавить создание и изменение задач.
8. Собрать `.sppkg` и развернуть его через тестовый App Catalog.
