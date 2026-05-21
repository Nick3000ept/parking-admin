# CLAUDE.md — операционная шпаргалка

Контекст для будущих сессий Claude в этой папке. Не дублирует [TZ.md](TZ.md) — содержит только операционные детали разработки и деплоя.

## Проект

Веб-приложение для учёта работ на паркинге СБ3. Подробности в [TZ.md](TZ.md). Pre-deploy проверки — в [DEPLOY_CHECKLIST.md](DEPLOY_CHECKLIST.md).

**Состояние**: развёрнуто на боевом контуре, работает. Доступ для подрядчиков по токенам.

## Подключения

| Сущность | Значение |
|---|---|
| Spreadsheet ID | `1zB1IQJlla93txI9o-sH-mxD3CeeO0Ugkv8mHPDzcqiI` |
| Apps Script ID | `1gJBrX63FdnXLHuqX8jcAN57OVoV3l9uvTGUsYyq4bVrNfC3sf4tJ7apD` |
| GAS Web App URL (production) | `https://script.google.com/macros/s/AKfycbxX2skPihWXv4Wu8SV4Ehj_IybE4PjfzsnWG9b9dGjEwCtjkYev9WMEtQNicP0-N71xaA/exec` |
| Production deployment ID | `AKfycbxX2skPihWXv4Wu8SV4Ehj_IybE4PjfzsnWG9b9dGjEwCtjkYev9WMEtQNicP0-N71xaA` |
| GAS аккаунт (clasp) | `kuzkin@acons.group` |
| GitHub репо | `https://github.com/Nick3000ept/parking-admin` |
| GitHub Pages URL | `https://nick3000ept.github.io/parking-admin/` |

## Листы в Google Sheets

| Лист | Назначение |
|---|---|
| `Главный` | Реестр помещений (A–L) + парные колонки работ (M+). **A–L read-only для GAS** — `assertWritableColumn(col)` падает с ошибкой если col ≤ 12 |
| `Ведомость_работ` | Справочник работ: Полное название / Поверхность / Название для площадки / ИД_работы |
| `Ведомость_подрядчиков` | Подрядчик / ИД_подрядчик / Статус / Токен. Зарезервированные `ИД_подрядчик`: `admin` (роль Админ), `viewer` (роль Наблюдатель) |
| `Типы_помещений` | Тип_помещения_детально / Тип_помещения_панель / Цвет. Связь с реестром через `Главный.Наименование = Типы_помещений.A` |

## Деплой

### Фронт (index.html)
```
cd "c:/Users/User/YandexDisk/VS_hub/СБ3_Паркинг_Админ"
git add index.html
git commit -m "..."
git push
```
GitHub Pages подхватит через ~1 минуту.

### Бэк (script.gs)
```
cd "c:/Users/User/YandexDisk/VS_hub/СБ3_Паркинг_Админ"
clasp push --force
clasp update-deployment --description "vN <что изменилось>" AKfycbxX2skPihWXv4Wu8SV4Ehj_IybE4PjfzsnWG9b9dGjEwCtjkYev9WMEtQNicP0-N71xaA
```
**Использую `update-deployment`, не `deploy`** — иначе создаётся НОВЫЙ deployment с новым URL, и фронт перестанет работать.

### Тестирование локально (без GitHub Pages)
```
cd "c:/Users/User/YandexDisk/VS_hub/СБ3_Паркинг_Админ"
python -m http.server 8000
```
Открыть `http://localhost:8000/?t=<токен>`.

## Тестовые токены

В лиcте `Ведомость_подрядчиков` (могут устареть, актуальные смотреть в Sheets):
- Овчинников: `81bc3299c7ceb6fb` (был в ТЗ как пример)
- Спектр: `ce3d0facb5187091`
- Альфа Строй: `8ca318c5521b5b66`
- Админ / Наблюдатель: создаются вручную в Sheets с произвольным токеном

## Архитектура коротко

- **Auth**: токен в URL `?t=` → localStorage → бэк проверяет на каждый запрос
- **doGet?action=load** возвращает весь снапшот (~1MB JSON: rooms, works, assignments, tip_panels)
- **doGet?action=ping** — диагностика, список всех листов в файле
- **doPost** — `setDate` / `clearDate` c hint `{row, colDate, colSp}` для fast-path. LockService только на `setValue`.
- **Optimistic UI**: клик в режиме отметки моментально красит карточку, fetch уходит параллельно. Race condition защищён через `cellSeq[num|idRaboty]`.
- **Кэш**: `localStorage.parking_data_<token>` с `version_hash` — мгновенный рендер при перезагрузке + фоновое обновление.

## Известные нюансы

- При переименовании листов в Sheets — синхронно поменять `CONFIG.SHEET_*` в `script.gs` и `clasp push`.
- Чтобы добавить новую работу: строка в `Ведомость_работ` + две колонки в `Главный`. UI подхватит при F5.
- Чтобы добавить подрядчика: строка в `Ведомость_подрядчиков` со свежим токеном. Имя автоматом попадёт в Data Validation колонок `*СП` (если он настроен).
- Зелёный цвет НЕЛЬЗЯ использовать в `Типы_помещений.Цвет` — он зарезервирован за «работа готова».
- На бэке имена сравниваются после `.trim()` — защита от trailing-пробелов в Sheets.

## Что сделано в текущей итерации (history)

- ✅ Базовый каркас + GAS подключение
- ✅ Матрица этаж × корпус с карточками
- ✅ Цветовая индикация по типу помещения
- ✅ Авторизация по токену, очистка URL после первого входа
- ✅ Три роли: Подрядчик / Админ / Наблюдатель (`viewer`)
- ✅ Optimistic UI + fast-path GAS + sequence для race condition
- ✅ Узкий LockService только на write
- ✅ Мультивыбор Ctrl+Click, нижняя панель с площадью и тегами
- ✅ Режим отметки (зелёная шапка, toggle всей карточкой)
- ✅ Прогресс в шапке: помещений / площадь, с цветовой разбивкой готово/остаток
- ✅ Чипсы работ над матрицей, single-select, сгруппированы по поверхности
- ✅ Тёмная панель фильтров (одного цвета с шапкой)
- ✅ GitHub Pages деплой через `git push`

## Что НЕ автоматизировано (намеренно)

- Авто-деплой GAS через GitHub Actions — лишняя инфраструктура; `clasp push` локально достаточно
- Unit-тесты GAS — слабый testing framework в Apps Script; полагаемся на smoke-test через ping
- CI/CD pipeline — `git push` → Pages обновляется самостоятельно
