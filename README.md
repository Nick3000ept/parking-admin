# СБ3 Паркинг — учёт работ

Веб-приложение для учёта выполненных работ по ~2000 помещениям паркинга. Подрядчики отмечают факт готовности по своим назначениям, админ видит общую картину и при необходимости корректирует данные.

## Документы

- **[TZ.md](TZ.md)** — техническое задание (актуальное)
- **[DEPLOY_CHECKLIST.md](DEPLOY_CHECKLIST.md)** — pre-deploy чеклист
- **[CLAUDE.md](CLAUDE.md)** — операционная шпаргалка для AI-сессий

## Стек

- Frontend: `index.html` (vanilla JS, без сборки)
- Backend: Google Apps Script (Web App)
- БД: Google Sheets
- Хостинг: GitHub Pages

## Структура файлов

```
СБ3_Паркинг_Админ/
├── index.html              # фронт (vanilla JS)
├── script.gs               # бэкенд GAS (бэкап; рабочая копия — в Apps Script)
├── README.md               # этот файл
├── TZ.md                   # ТЗ
├── DEPLOY_CHECKLIST.md     # чеклист деплоя
├── CLAUDE.md               # контекст для AI-сессий
└── .gitignore              # исключения для git
```

## Деплой

См. **CLAUDE.md** и **DEPLOY_CHECKLIST.md**.
