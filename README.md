# CEE Tender Intelligence v5

Новий Render-ready MVP.

## Конектори
- Prozorro: офіційний public API для України.
- TED Search API v3: офіційний відкритий API без ключа. Ринки PL, RO, CZ, LT, LV, EE, SK, SI, HR, BG, HU, DE, FR, IT та ES отримуються через TED country filters.
- BZP/e-Zamówienia: офіційний польський read endpoint. Польща отримується одночасно з BZP та TED.

Національні портали без підтвердженого стабільного відкритого API навмисно не скрейпляться. Система не обходить CAPTCHA, авторизацію або умови сайтів.

## Розгортання
1. Видаліть старі файли з GitHub-репозиторію та завантажте весь вміст цього архіву зі збереженням папок.
2. Render має побачити Dockerfile. Health check: `/health`.
3. Змінні наведені в `.env.example`; `render.yaml` додає основні автоматично.
4. Натисніть `Оновити дані`. Статус кожного конектора та точна помилка будуть у верхній частині вебпанелі та в Render Logs.

## Telegram
Polling/getUpdates не використовується, тому конфлікту двох екземплярів немає. У профілі KAM потрібно вказати власний Telegram Chat ID; бот використовується тільки для вихідних дайджестів.

## Render Free
`DB_PATH=/tmp/cee_tenders.db` є тестовим сховищем і може очищатися після restart/redeploy. Для production потрібні persistent disk або PostgreSQL.
