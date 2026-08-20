# Lead Parsers — Avito, Яндекс.Карты, HH.ru, Rusprofile

Четыре десктопных парсера для Windows с установкой в один клик. Собирают лиды и компании
в CSV / JSON / Excel:

| Парсер | Что собирает |
|---|---|
| **Avito Parser** | объявления Авито: название, цена, ссылка, локация, продавец, описание |
| **Yandex Maps Lead Parser** | организации с Яндекс.Карт: телефоны, сайты, адреса, соцсети (есть режим «Вся Россия» — обход 65 городов) |
| **HH.ru Parser** | вакансии hh.ru: должность, компания, зарплата, контакты (где открыты) |
| **Rusprofile Parser** | компании с rusprofile.ru по кодам ОКВЭД: название, ИНН, ОГРН, статус, адрес |

## Скачать

**[⬇ Скачать архив со всеми файлами](https://github.com/pastilab2017-code/lead-parsers/archive/refs/heads/main.zip)**

(или зелёная кнопка **Code → Download ZIP** вверху страницы — аккаунт GitHub не нужен)

## Установка (Windows 10/11)

1. Установите [Python 3.10+](https://www.python.org/downloads/windows/) — при установке
   обязательно отметьте галочку **«Add python.exe to PATH»**.
2. Установите [Google Chrome](https://www.google.com/chrome/), если его нет.
3. Распакуйте скачанный архив в постоянную папку, например `C:\Parsers`.
4. Запустите `install.bat` двойным кликом и дождитесь `=== Done! ===`.

Установщик сам создаст окружение Python, поставит зависимости и добавит четыре ярлыка
на рабочий стол. Подробности и ответы на вопросы — в файле [УСТАНОВКА.md](УСТАНОВКА.md).

## Запуск

Ярлыки на рабочем столе: **Avito Parser**, **Yandex Maps Lead Parser**, **HH.ru Parser**,
**Rusprofile Parser**. Введите запрос и город, нажмите «Запустить сбор» — Excel появится
в папке `results`.

Консольный режим для автоматизации:

```powershell
.\.venv\Scripts\python.exe .\avito_parser_app.py --cli --query "пастила" --city "Москва" --limit 100
.\.venv\Scripts\python.exe .\yandex_maps_parser_app.py --cli --query "кофейня" --city "Вся Россия" --limit 300
.\.venv\Scripts\python.exe .\hh_parser_app.py --cli --query "менеджер по закупкам" --city "Москва" --limit 100 --deep
.\.venv\Scripts\python.exe .\tools\rusprofile-scrapper\run_rusprofile_local.py 463900 472900 --max-items 300 --output-dir .\results\rusprofile
```

## Примечания

- Парсеры используют открытые страницы сайтов. Если сайт покажет капчу, решите её в
  открывшемся окне Chrome — сбор продолжится автоматически.
- Профиль Chrome сохраняется между запусками: если войти в аккаунт Авито или hh.ru один раз,
  вход сохранится и данных будет больше.
- Перед использованием собранных контактов для рассылок соблюдайте требования
  законодательства о рекламе и персональных данных.
