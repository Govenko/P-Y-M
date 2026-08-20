
| Парсер | Что собирает |
| **Avito Parser** | объявления Авито: название, цена, ссылка, локация, продавец, описание |
| **Yandex Maps Lead Parser** | организации с Яндекс.Карт: телефоны, сайты, адреса, соцсети (есть режим «Вся Россия» — обход 65 городов) |
| **HH.ru Parser** | вакансии hh.ru: должность, компания, зарплата, контакты (где открыты) |
| **Rusprofile Parser** | компании с rusprofile.ru по кодам ОКВЭД: название, ИНН, ОГРН, статус, адрес |


Консольный режим для автоматизации:

```powershell
.\.venv\Scripts\python.exe .\avito_parser_app.py --cli --query "пастила" --city "Москва" --limit 100
.\.venv\Scripts\python.exe .\yandex_maps_parser_app.py --cli --query "кофейня" --city "Вся Россия" --limit 300
.\.venv\Scripts\python.exe .\hh_parser_app.py --cli --query "менеджер по закупкам" --city "Москва" --limit 100 --deep
.\.venv\Scripts\python.exe .\tools\rusprofile-scrapper\run_rusprofile_local.py 463900 472900 --max-items 300 --output-dir .\results\rusprofile
```
