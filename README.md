# Конвертер Markdown

[![MIT License](https://img.shields.io/badge/License-MIT-yellow.svg)](https://choosealicense.com/licenses/mit/) [![version](https://img.shields.io/badge/version-1.0-blue)](https://img.shields.io/badge/version-1.0-blue)


**Markdown мультиконвертер** — это удобное десктопное приложение на Python и Flet, позволяющее массово конвертировать `.md` файлы в форматы **Word (.docx)** и **Excel (.xlsx)**. Поддерживает выбор нескольких файлов, автоматическую обработку заголовков и метаданных, а также гибкое указание выходной директории.


## 💡 Возможности

- ✅ Конвертация Markdown → DOCX (с сохранением структуры, таблиц, списков, заголовков и метаданных)
- ✅ Конвертация Markdown → Excel:
  - Каждая строка — это один `.md` файл
  - Каждому заголовку второго уровня (`##`) соответствует свой столбец
- ✅ Поддержка множественного выбора файлов
- ✅ Указание выходной директории (по умолчанию — папка `Документы`)
- ✅ Графический интерфейс на базе [Flet](https://flet.dev)



## 📦 Установка

1. Клонируй репозиторий:

```bash
git clone https://github.com/yourusername/markdown-multiconverter.git
cd markdown-multiconverter
```

2. Установи зависимости:

```bash
pip install -r requirements.txt
```

### `requirements.txt`

```text
flet
markdown2
beautifulsoup4
python-docx
pyyaml
pandas
openpyxl
```

3. Запусти приложение:

```bash
python ConvertMD.py
```


## 🖥 Использование

- Выбери `.md` файлы (можно несколько)
- Укажи выходную директорию (опционально)
- Выбери формат: `DOCX` или `Excel`
- Нажми кнопку **Конвертировать**

## 📁 Структура Excel-файла

При экспорте в Excel создаётся таблица:
- Первый столбец — имя файла
- Остальные — содержимое из заголовков `## Заголовок` в Markdown-файлах

## 🛠 Требования

- Python 3.11+
- ОС: Windows / macOS / Linux


## Структура проекта

```
markdown-мультиконвертер/
├── ConvertMD.py                   # основной GUI-скрипт
├── README.md                      # описание проекта
├── requirements.txt               # зависимости
├── .gitignore                     # исключения для Git
└── LICENSE                        # MIT-лицензия
```



## Разработка

Если вы хотите внести изменения в проект, используйте следующие команды для работы с Git:

Клонировать репозиторий:
```bash
git clone ...
```

Создать новую ветку для разработки:
```bash
git checkout -b new-feature
```
После внесения изменений сделайте коммит:
```bash
git commit -am "Добавлена новая фича"
```
Отправьте изменения в репозиторий:
```bash
git push origin new-feature
```


## Автор

👤 **SCaeR42@SpaceCoding**

* Сайт: [spacecoding.net](https://spacecoding.net/utm_source=github)
* Github: [@SCaeR42](https://github.com/SCaeR42)

## Покажите свою поддержку

Поставьте ⭐️, если этот проект вам помог!

## 🔖 Лицензия

Авторское право (C) 2025 [spacecoding.net](https://spacecoding.net/utm_source=github)

Лицензия [MIT](https://choosealicense.com/licenses/mit/).

Разработано с ❤️ на Python
