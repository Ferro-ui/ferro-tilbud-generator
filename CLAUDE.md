# Ferro Tilbud Generator — Контекст проекту

## Що це таке

Інструмент для компанії **Ferro Stålentreprenør AS** (Норвегія) — автоматично генерує комерційні пропозиції (tilbud) з вхідних файлів.

Власник: Vegard Landsdal, vegard@ferrostal.no, Tlf: 95 83 71 23  
Адреса: Ringsevja 3, 3830 Ulefoss

---

## Файли в цій папці

| Файл | Опис |
|------|------|
| `tilbud_html.html` | Ручний HTML-генератор tilbud (React, без сервера) |
| `tilbud_generator.py` | **Головна програма** — читає файли → Claude API → генерує .docx + .html |

---

## Що робить `tilbud_generator.py`

1. Приймає файли: `.pdf`, `.xlsx`, `.xls`, `.jpg`, `.jpeg`, `.png`, `.docx`
2. Витягує текст і зображення з кожного файлу
3. Відправляє все до Claude (claude-sonnet-4-6) через Vision API
4. Claude повертає структурований JSON з даними проекту
5. Генерує три файли в папці з вхідними файлами:
   - `Tilbud_[Назва]_[дата].docx` — готовий Word-документ
   - `Tilbud_[Назва]_[дата]_prefilled.html` — HTML-генератор з заповненими полями
   - `Tilbud_[Назва]_[дата]_extracted.json` — сирі дані від Claude

---

## Структура tilbud

Стандартний tilbud містить:
- Заголовок + дата + отримувач (kunde, kontakt, adresse)
- PRISSAMMENDRAG (sum eks. mva → 25% mva → sum ink. mva)
- ARBEIDSBESKRIVELSE по 18 секціях:
  bygget, grunn, betong, staal, yttervegg, innervegg, himling, gulv,
  trapper, vinduer, porter, tak, inventar, ventilasjon, brann,
  prosjektering, mengder, rigg
- OPSJONER (необов'язково)
- GENERELLE FORUTSETNINGER (стандартний текст)

---

## Запуск

```bash
# Встановити API ключ
export ANTHROPIC_API_KEY='sk-ant-...'

# З окремими файлами
python3 tilbud_generator.py креслення.pdf кошторис.xlsx фото.jpg

# Або вся папка
python3 tilbud_generator.py --folder /шлях/до/папки
```

---

## Встановлені залежності

```
anthropic      # Claude API
pymupdf        # читання PDF (fitz)
openpyxl       # читання Excel
pillow         # обробка зображень
python-docx    # читання і генерація .docx
```

Встановити: `pip3 install anthropic pymupdf openpyxl pillow python-docx`

---

## Що зроблено / що можна покращити

### Зроблено
- [x] Читання PDF (текст + рендер сторінок як зображення)
- [x] Читання Excel (всі аркуші)
- [x] Читання зображень (JPEG, PNG) через Claude Vision
- [x] Читання .docx
- [x] Генерація .docx tilbud
- [x] Генерація pre-filled HTML
- [x] Масштабування зображень до ліміту Claude Vision (1568px)

### Можливі покращення
- [ ] Додати підтримку `.dwg` / `.dxf` (AutoCAD креслення)
- [ ] Веб-інтерфейс (drag & drop файлів)
- [ ] Підтримка кількох мов (bokmål / nynorsk)
- [ ] Шаблони для різних типів будівель
- [ ] Автоматичний розрахунок ціни з Excel-кошторису

---

## Тестовий файл

`~/Downloads/Tilbud_Lagerbygg_Steinsholt_2026-04-23.docx` — приклад готового tilbud:
- Проект: Lagerbygg Steinsholt, 11×20m, gesimshøyde 5,5m
- Kunde: Reis AS, v/ Eskil Støvland, Hegdalringen 6B, 3261 Larvik
- Opsjoner: 51 100,- eks. mva (motor, radiomottaker, fjernkontroll, fotocelle)
