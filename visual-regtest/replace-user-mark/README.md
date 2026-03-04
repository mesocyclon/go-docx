# replace-user-mark

Заполняет пользовательский шаблон содержимым из файлов-меток.

## Структура

```
replace-user-mark/
  in/
    template.docx          ← шаблон с метками формата (<имя>)
    mark/
      invoice_header.docx  ← содержимое для (<invoice_header>)
      footer_info.docx     ← содержимое для (<footer_info>)
      ...                  ← любое количество меток
  out/
    result.docx            ← заполненный шаблон (генерируется)
    manifest.json          ← отчёт о заменах
```

## Правило маппинга

Имя файла в `mark/` (без `.docx`) → метка `(<имя>)` в шаблоне.

| Файл в `mark/`         | Заменяет метку           |
|-------------------------|--------------------------|
| `mark1.docx`           | `(<mark1>)`              |
| `header_content.docx`  | `(<header_content>)`     |
| `таблица.docx`         | `(<таблица>)`            |

## Запуск

```bash
cd visual-regtest
make replace-user-mark
```

Или напрямую:

```bash
go run ./visual-regtest/replace-user-mark
```

С пользовательскими путями:

```bash
go run ./visual-regtest/replace-user-mark \
  --input /path/to/input \
  --output /path/to/output
```