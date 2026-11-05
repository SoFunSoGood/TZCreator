# templates/

Положите сюда эталонный файл `заявка_шаблон.docx` с Jinja2-плейсхолдерами, например:

- `{{ number }}`, `{{ created }}` — номер и дата заявки
- `{% for c in candidates %}{{ c.full_name }} — {{ c.birth_date }}{% endfor %}` — список кандидатов
