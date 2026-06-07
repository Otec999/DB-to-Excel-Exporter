# SQL to Excel Exporter

GUI-приложение для экспорта таблиц из SQL базы данных в Excel файлы.

## Возможности

- Подключение к SQL Server, MySQL, PostgreSQL через connection string
- Автоматическая загрузка списка таблиц из базы данных
- Экспорт выбранной таблицы в Excel (.xlsx)
- Экспорт всех таблиц разом в отдельные Excel файлы
- Прогресс-бар и лог операций

## Установка

```bash
pip install -r requirements.txt
```

## Запуск

```bash
python main.py
```

## Примеры connection strings

| СУБД | Connection String |
|------|-------------------|
| SQL Server | `mssql+pyodbc://server/db?trusted_connection=yes&driver=ODBC+Driver+17+for+SQL+Server` |
| MySQL | `mysql+pymysql://root:password@localhost/your_database` |
| PostgreSQL | `postgresql://postgres:password@localhost:5432/your_database` |
