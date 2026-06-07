<div align="center">

# 📊 SQL to Excel Exporter

**Экспорт таблиц из SQL базы данных в Excel | Export SQL tables to Excel**

![Python](https://img.shields.io/badge/Python-3.10+-blue?logo=python)
![PyQt6](https://img.shields.io/badge/GUI-PyQt6-green)
![License](https://img.shields.io/badge/License-MIT-orange)

<br>

</div>

---

# 🇷🇺 Русский

## 📝 Описание

**SQL to Excel Exporter** — это десктопное GUI-приложение, которое подключается к SQL базе данных, загружает список таблиц и позволяет экспортировать их в формат Excel (`.xlsx`). Работает с **SQL Server**, **MySQL** и **PostgreSQL**.

## ✨ Возможности

| Функция | Описание |
|---------|----------|
| 🔌 **Подключение к БД** | SQL Server, MySQL, PostgreSQL через connection string |
| 📋 **Список таблиц** | Автоматическая загрузка всех таблиц из базы |
| 📤 **Экспорт таблицы** | Сохранение выбранной таблицы в `.xlsx` |
| 📦 **Экспорт всех таблиц** | Сохранение всех таблиц в отдельные Excel-файлы |
| ⚙️ **Настройки** | Включение/отключение индекса строк, автозапуск файла |
| 📜 **Лог операций** | Подробный лог с временными метками |
| 🎨 **Интерфейс** | Интуитивный UI на PyQt6 |

## 🚀 Установка и запуск

```bash
# Клонирование
git clone https://github.com/Otec999/DB-to-Excel-Exporter.git
cd DB-to-Excel-Exporter

# Установка зависимостей
pip install -r requirements.txt

# Запуск
python main.py
```

## 🔌 Примеры подключения

| СУБД | Connection String |
|------|-------------------|
| **SQL Server** | `mssql+pyodbc://server/db?trusted_connection=yes&driver=ODBC+Driver+17+for+SQL+Server` |
| **MySQL** | `mysql+pymysql://root:password@localhost/your_database` |
| **PostgreSQL** | `postgresql://postgres:password@localhost:5432/your_database` |

> **💡 Совет:** В приложении есть кнопки быстрой подстановки шаблонов для популярных СУБД.

## 📸 Как использовать

1. Введите **connection string** (или нажмите кнопку быстрой подстановки)
2. Нажмите **"Connect & Load Tables"** — загрузятся все таблицы
3. Выберите таблицу из выпадающего списка
4. Нажмите **"Export Selected Table"** и укажите путь
5. Либо нажмите **"Export ALL Tables"** для экспорта всего сразу

## 📦 Зависимости

```
PyQt6     — графический интерфейс
pandas    — работа с данными
openpyxl  — создание Excel-файлов
sqlalchemy — подключение к БД
pyodbc    — драйвер для SQL Server
```

---

# 🇬🇧 English

## 📝 Description

**SQL to Excel Exporter** is a desktop GUI application that connects to SQL databases, loads table lists, and exports them to Excel (`.xlsx`) format. Supports **SQL Server**, **MySQL**, and **PostgreSQL**.

## ✨ Features

| Feature | Description |
|---------|-------------|
| 🔌 **DB Connection** | SQL Server, MySQL, PostgreSQL via connection string |
| 📋 **Table List** | Auto-fetches all tables from database |
| 📤 **Single Export** | Export selected table to `.xlsx` |
| 📦 **Bulk Export** | Export all tables to separate Excel files |
| ⚙️ **Options** | Include row index, auto-open file after export |
| 📜 **Activity Log** | Detailed log with timestamps |
| 🎨 **GUI** | Clean PyQt6 interface |

## 🚀 Installation & Run

```bash
# Clone
git clone https://github.com/Otec999/DB-to-Excel-Exporter.git
cd DB-to-Excel-Exporter

# Install dependencies
pip install -r requirements.txt

# Run
python main.py
```

## 🔌 Connection Examples

| DB | Connection String |
|----|-------------------|
| **SQL Server** | `mssql+pyodbc://server/db?trusted_connection=yes&driver=ODBC+Driver+17+for+SQL+Server` |
| **MySQL** | `mysql+pymysql://root:password@localhost/your_database` |
| **PostgreSQL** | `postgresql://postgres:password@localhost:5432/your_database` |

> **💡 Tip:** The app has quick-template buttons for popular databases.

## 📸 How to Use

1. Enter a **connection string** (or click a template button)
2. Click **"Connect & Load Tables"** — all tables will load
3. Select a table from the dropdown
4. Click **"Export Selected Table"** and choose save location
5. Or click **"Export ALL Tables"** to export everything at once

## 📦 Dependencies

```
PyQt6     — graphical user interface
pandas    — data manipulation
openpyxl  — Excel file creation
sqlalchemy — database connection
pyodbc    — SQL Server driver
```

---

<div align="center">

## 👨‍💻 Author

**Otec999**

[![GitHub](https://img.shields.io/badge/GitHub-Otec999-181717?logo=github)](https://github.com/Otec999)

---

⭐ **If you like this project, give it a star!** ⭐

</div>
