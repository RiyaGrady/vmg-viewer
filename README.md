VMG Viewer
VMG Viewer — лёгкая кроссплатформенная утилита на Python для просмотра SMS, экспортированных в формате .vmg (обычно из телефонов Samsung). Проект минималистичен, не требует внешних зависимостей и предназначен для быстрого просмотра, поиска и дальнейшей обработки SMS-архивов.

Запуск:
В репозиторий добавлен предварительно собранный исполняемый файл для Windows: vmg_viewer.exe
Также можно запустить через исходный скрипт: python3 vmg_viewer.py

В приложении нажать «Выбрать папку с sms» и указать каталог с .vmg файлами.

Возможности
Чтение всех .vmg-файлов в указанной папке (по умолчанию — папка со скриптом).
<img width="839" height="746" alt="Screenshot_6" src="https://github.com/user-attachments/assets/c6a97d1d-5ca8-4b4d-a94f-2b04fec83710" />
Корректная обработка QUOTED-PRINTABLE и поддержка кодировок UTF-8, CP1251 и Latin‑1.
Извлечение номера контакта, даты и текста сообщения; поддержка перенесённых (folded) строк.

Простой GUI на tkinter: выбор папки, список контактов с количеством SMS, просмотр переписок в хронологическом порядке с временем и датой.
Зачем это полезно

Ноль внешних зависимостей — работает на любой системе с Python 3.

Удобно для резервного просмотра и анализа сообщений из старых устройств.
Код организован так, чтобы было просто добавить экспорт (CSV/JSON), полнотекстовый поиск или фильтры по дате/номеру.

Интерфейс  


<img width="674" height="473" alt="Screenshot_5" src="https://github.com/user-attachments/assets/767c9672-3bd2-4574-add6-7d5988e22509" />  
Сверху — кнопки указания путь к файлам sms формата vmg и кнопка экспорта sms в JSON.
Слева — список контактов и количество сообщений.
Справа — переписка выбранного контакта, отсортированная по дате, с заголовками времени/даты для каждого сообщения.


Экспорт всех текущих переписок в JSON-файл по контактам  
<img width="851" height="219" alt="Screenshot_7" src="https://github.com/user-attachments/assets/6c314e1d-476b-4254-bb80-362667ef72d1" />  
Экспортирует структуру вида: { "<номер>": [ { "date_raw": "...", "date_iso": "...", "text": "...", "source_file": "..." }, ... ], ... }
Для каждой записи сохраняются: исходная строка даты, дата в формате ISO‑8601 (если удалось распарсить), декодированный текст сообщения и имя исходного .vmg‑файла.
JSON записывается в выбранный пользователем файл с кодировкой UTF-8 и красивым отступом (ensure_ascii=False, indent=2) — русские символы читаются корректно.


ENGLISH
## VMG Viewer  

**VMG Viewer** is a lightweight, cross‑platform Python utility for viewing SMS messages exported in the *.vmg* format (commonly from Samsung phones). The project is minimalist, has no external dependencies, and is intended for quick viewing, searching, and further processing of SMS archives.

### Running the program  

* A pre‑built Windows executable is included in the repository: **`vmg_viewer.exe`**.  
* Alternatively, run the source script:  

```bash
python3 vmg_viewer.py
```

### How to use  

1. In the application click **«Выбрать папку с sms»** and point to the directory containing the *.vmg* files.  
2. The program reads all *.vmg* files in the selected folder (by default, the folder where the script resides).  

### Features  

| Feature | Details |
|---------|---------|
| **File handling** | Reads every *.vmg* file in the chosen directory. |
| **Encoding support** | Correctly processes **QUOTED‑PRINTABLE** and supports UTF‑8, CP1251, and Latin‑1 encodings. |
| **Data extraction** | Retrieves the contact number, date, and message text; handles folded lines. |
| **GUI** | Simple Tkinter interface: folder selection, contact list with message counts, chronological view of conversations with timestamps. |
| **Export** | Exports all current conversations to a JSON file, grouped by contact. |
| **Zero external dependencies** | Works on any system with Python 3. |

### Why it’s useful  

* **No third‑party libraries** – runs everywhere Python 3 is available.  
* Handy for quickly browsing and analysing messages from legacy devices.  
* Code is structured to make it easy to add CSV/JSON export, full‑text search, or date/number filters.

### Interface overview  

* **Top row** – buttons for setting the path to the *.vmg* SMS files and for exporting SMS to JSON.  
* **Left panel** – list of contacts with the number of messages per contact.  
* **Right panel** – the conversation of the selected contact, sorted by date, with a header showing the time/date for each message.

### JSON export  

Press the **«Экспорт sms в JSON»** button to generate a JSON file with the following structure:

```json
{
  "<phone_number>": [
    {
      "date_raw": "...",          // original date string from the .vmg file
      "date_iso": "...",          // ISO‑8601 representation (if parsing succeeded)
      "text": "...",              // decoded message text
      "source_file": "..."        // name of the source .vmg file
    },
    …
  ],
  …
}
```

* The JSON is written to the user‑specified file using **UTF‑8** encoding with pretty indentation (`ensure_ascii=False, indent=2`), so Cyrillic characters are displayed correctly.
