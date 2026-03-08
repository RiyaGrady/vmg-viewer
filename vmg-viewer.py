#!/usr/bin/env python3
import pathlib
import datetime
import re
import quopri
import tkinter as tk
from tkinter import ttk, scrolledtext, filedialog, messagebox


def decode_quoted_printable(s):
    return quopri.decodestring(s).decode('utf-8', errors='replace')

def read_text(path):
    b = path.read_bytes()
    for enc in ('utf-8', 'cp1251', 'latin-1'):
        try:
            return b.decode(enc)
        except Exception:
            continue
    return b.decode('utf-8', errors='replace')

def parse_vmg(path):
    with open(path, 'rb') as f:
        data = f.read()
    text = read_text(path)
    tel_match = re.search(r'^TEL[^:]*:(.+)$', text, flags=re.MULTILINE)
    tel = tel_match.group(1).strip() if tel_match else 'Unknown'
    date_match = re.search(r'^(Date|REV):\s*([0-9T:\.\-Z]+)', text, flags=re.MULTILINE)
    date = date_match.group(2).strip() if date_match else ''
    texts = []
    for b in re.finditer(r'BEGIN:VBODY(.*?)END:VBODY', text, flags=re.DOTALL):
        body = b.group(1)
        for m in re.finditer(r'(?m)^TEXT[^:]*:(.*(?:\n[ \t].*)*)', body):
            raw = m.group(1)
            raw = re.sub(r'\n[ \t]+', '', raw)
            decoded = decode_quoted_printable(raw)
            texts.append((date, decoded))
    if not texts:
        for m in re.finditer(r'(?m)^TEXT[^:]*:(.*)$', text):
            raw = m.group(1).rstrip()
            raw = re.sub(r'\n[ \t]+', '', raw)
            texts.append((date, decode_quoted_printable(raw)))
    return tel, texts

def export_messages_to_json(contacts: dict, out_path: pathlib.Path):
    """
    Экспортирует сообщения в JSON.

    Параметры:
    - contacts: словарь {tel: [(date_raw, text, source_file), ...], ...}
      где date_raw — исходная строка даты (или ''), text — декодированный текст.
    - out_path: pathlib.Path файла вывода (например, Path('export.json')).

    Формат JSON:
    {
      "<tel>": [
        { "date_raw": "...", "date_iso": "...", "text": "...", "source_file": "..." },
        ...
      ],
      ...
    }

    date_iso — дата в формате ISO 8601 (YYYY-MM-DDTHH:MM:SS) если возможно распарсить,
    иначе пустая строка.
    """
    import json
    import re
    import datetime
    import pathlib

    # регулярка для даты вида 2018.4.27.5.46.16 или 2018.04.27.05.46.16
    date_re = re.compile(r'^\s*([0-9]{4})\.([0-9]{1,2})\.([0-9]{1,2})\.([0-9]{1,2})\.([0-9]{1,2})\.([0-9]{1,2})')

    def to_iso(datestr: str) -> str:
        if not datestr:
            return ''
        m = date_re.match(datestr)
        if not m:
            return ''
        y, mo, d, H, M, S = map(int, m.groups())
        try:
            return datetime.datetime(y, mo, d, H, M, S).isoformat()
        except Exception:
            return ''

    out = {}
    for tel, msgs in contacts.items():
        entries = []
        for date_raw, text, source in msgs:
            entries.append({
                "date_raw": date_raw or "",
                "date_iso": to_iso(date_raw),
                "text": text,
                "source_file": source
            })
        out[str(tel)] = entries

    # Запись в файл с отступами, ensure_ascii=False чтобы русские символы читались нормально
    out_path = pathlib.Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with out_path.open('w', encoding='utf-8') as f:
        json.dump(out, f, ensure_ascii=False, indent=2)


def format_date(datestr):
    if not datestr:
        return ''
    m = re.match(r'^\s*([0-9]{4})\.([0-9]{1,2})\.([0-9]{1,2})\.([0-9]{1,2})\.([0-9]{1,2})\.([0-9]{1,2})', datestr)
    if not m:
        return ''
    y, mo, d, H, M, S = map(int, m.groups())
    try:
        dt = datetime.datetime(y, mo, d, H, M, S)
    except ValueError:
        return ''
    return f'[{dt.hour}:{dt.minute:02d} {dt.day:02d}.{dt.month:02d}.{dt.year}]'

def collect_messages(folder: pathlib.Path):
    import pathlib as _pathlib, re as _re, datetime as _datetime
    script_dir = folder
    contacts = {}
    date_re = _re.compile(r'^\s*([0-9]{4})\.([0-9]{1,2})\.([0-9]{1,2})\.([0-9]{1,2})\.([0-9]{1,2})\.([0-9]{1,2})')
    for path in sorted(script_dir.glob('*.vmg')):
        try:
            tel, texts = parse_vmg(path)
            for date_str, txt in texts:
                dt = None
                if date_str:
                    m = date_re.match(date_str)
                    if m:
                        try:
                            y, mo, d, H, M, S = map(int, m.groups())
                            dt = _datetime.datetime(y, mo, d, H, M, S)
                        except Exception:
                            dt = None
                contacts.setdefault(tel, []).append((dt, date_str or '', txt, path.name))
        except Exception:
            continue
    for tel in contacts:
        contacts[tel].sort(key=lambda x: (x[0] is None, x[0] or datetime.datetime.min, x[1]))
    result = {tel: [(item[1], item[2], item[3]) for item in contacts[tel]] for tel in contacts}
    return result

class App(ttk.Frame):
    def __init__(self, root):
        super().__init__(root)
        root.title('VMG Messages')
        root.geometry('900x600')
        self.pack(fill='both', expand=True)
        self.current_folder = pathlib.Path(__file__).resolve().parent
        self.contacts = collect_messages(self.current_folder)
        self.create_widgets()

    def create_widgets(self):
        topbar = ttk.Frame(self)
        topbar.pack(fill='x', padx=8, pady=6)

        self.folder_label = ttk.Label(topbar, text=str(self.current_folder))
        self.folder_label.pack(side='left', padx=(0,8))

        choose_btn = ttk.Button(topbar, text='Выбрать папку с sms', command=self.choose_folder)
        choose_btn.pack(side='left')

        export_btn = ttk.Button(topbar, text='Экспорт в JSON', command=self.export_json)
        export_btn.pack(side='left', padx=(6, 0))

        paned = ttk.Panedwindow(self, orient='horizontal')
        paned.pack(fill='both', expand=True)

        left = ttk.Frame(paned, width=250)
        right = ttk.Frame(paned)
        paned.add(left, weight=1)
        paned.add(right, weight=4)

        ttk.Label(left, text='Contacts').pack(anchor='nw', padx=8, pady=(8,0))
        self.listbox = tk.Listbox(left)
        self.listbox.pack(fill='both', expand=True, padx=8, pady=8)
        self.populate_contacts()
        self.listbox.bind('<<ListboxSelect>>', self.on_select)

        ttk.Label(right, text='Conversation').pack(anchor='nw', padx=8, pady=(8,0))
        self.txt = scrolledtext.ScrolledText(right, wrap='word', state='disabled')
        self.txt.pack(fill='both', expand=True, padx=8, pady=8)

    def populate_contacts(self):
        self.listbox.delete(0, 'end')
        for tel in sorted(self.contacts.keys()):
            self.listbox.insert('end', f'{tel} ({len(self.contacts[tel])})')

    def choose_folder(self):
        folder = filedialog.askdirectory(initialdir=str(self.current_folder))
        if not folder:
            return
        self.current_folder = pathlib.Path(folder)
        self.folder_label.config(text=str(self.current_folder))
        self.contacts = collect_messages(self.current_folder)
        self.populate_contacts()
        self.txt.configure(state='normal')
        self.txt.delete('1.0', 'end')
        self.txt.configure(state='disabled')

    def on_select(self, event):
        sel = self.listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        key = sorted(self.contacts.keys())[idx]
        msgs = self.contacts[key]
        self.txt.configure(state='normal')
        self.txt.delete('1.0', 'end')
        for date, text, fname in msgs:
            hdr = format_date(date)
            if hdr:
                self.txt.insert('end', hdr + '\n')
            self.txt.insert('end', text + '\n\n')
        self.txt.configure(state='disabled')

    def export_json(self):
        # спросить файл для сохранения
        path = filedialog.asksaveasfilename(
            initialdir=str(self.current_folder),
            initialfile = 'export.json',
            defaultextension = '.json',
            filetypes = [('JSON files', '\*.json'), ('All files', '\*.\*')]
        )
        if not path:
            return
        out_path = pathlib.Path(path)
        try:
            export_messages_to_json(self.contacts, out_path)
            messagebox.showinfo('Экспорт завершён', f'Экспортировано в: {out_path}')
        except Exception as e:
            messagebox.showerror('Ошибка экспорта', str(e))

if __name__ == '__main__':
    root = tk.Tk()
    app = App(root)
    root.mainloop()
