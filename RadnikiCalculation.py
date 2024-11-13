import tkinter as tk
from tkinter.filedialog import askopenfilename, asksaveasfilename
import pandas as pd
from jinja2 import Template
import math
import pdfkit
import os
import customtkinter as ctk
import win32api

xlsx = pd.DataFrame()
df = pd.DataFrame()


def open_file():
    global xlsx
    global list_length

    filepath = askopenfilename(
        filetypes=[("Excel fajlovi", "*.xlsx"), ("Sve fajlovi", "*.*")]
    )
    if not filepath:
        return
    txt_edit.delete("0.0", ctk.END)

    try:
        xlsx = pd.ExcelFile(filepath)
        list_of_sheets = xlsx.sheet_names
        reversed_list = []

        for item in reversed(list_of_sheets):
            reversed_list.append(item)

        list_of_sheets = reversed_list
        list_length = len(list_of_sheets)

    except UnicodeDecodeError:
        message = "Nevažeće kodiranje. Fajl mora biti u UTF-8."
        txt_edit.insert('0.0', message)
        return

    if len(list_of_sheets) == 1:
        convert_data()
    else:
        if listbox_of_frm_list.size() > 0:
            listbox_of_frm_list.delete(0, 'end')

        listbox_of_frm_list.insert('end', *list_of_sheets)
        if len(list_of_sheets) > 10:
            listbox_of_frm_list.configure(height=10)
        else:
            listbox_of_frm_list.configure(height=len(list_of_sheets))
        frm_list_sheets.place(anchor="center", relx=.5, rely=.5)


def bad_file(add_text=None):
    msg = "Netačna struktura datoteke"
    if add_text:
        msg += f"\nNedostajuća kolona {add_text}"
    txt_edit.insert('0.0', msg)


def convert_data(sheet=0):
    global df

    frm_list_sheets.place_forget()
    df = pd.read_excel(xlsx, sheet, header=None)

    # start_of_df = (0, 0)

    for i in range(df.shape[0]):
        for j in range(df.shape[1]):
            if df.iloc[i, j] == 'UKUPNO':
                end_column_of_df = (i, j)
                break
        if 'end_column_of_df' in locals():
            break
    if 'end_column_of_df' not in locals():
        return bad_file(add_text='UKUPNO')

    start_of_df = (end_column_of_df[0], 0)
    # print(f'start_of_df = {start_of_df}')

    for line in reversed(range(df.shape[0])):
        if df.iloc[line, end_column_of_df[1]] > 0:
            end_of_df = (line, end_column_of_df[1])
            break

    if 'end_of_df' not in locals():
        return bad_file()

    df = df.iloc[start_of_df[0]:end_of_df[0], start_of_df[1]:end_of_df[1] + 1]
    df = df.reset_index(drop=True)

    correct_columns = ['IME', 'PLATA', 'RACUN', 'BONUS', 'UMANJENJE', 'UKUPNO']

    for i in reversed(range(df.shape[1])):
        if df.iloc[0, i] not in correct_columns:
            # print(df.iloc[0, i])
            df = df.drop(df.columns[[i]], axis=1)

    df = df.iloc[start_of_df[0]:end_of_df[0], start_of_df[1]:end_of_df[1] + 1]
    df = df.reset_index(drop=True)

    for el in correct_columns:
        found = False
        for j in range(df.shape[1]):
            if el == df.iloc[0, j]:
                found = True
                break
        if not found:
            return bad_file(add_text=el)

    df = df.iloc[start_of_df[0] + 1:end_of_df[0], start_of_df[1]:end_of_df[1] + 1]
    df = df.reset_index(drop=True)

    """
    for i in range(df.shape[0]):
        val = df.iloc[i, 0]

        if type(val) is not str:
            continue

        if val.__contains__('SIFRA'):
            start_of_df = (0, 1)
            break

    for i in range(df.shape[0]):
        val = df.iloc[i, 1]

        if type(val) is not str:
            continue

        if val.__contains__('SIFRA'):
            start_of_df = (0, 2)
            break

    for i in range(df.shape[0]):
        for j in range(df.shape[1]):
            if df.iloc[i, j] == 'UKUPNO':
                end_column_of_df = (i, j)
                break
        if 'end_column_of_df' in locals():
            break

    if 'end_column_of_df' not in locals():
        # print('1')
        return bad_file()

    for line in reversed(range(df.shape[0])):
        if df.iloc[line, end_column_of_df[1]] > 0:
            end_of_df = (line, end_column_of_df[1])
            break

    if 'end_of_df' not in locals():
        # print('2')
        return bad_file()

    df = df.iloc[start_of_df[0] + 1:end_of_df[0] + 1, start_of_df[1]:end_of_df[1] + 1]
    df = df.reset_index(drop=True)

    # print(df)
    # return
    """


    df = df.round(0)

    # message = df.to_string()
    # txt_edit.insert('0.0', message)
    # return

    # Removing false lines
    lines_to_remove = []

    for line in range(df.shape[0]):
        pass_line = False
        name = df.iloc[line, 0]
        total = df.iloc[line, df.shape[1] - 1]

        for i in range(1, df.shape[1] - 1):
            if pass_line:
                break
            if type(df.iloc[line, i]) not in (int, float):
                lines_to_remove.append(line)
                pass_line = True

        if pass_line:
            continue

        try:
            if (len(str(name)) < 4) or ('UKUPNO' in name.strip().upper()) \
                    or (math.isnan(total)) or (math.isnan(name)) or (name == '') \
                    or (name == ' ') or (total == '') or (total == ' '):
                lines_to_remove.append(line)

        except TypeError:
            pass

    for index in lines_to_remove:
        df = df.drop(index)

    df = df.reset_index(drop=True)
    if df.size < 1:
        # print('3')
        return bad_file()

    df.columns = range(df.columns.size)

    # df = df.drop(labels=1, axis=1)
    df = df.reset_index(drop=True)

    # rename columns
    df.rename(columns={df.columns[0]: 'Ime'}, inplace=True)
    for i in range(1, df.shape[1]):
        df.rename(columns={df.columns[i]: f'field_{i}'}, inplace=True)

    # Changing all NaN to zero
    df = df.fillna(0)

    convert_dict = {}
    for i in range(1, df.shape[1]):
        convert_dict[f'{df.columns[i]}'] = int
    df = df.astype(convert_dict)

    message = df.to_string()
    txt_edit.insert('0.0', message)


def check_for_data():
    if df.empty:
        message = 'Nema podataka. Prvo izaberite fajl!'
        txt_edit.delete('0.0', tk.END)
        txt_edit.insert('0.0', message)
        return False
    return True


def save_file(file_name=''):
    if not check_for_data():
        return

    message = ''

    if not file_name:
        filepath = asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF fajlovi", "*.pdf"), ("Sve fajlovi", "*.*")],
        )
        if not filepath:
            return
    else:
        filepath = file_name

    # Define the Jinja2 template, with CSS
    template_str = '''
        <!DOCTYPE html>
        <html>
            <head>
                <meta http-equiv="Content-type" content="text/html; charset=utf-8" />
                <style>
                    table, tr, td, th, tbody, thead, tfoot {
                        page-break-inside: avoid !important;
                    }
                    thead {
                        display: none
                    }
                    tr {
                        display: inline-block;
                        border: 1px dashed grey;
                        padding: 5px;
                        border-collapse: collapse;
                        width: 10em;
                        word-break: break-word;
                    }
                    td {
                        display: block;
                    }
                    td:first-child {
                        height: 36px;
                        text-align: center;
                        border-bottom: 1px solid gray;
                        overflow: hidden;
                    }
                    td:not(:first-child){
                        text-align: right;
                        padding-right: 10px;
                    }
                    td:last-child {
                        font-weight: bold;
                    }
                </style>
             </head>
            <body>
                <table>
                  <thead>
                    <tr>
                      {% for col in cols %}
                        <th>{{ col }}</th>
                      {% endfor %}
                    </tr>
                  </thead>
                  <tbody>
                    {% for row in rows %}
                      <tr>
                        {% for field in row %}
                          <td>{{ field }}</td>
                        {% endfor %}
                      </tr>
                    {% endfor %}
                  </tbody>
                </table>
            </body>
        </html>
        '''

    template = Template(template_str)
    html_table = template.render(cols=df.columns, rows=df.values.tolist())

    html_path = 'temp.html'

    with open(html_path, 'w', encoding='utf-8') as file:
        file.write(html_table)

    pdf_options = {
        'page-size': 'A4',
        'margin-top': '5mm',
        'margin-right': '5mm',
        'margin-bottom': '5mm',
        'margin-left': '5mm',
        'encoding': 'UTF-8',
        'no-outline': None,
        'print-media-type': None,
        'disable-smart-shrinking': None
    }

    try:
        pdfkit.from_file(html_path, filepath, options=pdf_options)
        message = f"Gotovo!\nFajl {filepath} stvorio."
    except Exception as e:
        message = f"Greška u procesu.\n{e}"
    finally:
        txt_edit.delete('0.0', tk.END)
        txt_edit.insert('0.0', message)
        if os.path.exists(html_path):
            os.remove(html_path)


def print_file():
    if not check_for_data():
        return

    filepath = 'print.pdf'
    if os.path.exists(filepath):
        os.remove(filepath)

    save_file(filepath)
    message = "Štampanje..."
    txt_edit.delete('0.0', tk.END)
    txt_edit.insert('0.0', message)
    try:
        win32api.ShellExecute(0, "print", filepath, None, ".", 0)
    except (Exception,):
        win32api.ShellExecute(0, "", filepath, None, ".", 0)


def quit_app():
    window.destroy()


if __name__ == "__main__":
    global list_length
    window = tk.Tk()
    ctk.set_default_color_theme("green")
    window.title("Radnici Obračun")

    window.rowconfigure(0, minsize=600, weight=1)
    window.columnconfigure(1, minsize=1000, weight=1)

    txt_edit = tk.Text(window)
    frm_buttons = ctk.CTkFrame(window)

    frm_list_sheets = ctk.CTkFrame(window)

    btn_open = ctk.CTkButton(frm_buttons, text="Otvori Excel", command=open_file)
    btn_save = ctk.CTkButton(frm_buttons, text="Sačuvaj u PDF", command=save_file)
    btn_print = ctk.CTkButton(frm_buttons, text="Štampati", command=print_file)
    btn_quit = ctk.CTkButton(frm_buttons, text="Izađite", command=quit_app, fg_color='red', hover_color='darkred')

    label_of_frm_list = ctk.CTkLabel(frm_list_sheets, text="Izaberite Ekcel list")
    listbox_of_frm_list = tk.Listbox(frm_list_sheets,
                                     height=6,
                                     width=20,
                                     bg="white",
                                     activestyle='dotbox',
                                     font="Helvetica"
                                     )

    vsb = ctk.CTkScrollbar(frm_list_sheets, command=listbox_of_frm_list.yview)
    vsb.grid(row=1, column=1, sticky='ns')
    listbox_of_frm_list.configure(yscrollcommand=vsb.set)

    btn_select_list = ctk.CTkButton(frm_list_sheets, text="Izaberite list",
                                    command=lambda: convert_data(list_length - 1 - int(listbox_of_frm_list.curselection()[0])))

    frm_buttons.grid(row=0, column=0, sticky="ns")
    txt_edit.grid(row=0, column=1, sticky="nsew")

    btn_open.grid(row=0, column=0, sticky="ew", padx=5, pady=5)
    btn_save.grid(row=1, column=0, sticky="ew", padx=5, pady=5)
    btn_print.grid(row=2, column=0, sticky="ew", padx=5, pady=5)
    btn_quit.grid(row=3, column=0, sticky="ew", padx=5, pady=5)

    label_of_frm_list.grid(row=0, column=0, columnspan=2, sticky="ew", padx=5, pady=5)
    listbox_of_frm_list.grid(row=1, column=0, sticky="ew", padx=5, pady=5)
    btn_select_list.grid(row=2, column=0, columnspan=2, sticky="ew", padx=5, pady=5)

    window.mainloop()
