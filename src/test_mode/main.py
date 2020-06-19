from helpers import *
import traceback
from tempfile import TemporaryDirectory
from config import *
from threading import Thread
from time import sleep
import tkinter as tk
import tkinter.ttk as ttk
from tkinter.font import Font
from tkinter import messagebox, W, E, filedialog, HORIZONTAL

from string import ascii_letters, digits
from random import choice
from xlsxwriter import Workbook
from requests import get as requests_get


def docs_present(i, bids):
    docs_num = 0
    if 'documents' in bids[i].keys():
        docs = bids[i]['documents']
        if 'financialDocuments' in bids[i].keys():
            docs += bids[i]['financialDocuments']
        if 'eligibilityDocuments' in bids[i].keys():
            docs += bids[i]['eligibilityDocuments']
        if 'qualification_documents' in bids[i].keys():
            docs += bids[i]['qualification_documents']
        docs_num = sum([1 for i in docs if i['title'] != 'sign.p7s'])
    return docs_num


def run_checks(response):
    """
    Function checks whether the tender has documents to download.
    Also it saves the number of files that will be downloaded.
    :return: boolean
    """
    global docs_number
    docs_number = 0
    if 'bids' in response['data'].keys():
        bids = response['data']['bids']
        no_docs = True
        for i in range(len(bids)):
            docs_num = docs_present(i, bids)
            if docs_num > 0:
                no_docs = False
                docs_number += docs_num
        if no_docs:
            messagebox.showinfo("Відсутні файли", "У тендері відсутні файли пропозиції.")
            remove_folder(folder)
            return False
    else:
        messagebox.showinfo("Відсутні пропозиції", "У тендері відсутні пропозиції.")
        remove_folder(folder)
        return False
    return True


def run_progress_bar(response):
    """
    Runs the interface for progress bar, updating number of files.
    :return:
    """
    global docs_done
    docs_done = 0
    th = Thread(target=lambda: download_files(response))
    th.start()
    while docs_done < docs_number - 1:
        progress['value'] = (docs_done / docs_number) * 100
        docs_label['text'] = 'Завантажено: {} з {}'.format(docs_done, docs_number)
        window.update()
        sleep(1)
    while docs_done < docs_number:
        docs_label['text'] = 'Завершення роботи ...'
        window.update()
        sleep(1)
    th.join()
    window.destroy()


def bid_files(bids, lots, tmpdirname, folder, save_m, i, index, nonempty_lots, lots_list):
    global docs_done
    if bids[i]['status'] not in ['invalid', 'deleted']:
        lot_paths = []
        if save_m:
            index_csv = ''.join(choice(ascii_letters + digits) for _ in range(10))
            csvname = f'Значення атрибутів файлів пропозиції {index_csv}.xlsx'
            workbook = Workbook(f'{tmpdirname}{SLASH}{csvname}')
            worksheet = workbook.add_worksheet()
            worksheet.write_row(0, 0, METADATA)
        if lots:
            bid_lot_ids = bids[i]['lotValues']
            for lot in bid_lot_ids:
                lot_id = lot['relatedLot']
                lot = f'{[i[1] for i in lots_list if i[0] == lot_id][0]} {lot_id}'
                lot = replace_invalid_chars(lot) + SLASH
                nonempty_lots.append(f'{folder}{SLASH}{lot}')
                lot_paths.append(lot)
        else:
            lot_paths.append("")

        # participant = f'{bids[i]["tenderers"][0]["name"]} {bids[i]["tenderers"][0]["identifier"]["id"]}'
        participant = bids[i]["tenderers"][0]["identifier"]["id"]
        if 'documents' in bids[i].keys():
            docs = bids[i]['documents']
        else:
            docs = []
        filenames, dates, ids = [], [], []
        row = 0
        old_files = False
        for doc in docs:
            if STOP_EXECUTION:
                for root, dirs, files in os.walk(folder, topdown=False):
                    for name in files:
                        os.remove(os.path.join(root, name))
                    for name in dirs:
                        os.rmdir(os.path.join(root, name))
                os.rmdir(folder)
                return False
            filename = doc['title']
            if filename != "sign.p7s":
                docs_done += 1
                old_files, filename = append_file(doc, filenames, filename, old_files, dates, ids)
                # print(old_files)
                r = requests_get(doc['url'], allow_redirects=True)
                with open(f'{tmpdirname}{SLASH}{filename[0]}{SLASH}{filename[1]}', 'wb') as file_:
                    file_.write(r.content)
                if save_m:
                    row = write_metadata(f'{tmpdirname}{SLASH}', filename[0], filename[1], worksheet,
                                         row)
                    remove_folder(f'{tmpdirname}{SLASH}{filename[0]}{SLASH}tmp')
        if save_m:
            workbook.close()
        for lot in lot_paths:
            participant_path = f'{folder}{SLASH}{lot}{str(index)} {participant}{SLASH}'
            if not os.path.exists(participant_path):
                os.mkdir(participant_path)  # folder of bidder inside lot
                os.mkdir(f'{participant_path}Актуальні{SLASH}')
                if old_files:
                    os.mkdir(f'{participant_path}Видалені{SLASH}')
            for filename in filenames:
                copyfile(f'{tmpdirname}{SLASH}{filename[0]}{SLASH}{filename[1]}',
                         f'{participant_path}{filename[0]}{SLASH}{filename[1]}')
            if save_m:
                copyfile(f'{tmpdirname}{SLASH}{csvname}', f'{participant_path}{csvname}')


def download_files(response):
    """"
    Downloads the files of the tender using Prozorro API.
    :param response: request response json object that contains all tender information
    :return:
    """
    try:
        lots_list = []
        if 'lots' in response['data'].keys():
            lots = response['data']['lots']
            for lot in lots:
                lot_title = lot['title']
                lot_title = replace_invalid_chars(lot_title)[:100]
                path_of_lot = folder + SLASH + lot_title + " " + lot['id']
                os.mkdir(path_of_lot)  # folder of lot inside tender
                lots_list.append([lot['id'], lot_title, False])
        else:
            lots = False

        if 'bids' in response['data'].keys():
            nonempty_lots = []
            global docs_done
            bids = response['data']['bids']
            bids_with_docs = []
            for i in range(len(bids)):
                docs_n = docs_present(i, bids)
                if docs_n:
                    bids_with_docs.append(i)
            save_m = save_meta.get()
            index = 0
            for i in bids_with_docs:
                with TemporaryDirectory(dir=folder) as tmpdirname:
                    os.mkdir(f'{tmpdirname}{SLASH}Актуальні')
                    os.mkdir(f'{tmpdirname}{SLASH}Видалені')
                    bid_files(bids, lots, tmpdirname, folder, save_m, i, index, nonempty_lots, lots_list)
                    sleep(2)
                    index += 1
            docs_done += 1
            if lots:
                lots_all = [os.path.join(folder, o) + SLASH for o in os.listdir(folder)
                            if os.path.isdir(os.path.join(folder, o))]
                for i in lots_all:
                    if i not in nonempty_lots:
                        os.rmdir(i)

    except Exception as e:
        print(e, traceback.print_exc())
        messagebox.showinfo("Помилка", e)


def rClicker(e):
    try:
        def rClick_Copy(e):
            e.widget.event_generate('<<Copy>>')

        def rClick_Cut(e):
            e.widget.event_generate('<<Cut>>')

        def rClick_Paste(e):
            e.widget.event_generate('<<Paste>>')

        def rClick_SelectAll(e):
            e.widget.event_generate('<<SelectAll>>')

        e.widget.focus()

        nclst=[
               (' Вирізати', lambda e=e: rClick_Cut(e)),
               (' Скопіювати', lambda e=e: rClick_Copy(e)),
               (' Вставити', lambda e=e: rClick_Paste(e)),
               (' Виділити все', lambda e=e:rClick_SelectAll(e)),
               ]

        rmenu = tk.Menu(None, tearoff=0, takefocus=0)

        for (txt, cmd) in nclst:
            rmenu.add_command(label=txt, command=cmd)

        rmenu.tk_popup(e.x_root+40, e.y_root+10, entry="0")

    except tk.TclError:
        # print(' - rClick menu, something wrong')
        pass


def browse_button(entry):
    global folder_path
    foldername = filedialog.askdirectory()
    folder_path.set(foldername)
    entry.delete(0, "end")
    entry.insert(0, foldername)


def ok_button(event=None):
    global folder
    folder = PREFIX
    f_path = folder_path.get().strip()
    if SYSTEM == "Windows":
        f_path = f'{folder}{f_path.replace("/", SLASH)}'
    t_id = tender_id.get().strip()
    if not os.path.exists(f_path):
        try:
            os.mkdir(f_path)
        except:
            messagebox.showinfo("Невірний шлях", "Неіснуючий шлях.")
    if os.path.exists(f_path):
        if not t_id.isalnum() or len(t_id) < 32:
            messagebox.showinfo("Невірний ID", "Тендер із введеним ID не знайдено.")
        else:
            response = requests_get("https://public.api.openprocurement.org/api/2.3/tenders/" + t_id)
            if response.status_code != 200:
                messagebox.showinfo("Невірний ID", "Тендер із введеним ID не знайдено.")
            else:
                response = response.json()
                tenderId = response['data']['tenderID']
                # orderer = response['data']['procuringEntity']['name']
                # folder = f'{f_path}{SLASH}{orderer} {tenderId}'
                folder = f'{f_path}{SLASH}{tenderId}'
                if not os.path.exists(folder):
                    os.mkdir(folder)  # folder of tender
                    if run_checks(response):
                        container.grid_forget()
                        container2.grid(row=1, column=1)
                        run_progress_bar(response)
                else:
                    if run_checks(response):
                        container.grid_forget()
                        container2.grid(row=1, column=1)
                        confirmation_window(response)


def yes_button(top, response):
    top.destroy()
    for root, dirs, files in os.walk(folder, topdown=False):
        for name in files:
            os.remove(os.path.join(root, name))
        for name in dirs:
            os.rmdir(os.path.join(root, name))
    run_progress_bar(response)


def no_button(top, stop_execution):
    global STOP_EXECUTION
    top.destroy()
    if not stop_execution:
        container2.grid_forget()
        container.grid(row=1, column=1)
    else:
        STOP_EXECUTION = True
#    if window != top:
#        window.destroy()


def confirmation_window(response):
    top = tk.Toplevel(window)
    wd = 240  # width for the Tk root
    ht = 140  # height for the Tk root
    x1 = (ws // 2) - (wd // 2)  # x and y coordinates for the Tk root window
    y1 = (hs // 2) - (ht // 2)
    top.geometry(f'{wd}x{ht}+{x1}+{y1}')
    top.resizable(width=wd, height=ht)

    rewrite = tk.Label(master=top, text='Папка даного тендеру існує, \n перезаписати?')
    rewrite.grid(row=0, column=1, columnspan=2, pady=15, padx=15)
    yes = tk.Button(top, text="Так", command=lambda: yes_button(top, response))
    yes.grid(row=1, column=1, pady=10, padx=10, sticky=E)
    no = tk.Button(top, text="Ні", command=lambda: no_button(top, False))
    no.grid(row=1, column=2, pady=10, padx=10, sticky=W)
    top.bind('<Return>', lambda event: yes_button(top, response))
    ent_tender.unbind("<Return>", bind_id)
    for widg, color in zip([top, rewrite, yes, no], [BACKGROUND, BACKGROUND, GREEN, RED]):
        widg['bg'] = color
    top.mainloop()


# def key_(event):
#     print(event.keysym_num)
#     print(event.keysym)
#     print(event.keycode)


def keypress(event):
    ctrl = (event.state & 0x4) != 0
    if event.keycode == 86 and ctrl and event.keysym.lower() != "v":
        event.widget.event_generate('<<Paste>>')
    elif event.keycode == 67 and ctrl and event.keysym.lower() != "c":
        event.widget.event_generate('<<Copy>>')
    elif event.keycode == 88 and ctrl and event.keysym.lower() != "x":
        event.widget.event_generate('<<Cut>>')
    elif event.keycode == 65 and ctrl and event.keysym.lower() != "a":
        event.widget.event_generate('<<SelectAll>>')


if __name__ == '__main__':

    window = tk.Tk()
    # window.iconbitmap('favicon.ico')
    window.title("Вивантаження файлів")

    w = 425  # width for the Tk root
    h = 180  # height for the Tk root
    ws = window.winfo_screenwidth()  # width of the screen
    hs = window.winfo_screenheight()  # height of the screen
    x = (ws // 2) - (w // 2) # x and y coordinates for the Tk root window
    y = (hs // 2) - (h // 2)
    window.geometry(f'{w}x{h}+{x}+{y}')
    window.resizable(width=w, height=h)

    myFont = Font(family="Times New Roman", size=13)
    window.option_add("*Font", myFont)

    folder_path = tk.StringVar()

    try:
        with open('last_use_folder.txt', 'r') as file:
            folder_path.set(file.read())
    except IOError:
        with open('last_use_folder.txt', 'w'):
            pass

    container = tk.Frame(window)
    tender_id = tk.StringVar()
    id_tender = tk.Label(master=container, text='ID закупівлі: ')
    id_tender.grid(row=0, column=1, sticky=W, pady=20, padx=10)
    ent_tender = tk.Entry(container, width=25, textvariable=tender_id)
    # ent_tender.bind('<Any-KeyPress>', key_)
    ent_tender.grid(row=0, column=2, pady=20, padx=0, sticky=W)
    folder_ = tk.Label(master=container, text='Папка: ')
    folder_.grid(row=1, column=1, sticky=W, pady=5, padx=10)
    ent_folder = tk.Entry(container, width=25, textvariable=folder_path)
    ent_folder.grid(row=1, column=2, pady=5, padx=0, sticky=W)
    button = tk.Button(container, text="...", width=5, command=lambda: browse_button(ent_folder))
    button.grid(row=1, column=2, pady=10, padx=240, sticky=W)
    button2 = tk.Button(container, text="Завантажити", command=ok_button)
    button2.grid(row=2, column=1, pady=15, padx=10)
    button3 = tk.Button(container, text="Скасувати", command=lambda: no_button(window, True))
    button3.grid(row=2, column=2, pady=15, padx=10, sticky=W)
    save_meta = tk.IntVar()
    cbutton = tk.Checkbutton(container, text="зберегти метадані", variable=save_meta, onvalue=1, offvalue=0)
    cbutton.grid(row=2, column=2, padx=110, sticky=W)
    cbutton.select()
    if SYSTEM == 'Windows':
        for ent in [ent_folder, ent_tender]:
            ent.bind('<Button-3>', rClicker, add='')
            ent.bind('<Key>', keypress)
    bind_id = ent_tender.bind('<Return>', ok_button)

    for wid in [id_tender, folder_, container, cbutton]:
        wid['bg'] = BACKGROUND
    for wid, colour in zip([ent_tender, ent_folder, button2, button3, button],
                           [LIGHTBLUE, LIGHTBLUE, GREEN, RED, DARKRED]):
        wid['bg'] = colour
    button['fg'] = BACKGROUND

    container2 = tk.Frame(window)

    s = ttk.Style()
    s.theme_use('clam')
    s.configure("bar.Horizontal.TProgressbar", troughcolor=LIGHTBLUE, bordercolor=BACKGROUND, background=DARKRED,
                lightcolor=DARKRED, darkcolor=DARKRED)

    docs_label = tk.Label(master=container2, text='Завантажено: {} з {}'.format(docs_done, docs_number))
    docs_label.grid(row=0, column=1, sticky=W, pady=22, padx=120)
    progress = ttk.Progressbar(container2, style="bar.Horizontal.TProgressbar", orient=HORIZONTAL,
                               length=260, mode='determinate')
    progress.grid(row=1, column=1, pady=15, padx=85)

    button_stop = tk.Button(container2, text='Скасувати', command=lambda: no_button(window, True))
    button_stop.grid(row=2, column=1, pady=15, padx=80)
    container.grid(row=1, column=1)

    for wid, colour in zip([container2, docs_label, button_stop],
                           [BACKGROUND, BACKGROUND, RED]):
        wid['bg'] = colour

    window.mainloop()
    with open('last_use_folder.txt', 'w') as file:
        file.write(folder_path.get())
