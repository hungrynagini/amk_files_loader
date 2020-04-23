import requests
import os
import threading
import shutil
import tkinter as tk
import tkinter.ttk as ttk
from time import sleep
from tkinter import messagebox, W, E, filedialog, HORIZONTAL

INVALID_CHARS = [":", "\\", "|", "/", "?", "*", ">", "<"]
LONG_PATH_PREFIX = "\\\\?\\"
SLASH = "\\"
# SLASH = "/"  # ubuntu
DOCS_DONE = 0


def run_checks():
    """
    Function checks whether the tender has documents to download.
    Also it saves the number of files that will be downloaded.
    :return: boolean
    """
    global DOCS_NUMBER
    if 'bids' in response_gl['data'].keys():
        bids = response_gl['data']['bids']
        no_docs = True
        for i in range(len(bids)):
            if 'documents' in bids[i].keys():
                docs = bids[i]['documents']
                docs_num = sum([1 for i in docs if i['title'] != 'sign.p7s'])
                if docs_num > 0:
                    no_docs = False
                    DOCS_NUMBER += docs_num
                # for j in docs:
                #     site = urllib.request.urlopen(j['url'])
                #     meta = site.info()
                #     print(meta["Content-Length"])
        if no_docs:
            messagebox.showinfo("Відсутні файли", "У тендері відсутні файли пропозиції.")
            shutil.rmtree(folder, ignore_errors=True)
            return False
    else:
        messagebox.showinfo("Відсутні пропозиції", "У тендері відсутні пропозиції.")
        shutil.rmtree(folder, ignore_errors=True)
        return False
    return True


def run_progress_bar():
    """
    Runs the interface for progress bar, updating number of files.
    :return:
    """
    th = threading.Thread(target=download_files)
    th.start()
    while DOCS_DONE < DOCS_NUMBER - 1:
        progress['value'] = (DOCS_DONE / DOCS_NUMBER) * 100
        docs_label['text'] = 'Завантажено: {} з {}'.format(DOCS_DONE, DOCS_NUMBER)
        window.update()
        sleep(1)
    th.join()
    window.destroy()


def download_files():
    """
    Downloads the files of the tender using Prozorro API.
    :return:
    """
    prefix_and_folder = folder
    response = response_gl

    lots_list = []
    if 'lots' in response['data'].keys():
        lots = response['data']['lots']
        for lot in lots:
            lot_title = lot['title']
            for rep in INVALID_CHARS:
                lot_title = lot_title.replace(rep, "_")
            lot_title = lot_title.replace("\"", "'")
            path_of_lot = prefix_and_folder + SLASH + lot_title + " " + lot['id']
            os.mkdir(path_of_lot)  # folder of lot inside tender
            lots_list.append([lot['id'], lot_title, False])
    else:
        lots = False

    if 'bids' in response['data'].keys():
        nonempty_lots = []
        global DOCS_DONE
        bids = response['data']['bids']
        bids_with_docs = []
        for i in range(len(bids)):
            if 'documents' in bids[i].keys():
                docs = bids[i]['documents']
                if False in [i['title'] == 'sign.p7s' for i in docs]:
                    bids_with_docs.append(i)

        for i in bids_with_docs:
            lot_paths = []
            if bids[i]['status'] != 'invalid':

                if lots:
                    bid_lot_ids = bids[i]['lotValues']
                    for lot in bid_lot_ids:
                        lot_id = lot['relatedLot']
                        lot = [i[1] for i in lots_list if i[0] == lot_id][0] + " " + lot_id
                        for rep in INVALID_CHARS:
                            lot = lot.replace(rep, "_")
                        lot = lot.replace("\"", "'") + SLASH
                        nonempty_lots.append(prefix_and_folder + SLASH + lot)
                        lot_paths.append(lot)
                else:
                    lot_paths.append("")

                participant = bids[i]['tenderers'][0]['name'] + " " + bids[i]['tenderers'][0]['identifier']['id']
                for rep in INVALID_CHARS:
                    participant = participant.replace(rep, "_")
                participant = participant.replace("\"", "'")
                if 'documents' in bids[i].keys():
                    docs = bids[i]['documents']
                else:
                    docs = []

                filenames = []
                for doc in docs:
                    DOCS_DONE += 1
                    filename = doc['title']
                    for rep in INVALID_CHARS:
                        filename = filename.replace(rep, "_")
                    filename = filename.replace("\"", "'")
                    if filename != "sign.p7s":
                        filenames.append(filename)
                        if filenames.count(filename) > 1:
                            filename = str(filenames.count(filename) - 1) + " " + filename
                            filenames.append(filename)
                        try:
                            r = requests.get(doc['url'], allow_redirects=True)
                            with open(prefix_and_folder + SLASH + filename, 'wb') as file_:
                                file_.write(r.content)
                            print(filename)
                        except:
                            print('ERROR')
                            messagebox.showinfo("Виникла помилка",
                                                "При завантаженні файлів виникла помилка, вибачте за незручності :\\ .")

                filenames = list(set(filenames))

                for lot in lot_paths:
                    participant_path = prefix_and_folder + SLASH + lot + str(i) + " " + participant + SLASH
                    print(lot + str(i) + " " + participant)
                    if not os.path.exists(participant_path):
                        os.mkdir(participant_path)  # folder of bidder inside lot
                    for filename in filenames:
                        shutil.copyfile(prefix_and_folder + SLASH + filename, participant_path + filename)

                for filename in filenames:
                    if filename != "sign.p7s":
                        os.remove(prefix_and_folder + SLASH + filename)

        if lots:
            lots_all = [os.path.join(prefix_and_folder, o) + SLASH for o in os.listdir(prefix_and_folder)
                        if os.path.isdir(os.path.join(prefix_and_folder, o))]
            # print(lots_all, nonempty_lots)
            for i in lots_all:
                if i not in nonempty_lots:
                    print(i[50:], [i[50:] for i in nonempty_lots])
                    os.rmdir(i)

# except OSError:
#     print(os.error)
#     messagebox.showinfo("Помилка", "Перевірте будь ласка, чи у Вас не відкрита папка тендеру.")


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


def ok_button():
    global response_gl
    global folder
    f_path = folder_path.get()
    f_path = LONG_PATH_PREFIX + f_path.replace("/", SLASH)
    t_id = tender_id.get().strip(" ")
    if not os.path.exists(f_path):
        try:
            os.mkdir(f_path)
        except:
            messagebox.showinfo("Невірний шлях", "Неіснуючий шлях.")
    if os.path.exists(f_path):
        if not t_id.isalnum() or len(t_id) < 32:
            messagebox.showinfo("Невірний ID", "Тендер із введеним ID не знайдено.")
        else:
            response = requests.get("https://public.api.openprocurement.org/api/2.3/tenders/" + t_id)
            if response.status_code != 200:
                messagebox.showinfo("Невірний ID", "Тендер із введеним ID не знайдено.")
            else:
                response_gl = response.json()
                response = response.json()
                tenderId = response['data']['tenderID']
                orderer = response['data']['procuringEntity']['name']
                for rep in INVALID_CHARS:
                    orderer = orderer.replace(rep, '_')
                orderer = orderer.replace("\"", "'")
                folder = f_path + SLASH + orderer + " " + tenderId
                container.grid_forget()
                container2.grid(row=1, column=1)
                if not os.path.exists(folder):
                    os.mkdir(folder)  # folder of tender
                    if run_checks():
                        run_progress_bar()
                else:
                    if run_checks():
                        confirmation_window(folder)


def yes_button(top, folder):
    top.destroy()
    for root, dirs, files in os.walk(folder, topdown=False):
        for name in files:
            os.remove(os.path.join(root, name))

        for name in dirs:
            os.rmdir(os.path.join(root, name))
    run_progress_bar()


def no_button(top):
    top.destroy()


def confirmation_window(folder):
    top = tk.Toplevel(window)
    top.resizable(width=False, height=False)
    rewrite = tk.Label(master=top, text='Папка даного тендеру інсує, перезаписати?')
    rewrite.grid(row=0, column=1, columnspan=2, pady=10, padx=10)
    yes = tk.Button(top, text="Так", command=lambda: yes_button(top, folder))
    yes.grid(row=1, column=1, pady=10, padx=10, sticky=E)
    no = tk.Button(top, text="Ні", command=lambda: no_button(top))
    no.grid(row=1, column=2, pady=10, padx=10, sticky=W)
    top.mainloop()


def key_(event):
    print(event.keysym_num)
    print(event.keysym)
    print(event.keycode)


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
    window.iconbitmap('favicon.ico')
    window.title("Вивантаження файлів")
    window.resizable(width=False, height=False)

    folder_path = tk.StringVar()
    folder = ""
    response_gl = ""
    DOCS_NUMBER = 0

    try:
        with open('last_use_folder.txt', 'r') as file:
            folder_path.set(file.read())
    except IOError:
        with open('last_use_folder.txt', 'w'):
            pass

    container = tk.Frame(window, width=300, height=300)
    tender_id = tk.StringVar()
    id_tender = tk.Label(master=container, text='ID закупівлі: ')
    id_tender.grid(row=0, column=1, sticky=W, pady=10, padx=10)
    ent_tender = tk.Entry(container, width=30, textvariable=tender_id)
    # ent_tender.bind('<Any-KeyPress>', key_)
    ent_tender.grid(row=0, column=2, pady=10, padx=10)
    folder_ = tk.Label(master=container, text='Папка: ')
    folder_.grid(row=1, column=1, sticky=W, pady=10, padx=10)
    ent_folder = tk.Entry(container, width=30, textvariable=folder_path)
    ent_folder.grid(row=1, column=2, pady=10, padx=10)
    button = tk.Button(container, text="...", width=5, command=lambda: browse_button(ent_folder))
    button.grid(row=1, column=3, pady=10, padx=10)
    button2 = tk.Button(container, text="Завантажити", command=ok_button)
    button2.grid(row=2, column=1, pady=10, padx=10)
    button3 = tk.Button(container, text="Скасувати", command=lambda: no_button(window))
    button3.grid(row=2, column=2, pady=10, padx=10, sticky=W)
    for ent in [ent_folder, ent_tender]:
        ent.bind('<Button-3>', rClicker, add='')
        ent.bind('<Key>', keypress)

    container2 = tk.Frame(window)
    docs_label = tk.Label(master=container2, text='Завантажено: {} з {}'.format(DOCS_DONE, DOCS_NUMBER))
    docs_label.grid(row=0, column=1, sticky=W, pady=10, padx=10)
    progress = ttk.Progressbar(container2, orient=HORIZONTAL,
                               length=150, mode='determinate')
    progress.grid(row=1, column=1, pady=20, padx=40)

    button_stop = tk.Button(container2, text='Скасувати', command=lambda: no_button(window))
    button_stop.grid(row=2, column=1, pady=10, padx=10)
    container.grid(row=1, column=1)

    window.mainloop()
    with open('last_use_folder.txt', 'w') as file:
        file.write(folder_path.get())
