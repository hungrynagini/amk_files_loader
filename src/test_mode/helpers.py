import os
import re
import olefile
from pdf_metadata import BinaryPdfForensics, binary_string
from config import SLASH, INVALID_CHARS
from patoolib import extract_archive
from shutil import rmtree, copyfile
from time import sleep
import xml.etree.ElementTree as ET
from PyPDF2 import PdfFileReader


def replace_invalid_chars(string):
    for rep in INVALID_CHARS:
        string = string.replace(rep[0], rep[1])
    return string


def remove_folder(folder_path):
    """
    Removes all contents of folder
    :param folder_path: path to folder to remove
    :return:
    """
    for root, dirs, files in os.walk(folder_path, topdown=False):
        for name in files:
            try:
                os.remove(os.path.join(root, name))
            except:
                sleep(1)
                os.remove(os.path.join(root, name))
        for name in dirs:
            os.rmdir(os.path.join(root, name))
    os.rmdir(folder_path)


def append_file(doc, filenames, filename, old_files, dates, ids):
    """
    Add file to list, with its status (Актуальні/Видалені) and a
    new name in case of duplicates.
    :param doc: file object
    :param filenames: list of filenames and statuses
    :param filename: string, name of file
    :param old_files: bool, tells whether deleted files exist
    :param dates: list of dates of publishing
    :param ids: list of file ids
    :return: [boolean, list]
    """
    filename = replace_invalid_chars(filename)
    filename = ['Актуальні', filename]
    id = ['Актуальні', doc['id']]
    date = doc['datePublished']
    if ids.count(id):
        old_files = True
        ind_of_same = ids.index(id)
        if dates[ind_of_same] < date:
            filenames[ind_of_same][0] = 'Видалені'
            ids[ind_of_same][0] = 'Видалені'
        else:
            filename[0] = 'Видалені'
            id[0] = 'Видалені'
    if filenames.count(filename):
        filename[1] = str(filenames.count(filename) - 1) + " " + filename[1]
    filenames.append(filename)
    dates.append(date)
    ids.append(id)
    return old_files, filename


def old_doc_metadata(filename, filepath, row, worksheet):
    """
    Extracts metadata from doc files with Structured Storage or
    Microsoft Compound Document File Format.
    :param filepath: path to file
    :param filename: name of file
    :param worksheet: worksheet object of created xlsx file
    :param row: row number in worksheet for current file
    :return:
    """
    if filename.endswith('docx'): return ''
    try:
        with olefile.OleFileIO(f'{filepath}{filename}') as ole:
            meta = ole.get_metadata()
            props = [meta.title, meta.author, meta.subject, meta.create_time,
                     meta.last_saved_time, meta.last_saved_by, meta.manager, meta.company]
            for prop, index in zip(props, [2, 3, 4, 5, 6, 7, 12, 13]):
                try:
                    worksheet.write(row, index, prop.decode('unicode_escape'))
                except:
                    worksheet.write(row, index, str(prop))
    except Exception as e:
        print(e)


def doc_metadata(filepath, filename, worksheet, row):
    """
    Extracts metadata from doc files using core.xml and app.xml files of doc file
    (which appear in folder {SLASH}docProps after unzipping the file).
    :param filepath: path to file
    :param filename: name of file
    :param worksheet: worksheet object of created xlsx file
    :param row: row number in worksheet for current file
    :return:
    """
    os.rename(f'{filepath}{filename}', f'{filepath}{filename}.zip')
    try:
        extract_archive(f'{filepath}{filename}.zip', outdir=f'.{SLASH}.doc_unzipped', verbosity=-1)
        os.rename(f'{filepath}{filename}.zip', f'{filepath}{filename}')
        for file, propers, indexes in zip(['core.xml', 'app.xml'], [['title', 'creator', 'subject',
                 'created', 'modified', 'lastModifiedBy'], ['Manager', 'Company']], [[2, 3, 4, 5, 6, 7], [12, 13]]):
            try:
                tree = ET.parse(f".{SLASH}.doc_unzipped{SLASH}docProps{SLASH}{file}")
                root = tree.getroot()
                for child in root:
                    for prop, index in zip(propers, indexes):
                        if prop in child.tag:
                            worksheet.write(row, index, child.text)
            except Exception as e:
                print(e)
                old_doc_metadata(filename, filepath, row, worksheet)
        remove_folder(f'.{SLASH}.doc_unzipped')
    except Exception as e:
        print(e)
        os.rename(f'{filepath}{filename}.zip', f'{filepath}{filename}')
        old_doc_metadata(filename, filepath, row, worksheet)


def pdf_metadata(filepath, filename, worksheet, row):
    """
    Extracts pdf metadata using PyPDF2.
    :param filepath: path to file
    :param filename: name of file
    :param worksheet: worksheet object of created xlsx file
    :param row: row number in worksheet for current file
    :return:
    """

    def format_date(dates):
        for x in dates:
            if x[:2] == 'D:':
                yield f'{x[2:6]}-{x[6:8]}-{x[8:10]} {x[10:12]}:{x[12:14]}:{x[14:16]} {x[16:22]}'
            else:
                yield x

    try:
        pdf = BinaryPdfForensics(filepath+filename, '')
        version = pdf.pdf_magic()[1]
        worksheet.write(row, 10, version)
        objs_dict = pdf.get_info_obj()[1].values()
        xmps_dict = pdf.get_xmp_obj()[1].values()
        try:
            objs_decoded = ','.join(set([i.decode().replace(r'\000', '') for i in objs_dict]))
        except UnicodeDecodeError:
            objs_decoded = ''
        xmps = ','.join(set([binary_string(i) for i in xmps_dict]))
        objs = ','.join(set([binary_string(i) for i in objs_dict]))
        with open(filepath + filename, "rb") as f:
            pdf_toread = PdfFileReader(f, strict=False)
            pdf_info = pdf_toread.getDocumentInfo()
            for prop_xref, prop_xmp, index in zip(['/Title', '/Author', '/Subject', '/CreationDate', '/ModDate',
               '/Producer', "/Creator", '/Keywords'], ['Title', 'Author', 'Subject', 'CreateDate', 'ModifyDate',
               'Producer', 'CreatorTool', 'Keywords'], [2, 3, 4, 5, 6, 8, 9, 11]):
                try:
                    try:
                        to_write = [pdf_info[prop_xref].decode('unicode_escape')]
                    except:
                        to_write = [pdf_info[prop_xref]]
                except Exception as e:
                    to_write = []
                if not to_write:
                    regexref = fr'(?<={prop_xref})\s*\([^()]+(?=\))'
                    regexmp = fr'(?<={prop_xmp}>)[^</]+(?=</)'
                    to_write = set(re.findall(regexref, f'{objs} {objs_decoded}')).union(
                        set(re.findall(regexmp, xmps)))
                    to_write = [i.strip().strip("(") for i in to_write]
                if not to_write: to_write = ['']
                if index not in [5, 6]:
                    worksheet.write(row, index, ''.join(to_write))
                else:
                    worksheet.write(row, index, ''.join(list(format_date(to_write))))
    except Exception as e:
        print(e)


functions = {
        '.pdf': pdf_metadata,
        '.doc': doc_metadata,
    }


def run_func(file, *arg):
    functions[file](*arg)


def write_metadata(fpath, fstatus, fname, worksheet, row):
    """
    Writes metadata of DOC and PDF files to xlsx file of participant
    :param fpath: path to file
    :param fname: name of file
    :param worksheet: worksheet object of created xlsx file
    :param row: row number in worksheet for current file
    :return: row number for next file
    """
    fpath = f'{fpath}{fstatus}{SLASH}'
    if not os.path.exists(f'{fpath}tmp'):
        os.mkdir(f'{fpath}tmp')
    archive = True
    ext = str(os.path.splitext(fname)[1]).lower()
    if ext == '':
        os.rename(f'{fpath}{fname}', f'{fpath}{fname}.zip')
        try:
            extract_archive(f'{fpath}{fname}.zip', outdir=f'{fpath}tmp', verbosity=-1)
        except Exception as e:
            # print("1")
            # print(fname, e)
            archive = False
        os.rename(f'{fpath}{fname}.zip', f'{fpath}{fname}')
    # elif ext == '.zip':
    #     extract_archive(f'{fpath}{fname}', outdir=f'.{SLASH}.tmp', program='py_zipfile', verbosity=-1)
    # elif ext == '.rar':
        # rarfile.RarFile(f'{fpath}{fname}').extractall(f'.{SLASH}.tmp')
    elif ext[:4] not in ['.pdf', '.doc', '.jpg', '.png', '.jpe', '.txt', '.htm', '.csv', '.xls', '.ppt']:
        try:
            extract_archive(f'{fpath}{fname}', outdir=f'{fpath}tmp', verbosity=-1)
        except Exception as e:
            # print("2")
            # print(fname, e)
            archive = False
    else:
        archive = False
    filenames, paths = [], []
    if not archive:
        filenames = [fname]
        paths = [fpath]
    for root, dirs, files in os.walk(f'{fpath}tmp', topdown=False):
        for name in files:
            filenames.append(name)
            paths.append(f"{root}{SLASH}")
    for filename, filepath in zip(filenames, paths):
        file_extension = str(os.path.splitext(filename)[1]).lower()
        if file_extension in ['.pdf', '.doc', '.docx']:
            row += 1
            worksheet.write_row(row, 0, [filename, file_extension])
            worksheet.write(row, 14, fname * archive)
            worksheet.write(row, 15, fstatus)
            run_func(file_extension[:4], filepath, filename, worksheet, row)
    return row
