import os
from config import SLASH
# from unrar import rarfile
from patoolib import extract_archive
from shutil import rmtree, copyfile
from time import sleep
import xml.etree.ElementTree as ET
from PyPDF2 import PdfFileReader


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
    extract_archive(f'{filepath}{filename}.zip', outdir=f'.{SLASH}.doc_unzipped', program='py_zipfile', verbosity=-1)
    for file, propers, indexes in zip(['core.xml', 'app.xml'], [['title', 'creator', 'subject',
             'created', 'modified', 'lastModifiedBy'], ['Manager', 'Company']], [[2, 3, 4, 5, 6, 7], [12, 13]]):
        tree = ET.parse(f".{SLASH}.doc_unzipped{SLASH}docProps{SLASH}{file}")
        root = tree.getroot()
        for child in root:
            for prop, index in zip(propers, indexes):
                if prop in child.tag:
                    worksheet.write(row, index, child.text)
    remove_folder(f'.{SLASH}.doc_unzipped')
    os.rename(f'{filepath}{filename}.zip', f'{filepath}{filename}')


def pdf_metadata(filepath, filename, worksheet, row):
    """
    Extracts pdf metadata using PyPDF2.
    :param filepath: path to file
    :param filename: name of file
    :param worksheet: worksheet object of created xlsx file
    :param row: row number in worksheet for current file
    :return:
    """

    def format_date(x):
        if x:
            return f'{x[2:6]}-{x[6:8]}-{x[8:10]} {x[10:12]}:{x[12:14]}:{x[14:16]} {x[16:22]}'
        else:
            return ''

    try:
        with open(filepath + filename, "rb") as f:
            pdf_toread = PdfFileReader(f, strict=False)
            pdf_info = pdf_toread.getDocumentInfo()
            for prop, index in zip(['/Title', '/Author', '/Subject', '/CreationDate', '/ModDate',
               '/Producer', "/Creator", '/Version', '/Keywords'], [2, 3, 4, 5, 6, 8, 9, 10, 11]):
                try:
                    if index not in [5, 6]:
                        worksheet.write(row, index, pdf_info[prop])
                    else:
                        worksheet.write(row, index, format_date(pdf_info[prop]))
                except:
                    continue
        with open(filepath + filename, "rb") as f:
            read_file = f.read(10)
            magic_val = read_file[0:4].decode()
            pdf_version = read_file[1:8].decode()
            if magic_val == '%PDF':
                worksheet.write(row, 10, pdf_version)
    except:
        print("pdf failed")
    # try:
    #     pdf_info = pdf_toread.getXmpMetadata()
    #     worksheet.write(row, 11, pdf_info.pdf_keywords)
    # except:
    #     return


functions = {
        '.pdf': pdf_metadata,
        '.doc': doc_metadata,
    }


def run_func(file, *arg):
    functions[file](*arg)


def write_metadata(fpath, fname, worksheet, row):
    """
    Writes metadata of DOC and PDF files to xlsx file of participant
    :param fpath: path to file
    :param fname: name of file
    :param worksheet: worksheet object of created xlsx file
    :param row: row number in worksheet for current file
    :return: row number for next file
    """
    if not os.path.exists(f'.{SLASH}.tmp'):
        os.mkdir(f'.{SLASH}.tmp')
    archive = True
    ext = str(os.path.splitext(fname)[1]).lower()
    if ext == '':
        os.rename(f'{fpath}{fname}', f'{fpath}{fname}.zip')
        extract_archive(f'{fpath}{fname}.zip', outdir=f'.{SLASH}.tmp', verbosity=-1)
        os.rename(f'{fpath}{fname}.zip', f'{fpath}{fname}')
    # elif ext == '.zip':
    #     extract_archive(f'{fpath}{fname}', outdir=f'.{SLASH}.tmp', program='py_zipfile', verbosity=-1)
    # elif ext == '.rar':
        # rarfile.RarFile(f'{fpath}{fname}').extractall(f'.{SLASH}.tmp')
    elif ext[:4] not in ['.pdf', '.doc', '.jpg', '.png', '.jpe', '.txt', '.htm', '.csv', '.xls', '.ppt']:
        try:
            extract_archive(f'{fpath}{fname}', outdir=f'.{SLASH}.tmp', verbosity=-1)
        except:
            archive = False
    else:
        archive = False
    filenames, paths = [], []
    if not archive:
        filenames = [fname]
        paths = [fpath]
    for root, dirs, files in os.walk(f'.{SLASH}.tmp', topdown=False):
        for name in files:
            filenames.append(name)
            paths.append(f"{root}{SLASH}")
    for filename, filepath in zip(filenames, paths):
        file_extension = str(os.path.splitext(filename)[1]).lower()
        if file_extension in ['.pdf', '.doc', '.docx']:
            row += 1
            worksheet.write_row(row, 0, [filename, file_extension])
            worksheet.write(row, 14, fname * archive)
            run_func(file_extension[:4], filepath, filename, worksheet, row)
    return row
