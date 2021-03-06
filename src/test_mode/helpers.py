import os
import re
import traceback
import olefile
from pdf_metadata import BinaryPdfForensics, binary_string
from config import SLASH, INVALID_CHARS
from patoolib import extract_archive
from shutil import rmtree, copyfile
from time import sleep
import xml.etree.ElementTree as ET
from PyPDF2 import PdfFileReader


def docs_present(bid):
    """
    Counts number of documents in bid.
    :param bid: bid object
    :return: int, list
    """
    docs_num = 0
    docs = []
    if 'documents' in bid.keys():
        docs += bid['documents']
        if 'financialDocuments' in bid.keys():
            docs += bid['financialDocuments']
        if 'eligibilityDocuments' in bid.keys():
            docs += bid['eligibilityDocuments']
        if 'qualification_documents' in bid.keys():
            docs += bid['qualification_documents']
        docs_num = sum([1 for i in docs if i['title'] != 'sign.p7s'])
    return docs_num, docs


def replace_invalid_chars(string):
    for rep in INVALID_CHARS:
        string = string.replace(rep[0], rep[1])
    return string


def remove_folder(folder_path, rm_main=True):
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
            try:
                os.rmdir(os.path.join(root, name))
            except:
                sleep(1)
                os.rmdir(os.path.join(root, name))
    if rm_main: os.rmdir(folder_path)


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
    date = doc['dateModified']
    if ids.count(id) > 0:
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
        print(e, traceback.print_exc())


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


def format_date(dates):
    for x in dates:
        if x[:2] == 'D:':
            yield f'{x[2:6]}-{x[6:8]}-{x[8:10]} {x[10:12]}:{x[12:14]}:{x[14:16]} {x[16:22]}'
        else:
            yield x


def write_prop_value(objs, xmps, objs_decoded, xmps_decoded, worksheet, row, pdf_info):
    """
    Finds property values either in the object returned by PyPDF2 or using regex on xref
    or xmp strings extracted by pdf-metadata.
    :param objs: xref object string
    :param xmps: xmp string
    :param objs_decoded: xref decoded
    :param xmps_decoded: xmp decoded
    :param worksheet: worksheet object of the xls file
    :param row: row to write to
    :param pdf_info: pdf object with xref metadata created by PyPDF2
    :return:
    """
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
            with_styles = []
            if index == 2:
                regexmp_st = fr'(?<=title>)[^\/]+(?=(?:<\/[^<>]*>)*<\/[dcrf:]*title)'
                with_styles = re.findall(regexmp_st, xmps_decoded)
                for i in range(len(with_styles)):
                    with_styles[i] = with_styles[i][with_styles[i].rindex('>') + 1:]
                # print('title', with_styles)
            regexref = fr'(?<={prop_xref})\s*\([^()]+(?=\))'
            regexmp = fr'(?<={prop_xmp}>)[^</]+(?=</)'
            to_write = set(re.findall(regexref, f'{objs} {objs_decoded}')).union(set(re.findall(regexmp,
                                                                            f'{xmps_decoded} {xmps}')))
            # print('from xmp', prop_xmp,  xmp_to_write)
            to_write = [i.strip().strip("(") for i in to_write] + with_styles
        if not to_write:
            to_write = ['']
        if index not in [5, 6]:
            worksheet.write(row, index, ''.join(to_write))
        else:
            worksheet.write(row, index, ''.join(list(format_date(to_write))))


def pdf_metadata(filepath, filename, worksheet, row):
    """
    Extracts pdf metadata using PyPDF2.
    :param filepath: path to file
    :param filename: name of file
    :param worksheet: worksheet object of created xlsx file
    :param row: row number in worksheet for current file
    :return:
    """
    try:
        # print(filename)
        pdf = BinaryPdfForensics(filepath+filename, '')
        version = pdf.pdf_magic()[1]
        worksheet.write(row, 10, version)
        objs_dict = pdf.get_info_obj()[1].values()
        xmps_dict = pdf.get_xmp_obj()[1].values()
        # print(objs_dict, xmps_dict)
        try:
            objs_decoded = ','.join(set([i.decode().replace(r'\000', '') for i in objs_dict]))
        except UnicodeDecodeError:
            objs_decoded = ''
        xmps = ','.join(set([binary_string(i) for i in xmps_dict]))
        xmps_st = ','.join(set([i.decode(errors='ignore') for i in xmps_dict]))
        objs = ','.join(set([binary_string(i) for i in objs_dict]))
        # print(xmps)
        # print(objs)
        # print(objs_decoded)
        with open(filepath + filename, "rb") as f:
            pdf_toread = PdfFileReader(f, strict=False)
            pdf_info = pdf_toread.getDocumentInfo()
            write_prop_value(objs, xmps, objs_decoded, xmps_st, worksheet, row, pdf_info)
    except Exception as e:
        print(e, traceback.print_exc())


functions = {
        '.pdf': pdf_metadata,
        '.doc': doc_metadata,
    }


def run_func(file, *arg):
    functions[file](*arg)


def extract_any_archive(fpath, fname, archive=False, appended=False):
    ext = str(os.path.splitext(fname)[1]).lower()
    if not os.path.exists(f'{fpath}tmp'):
        os.mkdir(f'{fpath}tmp')
    if ext == '':
        os.rename(f'{fpath}{fname}', f'{fpath}{fname}.zip')
        fname += '.zip'
        appended = True
    if ext[:4] not in ['.pdf', '.doc', '.jpg', '.png', '.jpe', '.txt', '.htm', '.csv', '.xls', '.ppt', '.p7s', '.tiff']:
        try:
            extract_archive(f'{fpath}{fname}', outdir=f'{fpath}tmp', verbosity=-1)
            archive = True
            for root, dirs, files in os.walk(f'{fpath}tmp', topdown=False):
                for name in files:
                    extract_any_archive(f'{root}{SLASH}', name)
        except Exception as e:
            print(e)
            # print(fname, e)
    if appended:
        os.rename(f'{fpath}{fname}', f'{fpath}{fname[:-4]}')
    return archive


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
    archive = extract_any_archive(fpath, fname)
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
