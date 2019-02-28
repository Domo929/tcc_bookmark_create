import argparse
import tkinter as tk
import re
import os
from tkinter import filedialog, Listbox
from PyPDF3 import PdfFileReader, PdfFileWriter

initial_flag = '_INITIAL '

str_dict = {
    'A': '1',
    'B': '2',
    'C': '3',
    'D': '4',
    'E': '5',
    'F': '6',
    'G': '7',
    'H': '8',
    'I': '9',
    'J': '10',
    'K': '11',
    'L': '12',
    'M': '13',
    'N': '14',
    'O': '15',
    'P': '16',
    'Q': '17',
    'R': '18',
    'S': '19',
    'T': '20',
    'U': '21',
    'V': '22',
    'W': '23',
    'X': '24',
    'Y': '25',
    'Z': '26',
    '0': '0',
    '1': '1',
    '2': '2',
    '3': '3',
    '4': '4',
    '5': '5',
    '6': '6',
    '7': '7',
    '8': '8',
    '9': '9'
}


def argument_handler():
    parser = argparse.ArgumentParser(description='Creates a combined, book-marked report')
    parser.add_argument('-d', '--directory',
                        action='store',
                        help='Optional flag to provide the directory from the commandline',
                        type=str)
    return vars(parser.parse_args())


def choose_directory(opts):
    directory_path = ''
    if opts['directory']:
        directory_path = opts['directory']
    else:
        root = tk.Tk()
        root.withdraw()

        directory_path = filedialog.askdirectory(title='Please choose the directory with all the reports in it')

    return directory_path


def file_selection(directory_path):
    files = [f for f in os.listdir(directory_path)
             if os.path.isfile(os.path.join(directory_path, f)) and initial_flag not in f]

    root = tk.Tk()
    scrollbar = tk.Scrollbar(root, orient='vertical')

    listbox = Listbox(root, width=50, height=50, yscrollcommand=scrollbar.set, selectmode=tk.SINGLE)
    scrollbar.config(command=listbox.yview)
    scrollbar.pack(side='right', fill='y')
    listbox.pack(side='left', fill='both', expand=True)
    tk.Button(root, text='Delete', command=lambda lb=listbox: lb.delete(tk.ANCHOR)).pack(side='bottom', fill='both')
    tk.Button(root, text="Quit", command=root.destroy).pack(side='bottom', fill='both')
    listbox.pack()
    for f in files:
        listbox.insert(tk.END, f)

    root.mainloop()

    return listbox.get(tk.ANCHOR)


def directory_to_files_to_dict(directory_path):
    num_match = r"([\d]+.[\w\d]*)"
    temp = {}
    for f in os.listdir(directory_path):
        if os.path.isfile(os.path.join(directory_path, f)) and initial_flag not in f:
            pdf_num = re.match(num_match, f).group(0)
            temp[string_converter(pdf_num)] = f

    return_dict = {}
    for key in sorted(list(temp)):
        return_dict[key] = temp[key]
    return return_dict


def string_converter(word):
    list_of_char = []
    for char in word:
        char.upper()
        if char == ".":
            list_of_char.append(".")
        else:
            list_of_char.append(str_dict[char])
    num = float("".join(list_of_char))
    return num


def open_all_pdfs(directory, file_dict):
    bookmark_grep = r"([a-zA-Z\d.]*[ -]*[a-zA-Z& ]+[\d]*[.]?[\d]*)"
    list_of_pdf_obj = {}
    for num, filename in file_dict.items():
        full_path = os.path.join(directory, filename)
        bookmark_name = re.search(bookmark_grep, filename).group(1)
        if bookmark_name[-1] == ".":
            bookmark_name = bookmark_name[0:-1]
        list_of_pdf_obj[bookmark_name] = PdfFileReader(open(full_path, 'rb'), False)
    return list_of_pdf_obj


def combine_and_bookmark(file_dict, pdfs):
    out = PdfFileWriter()
    added_bookmarks = {}
    file_nums = list(file_dict.keys())
    counter = 0
    for name in pdfs:
        pdf = pdfs[name]
        pdf_num = int(file_nums[counter])
        out.addPage(pdf.getPage(0))
        if pdf_num in added_bookmarks:
            out.addBookmark(name, out.getNumPages() - 1, added_bookmarks[pdf_num])
        else:
            b = out.addBookmark(name, out.getNumPages() - 1)
            added_bookmarks[pdf_num] = b
        for page_num in range(1, pdf.getNumPages()):
            out.addPage(pdf.getPage(page_num))
        counter += 1
    return out


def main():
    opts = argument_handler()
    directory_path = choose_directory(opts)
    file_dict = directory_to_files_to_dict(directory_path)
    pdfs = open_all_pdfs(directory_path, file_dict)
    output = combine_and_bookmark(file_dict, pdfs)

    with open(os.path.join(directory_path, 'FullReport.pdf'), "wb") as w:
        output.write(w)


if __name__ == '__main__':
    main()
