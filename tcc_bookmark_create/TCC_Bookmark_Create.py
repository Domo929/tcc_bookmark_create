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


# This reads in a directory path and returns a dictionary of the pdf number keying the filename
def directory_to_files_to_dict(directory_path):
    # Pulls the float number from the front of the pdf title. Includes a possible letter or two
    num_match = r"([\d]+.[\w\d]*)"
    temp = {}

    # Iterate over the list of files in the directory
    for f in os.listdir(directory_path):
        # Check to make sure it's a file, and check that the title doesn't include the '_INITIAL ' name
        if os.path.isfile(os.path.join(directory_path, f)) and initial_flag not in f:
            # Pull the number out and store it
            pdf_num = re.match(num_match, f).group(0)
            # Store the converted float of the number as the key to the filename
            temp[string_converter(pdf_num)] = f

    return_dict = {}

    # Because os.listdir doesn't like in order, we need to sort it after
    for key in sorted(list(temp)):
        return_dict[key] = temp[key]
    return return_dict


# Converts a string of numbers and letters to an equivalent float number for sorting
def string_converter(word):
    list_of_char = []

    # Check each character individually to see if it's a . or a letter
    for char in word:
        # Force uppercase
        char.upper()

        # We want to add the period to the number string to act as the anchor for the float
        if char == ".":
            list_of_char.append(".")

        # Otherwise we convert the letters to their respective numbers (A=1, B=2, ...)
        else:
            list_of_char.append(str_dict[char])

    # Combine the characters and convert it to a float
    num = float("".join(list_of_char))
    return num


# Opens the PdfFileReader objects of the filenames and store them in a dict with the bookmark name as the key
def open_all_pdfs(directory, file_dict):
    # This is one hell of a grep line, but it works.
    # One issue, is that due to the table numbers, sometimes you end up with a trailing '.' on the name.
    # This will be fixed later
    bookmark_grep = r"([a-zA-Z\d.]*[ -]*[a-zA-Z& ]+[\d]*[.]?[\d]*)"
    list_of_pdf_obj = {}

    # Iterate through the dictionary of files, opening them as PdfFileReader objects
    for num, filename in file_dict.items():
        full_path = os.path.join(directory, filename)
        bookmark_name = re.search(bookmark_grep, filename).group(1)

        # If the last letter of the name is '.' then we remove it
        if bookmark_name[-1] == ".":
            bookmark_name = bookmark_name[0:-1]

        # Open the pdf and store it in the dict
        list_of_pdf_obj[bookmark_name] = PdfFileReader(open(full_path, 'rb'), False)
    return list_of_pdf_obj


# This actually builds the final pdf with bookmarks
def combine_and_bookmark(file_dict, pdfs):
    # Create the writer object
    out = PdfFileWriter()

    # This is used to track what bookmarks have been added, in order to add parent bookmarks as needed
    added_bookmarks = {}

    # Gives the numbers to store as keys in added_bookmarks
    file_nums = list(file_dict.keys())
    counter = 0

    # Do this for every PDF we've opened
    for name, pdf in pdfs.items():
        # Determine the number of the pdf chapter
        pdf_num = int(file_nums[counter])

        # Add the first page
        out.addPage(pdf.getPage(0))

        # If we already added a pdf bookmark from this chapter:
        if pdf_num in added_bookmarks:
            # We add the bookmark with the parent of the root of the chapter
            out.addBookmark(name, out.getNumPages() - 1, added_bookmarks[pdf_num])

        # Otherwise if we haven't added a bookmark from this chapter yet
        else:
            # Add the bookmark, and make sure to add that bookmark to the dict above
            b = out.addBookmark(name, out.getNumPages() - 1)
            added_bookmarks[pdf_num] = b
        # Then, we iterate through the rest of the pages and add the rest
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
