import argparse
import tkinter as tk
from os import listdir
from os.path import isfile, join
from tkinter import filedialog, Listbox

initial_flag = '_INITIAL '


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
    files = [f for f in listdir(directory_path) if isfile(join(directory_path, f)) and initial_flag not in f]

    root = tk.Tk()
    scrollbar = tk.Scrollbar(root, orient='vertical')

    listbox = Listbox(root, width=50, height=50, yscrollcommand=scrollbar.set, selectmode=tk.MULTIPLE)
    scrollbar.config(command=listbox.yview)
    scrollbar.pack(side='right', fill='y')
    listbox.pack(side='left', fill='both', expand=True)
    b = tk.Button(root, text='Delete',
                  command=lambda lb=listbox: lb.delete(tk.ANCHOR))
    listbox.pack()
    for f in files:
        listbox.insert(tk.END, f)

    root.mainloop()


def main():
    opts = argument_handler()
    directory_path = choose_directory(opts)
    file_selection(directory_path)


if __name__ == '__main__':
    main()
