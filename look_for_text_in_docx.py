import os
import docx # https://python-docx.readthedocs.io/en/latest/genindex.html

def pr_results(path, to_be_found_string, extension):
    """
    Gets the lists of the file names with the specified file extension
     in all the subdirectories of the given path.
    """
    final_list_file_names = []
    final_list_file_dirs = []
    for r, d, f in os.walk(path):
        for file in f:
            if extension in file:
                final_list_file_names.append(os.path.join(file))
                if r not in  final_list_file_dirs:
                    final_list_file_dirs.append(os.path.join(r))
    
    return result(final_list_file_names, final_list_file_dirs,
        to_be_found_string)

def result(final_list_file_names, final_list_file_dirs, 
    to_be_found_string):
    """
    Opens all the documenst in the list: final_list_file_names and
    prints list of all the directories and names of the docs, where
    the given string has been found.
    """
    to_be_printed = []
    for file_path in final_list_file_dirs:
            os.chdir(file_path)
            files = os.listdir()
            try:
                for i in final_list_file_names:
                    if i in files:
                        document = docx.Document(i)
                        for k in document.paragraphs:
                            if to_be_found_string in k.text:
                                added = [file_path, i]
                                if added not in to_be_printed:
                                    to_be_printed.append(added)
            #if given document is opened by dif user or other exceptions
            except:
                print("Not possible")

    print (to_be_printed)



if __name__ == "__main__":
    path = ''
    to_be_found_string = ""
    extension = ".docx" # not to be changed
    pr_results(path, to_be_found_string, extension)