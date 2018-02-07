#! /usr/bin/python3
# -*- coding: utf-8 -*-
'''
Search for a word or expression in docx documents from an input directory
and return the paths to the files where the word/expression is found.
'''

import os
import re
from sys import exit
from traceback import format_exc
from time import localtime, strftime
import docx


def dirSize(inPath):
    '''
    Return the total size (in bytes) of all files in the input directory.
    '''
    size = 0
    for dirpath, dirnames, filenames in os.walk(inPath):
        for filename in filenames:
            size += os.path.getsize(os.path.join(dirpath, filename))
    return size


def listaFiles(InPath):
    '''
    Return the list of '.docx' files in the input directory (strings).
    '''
    suffix = '.docx'
    listaFiles = []
    for dirpath, dirnames, filenames in os.walk(InPath):
        for filename in filenames:
            if filename.endswith(suffix):
                listaFiles.append(os.path.join(dirpath, filename))
    return listaFiles


def getText(filename):
    '''
    Open a docx file, convert it in a list of paragraphs,
    and return it as a text (string with delimited lines).
    '''
    doc = docx.Document(filename)
    fullText = []
    for para in doc.paragraphs:
        fullText.append(para.text)
    return '\n'.join(fullText)


print('''
      To save some output files, the application will create a new subfolder
      (named "Output") inside the current folder. In order to do this,
      you have to run the program from a folder where the user is allowed
      to make changes/ create directories.
      ''')
# setting the path to the folder where the output files will be saved in
# to avoid folder creation error (due to multiple runnings of the program) -->
# --> exact time is added to the name of the output folder
# (allowing unlimited output folders)
answer = input("Type 'y' to continue/ Type 'n' to quit: ")
answ = answer.lower()
if answ == 'y':
    currDir = os.getcwd()
    outfold = 'Output_' + strftime("%Y-%m-%d_%H-%M-%S", localtime())
    outDir = os.path.join(currDir, outfold)
    try:
        os.mkdir(outDir)
    except FileExistsError:
        exit('''\nThere is already a subfolder with this name.
Please wait a sec and try again.''')
    except Exception:
        exit('''\nUser does not have privileges to create directories
in the current folder. The program was aborted.''')
else:
    print('\nThank you for trying! Please come back when you are ready.')
    exit()

# setting the path to the folder to search in
while True:
    calea = input('''\n[To quit: Press Enter]
Enter the full path to the folder to search in: ''')
    if calea == '':
        os.rmdir(outDir)
        exit('\nThank you for trying! Please come back when you are ready.')
    elif os.path.exists(calea):
        if os.path.isdir(calea):
            print('\nTotal size of the files in the current folder:',
                  dirSize(calea), 'bytes')
            break
        else:
            print('\nPlease enter the valid path to a folder, not to a file.')
            continue
    else:
        print('\nPlease enter the valid path to an existing folder.')
        continue

# pointing for the word/expression to search for
cuv = input('''\n[To quit: Press Enter]
Enter the word or expression to search for: ''')
if cuv == '' or cuv == ' ':
    exit('\nThank you for trying! Please come back when you are ready.')

# setting the paths for output files (inside the output folder)
FileText = os.path.join(outDir, 'FileText.txt')
WordFoundFiles = os.path.join(outDir, 'WordFoundFiles.txt')
ErrorInfo = os.path.join(outDir, 'ErrorInfo.txt')
NotOpFiles = os.path.join(outDir, 'NotOpFiles.txt')
Matches = os.path.join(outDir, 'Matches.txt')

# iterate through the list of file paths, get the text from each file,
# count opened files (count_op) and total files (count_tot)
listaFis = listaFiles(calea)
listaFound = []
count_tot = 0
count_op = 0
count_unop = 0
count_foundfiles = 0
for filepath in listaFis:
    count_tot += 1
    with open(FileText, 'w') as fit:
        try:
            if os.path.isfile(filepath):
                fit.write(getText(filepath))
                count_op += 1
            else:
                errorFile = open(ErrorInfo, 'a')
                errorFile.write('Path: %s is not a file.\n' %filepath)
                errorFile.close()
                print('\nThe traceback info was written to ErrorInfo.txt.')
        except (KeyError, ValueError):
            unopefil = open(NotOpFiles, 'a')
            unopefil.write('Can\'t open file: %s\n' %filepath)
            unopefil.close()
            count_unop += 1
            print('''\nThe path to a non-valid docx file was written to
                  NotOpFiles.txt.''')
        except Exception:
            errorFile = open(ErrorInfo, 'a')
            errorFile.write(format_exc())
            errorFile.write('\n')
            errorFile.close()
            count_unop += 1
            print('\nThe traceback info was written to ErrorInfo.txt.')
    # search the input word using regular expressions
    # inside the temporary text file 'Filetext',
    # and returns a list of paths to the actual files containing the input word
    # AND a list of all matches in each file
    with open(FileText, 'r') as f:
        reader = f.read()
        lstcuv = re.findall(cuv, reader)
        if len(lstcuv) > 0:
            if filepath not in listaFound:
                listaFound.append(filepath)
                count_foundfiles += 1
            cuvlst = open(Matches, 'a')
            cuvlst.write('\nItems extracted from file - %s: \n' %filepath)
            cuvlst.write('\n'.join(lstcuv))
            cuvlst.close()

# write the paths to all files containing the input word
with open(WordFoundFiles, 'a') as fis:
    for pth in listaFound:
        fis.write('Expression "%s" found in file: %s\n' %(cuv, pth))

# remove the temporary file and print the counters for files
os.remove(FileText)
print('\nTotal docx files along the path:', count_tot)
print('\nProcessed files:', count_op)
print('\nUnprocessed files (due to invalid docx metadata):', count_unop)
print('\nFiles with matches:', count_foundfiles)

# print final info
print('\nThe list of matches was written to Matches.txt.')
print('\nThe paths to eligible files were written to WordFoundFiles.txt')
print('''\nTo check the results, open the txt files from the most recent
"Ouput" folder.\n''')