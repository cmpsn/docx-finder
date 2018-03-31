#! /usr/bin/env python
# -*- coding: utf-8 -*-
'''
Search for a word or expression in docx documents from an input directory
and return the paths to the files where the word/expression is found.
'''

import os
import sys
import re
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
        sys.exit('''\nThere is already a subfolder with this name.
Please wait a sec and try again.''')
    except Exception:
        sys.exit('''\nUser does not have privileges to create directories
in the current folder. The program was aborted.''')
else:
    print('\nThank you for trying! Please come back when you are ready.')
    sys.exit()

# setting the path to the folder to search in
while True:
    calea = input('''\n[To quit -> Press Enter]
To continue -> Type a full path to search in. \n''')
    if calea == '':
        os.rmdir(outDir)
        sys.exit('\nThank you for trying!Please come back when you are ready.')
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
cuv = input('''\n[To quit -> Press Enter]
To continue -> Type a word or an expression to search for. \n''')
if cuv == '' or cuv == ' ' or cuv in '!"$„”“%&\'()*+–,./:;<=>?@[\]^_`{|}~':
    sys.exit('\nThank you for trying! Please come back when you are ready.')
else:
    print('\nWorking...')

# setting the paths for output files (inside the output folder)
FileText = os.path.join(outDir, 'FileText.txt')
FoundFiles = os.path.join(outDir, 'FoundFiles.txt')
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
num_matches = 0
for filepath in listaFis:
    count_tot += 1
    with open(FileText, 'w') as fit:
        try:
            if os.path.isfile(filepath):
                fit.write(getText(filepath))
                count_op += 1
            else:
                errorFile = open(ErrorInfo, 'a')
                errorFile.write('Path: %s is not a file.\n' % filepath)
                errorFile.close()
                # print('\nTraceback info sent to ErrorInfo.txt.')
        except (KeyError, ValueError):
            unopefil = open(NotOpFiles, 'a')
            unopefil.write('Can\'t open file: %s\n' % filepath)
            unopefil.close()
            count_unop += 1
            # print('\nPath to a non-valid docx file sent to NotOpFiles.txt.')
        except Exception:
            errorFile = open(ErrorInfo, 'a')
            errorFile.write(format_exc())
            errorFile.write('\n')
            errorFile.close()
            count_unop += 1
            # print('\nTraceback info sent to ErrorInfo.txt.')
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
                num_matches += len(lstcuv)
            cuvlst = open(Matches, 'a')
            cuvlst.write('\nItems extracted from file - %s: \n' % filepath)
            cuvlst.write('\n'.join(lstcuv))
            cuvlst.close()
            # write the path to each file containing the input word
            fislst = open(FoundFiles, 'a')
            fislst.write('\nText "%s" found in file: %s\n' % (cuv, filepath))
            fislst.close()

# remove the temporary file and print the counters for files
os.remove(FileText)
print('\nDocx files along the path:', count_tot)
print('Files processed:', count_op)
print('Files not processed (docx not valid):', count_unop)
print('Files with matches:', count_foundfiles)
print('Total word/expression matches:', num_matches)

# check if the output files exists and print the final info
match_f_ex = os.path.exists(Matches)
found_f_ex = os.path.exists(FoundFiles)
notOp_f_ex = os.path.exists(NotOpFiles)
errInf_f_ex = os.path.exists(ErrorInfo)

if match_f_ex or found_f_ex or notOp_f_ex or errInf_f_ex:
    print('''\nTo check the results, open the txt files created inside
the most recent "Ouput_..." folder (from your current directory):\n''')
else:
    os.rmdir(outDir)

if match_f_ex:
    print(' - The list of words/expressions that match is in "Matches.txt".\n')
if found_f_ex:
    print(' - The paths to files with matches are in "FoundFiles.txt".\n')
if notOp_f_ex:
    print(' - The paths to not valid ".docx" files are in "NotOpFiles.txt".\n')
if errInf_f_ex:
    print(' - The traceback info is in ErrorInfo.txt.\n')
