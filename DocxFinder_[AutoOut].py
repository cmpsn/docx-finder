#! /usr/bin/env python
# -*- coding: utf-8 -*-
'''
Search for a word or expression inside docx documents
and return the paths to the files where the word/expression is found.
'''

import os
import sys
import re
import docx
from traceback import format_exc
from time import localtime, strftime


def dirSize(inPath):
    '''
    Return the total size (in bytes) of all files in the input directory.
    '''
    size = 0
    for dirpath, dirnames, filenames in os.walk(inPath):
        for filename in filenames:
            size += os.path.getsize(os.path.join(dirpath, filename))
    return size


def filesList(inPath):
    '''
    Return the list of '.docx' files inside in the input directory (paths).
    '''
    suffix = '.docx'
    lstFiles = []
    for dirpath, dirnames, filenames in os.walk(inPath):
        for filename in filenames:
            if filename.endswith(suffix):
                lstFiles.append(os.path.join(dirpath, filename))
    return lstFiles


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
      (named "Output") inside your current directory. In order to do this,
      you have to run the program from a folder where you are allowed
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
    thePath = input('''\n[To quit -> Press Enter]
To continue -> Type a full path to search in. \n''')
    if thePath == '':
        os.rmdir(outDir)
        sys.exit('\nThank you for trying!Please come back when you are ready.')
    elif os.path.exists(thePath):
        if os.path.isdir(thePath):
            print('\nTotal size of the files in the current folder:',
                  dirSize(thePath), 'bytes')
            break
        else:
            print('\nPlease enter the valid path to a folder, not to a file.')
            continue
    else:
        print('\nPlease enter the valid path to an existing folder.')
        continue

# pointing for the word/expression to search for
word = input('''\n[To quit -> Press Enter]
To continue -> Type a word or an expression to search for. \n''')
if word == '' or word == ' ' or word in '!"$„”“%&\'()*+–,./:;<=>?@[\]^_`{|}~':
    sys.exit('\nThank you for trying! Please come back when you are ready.')
else:
    print('\nWorking...')

# setting the paths for output files (inside the output folder)
fileText = os.path.join(outDir, 'FileText.txt')
foundFiles = os.path.join(outDir, 'FoundFiles.txt')
errorInfo = os.path.join(outDir, 'ErrorInfo.txt')
notOpFiles = os.path.join(outDir, 'NotOpFiles.txt')
matches = os.path.join(outDir, 'Matches.txt')

# iterate through the list of file paths, get the text from each file,
# count opened, unopened, found, and total files; count expression matches
listFiles = filesList(thePath)
filesFound = []
count_dict = {'count_tot': len(listFiles), 'count_op': 0, 'count_unop': 0,
              'count_foundfiles': 0, 'num_matches': 0}
for filepath in listFiles:
    if os.path.isfile(filepath):
        try:
            # using regular expressions, search the input expression
            # inside the temporary text file,
            # and returns a list of paths to the files with the input word
            # AND a list of all matches in each file
            with open(fileText, 'w') as textFile:
                textFile.write(getText(filepath))
                count_dict['count_op'] += 1
            with open(fileText, 'r') as textFile:
                reader = textFile.read()
                listMatches = re.findall(word, reader)
                if len(listMatches) > 0:
                    filesFound.append(filepath)
                    count_dict['count_foundfiles'] += 1
                    count_dict['num_matches'] += len(listMatches)
                    with open(matches, 'a') as mtch:
                        mtch.write(
                            '\nItems extracted from file %s : \n' % filepath)
                        mtch.write('\n'.join(listMatches))
                    # write the path to each file which contains the expression
                    with open(foundFiles, 'a') as fnd:
                        fnd.write(
                            '\n"%s" found in files: %s\n' % (word, filepath))
        except (KeyError, ValueError):
            with open(notOpFiles, 'a') as unopFile:
                unopFile.write('Can\'t open file: %s\n' % filepath)
            count_dict['count_unop'] += 1
        except Exception:
            with open(errorInfo, 'a') as errorFile:
                errorFile.write(format_exc())
                errorFile.write('\n')
            count_dict['count_unop'] += 1
    else:
        with open(errorInfo, 'a') as errorFile:
            errorFile.write('Path %s is not a file.\n' % filepath)

# remove the temporary file and print the counters for files
os.remove(fileText)
print('Done!')
print('\nDocx files along the path:', count_dict['count_tot'])
print('Files processed:', count_dict['count_op'])
print('Files not processed (docx not valid):', count_dict['count_unop'])
print('Files with matches:', count_dict['count_foundfiles'])
print('Total word/expression matches:', count_dict['num_matches'])

print('\nFiles with matches:')
if len(filesFound) > 0:
    for fil in filesFound:
        print(fil)
else:
    print("No file with matches.")

# check if the output files exist and print the final info
match_ex = os.path.exists(matches)
found_ex = os.path.exists(foundFiles)
notOp_ex = os.path.exists(notOpFiles)
errInf_ex = os.path.exists(errorInfo)

if match_ex or found_ex or notOp_ex or errInf_ex:
    print('''\nTo check the details, open the txt files created inside
folder "''' + outfold + '''" (in your current directory):\n''')
else:
    print('Nothing to output.')
    os.rmdir(outDir)

if match_ex:
    print(' - The list of expressions that match is in "Matches.txt".\n')
if found_ex:
    print(' - The paths to files with matches are in "FoundFiles.txt".\n')
if notOp_ex:
    print(' - The paths to not valid ".docx" files are in "NotOpFiles.txt".\n')
if errInf_ex:
    print(' - The traceback info is in ErrorInfo.txt.\n')
