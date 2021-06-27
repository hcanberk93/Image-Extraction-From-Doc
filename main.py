import docx2txt
import os
import shutil as sh

def getListOfFiles(dirName):
    # create a list of file and sub directories
    # names in the given directory
    listOfFile = sorted(os.listdir(dirName))
    allFiles = list()
    # Iterate over all the entries
    for entry in listOfFile:
        # Create full path

        fullPath = os.path.join(dirName, entry)
        # If entry is a directory then get the list of files in this directory
        if os.path.isdir(fullPath):
            allFiles = allFiles + getListOfFiles(fullPath)
        else:
            filename, file_extension = os.path.splitext(entry)
            if file_extension == '.docx':
                allFiles.append(fullPath)
    return allFiles

list =getListOfFiles("Kaynak/")

fileNo = 1
hasCode = False
for file in list:
    hasCode = False
    destPath = file.split('/')
    destPath.remove(destPath[-1])
    destPath = '/'.join(destPath)
    text = docx2txt.process(file, destPath)
    for line in text.splitlines():
        if ("Ürün Kodu" in line or "Ürün kodu" in line or "ÜRÜN KODU" in line):
            hasCode = True
            fileName = ''.join(line.split(':')[-1].split())

            if '/' in fileName:
                fileName = '-'.join(fileName.split('/'))
            print(fileName)

    if not hasCode:
        fileName = 'NoNumber' + str(fileNo)
        fileNo = fileNo +1
        filePath = destPath
        filePath = file.split('/')
        filePath.remove(filePath[-1])
        filePath.remove(filePath[-1])
        filePath.append('kodsuz')
        filePath = '/'.join(filePath)
        print(filePath)
        sh.copy(file,filePath)

    for x in os.listdir(destPath):
        if 'image1' in x:
            fileExt = x.split('.')[-1]

    os.rename(destPath+'/image1.'+fileExt,destPath+'/'+fileName)


