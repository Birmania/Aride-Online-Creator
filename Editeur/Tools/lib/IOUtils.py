import os

def ReadFile(path):
    file = open(path,"rb")

    lines = file.readlines()

    file.close()

    return lines

def WriteFile(path, lines):
    file = open(path,"wb")

    # Remove return character
    lines = [line.replace("\r","") for line in lines]
    lines = [line.replace("\n","") for line in lines]
    file.writelines([line+os.linesep for line in lines])

    file.close()