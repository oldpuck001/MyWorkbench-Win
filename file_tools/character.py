# character.py

import os

def character(request):

    location_character = request.get("data", {}).get("location_character", "")
    folderPath = request.get("data", {}).get("folderPath", "")

    filenameSet = set()
    result_text = ''

    if location_character != '':

        try:
            location_character = int(location_character)
        except:
            location_character = -1

        if location_character > 1:
            for filename in os.listdir(folderPath):
                filenameSet.add(filename[:location_character])

            for SpecifyCharactersfile in filenameSet:
                number = 0
                for filename in os.listdir(folderPath):
                    if filename[:location_character] == SpecifyCharactersfile:
                        number += 1
                        mainfilename = os.path.splitext(filename)[0]
                        extfilename = os.path.splitext(filename)[1]
                        newfilename = mainfilename[:location_character] + '-' + str(number) + extfilename
                        oldfilePath = os.path.join(folderPath, filename)
                        newfilePath = os.path.join(folderPath, newfilename)
                        os.rename(oldfilePath, newfilePath)
                        result_text = result_text + filename + ' to ' + newfilename + '\n'
            if result_text == '':
                result_text = '未找到可修改文件或文件夹！'

        else:
            result_text = '参数输入有误，请重新输入！'

    else:
        result_text = '参数输入有误，请重新输入！'

    return ['filename_character', result_text]