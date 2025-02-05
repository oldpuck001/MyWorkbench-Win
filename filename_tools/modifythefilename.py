# modifythefilename.py

import os
import re

def modify(request):
    select_function = request.get("data", {}).get("select_function", "")
    source_character = request.get("data", {}).get("source_character", "")
    location_character = request.get("data", {}).get("location_character", "")
    target_character = request.get("data", {}).get("target_character", "")
    folderPath = request.get("data", {}).get("folderPath", "")

    result_text = ''

    if select_function == 'find':

        if location_character != '':
            try:
                int(location_character)
            except:
                location_character = -1

        if location_character == '':
            for filename in os.listdir(folderPath):
                mainfilename = os.path.splitext(filename)[0]
                if source_character in mainfilename:
                    result_text = result_text + filename + '\n'
            if result_text == '':
                result_text = '未找到！'

        elif int(location_character) >= 1:
            location_character = int(location_character) - 1
            for filename in os.listdir(folderPath):
                if source_character == filename[location_character:location_character+len(source_character)]:
                    result_text = result_text + filename + '\n'
            if result_text == '':
                result_text = '未找到！'
        else:
            result_text = '参数输入有误，请重新输入！'

    elif select_function == 'add':

        try:
            int(location_character)
        except:
            location_character = -1

        if location_character == '':
            result_text = '参数输入有误，请重新输入！'

        elif int(location_character) == 0:
            for filename in os.listdir(folderPath):
                newfilename = source_character + filename
                oldfilePath = os.path.join(folderPath, filename)
                newfilePath = os.path.join(folderPath, newfilename)
                os.rename(oldfilePath, newfilePath)
                result_text = result_text + filename + ' to ' + newfilename + '\n'
            if result_text == '':
                result_text = '未找到可修改文件或文件夹！'

        elif int(location_character) >= 1:
            for filename in os.listdir(folderPath):
                if int(location_character) <= len(os.path.splitext(filename)[0]):
                    newfilename = filename[:int(location_character)-1] + source_character + filename[int(location_character)-1:]
                    oldfilePath = os.path.join(folderPath, filename)
                    newfilePath = os.path.join(folderPath, newfilename)
                    os.rename(oldfilePath, newfilePath)
                    result_text = result_text + filename + ' to ' + newfilename + '\n'
            if result_text == '':
                result_text = '未找到可修改文件或文件夹！'

        else:
            result_text = '参数输入有误，请重新输入！'

    elif select_function == 'del':

        if location_character != '':
            try:
                int(location_character)
            except:
                location_character = -1

        if location_character == '':
            for filename in os.listdir(folderPath):
                mainfilename = os.path.splitext(filename)[0]
                extfilename = os.path.splitext(filename)[1]
                if source_character in mainfilename:
                    newfilename = mainfilename.replace(source_character, '') + extfilename
                    oldfilePath = os.path.join(folderPath, filename)
                    newfilePath = os.path.join(folderPath, newfilename)
                    os.rename(oldfilePath, newfilePath)
                    result_text = result_text + filename + ' to ' + newfilename + '\n'
            if result_text == '':
                result_text = '未找到可修改文件或文件夹！'

        elif int(location_character) >= 1:
            try:
                int(source_character)
            except:
                source_character = -1

            if int(source_character) >= 1:
                for filename in os.listdir(folderPath):
                    mainfilename = os.path.splitext(filename)[0]
                    extfilename = os.path.splitext(filename)[1]
                    newfilename = filename[:int(location_character)-1] + filename[int(location_character)-1+int(source_character):]
                    oldfilePath = os.path.join(folderPath, filename)
                    newfilePath = os.path.join(folderPath, newfilename)
                    os.rename(oldfilePath, newfilePath)
                    result_text = result_text + filename + ' to ' + newfilename + '\n'
                if result_text == '':
                    result_text = '未找到可修改文件或文件夹！'
            
            else:
                result_text = '参数输入有误，请重新输入！'

    elif select_function == 'replace':

        if location_character != '':
            try:
                int(location_character)
            except:
                location_character = -1

        if location_character == '':
            for filename in os.listdir(folderPath):
                mainfilename = os.path.splitext(filename)[0]
                extfilename = os.path.splitext(filename)[1]
                if source_character in mainfilename:
                    newfilename = mainfilename.replace(source_character, target_character) + extfilename
                    oldfilePath = os.path.join(folderPath, filename)
                    newfilePath = os.path.join(folderPath, newfilename)
                    os.rename(oldfilePath, newfilePath)
                    result_text = result_text + filename + ' to ' + newfilename + '\n'
            if result_text == '':
                result_text = '未找到可修改文件或文件夹！'

        elif int(location_character) >= 1:
            for filename in os.listdir(folderPath):
                if source_character == filename[int(location_character)-1:int(location_character)-1+len(source_character)]:
                    newfilename = filename[:int(location_character)-1] + target_character + filename[int(location_character)-1+len(source_character):]
                    oldfilePath = os.path.join(folderPath, filename)
                    newfilePath = os.path.join(folderPath, newfilename)
                    os.rename(oldfilePath, newfilePath)
                    result_text = result_text + filename + ' to ' + newfilename + '\n'
            if result_text == '':
                result_text = '未找到可修改文件或文件夹！'
            
        else:
            result_text = '参数输入有误，请重新输入！'

    elif select_function == 'regex':

        try:
            if target_character == '':
                for filename in os.listdir(folderPath):
                    mainfilename = os.path.splitext(filename)[0]
                    regexcompile = re.compile(source_character)
                    if re.search(regexcompile,mainfilename):
                        result_text = result_text + filename + '\n'
                if result_text == '':
                    result_text = '未找到！'

            else:
                for filename in os.listdir(folderPath):
                    mainfilename = os.path.splitext(filename)[0]
                    extfilename = os.path.splitext(filename)[1]
                    regexcompile = re.compile(source_character)
                    if re.search(regexcompile, mainfilename):
                        newfilename = re.sub(regexcompile, target_character, mainfilename) + extfilename
                        oldfilePath = os.path.join(folderPath, filename)
                        newfilePath = os.path.join(folderPath, newfilename)
                        os.rename(oldfilePath, newfilePath)
                        result_text = result_text + filename + ' to ' + newfilename + '\n'
                if result_text == '':
                    result_text = '未找到可修改文件或文件夹！'
        except:
            result_text = '参数输入有误，请重新输入！'

    return ['filename_modify', result_text]