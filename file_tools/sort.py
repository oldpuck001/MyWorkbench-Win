# sort.py

import os
import shutil

def sort(request):

    location_character = request.get("data", {}).get("location_character", "")
    cut_or_copy = request.get("data", {}).get("cut_or_copy", "")
    sourcePath = request.get("data", {}).get("sourcePath", "")
    targetPath = request.get("data", {}).get("targetPath", "")

    filenameSet = set()
    result_text = ''

    if cut_or_copy == 'cut':
        if location_character != '':

            try:
                location_character = int(location_character)
            except:
                return ['filename_sort', '参数输入有误，请重新输入！']

            for filename in os.listdir(sourcePath):
                if filename != '.DS_Store':
                    filenameSet.add(filename[:location_character])
            for fileset in filenameSet:
                for filename in os.listdir(sourcePath):
                    if filename[:location_character] == fileset:
                        resultpath = f'{targetPath}/{fileset}'
                        os.makedirs(resultpath, exist_ok = True)
                        sourcefilePath = os.path.join(sourcePath, filename)
                        targetfilePath = os.path.join(resultpath, filename)
                        shutil.move(sourcefilePath,targetfilePath)
                        result_text = result_text + sourcePath + '/' + filename + ' move ' + resultpath + '/' + filename + '\n'
            if result_text == '':
                result_text = '未找到可移动文件或文件夹！'
        else:
            result_text = '参数输入有误，请重新输入！'
        
    elif cut_or_copy == 'copy':
        if location_character != '':

            try:
                location_character = int(location_character)
            except:
                return ['filename_sort', '参数输入有误，请重新输入！']
    
            for filename in os.listdir(sourcePath):
                if filename != '.DS_Store':
                    filenameSet.add(filename[:location_character])
            for fileset in filenameSet:
                for filename in os.listdir(sourcePath):
                    if filename[:location_character] == fileset:
                        resultpath = f'{targetPath}/{fileset}'
                        os.makedirs(resultpath, exist_ok = True)
                        sourcefilePath = os.path.join(sourcePath, filename)
                        targetfilePath = os.path.join(resultpath, filename)
                        shutil.copy(sourcefilePath,targetfilePath)
                        result_text = result_text + sourcePath + '/' + filename + ' copy ' + resultpath + '/' + filename + '\n'
            if result_text == '':
                result_text = '未找到可复制文件或文件夹！'
        else:
            result_text = '参数输入有误，请重新输入！'

    return ['filename_sort', result_text]