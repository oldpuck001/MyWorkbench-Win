# image.py

import os

def image(request):

    select_function = request.get("data", {}).get("select_function", "")

    result_text = ''

    if select_function == 'select':
        filePath = request.get("data", {}).get("filePath", "")
        folderPath = os.path.dirname(filePath)
        imageList = [os.path.join(folderPath, f) for f in os.listdir(folderPath) if f.endswith(('.png', '.jpg', '.jpeg'))]
        imageList.sort()
        num = imageList.index(filePath)
        result_text = [num, imageList]
        return ['filename_image', ['select', result_text]]

    elif select_function == 'rename':
    
        num = request.get("data", {}).get("num", "")
        imageList = request.get("data", {}).get("imagelist", "")
        name_1 = request.get("data", {}).get("name_1", "")
        name_2 = request.get("data", {}).get("name_2", "")
        name_3 = request.get("data", {}).get("name_3", "")
        name_4 = request.get("data", {}).get("name_4", "")
        name_5 = request.get("data", {}).get("name_5", "")
        result_text = request.get("data", {}).get("result_text", "")

        if name_1 == '' and name_2 == '' and name_3 == '' and name_4 == '' and name_5 == '':
            result_text = result_text + '请输入新文件名！' + '\n'
            return ['filename_image', ['no_new', result_text]]

        else:
            extfilename = os.path.splitext(imageList[num])[1]
            newfilename = ''
            if name_1 != '':
                newfilename = name_1
            if name_2 != '':
                newfilename = newfilename + '-' + name_2
            if name_3 != '':
                newfilename = newfilename + '-' + name_3
            if name_4 != '':
                newfilename = newfilename + '-' + name_4
            if name_5 != '':
                newfilename = newfilename + '-' + name_5
            newfilename = newfilename.strip('-')                                    # 如過第一個輸入框是空白的，文件名第一個字符會是“-”，使用strip刪除
            newfilename = newfilename + extfilename
            newFilePath = os.path.join(os.path.dirname(imageList[num]), newfilename)
            os.rename(imageList[num], newFilePath)

            oldfilename = imageList[num].split('/')[-1]
            result_text = result_text + oldfilename + ' to ' + newfilename + '\n'
            
            imageList[num] = newFilePath

            return ['filename_image', ['rename', [num, imageList, result_text]]]