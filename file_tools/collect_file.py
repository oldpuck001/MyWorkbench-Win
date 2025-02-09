# collect_file.py

import os
import shutil

def collect_file(request):
    
    sourcePath = request.get("data", {}).get("sourcePath", "")
    targetPath = request.get("data", {}).get("targetPath", "")
    
    result_text = ''

    # 定义递归函数，用于遍历文件夹及其子文件夹
    def get_files_recursive(folder_path):

        file_list = []

        for root, dirs, files in os.walk(folder_path):
            # 添加当前文件夹中的文件信息
            for file in files:
                file_path = os.path.join(root, file)
                file_list.append(file_path)
        return file_list


    items = get_files_recursive(sourcePath)

    for file in items:

        sourcefilePath = file
        file_name = os.path.basename(sourcefilePath)
        targetfilePath = os.path.join(targetPath, file_name)
        
        # 检查文件是否重名，并生成唯一文件名
        base_name, extension = os.path.splitext(targetfilePath)
        counter = 1
        while os.path.exists(targetfilePath):
            targetfilePath = f"{base_name}_{counter}{extension}"
            counter += 1

        shutil.copy(sourcefilePath,targetfilePath)
        result_text = result_text + sourcefilePath + ' copy ' + targetfilePath + '\n'

    return ['collect_file', result_text]