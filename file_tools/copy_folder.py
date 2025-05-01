# copy_folder.py

import os

def copy_folder(request):

    # 获取前端传递的参数
    sourcePath = request.get("data", {}).get("sourcePath", "")
    targetPath = request.get("data", {}).get("targetPath", "")

    # 定义递归函数，用于遍历文件夹及其子文件夹
    def copy_structure_recursive(src, dst):

        # 创建目标文件夹（如果不存在）
        os.makedirs(dst, exist_ok=True)

        # 遍历源文件夹中的所有子文件夹
        for item in os.listdir(src):
            src_item = os.path.join(src, item)
            if os.path.isdir(src_item):
                # 如果是子文件夹，递归调用
                dst_item = os.path.join(dst, item)
                copy_structure_recursive(src_item, dst_item)


    # 开始复制文件夹结构
    copy_structure_recursive(sourcePath, targetPath)

    return ['copy_folder']