import os
import json

def generate_full_path_file_list(root_path):
    """
    生成给定目录下所有符合特定格式的文件的完整路径列表。
    :param root_path: 要扫描的目录路径。
    :return: 包含完整路径的文件列表。
    """
    supported_formats = {'.docx', '.pdf', '.epub', '.txt', '.png', '.jpg', '.webp'}
    full_path_list = []

    for root, dirs, files in os.walk(root_path):
        for file in files:
            if any(file.endswith(ext) for ext in supported_formats):
                full_path_list.append(os.path.join(root, file))

    return full_path_list

def save_list_to_json(file_list, filename):
    """
    将文件列表保存为JSON格式。
    :param file_list: 文件列表。
    :param filename: 要保存的文件名。
    """
    with open(filename, 'w', encoding='utf-8') as f:
        json.dump(file_list, f, ensure_ascii=False, indent=4)



