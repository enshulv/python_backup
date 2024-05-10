import os
import shutil
import subprocess
import json
from datetime import datetime
import tempfile

def get_dir_size(directory):
    """
    计算指定目录的大小。
    :param directory: 目录路径。
    :return: 目录的总大小（字节）。
    """
    total_size = 0
    for dirpath, dirnames, filenames in os.walk(directory):
        for f in filenames:
            fp = os.path.join(dirpath, f)
            if os.path.exists(fp):
                total_size += os.path.getsize(fp)
    return total_size

def compress_with_haozip(file_list, output_directory,base_directory):
    """
    使用HaoZip命令行工具压缩文件，如果文件总大小超过4GB，则分卷压缩。
    :param file_list: 要压缩的文件列表。
    :param output_directory: 压缩文件存储的目录。
    """
    if not os.path.exists(output_directory):
        os.makedirs(output_directory)

    # 获取当前日期，格式为YYYYMMDD
    current_date = datetime.now().strftime("%Y%m%d")
    password = current_date  # 使用当前日期作为密码

    # 创建临时目录来存放即将被压缩的文件
    with tempfile.TemporaryDirectory(dir='E:\\') as temp_dir:
        zip_contents = []  # 存储当前压缩包的内容

        for file in file_list:
            # 保持文件的原始文件夹结构并复制文件到临时目录
            relative_path = os.path.relpath(file, start=base_directory)
            dest_path = os.path.join(temp_dir, relative_path)
            os.makedirs(os.path.dirname(dest_path), exist_ok=True)  # 创建所需的文件夹结构
            shutil.copy(file, dest_path)
            zip_contents.append(dest_path)

        # 计算临时目录的大小
        temp_dir_size = get_dir_size(temp_dir)
        max_size = 4 * 1024 * 1024 * 1024  # 4GB

        # 压缩包文件名
        zip_filename = os.path.join(output_directory, f"藏书备份-{current_date}.zip")
        haozip_path = 'E:\\HaoZip\\HaoZipC.exe'

        # 根据文件总大小决定是否分卷压缩
        if temp_dir_size > max_size:
            volume_size = "4g"  # 设置分卷大小为4GB
            subprocess.run([haozip_path, "a", "-p" + password, "-tzip", "-v" + volume_size, zip_filename, temp_dir + "\\*"])
        else:
            subprocess.run([haozip_path, "a", "-p" + password, "-tzip", zip_filename, temp_dir + "\\*"])

        # 保存压缩包内容到JSON文件
        save_zip_contents(zip_contents, output_directory, current_date)

def save_zip_contents(contents, output_directory, current_date):
    """
    保存压缩包内容到JSON文件。
    :param contents: 压缩包中的文件列表。
    :param output_directory: 储存JSON文件的目录。
    :param current_date: 当前日期。
    """
    json_filename = os.path.join(output_directory, f"藏书备份-{current_date}-目录.json")
    with open(json_filename, 'w', encoding='utf-8') as f:
        json.dump([os.path.basename(file) for file in contents], f, indent=4)

def compress_individual_files(file_list, output_directory):
    """
    将文件列表中的每个文件单独压缩，并将压缩包保存到指定目录。
    :param file_list: 要压缩的文件列表。
    :param output_directory: 压缩文件存储的目录。
    """
    if not os.path.exists(output_directory):
        os.makedirs(output_directory)

    current_date = datetime.now().strftime("%Y%m%d")
    file_mapping = {}
    haozip_path = 'E:\\HaoZip\\HaoZipC.exe'
    max_size = 4 * 1024 * 1024 * 1024

    for index, file in enumerate(file_list):
        temp_dir_size = get_dir_size(file)
        # 生成压缩包文件名
        name = f"{index}_{current_date}.zip"
        zip_filename = os.path.join(output_directory,name)
        password = current_date  # 使用日期作为密码

        # 压缩单个文件
        if temp_dir_size > max_size:
            volume_size = "4g"  # 设置分卷大小为4GB
            subprocess.run([haozip_path, "a", "-p" + password, "-tzip","-v" + volume_size, zip_filename, file + "\\*"])
        else:
            subprocess.run([haozip_path, "a", "-p" + password, "-tzip", zip_filename, file + "\\*"])

        # 添加到映射表
        file_mapping[name] = os.path.basename(file)

    # 保存映射表为JSON文件
    mapping_file_path = os.path.join(output_directory, f"file_mapping_{current_date}.json")
    with open(mapping_file_path, 'w', encoding='utf-8') as f:
        json.dump(file_mapping, f, indent=4, ensure_ascii=False)

    return file_mapping

if __name__ == '__main__':
    directory = 'E:\ghs\game'
    file_list = [os.path.join(directory,a) for a in os.listdir(directory)]  # 添加文件列表
    output_directory = 'E:\Desktop\备份\ghs'
    file_mapping = compress_individual_files(file_list, output_directory)
    print(file_mapping)
