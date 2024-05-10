import os
import hashlib

def split_and_save_file(file_path, output_dir, chunk_size=5*1024*1024):
    """
    将文件分割成分片，并在同名文件夹中保存每个分片为二进制文件，同时计算每个分片的MD5值。
    :param file_path: 要分割的文件的路径。
    :param output_dir: 总体的输出目录。
    :param chunk_size: 分片大小，默认为4MB。
    :return: 分片的MD5值列表。
    """
    file_base_name = os.path.splitext(os.path.basename(file_path))[0]
    # 创建一个同名的文件夹来存储分片
    chunk_folder = os.path.join(output_dir, file_base_name)
    if not os.path.exists(chunk_folder):
        os.makedirs(chunk_folder)

    block_list = []

    with open(file_path, 'rb') as f:
        chunk_index = 0
        chunk = f.read(chunk_size)
        while chunk:
            # 计算分片的MD5值
            md5 = hashlib.md5(chunk).hexdigest()
            block_list.append(md5)

            # 保存分片为二进制文件
            chunk_file_name = f"{md5}.bin"
            chunk_file_path = os.path.join(chunk_folder, chunk_file_name)
            with open(chunk_file_path, 'wb') as chunk_file:
                chunk_file.write(chunk)

            # 准备下一个分片
            chunk_index += 1
            chunk = f.read(chunk_size)

    return block_list

# 示例使用
file_path = ''
output_dir = 'E:\Desktop\备份'
block_list = split_and_save_file(file_path, output_dir)
print(block_list)
