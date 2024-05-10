import 列表比对
import 文件扫描与列表生成 as 扫描
import 文件压缩与加密 as 压缩
from datetime import datetime
import os


def main():
    """
    主程序执行流程。
    """
    # 基本路径
    base_directory = "E:\\Desktop"
    backup_folder = os.path.join(base_directory, "备份")

    # 当前日期
    current_date = datetime.now().strftime("%Y%m%d")

    # 文件列表路径
    old_list_file = os.path.join(backup_folder, "list.json")
    new_list_file = os.path.join(backup_folder, f"list_{current_date}.json")

    # 按日期创建的输出目录
    output_directory = os.path.join(backup_folder, current_date)

    # 需要扫描的目录
    directory_to_scan = base_directory  # 或者任何其他需要扫描的路径

    # 生成当前文件列表
    current_file_list = 扫描.generate_full_path_file_list(directory_to_scan)

    # 加载旧的文件列表
    old_file_list = 列表比对.load_file_list(old_list_file)

    # 比较并获取新增文件
    new_files = 列表比对.compare_file_lists(old_file_list, current_file_list)

    if len(new_files) == 0:
        return

    # 压缩新文件
    压缩.compress_with_haozip(new_files, output_directory, base_directory)



if __name__ == "__main__":
    main()
