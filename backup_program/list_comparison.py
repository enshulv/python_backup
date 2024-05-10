import json

def load_file_list(filename):
    """
    从JSON文件中加载文件列表。
    :param filename: JSON文件的路径。
    :return: 文件列表。
    """
    try:
        with open(filename, 'r', encoding='utf-8') as f:
            return json.load(f)
    except FileNotFoundError:
        return []

def compare_file_lists(old_list, new_list):
    """
    比较旧列表和新列表，找出新列表中新增的文件。
    :param old_list: 旧的文件列表。
    :param new_list: 新的文件列表。
    :return: 新增的文件列表。
    """
    set_old = set(old_list)
    set_new = set(new_list)
    return list(set_new - set_old)  # 返回新增的文件

