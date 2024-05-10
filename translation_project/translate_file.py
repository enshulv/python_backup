from translate_tools import translate
from full_formula_translate import run_translate
import time
import os
import shutil

def run_translate_process(name):
    source_path = 'E:\Desktop\Files\Save//'
    save_path = 'E:\Desktop\Files\Results//'
    try:
        # Translation engine for the main text, skipping references, table translate engine, one-to-one translate
        if 'en-' in name[:5]:
            translate(name, 'en', source_path, save_path, 1, 0, 1, 0)
        elif 'qk-' in name[:5]:
            translate(name, 'zh', source_path, save_path, 1, 0, 1, 0)
        elif 'gs-' in name[:5]:
            run_translate('zh', name, source_path, save_path)
        elif 'gn-' in name[:5]:
            run_translate('en', name, source_path, save_path)
        else:
            translate(name, 'zh', source_path, save_path, 0, 0, 1, 0)
    except Exception as error:
        print('\n' + str(error))

def run():
    source_path = 'E:\Desktop\Files\Save//'
    backup_path = 'E:\Desktop\Files\Save//Backup//'
    files_to_translate = os.listdir(source_path)
    for file in files_to_translate:
        file_path = source_path + file
        backup_file_path = backup_path + file
        if 'docx' in file and '~$' not in file:
            name = file.replace('.docx', '')
            print(name)
            run_translate_process(name)
            shutil.move(file_path, backup_file_path)

if __name__ == '__main__':
    run()