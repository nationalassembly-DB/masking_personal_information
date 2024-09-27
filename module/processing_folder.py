"""
입력받은 폴더를 순회하면서 pdf, hwp, xlsx 파일의 경로를 구하고 extract_information으로 넘깁니다
"""

import os
from natsort import natsorted


from module.create_excel import create_excel
from module.extract_information import processing_search


def processing_folder(folder_path, excel_file):
    """폴더 내부를 순회하며, pdf, hwp, xlsx 파일을 찾아 개인정보를 찾습니다."""
    infos_list = []
    extension_list = {'.pdf', '.hwp', '.hwpx', '.xlsx'}

    for root, _, files in os.walk(folder_path):
        for file in natsorted(files):
            if file.lower().endswith(tuple(extension_list)):
                file_path = os.path.join('\\\\?\\', root, file)
                search_result = processing_search(folder_path, file_path)
                infos_list.extend(search_result)

    create_excel(infos_list, excel_file)
