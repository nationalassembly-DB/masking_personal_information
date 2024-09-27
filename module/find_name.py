"""
위원회명과 피감기관명을 폴더경로에서 찾습니다
"""

import os


def find_cmt_org_name(file, folder_path):
    """위원회명과 피감기관명을 폴더명에서 찾아 반환합니다"""
    if os.path.basename(folder_path).find(' ') != -1:
        if os.path.basename(folder_path).find('_') != -1:
            cmt = os.path.basename(folder_path)[
                os.path.basename(folder_path).find(' ') + 1:os.path.basename(folder_path).find('_')]
        else:
            cmt = os.path.basename(folder_path)[
                os.path.basename(folder_path).find(' ') + 1:]
    else:
        cmt = os.path.basename(folder_path)

    relative_path = os.path.relpath(file, os.path.dirname(folder_path))
    org = relative_path.split(os.sep)[1]

    return cmt, org
