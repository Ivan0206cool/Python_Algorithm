import shutil
import ntpath
import re
import os
def copy_file_dest(src_file_path, dest_file_path):
    isExist = os.path.exists(dest_file_path)
    if not isExist:
        os.makedirs(dest_file_path)
    shutil.copy(src_file_path,dest_file_path)

def get_file_name(file_path):
    return ntpath.basename(file_path)

# print(get_file_name("../error_logs/ccpa/Report_logccpa1111.txt"))
