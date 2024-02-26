import paramiko
from scp import SCPClient
import shutil
import glob
import os
import ntpath
from pathlib import Path


# class Product:
#     def __init__(self,product_name:str,import_path:str, export_path:str,get_file_from:str,date:str) -> None:
#         self.product_name = product_name
#         self.import_path = import_path
#         self.export_path = export_path
#         self.get_file_from = get_file_from
#         self.date = date
#         self.get_logs()
#     def get_logs_from_remote(self):
#         ssh = paramiko.SSHClient()
#         ssh.load_system_host_keys()
#         ssh.connect(hostname='192.168.220.172',
#                     username='autotest',
#                     password='loop1234')
#         # SCPCLient takes a paramiko transport as its only argument
#         scp = SCPClient(ssh.get_transport())
#         scp.get(F'{self.import_path}/Report_log.txt', F'{self.import_path}/Report_log_{self.product_name}_{self.date}.txt')
#         scp.get(F'{self.import_path}/Error_log.txt', F'{self.import_path}/Error_log_{self.product_name}_{self.date}.txt')
#     def get_logs_from_local(self):
#         shutil.copyfile(F'{self.import_path}/Report_log.txt',F'{self.export_path}/Report_log_{self.product_name}_{self.date}.txt' )
#         shutil.copyfile(F'{self.import_path}/Error_log.txt',F'{self.export_path}/Error_log{self.product_name}_{self.date}.txt' )
#     def get_logs(self):
#         if self.get_file_from == "remote":
#             self.get_logs_from_remote()
#         elif self.get_file_from == "local":
#             self.get_logs_from_local()

# if __name__ == "__main__":
#     pro_g7800 = Product('G7800','D:/Jason-Huang/Repo/sw-autoit-2020/Test_Case/case_344CCPA/5_GE_ccpa_G7800','../error_logs/G7800',"local","0811")

def get_logs_from_remote(date,dest_folder):
    ssh = paramiko.SSHClient()
    ssh.load_system_host_keys()
    ssh.connect(hostname='192.168.220.172',
                username='autotest',
                password='loop1234')
    # SCPCLient takes a paramiko transport as its only argument
    scp = SCPClient(ssh.get_transport())

    # scp.put('file_path_on_local_machine', 'file_path_on_remote_machine')
    scp.get('D:/AutoTest/sw-autoit-2020/Test_Case/case_344CCPA/5_GE_ccpa/Report_log.txt', Path(F'{dest_folder}/ccpa/Report_logccpa{date}.txt'))
    scp.get('D:/AutoTest/sw-autoit-2020/Test_Case/case_344CCPA/5_GE_ccpa/Error_log.txt', Path(F'{dest_folder}/ccpa/Error_logccpa{date}.txt'))
    # scp.get('D:/AutoTest/sw-autoit-2020/Test_Case/case_344CCPA/5_GE_ccpa_G7800/Report_log.txt', F'../error_logs/g7800/Report_log_g7800_{date}.txt')
    # scp.get('D:/AutoTest/sw-autoit-2020/Test_Case/case_344CCPA/5_GE_ccpa_G7800/Error_log.txt', F'../error_logs/g7800/Error_log_g7800_{date}.txt')
    scp.get('D:/AutoTest/sw-autoit-2020/Test_Case/case_344CCPA/5_GE_ccpb_2ge/Report_log.txt', Path(F'{dest_folder}/ccpb/Report_logccpb{date}.txt'))
    scp.get('D:/AutoTest/sw-autoit-2020/Test_Case/case_344CCPA/5_GE_ccpb_2ge/Error_log.txt', Path(F'{dest_folder}/ccpb/Error_logccpb{date}.txt'))
    scp.get('D:/AutoTest/sw-autoit-2020/Test_Case/case_344CCPA/5_GE_Mix/Report_log.txt', Path(F'{dest_folder}/mix/Report_logmix{date}.txt'))
    scp.get('D:/AutoTest/sw-autoit-2020/Test_Case/case_344CCPA/5_GE_Mix/Error_log.txt', Path(F'{dest_folder}/mix/Error_logmix{date}.txt'))

    scp.close()

    # ssh.load_system_host_keys()
    # ssh.connect(hostname='192.168.220.170',
    #             username='auto-test',
    #             password='Default@0000')
    # scp = SCPClient(ssh.get_transport())
    # scp.get('D:/Autotest/8ge/sw-autoit-2020/Test_Case/case_344CCPA/5_GE_ccpb_8ge/Report_log.txt', Path(F'{dest_folder}/8ge/Report_log_8ge_{date}.txt'))
    # scp.get('D:/Autotest/8ge/sw-autoit-2020/Test_Case/case_344CCPA/5_GE_ccpb_8ge/Error_log.txt', Path(F'{dest_folder}/8ge/Error_log_8ge_{date}.txt'))


    # scp.close()


def get_logs_from_local(date,dest_folder):
    shutil.copyfile('D:/Jason-Huang/Repo/merge/sw-autoit-2020/Test_Case/case_344CCPA/5_GE_ccpa_G7800/Report_log.txt',Path(F'{dest_folder}/g7800/Report_log_g7800_{date}.txt' ))
    shutil.copyfile('D:/Jason-Huang/Repo/merge/sw-autoit-2020/Test_Case/case_344CCPA/5_GE_ccpa_G7800/Error_log.txt',Path(F'{dest_folder}/g7800/Error_log_g7800_{date}.txt' ))
    # shutil.copyfile('D:/Jason-Huang/Repo/sw-autoit-2020/Test_Case/case_344CCPA/5_GE_Mix/Report_log.txt',F'../error_logs/mix/Report_logmix{date}.txt' )
    # shutil.copyfile('D:/Jason-Huang/Repo/sw-autoit-2020/Test_Case/case_344CCPA/5_GE_Mix/Error_log.txt',F'../error_logs/mix/Error_logmix{date}.txt' )
    # shutil.copyfile('D:/Jason-Huang/Repo/sw-autoit-2020/AT_CCPA_CCPB/Test_Case/case_344CCPA/5_GE_ccpa/Report_log.txt',F'../error_logs/ccpa/Report_logccpa{date}.txt' )
    # shutil.copyfile('D:/Jason-Huang/Repo/sw-autoit-2020/AT_CCPA_CCPB/Test_Case/case_344CCPA/5_GE_ccpa/Error_log.txt',F'../error_logs/ccpa/Error_logccpa{date}.txt' )

def get_FW_test_json(dest_folder,product_name):
    ssh = paramiko.SSHClient()
    ssh.load_system_host_keys()
    ssh.connect(hostname='192.168.220.172',
                username='autotest',
                password='loop1234')


    latest = 0
    latestfile = None
    sftp = ssh.open_sftp()
    for fileattr in sftp.listdir_attr(F'D:/zack/sw-am3440-ssh/log/result/{product_name}'):
        if product_name == "am3440":
            if fileattr.filename.startswith('log-20') and fileattr.st_mtime > latest:
                latest = fileattr.st_mtime
                latestfile = fileattr.filename
        elif product_name == "g7800":
            if fileattr.filename.startswith('log-G7800') and fileattr.st_mtime > latest:
                latest = fileattr.st_mtime
                latestfile = fileattr.filename
    if latestfile is not None:
        print(F"Get FW download test log from D:/zack/sw-am3440-ssh/log/result/{product_name}/{latestfile}")
        sftp.get(F'D:/zack/sw-am3440-ssh/log/result/{product_name}/{latestfile}', Path(F'{dest_folder}/{latestfile}'))
    # list_fw_files = [f for f in glob.glob('K:/zack-chang_data/downloadRunCodeLog/*.json') if 'log-20' in f]
    # latest_file_file = max(list_fw_files,key=os.path.getctime)
    print(F'Put FW download test log to {dest_folder}/{latestfile}')
    # shutil.copyfile(latest_file_file,F'../error_logs/FWtest/{ntpath.basename(latest_file_file)}' )
    # return ntpath.basename(latest_file_file)
    return Path(F"{dest_folder}/{latestfile}")

def get_trap_test_json(dest_folder):
    list_trap_files = [f for f in glob.glob('K:/zack-chang_data/downloadRunCodeLog/*.json') if 'log-trap-' in f]
    latest_trap_file = max(list_trap_files,key=os.path.getctime)
    print(latest_trap_file)
    shutil.copyfile(latest_trap_file,Path(F'{dest_folder}/TrapTest/{ntpath.basename(latest_trap_file)}' ))
    return ntpath.basename(latest_trap_file)


def get_mib_diff_file(target_dir, mib_dir='K:/zack-chang_data/mibDiffLog/'):
    list_mib_dir = [os.path.join(mib_dir,d) for d in os.listdir(mib_dir)]
    latest_mib_dir = max(list_mib_dir,key=os.path.getctime)
    list_mib_log = [f for f in glob.glob(RF"{latest_mib_dir}/mib*.txt")]
    list_target_file = []
    for mib_log in list_mib_log:
        if not os.path.exists(F'{target_dir}/MIBdiff/{ntpath.basename(latest_mib_dir)}'):
            os.makedirs(F'{target_dir}/{ntpath.basename(latest_mib_dir)}')
        shutil.copyfile(mib_log,F'{target_dir}/MIBdiff/{ntpath.basename(latest_mib_dir)}/{ntpath.basename(mib_log)}' )
        list_target_file.append(F'{target_dir}/MIBdiff/{ntpath.basename(latest_mib_dir)}/{ntpath.basename(mib_log)}')
    if len(list_target_file) == 2:
        return list_target_file[0],list_target_file[1]
    else:
        return 'None','None'
    # return ntpath.basename(latest_trap_file)


if __name__ == "__main__":
    get_mib_diff_file()
    os._exit(0)
    get_FW_test_json()
    get_trap_test_json()