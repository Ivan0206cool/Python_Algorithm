from datetime import datetime
from EXCELgen import write_data_to_excel
from HTMLgen import export_to_html
from LogParse import *
# import GetLogs
from GetLogs import *
from FileTarans import *
from MailSend import *
import re
from distutils.dir_util import copy_tree

#import pyshark
import sys
import argparse
from pathlib import Path
#import shutil
#import json
import os



def check_fail_count(test_summary:dict):
    fail_count = 0
    for key in test_summary.keys():
        if 'fail' in test_summary[key]:
            fail_count += test_summary[key]['fail']
    return fail_count
def merge_dicts(dict1, dict2):
    for key in dict2:
        if key in dict1:
            if isinstance(dict1[key], dict) and isinstance(dict2[key], dict):
                merge_dicts(dict1[key], dict2[key])
            elif isinstance(dict1[key], list) and isinstance(dict2[key], list):
                dict1[key].extend(dict2[key])
            elif isinstance(dict1[key], int) and isinstance(dict2[key],int) :
                dict1[key]+= dict2[key]

        else:
            dict1[key] = dict2[key]
def merge_complex_dicts(dict_list):
    merged_dict = {}
    for d in dict_list:
        merge_dicts(merged_dict, d)
    return merged_dict

def main(argv):
    global DEBUG
    test_summary = {}
    test_result = {}
    total_fail_count = 0
    total_test_summary = []
    total_test_result = []
    cap_list = []
    pkt_list = {}
    #print(pkt.eth.src)
    #json_data = {}
    file_name = ""
    bd_type = 0
    id = 0
    root_directory = ""
    src_str = ""
    dst_str = ""
    #imported_file = []

    parser = argparse.ArgumentParser()
    parser.add_argument("-d", '--directory', default=".", help = "root directory of log files", action="store")
    parser.add_argument('-p', '--parse', help = "parse all log files from the root directory",action='store_true')
    parser.add_argument('-m', '--merge', help = "merge all json files from every test cases",action='store_true')
    parser.add_argument('-F', '--Firmware', help = "run frimware test case",action='store_true')
    parser.add_argument('-l', '--logFile', help = "log file",action='store')
    parser.add_argument('-P', '--Path', help = "excel file path and Name",action='store')
    #parser.add_argument('bd_type', help = "bundle type: eth or ip", type = str)
    #parser.add_argument('id', help = "ECID or UDP number", type = int)
    #parser.add_argument('src', help = "source mac/ip", type = str)
    #parser.add_argument('dst', help = "destination mac/ip", type = str)
    #parser.add_argument('--verbose', '-v', help = 'increase verbosity level', action = 'count', default = 0)
    #parser.add_argument('-l', '--log', help="generate log files", action = 'store_true')
    #parser.add_argument('-d', '--debug', help="generate files containing the dictionary data", action = 'store_true')
    args = parser.parse_args()
    #if args.verbose:
    #    print("verbosity is %d" % (args.verbose))
    #if args.log:
    #    print("log is turned on")
    #if args.debug:
    #    print("debug is turned on")
    #    DEBUG = args.debug
    #if args.imported_file and len(args.imported_file):
    #    print("imported mibs: %s" % (args.imported_file))

    root_directory = args.directory
    #bd_str = args.bd_type
    #id = args.id
    #src_str = args.src
    #dst_str = args.dst
    print(root_directory)
    #report_log_path = Path(root_directory).rglob('Report_log.txt')
    #print(report_log_path)

    file_fw_test = Path(root_directory)/get_FW_test_json(Path(root_directory),'g7800')
    # mib_diff_log_array = [Path(root_directory)/'MIBdiff'/'mib1.txt', Path(root_directory)/'MIBdiff'/'mib2.txt']
    
    if args.Firmware:
        test_file_parse("Firmware test", args.logFile ,test_result,test_summary)
        if args.Path:
            write_data_to_excel("", test_summary, test_result, excel_path=args.Path)
        else:
            write_data_to_excel("", test_summary, test_result)
        sys.exit()
    
    if args.parse:
        mibOnly_log = Path(root_directory)/'MIBcompare'/'mibOnly.txt'
        productOnly_log = Path(root_directory)/'MIBcompare'/'productOnly.txt'
        print(mibOnly_log)
        print(productOnly_log)
        for err_log, report_log in zip(Path(root_directory).rglob('Error_log.txt'), Path(root_directory).rglob('Report_log.txt')):
            print(err_log)
            print(report_log)
            log_dir_path = os.path.dirname(os.path.realpath(err_log))
            print(log_dir_path)
            current_folder_name = log_dir_path.split('\\')[-1]
            date = log_dir_path.split('\\')[-2]
            excel_path = Path(F"{log_dir_path}/summary_g7800_{current_folder_name}_{date}.xlsx")
            print(excel_path)
            #print(F"Start parsing {path}")
            dut_info = get_dut_info(report_log)
            init_arrays(test_summary,test_result)
            log_parse(err_log, test_summary, test_result)
            mib_parse(mibOnly_log,productOnly_log,test_result,test_summary)
            # mib_diff_file_parse(mib_diff_log_array,test_result,test_summary)
            test_file_parse("Firmware test",file_fw_test,test_result,test_summary)
            write_data_to_excel(dut_info, test_summary, test_result, excel_path)
            # export_to_html(excel_path,Path(F'{log_dir_path}/HTML'))
            total_fail_count += check_fail_count(test_summary)
            # print(total_fail_count)
        shutil.copytree(Path(root_directory),Path(F"K:\jason-huang_Data\AutoIt\Test summary\G7800\{ntpath.basename(Path(root_directory))}"))
        return total_fail_count

    if args.merge:
        for test_result_file, test_summary_file in zip(Path(root_directory).rglob('test_result.json'), Path(root_directory).rglob('test_summary.json')):
            # print(test_result_file)
            # print(test_summary_file)
            with open(test_result_file) as f:
                test_result = json.load(f)
            with open(test_summary_file) as f:
                test_summary = json.load(f)
            total_test_result.append(test_result.copy())
            total_test_summary.append(test_summary.copy())
        merged_test_result = {}
        merged_test_summary = {}
        merged_test_result = merge_complex_dicts(total_test_result)
        merged_test_summary = merge_complex_dicts(total_test_summary)
        # print(json.dumps(merged_test_summary))
        dut_info_file = next(Path(root_directory).rglob('dut_info.json'),None)
        # print(dut_info_file)
        with open(dut_info_file) as f:
            dut_info = json.load(f)
        if not os.path.exists(Path(F"{root_directory}/Merge")):
            os.makedirs(Path(F"{root_directory}/Merge"))
        write_data_to_excel(dut_info, merged_test_summary, merged_test_result,Path(F"{root_directory}/Merge/total_summary.xlsx"))
        export_to_html(Path(F"{root_directory}/Merge/total_summary.xlsx"),Path(F'{root_directory}/Merge/HTML'))
    
    #for filename in glob('src/**/*.c', recursive=True):
    #    print(filename) 
    

if __name__ == "__main__":
    main(sys.argv[1:])

"""
if __name__ == "__main__":
    date = datetime.today().strftime('%m%d')
    # get_logs_from_remote(date)
    # get_logs_from_local(date)
    log_ccpa = F"../error_logs/ccpa/Error_logccpa{date}.txt"
    log_ccpb = F"../error_logs/ccpb/Error_logccpb{date}.txt"
    log_g7800 = F"../error_logs/g7800/Error_log_g7800_{date}.txt"
    log_8ge = F"../error_logs/8ge/Error_log_8ge_{date}.txt"
    log_mix = F"../error_logs/mix/Error_logmix{date}.txt"
    summary_ccpa = F"./Summary/CCPA/summary_ccpa{date}.xlsx"
    summary_ccpb = F"./Summary/CCPB/summary_ccpb{date}.xlsx"
    summary_g7800 = F"./Summary/G7800/summary_g7800_{date}.xlsx"
    summary_mix = F"./Summary/CISCO/summary_cisco_{date}.xlsx"
    summary_8ge = F"./Summary/8GE/summary_8ge_{date}.xlsx"
    report_log_ccpa = F"../error_logs/ccpa/Report_logccpa{date}.txt"
    report_log_ccpb = F"../error_logs/ccpb/Report_logccpb{date}.txt"
    report_log_g7800 = F"../error_logs/g7800/Report_log_g7800_{date}.txt"
    report_log_8ge = F"../error_logs/8ge/Report_log_8ge_{date}.txt"
    report_log_mix = F"../error_logs/mix/Report_logmix{date}.txt"
    file_fw_test = F"../error_logs/FWtest/{get_FW_test_json()}"
    file_trap_test = F"../error_logs/TrapTest/{get_trap_test_json()}"
    mib_diff_log1 = F"../error_logs/MIBdiff/log1/mib1.txt"
    mib_diff_log2 = F"../error_logs/MIBdiff/log1/mib2.txt"
    logs_to_parse = [log_ccpa,log_ccpb,log_g7800,log_8ge,log_mix]
    file_to_export = [summary_ccpa,summary_ccpb,summary_g7800,summary_8ge,summary_mix]
    report_log = [report_log_ccpa,report_log_ccpb,report_log_g7800,report_log_8ge,report_log_mix]

    for log, summary, report_log in zip(logs_to_parse,file_to_export, report_log):
        path = re.findall(r'(ccpa|ccpb|g7800|cisco|8ge)',get_file_name(summary))[0]
        if path == 'g7800':
            continue
        print(F"Start parsing {path}")
        dut_info = get_dut_info(report_log)
        init_arrays(test_summary,test_result,result_cnt)
        if path != 'cisco':
            log_parse(log, test_summary, test_result, result_cnt)
            #if path != 'g7800':
                #mib_parse("../error_logs/MIBscan/mibOnly.txt","../error_logs/MIBscan/productOnly.txt",test_result,test_summary)
                #test_file_parse("Firmware test",file_fw_test,test_result,test_summary)
                #test_file_parse("Alarm Trap check",file_trap_test,test_result,test_summary)
            # if path == 'ccpb':
            #mib_diff_file_parse([mib_diff_log1,mib_diff_log2],test_result,test_summary)
        else:
            pdh_parse(log,test_result,test_summary)

        write_data_to_excel(dut_info, test_summary, test_result, result_cnt,summary)
"""
    #     copy_file_dest(summary,F"K:/jason-huang_Data/AutoIt/Test summary/{path.upper()}/")
    #     copy_file_dest(report_log,F"K:/jason-huang_Data/AutoIt/Test summary/{path.upper()}/")
    # SendMail()



    # create_main_page("./Summary/Main page/Main.html")
    # mib_parse("../error_logs/MIBscan/InMibOnly_2.txt","../error_logs/MIBscan/InProductOnly_2.txt",test_result)
    # pdh_parse(log_mix,test_result,test_summary)
    # write_data_to_excel(get_dut_info(report_log_mix),test_summary,test_result,result_cnt,summary_mix)
    # parse_mib(file_mib_scan)
    # parse_cfg(file_cfg_test)
    # write_data_to_excel(dut_info,test_summary,test_result,result_cnt,"./Summary/summary_ccpb0406.xlsx")

