from datetime import datetime
from EXCELgen import write_data_to_excel
from HTMLgen import export_to_html
from LogParse import *
# import GetLogs
from GetLogs import *
from FileTarans import *
from MailSend import *
import re
import subprocess
from pathlib import Path
test_summary = {}
test_result = {}
result_cnt = {}
normal_parse = ["ccpb"] #"g7800","8ge","ccpa"
other_parse = ["cisco"]

if __name__ == "__main__":
    date = datetime.today().strftime('%m%d')
    log_path = Path('C:/Users/jeff-hsu/Desktop/sw-reportsummary/error_logs')
    summary_path = Path('C:/Users/jeff-hsu/Desktop/sw-reportsummary/Summary')
    # get_logs_from_remote(date,log_path)
    # get_logs_from_local(date,log_path)
    #log_ccpa = F"{log_path}/ccpa/Error_logccpa{date}.txt"
    log_ccpb = F"{log_path}/ccpb/Error_logccpb{date}.txt"
    #log_g7800 = F"{log_path}/g7800/Error_log_g7800_{date}.txt"
    #log_8ge = F"{log_path}/8ge/Error_log_8ge_{date}.txt"
    #log_mix = F"{log_path}/mix/Error_logmix{date}.txt"
    # log_cisco = F"{log_path}/cisco/cisco{date}/Error_log.txt"
    #summary_ccpa = F"{summary_path}/CCPA/summary_ccpa{date}.xlsx"
    summary_ccpb = F"{summary_path}/CCPB/summary_ccpb{date}.xlsx"
    #summary_g7800 = F"{summary_path}/G7800/summary_g7800_{date}.xlsx"
    #summary_mix = F"{summary_path}/CISCO/summary_cisco_{date}.xlsx"
    #summary_8ge = F"{summary_path}/8GE/summary_8ge_{date}.xlsx"
    # summary_cisco = F"{summary_path}/cisco/summary_cisco_{date}.xlsx"
    #report_log_ccpa = F"{log_path}/ccpa/Report_logccpa{date}.txt"
    report_log_ccpb = F"{log_path}/ccpb/Report_logccpb{date}.txt"
    #report_log_g7800 = F"{log_path}/g7800/Report_log_g7800_{date}.txt"
    #report_log_8ge = F"{log_path}/8ge/Report_log_8ge_{date}.txt"
    #report_log_mix = F"{log_path}/mix/Report_logmix{date}.txt"
    # report_log_cisco = F"{log_path}/cisco/cisco{date}/Report_log.txt"
    # file_fw_test = get_FW_test_json(F"{log_path}/FWtest/am3440/",'am3440')
    # file_trap_test = F"{log_path}/TrapTest/{get_trap_test_json(log_path)}"
    # mib_diff_log1, mib_diff_log2 = get_mib_diff_file(log_path)
    logs_to_parse = [log_ccpb]
    file_to_export = [summary_ccpb]
    report_log = [report_log_ccpb]
    #return_code = subprocess.run(["python","main_serial.py", "-d", F"D:\\Jason-Huang\\Repo\\Rp_gen\\error_logs\\g7800_func\\{date}","-p"])
    #print(return_code)
    for log, summary, report_log in zip(logs_to_parse,file_to_export, report_log):
        print(log,summary,report_log)
        path = re.findall(r'(ccpa|ccpb|g7800|cisco|8ge)',get_file_name(summary))[0]
        if path == 'ccpa' or path == 'g7800' or path == '8ge':
            continue
        print(F"Start parsing {path}")
        dut_info = get_dut_info(report_log)
        init_arrays(test_summary,test_result)
        if path in normal_parse:
            log_parse(log,test_summary,test_result)
            #test_file_parse("Firmware test",file_fw_test,test_result,test_summary)
            #if path != 'g7800':
                #mib_parse(F"{log_path}/MIBscan/mibOnly.txt",F"{log_path}/MIBscan/productOnly.txt",test_result,test_summary)
                # test_file_parse("Firmware test",file_fw_test,test_result,test_summary)
                #test_file_parse("Alarm Trap check",file_trap_test,test_result,test_summary)
                #mib_diff_file_parse([mib_diff_log1,mib_diff_log2],test_result,test_summary)
        elif path in other_parse:
            pdh_parse(log,test_result,test_summary)
        write_data_to_excel(dut_info,test_summary,test_result,summary)
        # export_to_html(summary,F'./HTML/{path}/{date}')
        #copy_file_dest(summary,F"K:/jason-huang_Data/AutoIt/Test summary/{path.upper()}/")
        #copy_file_dest(report_log,F"K:/jason-huang_Data/AutoIt/Test summary/{path.upper()}/")
    #SendMail()

