import re
import os
import json
from argparse import ArgumentParser
import sys
import re
test_summary = {
    # "warmreset":{
    #     "pass":0,
    #     "fail":0,
    #     "50ms test data":[],
    #     "50ms sort data":{},
    #     "fail cnt":[]
    # },
    # "coldreset":{
    #     "pass":0,
    #     "fail":0,
    #     "50ms test data":[],
    #     "50ms sort data":{},
    #     "fail cnt":[]
    # },
    # "portdisable":{
    #     "pass":0,
    #     "fail":0,
    #     "50ms test data":[],
    #     "50ms sort data":{},
    #     "fail cnt":[]
    # }
}



keywords = [
    "ptn_mode = 5 svlan",
    "ptn_mode = 4 cvlan"
]

action_list = [
    "warmreset",
    "coldreset",
    "portdisable"
]

def sort_time_for_50ms_test(array):
    group = {
        "< 10 ms":0,
        "10-50 ms":0,
        ">= 50ms" :0
    }
    array.sort()
    for time in array:
        if float(time) < 10 :
            group["< 10 ms"] +=1
        elif float(time) > 10 and float(time) < 50:
            # print(time)
            group["10-50 ms"] +=1
        elif float(time) >= 50:
            group[">= 50ms"] +=1
    # f_write.write(F"{json.dumps(group, indent=4)}")
    return group
    # print(group)


def is_float(num):
    try:
        float(num)
        return True
    except ValueError:
        return False

def parse_50ms(source_path,export_path):
    action_info = ""
    f_read = open(source_path,'r',encoding="utf-8")
    export_dir = os.path.dirname(export_path)
    if not os.path.exists(export_dir):
        os.makedirs(export_dir)
    if not os.path.exists(export_path):
        open(export_path, 'w').close()
    f_write = open(export_path, 'a')
    while True:
        text_line = f_read.readline()
        if text_line == '':
            break

        if any(search_text in text_line for search_text in keywords):
            action_info = re.split(r'\ ',text_line)[-1].split('_cnt=')
            # print(action_info)
            vlan_mode = re.findall(r'(?<=ptn_mode = [\d] ).[^,]*',text_line)[0]
            # print(vlan_mode)
            if vlan_mode not in test_summary.keys():
                test_summary[vlan_mode] = {}
                for action in action_list:
                    test_summary[vlan_mode][action] = {
                        "pass":0,
                        "fail":0,
                        "50ms test data":[],
                        "50ms sort data":{},
                        "fail cnt":[]
                    }
            # print(test_summary)
            if len(action_info) >= 2 :
                cnt = action_info[-1].split('\n')[0]
                action_type = action_info[-2]
                if action_type in test_summary[vlan_mode]:
                    text_line = f_read.readline()
                    # print(text_line)
                    # Bit error = 0 -> pass
                    # Bit error != 0, Curr value < 10 ms -> pass
                    # Else -> fail
                    if "Bit error" in text_line:
                        info = re.split(r'[:,\n]',text_line)
                        bit_error_val = re.split(r'Bit error=',info[1])[1]
                        curr_val = re.split(r' Curr value=',info[2])[1]
                        # if bit_error_val = 0 set curr_val to 0
                        if int(bit_error_val) == 0:
                            curr_val = 0
                        if int(bit_error_val) == 0 or int(curr_val) < 50:
                            text_line = f_read.readline()
                            if "50ms test result" in text_line:
                                time = re.split(r':\ ',text_line)[-1].split('ms')[0]
                                # print(time)
                                if is_float(time):
                                    test_summary[vlan_mode][action_type]["pass"]+=1
                                    if int(curr_val) == 0:
                                        test_summary[vlan_mode][action_type]["50ms test data"].append(float(0))
                                    else:
                                        test_summary[vlan_mode][action_type]["50ms test data"].append(float(time))
                        else:
                            text_line = f_read.readline()
                            if "50ms test result" in text_line:
                                time = re.split(r':\ ',text_line)[-1].split('ms')[0]
                                # print(time)
                                if is_float(time):
                                    test_summary[vlan_mode][action_type]["fail"]+=1
                                    test_summary[vlan_mode][action_type]["50ms test data"].append(float(time))
                                    test_summary[vlan_mode][action_type]["fail cnt"].append(int(cnt))


    print(json.dumps(test_summary,indent=4))
    for vlan_mode in test_summary.keys():
        test_summary[vlan_mode]["coldreset"]["50ms sort data"] = sort_time_for_50ms_test(test_summary[vlan_mode]["coldreset"]["50ms test data"])
        test_summary[vlan_mode]["warmreset"]["50ms sort data"] = sort_time_for_50ms_test(test_summary[vlan_mode]["warmreset"]["50ms test data"])
        test_summary[vlan_mode]["portdisable"]["50ms sort data"] = sort_time_for_50ms_test(test_summary[vlan_mode]["portdisable"]["50ms test data"])
        f_write.write(F"==========={vlan_mode}===========\n")
        f_write.write("===========ColdReset===========\n")
        f_write.write(F'Pass: {test_summary[vlan_mode]["coldreset"]["pass"]}\n')
        f_write.write(F'Fail: {test_summary[vlan_mode]["coldreset"]["fail"]}\n')
        f_write.write(F'50ms sort data:\n {json.dumps(test_summary[vlan_mode]["coldreset"]["50ms sort data"], indent=4)}\n')
        f_write.write(F'Fail cnt:\n {json.dumps(test_summary[vlan_mode]["coldreset"]["fail cnt"], indent=4)}\n')
        f_write.write("===========WarmReset===========\n")
        f_write.write(F'Pass: {test_summary[vlan_mode]["warmreset"]["pass"]}\n')
        f_write.write(F'Fail: {test_summary[vlan_mode]["warmreset"]["fail"]}\n')
        f_write.write(F'50ms sort data:\n {json.dumps(test_summary[vlan_mode]["warmreset"]["50ms sort data"], indent=4)}\n')
        f_write.write(F'Fail cnt:\n {json.dumps(test_summary[vlan_mode]["warmreset"]["fail cnt"], indent=4)}\n')
        f_write.write("===========PortDisable===========\n")
        f_write.write(F'Pass: {test_summary[vlan_mode]["portdisable"]["pass"]}\n')
        f_write.write(F'Fail: {test_summary[vlan_mode]["portdisable"]["fail"]}\n')
        f_write.write(F'50ms sort data:\n {json.dumps(test_summary[vlan_mode]["portdisable"]["50ms sort data"], indent=4)}\n')
        f_write.write(F'Fail cnt:\n {json.dumps(test_summary[vlan_mode]["portdisable"]["fail cnt"], indent=4)}\n')
        f_write.write("===============================\n\n")
    # f_write.write(json.dumps(test_summary["coldreset"], indent=4))
    # f_write.write(json.dumps(test_summary["warmreset"], indent=4))
    # f_write.write("===========Fail CNT===========\n")
    # for i in fail_cnt:
    #     f_write.write(F"{i}\n")
    # f_write.write("==============================\n")
    # print(fail_cnt)


    f_read.close
def main(*argv):
    parser = ArgumentParser()
    parser.add_argument("input_error_log", help="error log want to parse")
    argv = parser.parse_args()
    file_to_parse = argv.input_error_log
    date = re.findall(r'[\d]{4}.txt', argv.input_error_log)[0].replace(".txt","")
    file_to_export = F"./Report_50ms_{date}.txt"
    parse_50ms(file_to_parse,file_to_export)

    # print(json.dumps(signal_dict,indent=4))

if __name__ == "__main__":
    main(sys.argv[1:])