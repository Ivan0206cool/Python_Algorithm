import re
import os
import json
from argparse import ArgumentParser
import sys
import re
from collections import namedtuple
import copy
import itertools
import numpy as np
import pandas as pd

CFGType = namedtuple("CFGTypeInfo", ["type", "code"])
CardType = namedtuple("CardTypeInfo", ["type", "code"])
cfg_type_code_map = {
    0: CFGType("Unframe", "00"),
    1: CFGType("Frame CRC CAS", "01"),
    2: CFGType("Frame CRC", "02"),
    3: CFGType("Frame", "03"),
    4: CFGType("Frame CAS", "04"),
    5: CFGType("Unframe192", "05"),
    6: CFGType("Unframe193", "06"),
    7: CFGType("D4 CAS", "07"),
    8: CFGType("ESF CAS", "08"),
    9: CFGType("D4", "00"),
    10: CFGType("ESF", "00")
}
card_type_code_map = {
    "QT1": CardType("QT1", "00"),
    "FT1": CardType("FT1", "01"),
    "MQT1": CardType("MQT1", "02"),
    "QE1": CardType("QE1", "10"),
    "FE1": CardType("FE1", "11"),
    "MQE1": CardType("MQE1", "12"),
    "4C37": CardType("4C37", "13"),
    "8CD": CardType("8CD", "14"),
    "FOM": CardType("FOM", "15"),
    "1C37": CardType("1C37", "16"),
    "Voice": CardType("Voice", "20")
}


def init_rs485_slot_info(dut_count):
    slot_list = ["A", "B", "C", "D"]
    slot_list.extend(range(1, 13))
    # slot_list = []
    # slot_list.extend(range(1,17))
    rs485_dut_all_slot_info = {}
    for dut_num in range(1, dut_count+1):
        rs485_dut_all_slot_info[F"DUT {dut_num}"] = {}
        for slot_num in slot_list:
            rs485_dut_all_slot_info[F"DUT {dut_num}"][F"slot {slot_num}"] = {
                "No RX count": 0,
                "Timeout count": 0,
                "Non-alive count": 0
            }
    # print(json.dumps(rs485_dut_all_slot_info,indent=4))
    return rs485_dut_all_slot_info


def init_arrays(test_summary, test_result):
    test_summary.clear()
    test_result.clear()
    # result_cnt.clear()

    # test_result["Bert"] = {"Total": 0, "Items": {}}
    # test_result["Ring"] = {"Total": 0, "Items": {}}
    # test_result["Tone"] = {"Total": 0, "Items": {}}
    # print(json.dumps(test_summary,indent=4))


def add_item_in_test_summary(test_summary, item, sheet_link_name):
    test_summary[item] = {
        "pass": 0,
        "fail": 0,
		"warning": 0,
        "sheet link": sheet_link_name
    }


def get_cfg_str_and_code(cfg_type):
    if cfg_type in cfg_type_code_map:
        return cfg_type_code_map[cfg_type].type, cfg_type_code_map[cfg_type].code
    else:
        return "no define", "99"


def get_card_type_str_and_code(card_type):
    if card_type in card_type_code_map:
        return card_type_code_map[card_type].type, card_type_code_map[card_type].code
    else:
        return "no define", "99"


def get_dut_info(input_report_path):
    f_read = open(input_report_path, 'r', encoding="utf-8")

    time_stamp, text_line = read_line_and_split_time(
        f_read)  # first line is 'begin'
    if 'begin' in text_line:
        time_stamp, text_line = read_line_and_split_time(f_read)
    dut_info = {}
    while "dut" in text_line.lower():
        if re.search(r'DUT[\d]{1} is', text_line):
            dut_num = text_line.split('is')[0].split('DUT')[1].strip()
            matches = re.findall(r'\b(LOOP|GE)([^\,]+)', text_line)
            for match in matches:
                dut_chassis = match[0] + ' ' + match[1].strip()
            print(dut_chassis)
            dut_slot_cnt = text_line.split(',slot#')[1].split(',')[
                0].split('\n')[0]
            dut_info[dut_num] = {
                "Chassis": dut_chassis,
                "Total slots": dut_slot_cnt,
                "Slot Info": []
            }
            # print(json.dumps(dut_info, indent=4))
        elif "dut:" in text_line.lower():
            dut_num = text_line.split(':')[1].split(' ')[0]
            dut_slot = text_line.split(':')[1].split(' ')[1]
            if "ctrl" in dut_slot:
                ctrl_hw_ver = text_line.split('"')[1].split(",")[0]
                ctrl_sw_ver = text_line.split('sw:"')[1].split('"')[0]
                ctrl_serial_ver = text_line.split('serial:')[1].split('\n')[0]
                if "ctrl" == dut_slot:  # ctrl1
                    ctrl_fpga1_ver = text_line.split('FPGA1:')[1].split(",")[0]
                    ctrl_fpga2_ver = text_line.split('FPGA2:')[1].split(",")[0]
                    ctrl_cpld_ver = text_line.split('CPLD:')[1].split(" ")[0]
                    dut_info[dut_num]["Slot Info"].append(
                        {
                            "Name": "Ctrl 1",
                            "HW Version": ctrl_hw_ver,
                            "SW Version": ctrl_sw_ver,
                            "Serial Version": ctrl_serial_ver,
                            "FPGA1 Version": ctrl_fpga1_ver,
                            "FPGA2 Version": ctrl_fpga2_ver,
                            "CPLD Version": ctrl_cpld_ver
                        })
                    # print(dut_cpld_ver,dut_slot)
                else:
                    dut_info[dut_num]["Slot Info"].append(
                        {
                            "Name": "Ctrl 2",
                            "HW Version": ctrl_hw_ver,
                            "SW Version": ctrl_sw_ver,
                            "Serial Version": ctrl_serial_ver
                        })
            elif "slot" in dut_slot:
                slot_info = re.findall(r'slot:(.*)\s+ver', text_line)
                # print(text_line,slot_info)
                slot_num = str(slot_info).split()[1].strip()
                slot_card_type = str(slot_info).split()[2].strip()
                # slot_card_type = re.split(r'\ ',text_line.split('slot: ')[1])[2]
                # print(slot_num, slot_card_type)
                slot_ver = text_line.split('ver:')[1].split('"')[1]
                # print(slot_ver)
                dut_info[dut_num]["Slot Info"].append(
                    {
                        "Name": F"Slot {slot_num}",
                        "Card Type": slot_card_type,
                        "Version": slot_ver
                    })
            #print(dut_slot)

        time_stamp, text_line = read_line_and_split_time(f_read)
    # print(list(dut_info.keys()))
    # for x , item in enumerate(list(dut_info.keys())):
    #     print(x,item,dut_info[item])
    f_read.close
    with open(F'{os.path.dirname(input_report_path)}/dut_info.json', 'w') as fw:
        fw.write(json.dumps(dut_info, indent=4))
    return dut_info
    # print(json.dumps(dut_info, indent=4))


def is_float(num):
    try:
        float(num)
        return True
    except ValueError:
        return False


def rs485_slot_info_parse(time_stamp, text_line, all_dut_slot_info):
    data = {}
    result = 'ok'
    # TXcount=538 RXcount=538 NoRXcount=0 TimeOutCount=0 AliveCount=538 NonAliveCount=0 OverSizeCount=0 InitCount=32 BadCMDCount=0 BadCMDNUM=0 SlotModel=qfxo
    dut = re.findall(r'\bDUT [^\s]*', text_line)[0]
    slot = re.findall(r'\bslot [^\s]*', text_line)[0]
    if "TXcount" in text_line:
        tx_count = re.findall(r'\bTXcount=[^\s]*', text_line)[0].split("=")[1]
    rx_count = re.findall(r'\bRXcount=[^\s]*', text_line)[0].split("=")[1]
    no_rx_count = re.findall(r'\bNoRXcount=[^\s]*', text_line)[0].split("=")[1]
    timeout_count = re.findall(
        r'\bTimeOutCount=[^\s]*', text_line)[0].split("=")[1]
    alive_count = re.findall(
        r'\bAliveCount=[^\s]*', text_line)[0].split("=")[1]
    non_alive_count = re.findall(
        r'\bNonAliveCount=[^\s]*', text_line)[0].split("=")[1]
    oversize_count = re.findall(
        r'\bOverSizeCount=[^\s]*', text_line)[0].split("=")[1]
    init_count = re.findall(r'\bInitCount=[^\s]*', text_line)[0].split("=")[1]
    bad_cmd_count = re.findall(
        r'\bBadCMDCount=[^\s]*', text_line)[0].split("=")[1]
    bad_cmd_num = re.findall(r'\bBadCMDNUM=[^\s]*', text_line)[0].split("=")[1]
    slot_model = re.findall(r'\bSlotModel=[^\s]*', text_line)[0].split("=")[1]
    slot_info = {
        F"{time_stamp}": F"{dut} {slot} {slot_model}"
    }

    if slot_model != 'empty':
        if "TXcount" in text_line:
            data["TX count"] = int(tx_count)
        data["RX count"] = int(rx_count)
        data["No RX count"] = int(no_rx_count)
        data["Timeout count"] = int(timeout_count)
        data["Alive count"] = int(alive_count)
        data["Non-alive count"] = int(non_alive_count)
        data["Oversize count"] = int(oversize_count)
        data["Init count"] = int(init_count)
        data["Bad CMD count"] = int(bad_cmd_count)
        data["Bad CMD num"] = int(bad_cmd_num)
        # if int(no_rx_count) > 0 and int(no_rx_count) > all_dut_slot_info[dut][slot]["No RX count"] :
        #     all_dut_slot_info[dut][slot]["No RX count"] = int(no_rx_count)
        #     result = 'fail'
        # if int(timeout_count) > 0 and int(timeout_count) > all_dut_slot_info[dut][slot]["Timeout count"]:
        #     all_dut_slot_info[dut][slot]["Timeout count"] = int(timeout_count)
        #     result = 'fail'
        # if int(non_alive_count) > 0 and int(non_alive_count) > all_dut_slot_info[dut][slot]["Non-alive count"]:
        #     all_dut_slot_info[dut][slot]["Non-alive count"] = int(non_alive_count)
        #     result = 'fail'
        if int(no_rx_count) > 0:
            result = 'warning'
        if int(timeout_count) > 0:
            result = 'warning'
        if int(non_alive_count) > 0:
            result = 'warning'
        data['result'] = result
    # if int(oversize_count) > 0:
    #     data["Oversize count"] = oversize_count
    # if int(bad_cmd_count) > 0:
    #     data["Bad CMD count"] = bad_cmd_count
    # if int(bad_cmd_num) > 0:
    #     data["Bad CMD num"] = bad_cmd_num
    if data:
        slot_info.update(data)
        data = slot_info
    return data


def rs485_dut_info_parse(time_stamp, text_line):
    data = {}
    result = 'fail'
    dut = re.findall(r'\bDUT [^\s]*', text_line)[0]

    break_time = re.findall(r'\bBreak=[^\s]*', text_line)[0].split("=")[1]
    framing = re.findall(r'\bFraming=[^\s]*', text_line)[0].split("=")[1]
    parity = re.findall(r'\bParity=[^\s]*', text_line)[0].split("=")[1]
    overrun = re.findall(r'\bOverrun=[^\s]*', text_line)[0].split("=")[1]
    FCS_error = re.findall(r'\bFCSerror=[^\s]*', text_line)[0].split("=")[1]
    slip_err = re.findall(r'\bSliperr=[^\s]*', text_line)[0].split("=")[1]
    buffer_error = re.findall(
        r'\bBuffererror=[^\s]*', text_line)[0].split("=")[1]
    cmd_Count_error = re.findall(
        r'\bCmdCounterror=[^\s]*', text_line)[0].split("=")[1]
    cmd_too_long = re.findall(
        r'\bCmdtooLong=[^\s]*', text_line)[0].split("=")[1]
    interrput = re.findall(r'\bInterrput=[^\s]*', text_line)[0].split("=")[1]
    record = re.findall(r'\bRecord=[^\s]*', text_line)[0].split("=")[1]
    dut_info = {
        F"{time_stamp}": F"{dut}"
    }
    data["FCS error"] = FCS_error
    data["Slip error"] = slip_err
    data["Buffer error"] = buffer_error
    # if int(FCS_error) > 0:
    #     data["FCS error"] = FCS_error
    # if int(slip_err) > 0:
    #     data["Slip error"] = slip_err
    # if int(buffer_error) > 0:
    #     data["Buffer error"] = buffer_error

    if int(FCS_error) > 0 or int(buffer_error):
       result = 'warning'
    elif int(FCS_error) > 0:
       result = 'fail'
    else:
        result = 'ok'
    data['result'] = result
    if data:
        dut_info.update(data)
        data = dut_info
    return data
def pw_statist_parse(time_stamp, text_line):
    # DUT 1 (AM3440_CCPA_Get_PW_Statist_log)  bid_idx=60 Eth1LinkStat=2 Eth2Linkstat=1 Eth1RxSeqNum=44777 Eth2RxSeqNum=0 Eth1RxSeqLos=0 Eth2RxSeqLos=0 Eth1RxJBLos=0 Eth2RxJBLos=0 DiffDelay="na."
    data = {}
    result = ''
    fail_reason = ''
    pri_sec = ''
    if 'Pri' in text_line:
        pri_sec = 'Primary'
    elif 'Sec' in text_line:
        pri_sec = 'Secondary'
    dut = re.findall(r'\bDUT [^\s]*', text_line)[0]
    bundle = int(re.findall(
        r'\bbid_idx=[^\s]*', text_line)[0].split("=")[1])

    pw_info = {
        F"{time_stamp}": F"{dut} Bundle {bundle} {pri_sec}"
    }
    parts = text_line.split()
    for part in parts:
        if '=' in part:
            key, value = part.split('=')
            data[key] = value
            if "RxJBLos" in key:
                if int(value) > 0:
                    result = 'fail'
                    fail_reason = key
    if result == 'fail':
        data['result'] = result
        data['fail_reason'] = fail_reason
    else:
        data['result'] = 'ok'
    pw_info.update(data)
    data = pw_info
    return data
    eth1_link_stat = re.findall(r'\bEth1LinkStat=[^\s]*', text_line)[0].split("=")[1]
    eth2_link_stat = re.findall(r'\bEth2Linkstat=[^\s]*', text_line)[0].split("=")[1]
    eth1_rx_seq_num = int(re.findall(r'\bEth1RxSeqNum=[^\s]*', text_line)[0].split("=")[1])
    eth2_rx_seq_num = int(re.findall(r'\bEth2RxSeqNum=[^\s]*', text_line)[0].split("=")[1])
    eth1_rx_seq_los = int(re.findall(r'\bEth1RxSeqLos=[^\s]*', text_line)[0].split("=")[1])
    eth2_rx_seq_los = int(re.findall(r'\bEth2RxSeqLos=[^\s]*', text_line)[0].split("=")[1])
    eth1_rx_jb_los = int(re.findall(r'\bEth1RxJBLos=[^\s]*', text_line)[0].split("=")[1])
    eth2_rx_jb_los = int(re.findall(r'\bEth2RxJBLos=[^\s]*', text_line)[0].split("=")[1])
    diff_delay = re.findall(r'\bDiffDelay=[^\s]*', text_line)[0].split("=")[1]
    pw_info = {
        F"{time_stamp}": F"{dut} Bundle {bundle}"
    }

    data['eth1_link_stat'] = eth1_link_stat
    data['eth2_link_stat'] = eth2_link_stat
    data['eth1_rx_seq_num'] = eth1_rx_seq_num
    data['eth2_rx_seq_num'] = eth2_rx_seq_num
    data['eth1_rx_seq_los'] = eth1_rx_seq_los
    data['eth2_rx_seq_los'] = eth2_rx_seq_los
    data['eth1_rx_jb_los'] = eth1_rx_jb_los
    data['eth2_rx_jb_los'] = eth2_rx_jb_los
    data['diff_delay'] = diff_delay

    if eth1_rx_jb_los > 0 :
        result = 'fail'
    else:
        result = 'ok'
    data['result'] = result
    if data:
        pw_info.update(data)
        data = pw_info
    return data

def pw_status_parse(time_stamp, text_line):
    data = {}
    result = 'fail'
    dut = re.findall(r'\bDUT [^\s]*', text_line)[0]
    # print(re.findall(r'\bCTRL-[^\s]*',text_line))
    if re.findall(r'\bCTRL-[^\s]*', text_line) != []:
        dut += " " + re.findall(r'\bCTRL-[^\s]*', text_line)[0]
    bid_index = int(re.findall(
        r'\bbid_idx=[^\s]*', text_line)[0].split("=")[1])
    bundle = re.findall(r'\bBundle=[^\s]*', text_line)[0].split("=")[1]
    bundle += " Primary " if "Pri" in text_line else " Secondary "
    jitter_buffer_under_run = float("{:.1f}".format(
        float(re.findall(r'\bJBUR=[^\s]*', text_line)[0].split("=")[1])))
    jitter_buffer_over_run = float("{:.1f}".format(
        float(re.findall(r'\bJBOR=[^\s]*', text_line)[0].split("=")[1])))
    jitter_buffer_min = float("{:.1f}".format(
        float(re.findall(r'\bJBMIN=[^\s]*', text_line)[0].split("=")[1])))
    jitter_buffer_max = float("{:.1f}".format(
        float(re.findall(r'\bJBMAX=[^\s]*', text_line)[0].split("=")[1])))
    jitter_buffer_depth = float("{:.1f}".format(
        float(re.findall(r'\bJBDP=[^\s]*', text_line)[0].split("=")[1])))
    pw_lbit_count = int(re.findall(
        r'\bLBitCnt=[^\s]*', text_line)[0].split("=")[1])
    pw_rbit_count = int(re.findall(
        r'\bRBitCnt=[^\s]*', text_line)[0].split("=")[1])
    tx_good_count = int(re.findall(
        r'\bTxGood=[^\s]*', text_line)[0].split("=")[1])
    rx_good_count = int(re.findall(
        r'\bRxGood=[^\s]*', text_line)[0].split("=")[1])
    rx_lost_count = int(re.findall(
        r'\bRxLost=[^\s]*', text_line)[0].split("=")[1])
    pw_link_status = re.findall(r'\bLink=[^\s]*', text_line)[0].split("=")[1]
    pw_size_error = re.findall(
        r'\bSizeErr=[^\s]*', text_line)[0].split("=")[1]

    pw_info = {
        F"{time_stamp}": F"{dut} Bundle {bundle}"
    }
    data["Jitter Buffer under run"] = jitter_buffer_under_run
    data["Jitter Buffer over run"] = jitter_buffer_over_run
    data["Jitter Buffer depth"] = F"{jitter_buffer_depth}({format(jitter_buffer_depth-8.0,'.1f')})"
    jitter_depth_differ = float(format(jitter_buffer_depth-8.0,'.1f'))
    data["Jitter Buffer min"] = jitter_buffer_min
    data["Jitter Buffer max"] = jitter_buffer_max
    data["PW L Bit_Count"] = pw_lbit_count
    data["PW R Bit_Count"] = pw_rbit_count
    data["PW TX Good Count"] = tx_good_count
    data["PW RX Good Count"] = rx_good_count
    data["RX lost count"] = rx_lost_count
    data["PW link status"] = pw_link_status
    # if jitter_buffer_under_run != 0.0:
    #     data["Jitter Buffer under run"] = jitter_buffer_under_run
    # if jitter_buffer_over_run != 0.0:
    #     data["Jitter Buffer over run"] = jitter_buffer_over_run
    # if jitter_buffer_depth > 12 or jitter_buffer_depth < 4:
    #     # print(jitter_buffer_depth)
    #     data["Jitter Buffer depth"] = F"{jitter_buffer_depth}({format(jitter_buffer_depth-8.0,'.1f')})"
    # if jitter_buffer_min < 4.0 or jitter_buffer_max > 20.0:
    #     # condition:jitter_buffer_min < (jiterr_delay/2)
    #     # condition:jitter_buffer_max > (jitter_size +2)
    #     data["Jitter Buffer min"] = jitter_buffer_min
    #     data["Jitter Buffer max"] = jitter_buffer_max

    # if int(pw_lbit_count) != 0:
    #     data["PW L Bit_Count"] = pw_lbit_count
    # if int(pw_rbit_count) != 0:
    #     data["PW R Bit_Count"] = pw_rbit_count
    # if int(rx_lost_count) > 0:
    #     data["RX lost count"] = rx_lost_count
    # if pw_link_status != "up":
    #     data["PW link status"] = pw_link_status

    if jitter_buffer_under_run != 0.0 or \
       jitter_buffer_over_run != 0.0 or \
       jitter_buffer_depth > 12.0 or jitter_buffer_depth < 4.0 or \
       jitter_buffer_min < 4.0 or jitter_buffer_max > 20.0 or \
       pw_lbit_count != 0 or \
       pw_rbit_count != 0 or \
       rx_lost_count > 0 or \
       pw_link_status != "up":
        result = 'fail'
        if pw_lbit_count != 0:
            fail_reason = 'L_bit'
        elif jitter_buffer_max > 20.0:
            fail_reason = 'jitter_max'
        elif jitter_buffer_min < 4.0:
            fail_reason = 'jitter_min'
        elif (-1 >= jitter_depth_differ or jitter_depth_differ >= 1) and round(jitter_depth_differ,2) == jitter_depth_differ and pw_link_status != 'down':
            fail_reason = 'DP+-1'
        elif (-2 >= jitter_depth_differ or jitter_depth_differ >= 2) and round(jitter_depth_differ,2) == jitter_depth_differ and pw_link_status != 'down':
            fail_reason = 'DP+-2'
        elif tx_good_count != 0 and abs(tx_good_count-rx_good_count) > 15:
            fail_reason = 'TX/RX good Gap > 15'
        elif pw_link_status == 'down':
            fail_reason = 'link down'
        else:
            fail_reason = 'others'
        data['fail_reason'] = fail_reason

    else:
        result = 'ok'
    data['result'] = result

    if data:
        pw_info.update(data)
        data = pw_info
    return data

def count_pw_test_summary(test_summary, working_pw_array,standby_pw_array):
    found_fail = False
    for working_status, standby_status in itertools.zip_longest(working_pw_array, standby_pw_array):
        if working_status is not None and working_status['result'] == 'fail':
            if working_status['fail_reason'] == 'L_bit':
                test_summary["PW Status(L-bit error)"]['fail'] += 1
            elif working_status['fail_reason'] == 'jitter_max':
                test_summary["PW Status(Jitter Max)"]['fail'] += 1
            elif working_status['fail_reason'] == 'DP+-1':
                test_summary["PW Status(DP +- 1)"]['fail'] += 1
            elif working_status['fail_reason'] == 'DP+-2':
                test_summary["PW Status(DP +- 2)"]['fail'] += 1
            elif working_status['fail_reason'] == 'link down':
                test_summary["PW Status(Link down)"]['fail'] += 1
            test_summary["PW Status"]['fail'] += 1
            found_fail = True
        if standby_status is not None and standby_status['result'] == 'fail':
            if standby_status['fail_reason'] == 'L_bit':
                test_summary["PW Status(L-bit error)"]['fail'] += 1
            elif standby_status['fail_reason'] == 'jitter_max':
                test_summary["PW Status(Jitter Max)"]['fail'] += 1
            elif standby_status['fail_reason'] == 'DP+-1':
                test_summary["PW Status(DP +- 1)"]['fail'] += 1
            elif standby_status['fail_reason'] == 'DP+-2':
                test_summary["PW Status(DP +- 2)"]['fail'] += 1
            elif standby_status['fail_reason'] == 'link down':
                test_summary["PW Status(Link down)"]['fail'] += 1
            test_summary["PW Status"]['fail'] += 1
            found_fail = True
    if found_fail != True:
        test_summary["PW Status"]['pass'] += 1

def count_bundle_pw_test_summary(test_summary, working_pw_array,standby_pw_array):
    found_fail = False
    for working_status, standby_status in itertools.zip_longest(working_pw_array, standby_pw_array):
        if working_status is not None and working_status['result'] == 'fail':
            # if working_status['fail_reason'] == 'L_bit':
            #     test_summary["Bundle test(PW status)(L-bit error)"]['fail'] += 1
            # elif working_status['fail_reason'] == 'jitter_max':
            #     test_summary["Bundle test(PW status)(Jitter Max)"]['fail'] += 1
            # elif working_status['fail_reason'] == 'DP+-1':
            #     test_summary["Bundle test(PW status)(DP +- 1)"]['fail'] += 1
            # elif working_status['fail_reason'] == 'DP+-2':
            #     test_summary["Bundle test(PW status)(DP +- 2)"]['fail'] += 1
            # elif working_status['fail_reason'] == 'link down':
            #     test_summary["Bundle test(PW status)(Link down)"]['fail'] += 1
            test_summary["Bundle test(PW status)"]['fail'] += 1
            found_fail = True
            break
        if standby_status is not None and standby_status['result'] == 'fail':
            # if standby_status['fail_reason'] == 'L_bit':
            #     test_summary["Bundle test(PW status)(L-bit error)"]['fail'] += 1
            # elif standby_status['fail_reason'] == 'jitter_max':
            #     test_summary["Bundle test(PW status)(Jitter Max)"]['fail'] += 1
            # elif standby_status['fail_reason'] == 'DP+-1':
            #     test_summary["Bundle test(PW status)(DP +- 1)"]['fail'] += 1
            # elif standby_status['fail_reason'] == 'DP+-2':
            #     test_summary["Bundle test(PW status)(DP +- 2)"]['fail'] += 1
            # elif standby_status['fail_reason'] == 'link down':
            #     test_summary["Bundle test(PW status)(Link down)"]['fail'] += 1
            test_summary["Bundle test(PW status)"]['fail'] += 1
            found_fail = True
            break
    if found_fail != True:
        test_summary["Bundle test(PW status)"]['pass'] += 1

def count_longtime_pw_test_summary(test_summary, working_pw_array,standby_pw_array):
    found_fail = False
    for working_status, standby_status in itertools.zip_longest(working_pw_array, standby_pw_array):
        if working_status is not None and working_status['result'] == 'fail':
            if working_status['fail_reason'] == 'L_bit':
                test_summary["Long Time(PW status)(L-bit error)"]['fail'] += 1
            elif working_status['fail_reason'] == 'jitter_max':
                test_summary["Long Time(PW status)(Jitter Max)"]['fail'] += 1
            elif working_status['fail_reason'] == 'DP+-1':
                test_summary["Long Time(PW status)(DP +- 1)"]['fail'] += 1
            elif working_status['fail_reason'] == 'DP+-2':
                test_summary["Long Time(PW status)(DP +- 2)"]['fail'] += 1
            elif working_status['fail_reason'] == 'link down':
                test_summary["Long Time(PW status)(Link down)"]['fail'] += 1
            test_summary["Long Time(PW status)"]['fail'] += 1
            found_fail = True
            break
        if standby_status is not None and standby_status['result'] == 'fail':
            if standby_status['fail_reason'] == 'L_bit':
                test_summary["Long Time(PW status)(L-bit error)"]['fail'] += 1
            elif standby_status['fail_reason'] == 'jitter_max':
                test_summary["Long Time(PW status)(Jitter Max)"]['fail'] += 1
            elif standby_status['fail_reason'] == 'DP+-1':
                test_summary["Long Time(PW status)(DP +- 1)"]['fail'] += 1
            elif standby_status['fail_reason'] == 'DP+-2':
                test_summary["Long Time(PW status)(DP +- 2)"]['fail'] += 1
            elif standby_status['fail_reason'] == 'link down':
                test_summary["Long Time(PW status)(Link down)"]['fail'] += 1
            test_summary["Long Time(PW status)"]['fail'] += 1
            found_fail = True
            break
    if found_fail != True:
        test_summary["Long Time(PW status)"]['pass'] += 1

def count_voice_pw_test_summary(test_summary, working_pw_array,standby_pw_array):
    found_fail = False
    for working_status, standby_status in itertools.zip_longest(working_pw_array, standby_pw_array):
        if working_status is not None and working_status['result'] == 'fail':
            if working_status['fail_reason'] == 'L_bit':
                test_summary["PW status(Voice test)(L-bit error)"]['fail'] += 1
            elif working_status['fail_reason'] == 'jitter_max':
                test_summary["PW status(Voice test)(Jitter Max)"]['fail'] += 1
            elif working_status['fail_reason'] == 'DP+-1':
                test_summary["PW status(Voice test)(DP +- 1)"]['fail'] += 1
            elif working_status['fail_reason'] == 'DP+-2':
                test_summary["PW status(Voice test)(DP +- 2)"]['fail'] += 1
            elif working_status['fail_reason'] == 'link down':
                test_summary["PW status(Voice test)(Link down)"]['fail'] += 1
            test_summary["PW status(Voice test)"]['fail'] += 1
            found_fail = True
        if standby_status is not None and standby_status['result'] == 'fail':
            if standby_status['fail_reason'] == 'L_bit':
                test_summary["PW status(Voice test)(L-bit error)"]['fail'] += 1
            elif standby_status['fail_reason'] == 'jitter_max':
                test_summary["PW status(Voice test)(Jitter Max)"]['fail'] += 1
            elif standby_status['fail_reason'] == 'DP+-1':
                test_summary["PW status(Voice test)(DP +- 1)"]['fail'] += 1
            elif standby_status['fail_reason'] == 'DP+-2':
                test_summary["PW status(Voice test)(DP +- 2)"]['fail'] += 1
            elif standby_status['fail_reason'] == 'link down':
                test_summary["PW status(Voice test)(Link down)"]['fail'] += 1
            test_summary["PW status(Voice test)"]['fail'] += 1
            found_fail = True
    if found_fail != True:
        test_summary["PW status(Voice test)"]['pass'] += 1

def dpll_status_parse(time_stamp, text_line):
    data = {}
    dut = re.findall(r'\bDUT \w+', text_line)[0]
    input_status = re.findall(r'\bInputStatus:\w+', text_line)[0].split(':')[1]
    lock_status = re.findall(r'\bLockStatus:\w+', text_line)[0].split(':')[1]
    if "CalStatus" in text_line:
    	cal_status = re.findall(r'\bCalStatus:\w+', text_line)[0].split(':')[1]
    input_sticky = re.findall(r'\bInputSticky:\w+', text_line)[0].split(':')[1]
    lock_sticky = re.findall(r'\bLockSticky:\w+', text_line)[0].split(':')[1]
    cal_sticky = re.findall(r'\bCalSticky:\w+', text_line)[0].split(':')[1]
    if "AM3440_CCPA_Get_dpllStabdbyStatus_log" in text_line:
        data[time_stamp] = dut + " standby"
    else:
        data[time_stamp] = dut + " working"
    data["Input status"] = input_status
    data["Lock status"] = lock_status
    data["Input sticky"] = input_sticky
    data["Lock sticky"] = lock_sticky
    if input_status != 'Normal' or \
       lock_status == 'Unlock' or \
       input_sticky != 'Normal' or \
       lock_sticky == 'Unlock':
        data['result'] = 'fail'
    else:
        data['result'] = 'ok'

    return data

def card_port_status_parse(time_stamp, text_line):
    data = {}
    result = ''
    fail_reason = ''
    dut = re.findall(r'\bDUT \w+', text_line)[0]
    card = text_line.split()[3]
    slot = re.findall(r'\bSlot \w+', text_line)[0]
    port = re.findall(r'\bPort \w+', text_line)[0]
    data[time_stamp] = F"{dut} {card} {slot} {port}"
    parts = text_line.split()
    for part in parts:
        if '=' in part:
            key, value = part.split('=')
            data[key] = value
            if "LOS" in key or "LOF" in key:
                if value == 'Yes':
                    result = 'fail'
                    fail_reason = key
    if result == 'fail':
        data['result'] = result
        #data['fail_reason'] = fail_reason
    else:
        data['result'] = 'ok'
    return data
    # print(data)


def pw_setting_parse(text_line):
    # (AM3440_CCPA_Get_PW_Setting) DUT 1 Pri BID=23 Format=Cesopsn Type= UdpEcid=0 NumofTimeslot="31." NumofFrame="1." JitterDelay=10 JitterSize=20 JitterDepth=10.0.
    data = {}
    data["DUT"] = re.findall(r'\bDUT \w+', text_line)[0]
    data["BID"] = re.findall(
            r'\bBID=[^ ]*', text_line)[0].split('=')[1]
    data["NumofFrame"] = re.findall(
            r'\bNumofFrame="[^ ]*', text_line)[0].split('=')[1].strip('"').rstrip('.').lstrip()
    data["JitterDelay"] = re.findall(
            r'\bJitterDelay=[^ ]*', text_line)[0].split('=')[1]
    data["JitterSize"] = re.findall(
            r'\bJitterSize=[^ ]*', text_line)[0].split('=')[1]
    if "JitterDepth" in text_line:
        data["JitterDepth"] = re.findall(
                r'\bJitterDepth=[^ ]*', text_line)[0].split('=')[1].strip('\n').strip('.')

    print(data)
    return data



def IP6704_aps_result_parse(time_stamp, text_line):
    data = {}
    # IP6704 APS Test Port_A_Result=Pass Port_B_Result=Pass Port_A_DelayTime="10875.0." Port_A_SwitchTime="7875.1." Port_B_DelayTime="10875.0." Port_B_SwitchTime="18884.1."

    data["TimeStamp"] = time_stamp
    data["Port A Result"] = re.findall(
            r'\bPortATestResult=[^ ]*', text_line)[0].split('=')[1]
    data["Port B Result"] = re.findall(
            r'\bPortBTestResult=[^ ]*', text_line)[0].split('=')[1]
    data["Port A Delaytime"] = re.findall(
            r'\bPortADelayTime=[^ ]*', text_line)[0].split('=')[1].rstrip('.')
    data["Port A SwitchTime"] = re.findall(
            r'\bPortASwitchTime=[^ ]*', text_line)[0].split('=')[1].rstrip('.')
    data["Port B Delaytime"] = re.findall(
            r'\bPortBDelayTime=[^ ]*', text_line)[0].split('=')[1].rstrip('.')
    data["Port B SwitchTime"] = re.findall(
            r'\bPortBSwitchTime=[^ ]*', text_line)[0].split('=')[1].strip('\n').rstrip('.')
    if data["Port A Result"] == 'Fail' or data["Port B Result"] == 'Fail':
        data["Result"] = "Fail"
    else:
        data["Result"] = "Pass"
    return data

def IP6704_aps_data_parse(time_stamp, text_line,f_read):
    data = {}
    data[time_stamp] = text_line

    time_stamp, cur_text_line = read_line_and_split_time(f_read)
    while "IP6704 APS" in cur_text_line:
        if cur_text_line != text_line:
            data[time_stamp] = cur_text_line
            text_line = cur_text_line
        time_stamp, cur_text_line = read_line_and_split_time(f_read)
    return data

def ctrl_bert_fail_data_parse(text_line, f_read):
    bit_error = ""
    unsync_sec = ""
    if "BitError" in text_line:
        bit_error = re.findall(
            r'\bBitError=[^ ]*', text_line)[0].split('=')[1]
    elif "UnsyncSecond" in text_line:
        unsync_sec = re.findall(
            r'\bUnsyncSecond=[^ ]*', text_line)[0].split('=')[1]
    error_id = re.findall(
        r'\berr_id=[^\n]*', text_line)[0].split('=')[1]
    # print(text_line)
    time_stamp, text_line = read_line_and_split_time(f_read)
    if "ElapsedTime" in text_line:
        elapsed_time = re.split(
            r'=', text_line)[-1].strip()
        if bit_error != "":
            data = {"Time": time_stamp,
                    "Bit Error": bit_error,
                    "Error ID": error_id,
                    "Elapsed Time": elapsed_time
                    }
        elif unsync_sec != "":
            data = {
                "Time": time_stamp,
                "UnsyncSecond": unsync_sec,
                "Error ID": error_id,
                "Elapsed Time": elapsed_time
            }
    # print(json.dumps(data,indent=4))
    return data


def qe1_bert_fail_data_parse(time_stamp, text_line):
    error_count = re.findall(
        r'ErrorCounts=[^ ]*', text_line)[0].split('=')[1]
    error_sec = re.findall(
        r'\bErrorSeconds=[^ ]*', text_line)[0].split('=')[1]
    if "err_id" in text_line:
        error_id = re.findall(
            r'\berr_id=[^\n]*', text_line)[0].split('=')[1]
    elapsed_sec = re.findall(
        r'\bElapseSeconds=[^ ]*', text_line)[0].split('=')[1]
    sync_status = re.findall(
        r'\bsPRBSStatus=[^ ]*', text_line)[0].split('=')[1]
    if 'pass' in text_line:
        result = 'pass'
    elif 'fail' in text_line:
        result = 'fail'
    data = {
        "Time": time_stamp,
        "PRBSStatus": sync_status,
		"ElapseSeconds":elapsed_sec,
        "ErrorCounts": error_count,
        #"Error ID": error_id,
        "ErrorSeconds": error_sec,
        "result":result
    }
    return data


def IP6704_50ms_fail_data_parse(time_stamp, text_line):
    # DUT 3	50ms pass :Bit error=46784, Curr value=18
    if "Bit error" in text_line:
        bit_error = re.findall(
            r'\bBit error=[^ ]*', text_line)[0].split('=')[1].strip(',')
    if "Curr value" in text_line:
        cur_val = re.findall(
            r'\bCurr value=[^ ]*', text_line)[0].split('=')[1].strip('\n')
    data = {
        "Time": time_stamp,
        "Bit error": bit_error,
        "Cur value": cur_val
    }
    return data

def IP6704_50ms_result_parse(time_stamp, text_line, f_read):
    result_str = time_stamp + " : " + "=======================================================\n"
    result_str += time_stamp + " : " + text_line
    text_line = f_read.readline()
    while 'End' not in text_line and text_line != '':
        result_str += text_line
        if "50ms result" in text_line:
            if 'Pass' in text_line:
                result = 'pass'
            else:
                result = 'fail'
        text_line = f_read.readline()
    result_str += text_line
    text_line = f_read.readline()
    result_str += text_line

    return result_str, result

def read_line_and_split_time(f_read):
    full_line = f_read.readline()
    # print(full_line)
    if not full_line:
        return full_line, ''
    else:
        text_line = re.split(
            r'\d{4}\-\d{2}-\d{2} \d{2}:\d{2}:\d{2} : ', full_line)[1]
        time_stamp = re.findall(
            r'\d{4}\-\d{2}-\d{2} \d{2}:\d{2}:\d{2} : ', full_line)[0].strip(" : ")
        return time_stamp, text_line


def log_parse(source_path, test_summary, test_result):
    dut_count = 2
    stages = ['Before setup check status',
              'Check setup status',
              'record before action',
              'record after action',
              'Reget action']
    case = ["warmreset",
            "coldreset(T1 Unframe)",
            "coldreset",
            "portdisable",
            "Ctrl1Eth1DisEn",
            "powereset",
            "systemreset",
            "ptn eth port disable",
            "Check setup status",
            "config_download"]
    for item in case:
        add_item_in_test_summary(test_summary, item, "Bert")
    add_item_in_test_summary(test_summary, "PW Status", "PW Status")
    add_item_in_test_summary(test_summary, "PW Status(DP +- 1)", "PW Status")
    add_item_in_test_summary(test_summary, "PW Status(DP +- 2)", "PW Status")
    add_item_in_test_summary(test_summary, "PW Status(L-bit error)", "PW Status")
    add_item_in_test_summary(test_summary, "PW Status(Jitter Max)", "PW Status")
    add_item_in_test_summary(test_summary, "PW Status(Link down)", "PW Status")
    add_item_in_test_summary(test_summary, "PW Statist", "PW Statist")
    add_item_in_test_summary(test_summary, "RS485 Status", "RS485 Status")
    add_item_in_test_summary(test_summary, "DPLL Status", "DPLL Status")
    add_item_in_test_summary(test_summary, "Card port status", "Card port status")
    add_item_in_test_summary(test_summary, "RS232 Status", "RS232 Status")
    add_item_in_test_summary(test_summary, "Long Time Bert", "Long Time Bert")
    
    rs485_all_dut_slot_info = init_rs485_slot_info(dut_count)

    with open(source_path, 'r', encoding="utf-8") as f_read:
        while True:
            round_num = 0
            working_pw_array = []
            standby_pw_array = []
            working_pw_statist_array = []
            standby_pw_statist_array = []
            cur_pw_setting_array = []
            rs485_array = []
            dpll_array = []
            card_port_status_array = []
            qe1_bert_array = []
            action_type = ""
            fail_record_before = False
            time_stamp, text_line = read_line_and_split_time(f_read)

            tone_cycle = 4
            ring_cycle = 12
            bert_cycle = 12

            # print(text_line)
            if not text_line:
                break
            if "TEST Begin" in text_line:
                #print(text_line)
                test_case = text_line.rsplit(' ', 2)[0]
                # f_write.write(F'{test_case}\n')
                if test_case == "tsi_in_out":
                    add_item_in_test_summary(
                        test_summary, "Basic TSI in out(Signalling)", "Basic TSI in out(Signalling)")
                    add_item_in_test_summary(
                        test_summary, "Basic TSI in out(Data)", "Basic TSI in out(Data)")
                    add_item_in_test_summary(
                        test_summary, "Cross TSI in out(Signalling)", "Cross TSI in out(Signalling)")
                    add_item_in_test_summary(
                        test_summary, "Cross TSI in out(Data)", "Cross TSI in out(Data)")
                    test_result["Basic TSI in out(Signalling)"] = {
                        "PW": {},
                        "Card": {}
                    }
                    test_result["Basic TSI in out(Data)"] = {
                        "PW": {},
                        "Card": {}
                    }
                    test_result["Cross TSI in out(Signalling)"] = {
                        "Card": {}
                    }
                    test_result["Cross TSI in out(Data)"] = {
                        "Card": {}
                    }

                    time_stamp, text_line = read_line_and_split_time(f_read)
                    while "TEST End" not in text_line and text_line != '':
                        if "Current test card" in text_line or "PW bundle setting" in text_line:
                            if "Current test card" in text_line:
                                card = re.findall(
                                    r'\btest card:\w+', text_line)[0].split(':')[1]
                                cfg = re.findall(
                                    r'\bcfg:\w+', text_line)[0].split(':')[1].strip()
                                cfg = card + " " + cfg
                                if cfg not in test_result["Basic TSI in out(Signalling)"]['Card']:
                                    test_result["Basic TSI in out(Signalling)"]['Card'][cfg] = {
                                    }
                                    test_result["Basic TSI in out(Data)"]['Card'][cfg] = {
                                    }
                                    test_result["Cross TSI in out(Signalling)"]['Card'][cfg] = {
                                    }
                                    test_result["Cross TSI in out(Data)"]['Card'][cfg] = {
                                    }
                            elif "PW bundle setting" in text_line:
                                cfg = text_line.split()[-1]
                                if cfg not in test_result["Basic TSI in out(Signalling)"]['PW']:
                                    test_result["Basic TSI in out(Signalling)"]['PW'][cfg] = {
                                    }
                                    test_result["Basic TSI in out(Data)"]['PW'][cfg] = {
                                    }
                        elif "Check card" in text_line or "Check cross connect" in text_line or \
                                "Check PW bundle" in text_line:
                            current_item = text_line
                        elif "Round" in text_line:
                            round_num = re.split(r'\-|:', text_line)[3]
                        elif "TSI CAS TS" in text_line or "TSI DATA TS" in text_line:
                            check_item = text_line.split()[1]
                            item_dict = {}
                            item_dict["Status"] = "Pass" if "--pass" in text_line else "Fail"
                            if check_item == 'CAS':
                                check_item = 'Signalling'
                                item_dict["OutCAS"] = "0x" + \
                                    re.findall(r'\boutCAS=\w+',
                                               text_line)[0].split('=')[1]
                                item_dict["InCAS"] = "0x" + \
                                    re.findall(r'\binCAS=\w+',
                                               text_line)[0].split('=')[1]
                            elif check_item == 'DATA':
                                check_item = "Data"
                                item_dict["OutData"] = "0x" + \
                                    re.findall(r'\boutData=\w+',
                                               text_line)[0].split('=')[1]
                                item_dict["InData"] = "0x" + \
                                    re.findall(r'\binData=\w+',
                                               text_line)[0].split('=')[1]
                            item_dict["TS"] = re.findall(
                                r'\bTS:\w+', text_line)[0].split(':')[1]
                            if 'PW bundle' in current_item:
                                if round_num not in test_result[F"Basic TSI in out({check_item})"]['PW'][cfg]:
                                    test_result[F"Basic TSI in out({check_item})"]['PW'][cfg][round_num] = {
                                        "Pass": 0,
                                        "Fail": 0,
                                        "Item": []
                                    }
                                if "--pass" in text_line:
                                    test_result[F"Basic TSI in out({check_item})"]['PW'][cfg][round_num]["Pass"] += 1
                                    test_summary[F"Basic TSI in out({check_item})"]["pass"] += 1
                                elif "--fail" in text_line:
                                    test_result[F"Basic TSI in out({check_item})"]['PW'][cfg][round_num]["Fail"] += 1
                                    test_summary[F"Basic TSI in out({check_item})"]["fail"] += 1
                                test_result[F"Basic TSI in out({check_item})"]['PW'][cfg][round_num]["Item"].append(
                                    item_dict.copy())
                            elif 'card' in current_item:
                                if round_num not in test_result[F"Basic TSI in out({check_item})"]['Card'][cfg]:
                                    test_result[F"Basic TSI in out({check_item})"]['Card'][cfg][round_num] = {
                                        "Pass": 0,
                                        "Fail": 0,
                                        "Item": []
                                    }
                                if "--pass" in text_line:
                                    test_result[F"Basic TSI in out({check_item})"]['Card'][cfg][round_num]["Pass"] += 1
                                    test_summary[F"Basic TSI in out({check_item})"]["pass"] += 1
                                elif "--fail" in text_line:
                                    test_result[F"Basic TSI in out({check_item})"]['Card'][cfg][round_num]["Fail"] += 1
                                    test_summary[F"Basic TSI in out({check_item})"]["fail"] += 1
                                test_result[F"Basic TSI in out({check_item})"]['Card'][cfg][round_num]["Item"].append(
                                    item_dict.copy())
                            elif 'cross connect' in current_item:
                                if round_num not in test_result[F"Cross TSI in out({check_item})"]['Card'][cfg]:
                                    test_result[F"Cross TSI in out({check_item})"]['Card'][cfg][round_num] = {
                                        "Pass": 0,
                                        "Fail": 0,
                                        "Item": []
                                    }
                                if "--pass" in text_line:
                                    test_result[F"Cross TSI in out({check_item})"]['Card'][cfg][round_num]["Pass"] += 1
                                    test_summary[F"Cross TSI in out({check_item})"]["pass"] += 1
                                elif "--fail" in text_line:
                                    test_result[F"Cross TSI in out({check_item})"]['Card'][cfg][round_num]["Fail"] += 1
                                    test_summary[F"Cross TSI in out({check_item})"]["fail"] += 1
                                test_result[F"Cross TSI in out({check_item})"]['Card'][cfg][round_num]["Item"].append(
                                    item_dict.copy())
                            # print(json.dumps(test_result[case],indent=4))
                        time_stamp, text_line = read_line_and_split_time(
                            f_read)
                    # print(json.dumps(test_result[case],indent=4))
                elif test_case == "E1/T1 au_law":
                    case = "TSI in out(AU_Law)"
                    add_item_in_test_summary(test_summary, case, case)
                    test_result[case] = {}
                    pass_count = 0
                    fail_count = 0
                    time_stamp, text_line = read_line_and_split_time(f_read)
                    while "TEST End" not in text_line:
                        if "au-law convert test" in text_line:
                            test_card = re.findall(
                                r'\[.*?\]', text_line)[0].strip('[]')
                            test_result[case][test_card] = {
                                "a to u(E1 to T1)": [],
                                "u to a(T1 to E1)": []
                            }
                        elif "a-law to u-law (E1 to T1) get" in text_line:
                            item_dict = {}
                            item_dict['status'] = re.findall(
                                r'(pass|fail)', text_line)[0]
                            item_dict['setData'] = "0x" + \
                                re.findall(
                                    r'\bset:\w+', text_line)[0].strip('set:')
                            item_dict['getData'] = "0x" + \
                                re.findall(
                                    r'\bget:\w+', text_line)[0].strip('get:')
                            test_result[case][test_card]["a to u(E1 to T1)"].append(
                                item_dict)
                            if item_dict['status'] == 'pass':
                                test_summary[case]["pass"] += 1
                            else:
                                test_summary[case]["fail"] += 1
                        elif "u-law to a-law (T1 to E1) get" in text_line:
                            item_dict = {}
                            item_dict['status'] = re.findall(
                                r'(pass|fail)', text_line)[0]
                            item_dict['setData'] = "0x" + \
                                re.findall(
                                    r'\bset:\w+', text_line)[0].strip('set:')
                            item_dict['getData'] = "0x" + \
                                re.findall(
                                    r'\bget:\w+', text_line)[0].strip('get:')
                            test_result[case][test_card]["u to a(T1 to E1)"].append(
                                item_dict)
                            if item_dict['status'] == 'pass':
                                test_summary[case]["pass"] += 1
                            else:
                                test_summary[case]["fail"] += 1
                        time_stamp, text_line = read_line_and_split_time(
                            f_read)
                elif test_case == "any_ds1":
                    stage = ""
                    case = "Any DS1"
                    add_item_in_test_summary(test_summary, case, case)
                    test_result[case] = []
                    time_stamp, text_line = read_line_and_split_time(f_read)
                    card = ""
                    while "TEST End" not in text_line:
                        text_line = re.sub(
                            r'\d{4}\-\d{2}-\d{2} \d{2}:\d{2}:\d{2} : ', "", text_line)
                        # text_line = text_line
                        result = {}
                        if "---any ds1 test begin Voice---with" in text_line:
                            card = text_line.split('---with ')[1].strip('\n')
                            result = {
                                "Voice datapath": "",
                                "Card": F"{card}",
                                "Pass count": 0,
                                "Fail count": 0,
                                "Init Pass": {
                                    "Tone": 0,
                                    "Bert": 0,
                                    "Ring": 0
                                },
                                "Init Fail": {
                                    "Tone": 0,
                                    "Bert": 0,
                                    "Ring": 0
                                },
                                "Card Disable Pass": {
                                    "Tone": 0,
                                    "Bert": 0,
                                    "Ring": 0
                                },
                                "Card Disable Fail": {
                                    "Tone": 0,
                                    "Bert": 0,
                                    "Ring": 0
                                },
                                "Eth Disable Pass": {
                                    "Tone": 0,
                                    "Bert": 0,
                                    "Ring": 0
                                },
                                "Eth Disable Fail": {
                                    "Tone": 0,
                                    "Bert": 0,
                                    "Ring": 0
                                }
                            }
                        time_stamp, text_line = read_line_and_split_time(
                            f_read)
                        while "---any ds1 test end Voice---" not in text_line:
                            # text_line = re.sub(r'\d{4}\-\d{2}-\d{2} \d{2}:\d{2}:\d{2} : ',"",text_line)
                            if "--------------Begin----" in text_line:
                                result["Voice datapath"] = text_line.split(
                                    " ")[0].split("--------------Begin----")[1]
                                stage = "Init"
                            elif "action" in text_line:
                                setup_pos = re.split(
                                    'set | port', text_line)[1]
                                stage = setup_pos
                                if "viacard" in setup_pos:
                                    stage = "Card Disable"
                                elif "eth" in setup_pos:
                                    stage = "Eth Disable"
                            else:
                                if "tone test" in text_line:
                                    if "pass" in text_line:
                                        result[F"{stage} Pass"]["Tone"] += 1
                                        result["Pass count"] += 1
                                    else:
                                        result[F"{stage} Fail"]["Tone"] += 1
                                        result["Fail count"] += 1
                                elif "ring test" in text_line:
                                    if "pass" in text_line:
                                        result[F"{stage} Pass"]["Ring"] += 1
                                        result["Pass count"] += 1
                                    else:
                                        result[F"{stage} Fail"]["Ring"] += 1
                                        result["Fail count"] += 1
                                elif "bert test" in text_line:
                                    if "Pass" in text_line:
                                        result[F"{stage} Pass"]["Bert"] += 1
                                        result["Pass count"] += 1
                                    else:
                                        result[F"{stage} Fail"]["Bert"] += 1
                                        result["Fail count"] += 1
                            time_stamp, text_line = read_line_and_split_time(
                                f_read)
                        test_result[case].append(result)

                        time_stamp, text_line = read_line_and_split_time(f_read)
                    for item in test_result[case]:
                        test_summary[case]["pass"] += item["Pass count"]
                        test_summary[case]["fail"] += item["Fail count"]

            elif "AM3440_CCPA_Rs485Delay_Clear_Snmp" in text_line:
                rs485_all_dut_slot_info = init_rs485_slot_info(dut_count)
            elif "Begin\n" in text_line:
                if "Bert" not in test_result:
                    test_result["Bert"] = {"Total": 0, "Items": {}}
                    #test_result["Ring"] = {"Total": 0, "Items": {}}
                    #test_result["Tone"] = {"Total": 0, "Items": {}}
                    test_result["50ms"] = {"Total": 0, "Items": {}}
                time_stamp, text_line = read_line_and_split_time(f_read)
                while "END\n" not in text_line and text_line != '':
                    if "-----Begin----" in text_line:  # get card type
                        if "test_card:" in text_line:
                            cur_pw_setting_array.clear()
                            # if there is a string means to capture the bert result(after record)
                            action_type = ""
                            cfg_type = re.split(r'\$cfg_type\ ', text_line)[
                                1].split(' ')[0]
                            str_cfg_type, str_cfg_type_code = get_cfg_str_and_code(
                                int(cfg_type))
                            card_type = re.split(
                                r'dut[\d+]_|_slot', text_line)[1]
                            str_card_type, str_card_type_code = get_card_type_str_and_code(
                                card_type)
                        elif "_voice_slot" in text_line:
                            action_type = ""
                            str_cfg_type, str_cfg_type_code = get_cfg_str_and_code(
                                99)
                            # str_card_type,str_card_type_code = get_card_type_str_and_code("Voice")
                            str_card_type = re.findall(
                                r'-----Begin----[^\ ]*', text_line)[0].split('/')[-1].upper()
                            # print(str_card_type)
                    elif "AM3440_CCPA_Get_PW_Setting" in text_line:
                        # (AM3440_CCPA_Get_PW_Setting) DUT 1 Pri BID=50 Format=Cesopsn Type= UdpEcid=0 NumofTimeslot="31." NumofFrame="1." JitterDelay=10 JitterSize=20
                        cur_pw_setting = pw_setting_parse(text_line)
                        cur_pw_setting_array.append(cur_pw_setting)
                        # print(text_line)
                    elif any((match := stage) in text_line for stage in stages):
                        working_pw_array.clear()
                        standby_pw_array.clear()
                        rs485_array.clear()
                        dpll_array.clear()
                        qe1_bert_array.clear()
                        time_stamp, text_line = read_line_and_split_time(
                            f_read)
                        #print(text_line)
                        while match +" end" not in text_line and text_line != '':
                            if match == 'Check setup status':
                                action_type = match
                                cur_aciton_info = str_card_type+'-'+str_cfg_type+'-'+action_type
                                if cur_aciton_info not in test_result["Bert"]["Items"] and str_card_type != "FXO":
                                    # print(text_line, str_card_type+'-'+str_cfg_type+'-'+action_type)
                                    test_result["Bert"]["Items"][cur_aciton_info] = {
                                        "pass": 0, "fail": 0, "fail_item": []}
                                    #test_result["50ms"]["Items"][cur_aciton_info] = {
                                        #"pass": 0, "fail": 0, "fail_item": []}
                                    #if str_card_type == "FXO" or str_card_type == "EM":
                                        # print(str_card_type+"---"+cur_aciton_info)
                                    #    test_result["Ring"]["Items"][cur_aciton_info] = {
                                    #        "pass": 0, "fail": 0, "fail_item": []}
                                    #    test_result["Tone"]["Items"][cur_aciton_info] = {
                                    #        "pass": 0, "fail": 0, "fail_item": []}
                            if "AM3440_CCPA_pw_status_get" in text_line or "AM3440_CCPA_pw_Standbystatus_get_log" in text_line:
                                case = "PW Status"
                                if case not in test_result:
                                    test_result[case] = []
                                pw_data = pw_status_parse(
                                    time_stamp, text_line)
                                if "AM3440_CCPA_pw_status_get" in text_line:
                                    if pw_data:
                                        working_pw_array.append(pw_data)
                                elif "AM3440_CCPA_pw_Standbystatus_get_log" in text_line:
                                    if pw_data:
                                        standby_pw_array.append(pw_data)
                            elif "AM3440_CCPA_Get_PW_Working_Statist_log" in text_line or "AM3440_CCPA_Get_PW_Standby_Statist_log" in text_line:
                                case = "PW Statist"
                                if case not in test_result:
                                    test_result[case] = []
                                pw_data = pw_statist_parse(time_stamp, text_line)
                                if "AM3440_CCPA_Get_PW_Working_Statist_log" in text_line:
                                    if pw_data:
                                        working_pw_statist_array.append(pw_data)
                                elif "AM3440_CCPA_Get_PW_Standby_Statist_log" in text_line:
                                    if pw_data:
                                        standby_pw_statist_array.append(pw_data)
                            elif "AM3440_CCPA_Get_Rs485_statist" in text_line:
                                case = "RS485 Status"
                                if case not in test_result:
                                    test_result[case] = []
                                rs485_data = rs485_dut_info_parse(
                                    time_stamp, text_line)
                                if rs485_data:
                                    rs485_array.append(rs485_data)
                            elif "AM3440_CCPA_rs485_get_statist_by_slot" in text_line:
                                case = "RS485 Status"
                                if case not in test_result:
                                    test_result[case] = []
                                rs485_data = rs485_slot_info_parse(
                                    time_stamp, text_line, rs485_all_dut_slot_info)
                                if rs485_data:
                                    rs485_array.append(rs485_data)
                            elif "AM3440_CCPA_Get_Card_Port_Status_log" in text_line:
                                case = "Card port status"
                                if case not in test_result:
                                    test_result[case] = []
                                card_port_data = card_port_status_parse(time_stamp, text_line)
                                card_port_status_array.append(card_port_data)
                                # print(card_port_status_array)
                            elif "AM3440_CCPA_Get_dpllstatus_log" in text_line or "AM3440_CCPA_Get_dpllStabdbyStatus_log" in text_line:
                                case = "DPLL Status"
                                if case not in test_result:
                                    test_result[case] = []
                                dpll_data = dpll_status_parse(
                                    time_stamp, text_line)
                                dpll_array.append(dpll_data)
                            elif "AM3440_CCPA_Rs485Delay_Clear_Snmp" in text_line:
                                rs485_all_dut_slot_info = init_rs485_slot_info(
                                    dut_count)
                            elif '(AM3440_CCPA_Get_QE1Bert_Snmp) begin' in text_line:
                                case = "Long Time Bert"
                                if case not in test_result:
                                    test_result[case] = []
                                time_stamp, text_line = read_line_and_split_time(f_read)
                                while '(AM3440_CCPA_Get_QE1Bert_Snmp) end' not in text_line and text_line != '':
                                    if "(AM3440_CCPA_Get_QE1Bert_Snmp) begin" not in text_line:
                                        bert_fail_data = qe1_bert_fail_data_parse(
                                            time_stamp, text_line)
                                        qe1_bert_array.append(bert_fail_data)
                                    time_stamp, text_line = read_line_and_split_time(
                                        f_read)
                            # Bert test parse
                            elif "(AM3440_CCPA_Get_FxsBert_Snmp) begin" in text_line and str_card_type != "FXO" and \
                                    (match == 'record after action' or match == 'record before action' or match == 'Check setup status'):
                                time_stamp, text_line = read_line_and_split_time(
                                    f_read)
                                while "(AM3440_CCPA_Get_FxsBert_Snmp) end" not in text_line:
                                    if "bert test Fail!!" in text_line:
                                        time_stamp, text_line = read_line_and_split_time(
                                            f_read)
                                        if "BitError" in text_line:
                                            if match == 'record after action' or match == 'Check setup status':
                                                bert_fail_data = {}
                                                bert_fail_data = ctrl_bert_fail_data_parse(
                                                    text_line, f_read)
                                                if fail_record_before:
                                                    bert_fail_data['Elapsed Time'] += "(Error before action)"
                                                test_result["Bert"]["Items"][cur_aciton_info]["fail_item"].append(
                                                    bert_fail_data)
                                                test_result["Bert"]["Total"] += 1
                                                test_result["Bert"]["Items"][cur_aciton_info]["fail"] += 1
                                                if "coldreset" in cur_aciton_info:
                                                    if "T1" in cur_aciton_info and "Unframe" in cur_aciton_info:
                                                        test_summary["coldreset(T1 Unframe)"]["fail"] += 1
                                                    else:
                                                        test_summary["coldreset"]["fail"] += 1
                                                else:
                                                    test_summary[action_type]["fail"] += 1
                                                fail_record_before = False
                                            elif match == 'record before action':
                                                fail_record_before = True
                                    elif "bert test Pass" in text_line:
                                        if match == 'record after action' or match == 'Check setup status':
                                            test_result["Bert"]["Total"] += 1
                                            # print(cur_aciton_info)
                                            test_result["Bert"]["Items"][cur_aciton_info]["pass"] += 1
                                            # print(json.dumps(test_result,indent=4))
                                            if "coldreset" in cur_aciton_info:
                                                if "T1" in cur_aciton_info and "Unframe" in cur_aciton_info:
                                                    test_summary["coldreset(T1 Unframe)"]["pass"] += 1
                                                else:
                                                    test_summary["coldreset"]["pass"] += 1
                                            else:
                                                test_summary[action_type]["pass"] += 1
                                    time_stamp, text_line = read_line_and_split_time(
                                        f_read)
                            elif ("BitError" in text_line or "UnsyncSecond" in text_line or
                                  "Bert pass" in text_line or re.search("Fireberd[\s\d]+result get", text_line)) and "IP6704 APS Test" not in text_line:
                                if ("BitError" in text_line or "UnsyncSecond" in text_line):
                                    if match == 'record after action' or match == 'Check setup status':
                                        bert_fail_data = {}
                                        bert_fail_data = ctrl_bert_fail_data_parse(
                                            text_line, f_read)
                                        if fail_record_before:
                                            bert_fail_data['Elapsed Time'] += "(Error before action)"
                                        test_result["Bert"]["Items"][cur_aciton_info]["fail_item"].append(
                                            bert_fail_data)
                                        test_result["Bert"]["Total"] += 1
                                        test_result["Bert"]["Items"][cur_aciton_info]["fail"] += 1
                                        if "coldreset" in cur_aciton_info:
                                            if "T1" in cur_aciton_info and "Unframe" in cur_aciton_info:
                                                test_summary["coldreset(T1 Unframe)"]["fail"] += 1
                                            else:
                                                test_summary["coldreset"]["fail"] += 1
                                        else:
                                            test_summary[action_type]["fail"] += 1
                                        fail_record_before = False
                                    elif match == 'record before action':
                                        fail_record_before = True
                                elif "Bert pass" in text_line or "bert test Pass" in text_line:
                                    if match == 'record after action' or match == 'Check setup status':
                                        test_result["Bert"]["Total"] += 1
                                        test_result["Bert"]["Items"][cur_aciton_info]["pass"] += 1
                                        if "coldreset" in cur_aciton_info:
                                            if "T1" in cur_aciton_info and "Unframe" in cur_aciton_info:
                                                test_summary["coldreset(T1 Unframe)"]["pass"] += 1
                                            else:
                                                test_summary["coldreset"]["pass"] += 1
                                        else:
                                            test_summary[action_type]["pass"] += 1
                                        if "switching time is " in text_line:
                                            recover_time = text_line.split()[-2] + ' sec'
                                            # print(recover_time)
                                            back_time_data = {time_stamp: "Switching time is " + recover_time}
                                            test_result["Bert"]["Items"][cur_aciton_info]["fail_item"].append(
                                                        back_time_data)
                                elif re.search("[\s]*Fireberd[\s\d]+result get", text_line):
                                    # print(action_type)
                                    time_stamp, text_line = read_line_and_split_time(
                                        f_read)
                                    fireberd_info = ""
                                    while "Fireberd check PASS." not in text_line and \
                                        "Fireberd get Error." not in text_line and \
                                            "Fireberd get SLIP." not in text_line and text_line != "":
                                        fireberd_info += re.sub(
                                            r'Fireberd RESULT[\s|\:]*', "", text_line).strip() + '\n'
                                        # fireberd_info += text_line.replace("Fireberd RESULT : ","").replace("Fireberd RESULT ","").replace("        ", " ").replace("     ","")
                                        time_stamp, text_line = read_line_and_split_time(
                                            f_read)
                                    # print(fireberd_info)

                                    if "Fireberd get Error." in text_line or \
                                            "Fireberd get SLIP." in text_line:
                                        test_result["Bert"]["Total"] += 1
                                        test_result["Bert"]["Items"][cur_aciton_info]["fail_item"].append(
                                            {"Fireberd error": time_stamp+'\n'+fireberd_info})
                                        test_result["Bert"]["Items"][cur_aciton_info]["fail"] += 1
                                        if "coldreset" in cur_aciton_info:
                                            if "T1" in cur_aciton_info and "Unframe" in cur_aciton_info:
                                                test_summary["coldreset(T1 Unframe)"]["fail"] += 1
                                            else:
                                                test_summary["coldreset"]["fail"] += 1
                                        else:
                                            test_summary[action_type]["fail"] += 1
                                    elif "Fireberd check PASS." in text_line:
                                        test_result["Bert"]["Total"] += 1
                                        test_result["Bert"]["Items"][cur_aciton_info]["pass"] += 1
                                        if "coldreset" in cur_aciton_info:
                                            if "T1" in cur_aciton_info and "Unframe" in cur_aciton_info:
                                                test_summary["coldreset(T1 Unframe)"]["pass"] += 1
                                            else:
                                                test_summary["coldreset"]["pass"] += 1
                                        else:
                                            test_summary[action_type]["pass"] += 1
                # --- RS232 stutas ---
                            elif "--------8RS232 ---------" in text_line and "TEST BEGIN" in text_line:
                                case = "RS232 Status"
                                if case not in test_result:
                                    test_result[case] = {"Total": 0, "Items": {}}
                                time_stamp, text_line = read_line_and_split_time(f_read)
                                if "Rate:" in text_line and "Async-" in text_line:
                                    randA = re.search(r'Rate: (\S+) Async-(\S+)', text_line)
                                    if randA:
                                        rate = randA.group(1)
                                        async1 = randA.group(2)
                                        rateandasync1 = f"{rate}_{async1}"
                                    if f"{rate}_{async1}" not in test_result[case]["Items"]:
                                        test_result[case]["Items"][f"{rate}_{async1}"] = {"pass": 0, "fail": 0}
                                    time_stamp, text_line = read_line_and_split_time(f_read)
                                    while "--------8RS232 --------- TEST END" not in text_line:
                                        time_stamp, text_line = read_line_and_split_time(f_read)
                                        if "Fireberd check" in text_line:
                                            if "PASS" in text_line:
                                                test_result[case]["Total"] += 1
                                                test_result[case]["Items"][f"{rate}_{async1}"]["pass"] += 1
                                                test_summary["RS232 Status"]["pass"] += 1
                                            else:
                                                test_result[case]["Total"] += 1
                                                test_result[case]["Items"][f"{rate}_{async1}"]["fail"] += 1
                                                test_summary["RS232 Status"]["fail"] += 1
                # --- RS232 stutas ---
                # --- tone-test --- 
                            elif "********tone-test*********" in text_line:
                                if "Tone" not in test_summary:
                                    add_item_in_test_summary(test_summary, "Tone", "Tone")
                                    test_result["Tone"] = {"Total": 0, "Items": {}}
                                if cur_aciton_info not in test_result["Tone"]["Items"]:
                                    test_result["Tone"]["Items"][cur_aciton_info] = {"u-law_900-ohm": {"pass": 0, "fail": 0, "fail_item": []}, 
                                                                                "u-law_600-ohm": {"pass": 0, "fail": 0, "fail_item": []},
                                                                                "a-law_900-ohm": {"pass": 0, "fail": 0, "fail_item": []},
                                                                                "a-law_600-ohm": {"pass": 0, "fail": 0, "fail_item": []}}
                                for aulaw in ["u-law","a-law"]:
                                    if aulaw in text_line:
                                        for ohm_type in ["900-ohm", "600-ohm"]:
                                            if ohm_type in text_line:
                                                time_stamp, text_line = read_line_and_split_time(f_read)
                                                while "********tone-test done*********" not in text_line:
                                                    if "tone pass" in text_line or "tone test pass" in text_line:
                                                        test_result["Tone"]["Total"] += 1
                                                        test_result["Tone"]["Items"][cur_aciton_info][f"{aulaw}_{ohm_type}"]["pass"] += 1
                                                        test_summary["Tone"]["pass"] += 1
                                                    elif "tone fail" in text_line:
                                                        test_result["Tone"]["Total"] += 1
                                                        test_result["Tone"]["Items"][cur_aciton_info][f"{aulaw}_{ohm_type}"]["fail"] += 1
                                                        test_result["Tone"]["Items"][cur_aciton_info][f"{aulaw}_{ohm_type}"]["fail_item"].append(text_line)
                                                        test_summary["Tone"]["fail"] += 1
                                                    time_stamp, text_line = read_line_and_split_time(f_read)
                # --- tone-test ---     
                # --- ring-test ---
                            elif "********ring-test*********" in text_line:
                                if "Ring" not in test_summary:
                                    add_item_in_test_summary(test_summary, "Ring", "Ring")
                                    test_result["Ring"] = {"Total": 0, "Items": {}}
                                if cur_aciton_info not in test_result["Ring"]["Items"]:
                                    test_result["Ring"]["Items"][cur_aciton_info] = {"u-law_900-ohm": {"pass": 0, "fail": 0, "fail_item": []}, 
                                                                                "u-law_600-ohm": {"pass": 0, "fail": 0, "fail_item": []},
                                                                                "a-law_900-ohm": {"pass": 0, "fail": 0, "fail_item": []},
                                                                                "a-law_600-ohm": {"pass": 0, "fail": 0, "fail_item": []}}
                                for aulaw in ["u-law","a-law"]:
                                    if aulaw in text_line:
                                        for ohm_type in ["900-ohm", "600-ohm"]:
                                            if ohm_type in text_line:
                                                time_stamp, text_line = read_line_and_split_time(f_read)
                                                while "********ring-test done*********" not in text_line:
                                                    if "TS" in text_line:
                                                        TS_text = text_line
                                                    #time_stamp, text_line = read_line_and_split_time(f_read)
                                                    if "ringing=yes,pass!" in text_line:
                                                        test_result["Ring"]["Total"] += 1
                                                        test_result["Ring"]["Items"][cur_aciton_info][f"{aulaw}_{ohm_type}"]["pass"] += 1
                                                        test_summary["Ring"]["pass"] += 1
                                                    elif "ringing=no" in text_line:
                                                        test_result["Ring"]["Total"] += 1
                                                        test_result["Ring"]["Items"][cur_aciton_info][f"{aulaw}_{ohm_type}"]["fail"] += 1
                                                        test_result["Ring"]["Items"][cur_aciton_info][f"{aulaw}_{ohm_type}"]["fail_item"].append(
                                                            time_stamp + ' ' + text_line.strip() + ' ' + TS_text)
                                                        test_summary["Ring"]["fail"] += 1
                                                    time_stamp, text_line = read_line_and_split_time(f_read)
                # --- ring-test ---
                # --- Voice BERT ---
                            elif "********bert*********" in text_line:
                                if "Voice Bert" not in test_summary:
                                    add_item_in_test_summary(test_summary, "Voice Bert", "Voice Bert")
                                    test_result["Voice Bert"] = {"Total": 0, "Items": {}}
                                if cur_aciton_info not in test_result["Voice Bert"]["Items"]:
                                    test_result["Voice Bert"]["Items"][cur_aciton_info] = {"u-law_900-ohm": {"pass": 0, "fail": 0, "fail_item": []}, 
                                                                                "u-law_600-ohm": {"pass": 0, "fail": 0, "fail_item": []},
                                                                                "a-law_900-ohm": {"pass": 0, "fail": 0, "fail_item": []},
                                                                                "a-law_600-ohm": {"pass": 0, "fail": 0, "fail_item": []}}
                                for aulaw in ["u-law","a-law"]:
                                    if aulaw in text_line:
                                        for ohm_type in ["900-ohm", "600-ohm"]:
                                            if ohm_type in text_line:
                                                time_stamp, text_line = read_line_and_split_time(f_read)
                                                while "********bert done*********" not in text_line:
                                                    if "TS" in text_line:
                                                        TS_text = text_line
                                                    #time_stamp, text_line = read_line_and_split_time(f_read)
                                                    if "Fail" in text_line:
                                                        test_result["Voice Bert"]["Total"] += 1
                                                        test_result["Voice Bert"]["Items"][cur_aciton_info][f"{aulaw}_{ohm_type}"]["fail"] += 1
                                                        test_result["Voice Bert"]["Items"][cur_aciton_info][f"{aulaw}_{ohm_type}"]["fail_item"].append(
                                                            time_stamp + ' ' + text_line.strip() + ' ' + TS_text)
                                                        test_summary["Voice Bert"]["fail"] += 1
                                                    elif "Pass" in text_line:
                                                        test_result["Voice Bert"]["Total"] += 1
                                                        test_result["Voice Bert"]["Items"][cur_aciton_info][f"{aulaw}_{ohm_type}"]["pass"] += 1
                                                        test_summary["Voice Bert"]["pass"] += 1
                                                    time_stamp, text_line = read_line_and_split_time(f_read)
                # --- Voice BERT ---
                            elif "50ms" in text_line or re.match(r'\s+Start\sDut\s\d+\sIP6704_DACS',text_line) != None or "IP6704 APS Test" in text_line :
                                # print(cur_aciton_info)
                                if "50ms" not in test_summary:
                                    add_item_in_test_summary(
                                        test_summary, "50ms", "50ms")
                                if "50ms pass" in text_line:
                                    test_result["50ms"]["Total"] += 1
                                    test_result["50ms"]["Items"][cur_aciton_info]["pass"] += 1
                                    test_summary["50ms"]["pass"] += 1
                                elif "50ms fail" in text_line:
                                    test_result["50ms"]["Total"] += 1
                                    test_result["50ms"]["Items"][cur_aciton_info]["fail"] += 1
                                    ip6704_bert_data = IP6704_50ms_fail_data_parse(
                                        time_stamp, text_line)
                                    test_result["50ms"]["Items"][cur_aciton_info]["fail_item"].append(
                                        ip6704_bert_data)
                                    test_summary["50ms"]["fail"] += 1
                                elif re.match(r'\s+Start\sDut\s\d+\sIP6704_DACS',text_line) != None:
                                    test_result["50ms"]["Total"] += 1
                                    ip6704_result_str, result = IP6704_50ms_result_parse(time_stamp,text_line,f_read)
                                    if result == 'pass':
                                        test_result["50ms"]["Items"][cur_aciton_info]["pass"] += 1
                                        test_summary["50ms"]["pass"] += 1
                                    else:
                                        test_result["50ms"]["Items"][cur_aciton_info]["fail"] += 1
                                        test_summary["50ms"]["fail"] += 1
                                    test_result["50ms"]["Items"][cur_aciton_info]["fail_item"].append(
                                        ip6704_result_str)
                                elif "IP6704 APS Test" in text_line:
                                    # print(text_line)
                                    if match == 'record after action':
                                        test_result["50ms"]["Total"] += 1
                                        back_time_data = IP6704_aps_result_parse(time_stamp, text_line)
                                        test_result["50ms"]["Items"][cur_aciton_info]["test_result"]
                                        test_result["50ms"]["Items"][cur_aciton_info]["test_result"].append(
                                                    back_time_data)
                                        if "Fail" in back_time_data["Result"]:
                                            test_result["50ms"]["Items"][cur_aciton_info]["fail"] += 1
                                            test_summary["50ms"]["fail"] += 1
                                        else:
                                            test_result["50ms"]["Items"][cur_aciton_info]["pass"] += 1
                                            test_summary["50ms"]["pass"] += 1
                                        # print(working_pw_array)
                                        test_result["50ms"]["Items"][cur_aciton_info]["pw_setting"]
                                        test_result["50ms"]["Items"][cur_aciton_info]["pw_setting"].append(
                                                    cur_pw_setting_array)
                            time_stamp, text_line = read_line_and_split_time(f_read)
                        # if len(working_pw_array) > 0 or len(rs485_array) > 0 or len(standby_pw_array) > 0 or len(dpll_array) > 0:
                        if match == 'Check setup status' or match == 'Before setup check status':
                            card_action_cfg_str = F"{match}-{str_card_type}-{str_cfg_type}"
                        else:
                            card_action_cfg_str = F"{match}-{cur_aciton_info}"
                        if len(working_pw_array) > 0 or len(standby_pw_array) > 0:
                            test_result["PW Status"].append({
                                card_action_cfg_str: {
                                    "Working PW status": [],
                                    "Standby PW status": []
                                }
                            })

                            count_pw_test_summary(test_summary, working_pw_array,standby_pw_array)
                            if len(working_pw_array) > 0:
                                test_result["PW Status"][-1][card_action_cfg_str]["Working PW status"] = working_pw_array.copy()
                            if len(standby_pw_array) > 0:
                                test_result["PW Status"][-1][card_action_cfg_str]["Standby PW status"] = standby_pw_array.copy()
                        if len(working_pw_statist_array) > 0 or len(standby_pw_statist_array) > 0:
                            test_result["PW Statist"].append({
                                card_action_cfg_str: {
                                    "Working PW Static": [],
                                    "Standby PW Static": []
                                }
                            })
                            if len(working_pw_statist_array) > 0:
                                test_result["PW Statist"][-1][card_action_cfg_str]["Working PW Static"] = working_pw_statist_array.copy()
                            if len(standby_pw_statist_array) > 0:
                                test_result["PW Statist"][-1][card_action_cfg_str]["Standby PW Static"] = standby_pw_statist_array.copy()
                            found = False
                            for pw_statist in working_pw_statist_array:
                                if pw_statist['result'] == 'fail':
                                    test_summary["PW Statist"]['fail'] += 1
                                    found = True
                                    break
                            for pw_statist in standby_pw_statist_array:
                                if pw_statist['result'] == 'fail':
                                    test_summary["PW Statist"]['fail'] += 1
                                    found = True
                                    break
                            if found != True:
                                test_summary["PW Statist"]['pass'] += 1
                        #     test_summary["PW Status"]['fail'] += 1
                        # else:
                        #     test_summary["PW Status"]['pass'] += 1
                        if len(rs485_array) > 0:
                            test_result["RS485 Status"].append({
                                card_action_cfg_str: rs485_array.copy()
                            })
                            found = False
                            for rs485_status in rs485_array:
                                if rs485_status['result'] == 'fail':
                                    test_summary["RS485 Status"]['fail'] += 1
                                    found = True
                                    break
                                elif rs485_status['result'] == 'warning':
                                    test_summary["RS485 Status"]['warning'] += 1
                                else:
                                    test_summary["RS485 Status"]['pass'] += 1
                            # print(json.dumps(test_result["RS485 Status"],indent=4))
                        # else:
                        #     test_summary["RS485 Status"]['pass'] += 1
                        if len(dpll_array) > 0:
                            test_result["DPLL Status"].append({
                                card_action_cfg_str: dpll_array.copy()
                            })
                            found = False
                            for dpll_status in dpll_array:
                                if dpll_status['result'] == 'fail':
                                    test_summary['DPLL Status']['fail'] += 1
                                    found = True
                                    break
                            if found != True:
                                test_summary["DPLL Status"]['pass'] += 1
                        if len(card_port_status_array) > 0:
                            test_result['Card port status'].append({card_action_cfg_str: card_port_status_array.copy()})
                            found = False
                            for card_port_status in card_port_status_array:
                                # print(card_port_status)
                                if card_port_status['result'] == 'fail':
                                    test_summary["Card port status"]['fail'] += 1
                                    found = True
                                    break
                            if found != True:
                                test_summary["Card port status"]['pass'] += 1
                        if len(qe1_bert_array) > 0:
                            test_result['Long Time Bert'].append({card_action_cfg_str: qe1_bert_array.copy()})
                            found = False
                            for qe1_bert_status in qe1_bert_array:
                                if qe1_bert_status['result'] == 'fail':
                                    test_summary['Long Time Bert']['fail'] += 1
                                    found = True
                                    break
                            if found != True:
                                test_summary["Long Time Bert"]['pass'] += 1
                        working_pw_array.clear()
                        standby_pw_array.clear()
                        working_pw_statist_array.clear()
                        standby_pw_statist_array.clear()
                        rs485_array.clear()
                        dpll_array.clear()
                        card_port_status_array.clear()
                        qe1_bert_array.clear()
                        # print(cur_aciton_info)
                    elif ("CLK Init" in text_line or "_cnt=" in text_line or
                          "system reset" in text_line or "power off/on" in text_line or
                          "ptn eth port disable" in text_line):
                        if "_cnt=" in text_line:
                            action_info = re.split(
                                r'\ ', text_line)[-1].split('_cnt=')
                            cnt = action_info[-1].split('\n')[0]
                            # remove $ when the action isCtrl1Eth1DisEn
                            action_type = action_info[-2].replace('$', '')
                            if "reboot active at CC1" in text_line:
                                action_type += "(C1 to C2)"
                            elif "reboot active at CC2" in text_line:
                                action_type += "(C2 to C1)"
                             # print(action_type,cnt)
                        elif "CLK Init" in text_line:
                            action_info = re.findall(
                                r'CLK Init [^\ ^/]*', text_line)[0]
                            action_type = action_info.replace(
                                'CLK Init ', 'SetClkSrc')
                            if action_type not in test_summary:
                                add_item_in_test_summary(
                                    test_summary, action_type, "Bert")
                        elif "system reset" in text_line:
                            action_type = "systemreset"
                        elif "power off/on" in text_line:
                            action_type = "powereset"
                        elif "ptn eth port disable" in text_line:
                            action_type = "ptn eth port disable"
                        cur_aciton_info = str_card_type+'-'+str_cfg_type+'-'+action_type
                        if cur_aciton_info not in test_result["Bert"]["Items"] and str_card_type != "FXO":
                            # print(text_line, str_card_type+'-'+str_cfg_type+'-'+action_type)
                            test_result["Bert"]["Items"][cur_aciton_info] = {
                                "pass": 0, "fail": 0, "fail_item": []}
                            test_result["50ms"]["Items"][cur_aciton_info] = {
                                        "pass": 0, "fail": 0, "fail_item": [], "test_result":[], "pw_setting":[]}
                            #if str_card_type == "FXO" or str_card_type == "EM":
                                # print(str_card_type+"---"+cur_aciton_info)
                                #test_result["Ring"]["Items"][cur_aciton_info] = {
                                    #"pass": 0, "fail": 0, "fail_item": []}
                                #test_result["Tone"]["Items"][cur_aciton_info] = {
                                    #"pass": 0, "fail": 0, "fail_item": []}
                    elif re.match(r'\s+Start\sDut\s\d+\sIP6704_DACS',text_line) != None:
                        test_result["50ms"]["Total"] += 1
                        ip6704_result_str, result = IP6704_50ms_result_parse(time_stamp,text_line,f_read)
                        if result == 'pass':
                            test_result["50ms"]["Items"][cur_aciton_info]["pass"] += 1
                            test_summary["50ms"]["pass"] += 1
                        else:
                            test_result["50ms"]["Items"][cur_aciton_info]["fail"] += 1
                            test_summary["50ms"]["fail"] += 1
                        test_result["50ms"]["Items"][cur_aciton_info]["fail_item"].append(
                            ip6704_result_str)

                    time_stamp, text_line = read_line_and_split_time(f_read)
    f_read.close

def mib_parse(in_mib_only_path, in_dut_only_path, test_result, test_summary):
    case = "MIB Scan"
    add_item_in_test_summary(test_summary, case, case)
    test_result[case] = {
        "MIB version": "",
        "Registered cards": "",
        "In MIB only result": [],
        "In DUT only result": []
    }

    with open(in_mib_only_path, 'r', encoding="utf-8") as f_read:
        while True:
            text_line = f_read.readline()
            if not text_line:
                break
            if '[in mib only]' in text_line:
                text_line = f_read.readline()
                while text_line != '':
                    if text_line.strip() != '':
                        test_result[case]["In MIB only result"].append(
                            text_line.strip())
                    text_line = f_read.readline()

    with open(in_dut_only_path, 'r', encoding="utf-8") as f_read:
        while True:
            text_line = f_read.readline()
            if not text_line:
                break
            if '[title]' in text_line:
                text_line = f_read.readline()
                if 'mibfilename' in text_line:
                    test_result[case]["MIB version"] = text_line.split(' = ')[
                        1].strip()
                text_line = f_read.readline()
                if 'firmware version' in text_line:
                    test_result[case]["FW version"] = text_line.split(' = ')[
                        1].strip('"')
                text_line = f_read.readline()
                if 'registerSlot' in text_line:
                    test_result[case]["Registered cards"] = text_line.split(' = ')[
                        1].strip(',')
            elif '[in product only]' in text_line:
                text_line = f_read.readline()
                while text_line != '':
                    if text_line.strip() != '':
                        test_result[case]["In DUT only result"].append(
                            text_line.split(' = ')[0])
                    text_line = f_read.readline()


def mib_diff_file_parse(file_path_array: list, test_result: dict, test_summary):
    test_case = "MIB Diff"
    add_item_in_test_summary(test_summary, test_case, test_case)
    if test_case not in test_result:
        test_result[test_case] = []
    for file_path in file_path_array:
        # print(file_path)
        with open(file_path, 'r', encoding='utf-8') as f_read:
            mib_info = {}
            while True:
                text_line = f_read.readline()
                if not text_line:
                    break
                if 'MIB' in text_line:
                    mib_info = json.loads(text_line)
                    mib_info['diff_oid'] = []
                else:
                    mib_info['diff_oid'].append(text_line)
        test_result[test_case].append(mib_info.copy())
    # print(json.dumps(test_result[test_case],indent=4))


def test_file_parse(test_case, file_path, test_result, test_summary):
    # print(test_case)
    # case = "Firmware test"
    add_item_in_test_summary(test_summary, test_case, test_case)
    f_read = open(file_path)
    test_result[test_case] = json.load(f_read)
    # print(json.dumps(test_result[case],indent=4))

    return
    IPmode = ""
    Connectmode = ""
    test_result[case] = {"IPV4": {}, "IPV6": {}}
    connection = ["VT100(COM)",
                  "SNMPV1",
                  "SNMPV3",
                  "Telnet",
                  "SSH",
                  "WEB"]
    main_items = ["ctrl firmware download",
                  "cfg upload/download",
                  "Download flash",
                  "ctrl fae upload"]
    unit_items = [
        "unit fpga download",
        "unit cfg/fae updownload",
        "unit firmware download"]
    card_unit = ["QE1", "QT1", "FXS", "FXO", "RTB"]
    card_unit_result = {key: {"Pass": 0, "Fail": 0} for key in card_unit}

    test_result[case]["IPV4"] = {key: {} for key in connection}
    test_result[case]["IPV6"] = {key: {} for key in connection}
    for connect in connection:
        test_result[case]["IPV4"][connect] = {
            key: {"Pass": 0, "Fail": 0} for key in main_items}
        test_result[case]["IPV6"][connect] = {
            key: {"Pass": 0, "Fail": 0} for key in main_items}
        test_result[case]["IPV4"][connect].update(
            {key: copy.deepcopy(card_unit_result) for key in unit_items})
        test_result[case]["IPV6"][connect].update(
            {key: copy.deepcopy(card_unit_result) for key in unit_items})
    # print(json.dumps(test_result,indent=4))

    try:
        f_read = open(file_path, 'r', encoding="utf-8")
    except Exception as e:
        print("error! {}".format(e))
    while True:
        text_line = f_read.readline()
        if text_line == '':
            break
        if "Current test mode" in text_line:
            IPmode = text_line.split()[-2]
            Connectmode = text_line.split()[-1]
            print(IPmode, Connectmode)

        elif "Config Test Start" in text_line:
            text_line = f_read.readline()
            while "Config Test End" not in text_line:
                if "Pass" in text_line:
                    test_result[case][IPmode][Connectmode]["cfg upload/download"]["Pass"] += 1
                elif "Fail" in text_line:
                    test_result[case][IPmode][Connectmode]["cfg upload/download"]["Fail"] += 1
                text_line = f_read.readline()
        elif "start download" in text_line:
            cur_card = re.findall(r'\bsnmp [^\s]*', text_line)[0].split()[1]
            text_line = f_read.readline()
            while "end download" not in text_line and text_line != '':
                if cur_card == "Controller":
                    if "pass" in text_line:
                        test_result[case][IPmode][Connectmode]["ctrl firmware download"]["Pass"] += 1
                    elif "fail" in text_line:
                        test_result[case][IPmode][Connectmode]["ctrl firmware download"]["Fail"] += 1
                else:
                    if "pass download" in text_line:
                        test_result[case][IPmode][Connectmode]["unit firmware download"][cur_card]["Pass"] += 1
                    elif "fail download" in text_line:
                        test_result[case][IPmode][Connectmode]["unit firmware download"][cur_card]["Fail"] += 1
                text_line = f_read.readline()
    # with open('./test.json','w') as f_write:
    #     json.dump(test_result[case],f_write)
    # print(json.dumps(test_result[case],indent=4))


def pdh_parse(file_path, test_result, test_summary):
    working_pw_array = []
    standby_pw_array = []
    rs485_array = []
    dut_count = 4
    pdh_stage = [
        'Check init bert status',
        'Set QE1 port disable',
        'Get bert status after QE1 port disable',
        'Set QE1 port enable',
        'Get bert status after QE1 port enable'
    ]
    voice_stages = ['Before setup check status',
            'Check setup status',
            'record before action',
            'record after action',
            'Reget action',
            ]
    cisco_actions = ['coldreset',
                     'power off/on',
                     'system reset',
                     'Ctrl1Eth1DisEn',
                     'CLK Init',
                     "portdisable",
                     "ptn eth port disable",
                     "config_download",
                     "warmreset"
                     ]
    for item in cisco_actions:
        add_item_in_test_summary(test_summary, item, "Bert")
    match_action = ''
    cur_action = 'Initial'
    rs485_all_dut_slot_info = init_rs485_slot_info(dut_count)
    try:
        f_read = open(file_path, 'r', encoding="utf-8")
    except Exception as e:
        print("error! {}".format(e))
    while True:
        time_stamp, text_line = read_line_and_split_time(f_read)
        if text_line == '':
            break
        if "bundle60_test" in text_line:
            status_item = [
                'Bundle test(bert)', 'Bundle test(PW status)', 'Bundle test(RS485 status)']
            for item in status_item:
                test_result[item] = {}
                add_item_in_test_summary(test_summary, item, item)
            vlan_mode = ""
            network_type = ""
            cur_dict = {}
            cisco_cur_action = ''
            time_stamp, text_line = read_line_and_split_time(f_read)
            while "-----------bundle60_test---------------End" not in text_line and text_line != '':
                if "Bundle test start with" in text_line:
                    if 'QE1 CAS' in text_line:
                        cfg = 'QE1 CAS'
                    else:
                        cfg = 'None'
                    vlan_mode = re.findall(
                        r'Vlan mode=\w+', text_line)[0].split('=')[1]
                    network_type = re.findall(
                        r'Network type=\w+', text_line)[0].split('=')[1].strip()

                    if cfg == 'QE1 CAS':
                        cur_test = F'{cfg}-{vlan_mode}-{network_type}'
                    else:
                        cur_test = F'{vlan_mode}-{network_type}'
                    test_result['Bundle test(bert)'][F'{cur_test}'] = {}
                    test_result['Bundle test(PW status)'][F'{cur_test}'] = {}
                    test_result['Bundle test(RS485 status)'][F'{cur_test}'] = {
                    }
                elif "Bundle $BID" in text_line:
                    # print(text_line.split("$BID="))
                    start_bundle = text_line.split("$BID=")[1].replace('~', '')
                    end_bundle = text_line.split("$BID=")[2].strip()
                    # print(start_bundle, end_bundle)
                    bundle_range = F'{start_bundle}-{end_bundle}'
                    if cfg == 'QE1 CAS':
                        test_result['Bundle test(bert)'][F'{cur_test}'][bundle_range] = {
                        }
                    else:
                        test_result['Bundle test(bert)'][F'{cur_test}'][bundle_range] = {
                        }
                    test_result['Bundle test(PW status)'][F'{cur_test}'][bundle_range] = {
                    }
                    test_result['Bundle test(RS485 status)'][F'{cur_test}'][bundle_range] = {
                    }
                elif '(AM3440_CCPA_Get_bert_log) begin' in text_line:
                    if cisco_cur_action not in test_result['Bundle test(bert)'][F'{cur_test}'][bundle_range]:
                        test_result['Bundle test(bert)'][F'{cur_test}'][bundle_range][cisco_cur_action] = {
                            'Ctrl Bert': {'pass': 0, 'fail': 0, 'fail_item': []},
                            'QE1 Bert': {'pass': 0, 'fail': 0, 'fail_item': []}
                        }
                    time_stamp, text_line = read_line_and_split_time(f_read)
                    cur_dict = test_result['Bundle test(bert)'][F'{cur_test}'][bundle_range][cisco_cur_action]
                    case = 'Bundle test(bert)'
                    while '(AM3440_CCPA_Get_bert_log) end' not in text_line and text_line != '':
                        if 'pass' in text_line:
                            cur_dict['Ctrl Bert']['pass'] += 1
                            test_summary[case]['pass'] += 1
                        elif 'fail' in text_line:
                            bert_fail_data = {}
                            cur_dict['Ctrl Bert']['fail'] += 1
                            test_summary[case]['fail'] += 1
                            bert_fail_data = ctrl_bert_fail_data_parse(
                                text_line, f_read)
                            cur_dict['Ctrl Bert']["fail_item"].append(
                                bert_fail_data)
                        time_stamp, text_line = read_line_and_split_time(
                            f_read)
                elif '(AM3440_CCPA_Get_QE1Bert_Snmp) begin' in text_line:
                    if cisco_cur_action not in test_result['Bundle test(bert)'][F'{cur_test}'][bundle_range]:
                        test_result['Bundle test(bert)'][F'{cur_test}'][bundle_range][cisco_cur_action] = {
                            'Ctrl Bert': {'pass': 0, 'fail': 0, 'fail_item': []},
                            'QE1 Bert': {'pass': 0, 'fail': 0, 'fail_item': []}
                        }
                    time_stamp, text_line = read_line_and_split_time(f_read)
                    cur_dict = test_result['Bundle test(bert)'][F'{cur_test}'][bundle_range][cisco_cur_action]
                    case = 'Bundle test(bert)'
                    while '(AM3440_CCPA_Get_QE1Bert_Snmp) end' not in text_line and text_line != '':
                        if 'pass' in text_line:
                            cur_dict['QE1 Bert']['pass'] += 1
                            test_summary[case]['pass'] += 1
                        elif 'fail' in text_line:
                            bert_fail_data = {}
                            cur_dict['QE1 Bert']['fail'] += 1
                            test_summary[case]['fail'] += 1
                            bert_fail_data = qe1_bert_fail_data_parse(
                                time_stamp, text_line)
                            cur_dict['QE1 Bert']["fail_item"].append(
                                bert_fail_data)
                        time_stamp, text_line = read_line_and_split_time(
                            f_read)
                elif 'Check PW status after bert start' in text_line or 'Check PW status before bert start' in text_line:
                    working_pw_array.clear()
                    standby_pw_array.clear()
                    rs485_array.clear()
                    cisco_cur_action = ''
                    if 'Check PW status after bert start' in text_line:
                        action = 'after bert'
                    elif 'Check PW status before bert start' in text_line:
                        action = 'before bert'
                    test_result['Bundle test(PW status)'][F'{cur_test}'][bundle_range][action] = {
                    }
                    test_result['Bundle test(PW status)'][F'{cur_test}'][bundle_range][action][cisco_cur_action] = {
                        'pass': 0,
                        'fail': 0,
                        'Working PW error status': [],
                        'Standby PW error status': []
                    }
                    test_result['Bundle test(RS485 status)'][F'{cur_test}'][bundle_range][action] = {
                    }
                    test_result['Bundle test(RS485 status)'][F'{cur_test}'][bundle_range][action][cisco_cur_action] = {
                        'pass': 0,
                        'fail': 0,
                        'RS485 error status': []
                    }
                elif 'Check PW status after bert end' in text_line or 'Check PW status before bert end' in text_line:
                    cur_dict = test_result['Bundle test(PW status)'][F'{cur_test}'][bundle_range][action][cisco_cur_action]
                    if len(working_pw_array) > 0 or len(standby_pw_array) > 0:
                        if len(working_pw_array) > 0:
                            cur_dict['Working PW error status'] = working_pw_array.copy()
                        if len(standby_pw_array) > 0:
                            cur_dict['Standby PW error status'] = standby_pw_array.copy()    
                        count_bundle_pw_test_summary(test_summary, working_pw_array,standby_pw_array)
                    for i in range(len(working_pw_array)):
                        if working_pw_array[i]['result'] == 'fail':
                            # test_summary['Bundle test(PW status)']['fail'] += 1
                            cur_dict['fail'] += 1
                        if working_pw_array[i]['result'] == 'ok':
                            # test_summary['Bundle test(PW status)']['pass'] += 1
                            cur_dict['pass'] += 1
                        if standby_pw_array[i]['result'] == 'fail':
                            # test_summary['Bundle test(PW status)']['fail'] += 1
                            cur_dict['fail'] += 1
                        if standby_pw_array[i]['result'] == 'ok':
                            # test_summary['Bundle test(PW status)']['pass'] += 1
                            cur_dict['pass'] += 1
                    cur_dict = test_result['Bundle test(RS485 status)'][F'{cur_test}'][bundle_range][action][cisco_cur_action]
                    if len(rs485_array) > 0:
                        cur_dict['RS485 error status'] = rs485_array.copy()
                        found = False
                        for a in range(len(rs485_array)):
                            if rs485_array[a]['result'] == 'fail':
                                test_summary['Bundle test(RS485 status)']['fail'] += 1
                                cur_dict['fail'] += 1
                                found = True
                            else:
                            #     test_summary['Bundle test(RS485 status)']['pass'] += 1
                                cur_dict['pass'] += 1
                        if found != True:
                            test_summary["Bundle test(RS485 status)"]['pass'] += 1
                    working_pw_array.clear()
                    standby_pw_array.clear()
                    rs485_array.clear()
                    action = ''
                    test_result['Bundle test(PW status)'][F'{cur_test}'][bundle_range][action] = {}
                    test_result['Bundle test(RS485 status)'][F'{cur_test}'][bundle_range][action] = {}  
                elif "AM3440_CCPA_pw_status_get" in text_line:
                    pw_data = pw_status_parse(time_stamp, text_line)
                    if pw_data:
                        working_pw_array.append(pw_data)
                elif "AM3440_CCPA_pw_Standbystatus_get_log" in text_line:
                    pw_data = pw_status_parse(time_stamp, text_line)
                    if pw_data:
                        standby_pw_array.append(pw_data)
                elif "AM3440_CCPA_rs485_get_statist_by_slot" in text_line:
                    rs485_data = rs485_slot_info_parse(
                        time_stamp, text_line, rs485_all_dut_slot_info)
                    if rs485_data:
                        rs485_array.append(rs485_data)
                elif "AM3440_CCPA_Get_Rs485_statist" in text_line:
                    rs485_data = rs485_dut_info_parse(
                        time_stamp, text_line)
                    if rs485_data:
                        rs485_array.append(rs485_data)
                elif "AM3440_CCPA_Rs485Delay_Clear_Snmp" in text_line:
                    rs485_all_dut_slot_info = init_rs485_slot_info(dut_count)
                elif '----------------- DUT' in text_line:
                    if any((curcard_action := caction) in text_line for caction in cisco_actions):
                        if curcard_action == 'power off/on':
                            cur_action = 'power reset'
                        elif curcard_action == 'system reset':
                            cur_action = 'system reset'
                        elif curcard_action == 'coldreset':
                            cur_action = 'coldreset'
                        elif curcard_action == 'Ctrl1Eth1DisEn':
                            cur_action = 'Ctrl1Eth1DisEn'
                        #elif curcard_action == 'CLK Init':
                            #cur_action = 'CLK Init'
                    else:
                        cur_action = ''         
                elif any((match := 'record after action') in text_line for stage in voice_stages) and " start" in text_line and " start" in text_line and cur_action != '':
                    cisco_cur_action = cur_action+'-'+match
                    if cisco_cur_action not in test_result['Bundle test(bert)'][F'{cur_test}'][bundle_range]:
                        test_result['Bundle test(bert)'][F'{cur_test}'][bundle_range][cisco_cur_action] = {
                            'Ctrl Bert': {'pass': 0, 'fail': 0, 'fail_item': []},
                            'QE1 Bert': {'pass': 0, 'fail': 0, 'fail_item': []}
                        }
                    if cisco_cur_action not in test_result['Bundle test(PW status)'][F'{cur_test}'][bundle_range][action]:
                        test_result['Bundle test(PW status)'][F'{cur_test}'][bundle_range][action][cisco_cur_action] = {
                            'pass': 0,
                            'fail': 0,
                            'Working PW error status': [],
                            'Standby PW error status': []
                        }
                    if cisco_cur_action not in test_result['Bundle test(RS485 status)'][F'{cur_test}'][bundle_range][action]:
                        test_result['Bundle test(RS485 status)'][F'{cur_test}'][bundle_range][action][cisco_cur_action] = {
                            'pass': 0,
                            'fail': 0,
                            'RS485 error status': []
                        }
                elif "record after action end" in text_line and text_line != '':
                    if cur_action != '':
                        cur_dict = test_result['Bundle test(PW status)'][F'{cur_test}'][bundle_range][action][cisco_cur_action]
                        if len(working_pw_array) > 0 or len(standby_pw_array) > 0:
                            if len(working_pw_array) > 0:
                                cur_dict['Working PW error status'].extend(working_pw_array)
                            if len(standby_pw_array) > 0:
                                cur_dict['Standby PW error status'].extend(standby_pw_array)
                            count_bundle_pw_test_summary(test_summary, working_pw_array,standby_pw_array)
                        for i in range(len(working_pw_array)):
                            if working_pw_array[i]['result'] == 'fail':
                                # test_summary['Bundle test(PW status)']['fail'] += 1
                                cur_dict['fail'] += 1
                            if working_pw_array[i]['result'] == 'ok':
                                # test_summary['Bundle test(PW status)']['pass'] += 1
                                cur_dict['pass'] += 1
                        for i in range(len(standby_pw_array)):
                            if standby_pw_array[i]['result'] == 'fail':
                                # test_summary['Bundle test(PW status)']['fail'] += 1
                                cur_dict['fail'] += 1
                            if standby_pw_array[i]['result'] == 'ok':
                                # test_summary['Bundle test(PW status)']['pass'] += 1
                                cur_dict['pass'] += 1
                        cur_dict = test_result['Bundle test(RS485 status)'][F'{cur_test}'][bundle_range][action][cisco_cur_action]
                        if len(rs485_array) > 0:
                            cur_dict['RS485 error status'] = rs485_array.copy()
                            found = False
                            for a in range(len(rs485_array)):
                                if rs485_array[a]['result'] == 'fail':
                                    test_summary['Bundle test(RS485 status)']['fail'] += 1
                                    cur_dict['fail'] += 1
                                    found = True
                                else:
                                #     test_summary['Bundle test(RS485 status)']['pass'] += 1
                                    cur_dict['pass'] += 1
                            if found != True:
                                test_summary["Bundle test(RS485 status)"]['pass'] += 1
                    working_pw_array.clear()
                    standby_pw_array.clear()
                    rs485_array.clear()
                        #time_stamp, text_line = read_line_and_split_time(f_read)
                elif any((match := stage) in text_line for stage in voice_stages) and match +" end" in text_line and text_line != '':
                    working_pw_array.clear()
                    standby_pw_array.clear()
                    rs485_array.clear()
                    cisco_cur_action = ''
                time_stamp, text_line = read_line_and_split_time(f_read)
            # print(json.dumps(test_result[case],indent=4))
        elif "pdh_ring_case" in text_line:
            case = 'PDH ring'
            test_result[case] = {}
            add_item_in_test_summary(test_summary, case, case)
            time_stamp, text_line = read_line_and_split_time(f_read)
            while "-----------pdh_ring_case---------------End" not in text_line and text_line != '':
                if any((match := status) in text_line for status in pdh_stage):
                    # if match == 'Get bert status after reset':
                    #     status = 'Reset bert after ' + status
                    #     # print(status)
                    # else:
                    status = match
                    if status not in test_result[case]:
                        test_result[case][status] = {
                            'pass': 0, 'fail': 0, 'fail_item': []}
                elif "(AM3440_CCPA_Get_bert_log) begin" in text_line:
                    time_stamp, text_line = read_line_and_split_time(f_read)
                    while '(AM3440_CCPA_Get_bert_log) end' not in text_line and text_line != '':
                        if 'Bert pass' in text_line:
                            if status == 'Set QE1 port disable' or status == 'Set QE1 port enable':
                                test_result[case][status]['fail'] += 1
                                test_summary[case]['fail'] += 1
                            else:
                                test_result[case][status]['pass'] += 1
                                test_summary[case]['pass'] += 1
                        elif 'fail' in text_line:
                            if status == 'Set QE1 port disable' or status == 'Set QE1 port enable':
                                test_result[case][status]['pass'] += 1
                                test_summary[case]['pass'] += 1
                            else:
                                bert_fail_data = {}
                                test_result[case][status]['fail'] += 1
                                test_summary[case]['fail'] += 1
                                bert_fail_data = ctrl_bert_fail_data_parse(
                                    text_line, f_read)
                                test_result[case][status]['fail_item'].append(
                                    bert_fail_data)
                        time_stamp, text_line = read_line_and_split_time(
                            f_read)
                time_stamp, text_line = read_line_and_split_time(f_read)
        elif "QE1 all case setup end" in text_line:
            # print(time_stamp,text_line)
            status_item = [
                'Long Time(bert)', 'Long Time(PW status)', 'Long Time(RS485 status)']
            round = 0
            for item in status_item:
                if item not in test_result and item not in test_summary:
                    test_result[item] = {}
                    add_item_in_test_summary(test_summary, item, item)
            # case = 'Long Time'
            # add_item_in_test_summary(test_summary, case, case)
            # test_result[case] = {}
            time_stamp, text_line = read_line_and_split_time(f_read)
            while "Round@" not in text_line and 'fxs/fxo' not in text_line:
                # print(time_stamp,text_line)
                if 'Current $bert_chk_cnt:' in text_line:
                    # print(time_stamp,text_line)
                    cisco_cur_action = ''
                    if round != 0:
                        if len(rs485_array) > 0:
                            test_result['Long Time(RS485 status)'][round][cisco_cur_action]['RS485 error status'] = rs485_array.copy(
                            )
                            found = False
                            for a in range(len(rs485_array)):
                                if rs485_array[a]['result'] == 'fail':
                                    test_summary['Long Time(RS485 status)']['fail'] += 1
                                    test_result['Long Time(RS485 status)'][round][cisco_cur_action]['fail'] += 1
                                    found = True
                                else:
                                    test_result['Long Time(RS485 status)'][round][cisco_cur_action]['pass'] += 1
                            if found != True:
                                test_summary["Long Time(RS485 status)"]['pass'] += 1
                        if len(working_pw_array) > 0 or len(standby_pw_array) > 0:
                            if len(working_pw_array) > 0:
                                test_result['Long Time(PW status)'][round][cisco_cur_action]['Working PW error status'] = working_pw_array.copy(
                                )
                            if len(standby_pw_array) > 0:
                                test_result['Long Time(PW status)'][round][cisco_cur_action]['Standby PW error status'] = standby_pw_array.copy(
                                )
                            count_longtime_pw_test_summary(test_summary, working_pw_array,standby_pw_array)
                        for i in range(len(working_pw_array)):
                            if working_pw_array[i]['result'] == 'fail':
                                # test_summary['Bundle test(PW status)']['fail'] += 1
                                test_result['Long Time(PW status)'][round][cisco_cur_action]['fail'] += 1
                            if working_pw_array[i]['result'] == 'ok':
                                # test_summary['Bundle test(PW status)']['pass'] += 1
                                test_result['Long Time(PW status)'][round][cisco_cur_action]['pass'] += 1
                            if standby_pw_array[i]['result'] == 'fail':
                                # test_summary['Bundle test(PW status)']['fail'] += 1
                                test_result['Long Time(PW status)'][round][cisco_cur_action]['fail'] += 1
                            if standby_pw_array[i]['result'] == 'ok':
                                # test_summary['Bundle test(PW status)']['pass'] += 1
                                test_result['Long Time(PW status)'][round][cisco_cur_action]['pass'] += 1                     
                    working_pw_array.clear()
                    standby_pw_array.clear()
                    rs485_array.clear()
                    round = text_line.split('Current $bert_chk_cnt:')[
                        1].strip('***\n')
                    # print(round)
                    test_result['Long Time(bert)'][round] = {}
                    test_result['Long Time(PW status)'][round] = {}
                    test_result['Long Time(RS485 status)'][round] = {}
                    if cisco_cur_action not in test_result['Long Time(bert)'][round]:
                        test_result['Long Time(bert)'][round][cisco_cur_action] = {
                            'Ctrl Bert': {'pass': 0, 'fail': 0, 'fail_item': []},
                            'QE1 Bert': {'pass': 0, 'fail': 0, 'fail_item': []}
                        }
                    if cisco_cur_action not in test_result['Long Time(PW status)'][round]:
                        test_result['Long Time(PW status)'][round][cisco_cur_action] = {
                            'pass': 0,
                            'fail': 0,
                            'Working PW error status': [],
                            'Standby PW error status': []
                        }
                    if cisco_cur_action not in test_result['Long Time(RS485 status)'][round]:
                        test_result['Long Time(RS485 status)'][round][cisco_cur_action] = {
                            'pass': 0,
                            'fail': 0,
                            'RS485 error status': []
                        }
                elif '***Current $bert_chk_cnt DONE***' in text_line:
                    if round != 0:
                        cur_dict = test_result['Long Time(PW status)'][round][cisco_cur_action]
                        if len(working_pw_array) > 0 or len(standby_pw_array) > 0:
                            if len(working_pw_array) > 0:
                                cur_dict['Working PW error status'].extend(working_pw_array)
                            if len(standby_pw_array) > 0:
                                cur_dict['Standby PW error status'].extend(standby_pw_array)
                            count_longtime_pw_test_summary(test_summary, working_pw_array,standby_pw_array)
                        for i in range(len(working_pw_array)):
                            if working_pw_array[i]['result'] == 'fail':
                                # test_summary['Long Time(PW status)']['fail'] += 1
                                cur_dict['fail'] += 1
                            if working_pw_array[i]['result'] == 'ok':
                                # test_summary['Long Time(PW status)']['pass'] += 1
                                cur_dict['pass'] += 1
                        for i in range(len(standby_pw_array)):
                            if standby_pw_array[i]['result'] == 'fail':
                                # test_summary['Long Time(PW status)']['fail'] += 1
                                cur_dict['fail'] += 1
                            if standby_pw_array[i]['result'] == 'ok':
                                # test_summary['Long Time(PW status)']['pass'] += 1
                                cur_dict['pass'] += 1
                        cur_dict = test_result['Long Time(RS485 status)'][round][cisco_cur_action]
                        if len(rs485_array) > 0:
                            cur_dict['RS485 error status'] = rs485_array.copy()
                            found = False
                            for a in range(len(rs485_array)):
                                if rs485_array[a]['result'] == 'fail':
                                    test_summary['Long Time(RS485 status)']['fail'] += 1
                                    cur_dict['fail'] += 1
                                    found = True
                                else:
                                #     test_summary['Long Time(RS485 status)']['pass'] += 1
                                    cur_dict['pass'] += 1
                            if found != True:
                                test_summary["Long Time(RS485 status)"]['pass'] += 1
                    working_pw_array.clear()
                    standby_pw_array.clear()
                    rs485_array.clear()  
                    cur_action = ''
                elif "AM3440_CCPA_Rs485Delay_Clear_Snmp" in text_line:
                    rs485_all_dut_slot_info = init_rs485_slot_info(dut_count)
                elif 'AM3440_CCPA_rs485_get_statist_by_slot' in text_line:
                    rs485_data = rs485_slot_info_parse(time_stamp, text_line, rs485_all_dut_slot_info)
                    if rs485_data:
                        rs485_array.append(rs485_data)
                elif 'AM3440_CCPA_Get_Rs485_statist' in text_line:
                    rs485_data = rs485_dut_info_parse(time_stamp, text_line)
                    if rs485_data:
                        rs485_array.append(rs485_data)
                elif 'AM3440_CCPA_pw_status_get' in text_line:
                    pw_data = pw_status_parse(time_stamp, text_line)
                    if pw_data:
                        working_pw_array.append(pw_data)
                elif 'AM3440_CCPA_pw_Standbystatus_get_log' in text_line:
                    pw_data = pw_status_parse(time_stamp, text_line)
                    if pw_data:
                        standby_pw_array.append(pw_data)
                elif 'get bert start' in text_line:
                    case = 'Long Time(bert)'
                    if cisco_cur_action not in test_result['Long Time(bert)'][round]:
                        test_result['Long Time(bert)'][round][cisco_cur_action] = {
                            'Ctrl Bert': {'pass': 0, 'fail': 0, 'fail_item': []},
                            'QE1 Bert': {'pass': 0, 'fail': 0, 'fail_item': []}
                        }
                    time_stamp, text_line = read_line_and_split_time(f_read)
                    while ('get bert end' not in text_line and text_line != ''):
                        if '(AM3440_CCPA_Get_QE1Bert_Snmp) begin' in text_line:
                            time_stamp, text_line = read_line_and_split_time(
                                f_read)
                            while '(AM3440_CCPA_Get_QE1Bert_Snmp) end' not in text_line and text_line != '':
                                if 'pass' in text_line:
                                    test_result[case][round][cisco_cur_action]['QE1 Bert']['pass'] += 1
                                    test_summary[case]['pass'] += 1
                                elif 'fail' in text_line:
                                    bert_fail_data = qe1_bert_fail_data_parse(
                                        time_stamp, text_line)
                                    test_result[case][round][cisco_cur_action]['QE1 Bert']['fail'] += 1
                                    test_summary[case]['fail'] += 1
                                    test_result[case][round][cisco_cur_action]['QE1 Bert']["fail_item"].append(bert_fail_data)
                                time_stamp, text_line = read_line_and_split_time(
                                    f_read)
                        elif '(AM3440_CCPA_Get_bert_log) begin' in text_line:
                            time_stamp, text_line = read_line_and_split_time(
                                f_read)
                            while '(AM3440_CCPA_Get_bert_log) end' not in text_line and text_line != '':
                                if 'Bert pass' in text_line:
                                    test_result[case][round][cisco_cur_action]['Ctrl Bert']['pass'] += 1
                                    test_summary[case]['pass'] += 1
                                elif 'fail' in text_line:
                                    bert_fail_data = {}
                                    test_result[case][round][cisco_cur_action]['Ctrl Bert']['fail'] += 1
                                    test_summary[case]['fail'] += 1
                                    bert_fail_data = ctrl_bert_fail_data_parse(
                                        text_line, f_read)
                                    test_result[case][round][cisco_cur_action]['Ctrl Bert']["fail_item"].append(
                                        bert_fail_data)
                                time_stamp, text_line = read_line_and_split_time(
                                    f_read)
                        time_stamp, text_line = read_line_and_split_time(f_read)
                elif 'get bert after reset start' in text_line:
                    case = 'Long Time(bert)'
                    if cisco_cur_action not in test_result['Long Time(bert)'][round]:
                        test_result['Long Time(bert)'][round][cisco_cur_action] = {
                            'Ctrl Bert': {'pass': 0, 'fail': 0, 'fail_item': []},
                            'QE1 Bert': {'pass': 0, 'fail': 0, 'fail_item': []}
                        }
                    time_stamp, text_line = read_line_and_split_time(f_read)
                    while 'get bert after reset end' not in text_line and text_line != '':
                        if '(AM3440_CCPA_Get_QE1Bert_Snmp) begin' in text_line:
                            time_stamp, text_line = read_line_and_split_time(
                                f_read)
                            while '(AM3440_CCPA_Get_QE1Bert_Snmp) end' not in text_line and text_line != '':
                                if 'pass' in text_line:
                                    test_result[case][round][cisco_cur_action]['QE1 Bert']['pass'] += 1
                                    test_summary[case]['pass'] += 1
                                elif 'fail' in text_line:
                                    bert_fail_data = qe1_bert_fail_data_parse(
                                        time_stamp, text_line)
                                    test_result[case][round][cisco_cur_action]['QE1 Bert']['fail'] += 1
                                    test_summary[case]['fail'] += 1
                                    test_result[case][round][cisco_cur_action]['QE1 Bert']["fail_item"].append(
                                        bert_fail_data)
                                time_stamp, text_line = read_line_and_split_time(
                                    f_read)
                        elif '(AM3440_CCPA_Get_bert_log) begin' in text_line:
                            time_stamp, text_line = read_line_and_split_time(
                                f_read)
                            while '(AM3440_CCPA_Get_bert_log) end' not in text_line and text_line != '':
                                if 'Bert pass' in text_line:
                                    test_result[case][round][cisco_cur_action]['Ctrl Bert']['pass'] += 1
                                    test_summary[case]['pass'] += 1
                                elif 'fail' in text_line:
                                    bert_fail_data = {}
                                    test_result[case][round][cisco_cur_action]['Ctrl Bert']['fail'] += 1
                                    test_summary[case]['fail'] += 1
                                    bert_fail_data = ctrl_bert_fail_data_parse(
                                        text_line, f_read)
                                    test_result[case][round][cisco_cur_action]['Ctrl Bert']["fail_item"].append(
                                        bert_fail_data)
                                time_stamp, text_line = read_line_and_split_time(
                                    f_read)
                        time_stamp, text_line = read_line_and_split_time(
                            f_read)
                elif '----------------- DUT' in text_line:
                    if any((curcard_action := caction) in text_line for caction in cisco_actions):
                        if curcard_action == 'power off/on':
                            cur_action = 'power reset'
                        elif curcard_action == 'system reset':
                            cur_action = 'system reset'
                        elif curcard_action == 'coldreset':
                            cur_action = 'coldreset'
                        elif curcard_action == 'Ctrl1Eth1DisEn':
                            cur_action = 'Ctrl1Eth1DisEn'
                        #elif curcard_action == 'CLK Init':
                            #cur_action = 'CLK Init'
                    else:
                        cur_action = ''        
                elif any((match := 'record after action') in text_line for stage in voice_stages) and " start" in text_line and cur_action != '':
                    cisco_cur_action = cur_action+'-'+match
                    if cisco_cur_action not in test_result['Long Time(bert)'][round]:
                        test_result['Long Time(bert)'][round][cisco_cur_action] = {
                            'Ctrl Bert': {'pass': 0, 'fail': 0, 'fail_item': []},
                            'QE1 Bert': {'pass': 0, 'fail': 0, 'fail_item': []}
                        }
                    if cisco_cur_action not in test_result['Long Time(PW status)'][round]:
                        test_result['Long Time(PW status)'][round][cisco_cur_action] = {
                            'pass': 0,
                            'fail': 0,
                            'Working PW error status': [],
                            'Standby PW error status': []
                        }
                    if cisco_cur_action not in test_result['Long Time(RS485 status)'][round]:
                        test_result['Long Time(RS485 status)'][round][cisco_cur_action] = {
                            'pass': 0,
                            'fail': 0,
                            'RS485 error status': []
                        }
                elif "record after action end" in text_line and text_line != '':
                    if cur_action != '':
                        cur_dict = test_result['Long Time(PW status)'][round][cisco_cur_action]
                        if len(working_pw_array) > 0 or len(standby_pw_array) > 0:
                            if len(working_pw_array) > 0:
                                cur_dict['Working PW error status'].extend(working_pw_array)
                            if len(standby_pw_array) > 0:
                                cur_dict['Standby PW error status'].extend(standby_pw_array)
                            count_longtime_pw_test_summary(test_summary, working_pw_array,standby_pw_array)
                        for i in range(len(working_pw_array)):
                            if working_pw_array[i]['result'] == 'fail':
                                # test_summary['Long Time(PW status)']['fail'] += 1
                                cur_dict['fail'] += 1
                            if working_pw_array[i]['result'] == 'ok':
                                # test_summary['Long Time(PW status)']['pass'] += 1
                                cur_dict['pass'] += 1
                        for i in range(len(standby_pw_array)):
                            if standby_pw_array[i]['result'] == 'fail':
                                # test_summary['Long Time(PW status)']['fail'] += 1
                                cur_dict['fail'] += 1
                            if standby_pw_array[i]['result'] == 'ok':
                                # test_summary['Long Time(PW status)']['pass'] += 1
                                cur_dict['pass'] += 1
                        cur_dict = test_result['Long Time(RS485 status)'][round][cisco_cur_action]
                        if len(rs485_array) > 0:
                            cur_dict['RS485 error status'].extend(rs485_array)
                            found = False
                            for a in range(len(rs485_array)):
                                if rs485_array[a]['result'] == 'fail':
                                    test_summary['Long Time(RS485 status)']['fail'] += 1
                                    cur_dict['fail'] += 1
                                    found = True
                                else:
                                #     test_summary['Long Time(RS485 status)']['pass'] += 1
                                    cur_dict['pass'] += 1
                            if found != True:
                                test_summary["Long Time(RS485 status)"]['pass'] += 1
                    working_pw_array.clear()
                    standby_pw_array.clear()
                    rs485_array.clear()  
                elif any((match := stage) in text_line for stage in voice_stages) and match +" end" in text_line and text_line != '':
                    working_pw_array.clear()
                    standby_pw_array.clear()
                    rs485_array.clear()
                    cur_action = ''
                time_stamp, text_line = read_line_and_split_time(f_read)
            time_stamp, text_line = read_line_and_split_time(f_read)
                # ======= voice test =======
            if "Round@" in text_line and "fxs/fxo" in text_line and "Begin" not in text_line:
                # print(time_stamp,text_line)
                status_item = ['PW status(Voice test)', 'RS485 status(Voice test)']
                for item in status_item:
                    test_result[item] = {}
                    add_item_in_test_summary(test_summary, item, item)
                round_num = re.findall(r'Round@\s+(\d+)', text_line)
                cur_action = ''
                for match in round_num:
                    cur_round = str(int(match))
                test_result['PW status(Voice test)'][cur_round] = {}
                test_result['RS485 status(Voice test)'][cur_round] = {}
                time_stamp, text_line = read_line_and_split_time(f_read)
                while 'DONE---------' not in text_line:
                    if '----------------- DUT' in text_line and any((curcard_action := action) in text_line for action in cisco_actions):
                        #print(time_stamp, text_line)
                        if curcard_action == 'power off/on':
                            cur_action = 'power reset'
                        elif curcard_action == 'system reset':
                            cur_action = 'system reset'
                        elif curcard_action == 'coldreset':
                            cur_action = 'coldreset'
                        #elif curcard_action == 'CLK Init':
                            #cur_action = 'CLK Initialize'
                    if any((match := 'record after action') in text_line for stage in voice_stages):
                        cur_round_stage = cur_round+'-'+cur_action+'-'+match
                        cisco_cur_action = cur_action+'-'+match
                        if cisco_cur_action not in test_result['PW status(Voice test)'][cur_round]:
                            test_result['PW status(Voice test)'][cur_round][cisco_cur_action] = {
                                'pass': 0,
                                'fail': 0,
                                'Working PW error status': [],
                                'Standby PW error status': []
                            }
                        if cisco_cur_action not in test_result['RS485 status(Voice test)']:
                            test_result['RS485 status(Voice test)'][cur_round][cisco_cur_action] = {
                                'pass': 0,
                                'fail': 0,
                                'RS485 error status': []
                            }
                        working_pw_array.clear()
                        standby_pw_array.clear()
                        rs485_array.clear()
                        time_stamp, text_line = read_line_and_split_time(f_read)
                        while match +" end" not in text_line and text_line != '':
                            # --- tone-test --- 
                            if "********tone-test*********" in text_line:
                                if "Tone" not in test_summary:
                                    add_item_in_test_summary(test_summary, "Tone", "Tone")
                                    test_result["Tone"] = {"Total": 0, "Items": {}}
                                if cur_round_stage not in test_result["Tone"]["Items"]:
                                    test_result["Tone"]["Items"][cur_round_stage] = {"u-law_900-ohm": {"pass": 0, "fail": 0, "fail_item": []}, 
                                                                                "u-law_600-ohm": {"pass": 0, "fail": 0, "fail_item": []},
                                                                                "a-law_900-ohm": {"pass": 0, "fail": 0, "fail_item": []},
                                                                                "a-law_600-ohm": {"pass": 0, "fail": 0, "fail_item": []}}
                                for aulaw in ["u-law","a-law"]:
                                    if aulaw in text_line:
                                        for ohm_type in ["900-ohm", "600-ohm"]:
                                            if ohm_type in text_line:
                                                time_stamp, text_line = read_line_and_split_time(f_read)
                                                while "********tone-test done*********" not in text_line:
                                                    if "tone pass" in text_line or "tone test pass" in text_line:
                                                        test_result["Tone"]["Total"] += 1
                                                        test_result["Tone"]["Items"][cur_round_stage][f"{aulaw}_{ohm_type}"]["pass"] += 1
                                                        test_summary["Tone"]["pass"] += 1
                                                    elif "tone fail" in text_line:
                                                        test_result["Tone"]["Total"] += 1
                                                        test_result["Tone"]["Items"][cur_round_stage][f"{aulaw}_{ohm_type}"]["fail"] += 1
                                                        test_result["Tone"]["Items"][cur_round_stage][f"{aulaw}_{ohm_type}"]["fail_item"].append(text_line)
                                                        test_summary["Tone"]["fail"] += 1
                                                    time_stamp, text_line = read_line_and_split_time(f_read)
                            # --- tone-test ---     
                            # --- ring-test ---
                            elif "********ring-test*********" in text_line:
                                if "Ring" not in test_summary:
                                    add_item_in_test_summary(test_summary, "Ring", "Ring")
                                    test_result["Ring"] = {"Total": 0, "Items": {}}
                                if cur_round_stage not in test_result["Ring"]["Items"]:
                                    test_result["Ring"]["Items"][cur_round_stage] = {"u-law_900-ohm": {"pass": 0, "fail": 0, "fail_item": []}, 
                                                                                "u-law_600-ohm": {"pass": 0, "fail": 0, "fail_item": []},
                                                                                "a-law_900-ohm": {"pass": 0, "fail": 0, "fail_item": []},
                                                                                "a-law_600-ohm": {"pass": 0, "fail": 0, "fail_item": []}}
                                for aulaw in ["u-law","a-law"]:
                                    if aulaw in text_line:
                                        for ohm_type in ["900-ohm", "600-ohm"]:
                                            if ohm_type in text_line:
                                                time_stamp, text_line = read_line_and_split_time(f_read)
                                                while "********ring-test done*********" not in text_line:
                                                    if "TS" in text_line:
                                                        TS_text = text_line
                                                    #time_stamp, text_line = read_line_and_split_time(f_read)
                                                    if "ringing=yes,pass!" in text_line:
                                                        test_result["Ring"]["Total"] += 1
                                                        test_result["Ring"]["Items"][cur_round_stage][f"{aulaw}_{ohm_type}"]["pass"] += 1
                                                        test_summary["Ring"]["pass"] += 1
                                                    elif "ringing=no" in text_line:
                                                        test_result["Ring"]["Total"] += 1
                                                        test_result["Ring"]["Items"][cur_round_stage][f"{aulaw}_{ohm_type}"]["fail"] += 1
                                                        test_result["Ring"]["Items"][cur_round_stage][f"{aulaw}_{ohm_type}"]["fail_item"].append(
                                                            time_stamp + ' ' + text_line.strip() + ' ' + TS_text)
                                                        test_summary["Ring"]["fail"] += 1
                                                    time_stamp, text_line = read_line_and_split_time(f_read)
                            # --- ring-test ---
                            # --- Voice BERT ---
                            elif "********bert*********" in text_line:
                                if "Voice Bert" not in test_summary:
                                    add_item_in_test_summary(test_summary, "Voice Bert", "Voice Bert")
                                    test_result["Voice Bert"] = {"Total": 0, "Items": {}}
                                if cur_round_stage not in test_result["Voice Bert"]["Items"]:
                                    test_result["Voice Bert"]["Items"][cur_round_stage] = {"u-law_900-ohm": {"pass": 0, "fail": 0, "fail_item": []}, 
                                                                                "u-law_600-ohm": {"pass": 0, "fail": 0, "fail_item": []},
                                                                                "a-law_900-ohm": {"pass": 0, "fail": 0, "fail_item": []},
                                                                                "a-law_600-ohm": {"pass": 0, "fail": 0, "fail_item": []}}
                                for aulaw in ["u-law","a-law"]:
                                    if aulaw in text_line:
                                        for ohm_type in ["900-ohm", "600-ohm"]:
                                            if ohm_type in text_line:
                                                time_stamp, text_line = read_line_and_split_time(f_read)
                                                while "********bert done*********" not in text_line:
                                                    if "TS" in text_line:
                                                        TS_text = text_line
                                                    #time_stamp, text_line = read_line_and_split_time(f_read)
                                                    if "Fail" in text_line:
                                                        test_result["Voice Bert"]["Total"] += 1
                                                        test_result["Voice Bert"]["Items"][cur_round_stage][f"{aulaw}_{ohm_type}"]["fail"] += 1
                                                        test_result["Voice Bert"]["Items"][cur_round_stage][f"{aulaw}_{ohm_type}"]["fail_item"].append(
                                                            time_stamp + ' ' + text_line.strip() + ' ' + TS_text)
                                                        test_summary["Voice Bert"]["fail"] += 1
                                                    elif "Pass" in text_line:
                                                        test_result["Voice Bert"]["Total"] += 1
                                                        test_result["Voice Bert"]["Items"][cur_round_stage][f"{aulaw}_{ohm_type}"]["pass"] += 1
                                                        test_summary["Voice Bert"]["pass"] += 1
                                                    time_stamp, text_line = read_line_and_split_time(f_read)
                            time_stamp, text_line = read_line_and_split_time(f_read)
                            # elif "AM3440_CCPA_Rs485Delay_Clear_Snmp" in text_line:
                            #     rs485_all_dut_slot_info = init_rs485_slot_info(dut_count)
                            # elif 'AM3440_CCPA_rs485_get_statist_by_slot' in text_line:
                            #     rs485_data = rs485_slot_info_parse(time_stamp, text_line, rs485_all_dut_slot_info)
                            #     if rs485_data:
                            #         rs485_array.append(rs485_data)
                            # elif 'AM3440_CCPA_Get_Rs485_statist' in text_line:
                            #     rs485_data = rs485_dut_info_parse(time_stamp, text_line)
                            #     if rs485_data:
                            #         rs485_array.append(rs485_data)
                            # elif 'AM3440_CCPA_pw_status_get' in text_line:
                            #     pw_data = pw_status_parse(time_stamp, text_line)
                            #     if pw_data:
                            #         working_pw_array.append(pw_data)
                            # elif 'AM3440_CCPA_pw_Standbystatus_get_log' in text_line:
                            #     pw_data = pw_status_parse(time_stamp, text_line)
                            #     if pw_data:
                            #         standby_pw_array.append(pw_data)                                
                            # time_stamp, text_line = read_line_and_split_time(f_read)
                            # if "record after action end" in text_line and text_line != '':
                            #     if cur_action != '':
                            #         print(cur_round,cisco_cur_action,time_stamp)
                            #         cur_dict = test_result['PW status(Voice test)'][cur_round][cisco_cur_action]
                            #         if len(working_pw_array) > 0 or len(standby_pw_array) > 0:
                            #             if len(working_pw_array) > 0:
                            #                 cur_dict['Working PW error status'].extend(working_pw_array)
                            #             if len(standby_pw_array) > 0:
                            #                 cur_dict['Standby PW error status'].extend(standby_pw_array)
                            #             count_voice_pw_test_summary(test_summary, working_pw_array,standby_pw_array)
                            #         for i in range(len(working_pw_array)):
                            #             if working_pw_array[i]['result'] == 'fail':
                            #                 # test_summary['Long Time(PW status)']['fail'] += 1
                            #                 cur_dict['fail'] += 1
                            #             if working_pw_array[i]['result'] == 'ok':
                            #                 # test_summary['Long Time(PW status)']['pass'] += 1
                            #                 cur_dict['pass'] += 1
                            #         for i in range(len(standby_pw_array)):
                            #             if standby_pw_array[i]['result'] == 'fail':
                            #                 # test_summary['Long Time(PW status)']['fail'] += 1
                            #                 cur_dict['fail'] += 1
                            #             if standby_pw_array[i]['result'] == 'ok':
                            #                 # test_summary['Long Time(PW status)']['pass'] += 1
                            #                 cur_dict['pass'] += 1
                            #         cur_dict = test_result['RS485 status(Voice test)'][cur_round][cisco_cur_action]
                            #         if len(rs485_array) > 0:
                            #             cur_dict['RS485 error status'] = rs485_array.copy()
                            #             found = False
                            #             for a in range(len(rs485_array)):
                            #                 if rs485_array[a]['result'] == 'fail':
                            #                     test_summary['RS485 status(Voice test)']['fail'] += 1
                            #                     cur_dict['fail'] += 1
                            #                     found = True
                            #                 else:
                            #                 #     test_summary['Long Time(RS485 status)']['pass'] += 1
                            #                     cur_dict['pass'] += 1
                            #             if found != True:
                            #                 test_summary["RS485 status(Voice test)"]['pass'] += 1
                            #         # print(test_summary['PW status(Voice test)'])
                            #     working_pw_array.clear()
                            #     standby_pw_array.clear()
                            #     rs485_array.clear()
                        match = ''
                        cur_action = ''
                    if any((match := stage) in text_line for stage in voice_stages) and match != 'record after action':
                        #print(time_stamp, text_line)
                        cur_round_stage = cur_round+'-'+' '+'-'+match
                        time_stamp, text_line = read_line_and_split_time(f_read)
                        while match +" end" not in text_line and text_line != '':
                            if "********tone-test*********" in text_line:
                                if "Tone" not in test_summary:
                                    add_item_in_test_summary(test_summary, "Tone", "Tone")
                                    test_result["Tone"] = {"Total": 0, "Items": {}}
                                if cur_round_stage not in test_result["Tone"]["Items"]:
                                    test_result["Tone"]["Items"][cur_round_stage] = {"u-law_900-ohm": {"pass": 0, "fail": 0, "fail_item": []}, 
                                                                                "u-law_600-ohm": {"pass": 0, "fail": 0, "fail_item": []},
                                                                                "a-law_900-ohm": {"pass": 0, "fail": 0, "fail_item": []},
                                                                                "a-law_600-ohm": {"pass": 0, "fail": 0, "fail_item": []}}
                                for aulaw in ["u-law","a-law"]:
                                    if aulaw in text_line:
                                        for ohm_type in ["900-ohm", "600-ohm"]:
                                            if ohm_type in text_line:
                                                time_stamp, text_line = read_line_and_split_time(f_read)
                                                while "********tone-test done*********" not in text_line:
                                                    if "tone pass" in text_line or "tone test pass" in text_line:
                                                        test_result["Tone"]["Total"] += 1
                                                        test_result["Tone"]["Items"][cur_round_stage][f"{aulaw}_{ohm_type}"]["pass"] += 1
                                                        test_summary["Tone"]["pass"] += 1
                                                    elif "tone fail" in text_line:
                                                        test_result["Tone"]["Total"] += 1
                                                        test_result["Tone"]["Items"][cur_round_stage][f"{aulaw}_{ohm_type}"]["fail"] += 1
                                                        test_result["Tone"]["Items"][cur_round_stage][f"{aulaw}_{ohm_type}"]["fail_item"].append(text_line)
                                                        test_summary["Tone"]["fail"] += 1
                                                    time_stamp, text_line = read_line_and_split_time(f_read)
                            # --- tone-test ---     
                            # --- ring-test ---
                            elif "********ring-test*********" in text_line:
                                if "Ring" not in test_summary:
                                    add_item_in_test_summary(test_summary, "Ring", "Ring")
                                    test_result["Ring"] = {"Total": 0, "Items": {}}
                                if cur_round_stage not in test_result["Ring"]["Items"]:
                                    test_result["Ring"]["Items"][cur_round_stage] = {"u-law_900-ohm": {"pass": 0, "fail": 0, "fail_item": []}, 
                                                                                "u-law_600-ohm": {"pass": 0, "fail": 0, "fail_item": []},
                                                                                "a-law_900-ohm": {"pass": 0, "fail": 0, "fail_item": []},
                                                                                "a-law_600-ohm": {"pass": 0, "fail": 0, "fail_item": []}}
                                for aulaw in ["u-law","a-law"]:
                                    if aulaw in text_line:
                                        for ohm_type in ["900-ohm", "600-ohm"]:
                                            if ohm_type in text_line:
                                                time_stamp, text_line = read_line_and_split_time(f_read)
                                                while "********ring-test done*********" not in text_line:
                                                    if "TS" in text_line:
                                                        TS_text = text_line
                                                    #time_stamp, text_line = read_line_and_split_time(f_read)
                                                    if "ringing=yes,pass!" in text_line:
                                                        test_result["Ring"]["Total"] += 1
                                                        test_result["Ring"]["Items"][cur_round_stage][f"{aulaw}_{ohm_type}"]["pass"] += 1
                                                        test_summary["Ring"]["pass"] += 1
                                                    elif "ringing=no" in text_line:
                                                        test_result["Ring"]["Total"] += 1
                                                        test_result["Ring"]["Items"][cur_round_stage][f"{aulaw}_{ohm_type}"]["fail"] += 1
                                                        test_result["Ring"]["Items"][cur_round_stage][f"{aulaw}_{ohm_type}"]["fail_item"].append(
                                                            time_stamp + ' ' + text_line.strip() + ' ' + TS_text)
                                                        test_summary["Ring"]["fail"] += 1
                                                    time_stamp, text_line = read_line_and_split_time(f_read)
                            # --- ring-test ---
                            # --- Voice BERT ---
                            elif "********bert*********" in text_line:
                                if "Voice Bert" not in test_summary:
                                    add_item_in_test_summary(test_summary, "Voice Bert", "Voice Bert")
                                    test_result["Voice Bert"] = {"Total": 0, "Items": {}}
                                if cur_round_stage not in test_result["Voice Bert"]["Items"]:
                                    test_result["Voice Bert"]["Items"][cur_round_stage] = {"u-law_900-ohm": {"pass": 0, "fail": 0, "fail_item": []}, 
                                                                                "u-law_600-ohm": {"pass": 0, "fail": 0, "fail_item": []},
                                                                                "a-law_900-ohm": {"pass": 0, "fail": 0, "fail_item": []},
                                                                                "a-law_600-ohm": {"pass": 0, "fail": 0, "fail_item": []}}
                                for aulaw in ["u-law","a-law"]:
                                    if aulaw in text_line:
                                        for ohm_type in ["900-ohm", "600-ohm"]:
                                            if ohm_type in text_line:
                                                time_stamp, text_line = read_line_and_split_time(f_read)
                                                while "********bert done*********" not in text_line:
                                                    if "TS" in text_line:
                                                        TS_text = text_line
                                                    #time_stamp, text_line = read_line_and_split_time(f_read)
                                                    if "Fail" in text_line:
                                                        test_result["Voice Bert"]["Total"] += 1
                                                        test_result["Voice Bert"]["Items"][cur_round_stage][f"{aulaw}_{ohm_type}"]["fail"] += 1
                                                        test_result["Voice Bert"]["Items"][cur_round_stage][f"{aulaw}_{ohm_type}"]["fail_item"].append(
                                                            time_stamp + ' ' + text_line.strip() + ' ' + TS_text)
                                                        test_summary["Voice Bert"]["fail"] += 1
                                                    elif "Pass" in text_line:
                                                        test_result["Voice Bert"]["Total"] += 1
                                                        test_result["Voice Bert"]["Items"][cur_round_stage][f"{aulaw}_{ohm_type}"]["pass"] += 1
                                                        test_summary["Voice Bert"]["pass"] += 1
                                                    time_stamp, text_line = read_line_and_split_time(f_read)
                            time_stamp, text_line = read_line_and_split_time(f_read)
                    time_stamp, text_line = read_line_and_split_time(f_read)
                # --- Voice BERT ---
                # time_stamp, text_line = read_line_and_split_time(f_read)
            # print(json.dumps(test_result,indent=4))
        elif "-----Begin----" in text_line and "test_card" in text_line:
            # print(time_stamp,text_line)
            if "Bert" not in test_result:
                test_result["Bert"] = {"Total": 0, "Items": {}}
            action_type = ""
            cfg_type = re.split(r'\$cfg_type\ ', text_line)[1].split(' ')[0].replace('\n','')
            str_cfg_type, str_cfg_type_code = get_cfg_str_and_code(int(cfg_type))
            card_type = re.split(r'test_card: ', text_line)[1].split("_slot")[0].split("_")[1]
            str_card_type, str_card_type_code = get_card_type_str_and_code(card_type)
        elif any((match_action := action) in text_line for action in cisco_actions):
            working_pw_array.clear()
            standby_pw_array.clear()
            rs485_array.clear()
            cur_action = match_action
        elif any((match_stage := stage) in text_line for stage in voice_stages):
            working_pw_array.clear()
            standby_pw_array.clear()
            rs485_array.clear()
            card_action_cfg_str = F'{str_card_type}-{str_cfg_type}-{match_stage}-{cur_action}'
            if card_action_cfg_str not in test_result["Bert"]["Items"]:
                test_result["Bert"]["Items"][card_action_cfg_str] = {
                                "pass": 0, "fail": 0, "fail_item": []}  
            time_stamp, text_line = read_line_and_split_time(f_read)
            while match_stage +" end" not in text_line and text_line != '':
                if "AM3440_CCPA_pw_status_get" in text_line or "AM3440_CCPA_pw_Standbystatus_get_log" in text_line:
                    case = "PW Status"
                    if case not in test_summary:
                        add_item_in_test_summary(test_summary, "PW Status", "PW Status")
                    if case not in test_result:
                        test_result[case] = []
                    pw_data = pw_status_parse(time_stamp, text_line)
                    if "AM3440_CCPA_pw_status_get" in text_line:
                        if pw_data:
                            working_pw_array.append(pw_data)
                    elif "AM3440_CCPA_pw_Standbystatus_get_log" in text_line:
                        if pw_data:
                            standby_pw_array.append(pw_data)
                elif "AM3440_CCPA_Get_Rs485_statist" in text_line:
                    case = "RS485 Status"
                    if case not in test_summary:
                        add_item_in_test_summary(test_summary, "RS485 Status", "RS485 Status")
                    if case not in test_result:
                        test_result[case] = []
                    rs485_data = rs485_dut_info_parse(
                        time_stamp, text_line)
                    if rs485_data:
                        rs485_array.append(rs485_data)
                elif "AM3440_CCPA_rs485_get_statist_by_slot" in text_line:
                    case = "RS485 Status"
                    if case not in test_summary:
                        add_item_in_test_summary(test_summary, "RS485 Status", "RS485 Status")
                    if case not in test_result:
                        test_result[case] = []
                    rs485_data = rs485_slot_info_parse(
                        time_stamp, text_line, rs485_all_dut_slot_info)
                    if rs485_data:
                        rs485_array.append(rs485_data)
                elif "Bert pass" in text_line or "bert test Pass" in text_line:                       
                    if match_stage == 'record after action' or match_stage == 'Check setup status':                    
                        test_result["Bert"]["Total"] += 1
                        test_result["Bert"]["Items"][card_action_cfg_str]["pass"] += 1
                        if "coldreset" in card_action_cfg_str:
                            if "T1" in card_action_cfg_str and "Unframe" in card_action_cfg_str:
                                test_summary["coldreset(T1 Unframe)"]["pass"] += 1
                            else:
                                test_summary["coldreset"]["pass"] += 1
                        else:
                            if cur_action != 'Initial' and cur_action != 'CLK Init' and cur_action != '':
                                test_summary[cur_action]["pass"] += 1
                        if "switching time is " in text_line:
                            recover_time = text_line.split()[-2] + ' sec'
                            # print(recover_time)
                            back_time_data = {time_stamp: "Switching time is " + recover_time}
                            test_result["Bert"]["Items"][card_action_cfg_str]["fail_item"].append(
                                        back_time_data)
                elif re.search("[\s]*Fireberd[\s\d]+result get", text_line):
                    # print(action_type)
                    time_stamp, text_line = read_line_and_split_time(
                        f_read)
                    fireberd_info = ""
                    while "Fireberd check PASS." not in text_line and \
                        "Fireberd get Error." not in text_line and \
                            "Fireberd get SLIP." not in text_line and text_line != "":
                        fireberd_info += re.sub(
                            r'Fireberd RESULT[\s|\:]*', "", text_line).strip() + '\n'
                        # fireberd_info += text_line.replace("Fireberd RESULT : ","").replace("Fireberd RESULT ","").replace("        ", " ").replace("     ","")
                        time_stamp, text_line = read_line_and_split_time(
                            f_read)
                    # print(fireberd_info)
                    if "Fireberd get Error." in text_line or \
                            "Fireberd get SLIP." in text_line:
                        test_result["Bert"]["Total"] += 1
                        test_result["Bert"]["Items"][card_action_cfg_str]["fail_item"].append(
                            {"Fireberd error": time_stamp+'\n'+fireberd_info})
                        test_result["Bert"]["Items"][card_action_cfg_str]["fail"] += 1
                        if "coldreset" in card_action_cfg_str:
                            if "T1" in card_action_cfg_str and "Unframe" in card_action_cfg_str:
                                test_summary["coldreset(T1 Unframe)"]["fail"] += 1
                            else:
                                test_summary["coldreset(others)"]["fail"] += 1
                        else:
                            if cur_action != 'Initial':
                                test_summary[cur_action]["fail"] += 1
                    elif "Fireberd check PASS." in text_line:
                        test_result["Bert"]["Total"] += 1
                        test_result["Bert"]["Items"][card_action_cfg_str]["pass"] += 1
                        if "coldreset" in card_action_cfg_str:
                            if "T1" in card_action_cfg_str and "Unframe" in card_action_cfg_str:
                                test_summary["coldreset(T1 Unframe)"]["pass"] += 1
                            else:
                                test_summary["coldreset(others)"]["pass"] += 1
                        else:
                            if cur_action != 'Initial':
                                test_summary[cur_action]["pass"] += 1

                time_stamp, text_line = read_line_and_split_time(f_read)
            if len(working_pw_array) > 0 or len(standby_pw_array) > 0:
                if card_action_cfg_str not in test_result["PW Status"]:
                    test_result["PW Status"].append({
                        card_action_cfg_str: {
                            "Working PW status": [],
                            "Standby PW status": []
                        }
                    })
                if len(working_pw_array) > 0:
                    test_result["PW Status"][-1][card_action_cfg_str]["Working PW status"] = working_pw_array.copy()
                if len(standby_pw_array) > 0:
                    test_result["PW Status"][-1][card_action_cfg_str]["Standby PW status"] = standby_pw_array.copy()
                count_pw_test_summary(test_summary, working_pw_array,standby_pw_array)
            if len(rs485_array) > 0:
                test_result["RS485 Status"].append({
                    card_action_cfg_str: rs485_array.copy()
                })
                found = False
                for rs485_status in rs485_array:
                    if rs485_status['result'] == 'fail':
                        test_summary["RS485 Status"]['fail'] += 1
                        found = True
                        break
                if found != True:
                    test_summary["RS485 Status"]['pass'] += 1
def main(*argv):
    test_result = {}
    # cfg_parse("./Error_logFW0428.txt",test_result)
    # pdh_parse("../error_logs/mix/Error_logmix0504.txt",test_result)
    mib_diff_file_parse(['../error_logs/MIBdiff/log1/mib1.txt',
                        '../error_logs/MIBdiff/log1/mib2.txt'], test_result)
    # mib_parse("../error_logs/MIBscan/InMibOnly_2.txt",
    #   "../error_logs/MIBscan/InProductOnly_2.txt", test_result)
    # parser = ArgumentParser()
    # parser.add_argument("input_error_log", help="error log want to parse")
    # argv = parser.parse_args()
    # file_to_parse = argv.input_error_log
    # date = re.findall(r'\d+', argv.input_error_log)[0]
    # file_to_export = F"./Report_{date}.txt"
    # log_parse(file_to_parse, file_to_export)

    # print(json.dumps(signal_dict,indent=4))


if __name__ == "__main__":
    main(sys.argv[1:])
