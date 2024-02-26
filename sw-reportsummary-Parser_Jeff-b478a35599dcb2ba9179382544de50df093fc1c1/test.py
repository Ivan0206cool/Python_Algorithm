

text_line='DUT=1 CTRL=CTRL1 SW=V23.12.01(0107) Serial=205 HW=Ver.C  FPGA1=V16(15)  FPGA2=V14(10)  CPLD=V3(2)'
print(text_line.split('HW=')[0].split(' ')[3].split('=')[1])
print(text_line.split('HW=')[0])