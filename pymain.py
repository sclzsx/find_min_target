from email import header
from operator import index
import numpy as np
import pandas as pd
import docx
import sys
import math


def excel_to_txt(input_excel_path, output_txt_path):
    DATA = np.array(pd.read_excel(input_excel_path, sheet_name='Sheet1'))
    data  = DATA[:,0].astype('float32')
    case1 = DATA[:,1].astype('uint8')
    case2 = DATA[:,2].astype('uint8')
    case3 = DATA[:,3].astype('uint8')

    lines = []
    for i in range(DATA.shape[0]):
        line = ('%.3f %d %d %d') % (data[i], case1[i], case2[i], case3[i])
        if line not in lines:
            lines.append(line)
            print(line)

    with open(output_txt_path, 'w') as f:
        for line in lines:
            f.write(line)
            f.write('\n')


def txt_to_excel(input_txt_path, output_excel_path):
    DATA = []
    header = ['data', 'case1', 'case2', 'case3', 'a', 'b', 'c', 'target']
    DATA.append(header)
    with open(input_txt_path, 'r') as f:
        lines = f.readlines()
        for line in lines:
            info = line[:-1].split(' ')
            if len(info) != len(header):
                continue
            data, case1, case2, case3, a, b, c, target = info[0], info[1], info[2], info[3], info[4], info[5], info[6], info[7]
            data_list = [data, case1, case2, case3, a, b, c, target]
            print(data_list)
            DATA.append(data_list)
    DATA = pd.DataFrame(DATA)
    writer = pd.ExcelWriter(output_excel_path)
    DATA.to_excel(writer,'Sheet1',float_format='%.3f', header=None, index=False)
    writer.save()


def do_search(excel_path, max_range):
    DATA = np.array(pd.read_excel(excel_path, sheet_name='Sheet1'))

    data  = DATA[:,0].astype('float32')
    case1 = DATA[:,1].astype('uint8')
    case2 = DATA[:,2].astype('uint8')
    case3 = DATA[:,3].astype('uint8')

    L = data.shape[0]

    for i in range(1):

        ABCT = []

        MIN = 10000

        # find = 0

        for a in range(max_range):
            for b in range(max_range):
                for c in range(max_range):

                    sum = a + b + c
                    ma = max(case1[i], case2[i], case3[i])
                    mi = min(case1[i], case2[i], case3[i])

                    down = np.floor(data[i] / ma)
                    up = np.ceil(data[i] / mi)

                    # print('case1[i]:%d, case2[i]:%d, case3[i]:%d, data[i]:%f,a:%d,b:%d,c:%d,sum:%d, ma:%d,mi:%d, down:%f, up:%f' % (case1[i], case2[i], case3[i], data[i],a,b,c,sum, ma,mi,down,up))

                    if not (sum >= down and sum <= up):
                        continue

                    target = case1[i] * a + case2[i] * b + case3[i] * c - data[i]

                    # print(target)

                    if target < 0:
                        continue
                    
                    # if find > 10:
                    #     break

                    if target <= MIN:
                        l = [a, b, c, target]
                        # find += 1
                        # print(l)
                        ABCT.append(l)

                        MIN = target

                    # CNT -= 1

        for l in ABCT:
            if l[-1] == MIN:
                print(l)


if __name__ == '__main__':

    MODE = 2

    if MODE == 1:
        excel_to_txt('./input.xlsx', './input.txt')
    elif MODE == 2:
        txt_to_excel('./output.txt', './output.xlsx')
    else:
        do_search('./input.xlsx', 1000)