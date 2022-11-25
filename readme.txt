使用说明：

1. 把原始数据转换为input.txt。这个过程去除了冗余数据
    具体步骤：
    a. 把pymain.py的第111行中的MODE置1
    b. 把input.xlsx替换为你的新数据文件，只需要有前四列即可(data, case1, case2, case3)。名字需要也叫input.xlsx
    c. 在pymain.py所在的文件夹的地址栏输入cmd回车，再输入python pymain.py所在的文件夹的地址栏输入cmd回车,运行结束即获得input.txt

2. 读取上一步生成的input.txt, 并开始用C++程序计算结果，结果保存在output.txt，该文件中存放了abc和最小target 
    具体步骤：
    a. 双击vs工程*.sln，打来它后，按F5，运行,等待运行结束，即得到output.txt

3. 把上一步生成的output.txt转化为output.xlsx
    具体步骤：
    a. 把pymain.py的第111行中的MODE置2
    b. 在pymain.py所在的文件夹的地址栏输入cmd回车，再输入python pymain.py所在的文件夹的地址栏输入cmd回车,运行结束即获得output.xlsx