from src.backend.read_case import *
from src.db.init_db import init_db

def main():
    init_db()
    readCaseList('data/test.xls')
    readCaseExclusiveAndBatch('data/人大案例第一批.xlsx', '中国人民大学', 1)
    readCaseExclusiveAndBatch('data/人大案例第二批.xlsx', '中国人民大学', 2)
    m, n = readCaseExclusiveAndBatch('data/浙大案例第一批.xlsx', '浙江大学', 1)
    p, q = readCaseExclusiveAndBatch('data/浙大案例第二批.xlsx', '浙江大学', 2)
    for case in n + q:
        print(case)
    # readCaseList('data/testCaseList.xlsx')
    # readBrowsingAndDownloadRecord_Tsinghua('data/testTsinghua.xls')
    # readBrowsingAndDownloadRecord_Tsinghua('data/testTsinghua.xlsx')
    # readBrowsingAndDownloadData_HuaTu('data/testHuaTu.xls', 2024)
    # readBrowsingAndDownloadData_HuaTu('data/testHuaTu.xls', 2023)
    # readBrowsingAndDownloadData_HuaTu('data/testHuaTu.xls', 2022)
    # exportCaseList('data/testCL.xlsx')
    # exportBrowsingAndDownloadRecord('data/testTH.xlsx')
    # exportHuaTuData('data/testHT.xlsx')
    # cases = getSimilarCases('“双碳”政策下，宏宝莱的转型发展之路')
    # for case in cases:
    #     print(case)

if __name__ == '__main__':
    main()