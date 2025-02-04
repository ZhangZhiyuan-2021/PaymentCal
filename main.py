from src.backend.read_case import *
from src.db.init_db import init_db

def main():
    init_db()
    readCaseList('data/test.xls')
    # readCaseList('data/testCaseList.xlsx')
    readBrowsingAndDownloadRecord_Tsinghua('data/testTsinghua.xls')
    # readBrowsingAndDownloadRecord_Tsinghua('data/testTsinghua.xlsx')
    # getTest()
    # cases = getSimilarCases('“双碳”政策下，宏宝莱的转型发展之路')
    # for case in cases:
    #     print(case)

if __name__ == '__main__':
    main()