from src.backend.read_case import *
from src.db.init_db import init_db

def main():
    # init_db()
    # readCaseList('./test.xls')
    # readCaseList('./testCaseList.xlsx')
    # readBrowsingAndDownloadRecord_Tsinghua('./testTsinghua.xls')
    # readBrowsingAndDownloadRecord_Tsinghua('./testTsinghua.xlsx')
    # getTest()
    cases = getSimilarCases('“双碳”政策下，宏宝莱的转型发展之路')
    for case in cases:
        print(case)

if __name__ == '__main__':
    main()