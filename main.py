from src.backend.read_case import *
from src.db.init_db import init_db
from src.frontend.main import app

from sqlalchemy.orm import sessionmaker
from sqlalchemy import create_engine


def main():
    init_db()
    app()
    # readCaseList('data/test.xls')
    # # readCaseList('data/testCaseList.xlsx')
    # readBrowsingAndDownloadRecord_Tsinghua('data/testTsinghua.xls')
    # # readBrowsingAndDownloadRecord_Tsinghua('data/testTsinghua.xlsx')
    # readBrowsingAndDownloadData_HuaTu('data/testHuaTu.xls', 2024)
    # # readBrowsingAndDownloadData_HuaTu('data/testHuaTu.xls', 2023)
    # # readBrowsingAndDownloadData_HuaTu('data/testHuaTu.xls', 2022)
    # # exportCaseList('data/testCL.xlsx')
    # # exportBrowsingAndDownloadRecord('data/testTH.xlsx')
    # # exportHuaTuData('data/testHT.xlsx')
    # # cases = getSimilarCases('“双碳”政策下，宏宝莱的转型发展之路')
    # # for case in cases:
    # #     print(case)

if __name__ == '__main__':
    main()