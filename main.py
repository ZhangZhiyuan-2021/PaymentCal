from src.backend.read_case import *
from src.db.init_db import init_db
from src.frontend.main import app
from PyQt5.QtWidgets import QInputDialog

from sqlalchemy.orm import sessionmaker
from sqlalchemy import create_engine, MetaData
from sqlalchemy.ext.declarative import declarative_base



def main():
    init_db()
    app()
    
    
    
    # # # 初始化数据库连接
    # engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=False)
    # maker = sessionmaker(bind=engine)
    # session = maker()
    # # PaymentCalculatedYear.__table__.drop(engine)

    # # 更新 payment_calculated_year 表中的所有行，将 is_calculated 列设置为 False
    # session.query(PaymentCalculatedYear).filter_by(year=2023).update({PaymentCalculatedYear.is_calculated: False})

    # # 提交更改
    # session.commit()

    # # 关闭会话
    # session.close()

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