from src.backend.read_case import *
from src.db.init_db import init_db
from src.frontend.app import app
from PyQt5.QtWidgets import QInputDialog

from sqlalchemy.orm import sessionmaker
from sqlalchemy import create_engine, MetaData
from sqlalchemy.ext.declarative import declarative_base



def main():
    init_db()
    app()
    
    # engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=False)
    # maker = sessionmaker(bind=engine)
    # session = maker()
    # # HuaTuData.__table__.drop(engine)

    # # init_db()

    # # 更新 payment_calculated_year 表中的所有行，将 is_calculated 列设置为 False
    # year = session.query(PaymentCalculatedYear).update({PaymentCalculatedYear.is_calculated: False})

    # session.commit()
    # session.close()
    
    
    # from sqlalchemy import create_engine
    # from sqlalchemy.orm import sessionmaker
    # from src.db.init_db import Base, Case, BrowsingRecord, DownloadRecord, HuaTuData, Payment
    # import json

    # # 初始化数据库连接
    # engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=False)
    # Session = sessionmaker(bind=engine)
    # session = Session()

    # # 1. 获取所有 owner_name 为 '浙江大学管理学院' 的 Case 实例
    # cases_to_delete = session.query(Case).filter(Case.owner_name == '浙江大学管理学院').all()

    # aliases = []
    # for case in cases_to_delete:
    #     if case.alias:
    #         try:
    #             alias_list = json.loads(case.alias)
    #             if isinstance(alias_list, list):
    #                 aliases.extend(alias_list)
    #             else:
    #                 aliases.append(case.alias)
    #         except json.JSONDecodeError:
    #             aliases.append(case.alias)

    # # 3. 删除关联表中的记录
    # if aliases:
    #     session.query(BrowsingRecord).filter(BrowsingRecord.case_name.in_(aliases)).delete(synchronize_session=False)
    #     session.query(DownloadRecord).filter(DownloadRecord.case_name.in_(aliases)).delete(synchronize_session=False)
    #     session.query(HuaTuData).filter(HuaTuData.case_name.in_(aliases)).delete(synchronize_session=False)
    #     session.query(Payment).filter(Payment.case_name.in_(aliases)).delete(synchronize_session=False)

    # # 4. 删除对应的 Case 记录（会自动级联删除，因为定义了 `ondelete='CASCADE'`）
    # for case in cases_to_delete:
    #     session.delete(case)

    # # 提交更改
    # session.commit()

    # print(f"已删除 {len(aliases)} 个 alias 对应的记录和 {len(cases_to_delete)} 个 Case。")


    # from sqlalchemy import create_engine
    # from sqlalchemy.orm import sessionmaker
    # from src.db.init_db import BrowsingRecord

    # # 初始化数据库连接
    # engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=False)
    # Session = sessionmaker(bind=engine)
    # session = Session()

    # # 要删除的邮箱用户列表
    # emails_to_delete = [
    #     'luoqiong@htxt.com.cn',
    #     'wangxuewei@htxt.com.cn',
    #     'lichw@sem.tsinghua.edu.cn'
    # ]

    # # 删除 BrowsingRecord 中指定用户的记录
    # session.query(BrowsingRecord).filter(BrowsingRecord.browser.in_(emails_to_delete)).delete(synchronize_session=False)

    # # 删除 DownloadRecord 中指定用户的记录
    # session.query(DownloadRecord).filter(DownloadRecord.downloader.in_(emails_to_delete)).delete(synchronize_session=False)

    # # 提交更改
    # session.commit()
    

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