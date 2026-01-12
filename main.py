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
    
    # history_payment_file = "D:\\4_项目\\稿酬计算程序\\数据0305\\已支付案例稿酬情况（2015-2023）2025.03.05更新.xlsx"
    # readHistoryRealPaymentData(history_payment_file)
    # print("历史实付数据已成功写入数据库。")

if __name__ == '__main__':
    main()