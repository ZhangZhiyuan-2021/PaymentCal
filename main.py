from src.backend.read_case import *
from src.db.init_db import init_db
from src.frontend.app import app

def main():
    init_db()
    app()
    
    # # 清空全部年份
    # engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=False)
    # Session = sessionmaker(bind=engine)
    # session = Session()

    # try:
    #     # 1. 清空 Payment 表
    #     deleted_count = session.query(Payment).delete()
    #     print(f"已删除 Payment 记录数: {deleted_count}")

    #     # 2. 重置 is_calculated 为 False
    #     updated_count = session.query(PaymentCalculatedYear).update(
    #         {PaymentCalculatedYear.is_calculated: False}
    #     )
    #     print(f"已重置年份记录数: {updated_count}")

    #     session.commit()
    #     print("操作完成：Payment 表已清空，is_calculated 已重置为 False")

    # except Exception as e:
    #     session.rollback()
    #     print("操作失败:", e)

    # finally:
    #     session.close()
    
    
    # # 清空指定年份
    # engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=False)
    # Session = sessionmaker(bind=engine)
    # session = Session()

    # try:
    #     # 1️⃣ 删除 某些 年的 payment
    #     deleted_count = session.query(Payment).filter(
    #         Payment.year.in_([2025])
    #     ).delete(synchronize_session=False)

    #     print(f"删除 Payment 记录: {deleted_count}")

    #     # 2️⃣ 重置年份状态
    #     updated_count = session.query(PaymentCalculatedYear).filter(
    #         PaymentCalculatedYear.year.in_([2025])
    #     ).update(
    #         {PaymentCalculatedYear.is_calculated: False},
    #         synchronize_session=False
    #     )

    #     print(f"重置年份记录: {updated_count}")

    #     session.commit()
    #     print("完成：部分年份 的 payment 已删除，is_calculated 已重置")

    # except Exception as e:
    #     session.rollback()
    #     print("操作失败:", e)

    # finally:
    #     session.close()
    

    
    # history_payment_file = "D:\\4_projects\\稿酬计算程序\\数据0305\\已支付案例稿酬情况（2015-2025）2025.03.17更新.xlsx"
    # readHistoryRealPaymentData(history_payment_file)
    # print("历史实付数据已成功写入数据库。")

if __name__ == '__main__':
    main()