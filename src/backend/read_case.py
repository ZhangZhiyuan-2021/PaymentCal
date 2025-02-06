import json
import pandas as pd
from sqlalchemy.orm import sessionmaker
from sqlalchemy import create_engine
import datetime
from fuzzywuzzy import process
import sys
import os

# 获取当前脚本的目录，向上找到 `src` 目录
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..")))
from src.db.init_db import CopyrightOwner, Case, BrowsingRecord, DownloadRecord, HuaTuData, Payment

# 读案例列表
def readCaseList(path):
    # 读取 Excel 文件，过滤“已入库”的记录
    df = pd.read_excel(path)
    data_dict_list = [r for r in df.to_dict(orient='records') if r.get('案例状态') == '已入库']
    wrong_cases = []

    engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=True)
    Session = sessionmaker(bind=engine)
    session = Session()

    # -------------------------------
    # 【优化】直接加载所有案例和版权方数据
    all_cases = session.query(Case).all()
    # 构建两个缓存：按案例标题和投稿编号索引
    cases_by_name = {case.name: case for case in all_cases if case.name}
    cases_by_submission = {case.submission_number: case for case in all_cases if case.submission_number}

    all_owners = session.query(CopyrightOwner).all()
    owner_cache = {owner.name: owner for owner in all_owners if owner.name}

    # -------------------------------
    # 遍历 Excel 中的每条记录
    for data_dict in data_dict_list:
        # 检查必填字段
        if (not data_dict.get('案例标题') or pd.isna(data_dict.get('案例标题')) or
            not data_dict.get('投稿编号') or pd.isna(data_dict.get('投稿编号')) or
            not data_dict.get('案例版权') or pd.isna(data_dict.get('案例版权')) or
            not data_dict.get('发布时间') or pd.isna(data_dict.get('发布时间')) or
            not data_dict.get('创建时间') or pd.isna(data_dict.get('创建时间'))):
            print('案例标题、投稿编号、案例版权、发布时间、创建时间不能为空')
            data_dict['错误信息'] = '案例标题、投稿编号、案例版权、发布时间、创建时间不能为空'
            wrong_cases.append(data_dict)
            continue

        owner_name = data_dict['案例版权']
        # 直接从缓存中获取版权方，如不存在则新建
        owner = owner_cache.get(owner_name)
        if not owner:
            owner = CopyrightOwner(name=owner_name)
            session.add(owner)
            owner_cache[owner_name] = owner

        case_title = data_dict['案例标题']
        submission_number = data_dict['投稿编号']

        # 先按案例标题查找，如不存在再按投稿编号查找
        case = cases_by_name.get(case_title)
        if not case:
            case = cases_by_submission.get(submission_number)
            if (case and 
                case.owner_name == owner.name and 
                case.release_time == datetime.datetime.strptime(data_dict['发布时间'], "%Y-%m-%d %H:%M:%S.%f") and
                case.create_time == datetime.datetime.strptime(data_dict['创建时间'], "%Y-%m-%d %H:%M:%S.%f")):
                cases_by_name.pop(case.name, None)
                # 已存在的案例：更新信息
                # 更新 alias（别名）记录
                if not case.alias:
                    case.alias = json.dumps([case_title])
                else:
                    try:
                        alias_list = json.loads(case.alias)
                    except Exception:
                        alias_list = []
                    if case_title not in alias_list:
                        alias_list.append(case_title)
                        case.alias = json.dumps(alias_list)
                # 更新其他字段
                case.name = case_title
                case.type = data_dict.get('产品类型')
                try:
                    case.release_time = datetime.datetime.strptime(data_dict['发布时间'], "%Y-%m-%d %H:%M:%S.%f")
                    case.create_time = datetime.datetime.strptime(data_dict['创建时间'], "%Y-%m-%d %H:%M:%S.%f")
                except Exception:
                    try:
                        case.release_time = datetime.datetime.strptime(data_dict['发布时间'], "%Y-%m-%d %H:%M:%S")
                        case.create_time = datetime.datetime.strptime(data_dict['创建时间'], "%Y-%m-%d %H:%M:%S")
                    except Exception as e:
                        print("发布时间解析错误", data_dict['发布时间'])
                        data_dict['错误信息'] = '发布时间或创建时间格式不正确'
                        wrong_cases.append(data_dict)
                        continue
                case.is_micro = True if data_dict.get('是否微案例') == '是' else False
                case.is_exclusive = True if data_dict.get('是否独家案例') == '是' else False
                try:
                    case.batch = int(data_dict.get('案例批次', 0))
                except:
                    case.batch = 0
                case.submission_source = data_dict.get('投稿来源')
                case.contain_TN = True if data_dict.get('是否含有教学说明') == '是' else False
                case.is_adapted_from_text = True if data_dict.get('是否由文字案例改编') == '是' else False
                case.owner_name = owner.name
                cases_by_name[case_title] = case
            else:
                # 新建案例
                try:
                    release_time = datetime.datetime.strptime(data_dict['发布时间'], "%Y-%m-%d %H:%M:%S.%f")
                    create_time = datetime.datetime.strptime(data_dict['创建时间'], "%Y-%m-%d %H:%M:%S.%f")
                except Exception:
                    try:
                        release_time = datetime.datetime.strptime(data_dict['发布时间'], "%Y-%m-%d %H:%M:%S")
                        create_time = datetime.datetime.strptime(data_dict['创建时间'], "%Y-%m-%d %H:%M:%S")
                    except Exception as e:
                        print("发布时间解析错误", data_dict['发布时间'])
                        data_dict['错误信息'] = '发布时间或创建时间格式不正确'
                        wrong_cases.append(data_dict)
                        continue
                case = Case(
                    name=case_title,
                    type=data_dict.get('产品类型'),
                    submission_number=submission_number,
                    release_time=release_time,
                    create_time=create_time,
                    is_micro=True if data_dict.get('是否微案例') == '是' else False,
                    is_exclusive=True if data_dict.get('是否独家案例') == '是' else False,
                    batch=int(data_dict.get('案例批次', 0)),
                    submission_source=data_dict.get('投稿来源'),
                    contain_TN=True if data_dict.get('是否含有教学说明') == '是' else False,
                    is_adapted_from_text=True if data_dict.get('是否由文字案例改编') == '是' else False,
                    owner_name=owner.name
                )
                session.add(case)
                # 同时更新缓存
                cases_by_name[case_title] = case
                cases_by_submission[submission_number] = case

    session.commit()
    session.close()
    return wrong_cases

# def readCaseList(path):
#     # 读取 xls 或 xlsx 文件
#     df = pd.read_excel(path)
#     data_dict_list = [x for x in df.to_dict(orient='records') if x['案例状态'] == '已入库']

#     wrong_cases = []

#     engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=True)
#     Session = sessionmaker(bind=engine)
#     session = Session()

#     for data_dict in data_dict_list:
#         if (not bool(data_dict['案例标题']) or pd.isna(data_dict['案例标题'])
#             or not bool(data_dict['投稿编号']) or pd.isna(data_dict['投稿编号'])
#             or not bool(data_dict['案例版权']) or pd.isna(data_dict['案例版权'])
#             or not bool(data_dict['发布时间']) or pd.isna(data_dict['发布时间'])):
#             print('案例标题、投稿编号、案例版权、发布时间不能为空')
#             wrong_cases.append(data_dict)
#             continue

#         owner = session.query(CopyrightOwner).filter_by(name=data_dict['案例版权']).first()
#         if not owner:
#             owner = CopyrightOwner(name=data_dict['案例版权'])
#             session.add(owner)
#             session.commit()

#         case = session.query(Case).filter_by(name=data_dict['案例标题']).first()
#         owner = session.query(CopyrightOwner).filter_by(name=data_dict['案例版权']).first()
#         if not case:
#             case = session.query(Case).filter_by(submission_number=data_dict['投稿编号']).first()
#             if case:
#                 if not case.alias:
#                     case.alias = json.dumps([data_dict['案例标题']])
#                 else:
#                     alias_list = json.loads(case.alias)
#                     alias_list.append(data_dict['案例标题'])
#                     case.alias = json.dumps(alias_list)
#                 case.name = data_dict['案例标题']
#                 case.type = data_dict['产品类型']
#                 case.release_time = datetime.datetime.strptime(data_dict['发布时间'], "%Y-%m-%d %H:%M:%S.%f")
#                 case.is_micro = True if data_dict['是否微案例'] == '是' else False
#                 case.is_exclusive = True if data_dict['是否独家案例'] == '是' else False
#                 case.batch = int(data_dict['案例批次'])
#                 case.submission_source = data_dict['投稿来源']
#                 case.contain_TN = True if data_dict['是否含有教学说明'] == '是' else False
#                 case.is_adapted_from_text = True if data_dict['是否由文字案例改编'] == '是' else False
#                 case.owner_name = owner.name
#                 session.commit()
#             else:
#                 case = Case(name=data_dict['案例标题'], 
#                             type=data_dict['产品类型'], 
#                             submission_number=data_dict['投稿编号'],
#                             release_time=datetime.datetime.strptime(data_dict['发布时间'], "%Y-%m-%d %H:%M:%S.%f"), 
#                             is_micro=True if data_dict['是否微案例'] == '是' else False,
#                             is_exclusive=True if data_dict['是否独家案例'] == '是' else False,
#                             batch=int(data_dict['案例批次']),
#                             submission_source=data_dict['投稿来源'],
#                             contain_TN=True if data_dict['是否含有教学说明'] == '是' else False,
#                             is_adapted_from_text=True if data_dict['是否由文字案例改编'] == '是' else False,
#                             owner_name=owner.name)
#                 session.add(case)
#                 session.commit()

#     session.close()

#     return wrong_cases

def getCopyrightOwner(name):
    engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=True)
    Session = sessionmaker(bind=engine)
    session = Session()

    owner = session.query(CopyrightOwner).filter_by(name=name).first()
    if not owner:
        print('版权方不存在')
        return None

    session.close()

    return owner

def getAllCopyrightOwners():
    engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=True)
    Session = sessionmaker(bind=engine)
    session = Session()

    owners = session.query(CopyrightOwner).all()

    session.close()

    return owners

def updateCopyrightOwner(name, new_name):
    engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=True)
    Session = sessionmaker(bind=engine)
    session = Session()

    owner = session.query(CopyrightOwner).filter_by(name=name).first()
    if not owner:
        print('版权方不存在')
        return None

    owner.name = new_name

    session.commit()
    session.close()

    return owner

def deleteCopyrightOwner(name):
    engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=True)
    Session = sessionmaker(bind=engine)
    session = Session()

    owner = session.query(CopyrightOwner).filter_by(name=name).first()
    if not owner:
        print('版权方不存在')
        return None

    session.delete(owner)
    session.commit()
    session.close()

    return owner

def deleteAllCopyrightOwners():
    engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=True)
    Session = sessionmaker(bind=engine)
    session = Session()

    for owner in session.query(Case).all():
        session.delete(owner)
    session.commit()
    session.close()

# 给定案例名，返回案例所有属性
def getCase(name):
    engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=True)
    Session = sessionmaker(bind=engine)
    session = Session()

    case = session.query(Case).filter_by(name=name).first()
    if not case:
        print('案例不存在', name)
        return None

    session.close()

    return case

def getSimilarCases(name):
    engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=True)
    Session = sessionmaker(bind=engine)
    session = Session()

    # 通过 fuzzywuzzy 模糊匹配案例名称
    # 直接搜索 Case.name 会给出一个长度为 1 的名称元组的列表
    cases = session.query(Case).all()
    case_names = [case.name for case in cases]
    similar_case_names = process.extract(name, case_names, limit=10)
    similar_cases = []
    for similar_case_name in similar_case_names:
        similar_case = session.query(Case).filter_by(name=similar_case_name[0]).first()
        similar_cases.append(similar_case)

    session.close()

    return similar_cases

def getAllCases():
    engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=True)
    Session = sessionmaker(bind=engine)
    session = Session()

    cases = session.query(Case).all()

    session.close()

    return cases

# 设置alias，则添加别名；其他则是更新属性
def updateCase(name, alias=None, submission_number=None, type=None, release_time=None, create_time=None, is_micro=None, is_exclusive=None, batch=None, submission_source=None, contain_TN=None, is_adapted_from_text=None, owner_name=None):
    engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=True)
    Session = sessionmaker(bind=engine)
    session = Session()

    case = session.query(Case).filter_by(name=name).first()
    if not case:
        print('案例不存在', name)
        return None

    if alias:
        if not case.alias:
            case.alias = json.dumps([alias], ensure_ascii=False)
        else:
            alias_list = json.loads(case.alias)
            alias_list.append(alias)
            case.alias = json.dumps(alias_list, ensure_ascii=False)
    if type:
        case.type = type
    if submission_number:
        case.submission_number = submission_number
    if release_time:
        case.release_time = datetime.datetime.strptime(release_time, "%Y-%m-%d %H:%M:%S.%f")
    if create_time:
        case.create_time = datetime.datetime.strptime(create_time, "%Y-%m-%d %H:%M:%S.%f")
    if is_micro:
        case.is_micro = True if is_micro == '是' else False
    if is_exclusive:
        case.is_exclusive = True if is_exclusive == '是' else False
    if batch:
        case.batch = int(batch)
    if submission_source:
        case.submission_source = submission_source
    if contain_TN:
        case.contain_TN = True if contain_TN == '是' else False
    if is_adapted_from_text:
        case.is_adapted_from_text = True if is_adapted_from_text == '是' else False
    if owner_name:
        owner = session.query(CopyrightOwner).filter_by(name=owner_name).first()
        if not owner:
            print('版权方不存在')
            return None
        case.owner_name = owner.name

    session.commit()
    session.close()

    return case

def deleteCase(name):
    engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=True)
    Session = sessionmaker(bind=engine)
    session = Session()

    case = session.query(Case).filter_by(name=name).first()
    if not case:
        print('案例不存在', name)
        return None

    session.delete(case)
    session.commit()
    session.close()

    return case

def deleteAllCases():
    engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=True)
    Session = sessionmaker(bind=engine)
    session = Session()

    for case in session.query(Case).all():
        session.delete(case)
    session.commit()
    session.close()

# 导出xlsx
def exportCaseList(path):
    engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=True)
    Session = sessionmaker(bind=engine)
    session = Session()

    cases = session.query(Case).all()
    case_list = []
    for case in cases:
        case_dict = {}
        case_dict['案例标题'] = case.name
        case_dict['产品类型'] = case.type
        case_dict['投稿编号'] = case.submission_number
        case_dict['发布时间'] = case.release_time.strftime("%Y-%m-%d %H:%M:%S.%f")
        case_dict['创建时间'] = case.create_time.strftime("%Y-%m-%d %H:%M:%S.%f")
        case_dict['是否微案例'] = '是' if case.is_micro else '否'
        case_dict['是否独家案例'] = '是' if case.is_exclusive else '否'
        case_dict['案例批次'] = case.batch
        case_dict['投稿来源'] = case.submission_source
        case_dict['是否含有教学说明'] = '是' if case.contain_TN else '否'
        case_dict['是否由文字案例改编'] = '是' if case.is_adapted_from_text else '否'
        case_dict['案例版权'] = case.owner.name
        case_list.append(case_dict)

    df = pd.DataFrame(case_list)
    df.to_excel(path, index=False)

    session.close()

def readBrowsingAndDownloadRecord_Tsinghua(path):
    # 解析 Excel 文件及相关工作表
    xls = pd.ExcelFile(path)
    browsingSheets = [s for s in xls.sheet_names if '浏览记录' in s]
    downloadSheets = [s for s in xls.sheet_names if '下载记录' in s]

    # 创建数据库连接与会话
    engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=True)
    Session = sessionmaker(bind=engine)
    session = Session()

    missingInformationBrowsingRecords = []
    missingInformationDownloadRecords = []
    wrongBrowsingRecords = []
    wrongDownloadRecords = []

    # 直接加载所有案例
    all_cases = session.query(Case).all()
    # 构建案例字典：名称 -> Case 对象
    case_dict = {case.name: case for case in all_cases}

    # 收集待插入的记录列表，减少频繁 commit
    new_browsing_records = []
    new_download_records = []

    # --------------------
    # 处理浏览记录
    for sheet in browsingSheets:
        df = xls.parse(sheet)
        for data_dict in df.to_dict(orient='records'):
            # 排除系统账号
            if data_dict.get('浏览人账号') in ['admin', 'anonymous']:
                continue

            # 检查必要字段
            if (not data_dict.get('案例名称') or pd.isna(data_dict.get('案例名称')) or
                not data_dict.get('浏览人账号') or pd.isna(data_dict.get('浏览人账号')) or
                not data_dict.get('浏览时间') or pd.isna(data_dict.get('浏览时间'))):
                print('案例名称、浏览人账号、浏览时间不能为空')
                missingInformationBrowsingRecords.append(data_dict)
                continue

            # 直接从加载的案例字典中查找案例
            case = case_dict.get(data_dict['案例名称'])
            if not case:
                print('案例不存在', data_dict['案例名称'])
                wrongBrowsingRecords.append(data_dict)
                continue

            # 解析浏览时间（仅解析一次）
            try:
                browsing_dt = datetime.datetime.strptime(data_dict['浏览时间'], "%Y-%m-%d %H:%M:%S")
            except Exception as e:
                print(f"日期解析错误: {data_dict['浏览时间']}")
                missingInformationBrowsingRecords.append(data_dict)
                continue

            # 检查是否已存在相同记录
            existing = session.query(BrowsingRecord).filter_by(
                case_name=case.name,
                browser=data_dict['浏览人账号'],
                datetime=browsing_dt
            ).first()
            if existing:
                continue

            record = BrowsingRecord(
                case_name=case.name,
                browser=data_dict['浏览人账号'], 
                browser_institution=data_dict.get('浏览人所在院校'),
                datetime=browsing_dt, 
                is_valid=None
            )
            new_browsing_records.append(record)

    # 批量插入浏览记录，减少数据库交互
    if new_browsing_records:
        session.bulk_save_objects(new_browsing_records)
        session.commit()

    # --------------------
    # 处理下载记录
    for sheet in downloadSheets:
        df = xls.parse(sheet)
        for data_dict in df.to_dict(orient='records'):
            if data_dict.get('下载人账号') in ['admin', 'anonymous']:
                continue

            if (not data_dict.get('案例名称') or pd.isna(data_dict.get('案例名称')) or
                not data_dict.get('下载人账号') or pd.isna(data_dict.get('下载人账号')) or
                not data_dict.get('下载时间') or pd.isna(data_dict.get('下载时间'))):
                print('案例名称、下载人账号、下载时间不能为空')
                missingInformationDownloadRecords.append(data_dict)
                continue

            case = case_dict.get(data_dict['案例名称'])
            if not case:
                print('案例不存在', data_dict['案例名称'])
                wrongDownloadRecords.append(data_dict)
                continue

            try:
                download_dt = datetime.datetime.strptime(data_dict['下载时间'], "%Y-%m-%d %H:%M:%S")
            except Exception as e:
                print(f"日期解析错误: {data_dict['下载时间']}")
                missingInformationDownloadRecords.append(data_dict)
                continue

            existing = session.query(DownloadRecord).filter_by(
                case_name=case.name,
                downloader=data_dict['下载人账号'],
                datetime=download_dt
            ).first()
            if existing:
                continue

            record = DownloadRecord(
                case_name=case.name,
                downloader=data_dict['下载人账号'], 
                downloader_institution=data_dict.get('下载人所在院校'),
                datetime=download_dt, 
                is_valid=None
            )
            new_download_records.append(record)

    if new_download_records:
        session.bulk_save_objects(new_download_records)
        session.commit()

    # --------------------
    # 判断记录有效性（浏览和下载分别处理）
    for case in session.query(Case).all():
        # 浏览记录有效性判断
        browsing_records = session.query(BrowsingRecord).filter_by(case_name=case.name).all()
        if browsing_records:
            browser_group = {}
            for record in browsing_records:
                browser_group.setdefault(record.browser, []).append(record)
            for browser, records in browser_group.items():
                records.sort(key=lambda r: r.datetime)
                records[0].is_valid = True
                last_valid_datetime = records[0].datetime

                for rec in records[1:]:
                    new_semester = False
                    if last_valid_datetime.month in [1]:
                        if (rec.datetime.year == last_valid_datetime.year and rec.datetime.month >= 2) or (rec.datetime.year > last_valid_datetime.year):
                            new_semester = True
                    elif last_valid_datetime.month in [2, 3, 4, 5, 6, 7]:
                        if (rec.datetime.year == last_valid_datetime.year and rec.datetime.month >= 8) or (rec.datetime.year > last_valid_datetime.year):
                            new_semester = True
                    elif last_valid_datetime.month in [8, 9, 10, 11, 12]:
                        if (rec.datetime.year == last_valid_datetime.year + 1 and rec.datetime.month >= 2) or (rec.datetime.year > last_valid_datetime.year + 1):
                            new_semester = True

                    if new_semester and (rec.datetime - last_valid_datetime).days >= 60:
                        rec.is_valid = True
                        last_valid_datetime = rec.datetime
                    else:
                        rec.is_valid = False
            session.commit()

        # 下载记录有效性判断
        download_records = session.query(DownloadRecord).filter_by(case_name=case.name).all()
        if download_records:
            downloader_group = {}
            for record in download_records:
                downloader_group.setdefault(record.downloader, []).append(record)
            for downloader, records in downloader_group.items():
                records.sort(key=lambda r: r.datetime)
                records[0].is_valid = True
                last_valid_datetime = records[0].datetime

                for rec in records[1:]:
                    new_semester = False
                    if last_valid_datetime.month in [1]:
                        if (rec.datetime.year == last_valid_datetime.year and rec.datetime.month >= 2) or (rec.datetime.year > last_valid_datetime.year):
                            new_semester = True
                    elif last_valid_datetime.month in [2, 3, 4, 5, 6, 7]:
                        if (rec.datetime.year == last_valid_datetime.year and rec.datetime.month >= 8) or (rec.datetime.year > last_valid_datetime.year):
                            new_semester = True
                    elif last_valid_datetime.month in [8, 9, 10, 11, 12]:
                        if (rec.datetime.year == last_valid_datetime.year + 1 and rec.datetime.month >= 2) or (rec.datetime.year > last_valid_datetime.year + 1):
                            new_semester = True

                    if new_semester and (rec.datetime - last_valid_datetime).days >= 60:
                        rec.is_valid = True
                        last_valid_datetime = rec.datetime
                    else:
                        rec.is_valid = False
            session.commit()

    session.close()

    return (missingInformationBrowsingRecords, missingInformationDownloadRecords,
            wrongBrowsingRecords, wrongDownloadRecords)

def getBrowsingRecord(case_name, browser=None, browser_institution=None, datetime=None, is_valid=None):
    engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=True)
    Session = sessionmaker(bind=engine)
    session = Session()

    browsing_records = session.query(BrowsingRecord).filter_by(case_name=case_name).all()
    if not browsing_records:
        print('浏览记录不存在')
        return None

    if browser:
        browsing_records = [browsing_record for browsing_record in browsing_records if browsing_record.browser == browser]
    if browser_institution:
        browsing_records = [browsing_record for browsing_record in browsing_records if browsing_record.browser_institution == browser_institution]
    if datetime:
        browsing_records = [browsing_record for browsing_record in browsing_records if browsing_record.datetime == datetime]
    if is_valid:
        browsing_records = [browsing_record for browsing_record in browsing_records if browsing_record.is_valid == is_valid]

    session.close()

    return browsing_records

def getAllBrowsingRecords():
    engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=True)
    Session = sessionmaker(bind=engine)
    session = Session()

    browsing_records = session.query(BrowsingRecord).all()

    session.close()

    return browsing_records

def deleteBrowsingRecord(case_name, browser=None, browser_institution=None, datetime=None, is_valid=None):
    engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=True)
    Session = sessionmaker(bind=engine)
    session = Session()

    browsing_records = session.query(BrowsingRecord).filter_by(case_name=case_name).all()
    if not browsing_records:
        print('浏览记录不存在')
        return None

    if browser:
        browsing_records = [browsing_record for browsing_record in browsing_records if browsing_record.browser == browser]
    if browser_institution:
        browsing_records = [browsing_record for browsing_record in browsing_records if browsing_record.browser_institution == browser_institution]
    if datetime:
        browsing_records = [browsing_record for browsing_record in browsing_records if browsing_record.datetime == datetime]
    if is_valid:
        browsing_records = [browsing_record for browsing_record in browsing_records if browsing_record.is_valid == is_valid]

    for browsing_record in browsing_records:
        session.delete(browsing_record)
    session.commit()
    session.close()

    return browsing_records

def deleteAllBrowsingRecords():
    engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=True)
    Session = sessionmaker(bind=engine)
    session = Session()

    for browsing_record in session.query(BrowsingRecord).all():
        session.delete(browsing_record)
    session.commit()
    session.close()

def getDownloadRecord(case_name, downloader=None, downloader_institution=None, datetime=None, is_valid=None):
    engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=True)
    Session = sessionmaker(bind=engine)
    session = Session()

    download_records = session.query(DownloadRecord).filter_by(case_name=case_name).all()
    if not download_records:
        print('下载记录不存在')
        return None

    if downloader:
        download_records = [download_record for download_record in download_records if download_record.downloader == downloader]
    if downloader_institution:
        download_records = [download_record for download_record in download_records if download_record.downloader_institution == downloader_institution]
    if datetime:
        download_records = [download_record for download_record in download_records if download_record.datetime == datetime]
    if is_valid:
        download_records = [download_record for download_record in download_records if download_record.is_valid == is_valid]

    session.close()

    return download_records

def getAllDownloadRecords():
    engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=True)
    Session = sessionmaker(bind=engine)
    session = Session()

    download_records = session.query(DownloadRecord).all()

    session.close()

    return download_records

def deleteDownloadRecord(case_name, downloader=None, downloader_institution=None, datetime=None, is_valid=None):
    engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=True)
    Session = sessionmaker(bind=engine)
    session = Session()

    download_records = session.query(DownloadRecord).filter_by(case_name=case_name).all()
    if not download_records:
        print('下载记录不存在')
        return None

    if downloader:
        download_records = [download_record for download_record in download_records if download_record.downloader == downloader]
    if downloader_institution:
        download_records = [download_record for download_record in download_records if download_record.downloader_institution == downloader_institution]
    if datetime:
        download_records = [download_record for download_record in download_records if download_record.datetime == datetime]
    if is_valid:
        download_records = [download_record for download_record in download_records if download_record.is_valid == is_valid]

    for download_record in download_records:
        session.delete(download_record)
    session.commit()
    session.close()

    return download_records

def deleteAllDownloadRecords():
    engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=True)
    Session = sessionmaker(bind=engine)
    session = Session()

    for download_record in session.query(DownloadRecord).all():
        session.delete(download_record)
    session.commit()
    session.close()

def exportBrowsingAndDownloadRecord(path):
    engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=True)
    Session = sessionmaker(bind=engine)
    session = Session()

    browsing_records = session.query(BrowsingRecord).all()
    # 每 50000 个记录为一个 sheet
    browsing_record_list = []
    for browsing_record in browsing_records:
        browsing_record_dict = {}
        browsing_record_dict['案例名称'] = browsing_record.case_name
        browsing_record_dict['浏览人账号'] = browsing_record.browser
        browsing_record_dict['浏览人所在院校'] = browsing_record.browser_institution
        browsing_record_dict['浏览时间'] = browsing_record.datetime.strftime("%Y-%m-%d %H:%M:%S")
        browsing_record_list.append(browsing_record_dict)
        
    download_records = session.query(DownloadRecord).all()
    # 每 50000 个记录为一个 sheet
    download_record_list = []
    for download_record in download_records:
        download_record_dict = {}
        download_record_dict['案例名称'] = download_record.case_name
        download_record_dict['下载人账号'] = download_record.downloader
        download_record_dict['下载人所在院校'] = download_record.downloader_institution
        download_record_dict['下载时间'] = download_record.datetime.strftime("%Y-%m-%d %H:%M:%S")
        download_record_list.append(download_record_dict)

    with pd.ExcelWriter(path) as writer:
        for i in range(0, len(browsing_record_list), 50000):
            df = pd.DataFrame(browsing_record_list[i:i+50000])
            df.to_excel(writer, sheet_name='浏览记录' + str(i // 50000 + 1), index=False)
        for i in range(0, len(download_record_list), 50000):
            df = pd.DataFrame(download_record_list[i:i+50000])
            df.to_excel(writer, sheet_name='下载记录' + str(i // 50000 + 1), index=False)

    session.close()

def readBrowsingAndDownloadData_HuaTu(path, year):
    df = pd.read_excel(path)
    data_dict_list = df.to_dict(orient='records')

    engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=True)
    Session = sessionmaker(bind=engine)
    session = Session()

    missingInformationData = []
    wrongData = []

    # -------------------------------
    # 预加载所有案例数据
    all_cases = session.query(Case).all()
    cases_dict = {case.name: case for case in all_cases if case.name}

    # -------------------------------
    # 一次性加载指定年份所有 HuaTuData 数据
    all_huatu = session.query(HuaTuData).filter_by(year=year).all()
    # 构建映射：案例名称 -> HuaTuData 对象
    huatu_dict = {record.case_name: record for record in all_huatu}

    # -------------------------------
    # 遍历 Excel 数据，更新或新增 HuaTuData
    for data_dict in data_dict_list:
        # 检查必填字段
        if (data_dict.get('标题') is None or pd.isna(data_dict.get('标题')) or
            data_dict.get('邮件数') is None or pd.isna(data_dict.get('邮件数')) or
            data_dict.get('查看数') is None or pd.isna(data_dict.get('查看数'))):  
            
            data_dict['错误信息'] = '标题、邮件数、查看数不能为空'
            missingInformationData.append(data_dict)
            continue

        case = cases_dict.get(data_dict['标题'])
        if not case:
            print('案例不存在:', data_dict['标题'])
            wrongData.append(data_dict)
            continue

        # 尝试从 huatu_dict 中获取该案例的 HuaTuData
        huatu_data = huatu_dict.get(case.name)
        if not huatu_data:
            # 不存在，则新建，并加入 huatu_dict 以便后续使用
            huatu_data = HuaTuData(
                case_name=case.name,
                year=year,
                views=int(data_dict['查看数']),
                downloads=int(data_dict['邮件数'])
            )
            session.add(huatu_data)
            huatu_dict[case.name] = huatu_data
        else:
            # 存在则更新数据
            huatu_data.views = int(data_dict['查看数'])
            huatu_data.downloads = int(data_dict['邮件数'])
    
    session.commit()
    session.close()

    return missingInformationData, wrongData

# def readBrowsingAndDownloadData_HuaTu(path, year):
#     df = pd.read_excel(path)
#     data_dict_list = df.to_dict(orient='records')

#     engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=True)
#     Session = sessionmaker(bind=engine)
#     session = Session()

#     missingInformationData = []
#     wrongData = []

#     for data_dict in data_dict_list:
#         if (not bool(data_dict['标题']) or pd.isna(data_dict['标题'])
#             or not bool(data_dict['邮件数']) or pd.isna(data_dict['邮件数'])
#             or not bool(data_dict['查看数']) or pd.isna(data_dict['查看数'])):
#             missingInformationData.append(data_dict)
#             continue

#         case = session.query(Case).filter_by(name=data_dict['标题']).first()

#         if not case:
#             print('案例不存在')
#             wrongData.append(data_dict)
#             continue

#         huatu_data = session.query(HuaTuData).filter_by(case_name=case.name, year=year).first()
#         if not huatu_data:
#             huatu_data = HuaTuData(case_name=case.name, 
#                                    year=year, 
#                                    views=int(data_dict['查看数']), 
#                                    downloads=int(data_dict['邮件数']))
#             session.add(huatu_data)
#         else:
#             huatu_data.views = data_dict['查看数']
#             huatu_data.downloads = data_dict['邮件数']
#         session.commit()

#     session.close()

#     return missingInformationData, wrongData

def addBrowsingAndDownloadData_HuaTu(case_name, year, views, downloads):
    engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=True)
    Session = sessionmaker(bind=engine)
    session = Session()

    case = session.query(Case).filter_by(name=case_name).first()
    if not case:
        print('案例不存在', case_name)
        return None

    huatu_data = session.query(HuaTuData).filter_by(case_name=case_name, year=year).first()
    if not huatu_data:
        huatu_data = HuaTuData(case_name=case.name, 
                               year=year, 
                               views=views, 
                               downloads=downloads)
        session.add(huatu_data)
    else:
        huatu_data.views = views
        huatu_data.downloads = downloads
    session.commit()

    session.close()

    return huatu_data

def getHuaTuYearData():
    engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=True)
    Session = sessionmaker(bind=engine)
    session = Session()

    huatu_data = sorted([x[0] for x in session.query(HuaTuData.year).distinct()], reverse=True)

    session.close()

    return huatu_data

def getHuaTuData(case_name, year, views=None, downloads=None):
    engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=True)
    Session = sessionmaker(bind=engine)
    session = Session()

    huatu_data = session.query(HuaTuData).filter_by(case_name=case_name, year=year).first()
    if not huatu_data:
        print('数据不存在')
        return None

    if views:
        huatu_data = [huatu_data for huatu_data in huatu_data if huatu_data.views == views]
    if downloads:
        huatu_data = [huatu_data for huatu_data in huatu_data if huatu_data.downloads == downloads]

    session.close()

    return huatu_data

def getAllHuaTuData():
    engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=True)
    Session = sessionmaker(bind=engine)
    session = Session()

    huatu_data = session.query(HuaTuData).all()

    session.close()

    return huatu_data

def updateHuaTuData(case_name, year, views=None, downloads=None):
    engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=True)
    Session = sessionmaker(bind=engine)
    session = Session()

    huatu_data = session.query(HuaTuData).filter_by(case_name=case_name, year=year).first()
    if not huatu_data:
        print('数据不存在')
        return None

    if views:
        huatu_data.views = views
    if downloads:
        huatu_data.downloads = downloads

    session.commit()
    session.close()

    return huatu_data

def deleteHuaTuData(case_name, year, views=None, downloads=None):
    engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=True)
    Session = sessionmaker(bind=engine)
    session = Session()

    huatu_data = session.query(HuaTuData).filter_by(case_name=case_name, year=year).first()
    if not huatu_data:
        print('数据不存在')
        return None

    if views:
        huatu_data = [huatu_data for huatu_data in huatu_data if huatu_data.views == views]
    if downloads:
        huatu_data = [huatu_data for huatu_data in huatu_data if huatu_data.downloads == downloads]

    for huatu_data in huatu_data:
        session.delete(huatu_data)
    session.commit()
    session.close()

    return huatu_data

def deleteAllHuaTuData():
    engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=True)
    Session = sessionmaker(bind=engine)
    session = Session()

    for huatu_data in session.query(HuaTuData).all():
        session.delete(huatu_data)
    session.commit()
    session.close()

def exportHuaTuData(path, year=None):
    engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=True)
    Session = sessionmaker(bind=engine)
    session = Session()

    if year:
        huatu_data = session.query(HuaTuData).filter_by(year=year).all()
        huatu_data_list = []
        for huatu_data in huatu_data:
            huatu_data_dict = {}
            huatu_data_dict['标题'] = huatu_data.case_name
            huatu_data_dict['查看数'] = huatu_data.views
            huatu_data_dict['邮件数'] = huatu_data.downloads
            huatu_data_list.append(huatu_data_dict)

        df = pd.DataFrame(huatu_data_list)
        df.to_excel(path, index=False)
    else:
        huatu_data_list = {}
        for year in sorted([x[0] for x in session.query(HuaTuData.year).distinct()], reverse=True):
            huatu_data = session.query(HuaTuData).filter_by(year=year).all()
            huatu_data_list[year] = []
            for huatu_data in huatu_data:
                huatu_data_dict = {}
                huatu_data_dict['标题'] = huatu_data.case_name
                huatu_data_dict['查看数'] = huatu_data.views
                huatu_data_dict['邮件数'] = huatu_data.downloads
                huatu_data_list[year].append(huatu_data_dict)

        with pd.ExcelWriter(path) as writer:
            for year in huatu_data_list:
                df = pd.DataFrame(huatu_data_list[year])
                df.to_excel(writer, sheet_name=str(year), index=False)

    session.close()
