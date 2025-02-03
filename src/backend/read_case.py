import json
import pandas as pd
from sqlalchemy.orm import sessionmaker
from sqlalchemy import create_engine
import datetime
from fuzzywuzzy import process

from src.db.init_db import CopyrightOwner, Case, BrowsingRecord, DownloadRecord

def readCaseList(path):
    # 读取 xls 或 xlsx 文件
    df = pd.read_excel(path)
    data_dict_list = [x for x in df.to_dict(orient='records') if x['案例状态'] == '已入库']

    wrong_cases = []

    engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=True)
    Session = sessionmaker(bind=engine)
    session = Session()

    for data_dict in data_dict_list:
        if (not bool(data_dict['案例标题']) or pd.isna(data_dict['案例标题'])
            or not bool(data_dict['案例版权']) or pd.isna(data_dict['案例版权'])
            or not bool(data_dict['发布时间']) or pd.isna(data_dict['发布时间'])):
            print('案例标题、案例版权、发布时间不能为空')
            wrong_cases.append(data_dict)
            continue


        owner = session.query(CopyrightOwner).filter_by(name=data_dict['案例版权']).first()
        if not owner:
            owner = CopyrightOwner(name=data_dict['案例版权'])
            session.add(owner)
            session.commit()

        case = session.query(Case).filter_by(name=data_dict['案例标题']).first()
        owner = session.query(CopyrightOwner).filter_by(name=data_dict['案例版权']).first()
        if not case:
            case = Case(name=data_dict['案例标题'], 
                        type=data_dict['产品类型'], 
                        create_time=datetime.datetime.strptime(data_dict['发布时间'], "%Y-%m-%d %H:%M:%S.%f"), 
                        is_micro=True if data_dict['是否微案例'] == '是' else False,
                        is_exclusive=True if data_dict['是否独家案例'] == '是' else False,
                        batch=int(data_dict['案例批次']),
                        submission_source=data_dict['投稿来源'],
                        contain_TN=True if data_dict['是否含有教学说明'] == '是' else False,
                        is_adapted_from_text=True if data_dict['是否由文字案例改编'] == '是' else False,
                        owner_id=owner.id)
            session.add(case)
            session.commit()

    session.close()

    return wrong_cases

def getCase(name):
    engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=True)
    Session = sessionmaker(bind=engine)
    session = Session()

    case = session.query(Case).filter_by(name=name).first()
    if not case:
        print('案例不存在')
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

def updateCase(name, alias=None, type=None, create_time=None, is_micro=None, is_exclusive=None, batch=None, submission_source=None, contain_TN=None, is_adapted_from_text=None, owner_id=None):
    engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=True)
    Session = sessionmaker(bind=engine)
    session = Session()

    case = session.query(Case).filter_by(name=name).first()
    if not case:
        print('案例不存在')
        return None

    if alias:
        if not case.alias:
            case.alias = json.dumps([alias])
        else:
            alias_list = json.loads(case.alias)
            alias_list.append(alias)
            case.alias = json.dumps(alias_list)
    if type:
        case.type = type
    if create_time:
        case.create_time = create_time
    if is_micro:
        case.is_micro = is_micro
    if is_exclusive:
        case.is_exclusive = is_exclusive
    if batch:
        case.batch = batch
    if submission_source:
        case.submission_source = submission_source
    if contain_TN:
        case.contain_TN = contain_TN
    if is_adapted_from_text:
        case.is_adapted_from_text = is_adapted_from_text
    if owner_id:
        case.owner_id = owner_id

    session.commit()
    session.close()

    return case

def deleteCase(name):
    engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=True)
    Session = sessionmaker(bind=engine)
    session = Session()

    case = session.query(Case).filter_by(name=name).first()
    if not case:
        print('案例不存在')
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

def readBrowsingAndDownloadRecord_Tsinghua(path):
    xls = pd.ExcelFile(path)
    browsingSheets = [x for x in xls.sheet_names if '浏览记录' in x]
    downloadSheets = [x for x in xls.sheet_names if '下载记录' in x]

    engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=True)
    Session = sessionmaker(bind=engine)
    session = Session()

    missingInformationBrowsingRecords = []
    missingInformationDownloadRecords = []
    wrongBrowsingRecords = []
    wrongDownloadRecords = []

    for sheet in browsingSheets:
        df = xls.parse(sheet)
        data_dict_list = df.to_dict(orient='records')

        for data_dict in data_dict_list:
            if data_dict['浏览人账号'] in ['admin', 'anonymous']:
                continue

            if (not bool(data_dict['案例名称']) or pd.isna(data_dict['案例名称'])
                or not bool(data_dict['浏览人账号']) or pd.isna(data_dict['浏览人账号'])
                or not bool(data_dict['浏览时间']) or pd.isna(data_dict['浏览时间'])):
                print('案例名称、浏览人账号、浏览时间不能为空')
                missingInformationBrowsingRecords.append(data_dict)
                continue

            case = session.query(Case).filter_by(name=data_dict['案例名称']).first()
            if not case:
                print('案例不存在')
                wrongBrowsingRecords.append(data_dict)
                continue

            if session.query(BrowsingRecord).filter_by(case_id=case.id, browser=data_dict['浏览人账号'], datetime=datetime.datetime.strptime(data_dict['浏览时间'], "%Y-%m-%d %H:%M:%S")).first():
                continue

            browsing_record = BrowsingRecord(case_id=case.id, 
                                             case_name=case.name,
                                             browser=data_dict['浏览人账号'], 
                                             browser_institution=data_dict['浏览人所在院校'], 
                                             datetime=datetime.datetime.strptime(data_dict['浏览时间'], "%Y-%m-%d %H:%M:%S"), 
                                             is_valid=None)
            session.add(browsing_record)
            session.commit()

    for sheet in downloadSheets:
        df = xls.parse(sheet)
        data_dict_list = df.to_dict(orient='records')

        for data_dict in data_dict_list:
            if data_dict['下载人账号'] in ['admin', 'anonymous']:
                continue

            if (not bool(data_dict['案例名称']) or pd.isna(data_dict['案例名称'])
                or not bool(data_dict['下载人账号']) or pd.isna(data_dict['下载人账号'])
                or not bool(data_dict['下载时间']) or pd.isna(data_dict['下载时间'])):
                print('案例名称、下载人账号、下载时间不能为空')
                missingInformationDownloadRecords.append(data_dict)
                continue

            case = session.query(Case).filter_by(name=data_dict['案例名称']).first()
            if not case:
                print('案例不存在')
                wrongDownloadRecords.append(data_dict)
                continue

            if session.query(DownloadRecord).filter_by(case_id=case.id, downloader=data_dict['下载人账号'], downloader_institution=data_dict['下载人所在院校'], datetime=datetime.datetime.strptime(data_dict['下载时间'], "%Y-%m-%d %H:%M:%S")).first():
                continue

            download_record = DownloadRecord(case_id=case.id, 
                                             case_name=case.name,
                                             downloader=data_dict['下载人账号'], 
                                             downloader_institution=data_dict['下载人所在院校'], 
                                             datetime=datetime.datetime.strptime(data_dict['下载时间'], "%Y-%m-%d %H:%M:%S"), 
                                             is_valid=None)
            session.add(download_record)
            session.commit()

    # 查询记录时间，判断有效性
    # 每年的 2 月 1 日到 7 月 31 日为一个学期，8 月 1 日到第二年 1 月 31 日为一个学期
    # 每个学期，每个账号对于每一篇文章只有第一次浏览/下载记录被记为有效
    # 若每个账号的浏览/下载记录时间前 2 个月内含有对同一篇文章有效的浏览/下载记录，则该浏览/下载记录无效
    for case in session.query(Case).all():
        browsing_records = session.query(BrowsingRecord).filter_by(case_name=case.name).all()
        if browsing_records:
            browser_group = {}
            for browsing_record in browsing_records:
                if browsing_record.browser not in browser_group:
                    browser_group[browsing_record.browser] = []
                browser_group[browsing_record.browser].append(browsing_record)
            for browser in browser_group:
                browser_group[browser].sort(key=lambda x: x.datetime)

                browser_group[browser][0].is_valid = True
                last_valid_datetime = browser_group[browser][0].datetime
                new_semester = False

                for i in range(1, len(browser_group[browser])):
                    if last_valid_datetime.month in [1]:
                        if (browser_group[browser][i].datetime.year == last_valid_datetime.year and browser_group[browser][i].datetime.month >= 2) or (browser_group[browser][i].datetime.year > last_valid_datetime.year):
                            new_semester = True
                    elif last_valid_datetime.month in [2, 3, 4, 5, 6, 7]:
                        if (browser_group[browser][i].datetime.year == last_valid_datetime.year and browser_group[browser][i].datetime.month >= 8) or (browser_group[browser][i].datetime.year > last_valid_datetime.year):
                            new_semester = True
                    elif last_valid_datetime.month in [8, 9, 10, 11, 12]:
                        if (browser_group[browser][i].datetime.year == last_valid_datetime.year + 1 and browser_group[browser][i].datetime.month >= 2) or (browser_group[browser][i].datetime.year > last_valid_datetime.year + 1):
                            new_semester = True

                    if new_semester and (browser_group[browser][i].datetime - last_valid_datetime).days >= 60:
                        browser_group[browser][i].is_valid = True
                        last_valid_datetime = browser_group[browser][i].datetime
                    else:
                        browser_group[browser][i].is_valid = False
            session.commit()
                    
        download_records = session.query(DownloadRecord).filter_by(case_name=case.name).all()
        if download_records:
            downloader_group = {}
            for download_record in download_records:
                if download_record.downloader not in downloader_group:
                    downloader_group[download_record.downloader] = []
                downloader_group[download_record.downloader].append(download_record)
            for downloader in downloader_group:
                downloader_group[downloader].sort(key=lambda x: x.datetime)

                downloader_group[downloader][0].is_valid = True
                last_valid_datetime = downloader_group[downloader][0].datetime
                new_semester = False

                for i in range(1, len(downloader_group[downloader])):
                    if last_valid_datetime.month in [1]:
                        if (downloader_group[downloader][i].datetime.year == last_valid_datetime.year and downloader_group[downloader][i].datetime.month >= 2) or (downloader_group[downloader][i].datetime.year > last_valid_datetime.year):
                            new_semester = True
                    elif last_valid_datetime.month in [2, 3, 4, 5, 6, 7]:
                        if (downloader_group[downloader][i].datetime.year == last_valid_datetime.year and downloader_group[downloader][i].datetime.month >= 8) or (downloader_group[downloader][i].datetime.year > last_valid_datetime.year):
                            new_semester = True
                    elif last_valid_datetime.month in [8, 9, 10, 11, 12]:
                        if (downloader_group[downloader][i].datetime.year == last_valid_datetime.year + 1 and downloader_group[downloader][i].datetime.month >= 2) or (downloader_group[downloader][i].datetime.year > last_valid_datetime.year + 1):
                            new_semester = True

                    if new_semester and (downloader_group[downloader][i].datetime - last_valid_datetime).days >= 60:
                        downloader_group[downloader][i].is_valid = True
                        last_valid_datetime = downloader_group[downloader][i].datetime
                    else:
                        downloader_group[downloader][i].is_valid = False
            session.commit()

    session.close()

    return missingInformationBrowsingRecords, missingInformationDownloadRecords, wrongBrowsingRecords, wrongDownloadRecords

def deleteAllBrowsingRecords():
    engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=True)
    Session = sessionmaker(bind=engine)
    session = Session()

    for browsing_record in session.query(BrowsingRecord).all():
        session.delete(browsing_record)
    session.commit()
    session.close()

def deleteAllDownloadRecords():
    engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=True)
    Session = sessionmaker(bind=engine)
    session = Session()

    for download_record in session.query(DownloadRecord).all():
        session.delete(download_record)
    session.commit()
    session.close()

def deleteAllBrowsingAndDownloadRecords():
    deleteAllBrowsingRecords()
    deleteAllDownloadRecords()

def getTest():
    engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=True)
    Session = sessionmaker(bind=engine)
    session = Session()
    cases = session.query(BrowsingRecord).filter_by(case_name='北方电机公司').all()
    print('--------------------------------')
    print('Browsing Records:', len(cases))
    for case in cases:
        print('--------------------------------')
        print(case)
    session.close()