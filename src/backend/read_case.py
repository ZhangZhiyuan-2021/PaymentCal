import json
import pandas as pd
from sqlalchemy.orm import sessionmaker
from sqlalchemy import create_engine
import datetime
from fuzzywuzzy import process

from src.db.init_db import CopyrightOwner, Case, BrowsingRecord, DownloadRecord, HuaTuData, Payment

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
        else:
            case = updateCase(name=data_dict['案例标题'],
                              type=data_dict['产品类型'], 
                              create_time=datetime.datetime.strptime(data_dict['发布时间'], "%Y-%m-%d %H:%M:%S.%f"), 
                              is_micro=True if data_dict['是否微案例'] == '是' else False,
                              is_exclusive=True if data_dict['是否独家案例'] == '是' else False,
                              batch=int(data_dict['案例批次']),
                              submission_source=data_dict['投稿来源'],
                              contain_TN=True if data_dict['是否含有教学说明'] == '是' else False,
                              is_adapted_from_text=True if data_dict['是否由文字案例改编'] == '是' else False,
                              owner_id=owner.id)
        session.commit()

    session.close()

    return wrong_cases

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

def getAllCases():
    engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=True)
    Session = sessionmaker(bind=engine)
    session = Session()

    cases = session.query(Case).all()

    session.close()

    return cases

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
        case_dict['发布时间'] = case.create_time.strftime("%Y-%m-%d %H:%M:%S.%f")
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

    for data_dict in data_dict_list:
        if (not bool(data_dict['标题']) or pd.isna(data_dict['标题'])
            or not bool(data_dict['邮件数']) or pd.isna(data_dict['邮件数'])
            or not bool(data_dict['查看数']) or pd.isna(data_dict['查看数'])):
            missingInformationData.append(data_dict)
            continue

        case = session.query(Case).filter_by(name=data_dict['标题']).first()
        if not case:
            print('案例不存在')
            wrongData.append(data_dict)
            continue

        huatu_data = session.query(HuaTuData).filter_by(case_id=case.id, year=year).first()
        if not huatu_data:
            huatu_data = HuaTuData(case_id=case.id, 
                                   case_name=case.name, 
                                   year=year, 
                                   views=int(data_dict['查看数']), 
                                   downloads=int(data_dict['邮件数']))
            session.add(huatu_data)
        else:
            huatu_data.views = data_dict['查看数']
            huatu_data.downloads = data_dict['邮件数']
        session.commit()

    session.close()

    return missingInformationData, wrongData

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
