import json
import pandas as pd
from sqlalchemy.orm import sessionmaker
from sqlalchemy import create_engine
import datetime
from fuzzywuzzy import process

from PyQt5.QtCore import QThread, pyqtSignal

from src.db.init_db import CopyrightOwner, Case, BrowsingRecord, DownloadRecord, HuaTuData, Payment, PaymentCalculatedYear

def parse_excel_float(v):
    if v is None or pd.isna(v):
        return None
    s = str(v).strip()
    if s == "":
        return None
    try:
        return float(s)
    except:
        return None
    
def parse_dt(v):
    if pd.isna(v):
        raise ValueError("时间为空")
    try:
        dt = pd.to_datetime(v)
        if isinstance(dt, pd.Timestamp):
            return dt.to_pydatetime()
        return dt
    except Exception:
        # 兜底：如果非标准格式再尝试你原来那几种 format
        s = str(v).strip()
        for fmt in ("%Y-%m-%d %H:%M:%S.%f", "%Y-%m-%d %H:%M:%S"):
            try:
                return datetime.datetime.strptime(s, fmt)
            except Exception:
                pass
        raise

# 读案例列表
def readCaseList(path):
    # 读取 Excel 文件，过滤“已入库”的记录
    df = pd.read_excel(path)
    # data_dict_list = [r for r in df.to_dict(orient='records') if (r.get('案例状态') == '已入库' and r.get('案例版权') != '浙江大学管理学院')]
    data_dict_list = [r for r in df.to_dict(orient='records') if (r.get('案例状态') == '已入库')]
    wrong_cases = []

    engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=False)
    Session = sessionmaker(bind=engine)
    session = Session()

    # -------------------------------
    # 【优化】直接加载所有案例和版权方数据
    all_cases = session.query(Case).all()
    # 构建两个缓存：按案例标题和投稿编号索引的案例字典
    # 构建以 name 和 alias 为键的案例字典，需要将 alias 转换为列表并逐个匹配
    cases_by_name_and_alias = {}
    for case in all_cases:
        name_and_alias = list(set(json.loads(case.alias) + [case.name]))
        for name_or_alias in name_and_alias:
            cases_by_name_and_alias[name_or_alias] = case
    cases_by_submission = {}
    for case in all_cases:
        cases_by_submission.setdefault(case.submission_number, []).append(case)

    all_owners = session.query(CopyrightOwner).all()
    owner_cache = {owner.name: owner for owner in all_owners if owner.name}

    # -------------------------------
    # 遍历 Excel 中的每条记录
    for data_dict in data_dict_list:
        # 检查必填字段
        title_val = data_dict.get('案例标题')
        sub_no_val = data_dict.get('投稿编号')
        owner_val = data_dict.get('案例版权')
        release_val = data_dict.get('发布时间')
        create_val = data_dict.get('创建时间')

        if (pd.isna(title_val) or str(title_val).strip() == '' or
            pd.isna(sub_no_val) or str(sub_no_val).strip() == '' or
            pd.isna(owner_val) or str(owner_val).strip() == '' or
            pd.isna(release_val) or str(release_val).strip() == '' or
            pd.isna(create_val) or str(create_val).strip() == ''):
            print('案例标题、投稿编号、案例版权、发布时间、创建时间不能为空')
            data_dict['错误信息'] = '案例标题、投稿编号、案例版权、发布时间、创建时间不能为空'
            wrong_cases.append(data_dict)
            continue

        owner_name = data_dict['案例版权']
        # if '浙江大学' in owner_name or '浙大' in owner_name or '达顿':
        #     continue
        owner = owner_cache.get(owner_name)
        if not owner:
            print('版权方不存在', owner_name)
            data_dict['错误信息'] = '版权方不存在'
            wrong_cases.append(data_dict)
            continue

        raw_title = str(data_dict['案例标题']).strip()          # Excel 
        normalized_title = raw_title.replace(' ', '').replace('　', '')  # 去空格标题，仅用于匹配/作为别名
        # case_title = data_dict['案例标题'].replace(' ', '').replace('　', '')
        submission_number = data_dict['投稿编号']

        # 先按案例标题查找，如不存在再按投稿编号查找
        # case = cases_by_name_and_alias.get(case_title)
        case = cases_by_name_and_alias.get(raw_title)
        found_by_normalized = False
        if not case:
            case = cases_by_name_and_alias.get(normalized_title)
            if case:
                found_by_normalized = True
                
        if case:
            # 如果是通过“去空格标题”匹配上的，那么把 Excel 中的标题作为标准标题，并把两者都加入别名
            if found_by_normalized:
                alias_list = json.loads(case.alias)
                # 把原始标题和去空格标题都放进 alias
                for t in [raw_title, normalized_title]:
                    if t not in alias_list:
                        alias_list.append(t)
                case.alias = json.dumps(alias_list, ensure_ascii=False)
                # 用 Excel 中的标题作为主标题
                case.name = raw_title
                # 更新缓存：保证以后既能用带空格的标题，也能用去空格标题匹配到
                cases_by_name_and_alias[raw_title] = case
                cases_by_name_and_alias[normalized_title] = case
                
            case.type = data_dict.get('产品类型', '文字')
            try:
                case.release_time = parse_dt(data_dict['发布时间'])
                case.create_time = parse_dt(data_dict['创建时间'])
            except Exception:
                print("发布时间解析错误", data_dict['发布时间'])
                wrong_cases.append(data_dict)
                continue
                
            for attr in data_dict_list[0]:
                if '正文范围' in attr:
                    text_scope = data_dict.get('正文范围')
                    break
            case.is_micro = True if text_scope == '5页以内' else False
            case.submission_source = data_dict.get('投稿来源')
            
            contain_TN = False
            for attr in data_dict:
                if '教学说明' in attr:
                    val = str(data_dict.get(attr) or '').strip()
                    contain_TN = (val == '是')
                    break
            case.contain_TN = contain_TN
            
            case.is_adapted_from_text = True if data_dict.get('是否由文字案例改编') == '是' else False
            case.owner_name = owner.name
            continue

        submission_cases = cases_by_submission.get(submission_number)
        if submission_cases:
            for case in submission_cases:
                if (case.owner_name == owner.name and 
                    case.release_time == datetime.datetime.strptime(data_dict['发布时间'], "%Y-%m-%d %H:%M:%S.%f") and
                    case.create_time == datetime.datetime.strptime(data_dict['创建时间'], "%Y-%m-%d %H:%M:%S.%f")):
                    # 已存在的案例：更新信息
                    print(f'案例已存在，改名并更新信息。原名：{case.name}，新名：{raw_title}')
                    # 更新 alias（别名）记录
                    alias_list = json.loads(case.alias)
                    # if case_title not in alias_list:
                    #     alias_list.append(case_title)
                    # if data_dict['案例标题'] not in alias_list:
                    #     alias_list.append(data_dict['案例标题'])
                    for t in [case.name, raw_title, normalized_title]:
                        if t not in alias_list:
                            alias_list.append(t)
                    case.alias = json.dumps(alias_list, ensure_ascii=False)
                    # 更新其他字段
                    case.name = data_dict['案例标题']
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
                            wrong_cases.append(data_dict)
                            continue
                    for attr in data_dict_list[0]:
                        if '正文范围' in attr:
                            text_scope = data_dict.get('正文范围')
                            break
                    case.is_micro = True if text_scope == '5页以内' else False
                    case.submission_source = data_dict.get('投稿来源')
                    
                    contain_TN = False
                    for attr in data_dict:
                        if '教学说明' in attr:
                            val = str(data_dict.get(attr) or '').strip()
                            contain_TN = (val == '是')
                            break
                    case.contain_TN = contain_TN
                    
                    case.is_adapted_from_text = True if data_dict.get('是否由文字案例改编') == '是' else False
                    case.owner_name = owner.name
                    
                    # 更新缓存
                    cases_by_name_and_alias[raw_title] = case
                    cases_by_name_and_alias[normalized_title] = case
                    continue
        
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
                wrong_cases.append(data_dict)
                continue
            
        alias_list = []
        # if '别名' in data_dict:
        #     alias = json.dumps(list(data_dict['别名'].split('，')), ensure_ascii=False)
        # else:
        #     alias = json.dumps([case_title], ensure_ascii=False)
        #     if data_dict['案例标题'] != case_title:
        #         alias = json.dumps([case_title, data_dict['案例标题']], ensure_ascii=False)
        if '别名' in data_dict and str(data_dict['别名']).strip() != '':
            alias_list.extend(str(data_dict['别名']).split('，'))
        if raw_title not in alias_list:
            alias_list.append(raw_title)
        if normalized_title != raw_title and normalized_title not in alias_list:
            alias_list.append(normalized_title)
        alias = json.dumps(alias_list, ensure_ascii=False)
        
        for attr in data_dict_list[0]:
            if '正文范围' in attr:
                text_scope = data_dict.get(attr)
                break
        contain_TN = False
        for attr in data_dict:
            if '教学说明' in attr:
                val = str(data_dict.get(attr) or '').strip()
                contain_TN = (val == '是')
                break
        case = Case(
            name=data_dict['案例标题'],
            alias=alias,
            type=data_dict.get('产品类型'),
            submission_number=submission_number,
            release_time=release_time,
            create_time=create_time,
            is_micro=True if text_scope == '5页以内' else False,
            is_exclusive = True, # 默认全为独家案例，从人大及浙大案例列表中获取非独家信息
            batch = 0,
            submission_source=data_dict.get('投稿来源'),
            contain_TN=contain_TN,
            is_adapted_from_text=True if data_dict.get('是否由文字案例改编') == '是' else False,
            owner_name=owner.name
        )
        session.add(case)
        # 同时更新缓存
        # cases_by_name_and_alias[case_title] = case
        # if data_dict['案例标题'] != case_title:
        #     cases_by_name_and_alias[data_dict['案例标题']] = case
        # cases_by_submission.setdefault(submission_number, []).append(case)
        cases_by_name_and_alias[raw_title] = case
        if normalized_title != raw_title:
            cases_by_name_and_alias[normalized_title] = case
        cases_by_submission.setdefault(submission_number, []).append(case)

    session.commit()
    session.close()
    return wrong_cases

def readCaseExclusiveAndBatch(path, owner_name, batch):
    df = pd.read_excel(path)
    df.dropna(axis=0, how='all', inplace=True)
    data_dict_list = df.to_dict(orient='records')
    title = None
    for attr in data_dict_list[0]:
        if '标题' in attr:
            title = attr
            break
    if not title:
        print('标题字段不存在')
        return None, None
    missingInformationCases = []
    wrong_cases = []

    engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=False)
    Session = sessionmaker(bind=engine)
    session = Session()

    owner_name = owner_name.replace(' ', '').replace('　', '')
    owner = session.query(CopyrightOwner).filter_by(name=owner_name).first()
    if not owner:
        print('版权方不存在')
        return None, None
    
    all_cases = session.query(Case).all()
    cases_by_name_and_alias = {}
    for case in all_cases:
        name_and_alias = list(set(json.loads(case.alias) + [case.name]))
        for name_or_alias in name_and_alias:
            cases_by_name_and_alias[name_or_alias] = case

    is_exclusive = True
    if '中国人民大学' in owner_name or '浙江大学' in owner_name or '浙大' in owner_name or '达顿' in owner_name:
        is_exclusive = False

    for data_dict in data_dict_list:
        if pd.isna(data_dict.get(title)) or data_dict.get(title).strip() == '':
            print('案例标题不能为空')
            missingInformationCases.append(data_dict)
            continue

        case_title = data_dict[title].replace(' ', '').replace('　', '')
        case = cases_by_name_and_alias.get(case_title)
        if not case:
            print('案例不存在', data_dict[title])
            wrong_cases.append(data_dict)
            continue

        case.is_exclusive = is_exclusive
        case.batch = int(batch)
        case.owner_name = owner.name
        
        if '中国人民大学' in owner_name and case.batch == 1:
            case.create_time = case.create_time.replace(year=2021)
        elif '中国人民大学' in owner_name and case.batch == 2:
            case.create_time = case.create_time.replace(year=2023)
        elif '浙江大学' in owner_name and case.batch == 1:
            case.create_time = case.create_time.replace(year=2021)
        elif '浙江大学' in owner_name and case.batch == 2:
            case.create_time = case.create_time.replace(year=2023)

    session.commit()
    session.close()
    return missingInformationCases, wrong_cases

def getCopyrightOwner(name):
    engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=False)
    Session = sessionmaker(bind=engine)
    session = Session()

    owner = session.query(CopyrightOwner).filter_by(name=name.replace(' ', '').replace('　', '')).first()
    if not owner:
        print('版权方不存在')
        return None

    session.close()

    return owner

def getAllCopyrightOwners():
    engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=False)
    Session = sessionmaker(bind=engine)
    session = Session()

    owners = session.query(CopyrightOwner).all()

    session.close()

    return owners

def updateCopyrightOwner(name, new_name):
    engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=False)
    Session = sessionmaker(bind=engine)
    session = Session()

    owner = session.query(CopyrightOwner).filter_by(name=name.replace(' ', '').replace('　', '')).first()
    if not owner:
        print('版权方不存在')
        return None

    owner.name = new_name

    session.commit()
    session.close()

    return owner

def deleteCopyrightOwner(name):
    engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=False)
    Session = sessionmaker(bind=engine)
    session = Session()

    owner = session.query(CopyrightOwner).filter_by(name=name.replace(' ', '').replace('　', '')).first()
    if not owner:
        print('版权方不存在')
        return None

    session.delete(owner)
    session.commit()
    session.close()

    return owner

def deleteAllCopyrightOwners():
    engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=False)
    Session = sessionmaker(bind=engine)
    session = Session()

    for owner in session.query(CopyrightOwner).all():
        session.delete(owner)
    session.commit()
    session.close()

# 给定案例名，返回案例所有属性
def getCase(name):
    engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=False)
    Session = sessionmaker(bind=engine)
    session = Session()

    all_cases = session.query(Case).all()
    cases_by_name_and_alias = {}
    for case in all_cases:
        name_and_alias = list(set(json.loads(case.alias) + [case.name]))
        for name_or_alias in name_and_alias:
            cases_by_name_and_alias[name_or_alias] = case
    normalized_name = name.replace(' ', '').replace('　', '')
    case = cases_by_name_and_alias.get(normalized_name)
    if not case:
        print('案例不存在', name)
        return None

    session.close()

    return case

def getSimilarCases(name):
    engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=False)
    Session = sessionmaker(bind=engine)
    session = Session()

    # 通过 fuzzywuzzy 模糊匹配案例名称
    # 直接搜索 Case.name 会给出一个长度为 1 的名称元组的列表
    all_cases = session.query(Case).all()
    cases_by_name_and_alias = {}
    for case in all_cases:
        name_and_alias = list(set(json.loads(case.alias) + [case.name]))
        for name_or_alias in name_and_alias:
            cases_by_name_and_alias[name_or_alias] = case
    similar_case_names = process.extract(name.replace(' ', '').replace('　', ''), cases_by_name_and_alias.keys(), limit=10)
    print(similar_case_names)
    similar_cases = []
    for similar_case_name in similar_case_names:
        similar_case = cases_by_name_and_alias.get(similar_case_name[0])
        similar_cases.append((similar_case_name[0], similar_case))

    session.close()

    return similar_cases

def getAllCases():
    engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=False)
    Session = sessionmaker(bind=engine)
    session = Session()

    cases = session.query(Case).all()

    session.close()

    return cases

# 设置alias，则添加别名；其他则是更新属性
def updateCase(name, alias=None, submission_number=None, type=None, release_time=None, create_time=None, is_micro=None, is_exclusive=None, batch=None, submission_source=None, contain_TN=None, is_adapted_from_text=None, owner_name=None):
    engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=False)
    Session = sessionmaker(bind=engine)
    session = Session()
    
    print(name, alias)

    all_cases = session.query(Case).all() 
    cases_by_name_and_alias = {}
    for case in all_cases:
        name_and_alias = list(set(json.loads(case.alias) + [case.name]))
        for name_or_alias in name_and_alias:
            # 原始形式
            cases_by_name_and_alias[name_or_alias] = case
            # 去掉空格的形式，也作为一个键，方便用去空格标题匹配
            normalized = name_or_alias.replace(' ', '').replace('　', '')
            if normalized not in cases_by_name_and_alias:
                cases_by_name_and_alias[normalized] = case
                
    # name = name.replace(' ', '').replace('　', '')
    raw_name = str(name).strip()
    normalized_name = raw_name.replace(' ', '').replace('　', '')
    case = cases_by_name_and_alias.get(raw_name)
    
    if not case:
        case = cases_by_name_and_alias.get(normalized_name)
    if not case:
        print('案例不存在', name)
        session.close()
        return None

    if alias:
        # alias_list = json.loads(case.alias)
        # if alias not in alias_list:
        #     alias_list.append(alias.replace(' ', '').replace('　', ''))
        #     case.alias = json.dumps(alias_list, ensure_ascii=False)
        alias_raw = str(alias).strip()
        alias_normalized = alias_raw.replace(' ', '').replace('　', '')
        alias_list = json.loads(case.alias)
        # 同时保存“原样别名”和“去空格别名”
        if alias_raw not in alias_list:
            alias_list.append(alias_raw)
        if alias_normalized not in alias_list:
            alias_list.append(alias_normalized)
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
    if batch is not None:
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

def deleteAlias(name, alias):
    engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=False)
    Session = sessionmaker(bind=engine)
    session = Session()

    all_cases = session.query(Case).all() 
    cases_by_name_and_alias = {}
    for case in all_cases:
        name_and_alias = list(set(json.loads(case.alias) + [case.name]))
        for name_or_alias in name_and_alias:
            # 原始形式
            cases_by_name_and_alias[name_or_alias] = case
            # 去掉空格的形式
            normalized = name_or_alias.replace(' ', '').replace('　', '')
            if normalized not in cases_by_name_and_alias:
                cases_by_name_and_alias[normalized] = case
                
    raw_name = str(name).strip()
    normalized_name = raw_name.replace(' ', '').replace('　', '')
    case = cases_by_name_and_alias.get(raw_name)
    
    if not case:
        case = cases_by_name_and_alias.get(normalized_name)
    if not case:
        print('案例不存在', name)
        session.close()
        return None

    alias_raw = str(alias).strip()
    alias_normalized = alias_raw.replace(' ', '').replace('　', '')
    alias_list = json.loads(case.alias)

    if alias_raw in alias_list:
        alias_list.remove(alias_raw)
    elif alias_normalized in alias_list:
        alias_list.remove(alias_normalized)
    else:
        print('别名不存在')
        session.close()
        return None

    case.alias = json.dumps(alias_list, ensure_ascii=False)

    session.commit()
    session.close()

    return case

def deleteCase(name):
    engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=False)
    Session = sessionmaker(bind=engine)
    session = Session()

    all_cases = session.query(Case).all() 
    cases_by_name_and_alias = {}
    for case in all_cases:
        name_and_alias = list(set(json.loads(case.alias) + [case.name]))
        for name_or_alias in name_and_alias:
            cases_by_name_and_alias[name_or_alias] = case
    name = name.replace(' ', '').replace('　', '')
    case = cases_by_name_and_alias.get(name)
    if not case:
        print('案例不存在', name)
        return None

    session.delete(case)
    session.commit()
    session.close()

    return case

def deleteAllCases():
    engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=False)
    Session = sessionmaker(bind=engine)
    session = Session()

    for case in session.query(Case).all():
        session.delete(case)
    session.commit()
    session.close()

# 导出xlsx
def exportCaseList(path):
    engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=False)
    Session = sessionmaker(bind=engine)
    session = Session()

    cases = session.query(Case).all()
    case_list = []
    for case in cases:
        case_dict = {}
        case_dict['案例标题'] = case.name
        case_dict['别名'] = '，'.join(json.loads(case.alias))
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

class ReadTsinghuaBrowsingAndDownloadThread(QThread):
    result = pyqtSignal(tuple)
    progress = pyqtSignal(int)
    
    def __init__(self, path):
        super().__init__()
        self.path = path
        
    def run(self):
        self.progress.emit(2)
        
        # 解析 Excel 文件及相关工作表
        xls = pd.ExcelFile(self.path)
        browsingSheets = [s for s in xls.sheet_names if '浏览记录' in s]
        downloadSheets = [s for s in xls.sheet_names if '下载记录' in s]

        # 创建数据库连接与会话
        engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=False)
        Session = sessionmaker(bind=engine)
        session = Session()

        missingInformationBrowsingRecords = []
        missingInformationDownloadRecords = []
        wrongBrowsingRecords = []
        wrongDownloadRecords = []

        # 直接加载所有案例
        all_cases = session.query(Case).all()
        
        self.progress.emit(5)
        
        # 构建案例字典：名称 -> Case 对象
        cases_by_name_and_alias = {}
        for case in all_cases:
            name_and_alias = list(set(json.loads(case.alias) + [case.name]))
            for name_or_alias in name_and_alias:
                cases_by_name_and_alias[name_or_alias] = case

        # 收集待插入的记录列表，减少频繁 commit
        new_browsing_records = []
        new_download_records = []
        
        browsing_sheets_count = len(browsingSheets)
        # --------------------
        # 处理浏览记录
        for sheet in browsingSheets:
            current_progress = 10 + (70-10) * browsingSheets.index(sheet) / browsing_sheets_count
            target_progress = 10 + (70-10) * (browsingSheets.index(sheet)+1) / browsing_sheets_count
            self.progress.emit(int(current_progress))
            
            df = xls.parse(sheet)
            df_dict = df.to_dict(orient='records')
            for data_dict in df_dict:
                self.progress.emit(int(current_progress + (target_progress - current_progress) * df_dict.index(data_dict) / len(df_dict)))
                
                # 排除系统账号
                if data_dict.get('浏览人账号') in ['admin', 'anonymous', 'luoqiong@htxt.com.cn', 'wangxuewei@htxt.com.cn', 'lichw@sem.tsinghua.edu.cn']:
                    continue

                # 检查必要字段
                name_val = data_dict.get('案例名称')
                browser_val = data_dict.get('浏览人账号')
                time_val = data_dict.get('浏览时间')

                if (pd.isna(name_val) or str(name_val).strip() == '' or
                    pd.isna(browser_val) or str(browser_val).strip() == '' or
                    pd.isna(time_val) or str(time_val).strip() == ''):
                    print('案例名称、浏览人账号、浏览时间不能为空')
                    missingInformationBrowsingRecords.append(data_dict)
                    continue

                data_dict['案例名称'] = data_dict['案例名称'].replace(' ', '').replace('　', '')
                # 直接从加载的案例字典中查找案例
                case = cases_by_name_and_alias.get(data_dict['案例名称'])
                if not case:
                    print('案例不存在', data_dict['案例名称'])
                    wrongBrowsingRecords.append(data_dict)
                    continue
                # 新增：如果是浙大的案例，忽略该浏览记录
                if ('浙江大学' in case.owner_name 
                        or '浙大' in case.owner_name):
                        # or '达顿' in case.owner_name):
                    continue

                # 解析浏览时间（仅解析一次）
                try:
                    browsing_dt = pd.to_datetime(data_dict['浏览时间'])
                    # 转成 Python 的 datetime，方便后面比较
                    if isinstance(browsing_dt, pd.Timestamp):
                        browsing_dt = browsing_dt.to_pydatetime()
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
                    datetime=browsing_dt, 
                    is_valid=None
                )
                new_browsing_records.append(record)

        # 批量插入浏览记录，减少数据库交互
        if new_browsing_records:
            session.bulk_save_objects(new_browsing_records)
            session.commit()
            
        self.progress.emit(70)

        download_sheets_count = len(downloadSheets)
        # --------------------
        # 处理下载记录
        for sheet in downloadSheets:
            current_progress = 70 + (98-70) * downloadSheets.index(sheet) / download_sheets_count
            target_progress = 70 + (98-70) * (downloadSheets.index(sheet)+1) / download_sheets_count
            self.progress.emit(int(current_progress))
            
            df = xls.parse(sheet)
            df_dict = df.to_dict(orient='records')
            for data_dict in df_dict:
                self.progress.emit(int(current_progress + (target_progress - current_progress) * df_dict.index(data_dict) / len(df_dict)))
                
                if data_dict.get('下载人账号') in ['admin', 'anonymous', 'luoqiong@htxt.com.cn', 'wangxuewei@htxt.com.cn', 'lichw@sem.tsinghua.edu.cn']:   # 忽略王雪伟，李承文
                    continue

                name_val = data_dict.get('案例名称')
                browser_val = data_dict.get('浏览人账号')
                time_val = data_dict.get('浏览时间')

                if (pd.isna(name_val) or str(name_val).strip() == '' or
                    pd.isna(browser_val) or str(browser_val).strip() == '' or
                    pd.isna(time_val) or str(time_val).strip() == ''):
                    print('案例名称、浏览人账号、浏览时间不能为空')
                    missingInformationBrowsingRecords.append(data_dict)
                    continue

                data_dict['案例名称'] = data_dict['案例名称'].replace(' ', '').replace('　', '')
                case = cases_by_name_and_alias.get(data_dict['案例名称'])
                if not case:
                    print('案例不存在', data_dict['案例名称'])
                    wrongDownloadRecords.append(data_dict)
                    continue
                # 新增：如果是浙大的案例，忽略该下载记录
                if ('浙江大学' in case.owner_name 
                        or '浙大' in case.owner_name):
                        # or '达顿' in case.owner_name):
                    continue

                try:
                    download_dt = pd.to_datetime(data_dict['下载时间'])
                    if isinstance(download_dt, pd.Timestamp):
                        download_dt = download_dt.to_pydatetime()
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
                    datetime=download_dt, 
                    is_valid=None
                )
                new_download_records.append(record)

        if new_download_records:
            session.bulk_save_objects(new_download_records)
            session.commit()

        session.close()

        self.progress.emit(100)
        self.result.emit((missingInformationBrowsingRecords, missingInformationDownloadRecords,
            wrongBrowsingRecords, wrongDownloadRecords))


def addBrowsingRecord_Tsinghua(case_name, browser, datetime):
    engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=False)
    Session = sessionmaker(bind=engine)
    session = Session()
    
    all_cases = session.query(Case).all() 
    cases_by_name_and_alias = {}
    for case in all_cases:
        name_and_alias = list(set(json.loads(case.alias) + [case.name]))
        for name_or_alias in name_and_alias:
            cases_by_name_and_alias[name_or_alias] = case
    case_name = case_name.replace(' ', '').replace('　', '')
    case = cases_by_name_and_alias.get(case_name)
    if not case:
        print('案例不存在')
        return None

    record = BrowsingRecord(
        case_name=case.name,
        browser=browser, 
        datetime=datetime,
        is_valid=None
    )
    session.add(record)
    session.commit()
    session.close()

    return record

def addDownloadRecord_Tsinghua(case_name, downloader, datetime):
    engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=False)
    Session = sessionmaker(bind=engine)
    session = Session()
    
    all_cases = session.query(Case).all() 
    cases_by_name_and_alias = {}
    for case in all_cases:
        name_and_alias = list(set(json.loads(case.alias) + [case.name]))
        for name_or_alias in name_and_alias:
            cases_by_name_and_alias[name_or_alias] = case
    case_name = case_name.replace(' ', '').replace('　', '')
    case = cases_by_name_and_alias.get(case_name)
    if not case:
        print('案例不存在')
        return None

    record = DownloadRecord(
        case_name=case.name,
        downloader=downloader, 
        datetime=datetime,
        is_valid=None
    )
    session.add(record)
    session.commit()
    session.close()

    return record

def getBrowsingRecord(case_name, browser=None, datetime=None, is_valid=None):
    engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=False)
    Session = sessionmaker(bind=engine)
    session = Session()

    all_cases = session.query(Case).all() 
    cases_by_name_and_alias = {}
    for case in all_cases:
        name_and_alias = list(set(json.loads(case.alias) + [case.name]))
        for name_or_alias in name_and_alias:
            cases_by_name_and_alias[name_or_alias] = case
    case_name = case_name.replace(' ', '').replace('　', '')
    case = cases_by_name_and_alias.get(case_name)
    if not case:
        print('案例不存在')
        return None
    
    browsing_records = case.browsing_records
    if not browsing_records:
        print('浏览记录不存在')
        return None

    if browser:
        browsing_records = [browsing_record for browsing_record in browsing_records if browsing_record.browser == browser]
    if datetime:
        browsing_records = [browsing_record for browsing_record in browsing_records if browsing_record.datetime == datetime]
    if is_valid is not None:
        browsing_records = [browsing_record for browsing_record in browsing_records if browsing_record.is_valid == is_valid]

    session.close()

    return browsing_records

def getAllBrowsingRecords():
    engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=False)
    Session = sessionmaker(bind=engine)
    session = Session()

    browsing_records = session.query(BrowsingRecord).all()

    session.close()

    return browsing_records

def deleteBrowsingRecord(case_name, browser=None, datetime=None, is_valid=None):
    engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=False)
    Session = sessionmaker(bind=engine)
    session = Session()

    all_cases = session.query(Case).all() 
    cases_by_name_and_alias = {}
    for case in all_cases:
        name_and_alias = list(set(json.loads(case.alias) + [case.name]))
        for name_or_alias in name_and_alias:
            cases_by_name_and_alias[name_or_alias] = case
    case_name = case_name.replace(' ', '').replace('　', '')
    case = cases_by_name_and_alias.get(case_name)
    if not case:
        print('案例不存在')
        return None
    
    browsing_records = case.browsing_records
    if not browsing_records:
        print('浏览记录不存在')
        return None

    if browser:
        browsing_records = [browsing_record for browsing_record in browsing_records if browsing_record.browser == browser]
    if datetime:
        browsing_records = [browsing_record for browsing_record in browsing_records if browsing_record.datetime == datetime]
    if is_valid is not None:
        browsing_records = [browsing_record for browsing_record in browsing_records if browsing_record.is_valid == is_valid]

    for browsing_record in browsing_records:
        session.delete(browsing_record)
    session.commit()
    session.close()

    return browsing_records

def deleteAllBrowsingRecords():
    engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=False)
    Session = sessionmaker(bind=engine)
    session = Session()

    for browsing_record in session.query(BrowsingRecord).all():
        session.delete(browsing_record)
    session.commit()
    session.close()

def getDownloadRecord(case_name, downloader=None, datetime=None, is_valid=None):
    engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=False)
    Session = sessionmaker(bind=engine)
    session = Session()

    all_cases = session.query(Case).all() 
    cases_by_name_and_alias = {}
    for case in all_cases:
        name_and_alias = list(set(json.loads(case.alias) + [case.name]))
        for name_or_alias in name_and_alias:
            cases_by_name_and_alias[name_or_alias] = case
    case_name = case_name.replace(' ', '').replace('　', '')
    case = cases_by_name_and_alias.get(case_name)
    if not case:
        print('案例不存在')
        return None
    
    download_records = case.download_records
    if not download_records:
        print('下载记录不存在')
        return None

    if downloader:
        download_records = [download_record for download_record in download_records if download_record.downloader == downloader]
    if datetime:
        download_records = [download_record for download_record in download_records if download_record.datetime == datetime]
    if is_valid is not None:
        download_records = [download_record for download_record in download_records if download_record.is_valid == is_valid]

    session.close()

    return download_records

def getAllDownloadRecords():
    engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=False)
    Session = sessionmaker(bind=engine)
    session = Session()

    download_records = session.query(DownloadRecord).all()

    session.close()

    return download_records

def deleteDownloadRecord(case_name, downloader=None, datetime=None, is_valid=None):
    engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=False)
    Session = sessionmaker(bind=engine)
    session = Session()

    all_cases = session.query(Case).all() 
    cases_by_name_and_alias = {}
    for case in all_cases:
        name_and_alias = list(set(json.loads(case.alias) + [case.name]))
        for name_or_alias in name_and_alias:
            cases_by_name_and_alias[name_or_alias] = case
    case_name = case_name.replace(' ', '').replace('　', '')
    case = cases_by_name_and_alias.get(case_name)
    if not case:
        print('案例不存在')
        return None
    
    download_records = case.download_records
    if not download_records:
        print('下载记录不存在')
        return None

    if downloader:
        download_records = [download_record for download_record in download_records if download_record.downloader == downloader]
    if datetime:
        download_records = [download_record for download_record in download_records if download_record.datetime == datetime]
    if is_valid is not None:
        download_records = [download_record for download_record in download_records if download_record.is_valid == is_valid]

    for download_record in download_records:
        session.delete(download_record)
    session.commit()
    session.close()

    return download_records

def deleteAllDownloadRecords():
    engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=False)
    Session = sessionmaker(bind=engine)
    session = Session()

    for download_record in session.query(DownloadRecord).all():
        session.delete(download_record)
    session.commit()
    session.close()

def exportBrowsingAndDownloadRecord(path):
    engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=False)
    Session = sessionmaker(bind=engine)
    session = Session()

    browsing_records = session.query(BrowsingRecord).all()
    # 每 50000 个记录为一个 sheet
    browsing_record_list = []
    for browsing_record in browsing_records:
        browsing_record_dict = {}
        browsing_record_dict['案例名称'] = browsing_record.case_name
        browsing_record_dict['浏览人账号'] = browsing_record.browser
        browsing_record_dict['浏览时间'] = browsing_record.datetime.strftime("%Y-%m-%d %H:%M:%S")
        browsing_record_list.append(browsing_record_dict)
        
    download_records = session.query(DownloadRecord).all()
    # 每 50000 个记录为一个 sheet
    download_record_list = []
    for download_record in download_records:
        download_record_dict = {}
        download_record_dict['案例名称'] = download_record.case_name
        download_record_dict['下载人账号'] = download_record.downloader
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

    engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=False)
    Session = sessionmaker(bind=engine)
    session = Session()
    
    year_payment = session.query(PaymentCalculatedYear).filter_by(year=int(year)).first()
    if not year_payment:
        print('年度数据不存在')
        return None, None
    year_payment.is_calculated = False

    missingInformationData = []
    wrongData = []

    # -------------------------------
    # 预加载所有案例数据
    all_cases = session.query(Case).all()
    cases_by_name_and_alias = {}
    for case in all_cases:
        name_and_alias = list(set(json.loads(case.alias) + [case.name]))
        for name_or_alias in name_and_alias:
            cases_by_name_and_alias[name_or_alias] = case

    # -------------------------------
    # 遍历 Excel 数据，更新或新增 HuaTuData
    for data_dict in data_dict_list:
        # 检查必填字段
        if (pd.isna(data_dict.get('标题')) or data_dict.get('标题').strip() == '' or
            pd.isna(data_dict.get('邮件数')) or str(data_dict.get('邮件数')).strip() == '' or
            pd.isna(data_dict.get('查看数')) or str(data_dict.get('查看数')).strip() == ''):
            missingInformationData.append(data_dict)
            continue

        data_dict['标题'] = data_dict['标题'].replace(' ', '').replace('　', '')
        case = cases_by_name_and_alias.get(data_dict['标题'])
        if not case:
            print('案例不存在:', data_dict['标题'])
            wrongData.append(data_dict)
            continue

        huatu_data = None
        for data in case.huatu_data:
            if data.year == int(year):
                huatu_data = data
                break

        if not huatu_data:
            # 不存在，则新建，并加入 huatu_dict 以便后续使用
            huatu_data = HuaTuData(
                case_name=case.name,
                year=year,
                views=int(data_dict['查看数']),
                downloads=int(data_dict['邮件数'])
            )
            session.add(huatu_data)
        else:
            # 存在则更新数据
            huatu_data.views = int(data_dict['查看数'])
            huatu_data.downloads = int(data_dict['邮件数'])
    
    session.commit()
    session.close()

    return missingInformationData, wrongData

def addBrowsingAndDownloadData_HuaTu(case_name, year, views, downloads):
    engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=False)
    Session = sessionmaker(bind=engine)
    session = Session()

    all_cases = session.query(Case).all()
    cases_by_name_and_alias = {}
    for case in all_cases:
        name_and_alias = list(set(json.loads(case.alias) + [case.name]))
        for name_or_alias in name_and_alias:
            cases_by_name_and_alias[name_or_alias] = case
    case_name = case_name.replace(' ', '').replace('　', '')
    case = cases_by_name_and_alias.get(case_name)
    if not case:
        print('案例不存在', case_name)
        return None

    huatu_data = None
    for data in case.huatu_data:
        if data.year == int(year):
            huatu_data = data
            break

    if not huatu_data:
        huatu_data = HuaTuData(
            case_name=case.name, 
            year=int(year), 
            views=int(views), 
            downloads=int(downloads)
        )
        session.add(huatu_data)
    else:
        huatu_data.views = views
        huatu_data.downloads = downloads
    session.commit()

    session.close()

    return huatu_data

def getHuaTuYearData():
    engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=False)
    Session = sessionmaker(bind=engine)
    session = Session()

    huatu_data = sorted([x[0] for x in session.query(HuaTuData.year).distinct()], reverse=True)

    session.close()

    return huatu_data

def getHuaTuData(case_name, year=None, views=None, downloads=None):
    engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=False)
    Session = sessionmaker(bind=engine)
    session = Session()

    all_cases = session.query(Case).all()
    cases_by_name_and_alias = {}
    for case in all_cases:
        name_and_alias = list(set(json.loads(case.alias) + [case.name]))
        for name_or_alias in name_and_alias:
            cases_by_name_and_alias[name_or_alias] = case
    case_name = case_name.replace(' ', '').replace('　', '')
    case = cases_by_name_and_alias.get(case_name)
    if not case:
        print('案例不存在')
        return None
    
    huatu_data = case.huatu_data
    if not huatu_data:
        print('数据不存在')
        return None

    if year is not None:
        huatu_data = [huatu_data for huatu_data in huatu_data if huatu_data.year == int(year)]
    if views is not None:
        huatu_data = [huatu_data for huatu_data in huatu_data if huatu_data.views == int(views)]
    if downloads is not None:
        huatu_data = [huatu_data for huatu_data in huatu_data if huatu_data.downloads == int(downloads)]

    session.close()

    return huatu_data

def getAllHuaTuData():
    engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=False)
    Session = sessionmaker(bind=engine)
    session = Session()

    huatu_data = session.query(HuaTuData).all()

    session.close()

    return huatu_data

def updateHuaTuData(case_name, year, views=None, downloads=None):
    engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=False)
    Session = sessionmaker(bind=engine)
    session = Session()

    all_cases = session.query(Case).all()
    cases_by_name_and_alias = {}
    for case in all_cases:
        name_and_alias = list(set(json.loads(case.alias) + [case.name]))
        for name_or_alias in name_and_alias:
            cases_by_name_and_alias[name_or_alias] = case
    case_name = case_name.replace(' ', '').replace('　', '')
    case = cases_by_name_and_alias.get(case_name)
    if not case:
        print('案例不存在')
        return None
    
    huatu_data = None
    for data in case.huatu_data:
        if data.year == int(year):
            huatu_data = data
            break

    if not huatu_data:
        print('数据不存在')
        return None

    if views is not None:
        huatu_data.views = int(views)
    if downloads is not None:
        huatu_data.downloads = int(downloads)

    session.commit()
    session.close()

    return huatu_data

def deleteHuaTuData(case_name, year=None, views=None, downloads=None):
    engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=False)
    Session = sessionmaker(bind=engine)
    session = Session()

    all_cases = session.query(Case).all()
    cases_by_name_and_alias = {}
    for case in all_cases:
        name_and_alias = list(set(json.loads(case.alias) + [case.name]))
        for name_or_alias in name_and_alias:
            cases_by_name_and_alias[name_or_alias] = case
    case_name = case_name.replace(' ', '').replace('　', '')
    case = cases_by_name_and_alias.get(case_name)
    if not case:
        print('案例不存在')
        return None
    
    huatu_data = case.huatu_data
    if not huatu_data:
        print('数据不存在')
        return None

    if year is not None:
        huatu_data = [huatu_data for huatu_data in huatu_data if huatu_data.year == int(year)]
    if views is not None:
        huatu_data = [huatu_data for huatu_data in huatu_data if huatu_data.views == int(views)]
    if downloads is not None:
        huatu_data = [huatu_data for huatu_data in huatu_data if huatu_data.downloads == int(downloads)]

    for huatu_data in huatu_data:
        session.delete(huatu_data)
    session.commit()
    session.close()

    return huatu_data

def deleteAllHuaTuData():
    engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=False)
    Session = sessionmaker(bind=engine)
    session = Session()

    for huatu_data in session.query(HuaTuData).all():
        session.delete(huatu_data)
    session.commit()
    session.close()

def exportHuaTuData(path, year=None):
    engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=False)
    Session = sessionmaker(bind=engine)
    session = Session()

    if year:
        huatu_data = session.query(HuaTuData).filter_by(year=int(year)).all()
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
        for export_year in sorted([x[0] for x in session.query(HuaTuData.year).distinct()], reverse=True):
            huatu_data = session.query(HuaTuData).filter_by(year=export_year).all()
            huatu_data_list[export_year] = []
            for huatu_data in huatu_data:
                huatu_data_dict = {}
                huatu_data_dict['标题'] = huatu_data.case_name
                huatu_data_dict['查看数'] = huatu_data.views
                huatu_data_dict['邮件数'] = huatu_data.downloads
                huatu_data_list[export_year].append(huatu_data_dict)

        with pd.ExcelWriter(path) as writer:
            for export_year in huatu_data_list:
                df = pd.DataFrame(huatu_data_list[export_year])
                df.to_excel(writer, sheet_name=str(export_year), index=False)

    session.close()

def readYiWeiData(path, year):
    engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=False)
    Session = sessionmaker(bind=engine)
    session = Session()

    missingInformationData = []
    wrongData = []

    # -------------------------------
    # 预加载所有案例数据
    all_cases = session.query(Case).all()
    cases_by_name_and_alias = {}
    for case in all_cases:
        name_and_alias = list(set(json.loads(case.alias) + [case.name]))
        for name_or_alias in name_and_alias:
            cases_by_name_and_alias[name_or_alias] = case

    # -------------------------------
    # 遍历 Excel 数据，更新或新增 YiWeiData
    df = pd.read_excel(path)
    data_dict_list = df.to_dict(orient='records')
    if not data_dict_list:
        return [], []
    
    title, payment = None, None    
    for attr in data_dict_list[0]:
        if '标题' in attr:
            title = attr
        if '收入' in attr:
            payment = attr
    if not title or not payment:
        print('标题或收入字段不存在')
        return [], []
    
    for data_dict in data_dict_list:
        # 检查必填字段
        if (pd.isna(data_dict.get(title)) or data_dict.get(title).strip() == '' or
            pd.isna(data_dict.get(payment)) or str(data_dict.get(payment)).strip() == ''):
            missingInformationData.append(data_dict)
            continue

        data_dict[title] = data_dict[title].replace(' ', '').replace('　', '')
        case = cases_by_name_and_alias.get(data_dict[title])
        if not case:
            print('案例不存在:', data_dict[title])
            wrongData.append(data_dict)
        elif case.owner_name != '毅伟':
            print('案例不属于毅伟:', data_dict[title])
            wrongData.append(data_dict)

        yiwei_data = None
        for data in case.payments:
            if data.year == int(year):
                yiwei_data = data
                break

        if not yiwei_data:
            # 不存在，则新建，并加入 yiwei_dict 以便后续使用
            yiwei_data = Payment(
                case_name=case.name,
                year=int(year),
                views=0,
                downloads=0,
                prepaid_payment=0,
                renew_payment=0,
                real_prepaid_payment=0,
                real_renew_payment=0,
                accumulated_payment=0,
                accumulated_lack_payment=0,
                retail_payment=float(data_dict[payment])
            )
            session.add(yiwei_data)
        else:
            # 存在则更新数据
            yiwei_data.retail_payment = float(data_dict[payment])

    session.commit()
    session.close()

    return missingInformationData, wrongData

def addYiWeiData(case_name, year, retail_payment):
    engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=False)
    Session = sessionmaker(bind=engine)
    session = Session()

    all_cases = session.query(Case).all()
    cases_by_name_and_alias = {}
    for case in all_cases:
        name_and_alias = list(set(json.loads(case.alias) + [case.name]))
        for name_or_alias in name_and_alias:
            cases_by_name_and_alias[name_or_alias] = case
    case_name = case_name.replace(' ', '').replace('　', '')
    case = cases_by_name_and_alias.get(case_name)
    if not case:
        print('案例不存在', case_name)
        return None

    yiwei_data = None
    for data in case.payments:
        if data.year == int(year):
            yiwei_data = data
            break

    if not yiwei_data:
        yiwei_data = Payment(
            case_name=case.name, 
            year=int(year), 
            views=0,
            downloads=0,
            prepaid_payment=0,
            renew_payment=0,
            real_prepaid_payment=0,
            real_renew_payment=0,
            accumulated_payment=0,
            accumulated_lack_payment=0,
            retail_payment=float(retail_payment)
        )
        session.add(yiwei_data)
    else:
        yiwei_data.retail_payment = retail_payment
    session.commit()

    session.close()

    return yiwei_data

def getYiWeiYearData():
    engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=False)
    Session = sessionmaker(bind=engine)
    session = Session()

    yiwei_data = sorted([x[0] for x in session.query(Payment.year).distinct()], reverse=True)

    session.close()

    return yiwei_data

def getYiWeiData(case_name=None, year=None, retail_payment=None):
    engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=False)
    Session = sessionmaker(bind=engine)
    session = Session()

    all_cases = session.query(Case).all()
    cases_by_name_and_alias = {}
    for case in all_cases:
        name_and_alias = list(set(json.loads(case.alias) + [case.name]))
        for name_or_alias in name_and_alias:
            cases_by_name_and_alias[name_or_alias] = case

    if case_name:
        case_name = case_name.replace(' ', '').replace('　', '')
        case = cases_by_name_and_alias.get(case_name)
        if not case:
            print('案例不存在')
            return None
        yiwei_data = case.payments
    else:
        yiwei_data = session.query(Payment).all()

    if not yiwei_data:
        print('数据不存在')
        return None
    
    if year is not None:
        yiwei_data = [yiwei_data for yiwei_data in yiwei_data if yiwei_data.year == int(year)]
    if retail_payment is not None:
        yiwei_data = [yiwei_data for yiwei_data in yiwei_data if yiwei_data.retail_payment == float(retail_payment)]

    session.close()

    return yiwei_data

def deleteYiWeiData(case_name=None, year=None, retail_payment=None):
    engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=False)
    Session = sessionmaker(bind=engine)
    session = Session()

    all_cases = session.query(Case).all()
    cases_by_name_and_alias = {}
    for case in all_cases:
        name_and_alias = list(set(json.loads(case.alias) + [case.name]))
        for name_or_alias in name_and_alias:
            cases_by_name_and_alias[name_or_alias] = case

    if case_name:
        case_name = case_name.replace(' ', '').replace('　', '')
        case = cases_by_name_and_alias.get(case_name)
        if not case:
            print('案例不存在')
            return None
        yiwei_data = case.payments
    else:
        yiwei_data = session.query(Payment).all()

    if not yiwei_data:
        print('数据不存在')
        return None

    if year:
        yiwei_data = [yiwei_data for yiwei_data in yiwei_data if yiwei_data.year == int(year)]
    if retail_payment:
        yiwei_data = [yiwei_data for yiwei_data in yiwei_data if yiwei_data.retail_payment == float(retail_payment)]

    for data in yiwei_data:
        session.delete(data)

    session.commit()
    session.close()

    return yiwei_data

def readRealPaymentData(path, year):
    engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=False)
    Session = sessionmaker(bind=engine)
    session = Session()

    missingInformationData = []
    wrongData = []

    # -------------------------------
    # 预加载所有案例数据
    all_cases = session.query(Case).all()
    cases_by_name_and_alias = {}
    for case in all_cases:
        name_and_alias = list(set(json.loads(case.alias) + [case.name]))
        for name_or_alias in name_and_alias:
            cases_by_name_and_alias[name_or_alias] = case
            normalized = name_or_alias.replace(' ', '').replace('　', '')
            if normalized not in cases_by_name_and_alias:
                cases_by_name_and_alias[normalized] = case

    # -------------------------------
    # 识别标题和收入字段
    df = pd.read_excel(path)
    data_dict_list = df.to_dict(orient='records')
    if len(data_dict_list) == 0:
        return missingInformationData, wrongData

    title, real_prepaid_payment_title, real_renew_payment_title = None, None, None
    for attr in data_dict_list[0]:
        if '标题' in attr:
            title = attr
        if '预付' in attr:
            real_prepaid_payment_title = attr
        if '续付' in attr:
            real_renew_payment_title = attr
    if not title or not real_prepaid_payment_title or not real_renew_payment_title:
        print('标题或收入字段不存在')
        return missingInformationData, wrongData

    # -------------------------------
    # 遍历 Excel 数据，更新或新增 Payment
    for data_dict in data_dict_list:
        # 检查必填字段
        if pd.isna(data_dict.get(title)) or data_dict.get(title).strip() == '':
            missingInformationData.append(data_dict)
            continue

        data_dict[title] = data_dict[title].replace(' ', '').replace('　', '')
        case = cases_by_name_and_alias.get(data_dict[title])
        if not case:
            print('案例不存在:', data_dict[title])
            wrongData.append(data_dict)
            continue

        real_prepaid_payment = parse_excel_float(data_dict.get(real_prepaid_payment_title))
        real_renew_payment = parse_excel_float(data_dict.get(real_renew_payment_title))
        if real_prepaid_payment is None:
            real_prepaid_payment = 0.0
        if real_renew_payment is None:
            real_renew_payment = 0.0
        
        payment = next((pay for pay in case.payments if pay.year == int(year)), None)
        if not payment:
            payment = Payment(
                case_name=case.name,
                year=int(year),
                views=0,
                downloads=0,
                prepaid_payment=0,
                renew_payment=0,
                real_prepaid_payment=real_prepaid_payment,
                real_renew_payment=real_renew_payment,
                accumulated_payment=0,
                accumulated_lack_payment=0,
                retail_payment=0,
            )
            session.add(payment)
        else:
            payment.real_prepaid_payment = real_prepaid_payment
            payment.real_renew_payment = real_renew_payment

        # 规则：
        # 1) 只要当年有实发（预付或续付任一 > 0），就认为历史欠款结清，
        #    当年的 accumulated_lack_payment 记为 0
        # 2) 如果当年应发却没发，则累计 = 上年累计欠款 + 当年应发预付 + 当年应发续付
        if payment.real_prepaid_payment > 1e-6 or payment.real_renew_payment > 1e-6:
            payment.accumulated_lack_payment = 0
        else:
            last_year_payment = next((pay for pay in case.payments if pay.year == int(year) - 1), None)
            last_lack = last_year_payment.accumulated_lack_payment if last_year_payment else 0
            payment.accumulated_lack_payment = last_lack + payment.prepaid_payment + payment.renew_payment

    session.commit()
    session.close()

    return missingInformationData, wrongData

def readHistoryRealPaymentData(path):
    engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=False)
    Session = sessionmaker(bind=engine)
    session = Session()
    
    missingInformationData = []
    wrongData = []

    # -------------------------------
    # 预加载所有案例数据
    all_cases = session.query(Case).all()
    cases_by_name_and_alias = {}
    for case in all_cases:
        name_and_alias = list(set(json.loads(case.alias) + [case.name]))
        for name_or_alias in name_and_alias:
            # 原始形式
            cases_by_name_and_alias[name_or_alias] = case
            # 去掉空格的形式，也作为一个键，方便使用“去空格标题”匹配
            normalized = name_or_alias.replace(' ', '').replace('　', '')
            if normalized not in cases_by_name_and_alias:
                cases_by_name_and_alias[normalized] = case

    xls = pd.ExcelFile(path)
    for sheet_name in sorted([s for s in xls.sheet_names if s.strip().isdigit()]):
        df = pd.read_excel(xls, sheet_name=sheet_name)
        data_dict_list = df.to_dict(orient='records')
        if len(data_dict_list) == 0:
            print('len(data_dict_list) == 0')
            continue

        title, real_prepaid_payment_title, real_renew_payment_title = None, None, None
        for attr in data_dict_list[0]:
            if '标题' in attr:
                title = attr
            if '预付' in attr:
                real_prepaid_payment_title = attr
            if '续付' in attr:
                real_renew_payment_title = attr
        if not title or (not real_prepaid_payment_title and not real_renew_payment_title):
            print('标题或收入字段不存在，跳过表单:', sheet_name)
            continue

        year = int(sheet_name.strip())

        for data_dict in data_dict_list:
            raw_title = data_dict.get(title)
            if pd.isna(raw_title) or str(raw_title).strip() == '':
                print('标题为空')
                missingInformationData.append(data_dict)
                continue
            
            raw_title = str(raw_title).strip()
            normalized_title = raw_title.replace(' ', '').replace('　', '')

            case = cases_by_name_and_alias.get(raw_title)
            if not case:
                case = cases_by_name_and_alias.get(normalized_title)
            if not case:
                print('案例不存在:', raw_title)
                wrongData.append(data_dict)
                continue

            payment = next((pay for pay in case.payments if pay.year == int(year)), None)
            if not payment:
                # 不存在，则新建，并加入 payment_dict 以便后续使用
                payment = Payment(
                    case_name=case.name,
                    year=int(year),
                    views=0,
                    downloads=0,
                    prepaid_payment=0,
                    renew_payment=0,
                    real_prepaid_payment=0,
                    real_renew_payment=0,
                    accumulated_payment=0,
                    accumulated_lack_payment=0,
                    retail_payment=0,
                )
                session.add(payment)

            # 预付
            if real_prepaid_payment_title is None:
                real_prepaid_payment = 0.0
            else:
                v = parse_excel_float(data_dict.get(real_prepaid_payment_title))
                real_prepaid_payment = 0.0 if (v is None or v <= 1e-6) else float(v)

            # 续付
            if real_renew_payment_title is None:
                real_renew_payment = 0.0
            else:
                v = parse_excel_float(data_dict.get(real_renew_payment_title))
                real_renew_payment = 0.0 if (v is None or v <= 1e-6) else float(v)
                
            payment.real_prepaid_payment = real_prepaid_payment
            payment.real_renew_payment = real_renew_payment

            # 历史导入逻辑：
            if payment.real_prepaid_payment > 1e-6:
                payment.prepaid_payment = payment.real_prepaid_payment
            else:
                payment.prepaid_payment = 0

            if payment.real_renew_payment > 1e-6:
                payment.renew_payment = payment.real_renew_payment
            else:
                payment.renew_payment = 0
                
            # 欠款逻辑：
            # 只要该年有任何实发（预付或续付），就认为历史欠款结清
            if payment.real_prepaid_payment > 1e-6 or payment.real_renew_payment > 1e-6:
                payment.accumulated_lack_payment = 0
                payment.accumulated_payment = 8001
            else:
                last_year_payment = next((pay for pay in case.payments if pay.year == int(year) - 1), None)
                last_lack = last_year_payment.accumulated_lack_payment if last_year_payment else 0
                payment.accumulated_lack_payment = last_lack + payment.prepaid_payment + payment.renew_payment

    session.commit()
    session.close()
    
    return missingInformationData, wrongData


class calculatePaymentThread(QThread):
    finished = pyqtSignal()
    progress = pyqtSignal(int)
    
    def __init__(self, years, total_payment, decimal_value, square_selected):
        super().__init__()
        self.years = years
        self.total_payment = total_payment
        self.decimal_value = decimal_value
        self.square_selected = square_selected
    
    def get_start_year(self, case):
        owner_name = case.owner_name or ""
        if owner_name == '清华大学经济管理学院':
            return case.release_time.year if case.release_time else None
        else:
            return case.create_time.year if case.create_time else None
        
    def get_prepaid_payment_for_year(self, case, year):
        """
        计算某案例在某一年“应发预付版税”。
        这里只计算本年预付，不计算续付。
        """
        start_year = self.get_start_year(case)
        if start_year is None or year < start_year:
            return 0

        owner_name = case.owner_name or ""
        submission_source = case.submission_source or ""

        # 清华
        if owner_name == '清华大学经济管理学院':
            if year != start_year:
                return 0

            prepaid_payment = 0
            if '独立开发' in submission_source:
                prepaid_payment = 8000
            elif '合作开发' in submission_source:
                prepaid_payment = 4000
            elif '学院外' in submission_source or '外校' in submission_source:
                prepaid_payment = 5000

            if not case.contain_TN or case.is_adapted_from_text:
                prepaid_payment *= 0.5

            return prepaid_payment

        # 达顿不发预付
        elif owner_name == '达顿商学院':
            return 0

        # 人大
        elif owner_name == '中国人民大学商学院':
            if year != start_year:
                return 0
            return 2000 if case.is_micro else 4000

        # 浙大：首发年一次性预付
        elif owner_name == '浙江大学管理学院':
            publish_years = year - start_year + 1

            if case.batch == 1:
                if publish_years == 1:
                    return 1000
                elif publish_years == 2:
                    return 800
                elif publish_years == 3:
                    return 600
                elif 4 <= publish_years <= 8:
                    return 400
                else:
                    return 0

            elif case.batch == 2:
                if 1 <= publish_years <= 3:
                    if publish_years == 1:
                        return 1000
                    return 1000
                elif publish_years == 4:
                    return 800
                elif publish_years == 5:
                    return 600
                elif 6 <= publish_years <= 10:
                    return 400
                else:
                    return 0

            return 0

        return 0    
        
    def run(self):
        # 判断清华记录有效性（浏览和下载分别处理）
        engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=False)
        Session = sessionmaker(bind=engine)
        session = Session()

        all_cases = session.query(Case).all()
        
        self.progress.emit(int(2))
        
        for case in all_cases:
            #新增：如果是浙大的案例，忽略该下载记录
            if ('浙江大学' in case.owner_name 
                    or '浙大' in case.owner_name): 
                    # or '达顿' in case.owner_name):
                continue
            
            current_progress = 2 + (20-2) * all_cases.index(case) / len(all_cases)
            self.progress.emit(int(current_progress))
            
            # 浏览记录有效性判断
            browsing_records = case.browsing_records
            if browsing_records:
                browser_group = {}
                for record in browsing_records:
                    browser_group.setdefault(record.browser, []).append(record)
                for _, records in browser_group.items():
                    records.sort(key=lambda r: r.datetime)
                    last_valid_datetime = records[0].datetime
                    records[0].is_valid = True
                    none_valid = [rec for rec in records if rec.is_valid is None]
                    for rec in reversed(records):
                        if rec.is_valid:
                            last_valid_datetime = rec.datetime
                            break

                    for rec in none_valid:
                        if (rec.datetime - last_valid_datetime).days >= 1:
                            rec.is_valid = True
                            last_valid_datetime = rec.datetime
                        else:
                            rec.is_valid = False
                session.commit()

            # 下载记录有效性判断
            download_records = case.download_records
            if download_records:
                downloader_group = {}
                for record in download_records:
                    downloader_group.setdefault(record.downloader, []).append(record)
                for _, records in downloader_group.items():
                    records.sort(key=lambda r: r.datetime)
                    last_valid_datetime = records[0].datetime
                    records[0].is_valid = True
                    none_valid = [rec for rec in records if rec.is_valid is None]
                    for rec in reversed(records):
                        if rec.is_valid:
                            last_valid_datetime = rec.datetime
                            break

                    for rec in none_valid:
                        if (rec.datetime - last_valid_datetime).days >= 60:
                            rec.is_valid = True
                            last_valid_datetime = rec.datetime
                        else:
                            rec.is_valid = False
                session.commit()
                
        self.progress.emit(int(40))

        all_payments = session.query(Payment).all()
        payment_map = {}
        for payment in all_payments:
            payment_map[(payment.case_name, payment.year)] = payment
            
        # 记录数据库中原本就存在的 payment，用于判断某 case 某 year 是否需要补算
        existing_payment_keys = set(payment_map.keys())
            
        self.progress.emit(int(45))    
        
        # cases_by_year = {}
        # for case in all_cases:
        #     cases_by_year.setdefault(self.get_start_year(case), []).append(case)
        cases_by_year = {}
        for case in all_cases:
            if case.owner_name not in ['清华大学经济管理学院', '中国人民大学商学院', '达顿商学院']:
                continue
            if case.submission_source is None:
                continue
            start_year = self.get_start_year(case)
            if start_year is not None:
                cases_by_year.setdefault(start_year, []).append(case)

        year_payments = session.query(PaymentCalculatedYear).all()
        for year_payment in year_payments:
            year_payment.new_case_number = len(cases_by_year.get(year_payment.year, []))
            
        year_payment_by_year = {year_payment.year: year_payment for year_payment in year_payments}
        year_payment_by_year[self.years[-1]].is_calculated = False
        year_payment_by_year[self.years[-1]].total_payment = self.total_payment
           
        for i, year in enumerate(self.years):
            k=0
            self.year = year
            
            year_info = year_payment_by_year.get(self.year)
            year_globally_calculated = bool(year_info and year_info.is_calculated)

            for case in all_cases:
                # 计算每个案例的浏览和下载次数
                views = sum(1 for record in case.browsing_records 
                                if record.is_valid and record.datetime.year == int(self.year))
                downloads = sum(1 for record in case.download_records 
                                if record.is_valid and record.datetime.year == int(self.year))
                
                # 直接查找对应年份的 HuaTu 数据记录
                huatu_data = next((hd for hd in case.huatu_data if hd.year == int(self.year)), None)
                if huatu_data:
                    views = views + huatu_data.views
                    downloads = downloads + huatu_data.downloads

                payment = payment_map.get((case.name, int(self.year)))  
                if payment is None:
                    if self.year >= self.get_start_year(case):
                        payment = Payment(
                            case_name=case.name,
                            year=int(self.year),
                            views=views,
                            downloads=downloads,
                            prepaid_payment=0,
                            renew_payment=0,
                            real_prepaid_payment=0,
                            real_renew_payment=0,
                            accumulated_payment=0,
                            accumulated_lack_payment=0,
                            retail_payment=0,
                        )
                        session.add(payment)
                        payment_map[(case.name, int(self.year))] = payment
                else:
                    payment.views = views
                    payment.downloads = downloads

            self.progress.emit(int(50 + (100-50) * i / len(self.years)))

            current_year_payments = [
                p for (case_name, y), p in payment_map.items()
                if y == self.year
            ]

            # 只统计参与 A/B 分配的案例（清华、人大、达顿）
            participant_payments = []
            for case in all_cases:
                if case.owner_name not in ['清华大学经济管理学院', '中国人民大学商学院', '达顿商学院']:
                    continue
                if case.submission_source is None:
                    continue
                p = payment_map.get((case.name, self.year))
                if p is not None:
                    participant_payments.append(p)

            total_views = sum(payment.views for payment in participant_payments)
            total_downloads = sum(payment.downloads for payment in participant_payments)

            # B 分母：必须和分子口径一致
            if self.square_selected:
                b_denominator = sum(
                    (payment.views * self.decimal_value) ** 0.5 + payment.downloads
                    for payment in participant_payments
                )
            else:
                b_denominator = sum(
                    payment.views * self.decimal_value + payment.downloads
                    for payment in participant_payments
                )
            
            # !先计算本年全部预付，再得到续付资金池
            # ============================
            current_total_payment = getattr(year_payment_by_year.get(int(self.year), object()), 'total_payment', 0)
            total_prepaid_payment_for_year = 0
            for case in all_cases:
                total_prepaid_payment_for_year += self.get_prepaid_payment_for_year(case, self.year)

            renew_pool = max(current_total_payment - total_prepaid_payment_for_year, 0)
            print('年:', self.year, '总资金池:', current_total_payment, '预付总额:', total_prepaid_payment_for_year, '续付资金池:', renew_pool)
            # =================================
            
            weight_payment = (
                0.35 * getattr(year_payment_by_year.get(self.year, object()), 'new_case_number', 0)
                + 0.3 * getattr(year_payment_by_year.get(self.year - 1, object()), 'new_case_number', 0)
                + 0.2 * getattr(year_payment_by_year.get(self.year - 2, object()), 'new_case_number', 0)
                + 0.1 * getattr(year_payment_by_year.get(self.year - 3, object()), 'new_case_number', 0)
                + 0.05 * getattr(year_payment_by_year.get(self.year - 4, object()), 'new_case_number', 0)
            )

            def calculatePaymentA(case, current_k):
                case_start_year = self.get_start_year(case)
                if case_start_year is None or weight_payment <= 0:
                    return 0, current_k

                diff = int(self.year) - case_start_year
                if diff == 0:
                    current_k += 0.7 * 0.35 / weight_payment
                    return renew_pool * 0.7 * 0.35 / weight_payment, current_k
                elif diff == 1:
                    current_k += 0.7 * 0.3 / weight_payment
                    return renew_pool * 0.7 * 0.3 / weight_payment, current_k
                elif diff == 2:
                    current_k += 0.7 * 0.2 / weight_payment
                    return renew_pool * 0.7 * 0.2 / weight_payment, current_k
                elif diff == 3:
                    current_k += 0.7 * 0.1 / weight_payment
                    return renew_pool * 0.7 * 0.1 / weight_payment, current_k
                elif diff == 4:
                    current_k += 0.7 * 0.05 / weight_payment
                    return renew_pool * 0.7 * 0.05 / weight_payment, current_k
                else:
                    return 0, current_k

            # process_views 为 1 时，按照浏览量 * 0.3 + 下载量 计算
            # process_views 为 2 时，按照浏览量开平方 + 下载量计算
            def calculatePaymentB(payment, current_k):
                if b_denominator <= 0:
                    return 0, current_k

                if self.square_selected:
                    numerator = (payment.views * self.decimal_value) ** 0.5 + payment.downloads
                else:
                    numerator = payment.views * self.decimal_value + payment.downloads

                return renew_pool * 0.3 * numerator / b_denominator, current_k

            # self.progress.emit(int(65/len(self.years)))
            self.progress.emit(int(55 + (100-55) * i / len(self.years)))
            
            sum_A = {
                '清华大学经济管理学院': 0.0,
                '中国人民大学商学院': 0.0,
                '达顿商学院': 0.0,
            }
            sum_B = {
                '清华大学经济管理学院': 0.0,
                '中国人民大学商学院': 0.0,
                '达顿商学院': 0.0,
            }
            sum_written_renew = {
                '清华大学经济管理学院': 0.0,
                '中国人民大学商学院': 0.0,
                '达顿商学院': 0.0,
            }

            for i, case in enumerate(all_cases):
                
                current_progress = 55 + (100-55) * i / len(self.years) + ((100-55) / len(all_cases)) * (all_cases.index(case) / len(all_cases))
                self.progress.emit(int(current_progress))
                
                payment = payment_map.get((case.name, self.year))
                if payment is None:
                    continue
                
                if case.submission_source is None:
                    print(f'案例 {case.name} 的 submission_source 为空，跳过该案例的预付计算')
                    prepaid_payment = 0
                    continue
                
                # 如果这一年全局已计算过，则只跳过“原本已有且链条完整”的案例
                start_year = self.get_start_year(case)
                was_existing = (case.name, self.year) in existing_payment_keys
                prev_ok = (
                    (start_year is not None and self.year <= start_year)
                    or self.year == 2015
                    or (case.name, self.year - 1) in payment_map
                )
                if year_globally_calculated and was_existing and prev_ok:
                    continue
                
                #! 清华的按发布时间开始算，外部的按创建时间开始算
                if case.owner_name == '清华大学经济管理学院':
                    A, k = calculatePaymentA(case, k)
                    B, k = calculatePaymentB(payment, k)
                    
                    prepaid_payment = self.get_prepaid_payment_for_year(case, self.year)
                    renew_payment = 0 if case.is_adapted_from_text else (A + B)

                    if start_year < 2015:
                        payment.prepaid_payment = 0
                        payment.renew_payment = renew_payment
                        if self.year == 2015:
                            payment.accumulated_payment = renew_payment
                        else:
                            last_year_payment = payment_map.get((case.name, self.year - 1))
                            if last_year_payment is None:
                                print(f'1：案例 {case.name} 缺少 {self.year - 1} 年计算结果，跳过 {self.year}')
                                continue
                            payment.accumulated_payment = renew_payment + last_year_payment.accumulated_payment

                    elif int(self.year) == start_year:
                        payment.prepaid_payment = prepaid_payment
                        payment.renew_payment = max(renew_payment - prepaid_payment, 0)
                        payment.accumulated_payment = renew_payment

                    else:
                        payment.prepaid_payment = 0
                        if self.year == 2015:
                            last_year_accumulated_payment = 0
                        else:
                            last_year_payment = payment_map.get((case.name, self.year - 1))
                            if last_year_payment is None:
                                print(f'2：案例 {case.name} 缺少 {self.year - 1} 年计算结果，跳过 {self.year}')
                                continue
                            last_year_accumulated_payment = last_year_payment.accumulated_payment

                        payment.accumulated_payment = renew_payment + last_year_accumulated_payment

                        if last_year_accumulated_payment > prepaid_payment:
                            payment.renew_payment = renew_payment
                        else:
                            payment.renew_payment = max(payment.accumulated_payment - prepaid_payment, 0)
                            
                    sum_A[case.owner_name] += A
                    sum_B[case.owner_name] += B
                    sum_written_renew[case.owner_name] += payment.renew_payment
                        
                        
                elif case.owner_name == '达顿商学院':
                    A, k = calculatePaymentA(case, k)
                    B, k = calculatePaymentB(payment, k)
                    
                    payment.prepaid_payment = 0
                    renew_payment = 0 if case.is_adapted_from_text else (A + B)

                    if start_year < 2015:
                        payment.renew_payment = renew_payment
                        if self.year == 2015:
                            payment.accumulated_payment = renew_payment
                        else:
                            last_year_payment = payment_map.get((case.name, self.year - 1))
                            if last_year_payment is None:
                                print(f'达顿1：案例 {case.name} 缺少 {self.year - 1} 年计算结果，跳过 {self.year}')
                                continue
                            payment.accumulated_payment = renew_payment + last_year_payment.accumulated_payment

                    elif int(self.year) == start_year:
                        payment.renew_payment = max(renew_payment, 0)
                        payment.accumulated_payment = renew_payment

                    else:
                        if self.year == 2015:
                            last_year_accumulated_payment = 0
                        else:
                            last_year_payment = payment_map.get((case.name, self.year - 1))
                            if last_year_payment is None:
                                print(f'达顿2：案例 {case.name} 缺少 {self.year - 1} 年计算结果，跳过 {self.year}')
                                continue
                            last_year_accumulated_payment = last_year_payment.accumulated_payment

                        payment.accumulated_payment = renew_payment + last_year_accumulated_payment
                        payment.renew_payment = renew_payment
                        
                    sum_A[case.owner_name] += A
                    sum_B[case.owner_name] += B
                    sum_written_renew[case.owner_name] += payment.renew_payment
                            

                elif case.owner_name == '中国人民大学商学院':
                    A, k = calculatePaymentA(case, k)
                    B, k = calculatePaymentB(payment, k)

                    if not case.is_exclusive:
                        A = A * 0.8

                    calculated_payment = A + B
                    if case.is_micro:
                        calculated_payment *= 0.5
                    
                    prepaid_payment = self.get_prepaid_payment_for_year(case, self.year)
                    renew_payment = calculated_payment

                    if start_year < 2015:
                        payment.prepaid_payment = 0
                        payment.renew_payment = renew_payment
                        if self.year == 2015:
                            payment.accumulated_payment = renew_payment
                        else:
                            last_year_payment = payment_map.get((case.name, self.year - 1))
                            if last_year_payment is None:
                                print(f'3：案例 {case.name} 缺少 {self.year - 1} 年计算结果，跳过 {self.year}')
                                continue
                            payment.accumulated_payment = renew_payment + last_year_payment.accumulated_payment

                    elif int(self.year) == start_year:
                        payment.prepaid_payment = prepaid_payment
                        payment.renew_payment = max(renew_payment - prepaid_payment, 0)
                        payment.accumulated_payment = renew_payment

                    else:
                        payment.prepaid_payment = 0
                        if self.year == 2015:
                            last_year_accumulated_payment = 0
                        else:
                            last_year_payment = payment_map.get((case.name, self.year - 1))
                            if last_year_payment is None:
                                print(f'4：案例 {case.name} 缺少 {self.year - 1} 年计算结果，跳过 {self.year}')
                                continue
                            last_year_accumulated_payment = last_year_payment.accumulated_payment

                        payment.accumulated_payment = renew_payment + last_year_accumulated_payment

                        if last_year_accumulated_payment > prepaid_payment:
                            payment.renew_payment = renew_payment
                        else:
                            payment.renew_payment = max(payment.accumulated_payment - prepaid_payment, 0)
                            
                    sum_A[case.owner_name] += A
                    sum_B[case.owner_name] += B
                    sum_written_renew[case.owner_name] += payment.renew_payment


                elif case.owner_name == '浙江大学管理学院':
                    publish_years = int(self.year) - start_year + 1 if start_year is not None else 0

                    if case.batch == 1:
                        payment.prepaid_payment = 4400 if self.year == start_year else 0
                        if publish_years > 0 and publish_years <= 1:
                            renew_payment = 1000
                        elif publish_years == 2:
                            renew_payment = 800
                        elif publish_years == 3:
                            renew_payment = 600
                        elif publish_years >= 4 and publish_years <= 8:
                            renew_payment = 400
                        else:
                            renew_payment = 0

                    elif case.batch == 2:
                        payment.prepaid_payment = 6400 if self.year == start_year else 0
                        if publish_years > 0 and publish_years <= 3:
                            renew_payment = 1000
                        elif publish_years == 4:
                            renew_payment = 800
                        elif publish_years == 5:
                            renew_payment = 600
                        elif publish_years >= 6 and publish_years <= 10:
                            renew_payment = 400
                        else:
                            renew_payment = 0
                    else:
                        payment.prepaid_payment = 0
                        renew_payment = 0
                        
                    payment.renew_payment = renew_payment

                    if int(self.year) == start_year:
                        payment.accumulated_payment = renew_payment
                    else:
                        if self.year == 2015:
                            payment.accumulated_payment = renew_payment
                        else:
                            last_year_payment = payment_map.get((case.name, self.year - 1))
                            if last_year_payment is None:
                                print(f'5：案例 {case.name} 缺少 {self.year - 1} 年计算结果，跳过 {self.year}')
                                continue
                            payment.accumulated_payment = renew_payment + last_year_payment.accumulated_payment
                    
            year_info = year_payment_by_year.get(self.year)
            if year_info:
                year_info.is_calculated = True
                
            if self.year == 2025:
                print("\n===== 2025 A/B 核对 =====")
            for owner in sum_A:
                print(owner)
                print("  A总和 =", sum_A[owner])
                print("  B总和 =", sum_B[owner])
                print("  A+B   =", sum_A[owner] + sum_B[owner])
                print("  最终写入renew总和 =", sum_written_renew[owner])

        session.commit()
        session.close()
        
        self.progress.emit(100)
        self.finished.emit()

def getCalculatedPaymentByYear(year):
    engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=False)
    Session = sessionmaker(bind=engine)
    session = Session()

    year_payments = session.query(Payment).filter_by(year=int(year)).all()
    if not year_payments:
        print('数据不存在')
        return None

    result = {}
    for payment in year_payments:
        result[payment.case_name] = {
            'prepaid_payment': format(payment.prepaid_payment, '.2f'),
            'renew_payment': format(payment.renew_payment, '.2f'),
            'real_prepaid_payment': format(payment.real_prepaid_payment, '.2f'),
            'real_renew_payment': format(payment.real_renew_payment, '.2f'),
            'real_payment': format(payment.real_prepaid_payment + payment.real_renew_payment, '.2f'),
            'retail_payment': format(payment.retail_payment, '.2f'),
        }

    session.close()

    return result

def getCalculatedPaymentByCase(case_name):
    engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=False)
    Session = sessionmaker(bind=engine)
    session = Session()

    case = session.query(Case).filter_by(name=case_name).first()
    if not case:
        print('案例不存在')
        return None
    
    payments = case.payments
    if not payments:
        print('数据不存在')
        return None

    result = {}
    for payment in payments:
        result[payment.year] = {
            'prepaid_payment': format(payment.prepaid_payment, '.2f'),
            'renew_payment': format(payment.renew_payment, '.2f'),
            'real_prepaid_payment': format(payment.real_prepaid_payment, '.2f'),
            'real_renew_payment': format(payment.real_renew_payment, '.2f'),
            'real_payment': format(payment.real_prepaid_payment + payment.real_renew_payment, '.2f'),
            'retail_payment': format(payment.retail_payment, '.2f'),
        }

    session.close()

    return result

def get_start_year(case):
    owner_name = case.owner_name or ""
    if owner_name == '清华大学经济管理学院':
        return case.release_time.year if case.release_time else None
    else:
        return case.create_time.year if case.create_time else None

def exportCalculatedPayment(path):
    engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=False)
    Session = sessionmaker(bind=engine)
    session = Session()

    all_payments = session.query(Payment).all()
    payment_by_year = {}
    for payment in all_payments:    
        if get_start_year(payment.case) <= payment.year:
            if payment.case.owner_name == '毅伟':
                payment_by_year.setdefault(payment.year, {}).setdefault('retail', []).append({
                    '案例标题': payment.case_name,
                    '出版年份': payment.case.release_time.year,
                    '零售版税': format(payment.retail_payment, '.2f'),
                })
            elif '浙大' in payment.case.owner_name or '浙江大学' in payment.case.owner_name:
                payment_by_year.setdefault(payment.year, {}).setdefault('royalty', []).append({
                    '案例标题': payment.case_name,
                    '版权：': payment.case.owner_name,
                    '出版年份': payment.case.release_time.year,
                    '本年度应付预付版税': format(0, '.2f'),
                    '本年度应付续付版税': format(payment.renew_payment, '.2f'),
                    '本年度实际预付版税': format(payment.real_prepaid_payment, '.2f'),
                    '本年度实际续付版税': format(payment.real_renew_payment, '.2f'),
                    '累积未支付': format(payment.accumulated_lack_payment, '.2f'),
                    '当年浏览量': format(payment.views),
                    '当年下载量': format(payment.downloads),
                })
            else:
                payment_by_year.setdefault(payment.year, {}).setdefault('royalty', []).append({
                    '案例标题': payment.case_name,
                    '版权：': payment.case.owner_name,
                    '出版年份': payment.case.release_time.year,
                    '本年度应付预付版税': format(payment.prepaid_payment, '.2f'),
                    '本年度应付续付版税': format(payment.renew_payment, '.2f'),
                    '本年度实际预付版税': format(payment.real_prepaid_payment, '.2f'),
                    '本年度实际续付版税': format(payment.real_renew_payment, '.2f'),
                    '累积未支付': format(payment.accumulated_lack_payment, '.2f'),
                    '当年浏览量': format(payment.views),
                    '当年下载量': format(payment.downloads),
                })

    if not payment_by_year:
        print("没有数据需要导出")
        session.close()
        return
    
    with pd.ExcelWriter(path) as writer:
            at_least_one_sheet = False  # 用于跟踪是否有有效的工作表

            years = sorted(payment_by_year.keys(), reverse=True)
            for year in years:
                royalty_list = payment_by_year[year].get('royalty', None)
                # 检查 DataFrame 是否为空
                if royalty_list:
                    at_least_one_sheet = True
                    royalty_df = pd.DataFrame(royalty_list)
                    royalty_df.to_excel(writer, sheet_name=str(year) + '-版税', index=False)
                else:
                    # 如果某一年的数据为空，可以选择创建一个空工作表
                    df_empty = pd.DataFrame(columns=['案例标题', '版权：', '出版年份', '本年度应付预付版税', '本年度应付续付版税', 
                                                     '本年度实际预付版税', '本年度实际续付版税', '累积未支付'])
                    df_empty.to_excel(writer, sheet_name=str(year), index=False)

                retail_list = payment_by_year[year].get('retail', None)
                # 检查 DataFrame 是否为空
                if retail_list:
                    at_least_one_sheet = True
                    retail_df = pd.DataFrame(retail_list)
                    retail_df.to_excel(writer, sheet_name=str(year) + '-零售', index=False)
            
            # 如果没有有效的工作表，则创建一个默认的工作表
            if not at_least_one_sheet:
                df_empty = pd.DataFrame(columns=['案例标题', '版权：', '出版年份', '本年度应付预付版税', '本年度应付续付版税', 
                                                 '本年度实际预付版税', '本年度实际续付版税', '累积未支付'])
                df_empty.to_excel(writer, sheet_name="默认工作表", index=False)

    session.close()

def getPaymentCalculatedYear():
    engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=False)
    Session = sessionmaker(bind=engine)
    session = Session()

    year_payments = session.query(PaymentCalculatedYear).all()

    result = {}
    for year_payment in year_payments:
        result[year_payment.year] = {
            'is_calculated': year_payment.is_calculated,
            'contain_total_payment': False if year_payment.year > 2014 and year_payment.total_payment == 0 else True,
        }

    session.close()

    return result

def updatePaymentCalculatedYear(year, total_payment):
    engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=False)
    Session = sessionmaker(bind=engine)
    session = Session()

    year_payment = session.query(PaymentCalculatedYear).filter_by(year=int(year)).first()
    if not year_payment:
        print('数据不存在')
        return None

    year_payment.total_payment = total_payment

    session.commit()
    session.close()

    return year_payment
