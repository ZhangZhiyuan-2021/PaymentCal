from sqlalchemy import create_engine
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy import Column, Integer, String, DateTime, Boolean
from sqlalchemy import ForeignKey
from sqlalchemy.orm import relationship
from sqlalchemy.orm import sessionmaker

import datetime

Base = declarative_base()

class CopyrightOwner(Base):
    __tablename__ = 'copyright_owner'

    name = Column(String, index=True, unique=True, primary_key=True)

    def __repr__(self):
        return "<CopyrightOwner(name='%s')>" % (self.name)
    
class Case(Base):
    __tablename__ = 'case'

    name = Column(String, index=True, unique=True, primary_key=True)
    alias = Column(String) # name 也会同时存在于 alias 中
    submission_number = Column(String)
    type = Column(String)
    release_time = Column(DateTime)
    create_time = Column(DateTime)
    is_micro = Column(Boolean)
    is_exclusive = Column(Boolean)
    batch = Column(Integer)
    submission_source = Column(String)
    contain_TN = Column(Boolean)
    is_adapted_from_text = Column(Boolean)
    owner_name = Column(Integer, ForeignKey('copyright_owner.name', ondelete='CASCADE'))
    owner = relationship("CopyrightOwner", backref="cases")

    def __repr__(self):
        return "<Case(name='%s', alias='%s', submission_number='%s, type='%s', release_time='%s', create_time='%s, is_micro='%s', is_exclusive='%s', batch='%s', submission_source='%s', contain_TN='%s', is_adapted_from_text='%s', owner_name='%s')>" % (self.name, self.alias, self.submission_number, self.type, self.release_time, self.create_time, self.is_micro, self.is_exclusive, self.batch, self.submission_source, self.contain_TN, self.is_adapted_from_text, self.owner_name)

class BrowsingRecord(Base):
    __tablename__ = 'browsing_record'

    id = Column(Integer, primary_key=True, autoincrement=True)
    case_name = Column(String, ForeignKey('case.name', ondelete='CASCADE'), index=True)
    case = relationship("Case", backref="browsing_records")
    browser = Column(String)
    datetime = Column(DateTime)
    is_valid = Column(Boolean)

    def __repr__(self):
        return "<BrowsingRecord(id='%s', case_name='%s', browser='%s', datetime='%s, is_valid='%s')>" % (self.id, self.case_name, self.browser, self.datetime, self.is_valid)
    
class DownloadRecord(Base):
    __tablename__ = 'download_record'

    id = Column(Integer, primary_key=True, autoincrement=True)
    case_name = Column(String, ForeignKey('case.name', ondelete='CASCADE'), index=True)
    case = relationship("Case", backref="download_records")
    downloader = Column(String)
    datetime = Column(DateTime)
    is_valid = Column(Boolean)

    def __repr__(self):
        return "<DownloadRecord(id='%s', case_name='%s', downloader='%s', datetime='%s, is_valid='%s')>" % (self.id, self.case_name, self.downloader, self.datetime, self.is_valid)
    
class HuaTuData(Base):
    __tablename__ = 'huatu_data'

    id = Column(Integer, primary_key=True, autoincrement=True)
    case_name = Column(String, ForeignKey('case.name', ondelete='CASCADE'), index=True)
    case = relationship("Case", backref="huatu_data")
    year = Column(Integer)
    views = Column(Integer)
    downloads = Column(Integer)

    def __repr__(self):
        return "<HuaTuData(id='%s', case_name='%s', year='%s', views='%s', downloads='%s')>" % (self.id, self.case_name, self.year, self.views, self.downloads)

class Payment(Base):
    __tablename__ = 'payment'

    id = Column(Integer, primary_key=True, autoincrement=True)
    case_name = Column(String, ForeignKey('case.name', ondelete='CASCADE'), index=True)
    case = relationship("Case", backref="payments")
    year = Column(Integer)
    views = Column(Integer)
    downloads = Column(Integer)
    prepaid_payment = Column(Integer)
    renew_payment = Column(Integer)
    accumulated_payment = Column(Integer)
    real_prepaid_payment = Column(Integer)
    real_renew_payment = Column(Integer)

    def __repr__(self):
        return "<Payment(id='%s', case_name='%s', year='%s', views='%s', downloads='%s', prepaid_payment='%s', renew_payment='%s', accumulated_payment='%s', real_prepaid_payment='%s', real_renew_payment='%s')>" % (self.id, self.case_name, self.year, self.views, self.downloads, self.prepaid_payment, self.renew_payment, self.accumulated_payment, self.real_prepaid_payment, self.real_renew_payment)

class PaymentCalculatedYear(Base):
    __tablename__ = 'payment_calculated_year'

    year = Column(Integer, primary_key=True)
    is_calculated = Column(Boolean)
    new_case_number = Column(Integer)

    def __repr__(self):
        return "<PaymentCalculatedYear(year='%s', is_calculated='%s', new_case_number='%s')>" % (self.year, self.is_calculated, self.new_case_number)
    
def init_db():
    engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=False)
    Base.metadata.create_all(engine)

    Session = sessionmaker(bind=engine)
    session = Session()
    all_owner_names = [owner.name for owner in session.query(CopyrightOwner).all()]
    for owner in ['清华大学经济管理学院', '中国人民大学商学院', '浙江大学管理学院']:
        if owner not in all_owner_names:
            session.add(CopyrightOwner(name=owner))
    session.commit()

    years = session.query(PaymentCalculatedYear).all()
    years_dict = {year_item.year: year_item for year_item in years}
    for year in range(2020, datetime.datetime.now().year + 1):
        if year not in years_dict:
            session.add(PaymentCalculatedYear(year=year, is_calculated=False))
    session.commit()
