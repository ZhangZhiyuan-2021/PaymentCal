from sqlalchemy import create_engine
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy import Column, Integer, String, DateTime, Boolean
from sqlalchemy import ForeignKey
from sqlalchemy.orm import relationship

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
    browser_institution = Column(String)
    datetime = Column(DateTime)
    is_valid = Column(Boolean)

    def __repr__(self):
        return "<BrowsingRecord(id='%s', case_name='%s', browser='%s', browser_institution='%s', datetime='%s, is_valid='%s')>" % (self.id, self.case_name, self.browser, self.browser_institution, self.datetime, self.is_valid)
    
class DownloadRecord(Base):
    __tablename__ = 'download_record'

    id = Column(Integer, primary_key=True, autoincrement=True)
    case_name = Column(String, ForeignKey('case.name', ondelete='CASCADE'), index=True)
    case = relationship("Case", backref="download_records")
    downloader = Column(String)
    downloader_institution = Column(String)
    datetime = Column(DateTime)
    is_valid = Column(Boolean)

    def __repr__(self):
        return "<DownloadRecord(id='%s', case_name='%s', downloader='%s', downloader_institution='%s', datetime='%s, is_valid='%s')>" % (self.id, self.case_name, self.downloader, self.downloader_institution, self.datetime, self.is_valid)
    
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
    payment = Column(Integer)

    def __repr__(self):
        return "<Payment(id='%s', case_id='%s', year='%s', views='%s', downloads='%s', payment='%s')>" % (self.id, self.case_id, self.year, self.views, self.downloads, self.payment)
    
def init_db():
    engine = create_engine('sqlite:///PaymentCal.db?check_same_thread=False', echo=True)
    Base.metadata.create_all(engine)
