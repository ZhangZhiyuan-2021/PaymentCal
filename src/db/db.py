from sqlalchemy import create_engine
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy import Column, Integer, String, DateTime, Boolean
from sqlalchemy import ForeignKey
from sqlalchemy.orm import relationship

Base = declarative_base()

class CopyrightOwner(Base):
    __tablename__ = 'copyright_owner'

    id = Column(Integer, primary_key=True, autoincrement=True)
    name = Column(String)

    def __repr__(self):
        return "<CopyrightOwner(name='%s')>" % (self.name)
    
class Case(Base):
    __tablename__ = 'case'

    id = Column(Integer, primary_key=True, autoincrement=True)
    case_number = Column(String)
    name = Column(String)
    owner_id = Column(Integer, ForeignKey('copyright_owner.id', ondelete='SET NULL'))
    owner = relationship("CopyrightOwner", backref="cases")

    def __repr__(self):
        return "<Case(case_number='%s', name='%s')>" % (self.case_number, self.name)
    
class BrowsingRecord(Base):
    __tablename__ = 'browsing_record'

    id = Column(Integer, primary_key=True, autoincrement=True)
    case_id = Column(Integer, ForeignKey('case.id', ondelete='SET NULL'))
    case = relationship("Case", backref="browsing_records")
    browser = Column(String)
    browser_department = Column(String)
    datetime = Column(DateTime)
    is_valid = Column(Boolean)

    def __repr__(self):
        return "<BrowsingRecord(id='%s', case_id='%s', browser='%s', browser_department='%s', datetime='%s')>" % (self.id, self.case_id, self.browser, self.browser_department, self.datetime)
    
class DownloadRecord(Base):
    __tablename__ = 'download_record'

    id = Column(Integer, primary_key=True, autoincrement=True)
    case_id = Column(Integer, ForeignKey('case.id', ondelete='SET NULL'))
    case = relationship("Case", backref="download_records")
    downloader = Column(String)
    downloader_department = Column(String)
    datetime = Column(DateTime)
    is_valid = Column(Boolean)

    def __repr__(self):
        return "<DownloadRecord(id='%s', case_id='%s', downloader='%s', downloader_department='%s', datetime='%s')>" % (self.id, self.case_id, self.downloader, self.downloader_department, self.datetime)
    
class Payment(Base):
    __tablename__ = 'payment'

    id = Column(Integer, primary_key=True, autoincrement=True)
    case_id = Column(Integer, ForeignKey('case.id', ondelete='SET NULL'))
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
