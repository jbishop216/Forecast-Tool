from sqlalchemy import create_engine, Column, Integer, String, Float, Date, ForeignKey, Enum
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker, relationship
import enum
import datetime

Base = declarative_base()

class EmployeeType(enum.Enum):
    FTE = "FTE"
    Contractor = "Contractor"

class Employee(Base):
    __tablename__ = 'employees'
    
    id = Column(Integer, primary_key=True)
    name = Column(String, nullable=False)
    team = Column(String)
    manager_code = Column(String, nullable=False)
    cost_center = Column(String, nullable=False)
    start_date = Column(Date, nullable=False)
    end_date = Column(Date)
    employment_type = Column(String, nullable=False)  # FTE or Contractor
    
    forecasts = relationship("Forecast", back_populates="employee")
    
    @property
    def weekly_hours(self):
        if self.employment_type == "FTE":
            return 34.5
        else:  # Contractor
            return 39.0

class GA01Week(Base):
    __tablename__ = 'ga01_weeks'
    
    id = Column(Integer, primary_key=True)
    year = Column(Integer, nullable=False)
    month = Column(Integer, nullable=False)  # 1-12
    weeks = Column(Float, nullable=False)

class WorkCode(Base):
    __tablename__ = 'work_codes'
    
    id = Column(Integer, primary_key=True)
    code = Column(String, nullable=False)
    description = Column(String)

class Forecast(Base):
    __tablename__ = 'forecasts'
    
    id = Column(Integer, primary_key=True)
    employee_id = Column(Integer, ForeignKey('employees.id'), nullable=False)
    work_code_id = Column(Integer, ForeignKey('work_codes.id'), nullable=False)
    year = Column(Integer, nullable=False)
    month = Column(Integer, nullable=False)  # 1-12
    hours = Column(Float, nullable=False)
    
    employee = relationship("Employee", back_populates="forecasts")
    work_code = relationship("WorkCode")

def init_db():
    engine = create_engine('sqlite:///forecast_tool.db')
    Base.metadata.create_all(engine)
    return engine

def get_session():
    engine = create_engine('sqlite:///forecast_tool.db')
    Session = sessionmaker(bind=engine)
    return Session() 