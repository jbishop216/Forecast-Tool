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
        session = get_session()
        settings = session.query(Settings).first()
        if not settings:
            if self.employment_type == "FTE":
                return 34.5
            else:  # Contractor
                return 39.0
        else:
            if self.employment_type == "FTE":
                hours = settings.fte_hours
            else:  # Contractor
                hours = settings.contractor_hours
            session.close()
            return hours

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
    employee_id = Column(Integer, ForeignKey('employees.id'), nullable=True)
    work_code_id = Column(Integer, ForeignKey('work_codes.id'), nullable=False)
    year = Column(Integer, nullable=False)
    month = Column(Integer, nullable=False)  # 1-12
    hours = Column(Float, nullable=False)
    notes = Column(String, nullable=True)  # For planned hires or other notes
    
    employee = relationship("Employee", back_populates="forecasts")
    work_code = relationship("WorkCode")

class ProjectAllocation(Base):
    __tablename__ = 'project_allocations'
    
    id = Column(Integer, primary_key=True)
    manager_code = Column(String, nullable=False)
    year = Column(Integer, nullable=False)
    month = Column(Integer, nullable=False)  # 1-12
    hours = Column(Float, nullable=False)

class Settings(Base):
    __tablename__ = 'settings'
    
    id = Column(Integer, primary_key=True)
    fte_hours = Column(Float, default=34.5, nullable=False)
    contractor_hours = Column(Float, default=39.0, nullable=False)

class ChangeType(enum.Enum):
    NEW_HIRE = "New Hire"
    CONVERSION = "Conversion"
    TERMINATION = "Termination"
    
class PlannedChange(Base):
    __tablename__ = 'planned_changes'
    
    id = Column(Integer, primary_key=True)
    description = Column(String, nullable=False)
    change_type = Column(String, nullable=False)  # New Hire, Conversion, Termination
    effective_date = Column(Date, nullable=False)
    employee_id = Column(Integer, ForeignKey('employees.id'), nullable=True)  # Null for new hires
    
    # Fields for new hires
    name = Column(String, nullable=True)
    team = Column(String, nullable=True)
    manager_code = Column(String, nullable=True)
    cost_center = Column(String, nullable=True)
    employment_type = Column(String, nullable=True)
    
    # Fields for conversions
    target_employment_type = Column(String, nullable=True)
    
    employee = relationship("Employee", foreign_keys=[employee_id])

def init_db():
    engine = create_engine('sqlite:///forecast_tool.db')
    Base.metadata.create_all(engine)
    return engine

def get_session():
    engine = create_engine('sqlite:///forecast_tool.db')
    Session = sessionmaker(bind=engine)
    return Session() 