import pyodbc
import pandas as pd
import datetime as dt
import dateutil as du

class TeraData():
    DRIVERNAME=''
    Database=''
    
    # Object Constructor
    def __init__(self, hostname, uid, pwd):
        self.hostname = hostname
        self.uid = uid
        self.pwd = pwd
    
    # Link Function
    def TeraDataLink(self, query, ReportName, ReportDate):
        print("EDW connecting...")
        link = 'DRIVER={DRIVERNAME};DBCNAME={hostname};UID={uid};PWD={pwd}'.format(
                      DRIVERNAME= self.DRIVERNAME,hostname=self.hostname,
                      uid=self.uid, Database=self.Database, pwd=self.pwd)
        conn = pyodbc.connect(link)
        print("Succesfully Connected, Welcome " + self.uid + "!!")
        print("Function Excuting...")
        df = pd.read_sql(query, conn)
        df.to_excel(str(ReportDate)+"_Report_"+str(ReportName)+".csv")
        print("Mining Finish! We create a data sheet:"+str(ReportName)+"!")
        return df

class ReportCode:
    
    # Object Constructor
    def __init__(self, year, month, day):
        self.year = year
        self.month = month
        self.day = day
#         TargetMonth = dt.datetime(self.year, self.month, self.day)
        self.ReportMonth = dt.datetime(self.year, self.month, self.day).strftime("%Y-%m-%d")
#         print("{REPORTMONTH}:", self.ReportMonth)
        self.time_1m_before = dt.datetime(self.year, self.month, 1).strftime("%Y-%m-%d")
#         print("{TIME_1m_BEFORE}:",self.time_1m_before)
        self.time_2m_after = (dt.datetime(self.year, self.month, 1) - du.relativedelta.relativedelta(days = 1)).strftime("%Y-%m-%d")
#         print("{TIME_2m_AFTER}:",self.time_2m_after)
        self.time_3m_before = (dt.datetime(self.year, self.month, 1) - du.relativedelta.relativedelta(months = 3)).strftime("%Y-%m-%d")
#         print("{TIME_3m_BEFORE}:",self.time_3m_before)
        self.time_6m_before = (dt.datetime(self.year, self.month, 1) - du.relativedelta.relativedelta(months = 5)).strftime("%Y-%m-%d")
#         print("{TIME_6m_BEFORE}:",self.time_6m_before)
        self.time_11m_before = (dt.datetime(self.year, self.month, 1) - du.relativedelta.relativedelta(years = 1) + du.relativedelta.relativedelta(months = 1)).strftime("%Y-%m-%d")
#         print("{TIME_11m_BEFORE}:",self.time_11m_before)
        self.time_1y_before = (dt.datetime(self.year, self.month, 1) - du.relativedelta.relativedelta(years = 1)).strftime("%Y-%m-%d")
#         print("{TIME_1y_BEFORE}:",self.time_1y_before)
        self.time_1y1m_after = (dt.datetime(self.year, self.month, 1) - du.relativedelta.relativedelta(years = 1, days = 1)).strftime("%Y-%m-%d")
#         print("{TIME_1y1m_AFTER}:",self.time_1y1m_after)
        print("report date:", self.ReportMonth)

    #Code Box    
    def Code1(self):
        query = """
                /*code*/
                SELECT A, B, Date
                FROM YOUR.TABLENAME
                WHERE A = filter1
                  AND B = filter2
                  AND Date BETWEEN '{TIME_1m_BEFORE}' AND '{REPORTMONTH}'

                """.format(TIME_1m_BEFORE = str(self.time_1m_before),
                           TIME_2m_AFTER = str(self.time_2m_after),
                           TIME_3m_BEFORE = str(self.time_3m_before),
                           TIME_6m_BEFORE = str(self.time_6m_before),
                           TIME_11m_BEFORE = str(self.time_11m_before),
                           TIME_1y_BEFORE = str(self.time_1y_before),
                           TIME_1y1m_AFTER = str(self.time_1y1m_after),
                           REPORTMONTH = str(self.ReportMonth)
                          )
        LoginINFO.TeraDataLink(query, "ReportName...", self.ReportMonth)
        path = 'CodeName....txt'
        f = open(path, 'w')
        f.write(query)
        f.close

#登入EDW
LoginINFO = TeraData('授權IP', '帳號', '密碼')
#報表日期
ReportDate = ReportCode('YYYY', 'MM', 'DD')
#撈取報表
ReportDate.Code1()
