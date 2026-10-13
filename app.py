import os
import sqlite3
import time
import json
import io
import re
from datetime import datetime, timedelta, time as dt_time
from flask import Flask, render_template, request, jsonify, send_file
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Resolve project paths relative to this file, not the process working directory.
# This prevents deployment servers from accidentally creating/reading a different empty DB.
APP_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATES_CANDIDATE = os.path.join(APP_DIR, 'templates')
STATIC_CANDIDATE = os.path.join(APP_DIR, 'static')
template_dir = TEMPLATES_CANDIDATE if os.path.isdir(TEMPLATES_CANDIDATE) else APP_DIR
static_dir = STATIC_CANDIDATE if os.path.isdir(STATIC_CANDIDATE) else APP_DIR

app = Flask(__name__, template_folder=template_dir, static_folder=static_dir)

# You may override this with PAYROLL_DB_FILE. Prefer the database beside app.py;
# if there is no DB there yet, reuse an existing payroll.db in the current folder.
_env_db = os.environ.get('PAYROLL_DB_FILE')
if _env_db:
    DB_FILE = os.path.abspath(_env_db if os.path.isabs(_env_db) else os.path.join(APP_DIR, _env_db))
else:
    _app_db = os.path.join(APP_DIR, 'payroll.db')
    _cwd_db = os.path.abspath('payroll.db')
    DB_FILE = _app_db if os.path.isfile(_app_db) or not os.path.isfile(_cwd_db) else _cwd_db

# Seed data extracted from the user's Salary Master 2.xlsx. Imported once only
# when the salary_master2 table is empty; existing saved rows are never overwritten.
SALARY_MASTER2_SEED = [{"sNo":1,"orgEmpCode":"101","payrollEmpCode":"101","employeeName":"JAGAN NATH YADAV","designation":"Helper","monthlyRate":30345.0,"remarksStatus":"Active"},{"sNo":2,"orgEmpCode":"103","payrollEmpCode":"103","employeeName":"JITENDRA KUMAR","designation":"Helper","monthlyRate":24970.0,"remarksStatus":"Active"},{"sNo":3,"orgEmpCode":"104","payrollEmpCode":"104","employeeName":"KRISHNA KUMAR","designation":"Helper","monthlyRate":29810.0,"remarksStatus":"Active"},{"sNo":4,"orgEmpCode":"106","payrollEmpCode":"106","employeeName":"ANGAREJ","designation":"Helper","monthlyRate":35310.0,"remarksStatus":"Active"},{"sNo":5,"orgEmpCode":"107","payrollEmpCode":"107","employeeName":"MANOJ KUMAR JPR","designation":"Helper","monthlyRate":22700.0,"remarksStatus":"Active"},{"sNo":6,"orgEmpCode":"109","payrollEmpCode":"109","employeeName":"GOURISHANKAR (Dilip Sharma)","designation":"Helper","monthlyRate":28160.0,"remarksStatus":"Active"},{"sNo":7,"orgEmpCode":"110","payrollEmpCode":"110","employeeName":"RAM PARVESH SINGH","designation":"Helper","monthlyRate":27615.0,"remarksStatus":"Active"},{"sNo":8,"orgEmpCode":"111","payrollEmpCode":"111","employeeName":"PRADEEP","designation":"Helper","monthlyRate":21525.0,"remarksStatus":"Active"},{"sNo":9,"orgEmpCode":"112","payrollEmpCode":"112","employeeName":"RAJKUMAR","designation":"Helper","monthlyRate":35490.0,"remarksStatus":"Active"},{"sNo":10,"orgEmpCode":"114","payrollEmpCode":"114","employeeName":"RAVI KUMAR","designation":"Helper","monthlyRate":23730.0,"remarksStatus":"Active"},{"sNo":11,"orgEmpCode":"115","payrollEmpCode":"115","employeeName":"RATAN LAL","designation":"Helper","monthlyRate":26180.0,"remarksStatus":"Active"},{"sNo":12,"orgEmpCode":"117","payrollEmpCode":"117","employeeName":"SHIV","designation":"Helper","monthlyRate":30555.0,"remarksStatus":"Active"},{"sNo":13,"orgEmpCode":"118","payrollEmpCode":"118","employeeName":"VISHWANATH (Suresh)","designation":"Helper","monthlyRate":34210.0,"remarksStatus":"Active"},{"sNo":14,"orgEmpCode":"119","payrollEmpCode":"119","employeeName":"SANOJ KUMAR YADAV","designation":"Helper","monthlyRate":26620.0,"remarksStatus":"Active"},{"sNo":15,"orgEmpCode":"120","payrollEmpCode":"120","employeeName":"TUNTUN YADAV","designation":"-","monthlyRate":24970.0,"remarksStatus":"Active"},{"sNo":16,"orgEmpCode":"121","payrollEmpCode":"121","employeeName":"NAVEEN KUMAR SHARMA","designation":"Helper","monthlyRate":29590.0,"remarksStatus":"Active"},{"sNo":17,"orgEmpCode":"122","payrollEmpCode":"122","employeeName":"SANJEEV KUMAR","designation":"Helper","monthlyRate":26070.0,"remarksStatus":"Active"},{"sNo":18,"orgEmpCode":"123","payrollEmpCode":"123","employeeName":"MANOJ KASHYAP","designation":"Helper","monthlyRate":24970.0,"remarksStatus":"Active"},{"sNo":19,"orgEmpCode":"126","payrollEmpCode":"126","employeeName":"DINESH","designation":"Helper","monthlyRate":24970.0,"remarksStatus":"Active"},{"sNo":20,"orgEmpCode":"128","payrollEmpCode":"128","employeeName":"BHARAT","designation":"Helper","monthlyRate":34210.0,"remarksStatus":"Active"},{"sNo":21,"orgEmpCode":"130","payrollEmpCode":"130","employeeName":"JITU","designation":"Helper","monthlyRate":23650.0,"remarksStatus":"Active"},{"sNo":22,"orgEmpCode":"135","payrollEmpCode":"135","employeeName":"SUSHIL KUMAR","designation":"-","monthlyRate":26620.0,"remarksStatus":"Active"},{"sNo":23,"orgEmpCode":"136","payrollEmpCode":"136","employeeName":"VISHAL GUPTA","designation":"Executive","monthlyRate":26070.0,"remarksStatus":"Active"},{"sNo":24,"orgEmpCode":"137","payrollEmpCode":"137","employeeName":"VIRENDRA","designation":"Helper","monthlyRate":24420.0,"remarksStatus":"Active"},{"sNo":25,"orgEmpCode":"139","payrollEmpCode":"139","employeeName":"DILIP","designation":"Helper","monthlyRate":24725.0,"remarksStatus":"Active"},{"sNo":26,"orgEmpCode":"140","payrollEmpCode":"140","employeeName":"KARAN  KUMAR","designation":"Helper","monthlyRate":23540.0,"remarksStatus":"Active"},{"sNo":27,"orgEmpCode":"141","payrollEmpCode":"141","employeeName":"PRDIP PANDEY","designation":"Helper","monthlyRate":23310.0,"remarksStatus":"Active"},{"sNo":28,"orgEmpCode":"142","payrollEmpCode":"142","employeeName":"RAM JAS","designation":"Helper","monthlyRate":17010.0,"remarksStatus":"Active"},{"sNo":29,"orgEmpCode":"147","payrollEmpCode":"147","employeeName":"AMIT","designation":"Helper","monthlyRate":21120.0,"remarksStatus":"Active"},{"sNo":30,"orgEmpCode":"148","payrollEmpCode":"148","employeeName":"DEEPANSHU","designation":"Executive","monthlyRate":55000.0,"remarksStatus":"Active"},{"sNo":31,"orgEmpCode":"149","payrollEmpCode":"149","employeeName":"JAI SINGH","designation":"Helper","monthlyRate":24970.0,"remarksStatus":"Active"},{"sNo":32,"orgEmpCode":"153","payrollEmpCode":"153","employeeName":"RISHABH GUPTA","designation":"Helper","monthlyRate":36080.0,"remarksStatus":"Active"},{"sNo":33,"orgEmpCode":"154","payrollEmpCode":"154","employeeName":"SUNIL SINGH","designation":"Helper","monthlyRate":21525.0,"remarksStatus":"Active"},{"sNo":34,"orgEmpCode":"156","payrollEmpCode":"156","employeeName":"JAGDAMMA","designation":"Helper","monthlyRate":26450.0,"remarksStatus":"Active"},{"sNo":35,"orgEmpCode":"157","payrollEmpCode":"157","employeeName":"BHAGWAN PRASAD","designation":"Helper","monthlyRate":24090.0,"remarksStatus":"Active"},{"sNo":36,"orgEmpCode":"159","payrollEmpCode":"159","employeeName":"RINKU","designation":"Helper","monthlyRate":23100.0,"remarksStatus":"Active"},{"sNo":37,"orgEmpCode":"160","payrollEmpCode":"160","employeeName":"SANTOSH KUMAR DASS","designation":"Helper","monthlyRate":24090.0,"remarksStatus":"Active"},{"sNo":38,"orgEmpCode":"163","payrollEmpCode":"163","employeeName":"NAVEEN KUMAR","designation":"Executive","monthlyRate":48000.0,"remarksStatus":"Active"},{"sNo":39,"orgEmpCode":"171","payrollEmpCode":"171","employeeName":"MANI BHUSHAN CHOUDHARI","designation":"Helper","monthlyRate":21120.0,"remarksStatus":"Active"},{"sNo":40,"orgEmpCode":"173","payrollEmpCode":"173","employeeName":"RAKESH KUMAR DAS","designation":"-","monthlyRate":22440.0,"remarksStatus":"Active"},{"sNo":41,"orgEmpCode":"176","payrollEmpCode":"176","employeeName":"PAWAN KUMAR","designation":"Helper","monthlyRate":20570.0,"remarksStatus":"Active"},{"sNo":42,"orgEmpCode":"178","payrollEmpCode":"178","employeeName":"AMIT KUMAR  AGGARWAL","designation":"Helper","monthlyRate":27830.0,"remarksStatus":"Active"},{"sNo":43,"orgEmpCode":"180","payrollEmpCode":"180","employeeName":"GUDIYA","designation":"Helper","monthlyRate":23650.0,"remarksStatus":"Active"},{"sNo":44,"orgEmpCode":"185","payrollEmpCode":"185","employeeName":"TARANG GUPTA","designation":"Helper","monthlyRate":26450.0,"remarksStatus":"Active"},{"sNo":45,"orgEmpCode":"186","payrollEmpCode":"186","employeeName":"RAMRAJ","designation":"Helper","monthlyRate":21450.0,"remarksStatus":"Active"},{"sNo":46,"orgEmpCode":"187","payrollEmpCode":"187","employeeName":"DEEPAK KUMAR","designation":"-","monthlyRate":20020.0,"remarksStatus":"Active"},{"sNo":47,"orgEmpCode":"189","payrollEmpCode":"189","employeeName":"DILEEP PANDEY","designation":"Helper","monthlyRate":23230.0,"remarksStatus":"Active"},{"sNo":48,"orgEmpCode":"191","payrollEmpCode":"191","employeeName":"JATIN BAGGA","designation":"Executive","monthlyRate":28080.0,"remarksStatus":"Active"},{"sNo":49,"orgEmpCode":"192","payrollEmpCode":"209","employeeName":"GOPAL","designation":"-","monthlyRate":20900.0,"remarksStatus":"Active"},{"sNo":50,"orgEmpCode":"194","payrollEmpCode":"194","employeeName":"MAYUR KESHWANI","designation":"-","monthlyRate":54890.0,"remarksStatus":"Active"},{"sNo":51,"orgEmpCode":"206","payrollEmpCode":"207","employeeName":"SHIVA","designation":"-","monthlyRate":23000.0,"remarksStatus":"Active"},{"sNo":52,"orgEmpCode":"212","payrollEmpCode":"223","employeeName":"VIKAS RAO","designation":"-","monthlyRate":20570.0,"remarksStatus":"Active"},{"sNo":53,"orgEmpCode":"215","payrollEmpCode":"206","employeeName":"DINESH JAIPUR","designation":"-","monthlyRate":17270.0,"remarksStatus":"Active"},{"sNo":54,"orgEmpCode":"217","payrollEmpCode":"210","employeeName":"LALIT JOSHI","designation":"-","monthlyRate":15301.0,"remarksStatus":"Active"},{"sNo":55,"orgEmpCode":"224","payrollEmpCode":"211","employeeName":"RINKU 2nd","designation":"-","monthlyRate":20460.0,"remarksStatus":"Active"},{"sNo":56,"orgEmpCode":"225","payrollEmpCode":"200","employeeName":"RAVI KANT","designation":"-","monthlyRate":19910.0,"remarksStatus":"Active"},{"sNo":57,"orgEmpCode":"226","payrollEmpCode":"208","employeeName":"CHANDER PAL","designation":"-","monthlyRate":18260.0,"remarksStatus":"Active"},{"sNo":58,"orgEmpCode":"232","payrollEmpCode":"224","employeeName":"DINESH SHARMA","designation":"-","monthlyRate":21340.0,"remarksStatus":"Active"},{"sNo":59,"orgEmpCode":"235","payrollEmpCode":"235","employeeName":"GEETA","designation":"-","monthlyRate":13310.0,"remarksStatus":"Active"},{"sNo":60,"orgEmpCode":"244","payrollEmpCode":"225","employeeName":"PREETI","designation":"-","monthlyRate":20020.0,"remarksStatus":"Active"},{"sNo":61,"orgEmpCode":"248","payrollEmpCode":"248","employeeName":"KIRAN SHARMA","designation":"-","monthlyRate":16800.0,"remarksStatus":"Active"},{"sNo":62,"orgEmpCode":"255","payrollEmpCode":"255","employeeName":"MANU VERMA","designation":"-","monthlyRate":13310.0,"remarksStatus":"Active"},{"sNo":63,"orgEmpCode":"256","payrollEmpCode":"227","employeeName":"RAKESH MAURYA","designation":"-","monthlyRate":19140.0,"remarksStatus":"Active"},{"sNo":64,"orgEmpCode":"257","payrollEmpCode":"257","employeeName":"SADHNA","designation":"-","monthlyRate":13310.0,"remarksStatus":"Active"},{"sNo":65,"orgEmpCode":"260","payrollEmpCode":"260","employeeName":"SHARDA","designation":"-","monthlyRate":13090.0,"remarksStatus":"Active"},{"sNo":66,"orgEmpCode":"261","payrollEmpCode":"220","employeeName":"VIJENDER KUMAR","designation":"-","monthlyRate":24530.0,"remarksStatus":"Active"},{"sNo":67,"orgEmpCode":"266","payrollEmpCode":"266","employeeName":"SANTOSH","designation":"-","monthlyRate":14960.0,"remarksStatus":"Active"},{"sNo":68,"orgEmpCode":"273","payrollEmpCode":"234","employeeName":"RAVINDER KUMAR PANDEY","designation":"-","monthlyRate":19580.0,"remarksStatus":"Active"},{"sNo":69,"orgEmpCode":"274","payrollEmpCode":"230","employeeName":"SONU PATEL","designation":"-","monthlyRate":19140.0,"remarksStatus":"Active"},{"sNo":70,"orgEmpCode":"275","payrollEmpCode":"282","employeeName":"YASH SHRIVASTAV","designation":"-","monthlyRate":17930.0,"remarksStatus":"Active"},{"sNo":71,"orgEmpCode":"277","payrollEmpCode":"235","employeeName":"LAL MANI","designation":"-","monthlyRate":20680.0,"remarksStatus":"Active"},{"sNo":72,"orgEmpCode":"278","payrollEmpCode":"278","employeeName":"PREETI 2ND","designation":"-","monthlyRate":14960.0,"remarksStatus":"Active"},{"sNo":73,"orgEmpCode":"280","payrollEmpCode":"280","employeeName":"DURGESH","designation":"-","monthlyRate":18150.0,"remarksStatus":"Active"},{"sNo":74,"orgEmpCode":"282","payrollEmpCode":"282","employeeName":"SANDWICH/RAMNARESH","designation":"-","monthlyRate":17380.0,"remarksStatus":"Active"},{"sNo":75,"orgEmpCode":"283","payrollEmpCode":"271","employeeName":"KASHISH","designation":"-","monthlyRate":17930.0,"remarksStatus":"Active"},{"sNo":76,"orgEmpCode":"285","payrollEmpCode":"285","employeeName":"NEETU KASHYAP","designation":"-","monthlyRate":13310.0,"remarksStatus":"Active"},{"sNo":77,"orgEmpCode":"288","payrollEmpCode":"246","employeeName":"BHAGWATI","designation":"-","monthlyRate":18700.0,"remarksStatus":"Active"},{"sNo":78,"orgEmpCode":"292","payrollEmpCode":"292","employeeName":"SUMAN","designation":"-","monthlyRate":13310.0,"remarksStatus":"Active"},{"sNo":79,"orgEmpCode":"298","payrollEmpCode":"247","employeeName":"RISHABH  S/O LAKSHMI CHAN","designation":"-","monthlyRate":22000.0,"remarksStatus":"Active"},{"sNo":80,"orgEmpCode":"300","payrollEmpCode":"237","employeeName":"RAHUL KUMAR","designation":"-","monthlyRate":24640.0,"remarksStatus":"Active"},{"sNo":81,"orgEmpCode":"312","payrollEmpCode":"241","employeeName":"RANJEET KUMAR","designation":"-","monthlyRate":17820.0,"remarksStatus":"Active"},{"sNo":82,"orgEmpCode":"313","payrollEmpCode":"239","employeeName":"MANOJ KUMAR","designation":"-","monthlyRate":17010.0,"remarksStatus":"Active"},{"sNo":83,"orgEmpCode":"315","payrollEmpCode":"260","employeeName":"AZAD","designation":"-","monthlyRate":17270.0,"remarksStatus":"Active"},{"sNo":84,"orgEmpCode":"316","payrollEmpCode":"238","employeeName":"KULDEEP SAIN","designation":"-","monthlyRate":18480.0,"remarksStatus":"Active"},{"sNo":85,"orgEmpCode":"319","payrollEmpCode":"243","employeeName":"NARAYAN","designation":"-","monthlyRate":19690.0,"remarksStatus":"Active"},{"sNo":86,"orgEmpCode":"323","payrollEmpCode":"245","employeeName":"ANKUR","designation":"-","monthlyRate":18150.0,"remarksStatus":"Active"},{"sNo":87,"orgEmpCode":"324","payrollEmpCode":"272","employeeName":"MAN MOHAN SINGH","designation":"-","monthlyRate":19320.0,"remarksStatus":"Active"},{"sNo":88,"orgEmpCode":"338","payrollEmpCode":"338","employeeName":"NITIN VERMA","designation":"-","monthlyRate":16390.0,"remarksStatus":"Active"},{"sNo":89,"orgEmpCode":"339","payrollEmpCode":"266","employeeName":"ASHISH SRIVASTAV","designation":"-","monthlyRate":19320.0,"remarksStatus":"Active"},{"sNo":90,"orgEmpCode":"340","payrollEmpCode":"270","employeeName":"GOVIND KUMAR","designation":"-","monthlyRate":28080.0,"remarksStatus":"Active"},{"sNo":91,"orgEmpCode":"341","payrollEmpCode":"264","employeeName":"ANKIT","designation":"-","monthlyRate":26000.0,"remarksStatus":"Active"},{"sNo":92,"orgEmpCode":"343","payrollEmpCode":"265","employeeName":"ARJUN KUMAR","designation":"-","monthlyRate":19850.0,"remarksStatus":"Active"},{"sNo":93,"orgEmpCode":"359","payrollEmpCode":"276","employeeName":"SATISH CHAND","designation":"-","monthlyRate":24530.0,"remarksStatus":"Active"},{"sNo":94,"orgEmpCode":"360","payrollEmpCode":"274","employeeName":"RAHUL SHARMA","designation":"-","monthlyRate":20240.0,"remarksStatus":"Active"},{"sNo":95,"orgEmpCode":"366","payrollEmpCode":"366","employeeName":"SACHIN S/O JAI PRAKASH","designation":"-","monthlyRate":13200.0,"remarksStatus":"Active"},{"sNo":96,"orgEmpCode":"369","payrollEmpCode":"369","employeeName":"ANSHUL JHA","designation":"-","monthlyRate":14300.0,"remarksStatus":"Active"},{"sNo":97,"orgEmpCode":"372","payrollEmpCode":"372","employeeName":"ARJUN","designation":"-","monthlyRate":14300.0,"remarksStatus":"Active"},{"sNo":98,"orgEmpCode":"375","payrollEmpCode":"375","employeeName":"RAHUL","designation":"-","monthlyRate":19250.0,"remarksStatus":"Active"},{"sNo":99,"orgEmpCode":"376","payrollEmpCode":"283","employeeName":"DEVINDER SINGH","designation":"-","monthlyRate":26250.0,"remarksStatus":"Active"},{"sNo":100,"orgEmpCode":"378","payrollEmpCode":"296","employeeName":"VISHAL","designation":"-","monthlyRate":16500.0,"remarksStatus":"Active"},{"sNo":101,"orgEmpCode":"380","payrollEmpCode":"380","employeeName":"USHA RANI","designation":"-","monthlyRate":12100.0,"remarksStatus":"Active"},{"sNo":102,"orgEmpCode":"389","payrollEmpCode":"389","employeeName":"AMITABH VERMA","designation":"-","monthlyRate":16500.0,"remarksStatus":"Active"},{"sNo":103,"orgEmpCode":"393","payrollEmpCode":"393","employeeName":"AMIT RATHOR","designation":"-","monthlyRate":16500.0,"remarksStatus":"Active"},{"sNo":104,"orgEmpCode":"394","payrollEmpCode":"291","employeeName":"MANISHA RAWAT","designation":"-","monthlyRate":19800.0,"remarksStatus":"Active"},{"sNo":105,"orgEmpCode":"395","payrollEmpCode":"293","employeeName":"SURESH KUMAR S/O RAM ROOP","designation":"-","monthlyRate":20350.0,"remarksStatus":"Active"},{"sNo":106,"orgEmpCode":"396","payrollEmpCode":"396","employeeName":"NISHA D/O DHEERAJ","designation":"-","monthlyRate":12100.0,"remarksStatus":"Active"},{"sNo":107,"orgEmpCode":"398","payrollEmpCode":"398","employeeName":"BHUPENDRA PRATAP SINGH","designation":"-","monthlyRate":22000.0,"remarksStatus":"Active"},{"sNo":108,"orgEmpCode":"400","payrollEmpCode":"400","employeeName":"SHALINI CHAURASIYA","designation":"-","monthlyRate":19800.0,"remarksStatus":"Active"},{"sNo":109,"orgEmpCode":"405","payrollEmpCode":"286","employeeName":"AALOK GUPTA","designation":"-","monthlyRate":16800.0,"remarksStatus":"Active"},{"sNo":110,"orgEmpCode":"409","payrollEmpCode":"409","employeeName":"MINTU KUMARI","designation":"-","monthlyRate":12100.0,"remarksStatus":"Active"},{"sNo":111,"orgEmpCode":"414","payrollEmpCode":"414","employeeName":"NAVIN KUMAR S/O PADAM KUM","designation":"-","monthlyRate":23100.0,"remarksStatus":"Active"},{"sNo":112,"orgEmpCode":"420","payrollEmpCode":"420","employeeName":"DINESH KUMAR S/O LATE RAM","designation":"-","monthlyRate":20000.0,"remarksStatus":"Active"},{"sNo":113,"orgEmpCode":"423","payrollEmpCode":"423","employeeName":"HARSH GAUR S/O LATE DEEPA","designation":"-","monthlyRate":18000.0,"remarksStatus":"Active"},{"sNo":114,"orgEmpCode":"424","payrollEmpCode":"424","employeeName":"MUKESH S/O RAM BHAROSE","designation":"-","monthlyRate":20000.0,"remarksStatus":"Active"},{"sNo":115,"orgEmpCode":"425","payrollEmpCode":"425","employeeName":"MOHIT VERMA S/O SHER SING","designation":"-","monthlyRate":14000.0,"remarksStatus":"Active"},{"sNo":116,"orgEmpCode":"427","payrollEmpCode":"427","employeeName":"PRADEEP PANDEY","designation":"-","monthlyRate":15000.0,"remarksStatus":"Active"},{"sNo":117,"orgEmpCode":"430","payrollEmpCode":"430","employeeName":"PARVINDER SINGH","designation":"-","monthlyRate":18000.0,"remarksStatus":"Active"},{"sNo":118,"orgEmpCode":"434","payrollEmpCode":"434","employeeName":"MANOJ","designation":"-","monthlyRate":14000.0,"remarksStatus":"Active"},{"sNo":119,"orgEmpCode":"438","payrollEmpCode":"438","employeeName":"MOHIT","designation":"-","monthlyRate":18000.0,"remarksStatus":"Active"},{"sNo":120,"orgEmpCode":"441","payrollEmpCode":"441","employeeName":"PUSHKAR KASHYAP","designation":"-","monthlyRate":14000.0,"remarksStatus":"Active"},{"sNo":121,"orgEmpCode":"442","payrollEmpCode":"442","employeeName":"KASHISH","designation":"-","monthlyRate":18000.0,"remarksStatus":"Active"},{"sNo":122,"orgEmpCode":"447","payrollEmpCode":"447","employeeName":"RISHI KUMAR","designation":"-","monthlyRate":14000.0,"remarksStatus":"Active"},{"sNo":123,"orgEmpCode":"449","payrollEmpCode":"449","employeeName":"HITEN S/O RAM SUBIR","designation":"-","monthlyRate":16500.0,"remarksStatus":"Active"},{"sNo":124,"orgEmpCode":"450","payrollEmpCode":"450","employeeName":"AMIT KUMAR","designation":"-","monthlyRate":14000.0,"remarksStatus":"Active"},{"sNo":125,"orgEmpCode":"453","payrollEmpCode":"453","employeeName":"AMAN KUMAR RANA S/O SUREN","designation":"-","monthlyRate":17000.0,"remarksStatus":"Active"},{"sNo":126,"orgEmpCode":"454","payrollEmpCode":"454","employeeName":"SHIVAM KUMAR S/O VINOD KU","designation":"-","monthlyRate":17000.0,"remarksStatus":"Active"},{"sNo":127,"orgEmpCode":"456","payrollEmpCode":"456","employeeName":"MAHESH KUMAR TIWARI","designation":"-","monthlyRate":20000.0,"remarksStatus":"Active"},{"sNo":128,"orgEmpCode":"459","payrollEmpCode":"459","employeeName":"DIXON JOSEPH S/O BARNAT J","designation":"-","monthlyRate":14000.0,"remarksStatus":"Active"},{"sNo":129,"orgEmpCode":"463","payrollEmpCode":"463","employeeName":"BHARAT MATHUR S/O RUSTAM","designation":"-","monthlyRate":17500.0,"remarksStatus":"Active"},{"sNo":130,"orgEmpCode":"465","payrollEmpCode":"297","employeeName":"PAWAN SINGH GOSAI","designation":"-","monthlyRate":22500.0,"remarksStatus":"Active"},{"sNo":131,"orgEmpCode":"469","payrollEmpCode":"298","employeeName":"MONU KUMAR S/O PANCHANAND","designation":"-","monthlyRate":30000.0,"remarksStatus":"Active"},{"sNo":132,"orgEmpCode":"470","payrollEmpCode":"470","employeeName":"RADHA D/O KANHAIYA LAL","designation":"-","monthlyRate":12000.0,"remarksStatus":"Active"},{"sNo":133,"orgEmpCode":"473","payrollEmpCode":"300","employeeName":"KAPIL KUMAR S/O ANAND SIN","designation":"-","monthlyRate":25000.0,"remarksStatus":"Active"},{"sNo":134,"orgEmpCode":"474","payrollEmpCode":"474","employeeName":"RAVI KUMAR S/O VINOD KUMA","designation":"-","monthlyRate":12500.0,"remarksStatus":"Active"},{"sNo":135,"orgEmpCode":"475","payrollEmpCode":"475","employeeName":"RITESH BISHT S/O SURENDRA","designation":"-","monthlyRate":14000.0,"remarksStatus":"Active"},{"sNo":136,"orgEmpCode":"477","payrollEmpCode":"477","employeeName":"NEERAJ S/O RAM KISHAN DAS","designation":"-","monthlyRate":15000.0,"remarksStatus":"Active"},{"sNo":137,"orgEmpCode":"478","payrollEmpCode":"478","employeeName":"VIKASH S/O RAMKESH","designation":"-","monthlyRate":13000.0,"remarksStatus":"Active"},{"sNo":138,"orgEmpCode":"479","payrollEmpCode":"479","employeeName":"VIRENDRA KUMAR MISHRA","designation":"-","monthlyRate":17000.0,"remarksStatus":"Active"},{"sNo":139,"orgEmpCode":"480","payrollEmpCode":"480","employeeName":"MUSKAN D/O RAKESH KUMAR","designation":"-","monthlyRate":20500.0,"remarksStatus":"Active"},{"sNo":140,"orgEmpCode":"482","payrollEmpCode":"482","employeeName":"SONU S/O HARI CHAND","designation":"-","monthlyRate":13000.0,"remarksStatus":"Active"},{"sNo":141,"orgEmpCode":"483","payrollEmpCode":"483","employeeName":"PANKAJ KUMAR S/O BRAHAMDE","designation":"-","monthlyRate":15000.0,"remarksStatus":"Active"},{"sNo":142,"orgEmpCode":"484","payrollEmpCode":"484","employeeName":"AJAY S/O RADHEY SHYAM","designation":"-","monthlyRate":20000.0,"remarksStatus":"Active"},{"sNo":143,"orgEmpCode":"486","payrollEmpCode":"486","employeeName":"ROHIT SAXENA S/O ANIL","designation":"-","monthlyRate":14500.0,"remarksStatus":"Active"},{"sNo":144,"orgEmpCode":"487","payrollEmpCode":"487","employeeName":"UPASANA WO VIKKI","designation":"-","monthlyRate":20000.0,"remarksStatus":"Active"},{"sNo":145,"orgEmpCode":"488","payrollEmpCode":"488","employeeName":"SUDHANSHU S/O SONU","designation":"-","monthlyRate":20000.0,"remarksStatus":"Active"},{"sNo":146,"orgEmpCode":"489","payrollEmpCode":"489","employeeName":"MANISH TIWARI","designation":"-","monthlyRate":18000.0,"remarksStatus":"Active"},{"sNo":147,"orgEmpCode":"491","payrollEmpCode":"491","employeeName":"LAXMAN SINGH BHANDARI","designation":"-","monthlyRate":18000.0,"remarksStatus":"Active"},{"sNo":148,"orgEmpCode":"493","payrollEmpCode":"493","employeeName":"PANKAJ KUMAR MISHRA","designation":"-","monthlyRate":17500.0,"remarksStatus":"Active"},{"sNo":149,"orgEmpCode":"495","payrollEmpCode":"495","employeeName":"UMA SHANKER MISHRA","designation":"-","monthlyRate":18000.0,"remarksStatus":"Active"},{"sNo":150,"orgEmpCode":"497","payrollEmpCode":"497","employeeName":"AMIT SINGH TOMER S/O RAMD","designation":"-","monthlyRate":18000.0,"remarksStatus":"Active"},{"sNo":151,"orgEmpCode":"500","payrollEmpCode":"500","employeeName":"SANJANA SINGH D/O BALKISH","designation":"-","monthlyRate":15000.0,"remarksStatus":"Active"},{"sNo":152,"orgEmpCode":"501","payrollEmpCode":"501","employeeName":"KUM KUM D/O AJAY KUMAR","designation":"-","monthlyRate":22000.0,"remarksStatus":"Active"},{"sNo":153,"orgEmpCode":"502","payrollEmpCode":"502","employeeName":"NEHA KUMARI D/O BIRENDER","designation":"-","monthlyRate":15000.0,"remarksStatus":"Active"},{"sNo":154,"orgEmpCode":"503","payrollEmpCode":"503","employeeName":"ADITYA TIWARI S/O DIWAKAR","designation":"-","monthlyRate":16000.0,"remarksStatus":"Active"},{"sNo":155,"orgEmpCode":"504","payrollEmpCode":"504","employeeName":"MAMTA D/O CHANDAN SINGH","designation":"-","monthlyRate":18000.0,"remarksStatus":"Active"},{"sNo":156,"orgEmpCode":"505","payrollEmpCode":"505","employeeName":"POOJA SACHDEVA W/O PRADEE","designation":"-","monthlyRate":42000.0,"remarksStatus":"Active"},{"sNo":157,"orgEmpCode":"506","payrollEmpCode":"506","employeeName":"PRACHI SHARMA D/O ANIL SH","designation":"-","monthlyRate":18000.0,"remarksStatus":"Active"},{"sNo":158,"orgEmpCode":"507","payrollEmpCode":"507","employeeName":"SUNIL S/O HARI KISHAN","designation":"-","monthlyRate":15000.0,"remarksStatus":"Active"},{"sNo":159,"orgEmpCode":"509","payrollEmpCode":"509","employeeName":"SHIVAM VERMA S/O RAJENDRA","designation":"-","monthlyRate":15000.0,"remarksStatus":"Active"},{"sNo":160,"orgEmpCode":"510","payrollEmpCode":"510","employeeName":"VANSHIKA SHARMA D/O AMIT","designation":"-","monthlyRate":17000.0,"remarksStatus":"Active"},{"sNo":161,"orgEmpCode":"511","payrollEmpCode":"511","employeeName":"PAWAN KUMAR S/O KAMLESH K","designation":"-","monthlyRate":16000.0,"remarksStatus":"Active"},{"sNo":162,"orgEmpCode":"512","payrollEmpCode":"512","employeeName":"KRITI DEVI D/O LAL SINGH","designation":"-","monthlyRate":23500.0,"remarksStatus":"Active"},{"sNo":163,"orgEmpCode":"513","payrollEmpCode":"513","employeeName":"ANKUSH KAPOOR S/O MAHENDR","designation":"-","monthlyRate":15000.0,"remarksStatus":"Active"},{"sNo":164,"orgEmpCode":"514","payrollEmpCode":"514","employeeName":"JITENDAR KASHYAP S/O NANN","designation":"-","monthlyRate":18500.0,"remarksStatus":"Active"},{"sNo":165,"orgEmpCode":"515","payrollEmpCode":"515","employeeName":"SANDEEP KUMAR S/O SHIV BA","designation":"-","monthlyRate":19000.0,"remarksStatus":"Active"},{"sNo":166,"orgEmpCode":"516","payrollEmpCode":"516","employeeName":"JAI KUMAR S/O ASHWANI KUM","designation":"-","monthlyRate":22000.0,"remarksStatus":"Active"},{"sNo":167,"orgEmpCode":"517","payrollEmpCode":"517","employeeName":"KHUSHBU D/O JAINATH","designation":"-","monthlyRate":33600.0,"remarksStatus":"Active"},{"sNo":168,"orgEmpCode":"519","payrollEmpCode":"519","employeeName":"ANISH","designation":"-","monthlyRate":30000.0,"remarksStatus":"Active"}]

def get_current_timestamp():
    return datetime.now().strftime("%d-%m-%Y %I:%M %p")

def get_db_connection():
    conn = sqlite3.connect(DB_FILE, timeout=30, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn

def is_manager_code(code_str):
    return bool(re.search(r'[a-zA-Z]', str(code_str or '')))

def init_db():
    conn = get_db_connection()
    cursor = conn.cursor()
    
    # 1. Config Table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS config (
            id INTEGER PRIMARY KEY,
            company_name TEXT,
            company_addr TEXT,
            pf_rate REAL,
            esic_rate REAL,
            hra_rate REAL,
            pt_amount REAL
        )
    ''')
    
    # 2. Employees Master Table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS employees (
            id TEXT PRIMARY KEY,
            code TEXT,
            payroll_code TEXT DEFAULT '',
            status TEXT DEFAULT 'Active',
            name TEXT,
            designation TEXT,
            dept TEXT,
            father_or_husband_name TEXT,
            dob TEXT,
            doj TEXT,
            dol TEXT,
            uan TEXT,
            epf_member_id TEXT,
            esic_ip_number TEXT,
            bank TEXT,
            acc TEXT,
            ifsc TEXT,
            mode_of_transfer TEXT DEFAULT 'Bank Transfer',
            basic_rate REAL DEFAULT 0,
            spl_allowance REAL DEFAULT 0,
            city_allow REAL DEFAULT 0,
            deduct_short_time REAL DEFAULT 0,
            overtime_arrear REAL DEFAULT 0,
            penalty_late REAL DEFAULT 0,
            last_increment TEXT,
            created_at TEXT,
            modified_at TEXT
        )
    ''')
    
    # Auto-migration for employees table
    cursor.execute("PRAGMA table_info(employees)")
    existing_cols = [row['name'] for row in cursor.fetchall()]
    new_cols = {
        'status': "TEXT DEFAULT 'Active'",
        'payroll_code': "TEXT DEFAULT ''",
        'father_or_husband_name': "TEXT",
        'dob': "TEXT",
        'doj': "TEXT",
        'dol': "TEXT",
        'uan': "TEXT",
        'epf_member_id': "TEXT",
        'esic_ip_number': "TEXT",
        'mode_of_transfer': "TEXT DEFAULT 'Bank Transfer'",
        'city_allow': "REAL DEFAULT 0",
        'deduct_short_time': "REAL DEFAULT 0",
        'overtime_arrear': "REAL DEFAULT 0",
        'penalty_late': "REAL DEFAULT 0",
        'last_increment': "TEXT",
        'created_at': "TEXT",
        'modified_at': "TEXT",
        # Biometric-discovered employees remain hidden from Employee Master until promoted.
        'master_approved': "INTEGER NOT NULL DEFAULT 1"
    }
    for col, col_def in new_cols.items():
        if col not in existing_cols:
            cursor.execute(f"ALTER TABLE employees ADD COLUMN {col} {col_def}")

    # Independent Salary Master. Salary edits live here and never write back to employees.
    cursor.execute("SELECT 1 FROM sqlite_master WHERE type='table' AND name='salary_master'")
    salary_master_table_existed = cursor.fetchone() is not None
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS salary_master (
            employee_id TEXT PRIMARY KEY,
            emp_code TEXT DEFAULT '',
            payroll_code TEXT DEFAULT '',
            employee_name TEXT DEFAULT '',
            basic_rate REAL DEFAULT 0,
            spl_allowance REAL DEFAULT 0,
            city_allow REAL DEFAULT 0,
            modified_at TEXT
        )
    ''')
    cursor.execute("PRAGMA table_info(salary_master)")
    existing_salary_master_cols = {row['name'] for row in cursor.fetchall()}
    salary_master_migrations = {
        'emp_code': "TEXT DEFAULT ''",
        'payroll_code': "TEXT DEFAULT ''",
        'employee_name': "TEXT DEFAULT ''",
        'basic_rate': 'REAL DEFAULT 0',
        'spl_allowance': 'REAL DEFAULT 0',
        'city_allow': 'REAL DEFAULT 0',
        'modified_at': 'TEXT'
    }
    for col, col_def in salary_master_migrations.items():
        if col not in existing_salary_master_cols:
            cursor.execute(f'ALTER TABLE salary_master ADD COLUMN {col} {col_def}')

    # 3. Monthly Payroll Records
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS payroll_records (
            record_key TEXT PRIMARY KEY,
            paid_days REAL DEFAULT 31,
            advance REAL DEFAULT 0,
            basic_rate REAL,
            spl_allowance REAL,
            city_allow REAL,
            deduct_short_time REAL DEFAULT 0,
            overtime_arrear REAL DEFAULT 0,
            penalty_late REAL DEFAULT 0
        )
    ''')

    # Auto-migration for payroll_records table
    cursor.execute("PRAGMA table_info(payroll_records)")
    existing_rec_cols = [row['name'] for row in cursor.fetchall()]
    new_rec_cols = {
        'basic_rate': "REAL",
        'spl_allowance': "REAL",
        'city_allow': "REAL",
        'deduct_short_time': "REAL DEFAULT 0",
        'overtime_arrear': "REAL DEFAULT 0",
        'penalty_late': "REAL DEFAULT 0"
    }
    for col, col_def in new_rec_cols.items():
        if col not in existing_rec_cols:
            cursor.execute(f"ALTER TABLE payroll_records ADD COLUMN {col} {col_def}")

    # 4. Biometric Months Tracking Table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS biometric_months (
            month_key TEXT PRIMARY KEY,
            year INTEGER,
            month INTEGER,
            month_name TEXT,
            total_employees INTEGER,
            total_paid_days REAL,
            uploaded_at TEXT,
            file_name TEXT
        )
    ''')

    # 5. Biometric Attendance Records Table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS biometric_records (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            month_key TEXT,
            year INTEGER,
            month INTEGER,
            emp_code TEXT,
            emp_name TEXT,
            present_days REAL DEFAULT 0,
            absent_days REAL DEFAULT 0,
            weekly_off REAL DEFAULT 0,
            holiday REAL DEFAULT 0,
            paid_days REAL DEFAULT 0,
            hours_worked TEXT,
            overtime_hours TEXT,
            daily_logs TEXT,
            created_at TEXT,
            UNIQUE(month_key, emp_code)
        )
    ''')
    
    # 6. Salary Master 2 roster (independent from Employees Master)
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS salary_master2 (
            s_no INTEGER PRIMARY KEY,
            org_emp_code TEXT DEFAULT '',
            payroll_emp_code TEXT DEFAULT '',
            employee_name TEXT NOT NULL DEFAULT '',
            designation TEXT DEFAULT '',
            monthly_rate REAL DEFAULT 0,
            remarks_status TEXT DEFAULT 'Active',
            modified_at TEXT
        )
    ''')
    # Safe migration for an older/incomplete Salary Master 2 table.
    cursor.execute('PRAGMA table_info(salary_master2)')
    salary_cols = {row['name'] for row in cursor.fetchall()}
    salary_migrations = {
        # Older databases may have a salary_master2 table created before s_no
        # was added. Add it explicitly; CREATE TABLE IF NOT EXISTS does not
        # change the schema of an existing table.
        's_no': 'INTEGER',
        'org_emp_code': "TEXT DEFAULT ''",
        'payroll_emp_code': "TEXT DEFAULT ''",
        'employee_name': "TEXT NOT NULL DEFAULT ''",
        'designation': "TEXT DEFAULT ''",
        'monthly_rate': 'REAL DEFAULT 0',
        'remarks_status': "TEXT DEFAULT 'Active'",
        'modified_at': 'TEXT',
        # Set when a blank Payroll Emp Code is first saved; cleared after Employee Master add.
        'pending_employee_master': 'INTEGER NOT NULL DEFAULT 0'
    }
    for col, col_def in salary_migrations.items():
        if col not in salary_cols:
            cursor.execute(f'ALTER TABLE salary_master2 ADD COLUMN {col} {col_def}')

    # Backfill S.No. for every legacy row. Use rowid so existing employee
    # records are preserved and each row remains addressable by the API.
    cursor.execute('UPDATE salary_master2 SET s_no = rowid WHERE s_no IS NULL')

    cursor.execute('SELECT COUNT(*) FROM salary_master2')
    if cursor.fetchone()[0] == 0:
        cursor.executemany('''
            INSERT INTO salary_master2
            (s_no, org_emp_code, payroll_emp_code, employee_name, designation, monthly_rate, remarks_status)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        ''', [
            (r['sNo'], r['orgEmpCode'], r['payrollEmpCode'], r['employeeName'],
             r['designation'], r['monthlyRate'], r['remarksStatus'])
            for r in SALARY_MASTER2_SEED
        ])

    # Default Config
    cursor.execute('SELECT COUNT(*) FROM config')
    if cursor.fetchone()[0] == 0:
        cursor.execute('''
            INSERT INTO config (id, company_name, company_addr, pf_rate, esic_rate, hra_rate, pt_amount)
            VALUES (1, 'PRANAV CREATION PRIVATE LIMITED', 'Sector 62, Electronic City, Noida, Uttar Pradesh - 201301', 12.0, 0.75, 40.0, 200.0)
        ''')

    # Default Employees
    cursor.execute('SELECT COUNT(*) FROM employees')
    if cursor.fetchone()[0] == 0:
        now_str = get_current_timestamp()
        default_emps = [
            ("EMP-101", "101", "101", "Active", "JAGAN NATH YADAV", "Helper", "Operations", "", "", "", "", "", "", "", "SBI", "30291029381", "SBIN0001234", "Bank Transfer", 30345, 0, 0, 0, 0, 0, "Jul-16", now_str, now_str),
            ("EMP-103", "103", "103", "Active", "JITENDRA KUMAR", "Helper", "Operations", "", "", "", "", "", "", "", "HDFC Bank", "50100293841", "HDFC0000412", "Bank Transfer", 20400, 0, 0, 0, 0, 0, "", now_str, now_str),
            ("EMP-104", "104", "104", "Active", "KRISHNA KUMAR", "Helper", "Operations", "", "", "", "", "", "", "", "ICICI Bank", "00120192837", "ICIC0000182", "Bank Transfer", 22500, 0, 0, 0, 0, 0, "", now_str, now_str),
            ("EMP-106", "106", "106", "Active", "ANGAREJ", "Helper", "Operations", "", "", "", "", "", "", "", "Axis Bank", "91802019283", "UTIB0000982", "Bank Transfer", 22500, 0, 0, 0, 0, 0, "Jul-16", now_str, now_str),
            ("EMP-107", "107", "107", "Active", "MANOJ KUMAR JPR", "Helper", "Operations", "", "", "", "", "", "", "", "Kotak Bank", "48102938102", "KKBK0000521", "Bank Transfer", 20400, 0, 0, 0, 0, 0, "", now_str, now_str),
            ("EMP-109", "109", "109", "Active", "GOURISHANKAR (Dilip Sharma)", "Helper", "Operations", "", "", "", "", "", "", "", "SBI", "48102938109", "SBIN0001234", "Bank Transfer", 20400, 0, 0, 0, 0, 0, "Jul-16", now_str, now_str),
            ("EMP-110", "110", "110", "Active", "RAM PARVESH SINGH", "Helper", "Operations", "", "", "", "", "", "", "", "PNB", "48102938110", "PUNB0001234", "Bank Transfer", 22500, 0, 0, 0, 0, 0, "Jul-16", now_str, now_str)
        ]
        cursor.executemany('''
            INSERT INTO employees (id, code, payroll_code, status, name, designation, dept, father_or_husband_name, dob, doj, dol, uan, epf_member_id, esic_ip_number, bank, acc, ifsc, mode_of_transfer, basic_rate, spl_allowance, city_allow, deduct_short_time, overtime_arrear, penalty_late, last_increment, created_at, modified_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', default_emps)

    # First-creation backfill only: do not undo deliberate Salary Master deletions
    # or reset independent salary values when the server restarts.
    if not salary_master_table_existed:
        cursor.execute('''
            INSERT OR IGNORE INTO salary_master
                (employee_id, emp_code, payroll_code, employee_name, basic_rate, spl_allowance, city_allow, modified_at)
            SELECT id, COALESCE(code, ''), COALESCE(payroll_code, ''), COALESCE(name, ''),
                   COALESCE(basic_rate, 0), COALESCE(spl_allowance, 0), COALESCE(city_allow, 0), modified_at
            FROM employees WHERE COALESCE(master_approved, 1) = 1
        ''')

    conn.commit()
    conn.close()

init_db()

@app.route('/')
def index():
    # Always load the template that ships with this project folder.
    template_file = os.path.join(APP_DIR, 'templates', 'index.html')
    root_file = os.path.join(APP_DIR, 'index.html')
    if os.path.isfile(template_file):
        return render_template('index.html')
    if os.path.isfile(root_file):
        return send_file(root_file, max_age=0)
    return jsonify({'error': 'index.html was not found. Put it in templates/index.html or beside app.py.'}), 500

@app.route('/static/script.js')
@app.route('/script.js')
def serve_script():
    # Serve the file under this app's static folder (the path provided in the ZIP).
    # Disable caching while developing so an older local script cannot mask updates.
    static_file = os.path.join(APP_DIR, 'static', 'script.js')
    root_file = os.path.join(APP_DIR, 'script.js')
    if os.path.isfile(static_file):
        response = send_file(static_file, mimetype='application/javascript', max_age=0)
    elif os.path.isfile(root_file):
        response = send_file(root_file, mimetype='application/javascript', max_age=0)
    else:
        return jsonify({'error': 'script.js was not found. Put it in static/script.js or beside app.py.'}), 404
    response.headers['Cache-Control'] = 'no-store, max-age=0'
    response.headers['X-Payroll-Script-Version'] = 'salary-master-independent-payroll-20261012'
    return response

@app.route('/api/health', methods=['GET'])
def api_health():
    """Small diagnostic endpoint to confirm the Flask app and DB are reachable."""
    conn = None
    try:
        conn = get_db_connection()
        conn.execute('SELECT 1').fetchone()
        return jsonify({'status': 'ok', 'database': os.path.basename(DB_FILE)})
    except Exception as exc:
        app.logger.exception('Health check failed')
        return jsonify({'status': 'error', 'message': str(exc)}), 500
    finally:
        if conn is not None:
            conn.close()

# API: Get Complete State
@app.route('/api/data', methods=['GET'])
def get_all_data():
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute('SELECT * FROM config WHERE id = 1')
        cfg_row = cursor.fetchone()
        config = {
            'companyName': cfg_row['company_name'] if cfg_row else 'PRANAV CREATION PRIVATE LIMITED',
            'companyAddr': cfg_row['company_addr'] if cfg_row else 'Sector 62, Noida',
            'pfRate': cfg_row['pf_rate'] if cfg_row else 12.0,
            'esicRate': cfg_row['esic_rate'] if cfg_row else 0.75,
            'hraRate': cfg_row['hra_rate'] if cfg_row else 40.0,
            'ptAmount': cfg_row['pt_amount'] if cfg_row else 200.0
        }
        
        cursor.execute('SELECT * FROM employees ORDER BY CAST(code AS INTEGER), code')
        emp_rows = cursor.fetchall()
        employees = [{
            'id': r['id'],
            'code': str(r['code'] or '').upper(),
            'payrollCode': str(r['payroll_code'] if 'payroll_code' in r.keys() else '').upper(),
            'status': r['status'] or 'Active',
            'name': r['name'],
            'designation': r['designation'] or '',
            'dept': r['dept'] or 'Operations',
            'fatherOrHusbandName': r['father_or_husband_name'] or '',
            'dob': r['dob'] or '',
            'doj': r['doj'] or '',
            'dol': r['dol'] or '',
            'uan': r['uan'] or '',
            'epfMemberId': r['epf_member_id'] or '',
            'esicIpNumber': r['esic_ip_number'] or '',
            'bank': r['bank'] or '',
            'acc': r['acc'] or '',
            'ifsc': r['ifsc'] or '',
            'modeOfTransfer': r['mode_of_transfer'] or 'Bank Transfer',
            'basicRate': r['basic_rate'] or 0,
            'splAllowance': r['spl_allowance'] or 0,
            'cityAllow': r['city_allow'] or 0,
            'deductShortTime': r['deduct_short_time'] or 0,
            'overtimeArrear': r['overtime_arrear'] or 0,
            'penaltyLate': r['penalty_late'] or 0,
            'lastIncrement': r['last_increment'] or '',
            'createdAt': r['created_at'] or '',
            'modifiedAt': r['modified_at'] or '',
            'masterApproved': bool(r['master_approved']) if 'master_approved' in r.keys() else True
        } for r in emp_rows]
        
        cursor.execute('SELECT * FROM payroll_records')
        rec_rows = cursor.fetchall()
        records = {
            r['record_key']: {
                'paidDays': r['paid_days'],
                'advance': r['advance'] or 0,
                'basicRate': r['basic_rate'],
                'splAllowance': r['spl_allowance'],
                'cityAllow': r['city_allow'],
                'deductShortTime': r['deduct_short_time'] or 0,
                'overtimeArrear': r['overtime_arrear'] or 0,
                'penaltyLate': r['penalty_late'] or 0
            } for r in rec_rows
        }

        biometric_months = []
        try:
            cursor.execute('SELECT * FROM biometric_months ORDER BY year DESC, month DESC')
            bio_m_rows = cursor.fetchall()
            biometric_months = [{
                'monthKey': r['month_key'],
                'year': r['year'],
                'month': r['month'],
                'monthName': r['month_name'],
                'totalEmployees': r['total_employees'],
                'totalPaidDays': r['total_paid_days'],
                'uploadedAt': r['uploaded_at'],
                'fileName': r['file_name']
            } for r in bio_m_rows]
        except Exception as e:
            pass

        biometric_records = {}
        try:
            cursor.execute('SELECT * FROM biometric_records')
            bio_rec_rows = cursor.fetchall()
            for r in bio_rec_rows:
                m_key = r['month_key']
                if m_key not in biometric_records:
                    biometric_records[m_key] = []
                
                parsed_logs = []
                if r['daily_logs']:
                    try:
                        parsed_logs = json.loads(r['daily_logs'])
                    except Exception:
                        parsed_logs = []

                biometric_records[m_key].append({
                    'code': r['emp_code'],
                    'name': r['emp_name'],
                    'presentDays': r['present_days'],
                    'absentDays': r['absent_days'],
                    'weeklyOff': r['weekly_off'],
                    'holiday': r['holiday'],
                    'paidDays': r['paid_days'],
                    'hoursWorked': r['hours_worked'] or '00:00',
                    'overtimeHours': r['overtime_hours'] or '00:00',
                    'dailyLogs': parsed_logs
                })
        except Exception as e:
            pass

        cursor.execute('SELECT * FROM salary_master ORDER BY CAST(emp_code AS INTEGER), emp_code, employee_name')
        salary_master = [{
            'employeeId': r['employee_id'],
            'empCode': r['emp_code'] or '',
            'payrollCode': r['payroll_code'] or '',
            'name': r['employee_name'] or '',
            'basicRate': r['basic_rate'] or 0,
            'splAllowance': r['spl_allowance'] or 0,
            'cityAllow': r['city_allow'] or 0,
            'modifiedAt': r['modified_at'] or ''
        } for r in cursor.fetchall()]

        cursor.execute('SELECT * FROM salary_master2 ORDER BY s_no')
        salary_master2 = [{
            'sNo': r['s_no'],
            'orgEmpCode': r['org_emp_code'] or '',
            'payrollEmpCode': r['payroll_emp_code'] or '',
            'employeeName': r['employee_name'] or '',
            'designation': r['designation'] or '',
            'monthlyRate': r['monthly_rate'] or 0,
            'remarksStatus': r['remarks_status'] or '',
            'modifiedAt': r['modified_at'] or '',
            'pendingEmployeeMaster': bool(r['pending_employee_master']) if 'pending_employee_master' in r.keys() else False
        } for r in cursor.fetchall()]

        return jsonify({
            'config': config,
            'employees': employees,
            'records': records,
            'biometricMonths': biometric_months,
            'biometricRecords': biometric_records,
            'salaryMaster': salary_master,
            'salaryMaster2': salary_master2
        })
    except Exception as ex:
        return jsonify({'error': str(ex)}), 500
    finally:
        conn.close()

# Independent Salary Master endpoints: these modify only salary_master, never employees.
@app.route('/api/salary-master/<path:employee_id>', methods=['PUT'])
def update_salary_master(employee_id):
    data = request.get_json(silent=True) or {}
    field_map = {'basicRate': 'basic_rate', 'splAllowance': 'spl_allowance', 'cityAllow': 'city_allow'}
    updates = {}
    for key, column in field_map.items():
        if key not in data:
            continue
        try:
            value = float(data[key] or 0)
        except (TypeError, ValueError):
            return jsonify({'status': 'error', 'message': f'Invalid value for {key}.'}), 400
        if value < 0:
            return jsonify({'status': 'error', 'message': 'Salary amounts cannot be negative.'}), 400
        updates[column] = value
    if not updates:
        return jsonify({'status': 'error', 'message': 'No salary fields were supplied.'}), 400
    conn = get_db_connection()
    try:
        cursor = conn.cursor()
        cursor.execute('SELECT employee_id FROM salary_master WHERE employee_id = ?', (employee_id,))
        if not cursor.fetchone():
            return jsonify({'status': 'error', 'message': 'Salary Master record not found.'}), 404
        updates['modified_at'] = get_current_timestamp()
        assignments = ', '.join(f'{col} = ?' for col in updates)
        cursor.execute(f'UPDATE salary_master SET {assignments} WHERE employee_id = ?', (*updates.values(), employee_id))
        conn.commit()
        return jsonify({'status': 'success', 'message': 'Salary Master updated. Employee Master data was not changed.'})
    except Exception as exc:
        conn.rollback()
        app.logger.exception('Salary Master update failed')
        return jsonify({'status': 'error', 'message': str(exc)}), 500
    finally:
        conn.close()


@app.route('/api/salary-master/bulk_update', methods=['POST'])
def bulk_update_salary_master():
    data = request.get_json(silent=True) or {}
    ids = data.get('employeeIds', [])
    if not isinstance(ids, list) or not ids:
        return jsonify({'status': 'error', 'message': 'Please select at least one Salary Master employee.'}), 400
    employee_ids = list(dict.fromkeys(str(value).strip() for value in ids if str(value).strip()))
    if not employee_ids:
        return jsonify({'status': 'error', 'message': 'No valid employees selected.'}), 400
    field_map = {'basicRate': 'basic_rate', 'splAllowance': 'spl_allowance', 'cityAllow': 'city_allow'}
    supplied = {}
    for key, column in field_map.items():
        raw = data.get(key, None)
        if raw is None or str(raw).strip() == '':
            continue
        try:
            value = float(raw)
        except (TypeError, ValueError):
            return jsonify({'status': 'error', 'message': f'Invalid value for {key}.'}), 400
        if value < 0:
            return jsonify({'status': 'error', 'message': 'Salary values/adjustments cannot be negative.'}), 400
        supplied[column] = value
    if not supplied:
        return jsonify({'status': 'error', 'message': 'Enter at least one salary field to update.'}), 400
    mode = str(data.get('mode', 'set')).strip().lower()
    if mode not in {'set', 'increase', 'decrease', 'percent_increase', 'percent_decrease'}:
        return jsonify({'status': 'error', 'message': 'Invalid bulk edit mode.'}), 400
    conn = get_db_connection()
    try:
        cursor = conn.cursor()
        placeholders = ','.join('?' for _ in employee_ids)
        cursor.execute(f'SELECT employee_id, basic_rate, spl_allowance, city_allow FROM salary_master WHERE employee_id IN ({placeholders})', employee_ids)
        rows = cursor.fetchall()
        if not rows:
            return jsonify({'status': 'error', 'message': 'No selected Salary Master records were found.'}), 404
        now = get_current_timestamp()
        updated = 0
        for row in rows:
            new_values = {}
            for column, value in supplied.items():
                old = float(row[column] or 0)
                if mode == 'set': new = value
                elif mode == 'increase': new = old + value
                elif mode == 'decrease': new = max(0, old - value)
                elif mode == 'percent_increase': new = old * (1 + value / 100)
                else: new = max(0, old * (1 - value / 100))
                new_values[column] = round(new, 2)
            assignments = ', '.join(f'{column} = ?' for column in new_values)
            cursor.execute(f'UPDATE salary_master SET {assignments}, modified_at = ? WHERE employee_id = ?', (*new_values.values(), now, row['employee_id']))
            updated += cursor.rowcount
        conn.commit()
        return jsonify({'status': 'success', 'updatedCount': updated, 'message': 'Salary Master updated only; Employee Master data unchanged.'})
    except Exception as exc:
        conn.rollback()
        app.logger.exception('Salary Master bulk update failed')
        return jsonify({'status': 'error', 'message': str(exc)}), 500
    finally:
        conn.close()


@app.route('/api/salary-master/bulk_delete', methods=['POST'])
def bulk_delete_salary_master():
    data = request.get_json(silent=True) or {}
    ids = data.get('employeeIds', [])
    if not isinstance(ids, list) or not ids:
        return jsonify({'status': 'error', 'message': 'No Salary Master rows selected.'}), 400
    employee_ids = list(dict.fromkeys(str(value).strip() for value in ids if str(value).strip()))
    if not employee_ids:
        return jsonify({'status': 'error', 'message': 'No valid Salary Master rows selected.'}), 400
    conn = get_db_connection()
    try:
        cursor = conn.cursor()
        placeholders = ','.join('?' for _ in employee_ids)
        cursor.execute(f'DELETE FROM salary_master WHERE employee_id IN ({placeholders})', employee_ids)
        deleted = cursor.rowcount
        conn.commit()
        return jsonify({'status': 'success', 'deletedCount': deleted, 'message': 'Salary Master rows deleted; Employee Master and payroll records unchanged.'})
    except Exception as exc:
        conn.rollback()
        app.logger.exception('Salary Master bulk delete failed')
        return jsonify({'status': 'error', 'message': str(exc)}), 500
    finally:
        conn.close()


# API: Bulk delete selected Salary Master 2 roster rows.
# This removes only Salary Master 2 entries; Employee Master and payroll history are untouched.
@app.route('/api/salary-master2/bulk_delete', methods=['POST'])
def bulk_delete_salary_master2():
    data = request.get_json(silent=True) or {}
    raw_s_nos = data.get('sNos', data.get('ids', []))
    if not isinstance(raw_s_nos, list) or not raw_s_nos:
        return jsonify({'status': 'error', 'message': 'No Salary Master 2 rows selected.'}), 400

    try:
        s_nos = sorted({int(value) for value in raw_s_nos if str(value).strip()})
    except (TypeError, ValueError):
        return jsonify({'status': 'error', 'message': 'Invalid Salary Master 2 row selection.'}), 400
    s_nos = [value for value in s_nos if value > 0]
    if not s_nos:
        return jsonify({'status': 'error', 'message': 'No valid Salary Master 2 rows selected.'}), 400

    conn = get_db_connection()
    try:
        cursor = conn.cursor()
        placeholders = ','.join('?' for _ in s_nos)
        cursor.execute(f'DELETE FROM salary_master2 WHERE s_no IN ({placeholders})', s_nos)
        deleted_count = cursor.rowcount
        conn.commit()
        return jsonify({'status': 'success', 'deletedCount': deleted_count})
    except Exception as exc:
        conn.rollback()
        app.logger.exception('Salary Master 2 bulk delete failed')
        return jsonify({'status': 'error', 'message': str(exc)}), 500
    finally:
        conn.close()


# API: Insert a new Salary Master 2 roster row from the Add button.
@app.route('/api/salary-master2', methods=['POST'])
def create_salary_master2():
    data = request.get_json(silent=True) or {}
    org_code = str(data.get('orgEmpCode', '') or '').strip().upper()
    payroll_code = str(data.get('payrollEmpCode', '') or '').strip().upper()
    employee_name = str(data.get('employeeName', '') or '').strip().upper()
    designation = str(data.get('designation', '') or '').strip().upper()
    remarks = str(data.get('remarksStatus', '') or 'Active').strip()
    if not org_code:
        return jsonify({'status': 'error', 'message': 'Please enter the ORG Emp Code.'}), 400
    if not employee_name:
        return jsonify({'status': 'error', 'message': 'Please enter the Employee Name.'}), 400
    try:
        monthly_rate = float(data.get('monthlyRate', 0) or 0)
        if monthly_rate < 0:
            return jsonify({'status': 'error', 'message': 'Monthly rate cannot be negative.'}), 400
    except (TypeError, ValueError):
        return jsonify({'status': 'error', 'message': 'Please enter a valid monthly rate.'}), 400

    conn = get_db_connection()
    try:
        cursor = conn.cursor()
        cursor.execute('SELECT s_no FROM salary_master2 WHERE UPPER(TRIM(org_emp_code)) = ? LIMIT 1', (org_code,))
        duplicate = cursor.fetchone()
        if duplicate:
            return jsonify({'status': 'error', 'message': f'ORG Emp Code {org_code} already exists in Salary Master 2 (S.No. {duplicate["s_no"]}). Use Edit instead.'}), 409
        cursor.execute('SELECT COALESCE(MAX(s_no), 0) + 1 AS next_s_no FROM salary_master2')
        next_s_no = int(cursor.fetchone()['next_s_no'] or 1)
        now = get_current_timestamp()
        # If a payroll code was entered while inserting, place the row into
        # Pending Employees; otherwise allow assigning it later in the row.
        pending_flag = 1 if payroll_code else 0
        cursor.execute('''
            INSERT INTO salary_master2
                (s_no, org_emp_code, payroll_emp_code, employee_name, designation, monthly_rate, remarks_status, modified_at, pending_employee_master)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (next_s_no, org_code, payroll_code, employee_name, designation, monthly_rate, remarks, now, pending_flag))
        conn.commit()
        return jsonify({'status': 'success', 'sNo': next_s_no, 'pendingEmployeeMaster': bool(pending_flag), 'message': 'Salary Master 2 record created.'}), 201
    except Exception as exc:
        conn.rollback()
        app.logger.exception('Salary Master 2 create failed')
        return jsonify({'status': 'error', 'message': str(exc)}), 500
    finally:
        conn.close()


# API: Save a Payroll Emp Code, or edit a Salary Master 2 row from its Edit modal.
@app.route('/api/salary-master2/<int:s_no>', methods=['POST', 'PUT'])
def update_salary_master2(s_no):
    data = request.get_json(silent=True) or {}
    is_full_edit = request.method == 'PUT' and bool(data.get('editExisting'))
    payroll_code = str(data.get('payrollEmpCode', '') or '').strip().upper()
    remarks = str(data.get('remarksStatus', '') or '').strip()
    try:
        monthly_rate = float(data.get('monthlyRate', 0) or 0)
        if monthly_rate < 0:
            return jsonify({'status': 'error', 'message': 'Monthly rate cannot be negative.'}), 400
    except (TypeError, ValueError):
        return jsonify({'status': 'error', 'message': 'Please enter a valid monthly rate.'}), 400

    if is_full_edit:
        org_code = str(data.get('orgEmpCode', '') or '').strip().upper()
        employee_name = str(data.get('employeeName', '') or '').strip().upper()
        designation = str(data.get('designation', '') or '').strip().upper()
        if not org_code:
            return jsonify({'status': 'error', 'message': 'Please enter the ORG Emp Code.'}), 400
        if not employee_name:
            return jsonify({'status': 'error', 'message': 'Please enter the Employee Name.'}), 400
    elif not payroll_code:
        return jsonify({'status': 'error', 'message': 'Please fill the Payroll Emp Code for this employee.'}), 400

    conn = get_db_connection()
    try:
        cursor = conn.cursor()
        cursor.execute('SELECT * FROM salary_master2 WHERE s_no = ?', (s_no,))
        existing = cursor.fetchone()
        if not existing:
            return jsonify({'status': 'error', 'message': 'Salary Master 2 row was not found.'}), 404

        now = get_current_timestamp()
        if is_full_edit:
            cursor.execute('SELECT s_no FROM salary_master2 WHERE UPPER(TRIM(org_emp_code)) = ? AND s_no <> ? LIMIT 1', (org_code, s_no))
            duplicate = cursor.fetchone()
            if duplicate:
                return jsonify({'status': 'error', 'message': f'ORG Emp Code {org_code} already exists in Salary Master 2 (S.No. {duplicate["s_no"]}).'}), 409
            old_payroll_code = str(existing['payroll_emp_code'] or '').strip()
            pending_flag = bool(existing['pending_employee_master']) if 'pending_employee_master' in existing.keys() else False
            if not old_payroll_code and payroll_code:
                pending_flag = True
            if not payroll_code:
                pending_flag = False
            cursor.execute('''
                UPDATE salary_master2
                SET org_emp_code = ?, payroll_emp_code = ?, employee_name = ?, designation = ?,
                    monthly_rate = ?, remarks_status = ?, pending_employee_master = ?, modified_at = ?
                WHERE s_no = ?
            ''', (org_code, payroll_code, employee_name, designation, monthly_rate, remarks, int(pending_flag), now, s_no))
        else:
            if str(existing['payroll_emp_code'] or '').strip():
                return jsonify({'status': 'error', 'message': 'This employee already has a Payroll Emp Code; use Edit to make changes.'}), 409
            cursor.execute('''
                UPDATE salary_master2
                SET payroll_emp_code = ?, monthly_rate = ?, remarks_status = ?,
                    pending_employee_master = 1, modified_at = ?
                WHERE s_no = ?
            ''', (payroll_code, monthly_rate, remarks, now, s_no))
        conn.commit()
        return jsonify({'status': 'success', 'sNo': s_no, 'pendingEmployeeMaster': bool(pending_flag) if is_full_edit else True})
    except Exception as exc:
        conn.rollback()
        app.logger.exception('Salary Master 2 save/edit failed')
        return jsonify({'status': 'error', 'message': str(exc)}), 500
    finally:
        conn.close()

# API: Add a Salary Master 2 saved employee to Employee Master after user approval.
@app.route('/api/salary-master2/<int:s_no>/add-to-employee-master', methods=['POST'])
def add_salary_master2_to_employee_master(s_no):
    conn = get_db_connection()
    try:
        cursor = conn.cursor()
        cursor.execute('SELECT * FROM salary_master2 WHERE s_no = ?', (s_no,))
        roster_row = cursor.fetchone()
        if not roster_row:
            return jsonify({'status': 'error', 'message': 'Salary Master 2 row was not found.'}), 404

        payroll_code = str(roster_row['payroll_emp_code'] or '').strip().upper()
        org_code = str(roster_row['org_emp_code'] or '').strip().upper()
        employee_name = str(roster_row['employee_name'] or '').strip().upper()
        designation = str(roster_row['designation'] or 'STAFF').strip().upper()
        if not payroll_code:
            return jsonify({'status': 'error', 'message': 'Please fill the Payroll Emp Code for this employee.'}), 400
        pending_flag = bool(roster_row['pending_employee_master']) if 'pending_employee_master' in roster_row.keys() else False
        if not pending_flag:
            return jsonify({'status': 'error', 'message': 'This employee is not waiting in Pending Employees.'}), 409
        if not org_code or not employee_name:
            return jsonify({'status': 'error', 'message': 'ORG Emp Code and Employee Name are required to add this employee.'}), 400

        now = get_current_timestamp()
        remark_upper = str(roster_row['remarks_status'] or '').upper()
        employee_status = 'INACTIVE' if 'INACTIVE' in remark_upper else ('LEFT' if 'LEFT' in remark_upper else 'ACTIVE')
        monthly_rate = float(roster_row['monthly_rate'] or 0)

        cursor.execute('SELECT id FROM employees WHERE UPPER(TRIM(code)) = ? ORDER BY created_at ASC LIMIT 1', (org_code,))
        existing_emp = cursor.fetchone()
        if existing_emp:
            emp_id = existing_emp['id']
            cursor.execute('''
                UPDATE employees
                SET payroll_code = ?, name = ?, designation = ?, status = ?, basic_rate = ?,
                    master_approved = 1, modified_at = ?
                WHERE id = ?
            ''', (payroll_code, employee_name, designation, employee_status, monthly_rate, now, emp_id))
        else:
            emp_id = f"EMP-{int(time.time()*1000)}-{s_no}"
            cursor.execute('''
                INSERT INTO employees
                (id, code, payroll_code, status, name, designation, dept, mode_of_transfer,
                 basic_rate, spl_allowance, city_allow, master_approved, created_at, modified_at)
                VALUES (?, ?, ?, ?, ?, ?, 'OPERATIONS', 'CASH', ?, 0, 0, 1, ?, ?)
            ''', (emp_id, org_code, payroll_code, employee_status, employee_name, designation, monthly_rate, now, now))

        cursor.execute('UPDATE salary_master2 SET pending_employee_master = 0, modified_at = ? WHERE s_no = ?', (now, s_no))
        cursor.execute('''
            INSERT OR IGNORE INTO salary_master
                (employee_id, emp_code, payroll_code, employee_name, basic_rate, spl_allowance, city_allow, modified_at)
            SELECT id, COALESCE(code, ''), COALESCE(payroll_code, ''), COALESCE(name, ''),
                   COALESCE(basic_rate, 0), COALESCE(spl_allowance, 0), COALESCE(city_allow, 0), modified_at
            FROM employees WHERE id = ? AND COALESCE(master_approved, 1) = 1
        ''', (emp_id,))
        cursor.execute('UPDATE salary_master SET emp_code = ?, payroll_code = ?, employee_name = ? WHERE employee_id = ?',
                       (org_code, payroll_code, employee_name, emp_id))
        conn.commit()
        return jsonify({'status': 'success', 'message': 'Employee added to Employee Master.', 'employeeId': emp_id, 'sNo': s_no})
    except Exception as exc:
        conn.rollback()
        app.logger.exception('Add pending Salary Master 2 employee failed')
        return jsonify({'status': 'error', 'message': str(exc)}), 500
    finally:
        conn.close()

# API: Save/Update Employee Master
@app.route('/api/employee', methods=['POST'])
def save_employee():
    data = request.json or {}
    conn = get_db_connection()
    cursor = conn.cursor()
    emp_id = data.get('id', f"EMP-{int(time.time()*1000)}")
    payroll_code = str(data.get('payrollCode', '')).strip().upper()

    cursor.execute("SELECT created_at FROM employees WHERE id = ?", (emp_id,))
    existing = cursor.fetchone()
    current_time = get_current_timestamp()

    if existing and existing['created_at']:
        created_at = existing['created_at']
        modified_at = current_time
    else:
        created_at = current_time
        modified_at = current_time

    cursor.execute('''
        INSERT INTO employees (id, code, payroll_code, status, name, designation, dept, father_or_husband_name, dob, doj, dol, uan, epf_member_id, esic_ip_number, bank, acc, ifsc, mode_of_transfer, basic_rate, spl_allowance, city_allow, deduct_short_time, overtime_arrear, penalty_late, last_increment, created_at, modified_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(id) DO UPDATE SET
            code=excluded.code,
            payroll_code=excluded.payroll_code,
            status=excluded.status,
            name=excluded.name,
            designation=excluded.designation,
            dept=excluded.dept,
            father_or_husband_name=excluded.father_or_husband_name,
            dob=excluded.dob,
            doj=excluded.doj,
            dol=excluded.dol,
            uan=excluded.uan,
            epf_member_id=excluded.epf_member_id,
            esic_ip_number=excluded.esic_ip_number,
            bank=excluded.bank,
            acc=excluded.acc,
            ifsc=excluded.ifsc,
            mode_of_transfer=excluded.mode_of_transfer,
            basic_rate=excluded.basic_rate,
            spl_allowance=excluded.spl_allowance,
            city_allow=excluded.city_allow,
            deduct_short_time=excluded.deduct_short_time,
            overtime_arrear=excluded.overtime_arrear,
            penalty_late=excluded.penalty_late,
            last_increment=excluded.last_increment,
            modified_at=excluded.modified_at
    ''', (
        emp_id,
        str(data.get('code', '')).strip().upper(),
        payroll_code,
        str(data.get('status', 'ACTIVE')).strip().upper(),
        str(data.get('name', '')).strip().upper(),
        str(data.get('designation', '')).strip().upper(),
        str(data.get('dept', 'OPERATIONS')).strip().upper(),
        str(data.get('fatherOrHusbandName', '')).strip().upper(),
        data.get('dob', ''),
        data.get('doj', ''),
        data.get('dol', ''),
        str(data.get('uan', '')).strip().upper(),
        str(data.get('epfMemberId', '')).strip().upper(),
        str(data.get('esicIpNumber', '')).strip().upper(),
        str(data.get('bank', '')).strip().upper(),
        str(data.get('acc', '')).strip().upper(),
        str(data.get('ifsc', '')).strip().upper(),
        str(data.get('modeOfTransfer', 'CASH')).strip().upper(),
        float(data.get('basicRate', 0) or 0),
        float(data.get('splAllowance', 0) or 0),
        float(data.get('cityAllow', 0) or 0),
        float(data.get('deductShortTime', 0) or 0),
        float(data.get('overtimeArrear', 0) or 0),
        float(data.get('penaltyLate', 0) or 0),
        str(data.get('lastIncrement', '')).strip().upper(),
        created_at,
        modified_at
    ))
    # Add new employee to Salary Master, but never update existing salary amounts from Employee Master edits.
    emp_code_value = str(data.get('code', '')).strip().upper()
    emp_name_value = str(data.get('name', '')).strip().upper()
    cursor.execute('''
        INSERT OR IGNORE INTO salary_master
            (employee_id, emp_code, payroll_code, employee_name, basic_rate, spl_allowance, city_allow, modified_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    ''', (emp_id, emp_code_value, payroll_code, emp_name_value,
          float(data.get('basicRate', 0) or 0), float(data.get('splAllowance', 0) or 0),
          float(data.get('cityAllow', 0) or 0), modified_at))
    cursor.execute('UPDATE salary_master SET emp_code = ?, payroll_code = ?, employee_name = ? WHERE employee_id = ?',
                   (emp_code_value, payroll_code, emp_name_value, emp_id))
    conn.commit()
    conn.close()
    return jsonify({'status': 'success', 'createdAt': created_at, 'modifiedAt': modified_at})

@app.route('/api/biometric/save', methods=['POST'])
def save_biometric_data():
    payload = request.json or {}
    month_key = payload.get('monthKey')
    year = int(payload.get('year', 2026))
    month = int(payload.get('month', 8))
    month_name = payload.get('monthName', 'August')
    file_name = payload.get('fileName', 'Biometric_Upload.xlsx')
    employees_data = payload.get('employees', [])

    if not month_key or not employees_data:
        return jsonify({'status': 'error', 'message': 'Invalid biometric data payload'}), 400

    conn = get_db_connection()
    cursor = conn.cursor()
    current_time = get_current_timestamp()

    total_paid_days = sum(float(e.get('paidDays', 0) or 0) for e in employees_data)

    cursor.execute('''
        INSERT INTO biometric_months (month_key, year, month, month_name, total_employees, total_paid_days, uploaded_at, file_name)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(month_key) DO UPDATE SET
            year=excluded.year,
            month=excluded.month,
            month_name=excluded.month_name,
            total_employees=excluded.total_employees,
            total_paid_days=excluded.total_paid_days,
            uploaded_at=excluded.uploaded_at,
            file_name=excluded.file_name
    ''', (month_key, year, month, month_name, len(employees_data), total_paid_days, current_time, file_name))

    for e_idx, e in enumerate(employees_data):
        code = str(e.get('code', '')).strip()
        name = str(e.get('name', '')).strip()
        present = float(e.get('presentDays', 0) or 0)
        absent = float(e.get('absentDays', 0) or 0)
        weekly_off = float(e.get('weeklyOff', 0) or 0)
        holiday = float(e.get('holiday', 0) or 0)
        paid_days = float(e.get('paidDays', 0) or 0)
        hours_worked = str(e.get('hoursWorked', '00:00'))
        overtime_hours = str(e.get('overtimeHours', '00:00'))
        daily_logs_json = json.dumps(e.get('dailyLogs', []))

        cursor.execute('''
            INSERT INTO biometric_records (month_key, year, month, emp_code, emp_name, present_days, absent_days, weekly_off, holiday, paid_days, hours_worked, overtime_hours, daily_logs, created_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(month_key, emp_code) DO UPDATE SET
                emp_name=excluded.emp_name,
                present_days=excluded.present_days,
                absent_days=excluded.absent_days,
                weekly_off=excluded.weekly_off,
                holiday=excluded.holiday,
                paid_days=excluded.paid_days,
                hours_worked=excluded.hours_worked,
                overtime_hours=excluded.overtime_hours,
                daily_logs=excluded.daily_logs,
                created_at=excluded.created_at
        ''', (month_key, year, month, code, name, present, absent, weekly_off, holiday, paid_days, hours_worked, overtime_hours, daily_logs_json, current_time))

        # Salary Master 2 is the first lookup source for biometric-discovered employee codes.
        # Add a roster row only when this ORG Emp Code is missing from Salary Master 2.
        cursor.execute('SELECT s_no, employee_name, designation FROM salary_master2 WHERE UPPER(TRIM(org_emp_code)) = ? LIMIT 1', (code.upper(),))
        sm2_row = cursor.fetchone()
        if not sm2_row:
            cursor.execute('SELECT COALESCE(MAX(s_no), 0) + 1 AS next_s_no FROM salary_master2')
            next_s_no = int(cursor.fetchone()['next_s_no'] or 1)
            desig = 'MANAGER' if is_manager_code(code) else 'STAFF'
            cursor.execute('''
                INSERT INTO salary_master2
                (s_no, org_emp_code, payroll_emp_code, employee_name, designation, monthly_rate, remarks_status, pending_employee_master, modified_at)
                VALUES (?, ?, '', ?, ?, 0, 'NEW EMPLOYEE - PAYROLL CODE REQUIRED', 0, ?)
            ''', (next_s_no, code.upper(), name.upper(), desig, current_time))
        elif not str(sm2_row['employee_name'] or '').strip():
            desig = 'MANAGER' if is_manager_code(code) else 'STAFF'
            cursor.execute("UPDATE salary_master2 SET employee_name = ?, designation = COALESCE(NULLIF(designation, ''), ?), modified_at = ? WHERE s_no = ?", (name.upper(), desig, current_time, sm2_row['s_no']))

        cursor.execute("SELECT id FROM employees WHERE UPPER(TRIM(code)) = ? ORDER BY created_at ASC LIMIT 1", (code.upper(),))
        emp_row = cursor.fetchone()
        if not emp_row:
            emp_id = f"EMP-{int(time.time()*1000) + e_idx}"
            desig = 'MANAGER' if is_manager_code(code) else 'STAFF'
            cursor.execute('''
                INSERT INTO employees (id, code, payroll_code, status, name, designation, dept, mode_of_transfer, basic_rate, spl_allowance, city_allow, master_approved, created_at, modified_at)
                VALUES (?, ?, '', 'ACTIVE', ?, ?, 'OPERATIONS', 'CASH', 0, 0, 0, 0, ?, ?)
            ''', (emp_id, code.upper(), name.upper(), desig, current_time, current_time))
        else:
            emp_id = emp_row['id']
        rec_key = f"{year}-{month}-{emp_id}"
        cursor.execute('''
            INSERT INTO payroll_records (record_key, paid_days, advance)
            VALUES (?, ?, 0)
            ON CONFLICT(record_key) DO UPDATE SET
                paid_days = excluded.paid_days
        ''', (rec_key, paid_days))

    conn.commit()
    conn.close()
    return jsonify({'status': 'success', 'monthKey': month_key, 'count': len(employees_data)})

# API: Cleanup Auto-Created Biometric Employees
@app.route('/api/employees/cleanup_auto_created', methods=['POST'])
def cleanup_auto_created():
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute('''
        DELETE FROM employees 
        WHERE designation = 'Staff' 
          AND (father_or_husband_name IS NULL OR father_or_husband_name = '')
          AND (uan IS NULL OR uan = '')
          AND (bank IS NULL OR bank = '')
    ''')
    deleted_count = cursor.rowcount
    conn.commit()
    conn.close()
    return jsonify({'status': 'success', 'deletedCount': deleted_count})

# API: Delete Biometric Month
@app.route('/api/biometric/<month_key>', methods=['DELETE'])
def delete_biometric_month(month_key):
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute('DELETE FROM biometric_months WHERE month_key = ?', (month_key,))
    cursor.execute('DELETE FROM biometric_records WHERE month_key = ?', (month_key,))
    conn.commit()
    conn.close()
    return jsonify({'status': 'success'})

# API: Save Month-Specific Record
@app.route('/api/monthly_record', methods=['POST'])
def save_monthly_record():
    data = request.json or {}
    rec_key = data.get('key')
    field = data.get('field')
    value = float(data.get('value', 0) or 0)
    
    col_mapping = {
        'paidDays': 'paid_days',
        'advance': 'advance',
        'basicRate': 'basic_rate',
        'splAllowance': 'spl_allowance',
        'cityAllow': 'city_allow',
        'deductShortTime': 'deduct_short_time',
        'overtimeArrear': 'overtime_arrear',
        'penaltyLate': 'penalty_late'
    }
    
    if field not in col_mapping:
        return jsonify({'status': 'error', 'message': 'Invalid field'}), 400
        
    db_col = col_mapping[field]
    conn = get_db_connection()
    cursor = conn.cursor()
    
    cursor.execute(f'''
        INSERT INTO payroll_records (record_key, paid_days, advance, {db_col})
        VALUES (?, 31, 0, ?)
        ON CONFLICT(record_key) DO UPDATE SET
            {db_col} = excluded.{db_col}
    ''', (rec_key, value))
    
    conn.commit()
    conn.close()
    return jsonify({'status': 'success'})

# API: Bulk Import Employees
@app.route('/api/employees/bulk', methods=['POST'])
def bulk_save_employees():
    employees_list = request.json.get('employees', [])
    if not employees_list:
        return jsonify({'status': 'error', 'message': 'No data provided'}), 400
    
    conn = get_db_connection()
    cursor = conn.cursor()
    current_time = get_current_timestamp()
    
    for idx, data in enumerate(employees_list):
        code = str(data.get('code', '') or data.get('payrollCode', '')).strip().upper()
        if not code:
            continue
        
        cursor.execute("SELECT id, created_at FROM employees WHERE code = ?", (code,))
        row = cursor.fetchone()

        if row:
            emp_id = row['id']
            created_at = row['created_at'] or current_time
            modified_at = current_time
        else:
            emp_id = f"EMP-{int(time.time()*1000) + idx}"
            created_at = current_time
            modified_at = current_time

        cursor.execute('''
            INSERT INTO employees (id, code, payroll_code, status, name, designation, dept, father_or_husband_name, dob, doj, dol, uan, epf_member_id, esic_ip_number, bank, acc, ifsc, mode_of_transfer, basic_rate, spl_allowance, city_allow, deduct_short_time, overtime_arrear, penalty_late, last_increment, created_at, modified_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(id) DO UPDATE SET
                code=excluded.code,
                payroll_code=excluded.payroll_code,
                status=excluded.status,
                name=excluded.name,
                designation=excluded.designation,
                dept=excluded.dept,
                father_or_husband_name=excluded.father_or_husband_name,
                dob=excluded.dob,
                doj=excluded.doj,
                dol=excluded.dol,
                uan=excluded.uan,
                epf_member_id=excluded.epf_member_id,
                esic_ip_number=excluded.esic_ip_number,
                bank=excluded.bank,
                acc=excluded.acc,
                ifsc=excluded.ifsc,
                mode_of_transfer=excluded.mode_of_transfer,
                basic_rate=excluded.basic_rate,
                spl_allowance=excluded.spl_allowance,
                city_allow=excluded.city_allow,
                deduct_short_time=excluded.deduct_short_time,
                overtime_arrear=excluded.overtime_arrear,
                penalty_late=excluded.penalty_late,
                last_increment=excluded.last_increment,
                modified_at=excluded.modified_at
        ''', (
            emp_id, code,
            str(data.get('payrollCode', '') or '').strip().upper(),
            data.get('status', 'Active'),
            data.get('name', ''),
            data.get('designation', ''),
            data.get('dept', 'Operations'),
            data.get('fatherOrHusbandName', ''),
            data.get('dob', ''),
            data.get('doj', ''),
            data.get('dol', ''),
            str(data.get('uan', '')),
            str(data.get('epfMemberId', '')),
            str(data.get('esicIpNumber', '')),
            data.get('bank', ''),
            str(data.get('acc', '')),
            str(data.get('ifsc', '')).upper(),
            data.get('modeOfTransfer', 'Bank Transfer'),
            float(data.get('basicRate', 0) or 0),
            float(data.get('splAllowance', 0) or 0),
            float(data.get('cityAllow', 0) or 0),
            float(data.get('deductShortTime', 0) or 0),
            float(data.get('overtimeArrear', 0) or 0),
            float(data.get('penaltyLate', 0) or 0),
            data.get('lastIncrement', ''),
            created_at, modified_at
        ))
        # Seed independent Salary Master values only for newly represented employees.
        cursor.execute('''
            INSERT OR IGNORE INTO salary_master
                (employee_id, emp_code, payroll_code, employee_name, basic_rate, spl_allowance, city_allow, modified_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        ''', (emp_id, code, str(data.get('payrollCode', '') or '').strip().upper(),
              str(data.get('name', '') or '').strip().upper(),
              float(data.get('basicRate', 0) or 0), float(data.get('splAllowance', 0) or 0),
              float(data.get('cityAllow', 0) or 0), modified_at))
        cursor.execute('UPDATE salary_master SET emp_code = ?, payroll_code = ?, employee_name = ? WHERE employee_id = ?',
                       (code, str(data.get('payrollCode', '') or '').strip().upper(), str(data.get('name', '') or '').strip().upper(), emp_id))
        
    conn.commit()
    conn.close()
    return jsonify({'status': 'success', 'count': len(employees_list)})

# API: Delete Employee
# API: Bulk Delete Employees
@app.route('/api/employees/bulk_delete', methods=['POST'])
def bulk_delete_employees():
    data = request.json or {}
    emp_ids = data.get('ids', [])
    if not emp_ids:
        return jsonify({'status': 'error', 'message': 'No employee IDs provided'}), 400
    
    conn = get_db_connection()
    cursor = conn.cursor()
    
    placeholders = ','.join(['?'] * len(emp_ids))
    cursor.execute(f"DELETE FROM employees WHERE id IN ({placeholders})", emp_ids)
    deleted_count = cursor.rowcount
    
    for emp_id in emp_ids:
        cursor.execute("DELETE FROM payroll_records WHERE record_key LIKE ?", (f"%{emp_id}%",))
        
    conn.commit()
    conn.close()
    return jsonify({'status': 'success', 'deletedCount': deleted_count})

@app.route('/api/employee/<emp_id>', methods=['DELETE'])
def delete_employee(emp_id):
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute('DELETE FROM employees WHERE id = ?', (emp_id,))
    cursor.execute("DELETE FROM payroll_records WHERE record_key LIKE ?", (f"%{emp_id}%",))
    conn.commit()
    conn.close()
    return jsonify({'status': 'success'})

# API: Update Attendance & Advance
@app.route('/api/attendance', methods=['POST'])
def update_attendance():
    data = request.json or {}
    rec_key = data['key']
    paid_days = float(data.get('paidDays', 31))
    advance = float(data.get('advance', 0))
    
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute('''
        INSERT INTO payroll_records (record_key, paid_days, advance)
        VALUES (?, ?, ?)
        ON CONFLICT(record_key) DO UPDATE SET
            paid_days = excluded.paid_days,
            advance = excluded.advance
    ''', (rec_key, paid_days, advance))
    conn.commit()
    conn.close()
    return jsonify({'status': 'success'})

# API: Save Configuration
@app.route('/api/config', methods=['POST'])
def save_config():
    data = request.json or {}
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute('''
        UPDATE config SET
            company_name = ?,
            company_addr = ?,
            pf_rate = ?,
            esic_rate = ?,
            hra_rate = ?,
            pt_amount = ?
        WHERE id = 1
    ''', (data['companyName'], data['companyAddr'], data['pfRate'], data['esicRate'], data['hraRate'], data['ptAmount']))
    conn.commit()
    conn.close()
    return jsonify({'status': 'success'})

# ==============================================================================
# SPECIAL EXCEL EXPORT: PRANAV CREATION SALARY SHEET WITH EMPLOYEE TABS
# (Matches exact structure of SALARY_SHEET_SEPTEMBER_2026_FINAL_CLEAN.xlsx)
# ==============================================================================
@app.route('/api/export_pranav_salary_sheet', methods=['GET'])
def export_pranav_salary_sheet():
    m_param = request.args.get('month', '8')
    y_param = request.args.get('year', '2026')
    month_num = int(m_param) if m_param.isdigit() else 8
    year = int(y_param) if y_param.isdigit() else 2026

    month_names = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]
    month_name = month_names[month_num - 1].upper()
    days_in_month = (datetime(year, month_num % 12 + 1, 1) - timedelta(days=1)).day if month_num != 12 else 31

    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute('SELECT company_name FROM config WHERE id = 1')
    cfg_row = cursor.fetchone()
    company_name = cfg_row['company_name'] if cfg_row else 'PRANAV CREATION PRIVATE LIMITED'

    cursor.execute('SELECT * FROM employees ORDER BY CAST(code AS INTEGER), code')
    emp_rows = cursor.fetchall()

    month_key = f"{year}-{month_num}"
    cursor.execute('SELECT * FROM biometric_records WHERE month_key = ?', (month_key,))
    bio_rows = cursor.fetchall()
    conn.close()

    bio_records_map = {}
    for r in bio_rows:
        parsed_logs = []
        if r['daily_logs']:
            try:
                parsed_logs = json.loads(r['daily_logs'])
            except Exception:
                parsed_logs = []
        bio_records_map[str(r['emp_code']).strip()] = {
            'presentDays': r['present_days'],
            'absentDays': r['absent_days'],
            'weeklyOff': r['weekly_off'],
            'holiday': r['holiday'],
            'paidDays': r['paid_days'],
            'dailyLogs': parsed_logs
        }

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    c_navy_dark = '1F3864'
    c_blue_med  = '2E5C9A'
    c_blue_hdr  = '4472C4'
    c_blue_tot  = 'D9E2F3'

    f_title     = Font(name='Arial', size=14, bold=True, color='FFFFFF')
    f_sub       = Font(name='Arial', size=11, bold=True, color='FFFFFF')
    f_hdr_main  = Font(name='Arial', size=9, bold=True, color='FFFFFF')
    f_hdr_emp   = Font(name='Arial', size=9, bold=True, color='000000')
    f_data      = Font(name='Arial', size=9, bold=False, color='000000')
    f_tot       = Font(name='Arial', size=9, bold=True, color='000000')
    f_bold      = Font(name='Arial', size=9, bold=True, color='000000')
    f_invis     = Font(name='Arial', size=9, color='FFFFFF')

    fill_title  = PatternFill(start_color=c_navy_dark, end_color=c_navy_dark, fill_type='solid')
    fill_sub    = PatternFill(start_color=c_blue_med,  end_color=c_blue_med,  fill_type='solid')
    fill_hdr    = PatternFill(start_color=c_blue_hdr,  end_color=c_blue_hdr,  fill_type='solid')
    fill_tot    = PatternFill(start_color=c_blue_tot,  end_color=c_blue_tot,  fill_type='solid')

    b_thin = Side(style='thin', color='000000')
    border_std = Border(left=b_thin, right=b_thin, top=b_thin, bottom=b_thin)

    al_center = Alignment(horizontal='center', vertical='center')
    al_left   = Alignment(horizontal='left',   vertical='center')
    al_right  = Alignment(horizontal='right',  vertical='center')
    al_hdr    = Alignment(horizontal='center', vertical='center', wrap_text=True)

    def parse_time(val):
        if val is None or str(val).strip() in ['', '-', 'None']:
            return dt_time(10, 0)
        if isinstance(val, dt_time):
            return val
        if isinstance(val, datetime):
            return val.time()
        val_str = str(val).strip()
        if ':' in val_str:
            try:
                parts = [int(p) for p in re.findall(r'\d+', val_str)]
                if len(parts) >= 2:
                    return dt_time(parts[0] % 24, parts[1], parts[2] if len(parts) > 2 else 0)
            except Exception:
                pass
        return dt_time(10, 0)

    existing_tabs = set()
    def get_clean_tab(name, code):
        base = re.sub(r'[\\/*?:\[\]]', '', str(name).split()[0]).upper()[:20]
        if not base or base.lower() == 'none':
            base = f"EMP_{code}"
        tab = base
        cnt = 1
        while tab in existing_tabs:
            tab = f"{base}_{cnt}"
            cnt += 1
        existing_tabs.add(tab)
        return tab

    # 1. MAIN SHEET Setup
    ws_main = wb.create_sheet(title='MAIN SHEET')

    ws_main.merge_cells('A1:O1')
    ws_main['A1'] = company_name.upper()
    ws_main['A1'].font = f_title
    ws_main['A1'].alignment = al_center
    for c in range(1, 16):
        ws_main.cell(1, c).fill = fill_title
    ws_main.row_dimensions[1].height = 24.0

    ws_main.merge_cells('A2:O2')
    ws_main['A2'] = f"DETAIL OF DEDUCTION FROM SALARY FOR THE MONTH OF {month_name} {year}"
    ws_main['A2'].font = f_sub
    ws_main['A2'].alignment = al_center
    for c in range(1, 16):
        ws_main.cell(2, c).fill = fill_sub
    ws_main.row_dimensions[2].height = 20.0

    ws_main.row_dimensions[3].height = 10.0

    main_headers = [
        'ORG Emp Code', 'PAYROLL EMP CODE', 'Name of employee', 'Designation',
        'Days In Month', 'Leaves Taken', 'Attendance', 'Total Hours',
        'Total Working Hours', 'Short Time', 'Overtime', 'Amount',
        'Short Deduction', 'Overtime', 'Remarks'
    ]
    ws_main.row_dimensions[4].height = 32.0
    for c_idx, h in enumerate(main_headers, start=1):
        cell = ws_main.cell(4, c_idx, h)
        cell.font = f_hdr_main
        cell.fill = fill_hdr
        cell.alignment = al_hdr
        cell.border = border_std

    col_widths = {
        'A': 12.0, 'B': 14.0, 'C': 30.0, 'D': 15.0, 'E': 9.0,
        'F': 12.0, 'G': 11.0, 'H': 12.0, 'I': 13.0, 'J': 11.0,
        'K': 12.0, 'L': 13.0, 'M': 14.0, 'N': 13.0, 'O': 26.0
    }
    for col_l, w in col_widths.items():
        ws_main.column_dimensions[col_l].width = w

    first_data_row = 5
    row_idx = first_data_row
    standard_hours = 9

    detected_holidays = {}
    for code, b_data in bio_records_map.items():
        for l in b_data.get('dailyLogs', []):
            dt_s = l.get('date')
            in_v = str(l.get('inTime', '')).strip().upper()
            out_v = str(l.get('outTime', '')).strip().upper()
            for cand in [in_v, out_v]:
                if cand and cand not in ['', 'P', 'A', 'WO', 'NONE', 'OFF', 'WEEK OFF', '-'] and not re.match(r'^\d{1,2}:\d{2}', cand):
                    if any(w in cand for w in ['RAKSHA', 'BANDHAN', 'INDEPENDENCE', 'DIWALI', 'HOLI', 'EID', 'FESTIVAL', 'HOLIDAY', 'GANDHI', 'REPUBLIC']) or len(cand) > 3:
                        try:
                            d_obj = datetime.strptime(dt_s, '%d/%m/%Y').date()
                            detected_holidays[d_obj] = cand
                        except Exception:
                            pass

    for emp in emp_rows:
        code_str = str(emp['code']).strip()
        name_str = emp['name']
        tab_name = get_clean_tab(name_str, code_str)
        ws = wb.create_sheet(title=tab_name)

        rate_val = emp['basic_rate']

        ws['A1'] = 'NAME'
        ws['B1'] = name_str
        ws['F1'] = None

        if rate_val is not None and float(rate_val) > 0:
            ws['G1'] = float(rate_val)
            ws['G1'].number_format = ';;;';
            ws['G1'].font = f_invis
        else:
            ws['G1'] = None

        ws['H1'] = 'DAYS'
        ws['I1'] = days_in_month
        ws['J1'] = 'HOURS'
        ws['K1'] = standard_hours

        ws['A2'] = 'CODE#'
        try:
            ws['B2'] = int(float(code_str))
        except Exception:
            ws['B2'] = code_str

        ws.row_dimensions[1].height = 18.0
        ws.row_dimensions[2].height = 18.0

        top_cells = ['A1', 'B1', 'H1', 'I1', 'J1', 'K1', 'A2', 'B2']
        for c_ref in top_cells:
            cell = ws[c_ref]
            cell.font = f_bold
            cell.border = border_std
            if c_ref in ['A1', 'H1', 'J1', 'A2', 'I1', 'K1', 'B2']:
                cell.alignment = al_center
            else:
                cell.alignment = al_left

        emp_headers = [
            'MONTH', 'DATE', 'DAY', 'INTIME', 'OUTTIME',
            'TOTAL WORKING HRS', 'ACTUAL WORKING HRS',
            'SHORT OR EXCESS WORKING HRS', 'RATE PER HOUR',
            'GROSS SALARY', '', 'LEAVE'
        ]
        ws.row_dimensions[4].height = 30.0
        for c_idx in list(range(1, 11)) + [12]:
            cell = ws.cell(4, c_idx, emp_headers[c_idx - 1])
            cell.font = f_hdr_emp
            cell.alignment = al_hdr
            cell.border = border_std

        b_emp = bio_records_map.get(code_str, {})
        daily_logs_map = {}
        for l in b_emp.get('dailyLogs', []):
            try:
                dt_obj = datetime.strptime(l['date'], '%d/%m/%Y').date()
                daily_logs_map[dt_obj] = l
            except Exception:
                pass

        def check_day_is_leave(check_date):
            if check_date in detected_holidays:
                return False
            rec_c = daily_logs_map.get(check_date)
            if not rec_c:
                return True
            st_c = str(rec_c.get('status', '')).strip().upper()
            in_c = str(rec_c.get('inTime', '')).strip()
            out_c = str(rec_c.get('outTime', '')).strip()
            if st_c == 'P' and in_c and in_c != '-' and out_c and out_c != '-':
                return False
            return True

        for d in range(1, days_in_month + 1):
            r_num = 4 + d
            ws.row_dimensions[r_num].height = 16.0
            dt = datetime(year, month_num, d).date()
            day_name = dt.strftime('%a').upper()

            ws.cell(r_num, 1, month_name).alignment = al_center
            c_date = ws.cell(r_num, 2, dt)
            c_date.number_format = 'dd-mm-yyyy'
            c_date.alignment = al_center
            ws.cell(r_num, 3, day_name).alignment = al_center

            is_leave = False
            is_wo = False
            is_holiday = (dt in detected_holidays)
            holiday_title = detected_holidays.get(dt, 'HOLIDAY')

            in_time_val = None
            out_time_val = None

            if not daily_logs_map:
                is_leave = True
            else:
                rec = daily_logs_map.get(dt)
                if is_holiday:
                    pass
                elif day_name == 'SUN' or (rec and str(rec.get('status', '')).strip().upper() == 'WO'):
                    next_dt = dt + timedelta(days=1)
                    if check_day_is_leave(next_dt):
                        is_leave = True
                        is_wo = False
                    else:
                        is_wo = True
                        is_leave = False
                elif rec:
                    st = str(rec.get('status', '')).strip().upper()
                    in_raw = str(rec.get('inTime', '')).strip()
                    out_raw = str(rec.get('outTime', '')).strip()
                    if st == 'P' and in_raw and in_raw != '-' and out_raw and out_raw != '-':
                        in_time_val = in_raw
                        out_time_val = out_raw
                    else:
                        is_leave = True
                else:
                    is_leave = True

            if is_holiday and not is_leave:
                ws.merge_cells(start_row=r_num, start_column=4, end_row=r_num, end_column=5)
                h_cell = ws.cell(r_num, 4, holiday_title)
                h_cell.alignment = al_center
                h_cell.font = f_bold
                ws.cell(r_num, 6, dt_time(standard_hours, 0)).number_format = '[h]:mm'
                ws.cell(r_num, 7, dt_time(standard_hours, 0)).number_format = '[h]:mm'
            elif is_wo and not is_leave:
                ws.merge_cells(start_row=r_num, start_column=4, end_row=r_num, end_column=5)
                wo_cell = ws.cell(r_num, 4, 'WEEK OFF')
                wo_cell.alignment = al_center
                wo_cell.font = f_bold
                ws.cell(r_num, 6, dt_time(standard_hours, 0)).number_format = '[h]:mm'
                ws.cell(r_num, 7, dt_time(standard_hours, 0)).number_format = '[h]:mm'
            elif is_leave:
                ws.merge_cells(start_row=r_num, start_column=4, end_row=r_num, end_column=5)
                lv_cell = ws.cell(r_num, 4, 'LEAVE')
                lv_cell.alignment = al_center
                lv_cell.font = f_bold
                ws.cell(r_num, 6, dt_time(0, 0)).number_format = '[h]:mm'
                ws.cell(r_num, 7, dt_time(0, 0)).number_format = '[h]:mm'
                lv_val = ws.cell(r_num, 12, 1)
                lv_val.alignment = al_center
                lv_val.font = f_bold
            else:
                t_in = parse_time(in_time_val)
                t_out = parse_time(out_time_val)

                c_in = ws.cell(r_num, 4, t_in)
                c_in.number_format = 'hh:mm'
                c_in.alignment = al_center

                c_out = ws.cell(r_num, 5, t_out)
                c_out.number_format = 'hh:mm'
                c_out.alignment = al_center

                ws.cell(r_num, 6, f"=IF(E{r_num}>=D{r_num}, E{r_num}-D{r_num}, 1+E{r_num}-D{r_num})").number_format = '[h]:mm'
                ws.cell(r_num, 7, dt_time(standard_hours, 0)).number_format = '[h]:mm'

            ws.cell(r_num, 6).alignment = al_right
            ws.cell(r_num, 7).alignment = al_right

            ws.cell(r_num, 8, f'=IF(F{r_num}>=G{r_num}, TEXT(F{r_num}-G{r_num}, "h:mm"), "-" & TEXT(G{r_num}-F{r_num}, "h:mm"))').alignment = al_right
            ws.cell(r_num, 9, "=$G$1/$I$1/$K$1").alignment = al_right
            ws.cell(r_num, 9).number_format = '0.00'

            ws.cell(r_num, 10, f'=IF(F{r_num}>=G{r_num}, (F{r_num}-G{r_num})*I{r_num}*24, -(G{r_num}-F{r_num})*I{r_num}*24)').alignment = al_right
            ws.cell(r_num, 10).number_format = '#,##0.00;-#,##0.00;0.00'

            for c in list(range(1, 11)) + [12]:
                cell = ws.cell(r_num, c)
                cell.border = border_std
                if not cell.font or not cell.font.bold:
                    cell.font = f_data

        tot_r = 4 + days_in_month + 1
        ws.row_dimensions[tot_r].height = 17.0

        ws.cell(tot_r, 1, '=B2').alignment = al_center
        ws.cell(tot_r, 6, f'=SUM(F5:F{tot_r-1})').number_format = '[h]:mm'
        ws.cell(tot_r, 6).alignment = al_right
        ws.cell(tot_r, 7, f'=SUM(G5:G{tot_r-1})').number_format = '[h]:mm'
        ws.cell(tot_r, 7).alignment = al_right

        ws.cell(tot_r, 8, f'=IF(F{tot_r}>=G{tot_r}, TEXT(F{tot_r}-G{tot_r}, "[h]:mm"), "-" & TEXT(G{tot_r}-F{tot_r}, "[h]:mm"))').alignment = al_right
        ws.cell(tot_r, 10, f'=SUM(J5:J{tot_r-1})').number_format = '#,##0.00;-#,##0.00;0.00'
        ws.cell(tot_r, 10).alignment = al_right
        ws.cell(tot_r, 12, f'=SUM(L5:L{tot_r-1})').alignment = al_center

        for c in list(range(1, 11)) + [12]:
            cell = ws.cell(tot_r, c)
            cell.font = f_tot
            cell.border = border_std

        emp_col_widths = {
            'A': 10.0, 'B': 12.0, 'C': 7.0,  'D': 11.0, 'E': 11.0, 'F': 14.0,
            'G': 14.0, 'H': 16.0, 'I': 13.0, 'J': 14.0, 'K': 3.0,  'L': 8.0
        }
        for col_l, w in emp_col_widths.items():
            ws.column_dimensions[col_l].width = w

        ws.auto_filter.ref = f"A4:L{tot_r-1}"

        # Main Sheet Entry
        ws_main.row_dimensions[row_idx].height = 16.0
        try:
            ws_main.cell(row_idx, 1, int(float(code_str))).alignment = al_center
            ws_main.cell(row_idx, 2, int(float(code_str))).alignment = al_center
        except Exception:
            ws_main.cell(row_idx, 1, code_str).alignment = al_center
            ws_main.cell(row_idx, 2, code_str).alignment = al_center

        ws_main.cell(row_idx, 3, name_str).alignment = al_left
        ws_main.cell(row_idx, 4, emp['designation'] or '-').alignment = al_left
        ws_main.cell(row_idx, 5, days_in_month).alignment = al_center

        ws_main.cell(row_idx, 6, f"='{tab_name}'!L{tot_r}").alignment = al_center
        ws_main.cell(row_idx, 7, f"=E{row_idx}-F{row_idx}").alignment = al_center
        ws_main.cell(row_idx, 8, f"='{tab_name}'!G{tot_r}").number_format = '[h]:mm'
        ws_main.cell(row_idx, 9, f"='{tab_name}'!F{tot_r}").number_format = '[h]:mm'
        ws_main.cell(row_idx, 10, f"=IF((H{row_idx}>I{row_idx}),(H{row_idx}-I{row_idx}),0)").number_format = '[h]:mm'
        ws_main.cell(row_idx, 11, f"=IF((I{row_idx}>H{row_idx}),(I{row_idx}-H{row_idx}),0)").number_format = '[h]:mm'
        ws_main.cell(row_idx, 12, f"='{tab_name}'!J{tot_r}").number_format = '#,##0.00;(#,##0.00);-'
        ws_main.cell(row_idx, 13, f"=IF(L{row_idx}<0,L{row_idx},0)").number_format = '#,##0.00;(#,##0.00);-'
        ws_main.cell(row_idx, 14, f"=IF(L{row_idx}>0,L{row_idx}*2,0)").number_format = '#,##0.00;(#,##0.00);-'

        if rate_val is None or float(rate_val or 0) <= 0:
            ws_main.cell(row_idx, 15, "Rate Missing in Master").font = Font(name='Arial', size=8, color='C00000', bold=True)
            ws_main.cell(row_idx, 15).alignment = al_left

        for c in range(1, 16):
            cell = ws_main.cell(row_idx, c)
            cell.border = border_std
            cell.font = f_data
            if c not in [3, 4, 15]:
                cell.alignment = al_center

        row_idx += 1

    # Grand Total Row
    total_row = row_idx + 1
    ws_main.row_dimensions[total_row].height = 18.0

    ws_main.merge_cells(start_row=total_row, start_column=1, end_row=total_row, end_column=4)
    ws_main.cell(total_row, 1, 'TOTAL DEDUCTION').alignment = al_left

    ws_main.cell(total_row, 5, f'=SUM(E{first_data_row}:E{row_idx-1})')
    ws_main.cell(total_row, 6, f'=SUM(F{first_data_row}:F{row_idx-1})')
    ws_main.cell(total_row, 7, f'=SUM(G{first_data_row}:G{row_idx-1})')
    ws_main.cell(total_row, 8, f'=SUM(H{first_data_row}:H{row_idx-1})').number_format = '[h]:mm'
    ws_main.cell(total_row, 9, f'=SUM(I{first_data_row}:I{row_idx-1})').number_format = '[h]:mm'
    ws_main.cell(total_row, 10, f'=SUM(J{first_data_row}:J{row_idx-1})').number_format = '[h]:mm'
    ws_main.cell(total_row, 11, f'=SUM(K{first_data_row}:K{row_idx-1})').number_format = '[h]:mm'
    ws_main.cell(total_row, 12, f'=SUM(L{first_data_row}:L{row_idx-1})').number_format = '#,##0.00;(#,##0.00);-'
    ws_main.cell(total_row, 13, f'=SUM(M{first_data_row}:M{row_idx-1})').number_format = '#,##0.00;(#,##0.00);-'
    ws_main.cell(total_row, 14, f'=SUM(N{first_data_row}:N{row_idx-1})').number_format = '#,##0.00;(#,##0.00);-'

    for c in range(1, 16):
        cell = ws_main.cell(total_row, c)
        cell.font = f_tot
        cell.fill = fill_tot
        cell.border = border_std
        if c > 4:
            cell.alignment = al_center

    output_stream = io.BytesIO()
    wb.save(output_stream)
    output_stream.seek(0)
    filename = f"SALARY_SHEET_{month_name}_{year}_FINAL_CLEAN.xlsx"
    return send_file(output_stream, as_attachment=True, download_name=filename, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port, debug=os.environ.get('FLASK_DEBUG', '0') == '1')
