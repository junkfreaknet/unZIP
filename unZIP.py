#-----
import os
import zipfile
from pathlib import Path
#from pathlib import PurePath
#from pathlib import PurePosixPath
#from pathlib import PureWindowsPath
import pyodbc
import csv
#-----import pandas

#-----  create sql string(insert) start
def createSQLStringInsert_TableChouChou(fields_name,row,Table_Name):

    #CONST_FIELD_NAME_TODOUFUKEN_CODE='都道府県コード'
    #CONST_FIELD_NAME_TODOUFUKEN_NAME='都道府県名'
    #CONST_FIELD_NAME_SIKUCHOUSON_CODE='市区町村コード'
    #CONST_FIELD_NAME_SIKUCHOUSON_NAME='市区町村名'
    #CONST_FIELD_NAME_OOAZACHOUME_CODE='大字町丁目コード'
    #CONST_FIELD_NAME_OOAZACHOUME_NAME='大字町丁目名'
    #CONST_FIELD_NAME_IDO='緯度'
    #CONST_FIELD_NAME_KEIDO='経度'
    #CONST_FIELD_NAME_GENTENSIRYOU_CODE='原典資料コード'
    #CONST_FIELD_NAME_OOAZAAZACHOUMEKUBUN_CODE='大字・字・丁目区分コード'

    rv='insert into '+Table_Name+' ('+fields_name[0]+','+fields_name[1]+','+fields_name[2]+','+fields_name[3]+','+fields_name[4]+','+fields_name[5]+','+fields_name[6]+','+fields_name[7]+','+fields_name[8]+','+fields_name[9]+') values ('+'\''+row[0]+'\','+'\''+row[1]+'\','+'\''+row[2]+'\','+'\''+row[3]+'\','+'\''+row[4]+'\','+'\''+row[5]+'\','+'\''+row[6]+'\','+'\''+row[7]+'\','+'\''+row[8]+'\','+'\''+row[9]+'\''+')'

    return rv
#-----  create sql string(insert) end

#-----  init table  start
def initTable(func_DB_Connection,Table_NAme):

    SQLString='delete * from '+Table_NAme+';'
    db_cursor=func_DB_Connection.cursor()
    db_cursor.execute(SQLString)
    func_DB_Connection.commit()

#-----  init table  end

#-----  create field name list start
def createFieldListFromCSVHeader(header):

    list_Fields=[]
    for row in header:
        list_Fields.append(row)
    return list_Fields
#-----  create field name list end

#-----  import from csv to databsse start
def importFromCSVToDatabaseChouChouData(func_DB_Connection,csv_File,Table_Name):
    
    print(csv_File)

    db_cursor=func_DB_Connection.cursor()    

    with open(csv_File, newline='', encoding='shift_jisx0213') as f:

        header=next(csv.reader(f))
        list_Fields=createFieldListFromCSVHeader(header)

        reader = csv.reader(f)
        for row in reader:
            SQLSTRING=createSQLStringInsert_TableChouChou(list_Fields,row,Table_Name)
            db_cursor.execute(SQLSTRING)
    db_cnnctn.commit()
    db_cursor.close()

#-----  import from csv to database end

#-----  get ToDouFuKen name start
def getToDouFuKenFromAccess(func_DB_Connection,ToDouFuKen_cd):

    db_cursor=func_DB_Connection.cursor()

    para_like=ToDouFuKen_cd+'%'

    for a_row in db_cursor.execute('Select * from TOToDouFuKen where DantaiCode like ?;',para_like).fetchall():
        a_return=a_row[1]
    db_cursor.close()

    return a_return
#-----  get ToDouFuKen name end
#-----  end functions

#-----  main

CONST_DIR_ZIP='../unZIP/町丁データCompress'
CONST_DIR_ZIP_EXTRACT='../unZIP/町丁データExtrace'
CONST_DIR_FINAL='../unzip/町丁データ'
#CONST_PATH_TO_ACCESS=r'C:\Users\junkf\Documents\tmp\dust\VisualStudio\projects\XLStoTXT_jpAddressCode\db0307.accdb'
CONST_PATH_TO_ACCESS=r'../unZIP/db0307.accdb'
CONST_TABLE_NAME_CHOUCHOU_DATA='TOChouChouData'
#-----  init target directory start

if not Path(CONST_DIR_FINAL).exists():
    Path(CONST_DIR_FINAL).mkdir()

for x in Path(CONST_DIR_FINAL).glob('*'):
    os.remove(x)
#-----  init target directory end

#-----  extract zip start

for x in Path(CONST_DIR_ZIP).glob('*.zip'):
    with zipfile.ZipFile(x) as zf:
        rv=zf.namelist()
        zf.extract(member=rv[0],path=CONST_DIR_ZIP_EXTRACT)
#-----  extract zip end

#-----  move to final directory start-----

str_DB_Connection=(
    r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
    #r'DBQ=../unZIP/db0307.accdb;'
    r'DBQ='+CONST_PATH_TO_ACCESS+';'''
)

db_cnnctn=pyodbc.connect(str_DB_Connection) #-----  connect database
#db_cnnctn.getinfo
for a_dir in Path(CONST_DIR_ZIP_EXTRACT).glob('*'):
    for a_csv_file in Path(a_dir).glob('*'):
        iter_a_csv_file_name=os.path.splitext(a_csv_file.name)
        str_ToDouFuKen=getToDouFuKenFromAccess(db_cnnctn,a_csv_file.name[0:2])
        #print(str_ToDouFuKen)
        Path(a_csv_file).rename(CONST_DIR_FINAL+'/'+iter_a_csv_file_name[0]+'_'+str_ToDouFuKen+iter_a_csv_file_name[1])
    a_dir.rmdir()
#-----  move to final directory end-----

#-----  import csv to database start-----

#-----  import csv to database end -----

initTable(db_cnnctn,CONST_TABLE_NAME_CHOUCHOU_DATA) #-----  init target table 

for a_csv_file in Path(CONST_DIR_FINAL).glob('*.csv'):
    #print(a_csv_file)

    importFromCSVToDatabaseChouChouData(db_cnnctn,a_csv_file,CONST_TABLE_NAME_CHOUCHOU_DATA)
#-----  terminate program

db_cnnctn.close()

print('at last current running directory is :'+str(Path.cwd()))