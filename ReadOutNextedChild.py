
import os
import csv
from bs4 import BeautifulSoup
import requests

def parse_HTML_Todoufuken(url_source_prefix,url_source,output_dir,output_file):

#
    const_FIELD_NAME_VALUE_STANDARD_AREA_CODE='標準地域コード'
    const_FIELD_NAME_VALUE_SICHOUSIN='市区町村'
    const_FIELD_NAME_VALUE_URL='URL'

#
    response=requests.get(url_source)
    contents=response.content
    soup=BeautifulSoup(contents,"html.parser")
    findtbody=soup.find('tbody')    

    trIdx=1
    
    with open(output_dir+os.sep+output_file,'w') as f:
        for a_tr in findtbody.find_all('tr'):

            if trIdx==1:
                buffHeader="\""+const_FIELD_NAME_VALUE_STANDARD_AREA_CODE+"\",\""+const_FIELD_NAME_VALUE_SICHOUSIN+"\",\""+const_FIELD_NAME_VALUE_URL+"\"\n"
                f.write(buffHeader)   #write header at onece

            tdIdx=1

            for a_td in a_tr.find_all('td'):
                if tdIdx==2:
                    buffAnchor=url_source_prefix+a_td.find('a').get('href')
                    buffCode=a_td.string
                if tdIdx==3:
                    buffAdd=a_td.string
                tdIdx=tdIdx+1

            #print(buffCode+","+buffAdd+","+url_source_prefix+buffAnchor)
            buffWrite="\""+buffCode+"\","+"\""+buffAdd+"\","+"\""+buffAnchor+"\""
            buffDetail="\""+buffCode+"\",\""+buffAdd+"\",\""+buffAnchor+"\"\n"
            f.write(buffDetail)
            trIdx=trIdx+1

        print('Number of tr tag in this file: '+ format(trIdx))
    f.close()
#===end function===
def parse_HTML_KyouKaiSen(source_dir,source_file,output_dir,output_file,error_dir,error_file):
#
    const_FIELD_NAME_VALUE_ID='ID'
    const_FIELD_NAME_VALUE_PLACE_NAME='地名'
    const_FIELD_NAME_VALUE_AREA='面積（㎡）'
    const_FIELD_NAME_VALUE_PERIMETER_LENGTH='周辺長（ｍ）'
    const_FIELD_NAME_VALUE_POPULATION='人口'
    const_FIELD_NAME_VALUE_NUM_OF_HOUSEHOLDS='世帯数'

    const_FIELD_POS_ID=1
    const_FIELD_POS_PLACE_NAME=2
    const_FIELD_POS_AREA=3
    const_FIELD_POS_PERIMETER_LENGTH=4
    const_FIELD_POS_POPULATION=5
    const_FIELD_POS_NUM_OF_HOUSEHOLDS=6

#
    const_Pos_URL=2     #pos of URL in a row.
    const_Header_Record_Pos=0
    idx_Row=const_Header_Record_Pos
    idx_Write=0

    with open(output_dir+os.sep+output_file,'w') as outf:
        buffHeader="\""+const_FIELD_NAME_VALUE_ID+"\",\""+const_FIELD_NAME_VALUE_PLACE_NAME+"\",\""+const_FIELD_NAME_VALUE_AREA+"\",\""+const_FIELD_NAME_VALUE_PERIMETER_LENGTH+"\",\""+const_FIELD_NAME_VALUE_POPULATION+"\",\""+const_FIELD_NAME_VALUE_NUM_OF_HOUSEHOLDS+"\"\n"
        outf.write(buffHeader)   #write header at onece
        print(buffHeader)
        f_Error_Head=False
        idx_Error=0
        with open(source_dir+os.sep+source_file) as inf:
            list_All=csv.reader(inf,delimiter=',',quotechar='"')
            for a_row in list_All:
                if idx_Row!=const_Header_Record_Pos:
                    #print(idx_row,a_row[const_Pos_URL]) #for test
                    response=requests.get(a_row[const_Pos_URL])
                    contents=response.content
                    soup=BeautifulSoup(contents,"html.parser")
                    findtbody=soup.find('tbody')

                    for a_tr in findtbody.find_all('tr'):

                        tdIdx=1

                        for a_td in a_tr.find_all('td'):
                            if tdIdx==const_FIELD_POS_ID:
                                buffID=a_td.string
                            if tdIdx==const_FIELD_POS_PLACE_NAME:
                                buffPlaceName=a_td.string
                            if tdIdx==const_FIELD_POS_AREA:
                                buffArea=a_td.string
                            if tdIdx==const_FIELD_POS_PERIMETER_LENGTH:
                                buffPerimeterLength=a_td.string
                            if tdIdx==const_FIELD_POS_POPULATION:
                                buffPopulation=a_td.string
                            if tdIdx==const_FIELD_POS_NUM_OF_HOUSEHOLDS:
                                buffNumOfHouseholds=a_td.string
                            
                            tdIdx=tdIdx+1

                        if not buffPlaceName:
                            if not f_Error_Head:
                                with open(error_dir+os.sep+error_file,'w') as errorf:
                                    #errorf.write("\""+"ID"+"\",\""+"Reason"+"\"\n")
                                    errorf.write("\""+const_FIELD_NAME_VALUE_ID+"\",\""+const_FIELD_NAME_VALUE_PLACE_NAME+"\",\""+const_FIELD_NAME_VALUE_AREA+"\",\""+const_FIELD_NAME_VALUE_PERIMETER_LENGTH+"\",\""+const_FIELD_NAME_VALUE_POPULATION+"\",\""+const_FIELD_NAME_VALUE_NUM_OF_HOUSEHOLDS+"\"\n")
                                    #errorf.write("\""+buffID+"\",\""+"***** no place name *****"+"\"\n")
                                    errorf.write("\""+buffID+"\",\"***** no place name *****"+""+"\","+buffArea+","+buffPerimeterLength+","+buffPopulation+","+buffNumOfHouseholds+"\n")
                                    print(buffID+',***** no place name *****')
                                errorf.close()
                                f_Error_Head=True
                                idx_Error=1
                            else:
                                with open(error_dir+os.sep+error_file,'a') as errorf:
                                    #errorf.write("\""+buffID+"\",\""+"***** no place name *****"+"\"\n")
                                    errorf.write("\""+buffID+"\",\"***** no place name *****"+""+"\","+buffArea+","+buffPerimeterLength+","+buffPopulation+","+buffNumOfHouseholds+"\n")                                    
                                    print(buffID+',***** no place name *****')
                                errorf.close()  
                                idx_Error=idx_Error+1
                        else:
                            buffDetail="\""+buffID+"\",\""+buffPlaceName+"\","+buffArea+","+buffPerimeterLength+","+buffPopulation+","+buffNumOfHouseholds+"\n"
                            outf.write(buffDetail)
                            idx_Write=idx_Write+1
                idx_Row=idx_Row+1
        inf.close()
        print('read in count is : '+ format(idx_Row))
        print('write out count is : '+ format(idx_Write))
        print('error out count is : '+ format(idx_Error))
    outf.close()
#===#
#===start main===#
#===#
const_URL_TOP = 'https://geoshape.ex.nii.ac.jp/ka/resource/'    #URL of ToDouFuKen
const_URL_PREFIX = 'https://geoshape.ex.nii.ac.jp'              #URL string prefix
const_OUTPUT_RESULT_DIR='htmls'                                 #Directory for Output file as result
const_OUTPUT_RESULT_TODOUFUKEN_FILE='ToDouFuKen.csv'
const_OUTPUT_RESULT_KYOUKAISEN_FILE='KyouKaiSen.csv'
const_OUTPUT_ERROR_FILE='KyouKaiSen_Error.csv'
#

###out_top='out_top.html'

#-----  pre check start
    #-----html folder is required

    #-----
#-----  pre check end
#===parse ToDouFuKen KyouKaiSen
if os.path.exists(const_OUTPUT_RESULT_DIR):
    parse_HTML_Todoufuken(const_URL_PREFIX,const_URL_TOP,const_OUTPUT_RESULT_DIR,const_OUTPUT_RESULT_TODOUFUKEN_FILE)
    parse_HTML_KyouKaiSen(const_OUTPUT_RESULT_DIR,const_OUTPUT_RESULT_TODOUFUKEN_FILE,const_OUTPUT_RESULT_DIR,const_OUTPUT_RESULT_KYOUKAISEN_FILE,const_OUTPUT_RESULT_DIR,const_OUTPUT_ERROR_FILE)
else:
    print('***** FAILED TO START. NOT EXIST FILE FOLDER,DIRECTORY : '+const_OUTPUT_RESULT_DIR)
#===#
#===end main===#
#===#