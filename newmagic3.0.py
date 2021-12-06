import dart_fss as dart

from urllib.request import urlopen
from io import BytesIO
from zipfile import ZipFile

import xml.etree.ElementTree as ET

import os
import openpyxl

import requests
import json

#날짜계산
from datetime import datetime

api_key = '50b58aad9293511aa5537749f773d90a2f1134aa'
url = 'https://opendart.fss.or.kr/api/corpCode.xml?crtfc_key=50b58aad9293511aa5537749f773d90a2f1134aa'
thisyear = datetime.now().year
lastyear = thisyear-1

#회사고유번호 추출
with urlopen(url) as zipresp:
    with ZipFile(BytesIO(zipresp.read())) as zfile:
        zfile.extractall('corp_num')
tree = ET.parse('.\corp_num\CORPCODE.xml')
root = tree.getroot()

#기존 회사고유번호 엑셀파일 삭제
if os.path.isfile('.\\fsdata\corp_num.xlsx'):
    os.remove('.\\fsdata\corp_num.xlsx')

#회사 고유번호 엑셀파일 생성
wb = openpyxl.Workbook()
xlname = '.\\fsdata\corp_num.xlsx'

wb.active.title = '회사 고유번호'
wb.active.append(['고유번호','회사명','주가총액','자본총계','매출총이익','당기순이익','부채총계','매출액','매출원가','판관비','비영업자산'])


dart_comp_url = 'https://opendart.fss.or.kr/api/fnlttSinglAcntAll.json?crtfc_key='
corp_code = '00126380'
reprt_codes_dict = {'11011':'사업','11012':'반기','11013':'1분기','11014':'3분기'}
reprt_codes = ['11011','11012','11013','11014']
fs_div = '&fs_div=CFS'
bsns_years = [lastyear, thisyear]
#비영업자산을 세분화한 그림보고 다 넣어야함 나중에 합시다 이건.
account_nm = ['주가총액','자본총계','자산총계','매출총이익','당기순이익','부채총계','매출액','매출원가','판관비']


for corp in root.iter("list"):
    corp_code = corp.findtext("corp_code")
    # wb.active.append(corp_code.split())
    corp_code_txt = '&corp_code=' + corp_code
    temp_data_dict = {}
    for bsns_year in bsns_years:
        for reprt_code in reprt_codes:
            res = requests.get(dart_comp_url + api_key + corp_code_txt + '&bsns_year=' + bsns_year + '&reprt_code=' + reprt_code + fs_div)
            item = json.loads(res.text)
            for _item in item.get("list"):
                if _item['account_nm'] in account_nm:
                    temp_data_dict[_item['account_nm']] = _item['thstrm_amount']
                    wb.active.append() #계좌명:값 dict를 엑셀에 넣는 방법을 고민해야함
                    print(_item['thstrm_amount'])
    # print(corp.findtext("corp_code"))
wb.save(xlname)


############################
#corp_code에는 위에서 추출한 각 기업 순서대로 하나씩 넣고
#bsns_year에는 현재 분기 고려해서 작년 또는 올해 넣고(4개 분기를 받아와야 함)
#reprt_code에는 bsns_year고려해서 분기 잘 넣고
#fs_div는 연결재무재표(CFS)고정
#account_nm은 자본총계,매출총이익,자산총계,당기순이익,부채,현금,비영업자산,매출원가,판매비,관리비 등등..
############################
