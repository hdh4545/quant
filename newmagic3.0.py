import dart_fss as dart

from urllib.request import urlopen
from io import BytesIO
from zipfile import ZipFile

import xml.etree.ElementTree as ET

from openpyxl import Workbook

api_key = '50b58aad9293511aa5537749f773d90a2f1134aa'
url = 'https://opendart.fss.or.kr/api/corpCode.xml?crtfc_key=50b58aad9293511aa5537749f773d90a2f1134aa'

# dart.set_api_key(api_key=api_key)
#
# corp_list = dart.get_corp_list()
#
# samsung = corp_list.find_by_corp_name('삼성전자', exactly=True)[0]
#
# fs = samsung.extract_fs(bgn_de='20190101')
#
# fs.save()

#회사고유번호 추출
with urlopen(url) as zipresp:
    with ZipFile(BytesIO(zipresp.read())) as zfile:
        zfile.extractall('corp_num')
tree = ET.parse('.\corp_num\CORPCODE.xml')
root = tree.getroot()

#회사 고유번호 엑셀저장

for corp in root.iter("list"):
    print(corp.findtext("corp_code"))