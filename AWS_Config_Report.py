###2020 플랫폼보안파트
import boto3
import json
import datetime
import requests
import time
import win32com.client
import os
import asyncio
from pywintypes import com_error
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill, Color, NamedStyle
from openpyxl.chart import BarChart, Reference
from openpyxl.styles.borders import Border, Side
from openpyxl.utils import get_column_letter
#pip install pywin32

client = '';

# Create Work Book(Whole Excel file)
wb = Workbook()
raw_wb = Workbook()
ws = wb.active
raw_ws_1 = raw_wb.active
raw_ws_1.title='전체 점검결과'
raw_ws_2 = raw_wb.create_sheet('미준수 데이터')
ws.title = 'init'
column_num = 3
ws_raw1_col = ws_raw2_col = 2
#compliant summary시 사용될 변수
#comp_num = 0
#noncomp_num = 0

#테두리 스타일
paintstyle = NamedStyle(name="border")
bd = Side(border_style='thin', color="000000")
paintstyle.border = Border(left=bd, top=bd, right=bd, bottom=bd)
paintstyle.alignment = Alignment(vertical='center')

#리소스 딕셔너리 정의
awsrsc={
    'AWS::Config::ResourceCompliance': 'AWS Config 정책 준수 상태의 리소스',
    'AWS::IAM::Role': 'AWS 계정에 대한 Role(역할) 리소스',
    'AWS::IAM::Policy': '자격 증명이나 리소스에 연결될 때 해당 권한을 정의하는 정책 리소스',
    'AWS::IAM::User': 'AWS 계정의 사용자 리소스',
    'AWS::EC2::SecurityGroup': '인스턴스에 대한 인바운드 및 아웃바운드 트래픽을 제어하는 가상 방화벽 역할의 보안 그룹 리소스',
    'AWS::EC2::NetworkInterface': 'EC2 인스턴스의 네트워크 인터페이스',
    'AWS::S3::Bucket': 'S3 버켓 리소스',
    'AWS::EC2::Subnet': 'VPC에서 사용될 서브넷 리소스',
    'AWS::KMS::Key': 'AWS KMS에서 사용하는 대칭키(symmetric customer master key, CMK)',
    'AWS::EC2::RouteTable': '라우팅 테이블과 연결된 서브넷을 떠나는 트래픽을 제어하는 라우팅 규칙 집합 리소스',
    'AWS::EC2::Volume': 'Amazon EBS(Elastic Block Store) 볼륨 리소스',
    'AWS::EC2::Instance': 'EC2 인스턴스',
    'AWS::IAM::Group': 'AWS 계정의 그룹 리소스',
    'AWS::EC2::InternetGateway': '네트워크를 인터넷에 연결하는 리소스',
    'AWS::EC2::VPC': '가상 사설 클라우드 리소스',
    'AWS::EC2::EIP': 'EC2 인스턴스와 연결 가능한 EIP(Elastic IP) 고정 IP 리소스',
    'AWS::EC2::NetworkAcl': 'VPC에 대한 네트워크 ACL 리소스',
    'AWS::Lambda::Function': '람다 함수 리소스',
    'AWS::EC2::VPCEndpoint': 'VPC와 서비스간의 사설 연결 엔드포인트 리소스',
    'AWS::EC2::NatGateway': '서브넷에서 사용될 NAT 게이트웨이 리소스',
    'AWS::RDS::DBSubnetGroup': '데이터베이스 서브넷 그룹 리소스',
    'AWS::AutoScaling::LaunchConfiguration': '오토 스케일링 그룹에서 EC2 인스턴스를 구성하는 데 사용하는 EC2 오토 스케일링 시작 구성 리소스',
    'AWS::Elasticsearch::Domain': 'Amazon ES(Elasticsearch, 검색엔진) 도메인 리소스',
    'AWS::EC2::FlowLog': '네트워크 인터페이스, 서브넷 또는 VPC에 대한 IP 트래픽을 캡쳐하는 흐름 로그 리소스',
    'AWS::RDS::DBSnapshot': 'DB인스턴스의 스토리지 볼륨 스냅샷 리소스',
    'AWS::CloudTrail::Trail': '지정한 Amazon S3 버킷에 이벤트를 제공하는 추적 리소스',
    'AWS::RDS::DBSecurityGroup': 'Amazon RDS DB 보안그룹 리소스',
    'AWS::SNS::Topic': '알림을 게시하는 주제(Topic) 리소스',
    'AWS::ElasticLoadBalancingV2::LoadBalancer': '네트워크 로드 밸런서 및 어플리케이션 로드 밸런서 리소스(NLB, ALB)',
    'AWS::SSM::ManagedInstanceInventory': 'SSM에서 AWS 계정의 관리형 인스턴스를 인벤토리하는 리소스',
    'AWS::WAFv2::WebACL': '웹 요청 검사 및 제어를 위한 규칙 모음 리소스',
    'AWS::CodePipeline::Pipeline':'Release 프로세스에 따른 소프트웨어 변경을 보여주는 파이프라인 리소스',
    'AWS::ACM::Certificate': '보안연결을 위한 AWS Certificate Manageer 인증서 리소스',
    'AWS::AutoScaling::AutoScalingGroup': 'EC2 오토 스케일링 리소스',
    'AWS::CodeBuild::Project':'AWS CodeBuild가 사용자의 소스코드를 어떻게 빌드할지 구성하는 프로젝트 리소스',
    'AWS::RDS::DBInstance':'Amazon RDS DB 보안그룹 리소스'
}

# Write Column Name
def writeColName(cName, row, ws):
    column_char = 'a'
    #write column name
    for name in cName:
        ws[column_char + str(row)] = name
        ws[column_char + str(row)].style = paintstyle
        ws[column_char + str(row)].alignment = Alignment(horizontal='center', vertical='center')
        column_char = chr(ord(column_char) + 1)

# Save Compliance Data to Excel
# 오버로딩이...
def saveContent(cName, ws, ws1=None, ws2 = None):
    column_char = 'a'
    global column_num, ws_raw1_col, ws_raw2_col
    #기본 엑셀
    for name in cName:
        ws[column_char + str(column_num)] = name
        ws[column_char + str(column_num)].style = paintstyle
        ws[column_char + str(column_num)].alignment = Alignment(vertical='center', wrap_text=True)

        if ws1 is not None and ws2 is None:
            ws1[column_char + str(ws_raw1_col)] = name
            ws1[column_char + str(ws_raw1_col)].style = paintstyle
            ws1[column_char + str(ws_raw1_col)].alignment = Alignment(vertical='center', wrap_text=True)

        elif ws2 is not None:
            ws1[column_char + str(ws_raw1_col)] = name
            ws1[column_char + str(ws_raw1_col)].style = paintstyle
            ws1[column_char + str(ws_raw1_col)].alignment = Alignment(vertical='center', wrap_text=True)
            ws2[column_char + str(ws_raw2_col)] = name
            ws2[column_char + str(ws_raw2_col)].style = paintstyle
            ws2[column_char + str(ws_raw2_col)].alignment = Alignment(vertical='center', wrap_text=True)

        column_char = chr(ord(column_char) + 1)
    column_num += 1
    if ws1 is not None and ws2 is None:
        ws_raw1_col += 1
    elif ws2 is not None:
        ws_raw1_col += 1
        ws_raw2_col += 1

#시트 첫행 타이틀 굵게, 가운데정렬, 배경색설정
def setTitleCell(cell):
    cell.font = Font(bold=True)
    cell.alignment = Alignment(horizontal='center', vertical='center')
    cell.fill = PatternFill(patternType='solid', fgColor=Color("CCCCCC"))
    return cell

#Page one for AWS Config Rule description(규칙 항목 및 설명)
def sheetOne(descConfigRulePgr):
    descConfig_iterator = descConfigRulePgr.paginate()
    global ws, column_num
    column_num = 3
    ws.merge_cells('A1:B1')
    ws['A1'] = "규칙 항목 및 설명"
    ws['A1'].style = paintstyle
    ws['B1'].style = paintstyle
    setTitleCell(ws.cell(row=1, column=1))
    ws.title = 'ConfigRuleList'
    # call export func by configrule for loop
    for configrule in descConfig_iterator:
        for rulename in configrule['ConfigRules']:
            # init column num for each Config Rule Result
            excelData = []
            excelData.append(rulename['ConfigRuleName'])
            excelData.append(rulename['Description'])
            saveContent(excelData, ws)

    #put the data title on sheet1
    excelColName = ['AWS Config 규칙명', 'Rule Description']
    writeColName(excelColName,2, ws)
    ws.row_dimensions[1].height = 25
    ws.column_dimensions['A'].width = 25
    ws.column_dimensions['B'].width = 45

    print('First sheet done')

#Sheet 2 for Evaluated resources - 평가된 리소스 현황
def sheetTwo():
    ws = wb.create_sheet('Evaluated Resources')
    discovered_rsc = client.get_disco
    vered_resource_counts()
    global column_num
    column_num = 3
    ws.merge_cells('A1:C1')
    ws['A1'] = "Evaluation Resources(평가 리소스 현황)"
    ws['A1'].style = paintstyle
    ws['B1'].style = paintstyle
    ws['C1'].style = paintstyle
    setTitleCell(ws.cell(row=1, column=1))
    # 리소스 설명칸 넣어서 총 개수 뺌
    #ws['A2'].style = paintstyle
    #ws['A2'].value = '총 리소스 수'
    #ws['B2'].style = paintstyle
    #ws['B2'].value = discovered_rsc['totalDiscoveredResources']

    for resource in discovered_rsc['resourceCounts']:
        excelData = []
        #print(resource)
        excelData.append(resource['resourceType'])
        try:
            excelData.append(awsrsc[resource['resourceType']])
        except:
            excelData.append('설명 추가 예정')
        excelData.append(resource['count'])
        saveContent(excelData, ws)

    excelColName = ['리소스 유형','설명', '합계']
    writeColName(excelColName, 2, ws)
    ws.row_dimensions[1].height = 25
    ws.column_dimensions['A'].width = 30
    ws.column_dimensions['B'].width = 35
    ws.column_dimensions['C'].width = 10

    print('Evaluated Resources sheet done')

#Sheet 3 for evaluation summary by config rule - Config Rule 별 평과 결과 막대그래프
def sheetThree():
    ws = wb.create_sheet('Evaluation Summary')
    ruleSummaryRes = client.get_compliance_summary_by_config_rule()
    global column_num
    column_num = 3
    ws.merge_cells('A1:D1')
    ws['A1'] = "Evaluation Summary(by config rule)"
    #border line

    comp = ruleSummaryRes['ComplianceSummary']['CompliantResourceCount']['CappedCount']
    noncomp = ruleSummaryRes['ComplianceSummary']['NonCompliantResourceCount']['CappedCount']
    #rows = [
    #    ('', '준수', '미준수', '전체'),
    #    ('Rule count',comp,noncomp,comp+noncomp)
    #]

    excelData = ['AWS Config rule 개수',comp,noncomp,comp+noncomp]
    saveContent(excelData, ws)

    #for row in rows:
    #    ws.append(row)

    bc = BarChart()
    bc.type = "col"
    bc.style = 10
    bc.title = "Evaluation Summary"
    bc.y_axis.title = 'Rule Count'
    bc.x_axis.title = 'Compliant Status'

    data = Reference(ws, min_col=2, min_row=1, max_row=4, max_col=3)
    cats = Reference(ws, min_col=1, min_row=2, max_row=4)
    bc.add_data(data, titles_from_data=True)
    bc.set_categories(cats)
    bc.shape = 4
    bc.width = 8
    bc.legend = None
    ws.add_chart(bc, "A6")

    rownum=0
    while rownum < 3:
        ws.cell(row=rownum+1, column=1).style = paintstyle
        ws.cell(row=rownum+1, column=2).style = paintstyle
        ws.cell(row=rownum+1, column=3).style = paintstyle
        ws.cell(row=rownum + 1, column=4).style = paintstyle
        rownum += 1
    setTitleCell(ws.cell(row=1, column=1))
    excelColName = ['', '준수', '미준수', '전체']
    writeColName(excelColName,2,ws)
    ws.row_dimensions[1].height = 25
    ws.column_dimensions['A'].width = 20
    ws.column_dimensions['B'].width = 15
    ws.column_dimensions['C'].width = 15

    print('Evaluation Summary sheet done')

#sheet 4 for numbers of compliance result - 평가 항목 별 compliant 요약
def sheetFour(descConfigRulePgr, compDetailPgr):
    descConfig_iterator = descConfigRulePgr.paginate()
    # Worksheet 생성
    global ws, column_num
    comp_num = noncomp_num = 0
    excelData=[]
    rulenametemp =''
    column_num = 3
    ws.merge_cells('A1:E1')
    ws['A1'] = "평가 항목 별 요약"
    #borderline
    ws['A1'].style = paintstyle
    ws['B1'].style = paintstyle
    ws['C1'].style = paintstyle
    ws['D1'].style = paintstyle
    ws['E1'].style = paintstyle
    setTitleCell(ws.cell(row=1, column=1))
    ws.title = 'Result by Rule'

    #describe rule의 rule name 추출해서 리소스들 get comp detail 정보 추출
    for configrule in descConfig_iterator:
        for rulename in configrule['ConfigRules']:
            responseConfigPage = compDetailPgr.paginate(
                ConfigRuleName=rulename['ConfigRuleName']
            )
            # 룰별 객체
            for page in responseConfigPage:
                try:
                    if not rulenametemp :
                        rulenametemp = page['EvaluationResults'][0]['EvaluationResultIdentifier']['EvaluationResultQualifier']['ConfigRuleName']
                        excelData.append(rulenametemp)
                        excelData.append(rulename['Description'])

                    for user in page['EvaluationResults']:
                        if user['ComplianceType'] == 'NON_COMPLIANT':
                            noncomp_num += 1
                        elif user['ComplianceType'] == 'COMPLIANT':
                            comp_num += 1
                    try:
                        #NextToken 값 존재(pagination)
                        if page['NextToken'] in page:
                            pass
                    except:
                        #룰네임존재, 리소스 pagination 아님 엑셀저장
                        excelData.append(comp_num)
                        excelData.append(noncomp_num)
                        excelData.append(comp_num+noncomp_num)
                        saveContent(excelData, ws)
                        comp_num = 0
                        noncomp_num = 0
                        rulenametemp = ''
                        excelData = []

                except:
                    if(rulenametemp == '') :
                        continue
                    else :
                        #마지막 페이지, nexttoken 없음 엑셀저장
                        excelData.append(comp_num)
                        excelData.append(noncomp_num)
                        excelData.append(comp_num + noncomp_num)
                        saveContent(excelData, ws)
                        comp_num = 0
                        noncomp_num = 0
                        rulenametemp =''
                        excelData = []

    #put the data title on sheet1
    excelColName = ['Config 규칙명', '설명', '충족', '불충족', '합계']
    writeColName(excelColName,2, ws)
    ws.row_dimensions[1].height = 25
    ws.column_dimensions['A'].width = 20
    ws.column_dimensions['B'].width = 33
    ws.column_dimensions['C'].width = 5
    ws.column_dimensions['D'].width = 7
    ws.column_dimensions['E'].width = 5
    #리소스 없는 룰 리스트 제외
    #ws['A'+str(ws.max_row+2)] = '평가되는 리소스가 없는 룰은 제외됩니다.'

    print('Compliance result by rule sheet done')

#aws 설정 파일 생성
async def makeConfigFile():
    try:
        if not (os.path.isdir(os.environ['USERPROFILE'] + '\.aws')):
            os.makedirs(os.path.join(os.environ['USERPROFILE'] + '\.aws'))
            print("Creating aws directory...")
    except OSError as e:
        if e.errno != errno.EEXIST:
            print("Failed to create directory")
            raise

    path = os.environ['USERPROFILE'] + r'\.aws\credentials'
    credentialsfile = open(path, 'w')
    path = os.environ['USERPROFILE'] + r'\.aws\config'
    configfile = open(path, 'w')
    print('Please put your account information(press enter)')
    print('######################################')
    # 컨피그 파일 입력용 문자열 생성
    configstr = []
    credstr = []

    configstr.append("[default]\n")
    credstr.append("[default]\n")
    # credentials 파일
    credstr.append("aws_access_key_id = " + input("aws_access_key_id = "))
    credstr.append("\naws_secret_access_key = " + input("aws_secret_access_key = "))
    # config 파일
    configstr.append("region = " + input("region = "))
    configstr.append("\noutput = json")
    print('######################################')

    #3줄씩 설정 파일 출력
    for i in range(len(credstr)):
        credentialsfile.write(credstr[i])
        configfile.write(configstr[i])

    credentialsfile.close()
    configfile.close()

async def mk_raw_data(configrulename, desc, compDetailPgr):

    global raw_wb
    raw_ws = raw_wb.create_sheet('Config Result')
    global column_num
    column_num = 2

    #if(raw_ws.title == 'init') :
    #    raw_ws.title = 'Config Result'
    #else :
    #    raw_ws = raw_wb.create_sheet('Config Result')

    #get detail paginator
    responseConfigPage = compDetailPgr.paginate(
        ConfigRuleName=configrulename
    )

    for page in responseConfigPage:
        for user in page['EvaluationResults']:
            excelData = []
            excelData.append(user['EvaluationResultIdentifier']['EvaluationResultQualifier']['ConfigRuleName'])
            excelData.append(user['EvaluationResultIdentifier']['EvaluationResultQualifier']['ResourceType'])
            excelData.append(str(user['EvaluationResultIdentifier']['OrderingTimestamp'])[:19])
            username = usernameDic.get(user['EvaluationResultIdentifier']['EvaluationResultQualifier']['ResourceId'])
            if not username:
                excelData.append('N/A')
            else:
                excelData.append(username)
            excelData.append(user['EvaluationResultIdentifier']['EvaluationResultQualifier']['ResourceId'])
            excelData.append(user['ComplianceType'])
            #미준수인경우
            if(user['ComplianceType']=='NON_COMPLIANT'):
                excelData.append(desc)
                saveContent(excelData, raw_ws, raw_ws_1, raw_ws_2)
            else:
                excelData.append('N/A')
                saveContent(excelData, raw_ws, raw_ws_1)

    max_col = raw_ws.max_column
    i=1;
    for i in range(i, max_col):
        raw_ws.column_dimensions[get_column_letter(i)].width = 20
    raw_ws.column_dimensions['G'].width = 70

    excelColName = ['AWS Config 규칙명', '리소스 타입', '평가시간','UserName(IAM::User)','ResourceId','컴플라이언스 결과', 'Rule Description']
    writeColName(excelColName,1, raw_ws)

###########################################Main############################################
print('creating excel file..')

#aws programmatic login
eventLoop = asyncio.get_event_loop()
eventLoop.run_until_complete(makeConfigFile())

#create config client
client = boto3.client('config')

#필요한 paginator들 호출
descConfigRulePgr = client.get_paginator('describe_config_rules')
descConfig_iterator = descConfigRulePgr.paginate()
compDetailPgr = client.get_paginator('get_compliance_details_by_config_rule')
userinfo = client.list_discovered_resources(
    resourceType='AWS::IAM::User'
)

###raw data 엑셀 작성
usernameDic = {}
for user_info in userinfo['resourceIdentifiers']:
    usernameDic[user_info['resourceId']] = user_info['resourceName']

#raw data 요약 페이지
excelColName = ['AWS Config 규칙명', '리소스 타입', '평가 시간', 'UserName(IAM::User)', 'ResourceId',
                '컴플라이언스 결과', 'Rule Description']
writeColName(excelColName, 1, raw_ws_1)
writeColName(excelColName, 1, raw_ws_2)

max_col = raw_ws_1.max_column
i=2
for i in range(i, max_col):
    raw_ws_1.column_dimensions[get_column_letter(i)].width = 25
    raw_ws_2.column_dimensions[get_column_letter(i)].width = 25
raw_ws_1.column_dimensions['A'].width = 35
raw_ws_2.column_dimensions['A'].width = 35
raw_ws_1.column_dimensions['G'].width = 70
raw_ws_2.column_dimensions['G'].width = 70

#call export func by configrule for loop
print('Export raw data..')
for configrule in descConfig_iterator:
    for rulename in configrule['ConfigRules']:
        #init column num for each Config Rule Result
        eventLoop.run_until_complete(mk_raw_data(rulename['ConfigRuleName'], rulename['Description'], compDetailPgr))

raw_wb.save("AWS_Config_Raw_Data.xlsx")
###
eventLoop.close()

###요약 레포트 엑셀 작성(총 sheet 4장)
#sheetOne(descConfigRulePgr)
sheetFour(descConfigRulePgr, compDetailPgr)
sheetTwo()
sheetThree()

time.sleep(1)
wb.save("aws-config-report-temp.xlsx")
print('Sheets all done')

PATH_TO_PDF = 'AWS_Config_Report.pdf'

excel = win32com.client.Dispatch("Excel.application")

try:
    print('Starting convert to PDF')
    wb=excel.Workbooks.Open(os.path.abspath('aws-config-report-temp.xlsx'))
    time.sleep(1)
    ws_list = ['Result by Rule', 'Evaluated Resources','Evaluation Summary']
    wb.WorkSheets(ws_list).Select()
    excel.ActiveSheet.ExportAsFixedFormat(0, os.path.abspath(PATH_TO_PDF))

except com_error as e:
    print('converting failed.')
    print(e)
else:
    print('PDF file saved')
finally:
    wb.Close(False)
    excel.Quit()

os.remove(os.path.abspath("aws-config-report-temp.xlsx"))
###fortest os.remove(os.environ['USERPROFILE'] + r'\.aws\credentials')
###fortest os.remove(os.environ['USERPROFILE'] + r'\.aws\config')
###fortest os.rmdir(os.environ['USERPROFILE'] + r'\.aws')

print("Automatically closing...")
###########################################################################################