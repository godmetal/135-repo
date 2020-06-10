# import boto3

# # Create IAM client
# iam = boto3.client('iam')

# # List users with the pagination interface
# paginator = iam.get_paginator('list_users')
# for response in paginator.paginate():
    # print(response)

import json
import os
os.system('cmd /k "aws configservice get-compliance-details-by-config-rule --config-rule-name CustomConfigRuleTest"')
# 호출된 결과를 json 형태로 저장
json_data = json.loads(con.text)
    # 결과를 담을 배열 초기화
detailCompliance = []
print('******************************')
  # 각 결과 값을 배열에 저장
detailCompliance.append(json_data['EvaluationResults']['EvaluationResultIdentifier'])
#  detailCompliance.append(json_data['whois']['countryCode'])
  
print(detailCompliance)