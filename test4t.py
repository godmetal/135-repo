
import boto3
import json
import datetime
import requests
from openpyxl import Workbook
from openpyxl import load_workbook

#create config client
client = boto3.client('config')

#################################Main############################################
#call config rules
response = client.describe_config_rules()
paginator = client.get_paginator('describe_config_rules')
response_iterator = paginator.paginate()


#just print response
print(response)

#call slack func by configrule for loop
for configrule in response_iterator:
    for configrulename in configrule['ConfigRules']:
        print(configrulename['ConfigRuleName'])

