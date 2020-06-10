
import boto3
import json
import datetime
import requests


webhook_url = "https://hooks.slack.com/services/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"

#create config client
client = boto3.client('config')
iamclient = boto3.client('iam')
paginator = client.get_paginator('get_compliance_details_by_config_rule')

# List users with the pagination interface
response_i = paginator.paginate(
    ConfigRuleName='CustomConfigRuleTest'
)
count=0


response = iamclient.list_users()
#print(count)
print(response['Users'])
response['date'] = datetime.datetime.now()

def myconverter(o):
    if isinstance(o, datetime.datetime):
        return o.__str__()

content = "from python test1.py\n"
content = response
payload = {"text": content}

requests.post(
    webhook_url, data=json.dumps(payload),
    headers={'Content-Type': 'application/json'}
)

#print(response_iterator)
