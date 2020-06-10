
import boto3
import json
import requests

webhook_url = "https://hooks.slack.com/services/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"

#create config client
client = boto3.client('config')
paginator = client.get_paginator('get_compliance_details_by_config_rule')

# List users with the pagination interface
response_i = paginator.paginate(
    ConfigRuleName='CustomConfigRuleTest'
)
count=0


for page in response_i:
    print(page)
    for user in page['EvaluationResults']:
        print(user['EvaluationResultIdentifier']['EvaluationResultQualifier']['ResourceId'])
        print(user['ComplianceType'])
        count += 1


print(count)

content = "from python"
content += ' test3.py'
payload = {"text": content}

requests.post(
    webhook_url, data=json.dumps(payload),
    headers={'Content-Type': 'application/json'}
)

#print(response_iterator)
