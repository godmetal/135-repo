
import boto3
import json
import datetime
import requests

#create config client
client = boto3.client('config')
webhook_url = "https://hooks.slack.com/services/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"

def sendMsgToSlack(payload):
    requests.post(
        webhook_url, data=json.dumps(payload),
        headers={'Content-Type': 'application/json'}
    )

#sending slack using configRuleName
def slackByCfgRuleName(configrulename):
    print(configrulename)
    content = ''
    content += '------------------------------\n'
    content += str(datetime.datetime.now()) + '\n'
    content += 'Config Rule Name :\n'
    content += '[' + configrulename+']\n'
    content += '------------------------------'
    #get detail paginator
    paginator = client.get_paginator('get_compliance_details_by_config_rule')
    responseConfigPage = paginator.paginate(
        ConfigRuleName=configrulename
    )
    for page in responseConfigPage:
        for user in page['EvaluationResults']:
            resourceid = user['EvaluationResultIdentifier']['EvaluationResultQualifier']['ResourceId']
            content += '\nResourceName : '
            content += str(usernameDic.get(resourceid))
            content += '\nResourceId : '
            content += resourceid
            content += '\nComplianceType : '
            content += user['ComplianceType'] + '\n'
    payload = {"text": content}
    sendMsgToSlack(payload)

#################################Main############################################
#call config rules
rulepaginator = client.get_paginator('describe_config_rules')
response_iterator = rulepaginator.paginate()
userinfo = client.list_discovered_resources(
    resourceType='AWS::IAM::User'
)

usernameDic = {}
for user_info in userinfo['resourceIdentifiers']:
    usernameDic[user_info['resourceId']] = user_info['resourceName']


#call slack func by configrule for loop
for configrule in response_iterator:
    for rulename in configrule['ConfigRules']:
        slackByCfgRuleName(rulename['ConfigRuleName'])

print("done")