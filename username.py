
import boto3
import json
import datetime
import requests

#create config client
client = boto3.client('config')

#################################Main############################################
#call config rules

userinfo = client.list_discovered_resources(
    resourceType='AWS::IAM::User'
)
#just print response
count =0
#call slack func by configrule for loop
for user in userinfo['resourceIdentifiers']:
    print(user)
    if(user['resourceId']== 'AIDA24JDS6OAYMFEIZR7E'):
        print(user['resourceName'])
    count += 1

print(usernameDic)