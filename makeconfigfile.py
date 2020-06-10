import os

try:
    if not(os.path.isdir(os.environ['USERPROFILE']+'\.aws')):
        os.makedirs(os.path.join(os.environ['USERPROFILE']+'\.aws'))
except OSError as e:
    if e.errno != errno.EEXIST:
        print("Failed to create directory")
        raise

path = os.environ['USERPROFILE'] + r'\.aws\credentials'
credentialsfile = open(path, 'w')
path = os.environ['USERPROFILE'] + r'\.aws\config'
configfile = open(path, 'w')

#컨피그 파일 입력용 문자열 생성
configstr = []
credstr = []

configstr.append("[default]\n")
credstr.append("[default]\n")
#credentials 파일
credstr.append("aws_access_key_id = " + input("aws_access_key_id = "))
credstr.append("\naws_secret_access_key = " + input("aws_secret_access_key = "))
#config 파일
configstr.append("region = " + input("region = "))
configstr.append("\noutput = json")

for i in range(len(credstr)):
    credentialsfile.write(credstr[i])

for i in range(len(configstr)):
    configfile.write(configstr[i])

credentialsfile.close()
configfile.close()

