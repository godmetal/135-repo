import boto3
import sys, traceback
import json
import logging
import time, datetime
import math

logger = logging.getLogger()
logger.setLevel(logging.INFO)

# put the id of your web acl here
_WebACLId = 'ade82d9d-0852-47bd-822a-fdb7587db3b4'
# set the interval between two invocations here
_Interval = 15
# a negative number stands for the gap between current time and the end time of the log window.
# e.g. if you set this to -30 and _Interval to 15, current time is 06:00, then this lambda function will retrieve log between 05:15 and 05:30.
# The reason we have this parameter is that there is always about 5 mins delay before the WAF log been updated.
_OFFSET = -5

_API_CALL_NUM_RETRIES = 3
_WINDOW = datetime.timedelta(minutes=_Interval)
_NOW = datetime.datetime.now() + datetime.timedelta(minutes=_OFFSET)
_DATE = datetime.datetime.now().strftime('%Y-%m-%d')
_TIMEFORMAT = '%Y-%m-%dT%H:%M:%S%z'


def trace_exception():
    print
    '=' * 80
    traceback.print_exc(file=sys.stdout)
    print
    '=' * 80
    exit()


def log_serialize(log):
    log_events = []

    if len(log):
        for l in log:
            timestamp = int(
                (l['Timestamp'].replace(tzinfo=None) - datetime.datetime(1970, 1, 1)).total_seconds() * 1000)
            l['Timestamp'] = l['Timestamp'].strftime(_TIMEFORMAT)
            log_events.append({'timestamp': timestamp, 'message': json.dumps(l)})

        # print log_events
        return sorted(log_events, key=lambda event: event['timestamp'])
    else:
        return []


def get_web_acl(client, acl_id):
    #3번동안 attempt
    for attempt in range(_API_CALL_NUM_RETRIES):
        try:
            #WebACL 배열값 반환(WAF 룰 id)
            return client.get_web_acl(Name='INYONG_TEST_WAF_Cloud-Front', Scope='CLOUDFRONT', Id=acl_id)['WebACL']

        except Exception:
            trace_exception()
            #실패했을때 쉬었다가 2번까지 더 재시도
            if attempt < _API_CALL_NUM_RETRIES:
                delay = math.pow(2, attempt)
                print
                u'[get_web_acl] Retrying in %d seconds...' % (delay)
                time.sleep(delay)

    else:
        #3번 실패 후 빈배열반환
        print
        u'[get_web_acl] Failed ALL attempts to call API'
        return {}


def retrieve_log(acl_id):
    waf_client = boto3.client('wafv2')

    #리턴 client.get_web_acl(WebACLId=acl_id)['WebACL']
    webacl = get_web_acl(waf_client, acl_id)
    rules = ['Default_Action']
    log = {}
    #logger.info(webacl['ARN'])
    #WAF내 모든 webacl 룰 id 배열저장
    for rule in webacl['Rules']:
        rules.append(rule['Name'])

    #룰 이름들 배열에서 룰꺼냄
    for rule in rules:
        for attempt in range(_API_CALL_NUM_RETRIES):
            try:
                #3번시도
                # print _NOW - _WINDOW
                # print _NOW
                #특정 시간(15분 동안) 내  첫 5000개 중 n 개 리턴 - 500개 리턴
                response = waf_client.get_sampled_requests(
                    WebAclArn='arn:aws:wafv2:us-east-1:747935822721:global/webacl/INYONG_TEST_WAF_Cloud-Front/ade82d9d-0852-47bd-822a-fdb7587db3b4',
                    RuleMetricName=rule,
                    Scope='CLOUDFRONT',
                    TimeWindow={
                        'StartTime': _NOW - _WINDOW,
                        'EndTime': _NOW
                    },
                    MaxItems=500
                )
                #룰 이름별 로그 추출
                log[rule] = log_serialize(response['SampledRequests'])

                break

            except Exception:
                trace_exception()

                if attempt < _API_CALL_NUM_RETRIES:
                    delay = math.pow(2, attempt)
                    print
                    u'[get_sampled_requests] Retrying in %d seconds...' % (delay)
                    time.sleep(delay)

        else:
            print
            u'[get_sampled_requests] Failed ALL attempts to call API'
    # print log
    logger.info(log)
    return log


def push_log(log={}):
    log_client = boto3.client('logs')
    log_groups = []
    log_name = 'WebACL-' + _WebACLId

    for attempt in range(_API_CALL_NUM_RETRIES):
        try:
            log_groups = log_client.describe_log_groups(
                logGroupNamePrefix=log_name,
            )['logGroups']
            # print log_groups
            break

        except Exception:
            trace_exception()

            if attempt < _API_CALL_NUM_RETRIES:
                delay = math.pow(2, attempt)
                print
                u'[describe_log_groups] Retrying in %d seconds...' % (delay)
                time.sleep(delay)

    else:
        print
        u'[describe_log_groups] Failed ALL attempts to call API'
        return

    if not log_groups:
        for attempt in range(_API_CALL_NUM_RETRIES):
            try:
                log_client.create_log_group(
                    logGroupName=log_name,
                    tags={
                        'Name': 'waf_log'
                    }
                )

                break

            except Exception:
                trace_exception()

                if attempt < _API_CALL_NUM_RETRIES:
                    delay = math.pow(2, attempt)
                    print
                    u'[create_log_group] Retrying in %d seconds...' % (delay)
                    time.sleep(delay)
        else:
            print
            u'[create_log_group] Failed ALL attempts to call API'
            return

    for rulenametmp, log_events in log.items():
        # print log_events
        if not log_events:
            continue

        rulename = 'Default_Action'
        if rulenametmp != 'Default_Action':
            rulename = rulenametmp

        stream_name = 'Rule [%s] %s' % (_DATE, rulename)

        log_streams = []
        nexttoken = ''

        for attempt in range(_API_CALL_NUM_RETRIES):
            try:
                log_streams = log_client.describe_log_streams(
                    logGroupName=log_name,
                    logStreamNamePrefix=stream_name,
                    orderBy='LogStreamName',
                    descending=False,
                )['logStreams']

                # print log_streams

                if log_streams and 'uploadSequenceToken' in log_streams[0]:
                    nexttoken = log_streams[0]['uploadSequenceToken']

                break

            except Exception:
                trace_exception()

                if attempt < _API_CALL_NUM_RETRIES:
                    delay = math.pow(2, attempt)
                    print
                    u'[describe_log_streams] Retrying in %d seconds...' % (delay)
                    time.sleep(delay)

        else:
            print
            u'[describe_log_streams] Failed ALL attempts to call API'
            return

        if not log_streams:
            for attempt in range(_API_CALL_NUM_RETRIES):
                try:
                    log_client.create_log_stream(
                        logGroupName=log_name,
                        logStreamName=stream_name
                    )

                    # print 'created ' + stream_name
                    break

                except Exception:
                    trace_exception()

                    if attempt < _API_CALL_NUM_RETRIES:
                        delay = math.pow(2, attempt)
                        print
                        u'[create_log_stream] Retrying in %d seconds...' % (delay)
                        time.sleep(delay)

            else:
                print
                u'[create_log_stream] Failed ALL attempts to call API'
                continue

        for attempt in range(_API_CALL_NUM_RETRIES):
            try:
                # print log_name, stream_name
                if nexttoken:
                    # print nexttoken
                    print
                    log_client.put_log_events(
                        logGroupName=log_name,
                        logStreamName=stream_name,
                        logEvents=log_events,
                        sequenceToken=nexttoken
                    )
                else:
                    print
                    log_client.put_log_events(
                        logGroupName=log_name,
                        logStreamName=stream_name,
                        logEvents=log_events,
                    )

                break
            except Exception:
                trace_exception()

                if attempt < _API_CALL_NUM_RETRIES:
                    delay = math.pow(2, attempt)
                    print
                    u'[put_log_events] Retrying in %d seconds...' % (delay)
                    time.sleep(delay)

        else:
            print
            u'[put_log_events] Failed ALL attempts to call API'

    return


def lambda_handler(event, context):
    # TODO implement
    return push_log(retrieve_log(_WebACLId))