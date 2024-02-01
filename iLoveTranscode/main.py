# coding=utf-8

import time
import uuid

import DaVinciResolveScript as dvr
import paho.mqtt.client as mqtt
import json
import qrcode
from PIL import Image
from enum import IntEnum, unique


class bcolors:
    HEADER = '\033[95m'
    OKBLUE = '\033[94m'
    OKCYAN = '\033[96m'
    OKGREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    ENDC = '\033[0m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'


@unique
class JobStatus(IntEnum):
    Ready = 0
    Rendering = 1
    Canceled = 2
    Failed = 3
    Finished = 4
    Unknown = 5

    @classmethod
    def from_string(cls, status_string):
        if status_string == '就绪' or status_string == 'Ready':
            return cls.Ready
        elif status_string == '渲染' or status_string == 'Rendering':
            return cls.Rendering
        elif status_string == '已取消' or status_string == 'Cancelled':
            return cls.Canceled
        elif status_string == '失败' or status_string == 'Failed':
            return cls.Failed
        elif status_string == '完成' or status_string == 'Complete':
            return cls.Finished
        else:
            return cls.Unknown



# Press the green button in the gutter to run the script.
if __name__ == '__main__':

    resolve = dvr.scriptapp("Resolve")
    projectManager = resolve.GetProjectManager()
    projectFromResolve = projectManager.GetCurrentProject()
    projectName = projectFromResolve.GetName()
    projectID = projectFromResolve.GetUniqueId()
    print(f'{projectName} - {projectID}')

    broker_address = "test.mosquitto.org"
    broker_port = 1883

    publish_topic = f"iLoveTranscode/{projectID}"
    qrcode_infomation = {
        "projectName": projectName,
        "brokerAddress": broker_address,
        "brokerPort": broker_port,
        "topicAddress": publish_topic
    }

    qr = qrcode.QRCode()
    qr.add_data(json.dumps(qrcode_infomation, ensure_ascii=False).encode('utf8'))
    print(json.dumps(qrcode_infomation))
    img = qr.make_image()
    img.show()


    json_data = {
        "projectName": projectName,
        "projectID": projectID,
        "status": "OK"
    }

    def on_publish(client, userdata, mid):
        # print("Message Published")
        i = 0

    def on_message(client, userdata, msg):
        message = msg.payload.decode()
        if not message.startswith('req@'):
            return
        jobId = message.split('@')[-1]
        print(f"Received request of job `{jobId}` from `{msg.topic}`")

        def is_request_job(job):
            return job['JobId'] == jobId
        try:
            requestJob = list(filter(is_request_job, projectFromResolve.GetRenderJobList()))[0]
        except IndexError:
            return

        jobDetails = {
            'id': requestJob['JobId'],
            'td': requestJob['TargetDir'],
            'v': requestJob['IsExportVideo'],
            'a': requestJob['IsExportAudio'],
            'w': requestJob['FormatWidth'],
            'h': requestJob['FormatHeight'],
            'fr': requestJob['FrameRate'],
            'pa': requestJob['PixelAspectRatio'],
            'abd': requestJob['AudioBitDepth'],
            'asr': requestJob['AudioSampleRate'],
            'ea': requestJob['ExportAlpha'],
            'ofn': requestJob['OutputFilename'],
            'rm': requestJob['RenderMode'],
            'pn': requestJob['PresetName'],
            'vf': requestJob['VideoFormat'],
            'vc': requestJob['VideoCodec'],
            'ac': requestJob['AudioCodec']
        }
        json_payload = json.dumps(jobDetails, ensure_ascii=False).encode('utf8')
        result = client.publish(publish_topic, json_payload)
        print(requestJob)

    def on_connect(client, userdata, flags, rc, _):
        if rc == 0:
            print("Connected to MQTT Broker!")
        else:
            print("Failed to connect, return code %d\n", rc)
        client.subscribe(f'{publish_topic}/inverse')

    # projectFromResolve.StartRendering(isInteractiveMode=False)

    client = mqtt.Client(projectID, protocol=mqtt.MQTTv5)
    client.on_message = on_message
    client.on_connect = on_connect
    client.on_publish = on_publish

    client.connect(broker_address, broker_port, 60)
    client.loop_start()

    while True:
        renderJobsInfo = []
        if not projectManager.GetCurrentProject().GetUniqueId() == projectID:
            break
        renderJobs = projectFromResolve.GetRenderJobList()
        for index, job in enumerate(renderJobs):
            jobId = job['JobId']
            jobName = job['RenderJobName']
            timelineName = job['TimelineName']
            status = projectFromResolve.GetRenderJobStatus(jobId)
            jobStatus = JobStatus.from_string(status['JobStatus'])
            jobProgress = status['CompletionPercentage']
            estimatedTime = 0
            try:
                estimatedTime = status['EstimatedTimeRemainingInMs']
            except KeyError:
                pass
            timeTaken = 0
            try:
                timeTaken = status['TimeTakenToRenderInMs']
            except KeyError:
                pass

            jobInfo = {
                'id': jobId,
                'jn': jobName,
                'tn': timelineName,
                'js': jobStatus,
                'jp': jobProgress,
                'et': estimatedTime,
                'tt': timeTaken,
                'od': index
            }
            json_payload = json.dumps(jobInfo, ensure_ascii=False).encode('utf8')
            result = client.publish(publish_topic, json_payload)
            renderJobsInfo.append(jobInfo)

        def is_ready(f_job):
            return f_job['js'] == JobStatus.Ready
        readyJobs = list(filter(is_ready, renderJobsInfo))

        def is_rendering(f_job):
            return f_job['js'] == JobStatus.Rendering
        currentJobs = list(filter(is_rendering, renderJobsInfo))
        if not currentJobs:
            try:
                currentJobs = [renderJobsInfo[-1]]
            except IndexError:
                pass
        currentJobId = ''
        try:
            currentJobId = currentJobs[0]['id']
        except IndexError:
            pass

        def is_finish(f_job):
            return f_job['js'] == JobStatus.Finished
        finishJobs = list(filter(is_finish, renderJobsInfo))

        def is_failed(f_job):  # 失败和取消的任务都归结到这里
            return f_job['js'] == JobStatus.Failed or f_job['js'] == JobStatus.Canceled
        failedJobs = list(filter(is_failed, renderJobsInfo))

        summary_info = {
            'rjn': len(readyJobs),
            'fjn': len(failedJobs),
            'fnjn': len(finishJobs),
            'cj': currentJobId,
            'ir': projectFromResolve.IsRenderingInProgress()
        }
        json_payload = json.dumps(summary_info, ensure_ascii=False).encode('utf8')
        result = client.publish(publish_topic, json_payload)

            # publish_info.append(jobInfo)

        # json_payload = json.dumps(publish_info, ensure_ascii=False).encode('utf8')
        # print(json_payload)
        # result = client.publish(publish_topic, json_payload)
        # result: [0, 1]
        # status = result[0]
        # if status == 0:
        #     # print(f"Send `{json_payload}` to topic `{publish_topic}`")
        #     i = 0
        # else:
        #     print(f"Failed to send message to topic {publish_topic}")

        time.sleep(1)

    client.loop_stop()
    client.disconnect()
