# -*- coding: utf-8 -*-

import DaVinciResolveScript as dvr
import os
# import csv
from datetime import datetime
import xlsxwriter as xls
import shutil
import sys


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


def calculate_timecode(start_timecode, frame_rate, frame_count):
    # 解析起始时间码
    hours, minutes, seconds, frames = map(int, start_timecode.split(':'))

    # 计算总帧数
    total_frames = (hours * 3600 + minutes * 60 + seconds) * frame_rate + frames + frame_count

    # 计算新的时间码
    new_hours, remaining_frames = divmod(total_frames, frame_rate * 3600)
    new_minutes, remaining_frames = divmod(remaining_frames, frame_rate * 60)
    new_seconds, new_frames = divmod(remaining_frames, frame_rate)

    # 格式化为时间码字符串
    new_timecode = f"{new_hours:02d}:{new_minutes:02d}:{new_seconds:02d}:{new_frames:02d}"

    return new_timecode


def frames_to_time(frames, frame_rate):
    total_seconds = frames // frame_rate
    remaining_frames = frames % frame_rate
    return f"{total_seconds}秒 {remaining_frames}帧"


def find_files_with_prefix_and_extension(directory, prefix, extension):
    matching_files = [os.path.join(directory, file) for file in os.listdir(directory)
                      if file.startswith(prefix) and file.endswith(extension)]
    return matching_files


class Timeline:
    name = ''
    resolutionWidth = 0
    resolutionHeight = 0
    frameRate = 0
    timeline = None

def timelineSort(elem: Timeline):
    return elem.name


# Press the green button in the gutter to run the script.
if __name__ == '__main__':

    resolve = dvr.scriptapp("Resolve")
    projectManager = resolve.GetProjectManager()
    projectFromResolve = projectManager.GetCurrentProject()
    print(projectFromResolve.GetName())

    arguments = sys.argv

    outputDir = ''
    if len(arguments) >= 2 and not arguments[1] == '':
        if os.path.exists(arguments[1]):
            outputDir = arguments[1]
    while outputDir == '':
        inputStr = input(f'{bcolors.HEADER}Please specify the output directory: {bcolors.ENDC}').strip()
        if not inputStr == '':
            if os.path.exists(inputStr):
                outputDir = inputStr
            else:
                print(f'{bcolors.FAIL}This directory is invalid.{bcolors.ENDC}')

    projectName = projectFromResolve.GetName()

    formatted_date = datetime.now().strftime("%Y%m%d-%H%M%S")
    outputDir = os.path.join(outputDir, f'{projectName}_{formatted_date}')
    if os.path.exists(outputDir):
        shutil.rmtree(outputDir)
    os.mkdir(outputDir)

    csvHeader = ['序号', '镜头示意图', '镜头编号', '绘制时长', '起始时间码', '结束时间码', '秒数', '帧数', '程序识别码']
    csvContent = {}

    stillList = {}
    timelineCount = projectFromResolve.GetTimelineCount()
    for timelineIndex in range(1, timelineCount + 1):
        resolveTimeline = projectFromResolve.GetTimelineByIndex(timelineIndex)
        projectFromResolve.SetCurrentTimeline(resolveTimeline)
        pyTimeline = Timeline()
        pyTimeline.timeline = resolveTimeline
        pyTimeline.name = resolveTimeline.GetName()
        pyTimeline.resolutionWidth = int(resolveTimeline.GetSetting('timelineResolutionWidth'))
        pyTimeline.resolutionHeight = int(resolveTimeline.GetSetting('timelineResolutionHeight'))
        pyTimeline.frameRate = int(resolveTimeline.GetSetting('timelineFrameRate'))

        startTimecode = resolveTimeline.GetStartTimecode()
        resolveTimeline.SetCurrentTimecode(startTimecode)

        markers = resolveTimeline.GetMarkers()
        for idx, marker in enumerate(markers):
            duration = markers[marker]['duration']
            print(str(idx + 1) + ' of ' + str(len(markers)), end=' ')
            gotoTimecode = calculate_timecode(startTimecode, pyTimeline.frameRate, marker)
            endTimecode = calculate_timecode(gotoTimecode, pyTimeline.frameRate, duration - 1)
            print('goto ' + gotoTimecode + ' grab still')
            resolveTimeline.SetCurrentTimecode(gotoTimecode)
            still = resolveTimeline.GrabStill()
            shotIndex = 'TL' + str(timelineIndex) + '_Shot' + str(idx + 1) + '_Dur' + str(duration)
            stillList[shotIndex] = still

            # Collect CSV information
            timelineItem = resolveTimeline.GetCurrentVideoItem()
            itemName = timelineItem.GetName()

            if pyTimeline in csvContent.keys():
                csvContent[pyTimeline].append((
                    idx + 1,
                    '参考图',
                    itemName,
                    gotoTimecode + ' - \n' + endTimecode + '\n' + frames_to_time(duration, pyTimeline.frameRate),
                    gotoTimecode,
                    endTimecode,
                    frames_to_time(duration, pyTimeline.frameRate),
                    duration,
                    shotIndex
                ))
            else:
                csvContent[pyTimeline] = [(
                    idx + 1,
                    '参考图',
                    itemName,
                    gotoTimecode + ' - \n' + endTimecode + '\n' + frames_to_time(duration, pyTimeline.frameRate),
                    gotoTimecode,
                    endTimecode,
                    frames_to_time(duration, pyTimeline.frameRate),
                    duration,
                    shotIndex
                )]

    gallery = projectFromResolve.GetGallery()
    album = gallery.GetCurrentStillAlbum()
    for key in stillList.keys():
        album.ExportStills([stillList[key]], outputDir, key, 'jpg')
        stillList[key] = find_files_with_prefix_and_extension(outputDir, key, 'jpg')[0]

    csvDirectory = os.path.join(outputDir, 'Summary_' + projectName + '.csv')
    xlsDirectory = os.path.join(outputDir, 'Summary_' + projectName + '.xlsx')

    workbook = xls.Workbook(xlsDirectory)
    worksheet = workbook.add_worksheet()
    text_format = workbook.add_format({
        'text_wrap': True,
        'align': 'center',
        'valign': 'vcenter',
        'font_name': '等线',
        'font_size': 11
    })
    text_format_sub = workbook.add_format({
        'font_size': 9
    })
    centered_format = workbook.add_format({
        'text_wrap': True,
        'align': 'center',
        'valign': 'vcenter'
    })
    worksheet.set_row(0, 40)
    worksheet.set_column('A:A', 4.83)
    worksheet.set_column('B:B', 19.17)
    worksheet.set_column('C:C', 28.67)
    worksheet.set_column('D:D', 13.17)
    worksheet.set_column('E:E', 11.67)
    worksheet.set_column('F:F', 11.67)
    worksheet.set_column('G:G', 9.17)
    worksheet.set_column('H:H', 7.67)
    worksheet.set_column('I:I', 13.33)

    for idx, content in enumerate(csvHeader):
        if idx == 2:
            worksheet.write_rich_string(0, idx, text_format, content, text_format_sub, '\n文件名', centered_format)
        else:
            worksheet.write(0, idx, content, text_format)

    startingLine = 1
    for idx, key in enumerate(csvContent.keys()):
        worksheet.set_row(startingLine, 20)
        worksheet.merge_range(startingLine, 1, startingLine, 7, key.name, text_format)
        startingLine += 1
        contents = csvContent[key]
        scale_x = 120 / key.resolutionWidth
        scale_y = 60 / key.resolutionHeight
        scale = min(scale_x, scale_y)
        for line, shot in enumerate(contents):
            worksheet.set_row(line + startingLine, 60)
            imageURL = stillList[shot[8]]
            for column, content in enumerate(shot):
                if column == 0:
                    worksheet.write(line + startingLine, column, f'{idx + 1}-{content}', text_format)
                elif column == 1:
                    worksheet.insert_image(line + startingLine, 1, imageURL, {'x_scale': scale, 'y_scale': scale})
                # elif column == 0 or column == 7:
                #     worksheet.write_number(line + 1, column, content, text_format)
                else:
                    worksheet.write(line + startingLine, column, content, text_format)
        startingLine += len(contents)

    worksheet.set_column(8, 8, width=0)

    workbook.close()
