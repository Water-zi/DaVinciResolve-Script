# This is a sample Python script.

# Press ⌃R to execute it or replace it with your code.
# Press Double ⇧ to search everywhere for classes, files, tool windows, actions, and settings.
import math
import time
import os
import DaVinciResolveScript as dvr
from distutils.dir_util import copy_tree
from multiprocessing import Process


# Colorful terminal output
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


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    resolve = dvr.scriptapp("Resolve")
    projectManager = resolve.GetProjectManager()
    projectFromResolve = projectManager.GetCurrentProject()
    print(projectFromResolve.GetName())
    mediaPool = projectFromResolve.GetMediaPool()
    folder = mediaPool.GetCurrentFolder()
    clips = folder.GetClipList()

    # It is recommended to put a single reel file in one folder to run this script
    timelineList = {}
    fpsList = set()

    for clip in clips:
        print('Analyzing clip ' + clip.GetName() + '                        ', end='\r')
        if clip.GetClipProperty('Type') == 'Timeline':
            mediaPool.DeleteClips([clip])
            continue
        clipResolution = clip.GetClipProperty('Resolution')
        width = int(clipResolution.split('x')[0])
        height = int(clipResolution.split('x')[1])
        # in case of error
        if width == 0 or height == 0:
            continue
        ratio = width / height
        if ratio in timelineList.keys():
            timelineList[ratio].append(clip)
        else:
            timelineList[ratio] = [clip]

        # Check if a clip's properties is abnormal

        pixelAspectRatioNormal = clip.GetClipProperty('PAR') == 'Square'
        if not pixelAspectRatioNormal:
            print(f'{bcolors.WARNING}' + clip.GetName() + f' PAR is not Square.{bcolors.ENDC}')
        fps = clip.GetClipProperty('FPS')
        fpsList.add(fps)

    print('\nEnd of Analyze')

    if len(fpsList) > 1:
        print(f'{bcolors.FAIL}Clips has multiple FPS. Modify them first.{bcolors.ENDC}')
        exit(1)

    for key in timelineList.keys():
        timeline = mediaPool.CreateTimelineFromClips(folder.GetName() + '_' + "{:.2f}".format(key), timelineList[key])
        width = 1920
        height = 1920
        if key < 1:
            width = round(height * key / 2) * 2
        else:
            height = round(width / key / 2) * 2
        timeline.SetSetting('useCustomSettings', '1')
        timeline.SetSetting('timelineResolutionHeight', str(height))
        timeline.SetSetting('timelineOutputResolutionHeight', str(height))
        timeline.SetSetting('timelineResolutionWidth', str(width))
        timeline.SetSetting('timelineOutputResolutionWidth', str(width))
        print(f'{bcolors.OKCYAN}Successfully created ' + timeline.GetName() + ' - ' + str(width) + 'x' + str(height) + f'{bcolors.ENDC}')

    print(f'🍺 {bcolors.OKGREEN}Operation Complete.{bcolors.ENDC}')
