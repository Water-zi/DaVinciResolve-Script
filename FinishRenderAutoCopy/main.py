# This is a sample Python script.

# Press ⌃R to execute it or replace it with your code.
# Press Double ⇧ to search everywhere for classes, files, tool windows, actions, and settings.
import math
import time
import os
import DaVinciResolveScript as dvr
from distutils.dir_util import copy_tree
from multiprocessing import Process
from kbhit import KBHit
import kbhit


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


def copy_folder(source, destination):
    copy_tree(source, destination)


def get_directory_size(directory):
    total_size = 0
    for dirpath, dirnames, filenames in os.walk(directory):
        for filename in filenames:
            file_path = os.path.join(dirpath, filename)
            total_size += os.path.getsize(file_path)
    return total_size


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    resolve = dvr.scriptapp("Resolve")
    projectManager = resolve.GetProjectManager()
    projectFromResolve = projectManager.GetCurrentProject()
    print(projectFromResolve.GetName())

    desPath = []
    while 1:
        path = input(f'{bcolors.WARNING}Enter a destination path (Leave blank if no more destination): {bcolors.ENDC}').strip()
        if path != '':
            if path in desPath:
                print(f'{bcolors.FAIL}ERROR: Path already exist.{bcolors.ENDC}')
            elif os.path.isdir(path):
                desPath.append(path)
                print(f'{bcolors.OKCYAN}Added to destination list. Total: ' + str(len(desPath)) + f'.{bcolors.ENDC}')
            else:
                print(f'{bcolors.FAIL}ERROR: Input is not a valid path.{bcolors.ENDC}')
        else:
            break

    if len(desPath) == 0:
        print(f'{bcolors.FAIL}ERROR: No destination path to run this script is meaningless.{bcolors.ENDC}')

    renderJobs = projectFromResolve.GetRenderJobList()

    waitForJobStartToRender = 10
    while not projectFromResolve.IsRenderingInProgress():
        print(f'Not rendering, waiting... Timeout in {bcolors.BOLD}' + str(waitForJobStartToRender) + f'{bcolors.ENDC} seconds.   ', end='\r', flush=True)
        waitForJobStartToRender -= 1
        time.sleep(1)
        if waitForJobStartToRender <= 0:
            print(f'\n{bcolors.FAIL}ERROR: Timeout. Render Jobs not started. Program EXIT.{bcolors.ENDC}')
            exit(1)

    print(f'{bcolors.HEADER}Captured render started, checking job list...                  {bcolors.ENDC}')

    # Update job list before loop
    renderJobs = projectFromResolve.GetRenderJobList()
    notCancelledJobIds = []
    # Check if job list is clean
    copySourceFolder = '?'
    for job in renderJobs:
        targetDir = job['TargetDir']
        jobId = job['JobId']
        status = projectFromResolve.GetRenderJobStatus(jobId)['JobStatus']
        if status != 'Cancelled':
            notCancelledJobIds.append(jobId)
        if copySourceFolder == '?':
            copySourceFolder = targetDir
        elif copySourceFolder == targetDir:
            continue
        else:
            print(f'{bcolors.FAIL}ERROR: Job list is dirty. Program EXIT.{bcolors.ENDC}')
            exit(1)

    if len(renderJobs) > len(notCancelledJobIds):
        print(f'{bcolors.WARNING}WARNING: It is recommended to Clear Render Status before rendering.{bcolors.ENDC}')
    else:
        print(f'{bcolors.BOLD}==========List OK==============={bcolors.ENDC}')

    waitTimeCounter = 0
    while projectFromResolve.IsRenderingInProgress():
        print(f'Waiting for all jobs to complete... ({bcolors.BOLD}' + str(waitTimeCounter) + f'{bcolors.ENDC})', end='\r', flush=True)
        waitTimeCounter += 1
        time.sleep(1)

    print(f'\n{bcolors.BOLD}==========Render Ended=========={bcolors.ENDC}')

    # Check if it is a Manual Stop
    isManualStop = False
    for jobId in notCancelledJobIds:
        status = projectFromResolve.GetRenderJobStatus(jobId)['JobStatus']
        if status == 'Cancelled':
            isManualStop = True
            break
    if isManualStop:
        print(f'{bcolors.FAIL}NOTE: This is a Manual Stop. Program EXIT.{bcolors.ENDC}')
        exit(1)
    else:
        copyCountDown = 10
        kb = KBHit()
        print(f'{bcolors.WARNING}Press ctrl+c to Cancel. Press any key to skip.{bcolors.ENDC}')
        while copyCountDown > 0:
            print(f'{bcolors.WARNING}Finish rendering. Will start copy in ' + str(copyCountDown) + f' seconds...    {bcolors.ENDC}', end='\r', flush=True)
            copyCountDown -= 1
            countDown = 100
            while countDown > 0:
                if kb.kbhit():
                    copyCountDown = 0
                    break
                time.sleep(0.01)
                countDown -= 1

    print(f'\n{bcolors.BOLD}==========Copy Started=========={bcolors.ENDC}')

    sourceFolderSize = get_directory_size(copySourceFolder)
    for index, des in enumerate(desPath):
        print(f'{bcolors.HEADER}' + str(index + 1) + ' of ' + str(len(desPath)) + ' \t Copying to ' + des + f'{bcolors.ENDC}')
        proc = Process(target=copy_folder, args=(copySourceFolder, des), daemon=True)
        proc.start()
        while proc.is_alive():
            time.sleep(0.02)
            # 计算进度浮点数
            desSize = get_directory_size(des)
            progress = desSize / sourceFolderSize
            # 计算进度条的方块数量
            bar_length = 50  # 进度条长度
            num_filled = int(progress * bar_length)
            num_empty = bar_length - num_filled
            filled_block = '■'
            empty_block = '□'
            progress_bar = filled_block * num_filled + empty_block * num_empty
            print(f"{progress_bar}", end='\t')
            print(str(int(progress * 100)) + '%', end='\r', flush=True)
        print('')

    print(f'🍺 {bcolors.OKGREEN}SUCCEED: Finish copy to all destination.{bcolors.ENDC}')
