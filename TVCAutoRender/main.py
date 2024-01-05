# -*- coding: utf-8 -*-
# 用于支持中文目录输入和输出
import os
import uuid

import DaVinciResolveScript as dvr


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
    print(f'{bcolors.OKGREEN}{bcolors.BOLD}' + projectFromResolve.GetName() + f'{bcolors.ENDC}')
    if projectFromResolve.IsRenderingInProgress():
        print(f'{bcolors.FAIL}Rendering is in progress. Program will exit.{bcolors.ENDC}')
        exit(0)
    mediaPool = projectFromResolve.GetMediaPool()

    # Check project settings
    # settings = projectFromResolve.GetSetting()
    # print(settings)
    # exit(0)

    # Read config file
    useOriginalResolution = False
    useOriginalColor = False
    allowMultipleFPS = True
    verboseConsole = True
    relativeRenderDestination = ''
    overrideExistingLUT = False
    try:
        configFile = open('config.conf', 'r', encoding='utf-8')
        for line in configFile:
            if line.startswith('#'):
                continue
            line = line.removesuffix('\n')
            key = line.split(':')[0]
            value = line.split(':')[-1]
            match key:
                case 'UseOriginalResolution':
                    useOriginalResolution = (value == '1')
                    if useOriginalResolution and 'TVC_Proxy_OR' not in projectFromResolve.GetRenderPresetList():
                        print(f'{bcolors.FAIL}You do not have a render preset called \'{bcolors.BOLD}TVC_Proxy_OR{bcolors.ENDC}{bcolors.FAIL}\'. You should create this preset first.{bcolors.ENDC}')
                        print(f'{bcolors.WARNING}Remember to have \'Render at source resolution\' checked.')
                        exit(0)
                    elif 'TVC_Proxy' not in projectFromResolve.GetRenderPresetList():
                        print(f'{bcolors.FAIL}You do not have a render preset called \'{bcolors.BOLD}TVC_Proxy{bcolors.ENDC}{bcolors.FAIL}\'. You should create this preset first.{bcolors.ENDC}')
                        print(f'{bcolors.WARNING}Remember to have \'Render at source resolution\' unchecked. And specify the correct format, audio tracks, etc.')
                        exit(0)
                case 'UseOriginalColor':
                    useOriginalColor = (value == '1')
                case 'AllowMultipleFPS':
                    allowMultipleFPS = (value == '1')
                case 'VerboseConsole':
                    verboseConsole = (value == '1')
                case 'RelativeRenderDestination':
                    relativeRenderDestination = value
                case 'OverrideExistingLUT':
                    overrideExistingLUT = (value == '1')
                case _:
                    continue
    except FileNotFoundError:
        print(f'{bcolors.FAIL}Failed to read config. Make sure config.conf is reachable.{bcolors.ENDC}')
        exit(0)

    # Import clips from terminal
    inputPaths = []
    requireCurrentFolderInResolve = False
    while 1:
        inputStr = input(f'{bcolors.HEADER}Drag a REEL (or REELs) to terminal to import. Leave blank to continue: {bcolors.ENDC}')
        if not inputStr == '':
            paths = inputStr.split('/Volumes')
            for path in paths:
                if not path.startswith('/'):
                    continue
                inputStr = '/Volumes' + path.strip()
                if inputStr in inputPaths:
                    print(f'{bcolors.FAIL}ERROR: folder already exist.{bcolors.ENDC}')
                elif os.path.isdir(inputStr):
                    inputPaths.append(inputStr)
                    reelName = os.path.basename(os.path.normpath(inputStr))
                    print(f'{bcolors.OKCYAN}Added ' + reelName + ' to list. Total: ' + str(len(inputPaths)) + f'.{bcolors.ENDC}')
                else:
                    print(f'{bcolors.FAIL}ERROR: Folder is not a valid path.{bcolors.ENDC}')
        else:
            if len(inputPaths) > 0:
                break
            else:
                currentFolder = mediaPool.GetCurrentFolder().GetName()
                inputStr = input(f'{bcolors.WARNING}Are you sure to continue with {bcolors.UNDERLINE}' + currentFolder + f'{bcolors.ENDC}{bcolors.WARNING}? (\'y\' or empty for yes): {bcolors.ENDC}')
                if inputStr == '' or inputStr == 'y':
                    break

    folderList = []
    folderRenderDestinationDictList = {}
    for path in inputPaths:
        reelName = os.path.basename(os.path.normpath(path))
        print(f'{bcolors.OKCYAN}Importing from ' + reelName + f'......{bcolors.ENDC}', end='')
        root = mediaPool.GetRootFolder()
        folder = mediaPool.AddSubFolder(root, reelName)
        folderList.append(folder)
        if not relativeRenderDestination == '':
            destination = os.path.normpath(os.path.join(path, relativeRenderDestination))
            folderRenderDestinationDictList[folder.GetName()] = destination
        mediaPool.SetCurrentFolder(folder)
        mediaList = []
        for root, dirs, _ in os.walk(path, topdown=False):
            for name in dirs:
                mediaList.append(os.path.join(root, name))
        clips = mediaPool.ImportMedia(mediaList)
        print(f'{bcolors.OKCYAN}{bcolors.UNDERLINE}' + str(len(clips)) + ' clip' + ('s' if len(clips) > 0 else '') + f'{bcolors.ENDC}{bcolors.OKCYAN} have imported.{bcolors.ENDC}')

    if len(folderList) == 0:
        folderList.append(mediaPool.GetCurrentFolder())

    # Select a render destination, leave blank may lead to unexpected error
    renderDestination = relativeRenderDestination
    if renderDestination == '':
        try:
            renderDesFile = open('renderDestination.default', 'r', encoding='utf-8')
            renderDestination = renderDesFile.read()
            if not os.access(renderDestination, os.W_OK):
                print(f'{bcolors.WARNING}You do not have write permission to the default render destination.{bcolors.ENDC}')
                renderDestination = ''
        except FileNotFoundError:
            renderDestination = ''

    requireEmptyDestination = False
    if renderDestination == '':
        while 1:
            inputStr = input(f'{bcolors.HEADER}Please specify a render destination: {bcolors.ENDC}').strip()
            if not inputStr == '':
                if os.path.isdir(inputStr):
                    renderDestination = inputStr
                    renderDesFile = open('renderDestination.default', 'w', encoding='utf-8')
                    renderDesFile.write(renderDestination)
                    renderDesFile.close()
                    print(f'{bcolors.OKCYAN}Set ' + renderDestination + f' as render destination.{bcolors.ENDC}')
                    break
                else:
                    print(f'{bcolors.FAIL}ERROR: Destination is not a valid path.{bcolors.ENDC}')
            else:
                if not requireEmptyDestination:
                    print(f'{bcolors.WARNING}Empty destination will lead to unexpected result. Leave blank again to confirm.{bcolors.ENDC}')
                    requireEmptyDestination = True
                else:
                    break
    elif verboseConsole:
        if relativeRenderDestination == '':
            print(f'{bcolors.OKBLUE}Your default render destination is: {bcolors.UNDERLINE}' + renderDestination + f'{bcolors.ENDC}')
            inputStr = input(f'{bcolors.HEADER}Please specify a render destination or leave blank to retain current: {bcolors.ENDC}').strip()
        else:
            renderDestinationList = {}
            for key in folderRenderDestinationDictList.keys():
                if folderRenderDestinationDictList[key] in renderDestinationList.keys():
                    renderDestinationList[folderRenderDestinationDictList[key]].append(key)
                else:
                    renderDestinationList[folderRenderDestinationDictList[key]] = [key]
            print(f'{bcolors.OKBLUE}You are using relative render destination: {bcolors.ENDC}')
            for key in renderDestinationList.keys():
                for reel in renderDestinationList[key]:
                    print(f'{bcolors.OKCYAN}' + reel + f' {bcolors.ENDC}', end='')
                print(f'{bcolors.OKCYAN}will be render to:{bcolors.UNDERLINE}')
                print(key + f'{bcolors.ENDC}')
    else:
        print(f'{bcolors.OKCYAN}Non Verbose Mode - Using Render Destination Preset.{bcolors.ENDC}')

    # Auto apply LUT part
    # Readout LUT list
    # if you are using a Windows Computer, you may need to change this path below
    lutFolder = '/Library/Application Support/Blackmagic Design/DaVinci Resolve/LUT'
    lutList = []
    lutListPath = {}
    lutListBackup = {}
    for root, dirs, files in os.walk(lutFolder, topdown=False):
        for name in files:
            actualName = name.rsplit('.', 1)[0]
            if actualName.startswith('.'):
                continue
            lutListPath[actualName] = os.path.join(root, actualName).split('/LUT/')[-1]
            lutList.append(actualName)
    lutListBackup = lutList

    applyLUT = ''
    manufacturer = ''

    lutPresetList = {}

    # Read primary LUT preset
    # Skip if user use original color
    if not useOriginalColor:
        try:
            lutFile = open('lut.default', 'r', encoding='utf-8')
            fileContent = lutFile.read()
            manufacturer = fileContent.split(':')[0]
            applyLUT = fileContent.split(':')[-1].split('\n')[0]
            if not applyLUT == '':
                print('Primary Lut Preset:')
                if manufacturer == '':
                    print(f'{bcolors.OKBLUE}' + applyLUT + f' -> Any Clip.{bcolors.ENDC}')
                else:
                    print(f'{bcolors.OKBLUE}' + applyLUT + ' -> ' + manufacturer + f'.{bcolors.ENDC}')
                if verboseConsole:
                    inputStr = input(f'{bcolors.HEADER}Would you like to use this preset? (\'y\' or empty for yes): {bcolors.ENDC}')
                    if not inputStr == '' and not inputStr == 'y':
                        applyLUT = ''
                        manufacturer = ''
        except FileNotFoundError:
            print('No Primary Lut Preset')

        # Read secondary LUT preset list
        try:
            lutListFile = open('luts.default', 'r', encoding='utf-8')
            for line in lutListFile:
                if ':' not in line:
                    continue
                listManufacturer = line.split(':')[0]
                listApplyLut = line.split(':')[-1].split('\n')[0]
                if not listApplyLut == '' and not listManufacturer == '':
                    lutPresetList[listManufacturer] = listApplyLut
            if len(lutPresetList) > 0:
                print('Secondary LUTs Preset:')
                for key in lutPresetList.keys():
                    print(f'{bcolors.OKBLUE}' + key + ' -> ' + lutPresetList[key] + f'.{bcolors.ENDC}')
        except FileNotFoundError:
            print('For Advanced User: No Secondary Lut Preset, ignore if you don\'t understand.')

        while applyLUT == '':
            inputStr = input(f'{bcolors.HEADER}Search for LUT (\'ls\' to list, empty to skip): {bcolors.ENDC}')
            if inputStr == '':
                break
            elif inputStr.lower() == 'ls':
                inputStr = ''
            newLutList = []
            if len(lutList) < 10 and inputStr.isnumeric() and int(inputStr) < 10 and int(inputStr) <= len(lutList) - 1:
                newLutList.append(lutList[int(inputStr)])
            else:
                for lut in lutList:
                    if not lut.lower().find(inputStr.lower()) == -1:
                        newLutList.append(lut)
                if inputStr.isnumeric():
                    if int(inputStr) <= len(lutList) - 1:
                        newLutList.append(lutList[int(inputStr)])
            lutList = newLutList
            if len(lutList) == 0:
                inputStr = input(f'{bcolors.WARNING}Did not find any LUT, press \'y\' if you want no lut to be applied.{bcolors.ENDC}')
                if inputStr == 'y':
                    break
                else:
                    lutList = lutListBackup
            else:
                if len(lutList) == 1:
                    inputStr = input(f'{bcolors.HEADER}Would you like to use {bcolors.OKCYAN}{bcolors.BOLD}' + lutListPath[lutList[0]] + f'{bcolors.ENDC}{bcolors.HEADER} (\'y\' or empty for yes): {bcolors.ENDC}')
                    if inputStr == 'y' or inputStr == '':
                        inputStr = input(f'{bcolors.HEADER}Specify a manufacturer if you want to: {bcolors.ENDC}')
                        manufacturer = inputStr
                        applyLUT = lutListPath[lutList[0]]
                        presetStr = manufacturer + ':' + applyLUT
                        presetFile = open('lut.default', 'w', encoding='utf-8')
                        presetFile.write(presetStr)
                        presetFile.close()
                        break
                    else:
                        lutList = lutListBackup
                else:
                    for idx, lut in enumerate(lutList):
                        print(f'{bcolors.BOLD}', end='')
                        print(idx, end=f' - {bcolors.ENDC}')
                        print(lut)

    fpsList = set()
    for folder in folderList:
        # It is recommended to put a single reel file in one folder to run this script
        mediaPool.SetCurrentFolder(folder)
        timelineList = {}
        clips = folder.GetClipList()
        for clip in clips:
            print('Analyzing clip in ' + folder.GetName() + ': ' + clip.GetName() + '                       ', end='\r')
            # Remove timeline, images, and other NOT Video clips
            if 'Video' not in clip.GetClipProperty('Type') and '视频' not in clip.GetClipProperty('Type'):
                mediaPool.DeleteClips([clip])
                continue
            # Abort when reel name is empty, require user to obtain a ReelName
            if clip.GetClipProperty('Reel Name') == '':
                print(f'{bcolors.FAIL}Reel Name is empty, this is not allowed, please specify a reel name.')
                exit(0)
            # Apply LUT if user needs to
            if not useOriginalColor and (not applyLUT == '' or len(lutPresetList) > 0):
                if overrideExistingLUT or clip.GetClipProperty('Input LUT') == '':
                    if applyLUT == '' or (len(lutPresetList) > 0 and not manufacturer == '' and not clip.GetMetadata('Camera Manufacturer') == manufacturer):
                        actucalManufacturer = clip.GetMetadata('Camera Manufacturer')
                        if actucalManufacturer in lutPresetList.keys():
                            actualLUT = lutPresetList[actucalManufacturer]
                        else:
                            actualLUT = ''
                        if not actualLUT == '':
                            clip.SetClipProperty('Input LUT', actualLUT)
                            if not clip.GetClipProperty('Input LUT') == actualLUT:
                                print(f'{bcolors.FAIL}Failed to apply LUT to {bcolors.UNDERLINE}' + clip.GetName() + f'{bcolors.ENDC}{bcolors.FAIL}. Probably because LUT does not exist.')
                        else:
                            print(f'{bcolors.WARNING}Skip applying LUT to {bcolors.UNDERLINE}' + clip.GetName() + f'{bcolors.ENDC}{bcolors.WARNING}. Manufacturer Not Match.{bcolors.ENDC}')
                    elif not applyLUT == '':
                        clip.SetClipProperty('Input LUT', applyLUT)
                        if not clip.GetClipProperty('Input LUT') == applyLUT:
                            print(f'{bcolors.FAIL}Failed to apply LUT to {bcolors.UNDERLINE}' + clip.GetName() + f'{bcolors.ENDC}{bcolors.FAIL}. Probably because LUT does not exist.')
                else:
                    print(f'{bcolors.WARNING}Skip applying LUT to {bcolors.UNDERLINE}' + clip.GetName() + f'{bcolors.ENDC}{bcolors.WARNING}. Please remove the LUT on it first.{bcolors.ENDC}')
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

        # Check if there are multiple frame rate
        if not allowMultipleFPS and len(fpsList) > 1:
            print(f'{bcolors.FAIL}Clips has multiple FPS. Use a traditional way and modify them first.{bcolors.ENDC}')
            exit(1)

        for key in timelineList.keys():
            timelineCount = projectFromResolve.GetTimelineCount()
            timelineNameList = []
            for idx in range(timelineCount):
                tempTimeline = projectFromResolve.GetTimelineByIndex(idx + 1)
                print(tempTimeline.GetName())
                timelineNameList.append(tempTimeline.GetName())
            actualTimelineName = folder.GetName() + '_' + "{:.2f}".format(key)
            if actualTimelineName in timelineNameList:
                actualTimelineName = actualTimelineName + '_' + str(uuid.uuid4())
            timeline = mediaPool.CreateTimelineFromClips(actualTimelineName, timelineList[key])
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
            projectFromResolve.SetCurrentTimeline(timeline)
            # You should have a RENDER PRESET named 'TVC_Proxy' or 'TVC_Proxy_OR', that setup for the right settings base on your own needs,
            # e.g. format, audio tracks, and fileName regulation
            if useOriginalResolution:
                if 'TVC_Proxy_OR' in projectFromResolve.GetRenderPresetList():
                    projectFromResolve.LoadRenderPreset('TVC_Proxy_OR')
            else:
                if 'TVC_Proxy' in projectFromResolve.GetRenderPresetList():
                    projectFromResolve.LoadRenderPreset('TVC_Proxy')
                # Set timeline output resolution base on itself
                projectFromResolve.SetRenderSettings({
                    'FormatWidth': width,
                    'FormatHeight': height
                })
            # Set render destination if its available
            if not relativeRenderDestination == '' and folder.GetName() in folderRenderDestinationDictList.keys():
                destination = folderRenderDestinationDictList[folder.GetName()]
                projectFromResolve.SetRenderSettings({
                    'TargetDir': str(os.path.join(destination, folder.GetName()))
                })
            elif not renderDestination == '':
                projectFromResolve.SetRenderSettings({
                    'TargetDir': str(os.path.join(renderDestination, folder.GetName()))
                })
            projectFromResolve.AddRenderJob()

    print(f'🍺 {bcolors.OKGREEN}Operation Complete.{bcolors.ENDC}')
