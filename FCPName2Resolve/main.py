import DaVinciResolveScript as dvr
import xml.etree.ElementTree as ET


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
    print(folder.GetName())

    clips = folder.GetClipList()

    # 10.xml is an example file, replace it to yours
    tree = ET.parse("10.xml")
    root = tree.getroot()
    items = root.iter("clipitem")
    clipTable = {}
    for item in items:
        binFileName = item.get('id')
        files = item.iter("file")
        for file in files:
            realFileName = file.get('id')
            if realFileName in clipTable.keys():
                continue
            clipTable[realFileName] = binFileName

    for clip in clips:
        clipName = clip.GetClipProperty('Clip Name')
        # Ignore file extension in this situation
        clipName = clipName.split('.')[0]
        print(f'{bcolors.HEADER}Finding ' + clipName + f' in ClipTable...{bcolors.ENDC}', end='\t')
        try:
            newClipName = clipTable[clipName]
            clip.SetClipProperty('Clip Name', newClipName)
            print(f'{bcolors.OKCYAN}Modified to ' + newClipName + f'{bcolors.ENDC}')
        except:
            print(f'{bcolors.FAIL}Not Found.{bcolors.ENDC}')

    exit(0)


# See PyCharm help at https://www.jetbrains.com/help/pycharm/
