import os.path

import DaVinciResolveScript as dvr


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

    gallery = projectFromResolve.GetGallery()
    album = gallery.GetCurrentStillAlbum()
    stills = album.GetStills()

    des = input(f'{bcolors.HEADER}Enter a output destination here: {bcolors.ENDC}')

    for idx, still in enumerate(stills):
        label = album.GetLabel(still).split('.')[0]
        album.ExportStills([still], des, label, 'jpg')

    exit(0)

    mediaPool = projectFromResolve.GetMediaPool()
    folder = mediaPool.GetCurrentFolder()
    print(folder.GetName())

    clips = folder.GetClipList()

    succeedClipCount = 0
    skippedItemCount = 0
    failedClips = []

    # Get a correspondence table between Clip and LUT
    for clip in clips:
        try:
            # For SONY Venice 2, the lut name in metadata should split as below
            lutName = clip.GetMetadata('Mon Color Space').split(':')[1].split('.cube')[0]
        except:
            # If item is a timeline or something else, it should not be modified
            print(f'{bcolors.WARNING}Skipped item: ' + clip.GetName() + f'{bcolors.ENDC}')
            skippedItemCount += 1
            continue
        originLUTName = clip.GetClipProperty('Input LUT')
        if originLUTName == '':
            originLUTName = 'No LUT'
        print(clip.GetName() + f'{bcolors.HEADER} with {bcolors.ENDC}' + originLUTName + f'{bcolors.HEADER} replacing to {bcolors.ENDC}' + lutName + ' ...... ', end='')
        # If you want to apply LUT automatically, you should place it in the LUT folder (or some sub-folder) of DaVinci Resove in advance
        # "QT" is the sub-folder in this example
        clip.SetClipProperty('Input LUT', 'QT/' + lutName)
        # Check if LUT is successfully applied, if not, probably because this LUT is not in the QT/ folder.
        if clip.GetClipProperty('Input LUT') == 'QT/' + lutName:
            print(f'{bcolors.OKGREEN}Succeed!{bcolors.ENDC}')
            succeedClipCount += 1
        else:
            print(f'{bcolors.FAIL}Failed!{bcolors.ENDC}')
            failedClips.append(clip.GetName())

    # Print out results
    if len(failedClips) == 0:
        print(f'{bcolors.OKCYAN}All (' + str(succeedClipCount) + f') clips have successfully applied LUT!{bcolors.ENDC}')
        if skippedItemCount > 0:
            print(f'{bcolors.WARNING}With ' + str(skippedItemCount) + f' item(s) skipped.{bcolors.ENDC}')
    else:
        print(f'{bcolors.FAIL}Clips (' + str(len(failedClips)) + ' / ' + str(len(failedClips) + succeedClipCount) + ') below failed to apply LUT:\n' + '\n'.join(failedClips) + f'{bcolors.ENDC}')
        if skippedItemCount > 0:
            print(f'{bcolors.WARNING}With ' + str(skippedItemCount) + f' item(s) skipped.{bcolors.ENDC}')
