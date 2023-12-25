# This is a sample Python script.

# Press ⌃R to execute it or replace it with your code.
# Press Double ⇧ to search everywhere for classes, files, tool windows, actions, and settings.

import DaVinciResolveScript as dvr

def print_hi(name):
    # Use a breakpoint in the code line below to debug your script.
    print(f'Hi, {name}')  # Press ⌘F8 to toggle the breakpoint.


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    print_hi('PyCharm')

    resolve = dvr.scriptapp("Resolve")
    projectManager = resolve.GetProjectManager()
    projectFromResolve = projectManager.GetCurrentProject()
    print(projectFromResolve.GetName())

    mediaPool = projectFromResolve.GetMediaPool()
    folder = mediaPool.GetCurrentFolder()
    print(folder.GetName())

    # tl = projectFromResolve.GetCurrentTimeline()
    # print(tl.GetSetting())

    clips = folder.GetClipList()
    anamorphicClips = []
    normalClips = []
    otherClips = []
    # Separate clips via Aspect Ratio Notes in metadata
    for clip in clips:
        pixelAspectRatio = clip.GetMetadata("Aspect Ratio Notes")
        if pixelAspectRatio == '2000/1000':
            anamorphicClips.append(clip)
        elif pixelAspectRatio == '1/1':
            normalClips.append(clip)
        else:
            otherClips.append(clip)
    # Create timelines with these clips separately
    anamorphicTimeline = mediaPool.CreateTimelineFromClips(folder.GetName() + "_ANA", anamorphicClips)
    normalTimeline = mediaPool.CreateTimelineFromClips(folder.GetName() + "_NORMAL", normalClips)
    otherTimeline = mediaPool.CreateTimelineFromClips(folder.GetName() + "_UNKNOWN", otherClips)

    # Scale to crop anamorphicClips to fit the release aspect ratio
    anamorphicClips = anamorphicTimeline.GetItemListInTrack('video', 1)
    for clip in anamorphicClips:
        clip.SetProperty('ZoomX', 1.356)
        clip.SetProperty('ZoomY', 1.356)

    # Set scale to crop to the square pixel clips timeline, to crop the blank edge in horizontal direction
    normalTimeline.SetSetting('useCustomSettings', '1')
    normalTimeline.SetSetting('timelineInputResMismatchBehavior', 'scaleToCrop')

    # Load render preset ( set in Resolve before, mainly about burn-in preset )
    projectFromResolve.SetCurrentTimeline(anamorphicTimeline)
    projectFromResolve.LoadRenderPreset("LSDDWY_ANA")
    projectFromResolve.AddRenderJob()

    projectFromResolve.SetCurrentTimeline(normalTimeline)
    projectFromResolve.LoadRenderPreset("LSDDWY_Normal")
    projectFromResolve.AddRenderJob()

    # Add grading nodes here if needed


# See PyCharm help at https://www.jetbrains.com/help/pycharm/
