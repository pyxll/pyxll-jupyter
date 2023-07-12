"""
Translated to Python from 
https://gist.githubusercontent.com/guwidoe/038398b6be1b16c458365716a921814d/raw/5bf858b2939e21e2b5731562bfb81768d7bfdb55/GetLocalOneDrivePath.bas.vb

Note: The variable names used in this file don't conform with the style
used in the rest of this project and it is not well structured.
This is intentional to make comparing this code with the original VB code easier.
"""

r"""
'Attribute VB_Name = "GetLocalOneDrivePath"
' Cross-platform VBA Function to get the local path of OneDrive/SharePoint
' synchronized Microsoft Office files (Works on Windows and on macOS)
'
' Author: Guido Witt-Dorring
' Created: 2022/07/01
' Updated: 2023/04/09
' License: MIT
'
' ----------------------------------------------------------------
' https://gist.github.com/guwidoe/038398b6be1b16c458365716a921814d
' https://stackoverflow.com/a/73577057/12287457
' ----------------------------------------------------------------
'
' Copyright (c) 2023 Guido Witt-Dorring
'
' MIT License:
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to
' deal in the Software without restriction, including without limitation the
' rights to use, copy, modify, merge, publish, distribute, sublicense, and/or
' sell copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in
' all copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
' FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS
' IN THE SOFTWARE.
'
'*******************************************************************************
' COMMENTS REGARDING THE IMPLEMENTATION:
' 1) Background and Alternative
'    This function was intended to be written as a single procedure without any
'    dependencies, for maximum portability between projects, as it implements a
'    functionality that is very commonly needed for many VBA applications
'    working inside OneDrive/SharePoint synchronized directories. I followed
'    this paradigm because it was not clear to me how complicated this simple
'    sounding endeavour would turn out to be.
'    Unfortunately, more and more complications arose, and little by little,
'    the procedure turned incredibly complicated. I do not condone the coding
'    style applied here, and that this is not how I usually write code,
'    nevertheless, I'm not open to rewriting this code in a different style,
'    because a clean implementation of this algorithm already exists, as pointed
'    out in the following.
'
'    If you would like to understand the underlying algorithm of how the local
'    path can be found with only the Url-path as input, I recommend following
'    the much cleaner implementation by Cristian Buse:
'    https://github.com/cristianbuse/VBA-FileTools
'    We developed the algorithm together and wrote separate implementations
'    concurrently. His solution is contained inside a module-level library,
'    split into many procedures and using features like private types and API-
'    functions, that are not available when trying to create a single procedure
'    without dependencies like below. This makes his code more readable.
'
'    Both of our solutions are well tested and actively supported with bugfixes
'    and improvements, so both should be equally valid choices for use in your
'    project. The differences in performance/features are marginal and they can
'    often be used interchangeably. If you need more file-system interaction
'    functionality, use Cristians library, and if you only need GetLocalPath,
'    just copy this function to any module in your project and it will work.
'
' 2) How does this function NOT work?
'    Most other solutions for this problem circulating online (A list of most
'    can be found here: https://stackoverflow.com/a/73577057/12287457) are using
'    one of two approaches, :
'     1. they use the environment variables set by OneDrive:
'         - Environ(OneDrive)
'         - Environ(OneDriveCommercial)
'         - Environ(OneDriveConsumer)
'        and replace part of the URL with it. There are many problems with this
'        approach:
'         1. They are not being set by OneDrive on MacOS.
'         2. It is unclear exactly which part of the URL needs to be replaced.
'         3. Environment variables can be changed by the user.
'         4. Only there three exist. If more onedrive accounts are logged in,
'            they just overwrite the previous ones.
'        or,
'     2. they use the mount points OneDrive writes to the registry here:
'         - \HKEY_CURRENT_USER\Software\SyncEngines\Providers\OneDrive\
'        this also has several drawbacks:
'         1. The registry is not available on MacOS.
'         2. It's still unclear exactly what part of the URL should be replaced.
'         3. These registry keys can contain mistakes, like for example, when:
'             - Synchronizing a folder called "Personal" from someone else's
'               personal OneDrive
'             - Synchronizing a folder called "Business1" from someone else's
'               personal OneDrive and then relogging your own first Business
'               OneDrive account
'             - Relogging you personal OneDrive can change the "CID" property
'               from a folderID formatted cid (e.g. 3DEA8A9886F05935!125) to a
'               regular private cid (e.g. 3dea8a9886f05935) for synced folders
'               from other people's OneDrives
'
'    For these reasons, this solution uses a completely different approach to
'    solve this problem.
'
' 3) How does this function work?
'    This function builds the Web to Local translation dictionary by extracting
'    the mount points from the OneDrive settings files.
'    It reads files from...
'    On Windows:
'        - the "...\AppData\Local\Microsoft" directory
'    On Mac:
'        - the "~/Library/Containers/com.microsoft.OneDrive-mac/Data/" & _
'              "Library/Application Support" directory
'        - and/or the "~/Library/Application Support"
'    It reads the following files:
'      - \OneDrive\settings\Personal\ClientPolicy.ini
'      - \OneDrive\settings\Personal\????????????????.dat
'      - \OneDrive\settings\Personal\????????????????.ini
'      - \OneDrive\settings\Personal\global.ini
'      - \OneDrive\settings\Personal\GroupFolders.ini
'      - \OneDrive\settings\Business#\????????-????-????-????-????????????.dat
'      - \OneDrive\settings\Business#\????????-????-????-????-????????????.ini
'      - \OneDrive\settings\Business#\ClientPolicy*.ini
'      - \OneDrive\settings\Business#\global.ini
'      - \Office\CLP\* (just the filename)
'
'    Where:
'     - "*" ... 0 or more characters
'     - "?" ... one character [0-9, a-f]
'     - "#" ... one digit
'     - "\" ... path separator, (= "/" on MacOS)
'     - The "???..." filenames represent CIDs)
'
'    On MacOS, the \Office\CLP\* exists for each Microsoft Office application
'    separately. Depending on whether the application was already used in
'    active syncing with OneDrive it may contain different/incomplete files.
'    In the code, the path of this directory is stored inside the variable
'    "clpPath". On MacOS, the defined clpPath might not exist or not contain
'    all necessary files for some host applications, because Environ("HOME")
'    depends on the host app.
'    This is not a big problem as the function will still work, however in
'    this case, specifying a preferredMountPointOwner will do nothing.
'    To make sure this directory and the necessary files exist, a file must
'    have been actively synchronized with OneDrive by the application whose
'    "HOME" folder is returned by Environ("HOME") while being logged in
'    to that application with the account whose email is given as
'    preferredMountPointOwner, at some point in the past!
'
'    If you are usually working with Excel but are using this function in a
'    different app, you can instead use an alternative (Excels CLP folder) as
'    the clpPath as it will most likely contain all the necessary information
'    The alternative clpPath is commented out in the code, if you prefer to
'    use Excels CLP folder per default, just un-comment the respective line
'    in the code.
'*******************************************************************************

'*******************************************************************************
' COMMENTS REGARDING THE USAGE:
' This function can be used as a User Defined Function (UDF) from the worksheet.
' (More on that, see "USAGE EXAMPLES")
'
' This function offers three optional parameters to the user, however using
' these should only be necessary in extremely rare situations.
' The best rule regarding their usage: Don't use them.
'
' In the following these parameters will still be explained.
'
'1) returnAll
'   In some exceptional cases it is possible to map one OneDrive WebPath to
'   multiple different localPaths. This can happen when multiple Business
'   OneDrive accounts are logged in on one device, and multiple of these have
'   access to the same OneDrive folder and they both decide to synchronize it or
'   add it as link to their MySite library.
'   Calling the function with returnAll:=True will return all valid localPaths
'   for the given WebPath, separated by two forward slashes (//). This should be
'   used with caution, as the return value of the function alone is, should
'   multiple local paths exist for the input webPath, not a valid local path
'   anymore.
'   An example of how to obtain all of the local paths could look like this:
'   Dim localPath as String, localPaths() as String
'   localPath = GetLocalPath(webPath, True)
'   If Not localPath Like "http*" Then
'       localPaths = Split(localPath, "//")
'   End If
'
'2) preferredMountPointOwner
'   This parameter deals with the same problem as 'returnAll'
'   If the function gets called with returnAll:=False (default), and multiple
'   localPaths exist for the given WebPath, the function will just return any
'   one of them, as usually, it shouldn't make a difference, because the result
'   directories at both of these localPaths are mirrored versions of the same
'   webPath. Nevertheless, this option lets the user choose, which mountPoint
'   should be chosen if multiple localPaths are available. Each localPath is
'  'owned' by an OneDrive Account. If a WebPath is synchronized twice, this can
'   only happen by synchronizing it with two different accounts, because
'   OneDrive prevents you from synchronizing the same folder twice on a single
'   account. Therefore, each of the different localPaths for a given WebPath
'   has a unique 'owner'. preferredMountPointOwner lets the user select the
'   localPath by specifying the account the localPath should be owned by.
'   This is done by passing the Email address of the desired account as
'   preferredMountPointOwner.
'   For example, you have two different Business OneDrive accounts logged in,
'   foo.bar@business1.com and foo.bar@business2.com
'   Both synchronize the WebPath:
'   webPath = "https://business1.sharepoint.com/sites/TestLib/Documents/" & _
              "Test/Test/Test/test.xlsm"
'
'   The first one has added it as a link to his personal OneDrive, the local
'   path looks like this:
'   C:\Users\username\OneDrive - Business1\TestLinkParent\Test - TestLinkLib\...
'   ...Test\test.xlsm
'
'   The second one just synchronized it normally, the localPath looks like this:
'   C:\Users\username\Business1\TestLinkLib - Test\Test\test.xlsm
'
'   Calling GetLocalPath like this:
'   GetLocalPath(webPath,,, "foo.bar@business1.com") will return:
'   C:\Users\username\OneDrive - Business1\TestLinkParent\Test - TestLinkLib\...
'   ...Test\test.xlsm
'
'   Calling it like this:
'   GetLocalPath(webPath,,, "foo.bar@business2.com") will return:
'   C:\Users\username\Business1\TestLinkLib - Test\Test\test.xlsm
'
'   And calling it like this:
'   GetLocalPath(webPath,, True) will return:
'   C:\Users\username\OneDrive - Business1\TestLinkParent\Test - TestLinkLib\...
'   ...Test\test.xlsm//C:\Users\username\Business1\TestLinkLib - Test\Test\...
'   ...test.xlsm
'
'   Calling it normally like this:
'   GetLocalPath(webPath) will return any one of the two localPaths, so:
'   C:\Users\username\OneDrive - Business1\TestLinkParent\Test - TestLinkLib\...
'   ...Test\test.xlsm
'   OR
'   C:\Users\username\Business1\TestLinkLib - Test\Test\test.xlsm
'
'3) rebuildCache
'   The function creates a "translation" dictionary from the OneDrive settings
'   files and then uses this dictionary to "translate" WebPaths to LocalPaths.
'   This dictionary is implemented as a static variable to the function doesn't
'   have to recreate it every time it is called. It is written on the first
'   function call and reused on all the subsequent calls, making them faster.
'   If the function is called with rebuildCache:=True, this dictionary will be
'   rewritten, even if it was already initialized.
'   Note that it is not necessary to use this parameter manually, even if a new
'   MountPoint was added to the OneDrive, or a new OneDrive account was logged
'   in since the last function call because the function will automatically
'   determine if any of those cases occurred, without sacrificing performance.
'*******************************************************************************
Option Explicit

''*******************************************************************************
'' USAGE EXAMPLES:
'' Excel:
'Private Sub TestGetLocalPathExcel()
'    Debug.Print GetLocalPath(ThisWorkbook.FullName)
'    Debug.Print GetLocalPath(ThisWorkbook.path)
'End Sub
'
' Usage as User Defined Function (UDF):
' You might have to replace ; with , in the formulas depending on your settings.
' Add this formula to any cell, to get the local path of the workbook:
' =GetLocalPath(LEFT(CELL("filename";A1);FIND("[";CELL("filename";A1))-1))
'
' To get the local path including the filename (the FullName), use this formula:
' =GetLocalPath(LEFT(CELL("filename";A1);FIND("[";CELL("filename";A1))-1) &
' TEXTAFTER(TEXTBEFORE(CELL("filename";A1);"]");"["))
'
''Word:
'Private Sub TestGetLocalPathWord()
'    Debug.Print GetLocalPath(ThisDocument.FullName)
'    Debug.Print GetLocalPath(ThisDocument.path)
'End Sub
'
''PowerPoint:
'Private Sub TestGetLocalPathPowerPoint()
'    Debug.Print GetLocalPath(ActivePresentation.FullName)
'    Debug.Print GetLocalPath(ActivePresentation.path)
'End Sub
'*******************************************************************************

'This Function will convert a OneDrive/SharePoint Url path, e.g. Url containing
'https://d.docs.live.net/; .sharepoint.com/sites; my.sharepoint.com/personal/...
'to the locally synchronized path on your current pc or mac, e.g. a path like
'C:\users\username\OneDrive\ on Windows; or /Users/username/OneDrive/ on MacOS,
'if you have the remote directory locally synchronized with the OneDrive app.
'If no local path can be found, the input value will be returned unmodified.
'Author: Guido Witt-Dorring
'Source: https://gist.github.com/guwidoe/038398b6be1b16c458365716a921814d
'        https://stackoverflow.com/a/73577057/12287457
"""
from pathlib import Path
from datetime import datetime
import re
import os

_locToWebCollAndUpdateTime = None, None


def get_onedrive_path(url: str,
                      return_all: bool=False,
                      preferred_mount_point_owner: str="",
                      rebuild_cache: bool=False):
    """
    Return a local path from a OneDrive URL.
    If the URL can't be resolved to a local path then returns None.
    """
    global _locToWebCollAndUpdateTime
    locToWebColl, lastCacheUpdate = _locToWebCollAndUpdateTime

    maxDirName = 255
    pmpo = preferred_mount_point_owner.lower()

    settPaths = [Path(os.environ["LOCALAPPDATA"], "Microsoft", "OneDrive", "settings")]
    clpPath = Path(os.environ["LOCALAPPDATA"], "Microsoft", "Office", "CLP")

    oneDriveSettDirs = []
    for settPath in settPaths:
        if settPath.exists() and settPath.is_dir():
            for dirName in settPath.iterdir():
                  if dirName.is_dir() \
                  and (dirName.name == "Personal" or re.match(r"^Business\d+$", dirName.name)):
                        oneDriveSettDirs.append(dirName)

    requiredFiles = []
    for vDir in oneDriveSettDirs:
         for fileName in vDir.iterdir():
              if fileName.is_file() and (
                re.match(r"^([\da-f]{16}|([\da-f]{8}-[\da-f]{4}-[\da-f]{4}-[\da-f]{4}-[\da-f]{12}))\.(ini|dat)$", fileName.name, re.I) \
                or re.match(r"^ClientPolicy.*\.ini$", fileName.name, re.I) \
                or re.match(r"^GroupFolders.ini$", fileName.name, re.I) \
                or re.match(r"^global.ini$", fileName.name, re.I)):
                   requiredFiles.append(fileName)

    # This part should ensure perfect accuracy despite the mount point cache
    # while sacrificing almost no performance at all by querying FileDateTimes.
    if not rebuild_cache:
        for vFile in requiredFiles:
            if lastCacheUpdate is None or vFile.stat().st_mtime > lastCacheUpdate:
                 rebuild_cache = True  # full cache refresh is required
                 break

    if locToWebColl is not None and not rebuild_cache:
        resColl = {}
        for vItem in locToWebColl.values():
            locRoot = vItem[0]
            webRoot = vItem[1]
            if webRoot in url:
                resColl[vItem[2]] = Path(url.replace(webRoot, str(locRoot)))

        if resColl:
            if return_all:
                return list(resColl.values())

            if pmpo in resColl:
                return resColl[pmpo]
            
            return next(iter(resColl.values()))
        
        return None

    # Rebuild cache
    lastCacheUpdate = datetime.now().timestamp()
    locToWebColl = {}        

    sig1 = b"\2"
    sig2 = b"\1\0\0\0"
    vbNullByte = b"\0"
    sig3 = b"\0\0"

    # Writing locToWebColl using .ini and .dat files in the OneDrive settings:
    for vDir in oneDriveSettDirs:  # One folder per logged in OD account
        globalIni = vDir / "global.ini"
        if not globalIni.exists():
            continue
        
        # Get the cid from the global.ini file
        cid = None
        with open(globalIni, "rt", encoding="utf-16le") as fh:
            for line in fh.readlines():
                match = re.match(r"cid\s*=\s*([\da-f]{16}|(?:[\da-f]{8}-[\da-f]{4}-[\da-f]{4}-[\da-f]{4}-[\da-f]{12}))", line, re.I)
                if match:
                        cid = match.group(1)
                        break
            else:
                continue
        
        dirName = vDir.name
        iniFile = vDir / f"{cid}.ini"
        datFile = vDir / f"{cid}.dat"
        if not iniFile.exists() or not datFile.exists():
            continue

        # Get email for business accounts
        email = None
        for fileName in clpPath.iterdir():
            if fileName.is_file():
                i = fileName.name.rfind(cid)
                if i > 0:
                    email = fileName.name[:i-1]
                    break

        # Read all the ClientPloicy*.ini files:
        cliPolColl = {}
        for fileName in vDir.iterdir():
            if re.match(r"^ClientPolicy.*\.ini", fileName.name, re.I):
                cliPol = cliPolColl[fileName.name] = {}
                with open(fileName, "rt", encoding="utf-16le") as fh:
                    for line in fh.readlines():
                        if "=" in line:
                            tag, s = re.split(r"\s*=\s*", line.strip(), 1)
                            if tag == "DavUrlNamespace":
                                cliPol[tag] = s
                            elif tag in ("SiteID", "IrmLibraryId", "WebID"): # Only used for backup method later
                                s = s.lower().replace("-", "")
                                if len(s) > 3:
                                    s = s[1:-1]
                                cliPol[tag] = s

        # Read cid.dat file (assume we'll have enough memory to read it all in)
        odFolders = {}
        with open(datFile, "rb") as fh:
            s = fh.read()

        if re.match("Personal", vDir.name, re.I):
            folderIdPattern = re.compile(r"^[a-z0-9]{16}!\d{3}.*$", re.I)
        else:
            folderIdPattern = re.compile(r"^[a-z0-9]{32}$", re.I)

        for vItem in range(16, 0, -8):
            i = s.find(sig2, vItem)               # Search pattern in cid.dat
            while i > vItem and i < len(s) - 168: # and confirm with another
                if s[i-vItem:i-vItem+1] == sig1:  # one
                    i = i + 8
                    n = s.find(vbNullByte, i) - i
                    n = max(min(n, 39), 0)
                    folderID = s[i:i+n].decode("ansi")

                    i = i + 39
                    n = s.find(vbNullByte, i) - i
                    n = max(min(n, 39), 0)
                    parentID = s[i:i+n].decode("ansi")

                    i = i + 121
                    if folderIdPattern.match(folderID) and folderIdPattern.match(parentID):
                        n = int((s.find(sig3, i) + 1 - i) / 2) * 2
                        n = max(min(n, maxDirName*2), 0)
                        folderName = s[i:i+n].decode("utf-16le")
                        odFolders[folderID] = (parentID, folderName)

                # Find next sig2 in cid.dat
                i = s.find(sig2, i+1)

        # Read cid.ini file
        with open(iniFile, "rt", encoding="utf-16le") as fh:
            if re.match(r"^Business\d+$", dirName, re.I):
                # Settings files for a business OD account
                # Max 9 Business OneDrive accounts can be signed in at a time.
                libNrToWebColl = {}
                locToWebColl = {}
                mainMount = ""
                for line in fh.readlines():
                    line = line.strip()
                    if "=" in line:
                        parts = line.split('"')
                        tag, s = re.split(r"\s*=\s*", line, 1)
                        if tag == "libraryScope":  # One line per synchronized library
                            locRoot = parts[9]
                            syncFind = locRoot
                            syncID = parts[10].split(" ")[2]
                            if not locRoot:
                                libNr = line.split(" ")[2]
                            folderType = parts[3]
                            parts = parts[8].split(" ")
                            siteID = parts[1]
                            webID = parts[2]
                            libID = parts[3]
                            if not mainMount and folderType == "ODB":
                                mainMount = locRoot
                                fileName = "ClientPolicy.ini"
                                mainSyncID = syncID
                                mainSyncFind = syncFind
                            else:
                                fileName = f"ClientPolicy_{libID}{siteID}.ini"

                            webRoot = cliPolColl.get(fileName, {}).get("DavUrlNamespace")
                            if not webRoot:  # Backup method to find webRoot
                                for vItem in cliPolColl.values():
                                    if vItem.get("SiteID") == siteID \
                                    and vItem.get("WebID") == webID \
                                    and vItem.get("IrmLibraryId") == libID:   
                                        webRoot = vItem.get("DavUrlNamespace")
                                        break

                            if not webRoot:
                                continue

                            if not locRoot:
                                libNrToWebColl[libID] = (libID, webRoot)
                            else:
                                locToWebColl[locRoot] = (locRoot, webRoot, email, syncID, syncFind, dirName)

                        elif tag == "libraryFolder":   # One line per synchronized library folder
                            libNr = re.split(r"\s+", line)[3]
                            locRoot = parts[1]
                            syncFind = locRoot
                            s = ""
                            parentID = re.split(r"\s+", line)[4][:32]
                            while parentID in odFolders:
                                s = f"{odFolders[parentID][1]}/{s}"
                                parentID = odFolders[parentID][0]
                            webRoot = f"{libNrToWebColl[libNr][1]}/{s}"
                            locToWebColl[locRoot] = (locRoot, webRoot, email, syncID, syncFind, dirName)

                        elif tag == "AddedScope":  # One line per folder added as link to personal
                            relPath = parts[5].strip()
                            parts = re.split(r"\s+", parts[4])
                            siteID = parts[1]
                            webID = parts[2]
                            libID = parts[3]
                            linkID = parts[4]
                            fileName = f"ClientPolicy_{libID}{siteID}{linkID}.ini"
                            webRoot = cliPolColl.get(fileName, {}).get("DavUrlNamespace")
                            if not webRoot:  # Backup method to find webRoot
                                for vItem in cliPolColl.values():
                                    if vItem.get("SiteID") == siteID \
                                    and vItem.get("WebID") == webID \
                                    and vItem.get("IrmLibraryId") == libID:   
                                        webRoot = vItem.get("DavUrlNamespace")
                                        break

                            if not webRoot:
                                continue

                            webRoot = Path(webRoot) / relPath

                            s = ""
                            parentID = re.split(r"\s+", line)[3][:32]
                            while parentID in odFolders:  # If link is not at the bottom of the personal library
                                s = Path(odFolders[parentID][1]) / s # add folders below mount point to locRoot
                                parentID = odFolders[parentID][0]

                            locRoot = Path(mainMount) / s
                            locToWebColl[locRoot] = (locRoot, webRoot, email, mainSyncID, mainSyncFind, dirName)

            elif dirName == "Personal":  # Settings files for a personal OD account
                # Only one Personal OneDrive account can be signed in at a time.
                for line in fh.readlines():  # Loop should exit at first line
                    if re.match(r"^Library\s+=\s.*$", line, re.I):
                        parts = re.split("\"", line)
                        locRoot = parts[3]
                        syncFind = locRoot
                        syncID = re.split(r"\s+", parts[4])[2]
                        break

                # This file may be missing if the personal OD was logged out of the OneDrive app
                webRoot = cliPolColl.get("ClientPolicy.ini", {}).get("DavUrlNamespace")
                if not webRoot:
                    continue

                locToWebColl[locRoot] = (locRoot, f"{webRoot}/{cid}", email, syncID, syncFind, dirName)

                # Read GroupFolders.ini file
                groupFoldersIni = vDir / "GroupFolders.ini"
                if not groupFoldersIni.exists():
                    continue

                cid = None
                with open(groupFoldersIni, "rt", encoding="utf-16le") as gfh:
                    for line in gfh.readlines():
                        line = line.rstrip()
                        if re.match(r"^.*_BaseUri\s+=\s+.*%$", line, re.I):
                            cid = line.rsplit("/", 1)[1][:16]
                            folderID = line.split("_", 1)[0]
                            locToWebColl[Path(locRoot) / odFolders[folderID][1]] = (
                                Path(locRoot) / odFolders[folderID][1], 
                                f"{webRoot}/{cid}/{line[len(folderID) + 9:]}",
                                email,
                                syncID,
                                syncFind,
                                dirName
                            )

    # Clean the finished "dictionary" up, remove trailing "\" and "/"
    tmpColl = {}
    for vItem in locToWebColl.values():
        locRoot = Path(vItem[0])
        webRoot = vItem[1].rstrip("/")
        syncFind = Path(vItem[4])
        tmpColl[locRoot] = (locRoot, webRoot, vItem[2], vItem[3], syncFind, locRoot)
    locToWebColl = tmpColl

    # Update the global cache object
    _locToWebCollAndUpdateTime = (locToWebColl, lastCacheUpdate)

    # Return the cached result        
    return get_onedrive_path(url, return_all, preferred_mount_point_owner, rebuild_cache=False)
