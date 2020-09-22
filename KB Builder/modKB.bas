Attribute VB_Name = "modKB"
Option Explicit
Private Const PHYSICALOFFSETX As Long = 112
Private Const PHYSICALOFFSETY As Long = 113
Private Const SWP_NOMOVE = &H2
Private Const HWND_TOPMOST = -1
Private Const SWP_NOSIZE As Long = 1
Private Const LB_DELETESTRING = &H182
Private Const CB_DELETESTRING = &H144
Private Const LB_FINDSTRINGEXACT = &H1A2
Private Const CB_FINDSTRINGEXACT = &H158

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type PALETTEENTRY
    peRed As Byte
    peGreen As Byte
    peBlue As Byte
    peFlags As Byte
End Type
Private Type LOGPALETTE
    palVersion As Integer
    palNumEntries As Integer
    palPalEntry(255) As PALETTEENTRY
End Type
Private Type PicBmp
    Size As Long
    Type As Long
        hBmp As Long
        hPal As Long
        reserved As Long
End Type
Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type
Private Type BrowseInfo
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type
'for switching task panel icon
Private Declare Function LoadImageAsString Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal uType As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal fuLoad As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc&, ByVal iCapabilitiy&) As Long
Private Declare Function GetSystemPaletteEntries Lib "gdi32" (ByVal hdc As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
Private Declare Function CreatePalette Lib "gdi32" (lpLogPalette As LOGPALETTE) As Long
Private Declare Function SelectPalette Lib "gdi32" (ByVal hdc&, ByVal hPalette&, ByVal bForceBackground&) As Long
Private Declare Function RealizePalette Lib "gdi32" (ByVal hdc&) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC&, ByVal x&, ByVal Y&, ByVal nWidth&, ByVal nHeight&, ByVal hSrcDC&, ByVal xSrc&, ByVal ySrc&, ByVal dwRop&) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc&) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Const WindowFlags As Long = SWP_NOMOVE Or SWP_NOSIZE
Private Const LR_LOADFROMFILE = &H10
Private Const LR_SHARED = &H8000&
Private Const IMAGE_ICON = 1
Private Const RASTERCAPS As Long = 38
Private Const RC_PALETTE As Long = &H100
Private Const SIZEPALETTE As Long = 104
Private Const BIF_RETURNONLYFSDIRS = 1
'detect if running in IDE (can't load from resources)
Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Public Type POINTAPI
    x       As Long
    Y       As Long
End Type
Public Type WINDOWPLACEMENT
    Length            As Long
    flags             As Long
    showCmd           As Long
    ptMinPosition     As POINTAPI
    ptMaxPosition     As POINTAPI
    rcNormalPosition  As RECT
End Type
Public Const GW_HWNDNEXT = 2
Public Const GW_CHILD = 5
Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Public Const TH32CS_SNAPPROCESS As Long = 2&
Public Const MAX_PATH As Long = 260
Public Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwflags As Long
    szexeFile As String * MAX_PATH
End Type
Public Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlgas As Long, ByVal lProcessID As Long) As Long
Public Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Public Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Public Declare Sub CloseHandle Lib "kernel32" (ByVal hPass As Long)
Private IsOnTop As Long
Private ViewHeadings(1 To 83) As String
Private Const LVM_FIRST As Long = &H1000
Private Const LVSCW_AUTOSIZE_USEHEADER As Long = -2
Public Enum FindWhere
    search_Text = 0
    search_SubItem = 1
    Search_Tag = 2
End Enum
Public Enum SearchType
    search_Partial = 1
    search_Whole = 0
End Enum
Private Const LVM_SETCOLUMNWIDTH As Long = (LVM_FIRST + 30)
Private Declare Function SendMessageLONG Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Const WM_USER = &H400
Private Const SB_GETRECT As Long = (WM_USER + 10)
Private Const EM_FORMATRANGE As Long = WM_USER + 57
Private Const LB_RESETCONTENT = &H184
Private Const CB_RESETCONTENT = &H14B
Public sProject As String
Public pPath As String
Public sProjPath As String
Public sProjHTML As String
Public sProjDb As String
Private xNode As Node
Public PropertiesFlds(1 To 10) As String
Public sProjContents As String
Public sProjRTF As String
Public sProjCnt As String
Public sProjHPJ As String
Public sProjHLP As String
Public FontHeadline As String
Public FontText As String
Public FontHeadlineSize As Long
Public FontTextSize As Long
Public FontHeadlineBold As Long
Public HeadlineColor As Long
Public PictureHeight As Long
Public PictureWidth As Long
Public HeadlineBackColor As Long
Public TextColor As Long
Public TextBackColor As Long
Public Title As String
Public Compression As String
Public Author As String
Public CompilerLocation As String
Public sProjIni As String
Public sProjLog As String
Public sStart As Long
Public sLength As Long
Public Quote As String
Public mnIndex As Long
Private cntNodes As Long
Private Const TV_FIRST As Long = &H1100
Private Const TVGN_ROOT As Long = &H0
Private Const TVM_GETNEXTITEM As Long = (TV_FIRST + 10)
Private Const TVM_DELETEITEM As Long = (TV_FIRST + 1)
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As Any) As Long
Private Const WM_PASTE = &H302
Public Const EM_LINEINDEX = &HBB
Public Type CharRange
    cpMin As Long     ' First character of range (0 for start of doc)
    cpMax As Long     ' Last character of range (-1 for end of doc)
End Type
Public Type FormatRange
    hdc As Long       ' Actual DC to draw on
    hdcTarget As Long ' Target DC for determining text formatting
    rc As RECT        ' Region of the DC to draw to (in twips)
    rcPage As RECT    ' Region of the entire DC (page size) (in twips)
    chrg As CharRange ' Range of text to draw (see above declaration)
End Type
Public Enum WordOperation
    wSpelling = 1
    wGrammar = 2
    wThesaurus = 3
    wSaveHTML = 4
End Enum
Public Enum FileOps
    foDelete = 0
    foMove = 1
    foCopy = 2
    foRename = 3
End Enum
Public mobjWord97 As Word.Application
Public blnWord97Loaded As Boolean
Public FM As String
Public KM As String
Public VM As String
Public AppTitle As String
Public retAnswer As Long
Public Type SHFILEOPSTRUCT
    hWnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Boolean
End Type
Private Const FO_DELETE = &H3
Private Const FO_MOVE As Long = &H1
Private Const FO_COPY As Long = &H2
Private Const FO_RENAME As Long = &H4
Private Const FOF_ALLOWUNDO = &H40
Public iPath As String
Public TopicName As String

Public Function GetExeFromHandle(hWnd As Long) As String
    On Error Resume Next
    Dim threadID As Long, processID As Long, hSnapshot As Long
    Dim uProcess As PROCESSENTRY32, rProcessFound As Long
    Dim i As Integer, szExename As String
    ' Get ID for window thread
    threadID = GetWindowThreadProcessId(hWnd, processID)
    ' Check if valid
    If threadID = 0 Or processID = 0 Then Exit Function
    ' Create snapshot of current processes
    hSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
    ' Check if snapshot is valid
    If hSnapshot = -1 Then Exit Function
    'Initialize uProcess with correct size
    uProcess.dwSize = Len(uProcess)
    'Start looping through processes
    rProcessFound = ProcessFirst(hSnapshot, uProcess)
    Do While rProcessFound
        If uProcess.th32ProcessID = processID Then
            'Found it, now get name of exefile
            i = InStr(1, uProcess.szexeFile, Chr$(0))
            If i > 0 Then szExename = Left$(uProcess.szexeFile, i - 1)
            Exit Do
        Else
            'Wrong ID, so continue looping
            rProcessFound = ProcessNext(hSnapshot, uProcess)
        End If
        Err.Clear
    Loop
    Call CloseHandle(hSnapshot)
    GetExeFromHandle = szExename
    Err.Clear
End Function
'fix app.path bug(gery)
Public Function AppPath() As String
    On Error Resume Next
    Dim sPath As String
    sPath = App.Path
    If Not Right$(sPath, 1) = "\" Then
        sPath = sPath & "\"
    End If
    AppPath = sPath
    Err.Clear
End Function
Public Sub HandleError(sComponent As String, sMethod As String, lErrLine As Long)
    On Error Resume Next
    MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "in " & sComponent & "." & sMethod & " at line " & lErrLine, vbExclamation Or vbOKOnly, "Error"
    Err.Clear
End Sub
'if running in IDE then cannot load from internal resources
Public Function IsRunningInIDE(Optional ByVal VbExeName As String = "VB6.EXE") As Boolean
    On Error Resume Next
    Dim lRet    As Long
    Dim sBuffer As String
    sBuffer = Space$(2048)
    lRet = GetModuleFileName(0&, sBuffer, Len(sBuffer))
    sBuffer = Left$(sBuffer, lRet)
    If StrComp(Right$(sBuffer, Len(VbExeName) + 1), "\" & VbExeName, vbTextCompare) = 0 Then
        IsRunningInIDE = True
    Else
        IsRunningInIDE = False
    End If
    Err.Clear
End Function
Public Function GetResIconHandle(ByVal sIconResName As String, ByVal DesiredSizeX As Long, ByVal DesiredSizeY As Long) As Long
    On Error GoTo ErrHandler
    If IsRunningInIDE Then
        GetResIconHandle = LoadImageAsString(0, App.Path & "\resources\" & sIconResName & ".ico", IMAGE_ICON, DesiredSizeX, DesiredSizeY, LR_LOADFROMFILE)
    Else
        GetResIconHandle = LoadImageAsString(App.hInstance, sIconResName, IMAGE_ICON, DesiredSizeX, DesiredSizeY, LR_SHARED)
    End If
    Err.Clear
    Exit Function
ErrHandler:
    HandleError "MIcon", "GetResIconHandle", Erl
    Err.Clear
End Function
Public Function Dir_Exists(ByVal strFile As String) As Boolean
    On Error Resume Next
    Dim fs As FileSystemObject
    Set fs = New FileSystemObject
    Dir_Exists = fs.FolderExists(strFile)
    Set fs = Nothing
    Err.Clear
End Function

'    On Error Resume Next
'    ' Get the dimensions of the window.
'    R = GetWindowRect(hWndActive, RectActive)
'    ' handle and return the Resulting Picture object.
'    Set CaptureScreen = CaptureWindow(hWndActive, False, 0, 0, RectActive.Right - RectActive.Left, RectActive.Bottom - RectActive.Top)
'    Err.Clear
'End Function
Public Sub SaveScreen(picControl As PictureBox, ByVal strFile As String, hWnd As Long)
    On Error Resume Next
    Dim EndTime As Date
    EndTime = DateAdd("s", 2, Now)
    Do Until Now > EndTime
        DoEvents
        Err.Clear
    Loop
    picControl.AutoRedraw = True
    picControl.AutoSize = True
    ApplicationOnTop hWnd
    DoEvents
    Set picControl.Picture = CaptureActiveWindow
    DoEvents
    SavePicture picControl.Picture, strFile
    DoEvents
    Err.Clear
End Sub
Private Function CreateBitmapPicture(ByVal hBmp As Long, ByVal hPal As Long) As Picture
    On Error Resume Next
    Dim Pic As PicBmp
    ' IPicture requires a reference to "Standard OLE Types."
    Dim IPic As IPicture
    Dim IID_IDispatch As GUID
    Dim R As Long
    ' Fill in with IDispatch Interface ID.
    With IID_IDispatch
        .Data1 = &H20400
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With
    ' Fill Pic with necessary parts.
    With Pic
        .Size = Len(Pic)          ' Length of structure.
        .Type = vbPicTypeBitmap   ' Type of Picture (bitmap).
        .hBmp = hBmp              ' Handle to bitmap.
        .hPal = hPal              ' Handle to palette (may be null).
    End With
    ' Create Picture object.
    R = OleCreatePictureIndirect(Pic, IID_IDispatch, 1, IPic)
    ' Return the new Picture object.
    Set CreateBitmapPicture = IPic
    Err.Clear
End Function
Private Function CaptureWindow(ByVal hWndSrc As Long, ByVal Client As Boolean, ByVal LeftSrc As Long, ByVal TopSrc As Long, ByVal WidthSrc As Long, ByVal HeightSrc As Long) As Picture
    On Error Resume Next
    Dim hDCMemory As Long
    Dim hBmp As Long
    Dim hBmpPrev As Long
    Dim R As Long
    Dim hDCSrc As Long
    Dim hPal As Long
    Dim hPalPrev As Long
    Dim RasterCapsScrn As Long
    Dim HasPaletteScrn As Long
    Dim PaletteSizeScrn As Long
    Dim LogPal As LOGPALETTE
    ' Depending on the value of Client get the proper device context.
    If Client Then
        hDCSrc = GetDC(hWndSrc) ' Get device context for client area.
    Else
        hDCSrc = GetWindowDC(hWndSrc) ' Get device context for entire
        ' window.
    End If
    ' Create a memory device context for the copy process.
    hDCMemory = CreateCompatibleDC(hDCSrc)
    ' Create a bitmap and place it in the memory DC.
    hBmp = CreateCompatibleBitmap(hDCSrc, WidthSrc, HeightSrc)
    hBmpPrev = SelectObject(hDCMemory, hBmp)
    ' Get screen properties.
    RasterCapsScrn = GetDeviceCaps(hDCSrc, RASTERCAPS) ' Raster
    ' capabilities.
    HasPaletteScrn = RasterCapsScrn And RC_PALETTE       ' Palette
    ' support.
    PaletteSizeScrn = GetDeviceCaps(hDCSrc, SIZEPALETTE) ' Size of
    ' palette.
    ' If the screen has a palette make a copy and realize it.
    If HasPaletteScrn And (PaletteSizeScrn = 256) Then
        ' Create a copy of the system palette.
        LogPal.palVersion = &H300
        LogPal.palNumEntries = 256
        R = GetSystemPaletteEntries(hDCSrc, 0, 256, LogPal.palPalEntry(0))
        hPal = CreatePalette(LogPal)
        ' Select the new palette into the memory DC and realize it.
        hPalPrev = SelectPalette(hDCMemory, hPal, 0)
        R = RealizePalette(hDCMemory)
    End If
    ' Copy the on-screen image into the memory DC.
    R = BitBlt(hDCMemory, 0, 0, WidthSrc, HeightSrc, hDCSrc, LeftSrc, TopSrc, vbSrcCopy)
    ' Remove the new copy of the  on-screen image.
    hBmp = SelectObject(hDCMemory, hBmpPrev)
    ' If the screen has a palette get back the palette that was
    ' selected in previously.
    If HasPaletteScrn And (PaletteSizeScrn = 256) Then
        hPal = SelectPalette(hDCMemory, hPalPrev, 0)
    End If
    ' Release the device context resources back to the system.
    R = DeleteDC(hDCMemory)
    R = ReleaseDC(hWndSrc, hDCSrc)
    ' bitmap and palette handles. Then return the resulting picture
    ' object.
    Set CaptureWindow = CreateBitmapPicture(hBmp, hPal)
    Err.Clear
End Function
Private Function IsPathFile(ByVal strPath As String) As Boolean
    On Error Resume Next
    Dim fso As Scripting.FileSystemObject
    Set fso = New Scripting.FileSystemObject
    Dim fsoFile As Scripting.File
    If fso.FileExists(strPath) = True Then
        Set fsoFile = fso.GetFile(strPath)
        If TypeName(fsoFile) = "Nothing" Then
            IsPathFile = False
        Else
            IsPathFile = True
        End If
    Else
        IsPathFile = False
    End If
    Set fsoFile = Nothing
    Set fso = Nothing
    Err.Clear
End Function
Public Function NextNewFile(ByVal StrFilePath As String, Optional IsCopy As Boolean = False) As String
    On Error Resume Next
    Dim bExist As Boolean
    Dim fCnt As Long
    Dim sExtension As String
    Dim fso As Scripting.FileSystemObject
    Set fso = New Scripting.FileSystemObject
    Dim fsoFile As Scripting.File
    Dim pPath As String
    Dim fName As String
    Set fsoFile = fso.GetFile(StrFilePath)
    pPath = fso.GetParentFolderName(StrFilePath)
    If Right$(pPath, 1) = "\" Then
        pPath = Left$(pPath, Len(pPath) - 1)
    End If
    fName = fso.GetBaseName(StrFilePath)
    sExtension = fso.GetExtensionName(StrFilePath)
    fCnt = 0
    bExist = IsPathFile(StrFilePath)
    Do Until bExist = False
        fCnt = fCnt + 1
        StrFilePath = pPath & "\" & fName & " " & CStr(fCnt) & "." & sExtension
        If IsCopy = True Then
            If fCnt - 1 < 0 Then
                StrFilePath = pPath & "\Copy Of " & fName & "." & sExtension
            Else
                StrFilePath = pPath & "\Copy " & CStr(fCnt) & " Of " & fName & "." & sExtension
            End If
        End If
        DoEvents
        bExist = IsPathFile(StrFilePath)
        Err.Clear
    Loop
    NextNewFile = StrFilePath
    Err.Clear
End Function
Function FindWindowLike(TheWindows As clsWindows, hWndArray() As Long, ByVal hWndStart As Long, ByVal WindowText As String, Optional ByVal Classname As String = "ThunderRT6FormDC") As Integer
    On Error Resume Next
    Dim hWnd As Long
    Dim swindowtext As String
    Dim sClassname As String
    Dim R As Long
    Static level As Integer
    Static found As Integer
    'Initialize if necessary
    If level = 0 Then
        found = 0
        ReDim hWndArray(0 To 0)
        If hWndStart = 0 Then
            hWndStart = GetDesktopWindow()
            swindowtext = Space$(255)
            R = GetWindowText(hWndStart, swindowtext, 255)
            swindowtext = Left$(swindowtext, R)
            sClassname = Space$(255)
            R = GetClassName(hWndStart, sClassname, 255)
            sClassname = Left$(sClassname, R)
            TheWindows.Add hWndStart, sClassname, swindowtext
        End If
    End If
    'Increase recursion counter
    level = level + 1
    'Get first child window
    hWnd = GetWindow(hWndStart, GW_CHILD)
    Do Until hWnd = 0
        'Search children by recursion
        R = FindWindowLike(TheWindows, hWndArray, hWnd, WindowText, Classname)
        'Get the window text and class name
        swindowtext = Space$(255)
        R = GetWindowText(hWnd, swindowtext, 255)
        swindowtext = Left$(swindowtext, R)
        sClassname = Space$(255)
        R = GetClassName(hWnd, sClassname, 255)
        sClassname = Left$(sClassname, R)
        'Check that window matches the search parameters:
        If (swindowtext Like WindowText) And (sClassname Like Classname) Then
            found = found + 1
            ReDim Preserve hWndArray(0 To found)
            hWndArray(found) = hWnd
            If GetParent(hWnd) = 0 Then
                TheWindows.Add hWndArray(found), sClassname, swindowtext
            Else
                TheWindows.Add hWndArray(found), sClassname, swindowtext
            End If
        End If
        'Get next child window:
        hWnd = GetWindow(hWnd, GW_HWNDNEXT)
        Err.Clear
    Loop
    'Decrement recursion counter
    level = level - 1
    'Return the number of windows found
    FindWindowLike = found
    Err.Clear
End Function
Public Function FileName_Validate(ByVal strValue As String) As String
    On Error Resume Next
    Dim fPath As String
    Dim fFileN As String
    Dim fExt As String
    fPath = File_Token(strValue, "p")
    fFileN = File_Token(strValue, "fo")
    fExt = File_Token(strValue, "e")
    fFileN = Replace$(fFileN, "\", "")
    fFileN = Replace$(fFileN, "/", "")
    fFileN = Replace$(fFileN, ":", "")
    fFileN = Replace$(fFileN, "*", "")
    fFileN = Replace$(fFileN, "?", "")
    fFileN = Replace$(fFileN, Chr$(34), "")
    fFileN = Replace$(fFileN, "<", "")
    fFileN = Replace$(fFileN, ">", "")
    fFileN = Replace$(fFileN, "|", "")
    FileName_Validate = fPath & "\" & fFileN & "." & fExt
    Err.Clear
End Function
Public Function File_Token(ByVal strFileName As String, Optional ByVal Sretrieve As String = "F", Optional ByVal Delim As String = "\") As String
    On Error Resume Next
    Dim intNum As Long
    Dim sNew As String
    File_Token = strFileName
    Select Case UCase$(Sretrieve)
    Case "D"
        File_Token = Left$(strFileName, 3)
    Case "F"
        intNum = InStrRev(strFileName, Delim)
        If intNum <> 0 Then
            File_Token = Mid$(strFileName, intNum + 1)
        End If
    Case "P"
        intNum = InStrRev(strFileName, Delim)
        If intNum <> 0 Then
            File_Token = Mid$(strFileName, 1, intNum - 1)
        End If
    Case "E"
        intNum = InStrRev(strFileName, ".")
        If intNum <> 0 Then
            File_Token = Mid$(strFileName, intNum + 1)
        End If
    Case "FO"
        sNew = strFileName
        intNum = InStrRev(sNew, Delim)
        If intNum <> 0 Then
            sNew = Mid$(sNew, intNum + 1)
        End If
        intNum = InStrRev(sNew, ".")
        If intNum <> 0 Then
            sNew = Left$(sNew, intNum - 1)
        End If
        File_Token = sNew
    Case "PF"
        intNum = InStrRev(strFileName, ".")
        If intNum <> 0 Then
            File_Token = Left$(strFileName, intNum - 1)
        End If
    End Select
    Err.Clear
End Function
Public Function BrowseForFolder(hWndOwner As Long, sPrompt As String) As String
    On Error Resume Next
    Dim iNull As Integer
    Dim lpIDList As Long
    Dim lResult As Long
    Dim sPath As String
    Dim udtBI As BrowseInfo
    With udtBI
        .hWndOwner = hWndOwner
        .lpszTitle = lstrcat(sPrompt, "")
        .ulFlags = BIF_RETURNONLYFSDIRS
    End With
    lpIDList = SHBrowseForFolder(udtBI)
    If lpIDList Then
        sPath = String$(MAX_PATH, 0)
        lResult = SHGetPathFromIDList(lpIDList, sPath)
        Call CoTaskMemFree(lpIDList)
        iNull = InStr(sPath, vbNullChar)
        If iNull Then
            sPath = Left$(sPath, iNull - 1)
        End If
    End If
    BrowseForFolder = sPath
    Err.Clear
End Function
Function lngStartDoc(ByVal Docname As String, Optional ByVal Operation As String = "Open", Optional ByVal WindowState As Long = 1) As Long
    On Error Resume Next
    Dim Scr_hDC As Long
    Dim sDir As String
    sDir = File_Token(Docname, "d")
    Scr_hDC = GetDesktopWindow()
    lngStartDoc = ShellExecute(Scr_hDC, Operation, Docname, "", sDir, WindowState)
    Err.Clear
End Function
Public Function ViewFile(ByVal FileName As String, Optional ByVal Operation As String = "Open", Optional ByVal WindowState As Long = 1) As Boolean
    On Error Resume Next
    Dim R As Long
    R = lngStartDoc(FileName, Operation, WindowState)
    If R <= 32 Then
        ' there was an error
        Beep
        MsgBox "An error occurred while opening your file." & vbCr & "The possibility is that the selected entry does not have" & vbCr & "a link in the registry to open it with.", vbOKOnly + vbExclamation + vbApplicationModal, "Viewer Error"
        ViewFile = False
    Else
        ViewFile = True
    End If
    Err.Clear
End Function
Public Sub IniPropertiesFlds()
    On Error Resume Next
    ' Initializing Properties
    PropertiesFlds(1) = "FullPath"
    PropertiesFlds(2) = "Title"
    PropertiesFlds(3) = "Context"
    PropertiesFlds(4) = "Number"
    PropertiesFlds(5) = "Browse"
    PropertiesFlds(6) = "Keywords"
    PropertiesFlds(7) = "Macros"
    PropertiesFlds(8) = "FootNotes"
    PropertiesFlds(9) = "Contents"
    PropertiesFlds(10) = "Popup"
    Err.Clear
End Sub
Public Sub ResizePictureSource(startingPic As PictureBox, destinationPic As PictureBox)
    On Error Resume Next
    Dim ratioX As Double
    Dim ratioY As Double
    Dim x As Double
    Dim Y As Double
    Dim xTot As Double
    Dim YTot As Double
    Dim theColor As Double
    Dim realX As Double
    Dim realY As Double
    ratioX = startingPic.ScaleWidth / destinationPic.ScaleWidth
    ratioY = startingPic.ScaleHeight / destinationPic.ScaleHeight
    xTot = startingPic.ScaleWidth
    YTot = startingPic.ScaleHeight
    For x = 0 To xTot Step ratioX
        For Y = 0 To YTot Step ratioY
            'get the color of the startingPic
            theColor = startingPic.Point(x, Y)
            'find the corresponding x and y values
            ' for the resized destination pic
            realX = ratioX ^ -1 * x
            realY = ratioY ^ -1 * Y
            destinationPic.PSet (realX, realY), theColor
            Err.Clear
        Next
        Err.Clear
    Next
    Err.Clear
End Sub
Public Function StringAsc(ByVal strValue As String) As String
    On Error Resume Next
    Dim rsCnt As Long
    Dim rsTot As Long
    Dim rsStr As String
    Dim rsCur As String
    Dim rsInt As Integer
    rsStr = ""
    rsTot = Len(strValue)
    For rsCnt = 1 To rsTot
        rsCur = Mid$(strValue, rsCnt, 1)
        rsInt = Asc(rsCur)
        rsStr = rsStr & CStr(rsInt)
        Err.Clear
    Next
    StringAsc = rsStr
    Err.Clear
End Function
Public Function Topic_Validate(ByVal strValue As String) As String
    On Error Resume Next
    strValue = Replace$(strValue, "#", "")
    strValue = Replace$(strValue, "=", "")
    strValue = Replace$(strValue, "+", "")
    strValue = Replace$(strValue, "@", "")
    strValue = Replace$(strValue, "*", "")
    strValue = Replace$(strValue, "%", "")
    strValue = Replace$(strValue, ">", "")
    strValue = Replace$(strValue, "\", "")
    strValue = Replace$(strValue, "/", "")
    strValue = Replace$(strValue, "$", "")
    Topic_Validate = Trim$(Replace$(strValue, "!", ""))
    Err.Clear
End Function
Public Function AlphaNumericOnly(ByVal strValue As String) As String
    On Error Resume Next
    Dim rsCnt As Long
    Dim rsTot As Long
    Dim rsStr As String
    Dim rsCur As String
    rsStr = ""
    rsTot = Len(strValue)
    For rsCnt = 1 To rsTot
        rsCur = Mid$(strValue, rsCnt, 1)
        If InStr(1, "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ ", UCase$(rsCur)) > 0 Then
            rsStr = rsStr & rsCur
        End If
        Err.Clear
    Next
    AlphaNumericOnly = rsStr
    Err.Clear
End Function
Public Function Context_Validate(ByVal strValue As String) As String
    On Error Resume Next
    strValue = Topic_Validate(strValue)
    strValue = AlphaNumericOnly(strValue)
    Context_Validate = Trim$(Replace$(strValue, " ", ""))
    Err.Clear
End Function
Public Function TreeView_AddNode(TV As TreeView, ByVal ParentName As String, ByVal ChildText As String, ByVal Image As String, ByVal strRelationship As String) As String
    On Error Resume Next
    Select Case LCase$(strRelationship)
    Case "child"
        Set xNode = TV.Nodes.Add(ParentName, tvwChild, , ChildText, Image, Image)
    Case "first"
        Set xNode = TV.Nodes.Add(ParentName, tvwFirst, , ChildText, Image, Image)
    Case "last"
        Set xNode = TV.Nodes.Add(ParentName, tvwLast, , ChildText, Image, Image)
    Case "next"
        Set xNode = TV.Nodes.Add(ParentName, tvwNext, , ChildText, Image, Image)
    Case "previous"
        Set xNode = TV.Nodes.Add(ParentName, tvwPrevious, , ChildText, Image, Image)
    End Select
    xNode.Key = xNode.Fullpath
    TV.Nodes.Item(1).Expanded = True
    TreeView_AddNode = xNode.Fullpath
    Err.Clear
End Function
Public Sub TreeView_AddNodePopup(TV As TreeView, ByVal ParentName As String, ByVal ChildText As String)
    On Error Resume Next
    TV.Nodes.Add ParentName, tvwChild, , ChildText, "Popup"
    TV.Nodes.Item(1).Expanded = True
    Err.Clear
End Sub
Public Sub TreeView_AddParent(TV As TreeView, ByVal ParentName As String, ByVal ParentText As String, Optional ByVal Image As String = "project")
    On Error Resume Next
    Set xNode = TV.Nodes.Add(, , ParentName, ParentText, Image, Image)
    xNode.Key = xNode.Fullpath
    TV.Nodes.Item(1).Bold = True
    TV.Nodes.Item(1).Expanded = True
    Err.Clear
End Sub
Public Function TreeView_NodeIndex(TV As TreeView, ByVal ChildName As String) As Long
    On Error Resume Next
    Dim x As Long
    Dim x_tot As Long
    x_tot = TV.Nodes.Count
    For x = 1 To x_tot
        If TV.Nodes.Item(x).Text = ChildName Then
            TreeView_NodeIndex = x
        End If
        Err.Clear
    Next
    Err.Clear
End Function
Public Sub TreeView_DeleteNode(TV As TreeView, ByVal NodeText As String)
    On Error Resume Next
    TV.Nodes.Remove TreeView_NodeIndex(TV, NodeText)
    Err.Clear
End Sub
Function RichToHTML(rtbRichTextBox As RichTextLib.RichTextBox, Optional lngStartPosition As Long = 0, Optional lngEndPosition As Long = 0) As String
    On Error Resume Next
    '**********************************************************
    '*            Draw Percent by Joseph Huntley              *
    '*               joseph_huntley@email.com                 *
    '*                http://joseph.vr9.com                   *
    '**********************************************************
    '*   You may use this code freely as long as credit is    *
    '* given to the author, and the header remains intact.    *
    '**********************************************************
    '--------------------- The Arguments -----------------------
    'rtbRichTextBox     - The rich textbox control to convert.
    'lngStartPosition   - The character position to start from.
    'lngEndPosition     - The character position to end at.
    '-----------------------------------------------------------
    'Returns:     The rich text converted to HTML.
    'Description: Converts rich text to HTML.
    Dim blnBold As Boolean
    Dim blnUnderline As Boolean
    Dim blnStrikeThru As Boolean
    Dim blnItalic As Boolean
    Dim strLastFont As String
    Dim lngLastFontColor As Long
    Dim strHTML As String
    Dim lngCurText As Long
    Dim strHex As String
    Dim intLastAlignment As Integer
    Const AlignLeft = 0, AlignRight = 1, AlignCenter = 2
    'check for lngStartPosition ad lngEndPosition
    If IsMissing(lngStartPosition&) Then lngStartPosition& = 0
    If IsMissing(lngEndPosition&) Then lngEndPosition& = Len(rtbRichTextBox.Text)
    lngLastFontColor& = -1 'no color
    For lngCurText& = lngStartPosition& To lngEndPosition&
        rtbRichTextBox.SelStart = lngCurText&
        rtbRichTextBox.SelLength = 1
        If intLastAlignment% <> rtbRichTextBox.SelAlignment Then
            intLastAlignment% = rtbRichTextBox.SelAlignment
            Select Case rtbRichTextBox.SelAlignment
            Case AlignLeft: strHTML$ = strHTML$ & "<p align=left>"
            Case AlignRight: strHTML$ = strHTML$ & "<p align=right>"
            Case AlignCenter: strHTML$ = strHTML$ & "<p align=center>"
            End Select
        End If
        If blnBold <> rtbRichTextBox.SelBold Then
            If rtbRichTextBox.SelBold = True Then
                strHTML$ = strHTML$ & "<b>"
            Else
                strHTML$ = strHTML$ & "</b>"
            End If
            blnBold = rtbRichTextBox.SelBold
        End If
        If blnUnderline <> rtbRichTextBox.SelUnderline Then
            If rtbRichTextBox.SelUnderline = True Then
                strHTML$ = strHTML$ & "<u>"
            Else
                strHTML$ = strHTML$ & "</u>"
            End If
            blnUnderline = rtbRichTextBox.SelUnderline
        End If
        If blnItalic <> rtbRichTextBox.SelItalic Then
            If rtbRichTextBox.SelItalic = True Then
                strHTML$ = strHTML$ & "<i>"
            Else
                strHTML$ = strHTML$ & "</i>"
            End If
            blnItalic = rtbRichTextBox.SelItalic
        End If
        If blnStrikeThru <> rtbRichTextBox.SelStrikeThru Then
            If rtbRichTextBox.SelStrikeThru = True Then
                strHTML$ = strHTML$ & "<s>"
            Else
                strHTML$ = strHTML$ & "</s>"
            End If
            blnStrikeThru = rtbRichTextBox.SelStrikeThru
        End If
        If strLastFont$ <> rtbRichTextBox.SelFontName Then
            strLastFont$ = rtbRichTextBox.SelFontName
            strHTML$ = strHTML$ + "<font face=""" & strLastFont$ & """>"
        End If
        If lngLastFontColor& <> rtbRichTextBox.SelColor Then
            lngLastFontColor& = rtbRichTextBox.SelColor
            ''Get hexidecimal value of color
            strHex$ = Hex$(rtbRichTextBox.SelColor)
            strHex$ = String$(6 - Len(strHex$), "0") & strHex$
            strHex$ = Right$(strHex$, 2) & Mid$(strHex$, 3, 2) & Left$(strHex$, 2)
            strHTML$ = strHTML$ + "<font color=#" & strHex$ & ">"
        End If
        strHTML$ = strHTML$ + rtbRichTextBox.SelText
        Err.Clear
    Next
    RichToHTML = strHTML$
    Err.Clear
End Function
Function RGBtoHEX(lngColor As Long) As String
    On Error Resume Next
    Dim strHex As String
    'get hexidecimal value
    strHex = Hex$(lngColor)
    'fill in
    strHex = String$(6 - Len(strHex), "0") & strHex
    'swap first and third hex values.
    strHex = Right$(strHex, 2) & Mid$(strHex, 3, 2) & Left$(strHex, 2)
    RGBtoHEX = strHex
    Err.Clear
End Function
Public Function BoxConversion(Box As String) As String
    On Error Resume Next
    Dim FindBox() As String
    Dim Data As String
    FindBox = Split(Box, "<boxl>")
    Data = FindBox(1)
    FindBox = Split(Data, "</box>")
    Data = FindBox(0)
    If Data > "" Then
        BoxConversion = "\pard\par\box\ql " & Data & "\par\pard "
        Err.Clear
        Exit Function
    End If
    FindBox = Split(Box, "<boxc>")
    Data = FindBox(1)
    FindBox = Split(Data, "</box>")
    Data = FindBox(0)
    If Data > "" Then
        BoxConversion = "\pard\par\box\qc " & Data & "\par\pard "
        Err.Clear
        Exit Function
    End If
    FindBox = Split(Box, "<boxr>")
    Data = FindBox(1)
    FindBox = Split(Data, "</box>")
    Data = FindBox(0)
    If Data > "" Then
        BoxConversion = "\pard\par\box\qr " & Data & "\par\pard "
        Err.Clear
        Exit Function
    End If
    Err.Clear
End Function
Public Function EXEConversion(exe As String) As String
    On Error Resume Next
    Dim FindContents() As String
    Dim FindContents2() As String
    Dim FindJumpName() As String
    Dim ContentPage As String
    Dim PageName As String
    Dim Data As String
    'Find the ContentPage Bracket
    FindContents = Split(exe, "<exe=")
    'Set the Conetnts Page to the Split/Non Complet ContentPage
    ContentPage = FindContents(1)
    'Split it again to Complete the Content Page Name
    FindContents2 = Split(ContentPage, ">")
    'Set ContentPage to the Name on the Content Page
    ContentPage = FindContents2(0)
    'Set the second part of the split to the page name
    PageName = FindContents2(1)
    'Split it again the get the page name
    FindJumpName = Split(PageName, "</exe")
    'Complete the Page Name
    PageName = FindJumpName(0)
    Data = "{\uldb " & PageName & "}" & "{\v !ExecFile(" & ContentPage & ",)}"
    EXEConversion = Data
    Err.Clear
End Function
Public Function URLConversion(url As String) As String
    On Error Resume Next
    Dim FindContents() As String
    Dim FindContents2() As String
    Dim FindJumpName() As String
    Dim ContentPage As String
    Dim PageName As String
    Dim Data As String
    'Find the ContentPage Bracket
    FindContents = Split(url, "<url=")
    'Set the Conetnts Page to the Split/Non Complet ContentPage
    ContentPage = FindContents(1)
    'Split it again to Complete the Content Page Name
    FindContents2 = Split(ContentPage, ">")
    'Set ContentPage to the Name on the Content Page
    ContentPage = FindContents2(0)
    'Set the second part of the split to the page name
    PageName = FindContents2(1)
    'Split it again the get the page name
    FindJumpName = Split(PageName, "</url")
    'Complete the Page Name
    PageName = FindJumpName(0)
    Data = "{\uldb " & PageName & "}" & "{\v !ExecFile(" & ContentPage & ",)}"
    URLConversion = Data
    Err.Clear
End Function
Public Function TargetConversion(Target As String) As String
    On Error Resume Next
    Dim FindTarget() As String
    Dim FindTarget1 As String
    Dim FindTarget2() As String
    Dim Data As String
    'Find the start of the target
    FindTarget = Split(Target, "<target>")
    FindTarget1 = FindTarget(1)
    FindTarget2 = Split(FindTarget1, "</target>")
    'Get the Target Keyword
    FindTarget1 = FindTarget2(0)
    'Compile the String
    Data = "K{\footnote " & FindTarget1 & "}" & "#{\footnote " & FindTarget1 & "}"
    TargetConversion = Data
    Err.Clear
End Function
Public Function JumpConversion(Jump As String) As String
    On Error Resume Next
    Dim FindContents() As String
    Dim FindContents2() As String
    Dim FindJumpName() As String
    Dim ContentPage As String
    Dim PageName As String
    Dim Data As String
    'Find the ContentPage Bracket
    FindContents = Split(Jump, "<jump=")
    'Set the Conetnts Page to the Split/Non Complet ContentPage
    ContentPage = FindContents(1)
    'Split it again to Complete the Content Page Name
    FindContents2 = Split(ContentPage, ">")
    'Set ContentPage to the Name on the Content Page
    ContentPage = FindContents2(0)
    'Set the second part of the split to the page name
    PageName = FindContents2(1)
    'Split it again the get the page name
    FindJumpName = Split(PageName, "</jump")
    'Complete the Page Name
    PageName = FindJumpName(0)
    Data = "{\uldb " & PageName & "}" & "{\v " & ContentPage & "}"
    JumpConversion = Data
    Err.Clear
End Function
Public Function JumpContext(ByVal Jump As String, ByVal JumpSelRTF As String, ByVal Topic As String, Optional JumpType As String = "j", Optional NoUnderLine As Boolean = False, Optional NotGreen As Boolean = False, Optional sWindow As String = "main") As String
    On Error Resume Next
    Dim strValue As String
    If NoUnderLine = True Then
        Topic = "%" & Topic
    End If
    If NotGreen = True Then
        Topic = "*" & Topic
    End If
    Select Case LCase$(JumpType)
    Case "j"
        Topic = Topic & ">" & sWindow
        strValue = Replace$(JumpSelRTF, " " & Jump, "{\uldb " & Jump & "}{\v " & Topic & "}")
    Case "p"
        Topic = Topic & ">" & sWindow
        strValue = Replace$(JumpSelRTF, " " & Jump, "{\ul " & Jump & "}{\v " & Topic & "}")
    Case "exe", "url"
        Jump = Replace$(Jump, "\", "\\")
        strValue = Replace$(JumpSelRTF, " " & Jump, "{\uldb " & Jump & "}{\v !ExecFile(" & Jump & ")}")
    Case "file"
        Jump = Replace$(Jump, "\", "\\")
        strValue = Replace$(JumpSelRTF, " " & Jump, "{\ul " & Jump & "}{\v " & Topic & "}")
    End Select
    JumpContext = strValue
    Err.Clear
End Function
Public Function PopConversion(pop As String) As String
    On Error Resume Next
    Dim FindContents() As String
    Dim FindContents2() As String
    Dim FindJumpName() As String
    Dim ContentPage As String
    Dim PageName As String
    Dim Data As String
    'Find the ContentPage Bracket
    FindContents = Split(pop, "<pop=")
    'Set the Conetnts Page to the Split/Non Complet ContentPage
    ContentPage = FindContents(1)
    'Split it again to Complete the Content Page Name
    FindContents2 = Split(ContentPage, ">")
    'Set ContentPage to the Name on the Content Page
    ContentPage = FindContents2(0)
    'Set the second part of the split to the page name
    PageName = FindContents2(1)
    'Split it again the get the page name
    FindJumpName = Split(PageName, "</pop")
    'Complete the Page Name
    PageName = FindJumpName(0)
    Data = "{\ul " & PageName & "}" & "{\v " & ContentPage & "}"
    PopConversion = Data
    Err.Clear
End Function
Public Function Hex2Dec(HexValue As String) As Integer
    On Error Resume Next
    Dim lsChar As String
    Dim liHBit As Integer
    Dim liLBit As Integer
    ' H Byte
    lsChar = Mid$(HexValue, 1, 1)
    Select Case Asc(lsChar)
    Case 48 To 57 '0..9
        liHBit = Val(lsChar) * 16
    Case 65 To 70 'A..F
        liHBit = (((65 - Asc(lsChar)) * -1) + 10) * 16
    End Select
    ' L Byte
    lsChar = Mid$(HexValue, 2, 1)
    Select Case Asc(lsChar)
    Case 48 To 57 '0..9
        liLBit = Val(lsChar)
    Case 65 To 70 'A..F
        liLBit = (65 - Asc(lsChar)) * -1 + 10
    End Select
    Hex2Dec = liHBit + liLBit
    Err.Clear
End Function
Public Sub Compile_Hpj(MeForm As Form, progBar As ProgressBar)
    On Error Resume Next
    Dim nTot As Long
    Dim nCnt As Long
    Dim nPth As String
    Dim db As DAO.Database
    Dim tb As DAO.Recordset
    Dim tbProject As DAO.Recordset
    Dim tContext As String
    Dim tNumber As String
    Dim CompiledHPJ As String
    Dim strBaggage As String
    If boolFileExists(sProjHPJ) = True Then Kill sProjHPJ
    'strBaggage = Compile_Baggage(MeForm, progBar)
    StatusMessage MeForm, "Compiling the project file, please wait..."
    Set db = DAO.OpenDatabase(sProjDb)
    Set tb = db.OpenRecordset("Properties")
    tb.Index = "FullPath"
    Set tbProject = db.OpenRecordset("select * from [" & sProject & "] order by sequence;")
    tbProject.MoveLast
    nTot = tbProject.RecordCount
    tbProject.MoveFirst
    CompiledHPJ = CompiledHPJ & "[OPTIONS]" & vbCrLf
    CompiledHPJ = CompiledHPJ & "CNT=" & StringGetFileToken(sProjCnt, "f") & vbCrLf
    CompiledHPJ = CompiledHPJ & "ERRORLOG=" & StringGetFileToken(sProjLog, "f") & vbCrLf
    CompiledHPJ = CompiledHPJ & "NOTES=yes" & vbCrLf
    CompiledHPJ = CompiledHPJ & "REPORT=yes" & vbCrLf
    CompiledHPJ = CompiledHPJ & "TITLE=" & Title & vbCrLf
    CompiledHPJ = CompiledHPJ & "FTS=29" & vbCrLf
    CompiledHPJ = CompiledHPJ & "COMPRESS=60 Hall Zeck" & vbCrLf
    CompiledHPJ = CompiledHPJ & "OLDKEYPHRASE=NO" & vbCrLf
    CompiledHPJ = CompiledHPJ & "COPYRIGHT=" & Author & vbCrLf
    CompiledHPJ = CompiledHPJ & "CITATION=" & Author & vbCrLf
    'CompiledHPJ = CompiledHPJ & "DEFFONT=" & FontText & "," & FontTextSize & ",ANSI" & vbCrLf
    CompiledHPJ = CompiledHPJ & "LCID=0x1c09 0x6 0x0 ; English (South Africa)" & vbCrLf
    CompiledHPJ = CompiledHPJ & "HLP=" & StringGetFileToken(sProjHLP, "f") & vbCrLf & vbCrLf
    CompiledHPJ = CompiledHPJ & "[BAGGAGE]" & vbCrLf
    'CompiledHPJ = CompiledHPJ & strBaggage & vbCrLf & vbCrLf
    CompiledHPJ = CompiledHPJ & "[FILES]" & vbCrLf
    CompiledHPJ = CompiledHPJ & StringGetFileToken(sProjRTF, "f") & vbCrLf & vbCrLf
    CompiledHPJ = CompiledHPJ & "[CONFIG]" & vbCrLf
    CompiledHPJ = CompiledHPJ & "BrowseButtons()" & vbCrLf & vbCrLf
    CompiledHPJ = CompiledHPJ & "[WINDOWS]" & vbCrLf
    CompiledHPJ = CompiledHPJ & "main=" & """" & Title & """,(20,71,600,783),60677,(r" & TextBackColor & "),(r" & HeadlineBackColor & "),f3" & vbCrLf & vbCrLf
    CompiledHPJ = CompiledHPJ & "[MAP]" & vbCrLf
    Call ProgBarInit(MeForm, progBar, nTot)
    For nCnt = 1 To nTot
        Call UpdateProgress(MeForm, nCnt, progBar, "Compiling the hpj file")
        nPth = StringProperCase(StringRemNull(tbProject!Fullpath))
        tb.Seek "=", nPth
        Select Case tb.NoMatch
        Case False
            tContext = StringRemNull(tb!Context)
            tNumber = StringRemNull(tb!Number)
            CompiledHPJ = CompiledHPJ & "#define " & tContext & " " & tNumber & vbCrLf
        End Select
        tbProject.MoveNext
        DoEvents
        Err.Clear
    Next
    ProgBarClose MeForm, progBar
    tb.Close
    db.Close
    Set tb = Nothing
    Set db = Nothing
    CompiledHPJ = CompiledHPJ & "" & vbCrLf
    FileUpdate sProjHPJ, CompiledHPJ
    Err.Clear
End Sub
Public Function Compile_Baggage(MeForm As Form, progBar As ProgressBar) As String
    On Error Resume Next
    Dim db As DAO.Database
    Dim tb As DAO.Recordset
    Dim nCnt As Long
    Dim nTot As Long
    Dim nFile As String
    Dim nNew As String
    nNew = ""
    Set db = DAO.OpenDatabase(sProjDb)
    Set tb = db.OpenRecordset("DataFiles")
    tb.MoveFirst
    nTot = tb.RecordCount
    Call ProgBarInit(MeForm, progBar, nTot)
    For nCnt = 1 To nTot
        Call UpdateProgress(MeForm, nCnt, progBar, "Updating baggage files...")
        nFile = StringRemNull(tb!FileNames)
        nNew = nNew & nFile & Chr$(253)
        tb.MoveNext
        Err.Clear
    Next
    tb.Close
    db.Close
    Set tb = Nothing
    Set db = Nothing
    ProgBarClose MeForm, progBar
    Compile_Baggage = Replace$(nNew, Chr$(253), vbNewLine)
    Err.Clear
End Function
Public Sub Compile_Rtf(MeForm As Form, progBar As ProgressBar)
    On Error Resume Next
    Dim db As DAO.Database
    Dim tb As DAO.Recordset
    Dim tbProject As DAO.Recordset
    Dim nTot As Long
    Dim nCnt As Long
    Dim pTitle As String
    Dim pContext As String
    Dim pNumber As String
    Dim pBrowse As String
    Dim pKeyWords As String
    Dim pContents As String
    Dim pPopUp As String
    Dim pMacros As String
    Dim MainText As String
    Dim arCode() As String
    Dim iLineCount As Long
    Dim Data As Long
    Dim Data2 As String
    Dim tPath As String
    Dim pFootNotes As String
    Dim CompiledRTF As String
    Dim iLineCount_Tot As Long
    Dim iLineCount_Cnt As Long
    Dim TopicPath As String
    Dim sParent As String
    If boolFileExists(sProjRTF) = True Then Kill sProjRTF
    StatusMessage MeForm, "Selecting topics, please wait..."
    Set db = DAO.OpenDatabase(sProjDb)
    Set tb = db.OpenRecordset("Properties")
    Set tbProject = db.OpenRecordset("select * from [" & sProject & "] order by sequence;")
    tbProject.MoveLast
    nTot = tbProject.RecordCount
    tbProject.MoveFirst
    tb.Index = "FullPath"
    Call ProgBarInit(MeForm, progBar, nTot)
    If FontHeadline = "" Then FontHeadline = "Tahoma"
    If FontText = "" Then FontText = "Tahoma"
    CompiledRTF = "{\rtf1\ansi\deff0{\fonttbl{\f0\fnil\fcharset0 " & FontHeadline & ";}{\f1\fnil\fcharset0 " & FontText & ";}}"
    'Compile the Pages
    For nCnt = 1 To nTot
        Call UpdateProgress(MeForm, nCnt, progBar, "Compiling the rtf file")
        tPath = StringProperCase(StringRemNull(tbProject!Fullpath))
        sParent = StringRemNull(tbProject!Parent)
        If Val(sParent) = 0 Then GoTo NextItem
        tb.Seek "=", tPath
        If tb.NoMatch = True Then
            GoTo NextItem
        End If
        pTitle = StringProperCase(StringRemNull(tb!Title))
        pContext = StringRemNull(tb!Context)
        pNumber = StringRemNull(tb!Number)
        pBrowse = StringRemNull(tb!browse)
        pKeyWords = StringRemNull(tb!Keywords)
        pContents = StringRemNull(tb!Contents)
        pFootNotes = StringRemNull(tb!Footnotes)
        pMacros = StringRemNull(tb!Macros)
        pPopUp = StringRemNull(tb!popup)
        pKeyWords = Replace$(pKeyWords, FM, ";")
        pFootNotes = Replace$(pFootNotes, FM, ";")
        If Len(pFootNotes) > 0 Then
            CompiledRTF = CompiledRTF & "A{\footnote " & MvRemoveDuplicates(pTitle & ";" & pFootNotes, ";") & "}"
        End If
        If Len(pMacros) > 0 Then
            CompiledRTF = CompiledRTF & "!{\footnote " & pMacros & "}"
        End If
        If pBrowse = "1" Then
            CompiledRTF = CompiledRTF & "+{\footnote auto}"
        End If
        CompiledRTF = CompiledRTF & "K{\footnote " & MvRemoveDuplicates(pTitle & ";" & pKeyWords, ";") & "}"
        If pPopUp <> "1" Then
            CompiledRTF = CompiledRTF & "${\footnote " & pTitle & "}"
        End If
        CompiledRTF = CompiledRTF & "#{\footnote " & pContext & "}"
        TopicPath = MvFromMv(tPath, 2, , "\")
        TopicPath = Replace$(TopicPath, "\", "\\")
        If FontHeadlineBold = "1" Then
            CompiledRTF = CompiledRTF & "{\keepn\cb3\cf1\b\fs" & FontHeadlineSize * 2 & "\f0 " & TopicPath & "\fs" & FontTextSize * 2 & "\b0\cb2\par\pard\plain"
            CompiledRTF = CompiledRTF & "\line\cf0\cb2\fs20\f1}"
        Else
            CompiledRTF = CompiledRTF & "{\keepn\cb3\cf1\fs" & FontHeadlineSize * 2 & "\f0 " & TopicPath & "\fs" & FontTextSize * 2 & "\cb2\par\pard\plain"
            CompiledRTF = CompiledRTF & "\line\cf0\cb2\fs20\f1}"
        End If
        MainText = pContents
        If Len(MainText) = 0 Then GoTo NextTopic
        '=================Replace Color Codes======================
        'Do a Color Replacement and Line return replacement
        MainText = Replace$(MainText, "<color=1>", "{\cf4 ")
        MainText = Replace$(MainText, "<color=2>", "{\cf5 ") 'LIGHT PINK
        MainText = Replace$(MainText, "<color=3>", "{\cf6 ") 'PURPLE
        MainText = Replace$(MainText, "<color=4>", "{\cf7 ") 'RED
        MainText = Replace$(MainText, "<color=5>", "{\cf8 ") 'TEAL
        MainText = Replace$(MainText, "<color=6>", "{\cf9 ") 'DARK BLUE
        MainText = Replace$(MainText, "<color=7>", "{\cf10 ") 'GREEN
        MainText = Replace$(MainText, "<color=8>", "{\cf11 ") 'GOLD
        MainText = Replace$(MainText, "<color=9>", "{\cf12 ") 'GREY
        MainText = Replace$(MainText, "<color=10>", "{\cf13 ") 'BLUE
        MainText = Replace$(MainText, "<color=11>", "{\cf14 ") 'LIGHT GREEN
        MainText = Replace$(MainText, "<color=12>", "{\cf15 ") 'PINK
        MainText = Replace$(MainText, "<color=13>", "{\cf16 ") 'WHITE
        MainText = Replace$(MainText, "<color=14>", "{\cf17 ") 'LIGHT PURPLE
        MainText = Replace$(MainText, "<color=15>", "{\cf18 ") 'YELLOW
        MainText = Replace$(MainText, "<color=16>", "{\cf19 ") 'DARK GREY
        MainText = Replace$(MainText, "<\color>", "}")
        MainText = Replace$(MainText, "</color>", "}")
        '=================End Replace Jump Codes======================
        '=================Replace Jump Codes======================
        If InStr(1, MainText, "<jump=", vbTextCompare) > 0 Then
            arCode = Split(MainText, vbCr)
            'Loop through th lines 1 by 1
            iLineCount_Tot = UBound(arCode)
            iLineCount_Cnt = LBound(arCode)
            For iLineCount = iLineCount_Cnt To iLineCount_Tot
                DoEvents
                'If there is Text in the line then Look through it
                If Len(Trim$(arCode(iLineCount))) > 0 Then
                    'If <jump= is found the change it
                    If InStr(1, Trim$(arCode(iLineCount)), "<jump=") Then
                        'Set data as the number of letters into the text the the word <jump= starts
                        Data = InStr(1, Trim$(arCode(iLineCount)), "<jump=")
                        'Set Data2 to Equal the Entire line of Text
                        Data2 = Trim$(arCode(iLineCount))
                        'Replace the Text with what we want it to be
                        MainText = Replace$(MainText, Mid$(Data2, Data), JumpConversion(Mid$(Data2, Data)))
                    End If
                End If
                Err.Clear
            Next
        End If
        '=================End Replace Jump Codes======================
        '=================Replace Target Codes======================
        'Split the Text in to lines
        If InStr(1, MainText, "<target>", vbTextCompare) > 0 Then
            arCode = Split(MainText, vbCr)
            'Loop through th lines 1 by 1
            iLineCount_Tot = UBound(arCode)
            iLineCount_Cnt = LBound(arCode)
            For iLineCount = iLineCount_Cnt To iLineCount_Tot
                DoEvents
                'If there is Text in the line then Look through it
                If Len(Trim$(arCode(iLineCount))) > 0 Then
                    'If <jump= is found the change it
                    If InStr(1, Trim$(arCode(iLineCount)), "<target>") Then
                        'Set data as the number of letters into the text the the word <jump= starts
                        Data = InStr(1, Trim$(arCode(iLineCount)), "<target>")
                        'Set Data2 to Equal the Entire line of Text
                        Data2 = Trim$(arCode(iLineCount))
                        'Replace the Text with what we want it to be
                        MainText = Replace$(MainText, Mid$(Data2, Data), TargetConversion(Mid$(Data2, Data)))
                    End If
                End If
                Err.Clear
            Next
        End If
        '=================End Replace Target Codes======================
        '=================Replace URL Codes======================
        'Split the Text in to lines
        If InStr(1, MainText, "<url=", vbTextCompare) > 0 Then
            arCode = Split(MainText, vbCr)
            'Loop through th lines 1 by 1
            iLineCount_Tot = UBound(arCode)
            iLineCount_Cnt = LBound(arCode)
            For iLineCount = iLineCount_Cnt To iLineCount_Tot
                DoEvents
                'If there is Text in the line then Look through it
                If Len(Trim$(arCode(iLineCount))) > 0 Then
                    'If <jump= is found the change it
                    If InStr(1, Trim$(arCode(iLineCount)), "<url=") Then
                        'Set data as the number of letters into the text the the word <jump= starts
                        Data = InStr(1, Trim$(arCode(iLineCount)), "<url=")
                        'Set Data2 to Equal the Entire line of Text
                        Data2 = Trim$(arCode(iLineCount))
                        'Replace the Text with what we want it to be
                        MainText = Replace$(MainText, Mid$(Data2, Data), URLConversion(Mid$(Data2, Data)))
                    End If
                End If
                Err.Clear
            Next
        End If
        '=================End Replace URL Codes======================
        '=================Replace Pop Codes======================
        'Split the Text in to lines
        If InStr(1, MainText, "<pop=", vbTextCompare) > 0 Then
            arCode = Split(MainText, vbCr)
            'Loop through th lines 1 by 1
            iLineCount_Tot = UBound(arCode)
            iLineCount_Cnt = LBound(arCode)
            For iLineCount = iLineCount_Cnt To iLineCount_Tot
                DoEvents
                'If there is Text in the line then Look through it
                If Len(Trim$(arCode(iLineCount))) > 0 Then
                    'If <jump= is found the change it
                    If InStr(1, Trim$(arCode(iLineCount)), "<pop=") Then
                        'Set data as the number of letters into the text the the word <jump= starts
                        Data = InStr(1, Trim$(arCode(iLineCount)), "<pop=")
                        'Set Data2 to Equal the Entire line of Text
                        Data2 = Trim$(arCode(iLineCount))
                        'Replace the Text with what we want it to be
                        MainText = Replace$(MainText, Mid$(Data2, Data), PopConversion(Mid$(Data2, Data)))
                    End If
                End If
                Err.Clear
            Next
        End If
        '=================End Replace Pop Codes======================
        '=================Replace Box Codes======================
        'Split the Text in to lines
        If InStr(1, MainText, "<box", vbTextCompare) > 0 Then
            arCode = Split(MainText, vbCr)
            'Loop through th lines 1 by 1
            iLineCount_Tot = UBound(arCode)
            iLineCount_Cnt = LBound(arCode)
            For iLineCount = iLineCount_Cnt To iLineCount_Tot
                DoEvents
                'If there is Text in the line then Look through it
                If Len(Trim$(arCode(iLineCount))) > 0 Then
                    'If <jump= is found the change it
                    If InStr(1, Trim$(arCode(iLineCount)), "<box") Then
                        'Set data as the number of letters into the text the the word <jump= starts
                        Data = InStr(1, Trim$(arCode(iLineCount)), "<box")
                        'Set Data2 to Equal the Entire line of Text
                        Data2 = Trim$(arCode(iLineCount))
                        'Replace the Text with what we want it to be
                        MainText = Replace$(MainText, Mid$(Data2, Data), BoxConversion(Mid$(Data2, Data)))
                    End If
                End If
                Err.Clear
            Next
        End If
        '=================End Replace Box Codes======================
        '=================Replace EXE Codes======================
        'Split the Text in to lines
        If InStr(1, MainText, "<exe=", vbTextCompare) > 0 Then
            arCode = Split(MainText, vbCr)
            'Loop through th lines 1 by 1
            iLineCount_Tot = UBound(arCode)
            iLineCount_Cnt = LBound(arCode)
            For iLineCount = iLineCount_Cnt To iLineCount_Tot
                DoEvents
                'If there is Text in the line then Look through it
                If Len(Trim$(arCode(iLineCount))) > 0 Then
                    'If <jump= is found the change it
                    If InStr(1, Trim$(arCode(iLineCount)), "<exe=") Then
                        'Set data as the number of letters into the text the the word <jump= starts
                        Data = InStr(1, Trim$(arCode(iLineCount)), "<exe=")
                        'Set Data2 to Equal the Entire line of Text
                        Data2 = Trim$(arCode(iLineCount))
                        'Replace the Text with what we want it to be
                        MainText = Replace$(MainText, Mid$(Data2, Data), EXEConversion(Mid$(Data2, Data)))
                    End If
                End If
                Err.Clear
            Next
        End If
        '=================End Replace EXE Codes======================
        '=================Replace Format Codes======================
        MainText = Replace$(MainText, "<b>", "{\b ") 'Bold
        MainText = Replace$(MainText, "<i>", "{\i ") 'Italic
        MainText = Replace$(MainText, "<u>", "{\ul ") 'UNderline
        MainText = Replace$(MainText, "<l>", "\pard\par\ql ") 'Align L
        MainText = Replace$(MainText, "<c>", "\pard\par\qc ") 'Align Center
        MainText = Replace$(MainText, "<r>", "\pard\par\qr ") 'Align Right
        MainText = Replace$(MainText, "</b>", "}") 'Align Right
        MainText = Replace$(MainText, "</i>", "}") 'Align Right
        MainText = Replace$(MainText, "</u>", "}") 'Align Right
        MainText = Replace$(MainText, "</l>", "\par\pard") 'Align Right
        MainText = Replace$(MainText, "</c>", "\par\pard") 'Align Right
        MainText = Replace$(MainText, "</r", "\par\pard") 'Align Right
NextTopic:
        '=================End Replace Format Codes======================
        CompiledRTF = CompiledRTF & MainText & "\line"
        CompiledRTF = CompiledRTF & "\plain\par\page"
NextItem:
        tbProject.MoveNext
        DoEvents
        Err.Clear
    Next
    tb.Close
    db.Close
    ProgBarClose MeForm, progBar
    Set tb = Nothing
    Set db = Nothing
    'Close the Compiled RTF
    CompiledRTF = CompiledRTF & "}"
    FileUpdate sProjRTF, CompiledRTF, "w"
    Err.Clear
End Sub
Public Sub Compile_CNT(MeForm As Form, progBar As ProgressBar)
    On Error Resume Next
    Dim nTot As Long
    Dim nCnt As Long
    Dim nPth As String
    Dim CompiledCNT As String
    Dim db As DAO.Database
    Dim tb As DAO.Recordset
    Dim tbProject As DAO.Recordset
    Dim tContext As String
    Dim tTitle As String
    Dim tImage As String
    Dim pCount As Long
    Dim sParent As String
    Dim strLine As String
    Dim strContents As String
    If boolFileExists(sProjCnt) = True Then Kill sProjCnt
    StatusMessage MeForm, "Selecting topics, please be patient..."
    Set db = DAO.OpenDatabase(sProjDb)
    Set tb = db.OpenRecordset("Properties")
    tb.Index = "FullPath"
    Set tbProject = db.OpenRecordset("select * from [" & sProject & "] order by sequence;")
    tbProject.MoveLast
    nTot = tbProject.RecordCount
    tbProject.MoveFirst
    CompiledCNT = CompiledCNT & ":Base " & StringGetFileToken(sProjHLP, "f") & vbCrLf
    CompiledCNT = CompiledCNT & ":Title " & sProject & vbCrLf
    CompiledCNT = CompiledCNT & ":Index " & sProject & "=" & sProject & ".hlp ;(This is necessary when using KLinks - because of a bug in the help" & vbCrLf
    CompiledCNT = CompiledCNT & ":Link " & sProject & ".hlp        ; compiler. Without this only the first keyword would be looked up!)" & vbCrLf
    Call ProgBarInit(MeForm, progBar, nTot)
    For nCnt = 1 To nTot
        Call UpdateProgress(MeForm, nCnt, progBar, "Compiling the cnt file")
        nPth = StringProperCase(StringRemNull(tbProject!Fullpath))
        tImage = StringRemNull(tbProject!Image)
        tTitle = StringProperCase(StringRemNull(tbProject!Text))
        sParent = StringRemNull(tbProject!Parent)
        tContext = Context_Validate(MvFromMv(nPth, 2, , "\"))
        If Val(sParent) = 0 Then GoTo NextTopic
        pCount = MvCount(nPth, "\") - 1
        Select Case tImage
        Case "book"
            strLine = CStr(pCount) & " " & tTitle & vbCrLf
            CompiledCNT = CompiledCNT & strLine
            tb.Seek "=", nPth
            If tb.NoMatch = False Then
                strContents = Trim$(tb!Contents.Value & "")
                If Len(strContents) > 0 Then
                    strLine = CStr(pCount + 1) & " An Introduction To " & tTitle & "=" & tContext & "@" & StringGetFileToken(sProjHLP, "f") & vbCrLf
                    CompiledCNT = CompiledCNT & strLine
                End If
            End If
        Case "leaf"
            strLine = CStr(pCount) & " " & tTitle & "=" & tContext & "@" & StringGetFileToken(sProjHLP, "f") & vbCrLf
            CompiledCNT = CompiledCNT & strLine
        End Select
NextTopic:
        tbProject.MoveNext
        DoEvents
        Err.Clear
    Next
    CompiledCNT = StringRemoveDelim(CompiledCNT, vbCrLf)
    ProgBarClose MeForm, progBar
    tb.Close
    db.Close
    Set tb = Nothing
    Set db = Nothing
    FileUpdate sProjCnt, CompiledCNT, "w"
    pCount = TOC_Errors(MeForm, progBar)
    Do Until pCount = 0
        TOC_Fix MeForm, progBar
        pCount = TOC_Errors(MeForm, progBar)
        Err.Clear
    Loop
    Err.Clear
End Sub

Public Sub Compile_HTML(MeForm As Form, progBar As ProgressBar, FileT As RichTextBox)
    On Error Resume Next
    Dim nTot As Long
    Dim nCnt As Long
    Dim db As DAO.Database
    Dim tb As DAO.Recordset
    Dim tContext As String
    Dim tContents As String
    
    Kill sProjHTML & "\*.*"
    StatusMessage MeForm, "Selecting topics, please be patient..."
    Set db = DAO.OpenDatabase(sProjDb)
    Set tb = db.OpenRecordset("Properties")
    nTot = tb.RecordCount
    Call ProgBarInit(MeForm, progBar, nTot)
    For nCnt = 1 To nTot
        Call UpdateProgress(MeForm, nCnt, progBar, "Compiling the html files, please wait")
        tContext = tb!Context & ""
        tContents = tb!Contents & ""
        FileT.TextRTF = tContents
        Word97Do wSaveHTML, FileT, tContext
        tb.MoveNext
        DoEvents
        Err.Clear
    Next
    tb.Close
    db.Close
    Set tb = Nothing
    Set db = Nothing
    Err.Clear
End Sub
Public Function DelimCount(ByVal StringMv As String, Optional ByVal Delim As String = "") As Long
    On Error Resume Next
    Dim rsTot As Long
    Dim rsCnt As Long
    Dim rsStr As String
    Dim dmCount As Long
    If Len(Delim) = 0 Then
        Delim = VM
    End If
    dmCount = 0
    rsTot = Len(StringMv)
    For rsCnt = 1 To rsTot
        rsStr = Mid$(StringMv, rsCnt, Len(Delim))
        Select Case LCase$(rsStr)
        Case LCase$(Delim)
            dmCount = dmCount + 1
        End Select
        Err.Clear
    Next
    DelimCount = dmCount
    Err.Clear
End Function
Sub Compile_Project(Location As String)
    On Error Resume Next
    Dim ReturnValue As Long
    '/A  Specifies that additional information is to be added to the Help file. For version 4.0, the additional information is the topic ID of each topic and the source file the topic appears in. If the Help Author command on the File menu in Help Workshop is checked, you can display this information by clicking a topic using your right mouse button, and then clicking Topic Information.
    '/C  Starts compiling the specified project (.hpj) file or Help makefile (.hmk).
    '/E  Quits the Hcw.exe program after compiling the specified file. Use this switch in conjunction with the /C or /M switch.
    '/M  Minimizes the Hcw.exe program while compiling the specified project file.
    '/N  Turns off compression when compiling the specified file, no matter what value is specified for the COMPRESS option in the project file(s).
    '/R  Specifies that the Help file is to be displayed in WinHelp as soon as it is compiled.
    '/T  Turns on Translation mode. This limits what can be changed in a contents (.cnt) or project file. Only text that should be translated to another language can be changed. This can also be set by clicking the Translation command on Help Workshop’s File menu.
    'filename    Specifies the name of one or more project, Help makefile, or contents files to open or compile.
    'Example
    'The following command compiles the New.hpj project file with no compression, minimizing the window while compiling, and adding additional information to the Help file:
    'hcw /c /m /n /a new
    ReturnValue = Shell(CompilerLocation & " /c /m /a /e /r " & Location)
    Err.Clear
End Sub
Public Sub Compile_Contents(MeForm As Form, progBar As ProgressBar, treeDms As TreeView)
    On Error Resume Next
    Dim nTot As Long
    Dim nCnt As Long
    Dim nPth As String
    Dim CompiledContents As String
    Dim db As DAO.Database
    Dim tb As DAO.Recordset
    Dim tRecs As Long
    Dim tPopup As Integer
    Dim tContext As String
    Dim tTitle As String
    If boolFileExists(sProjContents) = True Then Kill sProjContents
    Set db = DAO.OpenDatabase(sProjDb)
    Set tb = db.OpenRecordset("Properties")
    tb.Index = "FullPath"
    tRecs = tb.RecordCount - 1
    CompiledContents = CompiledContents & "##v##00.4.50" & vbCrLf
    CompiledContents = CompiledContents & tRecs & vbCrLf
    nTot = treeDms.Nodes.Count
    Call ProgBarInit(MeForm, progBar, nTot)
    For nCnt = 2 To nTot
        Call UpdateProgress(MeForm, nCnt, progBar, "Compiling the contents file")
        nPth = treeDms.Nodes(nCnt).Fullpath
        tb.Seek "=", nPth
        Select Case tb.NoMatch
        Case False
            tPopup = Val(StringRemNull(tb!popup))
            tContext = StringRemNull(tb!Context)
            tTitle = StringProperCase(StringRemNull(tb!Title))
            CompiledContents = CompiledContents & tTitle & vbCrLf
            CompiledContents = CompiledContents & nCnt - 2 & vbCrLf
            CompiledContents = CompiledContents & tPopup & vbCrLf
            CompiledContents = CompiledContents & "0" & vbCrLf
        End Select
        Err.Clear
    Next
    ProgBarClose MeForm, progBar
    tb.Close
    db.Close
    Set tb = Nothing
    Set db = Nothing
    FileUpdate sProjContents, CompiledContents, "w"
    Err.Clear
End Sub
Public Sub CopyPictureNew(DestinationP As PictureBox, cd1 As MSComDlg.CommonDialog, docWord As RichTextLib.RichTextBox)
    On Error GoTo Cancel
    Dim oW As Long
    Dim oH As Long
    With cd1
        .CancelError = True
        .DialogTitle = "Select Picture..."
        .Filter = "Picture Files (*.bmp, *.jpg, *.gif, etc)|*.bmp;*.dib;*.jpeg;*.jpg;*.jpe;*.jfif;*.gif;*.tif;*.tiff;*.png"
        .ShowOpen
        DestinationP.Picture = LoadPicture(.FileName)
        oW = DestinationP.Width
        oH = DestinationP.Height
        If oW > PictureWidth Then oW = PictureWidth
        If oH > PictureHeight Then oH = PictureHeight
        DestinationP.Width = PictureWidth
        DestinationP.Height = PictureHeight
        DestinationP.PaintPicture DestinationP, 0, 0, oW, oH
        DestinationP.Picture = DestinationP.Image
        Clipboard.Clear
        Clipboard.SetData DestinationP.Picture
        SendMessage docWord.hWnd, WM_PASTE, 0, 0&
        Set DestinationP.Picture = LoadPicture()
    End With
Cancel:
    Err.Clear
End Sub
Public Sub ResizePicture(DestinationP As PictureBox, docWord As RichTextLib.RichTextBox, SourceName As String)
    On Error Resume Next
    Dim oW As Long
    Dim oH As Long
    DestinationP.AutoSize = True
    DestinationP.Picture = LoadPicture(SourceName)
    oW = DestinationP.Width
    oH = DestinationP.Height
    If oW > PictureWidth Then oW = PictureWidth
    If oH > PictureHeight Then oH = PictureHeight
    DestinationP.AutoSize = False
    DestinationP.Width = PictureWidth
    DestinationP.Height = PictureHeight
    DestinationP.PaintPicture DestinationP, 0, 0, oW, oH
    DestinationP.Picture = DestinationP.Image
    Clipboard.Clear
    Clipboard.SetData DestinationP.Picture
    SendMessage docWord.hWnd, WM_PASTE, 0, 0&
    Set DestinationP.Picture = LoadPicture()
    Err.Clear
End Sub

Public Function File_Exists(ByVal strFile As String) As Boolean
    On Error Resume Next
    Dim fs As FileSystemObject
    Set fs = New FileSystemObject
    File_Exists = fs.FileExists(strFile)
    Set fs = Nothing
    Err.Clear
End Function

Public Sub InsertFile(cd1 As MSComDlg.CommonDialog, docWord As RichTextLib.RichTextBox, Optional InsertAsLink As Boolean = True)
    On Error GoTo Cancel
    Dim s As String
    Dim f As String
    With cd1
        .DialogTitle = "Select a File..."
        .Filter = "All Files (*.*)|*.*"
        .flags = cdlOFNFileMustExist Or cdlOFNNoDereferenceLinks
        .CancelError = True
        .ShowOpen
        If Len(.FileName) = 0 Then Exit Sub
        If InsertAsLink = False Then
            docWord.OLEObjects.Add , , .FileName
        Else
            f = File_Token(.FileName, "f")
            s = sProjPath & "\" & f
            s = NextNewFile(s, False)
            Do Until File_Exists(s) = True
                FileCopy .FileName, s
                DoEvents
                Err.Clear
            Loop
            'UpdateTopicDataFiles iPath, File_Token(s, "f")
            docWord.SelText = "{mci PLAY, " & File_Token(s, "f") & "}"
        End If
    End With
Cancel:
    Err.Clear
End Sub
Public Sub InsertPicture(cd1 As MSComDlg.CommonDialog, docWord As RichTextLib.RichTextBox)
    On Error GoTo Cancel
    With cd1
        .DialogTitle = "Select Picture To Insert..."
        .CancelError = True
        .Filter = "Picture Files (*.bmp, *.jpg, *.gif, etc)|*.bmp;*.dib;*.jpeg;*.jpg;*.jpe;*.jfif;*.gif;*.tif;*.tiff;*.png"
        .ShowOpen
        docWord.OLEObjects.Add , , .FileName
    End With
Cancel:
    Err.Clear
End Sub
Public Function IsFilePicture(ByVal nStr As String) As Boolean
    On Error Resume Next
    Dim ext As String
    ext = StringGetFileToken(nStr, "e")
    Select Case ext
    Case "bmp", "jpg", "gif", "dib", "jpeg", "jpe", "jfif", "tif", "tiff", "png"
        IsFilePicture = True
    Case Else
        IsFilePicture = False
    End Select
    Err.Clear
End Function
Public Sub InsertTextFile(cd1 As MSComDlg.CommonDialog, docWord As RichTextLib.RichTextBox, FileT As RichTextLib.RichTextBox)
    On Error GoTo Cancel
    With cd1
        .DialogTitle = "Select Text File To Insert..."
        .CancelError = True
        .Filter = "Text Files (*.txt, *.csv)|*.txt;*.csv"
        .ShowOpen
        FileT.LoadFile .FileName
    End With
    docWord.SelRTF = FileT.TextRTF
Cancel:
    Err.Clear
End Sub
Public Function WordsInBetween(ByVal Sentence As String, StartWord As String, EndWord As String) As Collection
    On Error Resume Next
    Dim colSection As New Collection
    Dim sSection As String
    Dim StartLength As Long
    Dim EndLength As Long
    Dim StartWordPos As Long
    Dim EndWordPos As Long
    StartLength = Len(StartWord)
    EndLength = Len(EndWord)
    StartWordPos = InStr(1, Sentence, StartWord, vbTextCompare)
    Do Until StartWordPos = 0
        EndWordPos = InStr(StartWordPos, Sentence, EndWord, vbTextCompare)
        sSection = Mid$(Sentence, StartWordPos, (EndWordPos - StartWordPos) + EndLength)
        colSection.Add sSection
        Sentence = Mid$(Sentence, EndWordPos + EndLength)
        StartWordPos = InStr(1, Sentence, StartWord, vbTextCompare)
        Err.Clear
    Loop
    Set WordsInBetween = colSection
    Err.Clear
End Function
Public Function StrWordsInBetween(ByVal Sentence As String, StartWord As String, EndWord As String) As String
    On Error Resume Next
    Dim colSection As String
    Dim sSection As String
    Dim StartLength As Long
    Dim EndLength As Long
    Dim StartWordPos As Long
    Dim EndWordPos As Long
    StartLength = Len(StartWord)
    EndLength = Len(EndWord)
    colSection = ""
    StartWordPos = InStr(1, Sentence, StartWord, vbTextCompare)
    Do Until StartWordPos = 0
        EndWordPos = InStr(StartWordPos, Sentence, EndWord, vbTextCompare)
        sSection = Mid$(Sentence, StartWordPos, (EndWordPos - StartWordPos) + EndLength)
        colSection = colSection & sSection & VM
        Sentence = Mid$(Sentence, EndWordPos + EndLength)
        StartWordPos = InStr(1, Sentence, StartWord, vbTextCompare)
        Err.Clear
    Loop
    StrWordsInBetween = StringRemoveDelim(colSection, VM)
    Err.Clear
End Function
Public Function SaveRtfTopic(iPath As String) As String
    On Error Resume Next
    Dim db As DAO.Database
    Dim tb As DAO.Recordset
    Dim pTitle As String
    Dim pContext As String
    Dim pNumber As String
    Dim pBrowse As String
    Dim pKeyWords As String
    Dim pContents As String
    Dim pPopUp As String
    Dim MainText As String
    Dim arCode() As String
    Dim iLineCount As Long
    Dim Data As Long
    Dim Data2 As String
    Dim pFootNotes As String
    Dim CompiledRTF As String
    Dim iLineCount_Tot As Long
    Dim iLineCount_Cnt As Long
    Dim rFile As String
    rFile = ExactPath(StringGetFileToken(sProjDb, "p")) & "\" & StringGetFileToken(iPath, "fo") & ".rtf"
    Set db = DAO.OpenDatabase(sProjDb)
    Set tb = db.OpenRecordset("Properties")
    tb.Index = "FullPath"
    '    strHex$ = Hex$(HeadlineColor)
    '    strHex$ = String$(6 - Len(strHex$), "0") & strHex$
    '    strHex$ = Right$(strHex$, 2) & Mid$(strHex$, 3, 2) & Left$(strHex$, 2)
    '    R = Hex2Dec(Mid$(strHex$, 1, 2))
    '    g = Hex2Dec(Mid$(strHex$, 3, 2))
    '    b = Hex2Dec(Mid$(strHex$, 5, 2))
    '    hc = "\red" & R & "\green" & g & "\blue" & b & ";"
    '    strHex$ = Hex$(TextColor)
    '    strHex$ = String$(6 - Len(strHex$), "0") & strHex$
    '    strHex$ = Right$(strHex$, 2) & Mid$(strHex$, 3, 2) & Left$(strHex$, 2)
    '    R = Hex2Dec(Mid$(strHex$, 1, 2))
    '    g = Hex2Dec(Mid$(strHex$, 3, 2))
    '    b = Hex2Dec(Mid$(strHex$, 5, 2))
    '    tc = "\red" & R & "\green" & g & "\blue" & b & ";"
    '    strHex$ = Hex$(HeadlineBackColor)
    '    strHex$ = String$(6 - Len(strHex$), "0") & strHex$
    '    strHex$ = Right$(strHex$, 2) & Mid$(strHex$, 3, 2) & Left$(strHex$, 2)
    '    R = Hex2Dec(Mid$(strHex$, 1, 2))
    '    g = Hex2Dec(Mid$(strHex$, 3, 2))
    '    b = Hex2Dec(Mid$(strHex$, 5, 2))
    '    hbc = "\red" & R & "\green" & g & "\blue" & b & ";"
    '    strHex$ = Hex$(TextBackColor)
    '    strHex$ = String$(6 - Len(strHex$), "0") & strHex$
    '    strHex$ = Right$(strHex$, 2) & Mid$(strHex$, 3, 2) & Left$(strHex$, 2)
    '    R = Hex2Dec(Mid$(strHex$, 1, 2))
    '    g = Hex2Dec(Mid$(strHex$, 3, 2))
    '    b = Hex2Dec(Mid$(strHex$, 5, 2))
    '    tbc = "\red" & R & "\green" & g & "\blue" & b & ";"
    If FontHeadline = "" Then FontHeadline = "Tahoma"
    If FontText = "" Then FontText = "Tahoma"
    CompiledRTF = CompiledRTF & "{\rtf1\ansi\deff0" & vbCrLf
    CompiledRTF = CompiledRTF & "{\fonttbl{\f0\fnil\fcharset0 " & FontHeadline & ";}{\f1\fnil\fcharset0 " & FontText & ";}}" & vbCrLf
    'CompiledRTF = CompiledRTF & "{\colortbl " & tc & hc & tbc & hbc & "}" & vbCrLf
    tb.Seek "=", iPath
    pTitle = StringProperCase(StringRemNull(tb!Title))
    pContext = StringRemNull(tb!Context)
    pNumber = StringRemNull(tb!Number)
    pBrowse = StringRemNull(tb!browse)
    pKeyWords = StringRemNull(tb!Keywords)
    pContents = StringRemNull(tb!Contents)
    pFootNotes = StringRemNull(tb!Footnotes)
    pPopUp = StringRemNull(tb!popup)
    pKeyWords = Replace$(pKeyWords, FM, ";")
    pFootNotes = Replace$(pFootNotes, FM, ";")
    CompiledRTF = CompiledRTF & "#{\footnote " & pContext & "}" & vbCrLf
    CompiledRTF = CompiledRTF & "${\footnote " & pTitle & "}" & vbCrLf
    CompiledRTF = CompiledRTF & "K{\footnote " & pTitle & ";" & pKeyWords & "}" & vbCrLf
    If pBrowse = "1" Then
        CompiledRTF = CompiledRTF & "+{\footnote auto}" & vbCrLf
    End If
    If Len(pFootNotes) > 0 Then
        CompiledRTF = CompiledRTF & "A{\footnote " & pTitle & ";" & pFootNotes & "}" & vbCrLf
    End If
    If FontHeadlineBold = "1" Then
        CompiledRTF = CompiledRTF & "\keepn\cb3\cf1\b\fs" & FontHeadlineSize * 2 & "\f0 " & pTitle & "\fs" & FontTextSize * 2 & "\b0\cb2\par\pard\plain" & vbCrLf
        CompiledRTF = CompiledRTF & "\line\cf0\cb2\fs20\f1" & vbCrLf
    Else
        CompiledRTF = CompiledRTF & "\keepn\cb3\cf1\fs" & FontHeadlineSize * 2 & "\f0 " & pTitle & "\fs" & FontTextSize * 2 & "\cb2\par\pard\plain" & vbCrLf
        CompiledRTF = CompiledRTF & "\line\cf0\cb2\fs20\f1" & vbCrLf
    End If
    MainText = pContents
    If Len(MainText) = 0 Then GoTo NextTopic
    '=================Replace Color Codes======================
    'Do a Color Replacement and Line return replacement
    MainText = Replace$(MainText, "<color=1>", "{\cf4 ")
    MainText = Replace$(MainText, "<color=2>", "{\cf5 ") 'LIGHT PINK
    MainText = Replace$(MainText, "<color=3>", "{\cf6 ") 'PURPLE
    MainText = Replace$(MainText, "<color=4>", "{\cf7 ") 'RED
    MainText = Replace$(MainText, "<color=5>", "{\cf8 ") 'TEAL
    MainText = Replace$(MainText, "<color=6>", "{\cf9 ") 'DARK BLUE
    MainText = Replace$(MainText, "<color=7>", "{\cf10 ") 'GREEN
    MainText = Replace$(MainText, "<color=8>", "{\cf11 ") 'GOLD
    MainText = Replace$(MainText, "<color=9>", "{\cf12 ") 'GREY
    MainText = Replace$(MainText, "<color=10>", "{\cf13 ") 'BLUE
    MainText = Replace$(MainText, "<color=11>", "{\cf14 ") 'LIGHT GREEN
    MainText = Replace$(MainText, "<color=12>", "{\cf15 ") 'PINK
    MainText = Replace$(MainText, "<color=13>", "{\cf16 ") 'WHITE
    MainText = Replace$(MainText, "<color=14>", "{\cf17 ") 'LIGHT PURPLE
    MainText = Replace$(MainText, "<color=15>", "{\cf18 ") 'YELLOW
    MainText = Replace$(MainText, "<color=16>", "{\cf19 ") 'DARK GREY
    MainText = Replace$(MainText, "<\color>", "}")
    MainText = Replace$(MainText, "</color>", "}")
    MainText = Replace$(MainText, Chr$(13), "\line" & Chr$(13))
    '=================End Replace Jump Codes======================
    '=================Replace Jump Codes======================
    If InStr(1, MainText, "<jump=", vbTextCompare) > 0 Then
        arCode = Split(MainText, vbCrLf)
        'Loop through th lines 1 by 1
        iLineCount_Tot = UBound(arCode)
        iLineCount_Cnt = LBound(arCode)
        For iLineCount = iLineCount_Cnt To iLineCount_Tot
            DoEvents
            'If there is Text in the line then Look through it
            If Len(Trim$(arCode(iLineCount))) > 0 Then
                'If <jump= is found the change it
                If InStr(1, Trim$(arCode(iLineCount)), "<jump=") Then
                    'Set data as the number of letters into the text the the word <jump= starts
                    Data = InStr(1, Trim$(arCode(iLineCount)), "<jump=")
                    'Set Data2 to Equal the Entire line of Text
                    Data2 = Trim$(arCode(iLineCount))
                    'Replace the Text with what we want it to be
                    MainText = Replace$(MainText, Mid$(Data2, Data), JumpConversion(Mid$(Data2, Data)))
                End If
            End If
            Err.Clear
        Next
    End If
    '=================End Replace Jump Codes======================
    '=================Replace Target Codes======================
    'Split the Text in to lines
    If InStr(1, MainText, "<target>", vbTextCompare) > 0 Then
        arCode = Split(MainText, vbCrLf)
        'Loop through th lines 1 by 1
        iLineCount_Tot = UBound(arCode)
        iLineCount_Cnt = LBound(arCode)
        For iLineCount = iLineCount_Cnt To iLineCount_Tot
            DoEvents
            'If there is Text in the line then Look through it
            If Len(Trim$(arCode(iLineCount))) > 0 Then
                'If <jump= is found the change it
                If InStr(1, Trim$(arCode(iLineCount)), "<target>") Then
                    'Set data as the number of letters into the text the the word <jump= starts
                    Data = InStr(1, Trim$(arCode(iLineCount)), "<target>")
                    'Set Data2 to Equal the Entire line of Text
                    Data2 = Trim$(arCode(iLineCount))
                    'Replace the Text with what we want it to be
                    MainText = Replace$(MainText, Mid$(Data2, Data), TargetConversion(Mid$(Data2, Data)))
                End If
            End If
            Err.Clear
        Next
    End If
    '=================End Replace Target Codes======================
    '=================Replace URL Codes======================
    'Split the Text in to lines
    If InStr(1, MainText, "<url=", vbTextCompare) > 0 Then
        arCode = Split(MainText, vbCrLf)
        'Loop through th lines 1 by 1
        iLineCount_Tot = UBound(arCode)
        iLineCount_Cnt = LBound(arCode)
        For iLineCount = iLineCount_Cnt To iLineCount_Tot
            DoEvents
            'If there is Text in the line then Look through it
            If Len(Trim$(arCode(iLineCount))) > 0 Then
                'If <jump= is found the change it
                If InStr(1, Trim$(arCode(iLineCount)), "<url=") Then
                    'Set data as the number of letters into the text the the word <jump= starts
                    Data = InStr(1, Trim$(arCode(iLineCount)), "<url=")
                    'Set Data2 to Equal the Entire line of Text
                    Data2 = Trim$(arCode(iLineCount))
                    'Replace the Text with what we want it to be
                    MainText = Replace$(MainText, Mid$(Data2, Data), URLConversion(Mid$(Data2, Data)))
                End If
            End If
            Err.Clear
        Next
    End If
    '=================End Replace URL Codes======================
    '=================Replace Pop Codes======================
    'Split the Text in to lines
    If InStr(1, MainText, "<pop=", vbTextCompare) > 0 Then
        arCode = Split(MainText, vbCrLf)
        'Loop through th lines 1 by 1
        iLineCount_Tot = UBound(arCode)
        iLineCount_Cnt = LBound(arCode)
        For iLineCount = iLineCount_Cnt To iLineCount_Tot
            DoEvents
            'If there is Text in the line then Look through it
            If Len(Trim$(arCode(iLineCount))) > 0 Then
                'If <jump= is found the change it
                If InStr(1, Trim$(arCode(iLineCount)), "<pop=") Then
                    'Set data as the number of letters into the text the the word <jump= starts
                    Data = InStr(1, Trim$(arCode(iLineCount)), "<pop=")
                    'Set Data2 to Equal the Entire line of Text
                    Data2 = Trim$(arCode(iLineCount))
                    'Replace the Text with what we want it to be
                    MainText = Replace$(MainText, Mid$(Data2, Data), PopConversion(Mid$(Data2, Data)))
                End If
            End If
            Err.Clear
        Next
    End If
    '=================End Replace Pop Codes======================
    '=================Replace Box Codes======================
    'Split the Text in to lines
    If InStr(1, MainText, "<box", vbTextCompare) > 0 Then
        arCode = Split(MainText, vbCrLf)
        'Loop through th lines 1 by 1
        iLineCount_Tot = UBound(arCode)
        iLineCount_Cnt = LBound(arCode)
        For iLineCount = iLineCount_Cnt To iLineCount_Tot
            DoEvents
            'If there is Text in the line then Look through it
            If Len(Trim$(arCode(iLineCount))) > 0 Then
                'If <jump= is found the change it
                If InStr(1, Trim$(arCode(iLineCount)), "<box") Then
                    'Set data as the number of letters into the text the the word <jump= starts
                    Data = InStr(1, Trim$(arCode(iLineCount)), "<box")
                    'Set Data2 to Equal the Entire line of Text
                    Data2 = Trim$(arCode(iLineCount))
                    'Replace the Text with what we want it to be
                    MainText = Replace$(MainText, Mid$(Data2, Data), BoxConversion(Mid$(Data2, Data)))
                End If
            End If
            Err.Clear
        Next
    End If
    '=================End Replace Box Codes======================
    '=================Replace EXE Codes======================
    'Split the Text in to lines
    If InStr(1, MainText, "<exe=", vbTextCompare) > 0 Then
        arCode = Split(MainText, vbCrLf)
        'Loop through th lines 1 by 1
        iLineCount_Tot = UBound(arCode)
        iLineCount_Cnt = LBound(arCode)
        For iLineCount = iLineCount_Cnt To iLineCount_Tot
            DoEvents
            'If there is Text in the line then Look through it
            If Len(Trim$(arCode(iLineCount))) > 0 Then
                'If <jump= is found the change it
                If InStr(1, Trim$(arCode(iLineCount)), "<exe=") Then
                    'Set data as the number of letters into the text the the word <jump= starts
                    Data = InStr(1, Trim$(arCode(iLineCount)), "<exe=")
                    'Set Data2 to Equal the Entire line of Text
                    Data2 = Trim$(arCode(iLineCount))
                    'Replace the Text with what we want it to be
                    MainText = Replace$(MainText, Mid$(Data2, Data), EXEConversion(Mid$(Data2, Data)))
                End If
            End If
            Err.Clear
        Next
    End If
    '=================End Replace EXE Codes======================
    '=================Replace Format Codes======================
    MainText = Replace$(MainText, "<b>", "{\b ") 'Bold
    MainText = Replace$(MainText, "<i>", "{\i ") 'Italic
    MainText = Replace$(MainText, "<u>", "{\ul ") 'UNderline
    MainText = Replace$(MainText, "<l>", "\pard\par\ql ") 'Align L
    MainText = Replace$(MainText, "<c>", "\pard\par\qc ") 'Align Center
    MainText = Replace$(MainText, "<r>", "\pard\par\qr ") 'Align Right
    MainText = Replace$(MainText, "</b>", "}") 'Align Right
    MainText = Replace$(MainText, "</i>", "}") 'Align Right
    MainText = Replace$(MainText, "</u>", "}") 'Align Right
    MainText = Replace$(MainText, "</l>", "\par\pard") 'Align Right
    MainText = Replace$(MainText, "</c>", "\par\pard") 'Align Right
    MainText = Replace$(MainText, "</r", "\par\pard") 'Align Right
NextTopic:
    '=================End Replace Format Codes======================
    CompiledRTF = CompiledRTF & MainText & "\line" & vbCrLf
    CompiledRTF = CompiledRTF & "\plain\par\page"
    CompiledRTF = CompiledRTF & "}"
    tb.Close
    db.Close
    Set tb = Nothing
    Set db = Nothing
    FileUpdate rFile, CompiledRTF, "w"
    SaveRtfTopic = rFile
    Err.Clear
End Function
Private Sub RunMsWord()
    On Error GoTo ErrHandler
    Dim intRet As Long
    Dim blnPause As Boolean
    Dim lngLoop As Long
    Set mobjWord97 = New Word.Application
    mobjWord97.WindowState = wdWindowStateMinimize
    mobjWord97.Visible = False
    blnWord97Loaded = True
    Err.Clear
    Exit Sub
ErrHandler:
    Select Case Err
    Case -2147023179, -2147023174, 462
        Set mobjWord97 = Nothing
        If Not blnPause Then
            For lngLoop = 0 To 50000
                Err.Clear
            Next
            blnPause = True
            Resume
        Else
            intRet = MsgBox("Your system is busy at the moment and " & "Microsoft Word couldn't be loaded." & vbCrLf & vbCrLf & "Would you like to try again?", vbYesNo + vbExclamation)
            If intRet = vbYes Then
                Resume
            Else
                blnWord97Loaded = False
                Err.Clear
                Exit Sub
            End If
        End If
    Case 429
        MsgBox "Microsoft Word is not installed on this computer.", vbCritical
        blnWord97Loaded = False
        Err.Clear
        Exit Sub
    Case Else
        Err.Clear
        Exit Sub
    End Select
    Err.Clear
End Sub
Sub Word97Do(strMode As WordOperation, docWord As RichTextBox, Optional ByVal TopicContext As String = "")
    On Error Resume Next
    Dim objDoc As Word.Document
    Dim rFile As String
    rFile = App.Path & "\word.rtf"
    If boolFileExists(rFile) = True Then Kill rFile
    docWord.SaveFile rFile, rtfRTF
    mobjWord97.WindowState = wdWindowStateMinimize
    Set objDoc = New Word.Document
    Set objDoc = mobjWord97.Documents.Open(rFile)
    mobjWord97.Visible = True
    mobjWord97.Activate
    objDoc.Activate
    Select Case strMode
    Case wSpelling
        If objDoc.SpellingErrors.Count > 0 Then
            objDoc.CheckSpelling
            objDoc.Save
            objDoc.Close
            Set objDoc = Nothing
            docWord.LoadFile rFile, rtfRTF
            DoEvents
        Else
            mobjWord97.Visible = False
            objDoc.Close
            Set objDoc = Nothing
            MsgBox "No spelling errors found!", vbInformation, "MyHelp"
            Err.Clear
            Exit Sub
        End If
    Case wGrammar
        If objDoc.GrammaticalErrors.Count > 0 Then
            objDoc.CheckGrammar
            objDoc.Save
            objDoc.Close
            Set objDoc = Nothing
            docWord.LoadFile rFile, rtfRTF
            DoEvents
        Else
            mobjWord97.Visible = False
            objDoc.Close
            Set objDoc = Nothing
            MsgBox "No grammatical errors found!", vbInformation, "MyHelp"
            Err.Clear
            Exit Sub
        End If
    Case wSaveHTML
        rFile = sProjHTML & "\" & TopicContext & ".html"
        objDoc.SaveAs rFile, Word.wdFormatHTML
        objDoc.Close
        Set objDoc = Nothing
    End Select
    Err.Clear
End Sub
Public Sub LstViewCheckAll(lstView As ListView, Optional ByVal bOp As Boolean = True)
    On Error Resume Next
    Dim lstTot As Long
    Dim lstCnt As Long
    lstTot = lstView.ListItems.Count
    For lstCnt = 1 To lstTot
        lstView.ListItems(lstCnt).Checked = bOp
        Err.Clear
    Next
    Err.Clear
End Sub
Public Function InternalFileLink(ByVal sFile As String) As String
    On Error Resume Next
    Dim spSections() As String
    Dim spTot As Long
    Dim spCnt As Long
    Dim spNew As String
    Dim xNew As Long
    Dim lNew As Long
    spNew = ""
    spSections = Split(sFile, "\")
    spTot = UBound(spSections)
    Select Case spTot
    Case 1
        InternalFileLink = spSections(1)
    Case Else
        xNew = spTot - 2
        For spCnt = 0 To xNew
            spNew = spNew & ".. \"
            Err.Clear
        Next
        lNew = xNew + 1
        For spCnt = lNew To spTot
            spNew = spNew & spSections(spCnt) & "\"
            Err.Clear
        Next
        spNew = StringRemoveDelim(spNew, "\")
    End Select
    InternalFileLink = spNew
    Erase spSections
    Err.Clear
End Function
Public Sub SaveReg(ByVal sKey As String, ByVal sValue As String, Optional ByVal sSection As String = "account", Optional ByVal sAppName As String = "")
    On Error Resume Next
    If Len(sAppName) = 0 Then sAppName = App.Title
    sValue = Replace$(sValue, vbCr, vbNullChar)
    sValue = Replace$(sValue, vbLf, vbNullChar)
    SaveSetting sAppName, sSection, sKey, sValue
    Err.Clear
End Sub
Public Function ReadReg(ByVal sKey As String, Optional ByVal sSection As String = "account", Optional ByVal sAppName As String = "") As String
    On Error Resume Next
    If Len(sAppName) = 0 Then sAppName = App.Title
    ReadReg = GetSetting(sAppName, sSection, sKey)
    Err.Clear
End Function
Function ExactPath(ByVal strValue As String) As String
    On Error Resume Next
    If Right$(strValue, 1) = "\" Then
        strValue = Left$(strValue, Len(strValue) - 1)
    End If
    ExactPath = strValue
    Err.Clear
End Function
Public Sub MakeDirectory(ByVal Sdirectory As String)
    On Error GoTo CreateDirectory_ErrorHandler
    CreateNestedDirectory Sdirectory
    Err.Clear
    Exit Sub
CreateDirectory_ErrorHandler:
    Select Case Err
    Case 0
    Case 75
    Case Else
        retAnswer = MyPrompt("Directory Name : " & Sdirectory & vbCr & vbCr & "Error " & VBA.CStr(Err) & ":" & "  " & error$ & vbCr & "Please check your drive and disk." & vbCr & vbCr & "Directory Name" & Sdirectory, "o", "w", "Create Directory")
    End Select
    Err.Clear
End Sub
Public Function boolDirExists(ByVal Sdirname As String) As Boolean
    On Error Resume Next
    Dim sDir As String
    boolDirExists = False
    sDir = Dir$(Sdirname, vbDirectory)
    If (Len(sDir) > 0) And (Err = 0) Then
        boolDirExists = True
    End If
    Err.Clear
End Function
Public Sub KillFolderTree(ByVal sFolder As String)
    On Error Resume Next
    Dim sCurrFilename As String
    sCurrFilename = Dir$(sFolder & "\*.*", vbDirectory Or vbArchive Or vbNormal Or vbHidden Or vbReadOnly)
    Do While sCurrFilename <> ""
        If sCurrFilename <> "." And sCurrFilename <> ".." Then
            If (GetAttr(sFolder & "\" & sCurrFilename) And vbDirectory) = vbDirectory Then
                Call KillFolderTree(sFolder & "\" & sCurrFilename)
                sCurrFilename = Dir$(sFolder & "\*.*", vbDirectory Or vbArchive Or vbNormal Or vbHidden Or vbReadOnly)
            Else
                Kill sFolder & "\" & sCurrFilename
                sCurrFilename = Dir$
            End If
        Else
            sCurrFilename = Dir$
        End If
        Err.Clear
    Loop
    RmDir sFolder
    Err.Clear
End Sub
Sub TreeViewSaveToTable(MeForm As Form, progBar As ProgressBar, ByVal Dbase As String, ByVal TbName As String, TreeV As TreeView)
    On Error Resume Next
    'Ask the user for the name of a mdb and table if it does not exist create it.
    'Then store all of the nodes from the TreeView into the table.
    StatusMessage MeForm, "Saving treeview, please be patient..."
    Dim mDB As DAO.Database
    Dim mRs As DAO.Recordset
    If boolFileExists(Dbase) = False Then
        dbCreate Dbase
    End If
    'Do Until dbTableExists(Dbase, TbName) = True
    dbCreateTable Dbase, TbName, _
    "Sequence,Key,Parent,Text,Image,SelectedImage,Tag,Checked,Bold,ForeColor,FullPath", _
    "long,lo,lo,me,te,te,me,lo,lo,lo,me", _
    ",,,,255,255,,,,255,", _
    "1,2,3,4,11"
    '    Err.Clear
    'Loop
    Set mDB = DAO.OpenDatabase(Dbase)
    Set mRs = mDB.OpenRecordset(TbName)
    TreeViewWriteToTable MeForm, progBar, mRs, TreeV
    mRs.Close
    mDB.Close 'close the database
    ProgBarClose MeForm, progBar
    Err.Clear
End Sub
Public Sub dbExecute(ByVal DbName As String, ByVal Strqry As String)
    On Error Resume Next
    Dim db As DAO.Database
    Set db = DAO.OpenDatabase(DbName)
    db.Execute Strqry
    db.Close
    Set db = Nothing
    DAO.DBEngine.Idle
    Err.Clear
End Sub
Public Sub CleanAllControls(Thisform As Form)
    On Error Resume Next
    Dim ctlControl As Control
    For Each ctlControl In Thisform.Controls
        Select Case ctlControl.Name
        Case "cmbNames", "cmbLists", "cmbSort", "cmbWork"
            GoTo NextSection
        Case "cmbType", "cmbFormat", "cmbJustify", "cmbDepth"
            GoTo NextSection
        End Select
        ctlControl.Text = ""
        ctlControl.Clear
        ctlControl.ListIndex = -1
        ctlControl.Value = 0
        ctlControl.Enabled = True
NextSection:
        Err.Clear
    Next
    Err.Clear
End Sub
Public Function RecycleFile(OwnerForm As Form, fromPaths As String, Optional toPaths As String = "", Optional intPerform As FileOps = foCopy) As Boolean
    On Error Resume Next
    DoEvents
    Dim FileOperation As SHFILEOPSTRUCT
    Dim lReturn As Long
    With FileOperation
        .hWnd = OwnerForm.hWnd
        Select Case intPerform
        Case 0
            .wFunc = FO_DELETE
        Case 1
            .wFunc = FO_MOVE
        Case 2
            .wFunc = FO_COPY
        Case 3
            .wFunc = FO_RENAME
        End Select
        .pFrom = fromPaths & vbNullChar & vbNullChar
        '.fFlags = FOF_SIMPLEPROGRESS Or FOF_ALLOWUNDO Or FOF_CREATEPROGRESSDLG
        .fFlags = FOF_ALLOWUNDO
        If Len(toPaths) > 0 Then
            .pTo = toPaths & vbNullChar & vbNullChar
        End If
    End With
    lReturn = SHFileOperation(FileOperation)
    RecycleFile = True
    If lReturn <> 0 Then
        ' Operation failed
        RecycleFile = False
    Else
        If FileOperation.fAnyOperationsAborted <> 0 Then
            RecycleFile = False
        End If
    End If
    Err.Clear
End Function
Sub RecurseFolderToComboBox(ByVal StartDirectory As String, cboBox As ComboBox, Optional FilesOnly As Boolean = True, Optional NamesOnly As Boolean = False)
    On Error Resume Next
    Dim fso As New Scripting.FileSystemObject
    Dim objFolder As Scripting.Folder
    Dim objFolders As Scripting.Folders
    Dim objFiles As Scripting.Files
    Dim objEachFolder As Scripting.Folder
    Dim objEachFile As Scripting.File
    Set objFolder = fso.GetFolder(StartDirectory)
    Set objFolders = objFolder.SubFolders
    Set objFiles = objFolder.Files
    For Each objEachFolder In objFolders
        DoEvents
        If FilesOnly = False Then
            If NamesOnly = True Then
                cboBox.AddItem StringGetFileToken(objEachFolder, "f")
            Else
                cboBox.AddItem objEachFolder
            End If
        End If
        RecurseFolderToComboBox objEachFolder, cboBox, FilesOnly, NamesOnly
        Err.Clear
    Next
    For Each objEachFile In objFiles
        DoEvents
        If FilesOnly = True Then
            If NamesOnly = True Then
                cboBox.AddItem StringGetFileToken(objEachFile, "f")
            Else
                cboBox.AddItem objEachFile
            End If
        End If
        Err.Clear
    Next
    Set fso = Nothing
    Set objFolder = Nothing
    Set objFolders = Nothing
    Set objFiles = Nothing
    Set objEachFolder = Nothing
    Set objEachFile = Nothing
    Err.Clear
End Sub
Sub ReloadMenu(MenuControl As Variant, LstFrom As ComboBox, Optional ByVal StrOwn As String = "", Optional ByVal StrBlank As String = "", Optional ByVal Ascending As Boolean = True, Optional ByVal NumberOfItemsToLoad As Long = -1, Optional ProperCase As Boolean = False)
    On Error Resume Next
    Dim menuTot As Long
    Dim menuCnt As Long
    Dim menuItm As Long
    Dim cntItem As Long
    Dim menuCnt_Cnt As Long
    cntItem = 0
    menuCnt = 0
    ' how many items are there yet
    menuTot = MenuControl.Count - 1
    For menuCnt = menuTot To 1 Step -1
        Unload MenuControl(menuCnt)
        Err.Clear
    Next
    menuItm = LstFrom.ListCount - 1
    If Ascending = True Then
        MenuControl(0).Caption = IIf((ProperCase = True), StringProperCase(LstFrom.List(0)), LstFrom.List(0))
        MenuControl(0).Checked = False
    Else
        MenuControl(0).Caption = IIf((ProperCase = True), StringProperCase(LstFrom.List(menuItm)), LstFrom.List(menuItm))
        MenuControl(0).Checked = False
    End If
    If MenuControl(0).Caption = "" Then
        MenuControl(0).Caption = "<No Records>"
    End If
    If menuItm = -1 Then
        MenuControl(0).Enabled = False
    Else
        MenuControl(0).Enabled = True
    End If
    Select Case Ascending
    Case True
        If NumberOfItemsToLoad < 0 Then
            NumberOfItemsToLoad = menuItm
        End If
        For menuCnt = 1 To NumberOfItemsToLoad
            Load MenuControl(menuCnt)
            MenuControl(menuCnt).Caption = IIf((ProperCase = True), StringProperCase(LstFrom.List(menuCnt)), LstFrom.List(menuCnt))
            Err.Clear
        Next
    Case Else
        If NumberOfItemsToLoad < 0 Then
            menuCnt_Cnt = menuItm - 1
            For menuCnt = menuCnt_Cnt To 1 Step -1
                cntItem = cntItem + 1
                Load MenuControl(cntItem)
                MenuControl(cntItem).Caption = IIf((ProperCase = True), StringProperCase(LstFrom.List(menuCnt)), LstFrom.List(menuCnt))
                Err.Clear
            Next
        Else
            menuCnt_Cnt = menuItm - 1
            For menuCnt = menuCnt_Cnt To 1 Step -1
                cntItem = cntItem + 1
                If cntItem + 1 > NumberOfItemsToLoad Then
                    Exit For
                End If
                Load MenuControl(cntItem)
                MenuControl(cntItem).Caption = IIf((ProperCase = True), StringProperCase(LstFrom.List(menuCnt)), LstFrom.List(menuCnt))
                Err.Clear
            Next
        End If
    End Select
    If Len(StrOwn) > 0 Then
        Select Case MenuControl(0).Caption
        Case "<No Records>"
            MenuControl(0).Enabled = True
            MenuControl(0).Caption = "<Enter Own>"
        Case Else
            menuCnt = MenuControl.Count - 1
            menuCnt = menuCnt + 1
            Load MenuControl(menuCnt)
            MenuControl(menuCnt).Caption = "<Enter Own>"
        End Select
    End If
    If Len(StrBlank) > 0 Then
        Select Case MenuControl(0).Caption
        Case "<No Records>"
            MenuControl(0).Enabled = True
            MenuControl(0).Caption = "<Blank>"
        Case Else
            menuCnt = MenuControl.Count - 1
            menuCnt = menuCnt + 1
            Load MenuControl(menuCnt)
            MenuControl(menuCnt).Caption = "<Blank>"
        End Select
    End If
    Err.Clear
End Sub
Public Sub StatusMessage(Thisform As Form, Optional ByVal Rsmsg As String = "", Optional ByVal pos As Integer = 5)
    On Error Resume Next
    If Val(pos) = 0 Then
        pos = 1
    End If
    Thisform.StatusBar1.Panels.Item(pos) = StringProperCase(Rsmsg)
    Thisform.StatusBar1.Refresh
    Err.Clear
End Sub
Public Sub dbCreate(ByVal DbName As String, Optional ByVal Overwrite As Boolean = False, Optional ByVal Version As DAO.DatabaseTypeEnum = dbVersion40)
    On Error Resume Next
    Dim fExist As Boolean
    Dim dbPath As String
    Dim wrkDefault As DAO.Workspace
    Dim dbsNew As DAO.Database
    dbPath = StringGetFileToken(DbName, "p")
    CreateNestedDirectory dbPath
    fExist = boolFileExists(DbName)
    If fExist = True Then
        If Overwrite = True Then
            Kill DbName
        Else
            Err.Clear
            Exit Sub
        End If
    End If
    Set wrkDefault = DAO.DBEngine.Workspaces(0)
    DbName = StringProperCase(DbName)
    Set dbsNew = wrkDefault.CreateDatabase(DbName, DAO.dbLangGeneral, Version)
    Set dbsNew = Nothing
    Set wrkDefault = Nothing
    Err.Clear
End Sub
Public Sub dbDeleteTables(ByVal DbName As String, ParamArray Items())
    On Error Resume Next
    Dim test As String
    Dim Item As Variant
    Dim db As DAO.Database
    Set db = DAO.OpenDatabase(DbName)
    For Each Item In Items
        test = LCase$(CStr(Item))
        test = db.TableDefs(test).Name
        If Err = 0 Then
            db.TableDefs.Delete test
        End If
        Err = 0
        Err.Clear
    Next
    db.Close
    Set db = Nothing
    Err.Clear
End Sub
Public Function dbTableExists(ByVal Dbase As String, ByVal TbName As String) As Boolean
    On Error Resume Next
    Dim DatCt As Long
    Dim StrDt As String
    Dim zCnt As Long
    Dim db As DAO.Database
    TbName = StringProperCase(TbName)
    TbName = StringIconv(TbName, "t")
    dbTableExists = False
    Set db = DAO.OpenDatabase(Dbase)
    With db
        zCnt = .TableDefs.Count - 1
        For DatCt = 0 To zCnt
            StrDt = StringProperCase(.TableDefs(DatCt).Name)
            Select Case StrDt
            Case TbName
                dbTableExists = True
                Exit For
            End Select
            Err.Clear
        Next
    End With
    db.Close
    Set db = Nothing
    Err.Clear
End Function
Public Sub dbCreateTable(ByVal DbName As String, ByVal dbTable As String, ByVal FldName As String, Optional ByVal fldType As String = "", Optional ByVal fldSize As String = "", Optional ByVal Fldidx As String = "", Optional ByVal FldAutoIncrement As String = "")
    On Error Resume Next
Recreate:
    DoEvents
    Dim spFlds() As String
    Dim spType() As String
    Dim spSize() As String
    Dim spIndx() As String
    Dim spAuto() As String
    Dim totFld As Integer
    Dim totIdx As Integer
    Dim NewFld As DAO.Field
    Dim NewIdx As DAO.Index
    Dim NewTb As DAO.TableDef
    Dim NewDb As DAO.Database
    Dim newCnt As Integer
    Dim newPos As Integer
    Dim NewType As Integer
    FldName = MvRemoveBlanks(FldName, ",")
    Fldidx = MvRemoveBlanks(Fldidx, ",")
    FldAutoIncrement = MvRemoveBlanks(FldAutoIncrement, ",")
    Call StringParse(spFlds, FldName, ",")
    Call StringParse(spType, fldType, ",")
    Call StringParse(spSize, fldSize, ",")
    Call StringParse(spIndx, Fldidx, ",")
    Call StringParse(spAuto, FldAutoIncrement, ",")
    ArrayTrimItems spFlds
    ArrayTrimItems spType
    ArrayTrimItems spSize
    ArrayTrimItems spIndx
    ArrayTrimItems spAuto
    totFld = UBound(spFlds)
    totIdx = UBound(spIndx)
    ReDim Preserve spType(totFld)
    ReDim Preserve spSize(totFld)
    dbTable = StringIconv(dbTable, "t")
    Set NewDb = DAO.OpenDatabase(DbName)
    Set NewTb = NewDb.CreateTableDef(dbTable)
    For newCnt = 1 To totFld
        spType(newCnt) = Trim$(spType(newCnt))
        spFlds(newCnt) = Trim$(spFlds(newCnt))
        spSize(newCnt) = Trim$(spSize(newCnt))
        If Len(spType(newCnt)) = 0 Then
            spType(newCnt) = "Text"
        End If
        If Len(spSize(newCnt)) = 0 Then
            spSize(newCnt) = "255"
        End If
        NewType = dbType(spType(newCnt))
        spFlds(newCnt) = spFlds(newCnt)
        Set NewFld = NewTb.CreateField(spFlds(newCnt), NewType)
        Select Case NewType
        Case dbText
            NewFld.AllowZeroLength = True
            NewFld.Size = spSize(newCnt)
        Case dbMemo
            NewFld.AllowZeroLength = True
        Case dbLong, dbInteger, dbDouble
            NewFld.DefaultValue = ""
        End Select
        If MvSearch(FldAutoIncrement, spFlds(newCnt), ",") > 0 Then
            NewFld.Type = DAO.dbLong
            NewFld.Attributes = DAO.dbAutoIncrField
        End If
        NewTb.Fields.Append NewFld
        Err.Clear
    Next
    For newCnt = 1 To totIdx
        spIndx(newCnt) = Trim$(spIndx(newCnt))
        newPos = spIndx(newCnt)
        Set NewIdx = NewTb.CreateIndex(spFlds(newPos))
        Set NewFld = NewIdx.CreateField(spFlds(newPos))
        NewIdx.Fields.Append NewFld
        NewTb.Indexes.Append NewIdx
        Err.Clear
    Next
NextSection:
    NewDb.TableDefs.Append NewTb
    Select Case Err
    Case 3010, 3012       ' already exists/locked
        NewDb.TableDefs.Delete dbTable
        GoTo NextSection
    Case 3006, 3009, 3008
        GoTo NextSection1
    Case 3001
        Err.Clear
        Exit Sub
    End Select
NextSection1:
    NewDb.Close
    Set NewFld = Nothing
    Set NewIdx = Nothing
    Set NewTb = Nothing
    Set NewDb = Nothing
    DAO.DBEngine.Idle
    DoEvents
    If dbTableExists(DbName, dbTable) = False Then GoTo Recreate
    Err.Clear
End Sub
Sub TreeViewLoadFromTable(MeForm As Form, progBar As ProgressBar, ByVal Dbase As String, TreeFor As TreeView, ByVal TbName As String, Optional ByVal TreeClear As String = "y")
    On Error Resume Next
    If Len(TbName) = 0 Then
        Err.Clear
        Exit Sub
    End If
    Dim oNodex As Node
    Dim nImage As String
    Dim nSelectedImage As String
    Dim nText As String
    Dim nTag As String
    Dim nKey As String
    Dim nParent As String
    Dim mDB As DAO.Database
    Dim mRs As DAO.Recordset
    Dim qrySql As String
    Dim nBold As String
    Dim nChecked As String
    Dim nForeColor As String
    Dim rsTot As Long
    Dim rsCnt As Long
    TreeClear = LCase$(TreeClear)
    If TreeClear = "y" Then
        TreeViewClearAPI TreeFor
    End If
    qrySql = "select * from [" & TbName & "] order by sequence;"
    Set mDB = DAO.OpenDatabase(Dbase)
    Set mRs = mDB.OpenRecordset(qrySql)
    With mRs
        .MoveLast
        rsTot = .RecordCount
        .MoveFirst
        Call ProgBarInit(MeForm, progBar, rsTot)
        For rsCnt = 1 To rsTot
            Call UpdateProgress(MeForm, rsCnt, progBar, "Loading tree view from table...")
            nBold = ""
            nChecked = ""
            nForeColor = ""
            nImage = StringRemNull(!Image)
            nSelectedImage = StringRemNull(!SelectedImage)
            nText = StringRemNull(!Text)
            nTag = StringRemNull(!Tag)
            nBold = StringRemNull(!Bold)
            nChecked = StringRemNull(!Checked)
            nForeColor = StringRemNull(!ForeColor)
            nKey = "K" & StringRemNull(!Key)
            nParent = "K" & StringRemNull(!Parent)
            Select Case nImage
            Case "28": nImage = "book"
            Case "29": nImage = "openbook"
            Case "31": nImage = "file"
            Case "32": nImage = "openfile"
            End Select
            Select Case nSelectedImage
            Case "28": nSelectedImage = "book"
            Case "29": nSelectedImage = "openbook"
            Case "31": nSelectedImage = "file"
            Case "32": nSelectedImage = "openfile"
            End Select
            Select Case nParent
            Case "K0"          ' parent node
                Set oNodex = TreeFor.Nodes.Add(, 1, nKey, nText, nImage, nSelectedImage)
                oNodex.Tag = nTag
                If Val(nBold) = 1 Then oNodex.Bold = True
                If Val(nChecked) = 1 Then oNodex.Checked = True
                If Val(nForeColor) = 0 Then
                    oNodex.ForeColor = vbBlack
                Else
                    oNodex.ForeColor = Val(nForeColor)
                End If
            Case Else          ' child nodes have a parent reference
                Set oNodex = TreeFor.Nodes.Add(nParent, 4, nKey, nText, nImage, nSelectedImage)
                oNodex.Tag = nTag
                If Val(nBold) = 1 Then oNodex.Bold = True
                If Val(nChecked) = 1 Then oNodex.Checked = True
                If Val(nForeColor) = 0 Then
                    oNodex.ForeColor = vbBlack
                Else
                    oNodex.ForeColor = Val(nForeColor)
                End If
            End Select
            .MoveNext
            Err.Clear
        Next
    End With
closetb:
    mRs.Close 'Close the table
    mDB.Close 'Close the database
    Set oNodex = Nothing
    ProgBarClose MeForm, progBar
    Err.Clear
End Sub
'Function ProgBarInitX(ByVal totRecords As Long, Optional OnTop As Boolean = True, Optional StopEnable As Boolean = True, Optional ByVal sNote As String = "") As Boolean
'    On Error Resume Next
'    ProgBarInit = True
'    Running = True
'    frmPg.chkPrg.Width = 0
'    frmPg.lblTime.Caption = ""
'    frmPg.lblPerc.Caption = ""
'    frmPg.lblRecs.Caption = ""
'    frmPg.lblHeader.Caption = ""
'    frmPg.cmdStop.Tag = ""
'    frmPg.cmdStop.Enabled = StopEnable
'    If totRecords <= 0 Then
'        ProgBarInit = False
'        Screen.MousePointer = vbDefault
'        frmPg.Hide
'        Err.Clear
'        Exit Function
'    End If
'    frmPg.Show
'    frmPg.Refresh
'    If OnTop = True Then
'        aHandle = FindWindowByTitle("", "Processing")
'        If aHandle > 0 Then
'            ApplicationOnTop aHandle
'        End If
'    End If
'    If Len(sNote) > 0 Then
'        frmPg.ProgressShow totRecords, totRecords, StringProperCase(sNote)
'    End If
'    Err.Clear
'End Function
'Function UpdateProgress(ByVal minValue As Long, ByVal maxValue As Long, Optional ByVal Note As String = "") As Boolean
'    On Error Resume Next
'    frmPg.ProgressShow minValue, maxValue, StringProperCase(Note)
'    If frmPg.cmdStop.Tag = "s" Then
'        UpdateProgress = False
'        Unload frmPg
'        Running = False
'    Else
'        Running = True
'        UpdateProgress = True
'    End If
'    If minValue > maxValue Then
'        Running = False
'        UpdateProgress = False
'        Unload frmPg
'    End If
'    DoEvents
'    Err.Clear
'End Function
'Sub CloseProgress()
'    On Error Resume Next
'    frmPg.cmdStop.Tag = "s"
'    Unload frmPg
'    Running = False
'    Err.Clear
'End Sub
Public Sub ProgBarInit(MeForm As Form, progBar As ProgressBar, totItems As Long, Optional Note As String = "")
    On Error Resume Next
    progBar.Value = 0
    If totItems > 0 Then
        progBar.Max = totItems
        progBar.Min = 0
    End If
    If Len(Note) > 0 Then StatusMessage MeForm, Note
    Err.Clear
End Sub
Public Function boolFileExists(ByVal FileName As String) As Boolean
    On Error Resume Next
    boolFileExists = False
    If Len(FileName) = 0 Then
        Err.Clear
        Exit Function
    End If
    boolFileExists = IIf(Dir$(FileName) <> "", True, False)
    Err.Clear
End Function
Public Function boolViewFile(ByVal FileName As String, Optional ByVal Operation As String = "Open", Optional ByVal WindowState As Long = 1) As Boolean
    On Error Resume Next
    Dim R As Long
    R = lngStartDoc(FileName, Operation, WindowState)
    If R <= 32 Then
        ' there was an error
        Beep
        retAnswer = MyPrompt("An error occurred while opening your document." & vbCr & "The possibility is that the selected entry does not have" & vbCr & "a link in the registry to open it with.", "o", "w", "Viewer Error")
        boolViewFile = False
    Else
        boolViewFile = True
        Pause 1
    End If
    Err.Clear
End Function
Public Function StringBrowseForFolder(hWndOwner As Long, sPrompt As String) As String
    On Error Resume Next
    Dim iNull As Integer
    Dim lpIDList As Long
    Dim lResult As Long
    Dim sPath As String
    Dim udtBI As BrowseInfo
    With udtBI
        .hWndOwner = hWndOwner
        .lpszTitle = lstrcat(sPrompt, "")
        .ulFlags = BIF_RETURNONLYFSDIRS
    End With
    lpIDList = SHBrowseForFolder(udtBI)
    If lpIDList Then
        sPath = String$(MAX_PATH, 0)
        lResult = SHGetPathFromIDList(lpIDList, sPath)
        Call CoTaskMemFree(lpIDList)
        iNull = InStr(sPath, vbNullChar)
        If iNull Then
            sPath = Left$(sPath, iNull - 1)
        End If
    End If
    StringBrowseForFolder = sPath
    Err.Clear
End Function
Public Function StringGetFileToken(ByVal strFileName As String, Optional ByVal Sretrieve As String = "", Optional ByVal Delim As String = "\") As String
    On Error Resume Next
    Dim intNum As Long
    Dim sNew As String
    StringGetFileToken = strFileName
    If Len(Sretrieve) = 0 Then
        Sretrieve = "F"
    End If
    If Len(Delim) = 0 Then
        Delim = "\"
    End If
    Select Case UCase$(Sretrieve)
    Case "D"
        StringGetFileToken = Left$(strFileName, 3)
    Case "F"
        intNum = InStrRev(strFileName, Delim)
        If intNum <> 0 Then
            StringGetFileToken = Mid$(strFileName, intNum + 1)
        End If
    Case "P"
        intNum = InStrRev(strFileName, Delim)
        If intNum <> 0 Then
            StringGetFileToken = Mid$(strFileName, 1, intNum - 1)
        End If
    Case "E"
        intNum = InStrRev(strFileName, ".")
        If intNum <> 0 Then
            StringGetFileToken = Mid$(strFileName, intNum + 1)
        End If
    Case "FO"
        sNew = strFileName
        intNum = InStrRev(sNew, Delim)
        If intNum <> 0 Then
            sNew = Mid$(sNew, intNum + 1)
        End If
        intNum = InStrRev(sNew, ".")
        If intNum <> 0 Then
            sNew = Left$(sNew, intNum - 1)
        End If
        StringGetFileToken = sNew
    Case "PF"
        intNum = InStrRev(strFileName, ".")
        If intNum <> 0 Then
            StringGetFileToken = Left$(strFileName, intNum - 1)
        End If
    End Select
    Err.Clear
End Function
Public Sub ShowFrame(fraObject As Variant, ByVal L As Long, ByVal t As Long, ByVal W As Long, ByVal h As Long, Optional ByVal Caption As String = "", Optional ShowBorder As Boolean = True)
    On Error Resume Next
    With fraObject
        .Left = L
        .Top = t
        .Width = W
        .Height = h
        .Visible = True
        If ShowBorder = False Then .BorderStyle = 0
        If IsMissing(Caption) = False Then .Caption = Caption
        .ZOrder 0
    End With
    Err.Clear
End Sub
Public Sub TreeViewMoveNode(tvw As TreeView, nodX As Node, ByVal Direction As String)
    On Error Resume Next
    Dim nodN As Node
    Dim strKey As String
    'All we do here is copy the node and set it as the previous
    'Nodes previous node. A little confusing, but it works.
    'We then add all the children and delete the original
    'Node
    With tvw
        Select Case Direction
        Case "UP"
            If Not nodX.Previous Is Nothing Then
                Set nodN = .Nodes.Add(nodX.Previous, tvwPrevious, , nodX.Text, nodX.Image, nodX.SelectedImage)
                nodN.Tag = nodX.Tag
            Else
                Err.Clear
                Exit Sub
            End If
        Case "DOWN"
            If Not nodX.Next Is Nothing Then
                Set nodN = .Nodes.Add(nodX.Next, tvwNext, , nodX.Text, nodX.Image, nodX.SelectedImage)
                nodN.Tag = nodX.Tag
            Else
                Err.Clear
                Exit Sub
            End If
        End Select
        nodN.Selected = True
        If nodX.Children <> 0 Then
            TreeViewGetChildren tvw, nodX, nodN
        End If
        strKey = nodX.Key
        .Nodes.Remove nodX.Index
        Set nodX = Nothing
        nodN.Key = strKey
    End With
    Err.Clear
End Sub
Public Sub Dao_ViewSQLNew(MeForm As Form, progBar As ProgressBar, ByVal DataSource As String, ByVal RecordSource As String, lstView As ListView, Optional ByVal Headings As String = "", Optional ByVal lstClear As String = "", Optional ByVal strImage As String = "", Optional ByVal strSelImage As String = "", Optional varCorrelate As Variant, Optional ByVal UpperCase As Boolean = False, Optional ByVal TagFldName As String = "", Optional ByVal ShowDay As String = "", Optional ByVal AmountFlds As String = "")
    On Error Resume Next
    Dim fldArray() As String
    Dim FldCnt As Long
    Dim FldTot As Long
    Dim recTot As Long
    Dim recCnt As Long
    Dim recStr As String
    Dim varTot As Integer
    Dim varCnt As Integer
    Dim varStr As String
    Dim tableP As String
    Dim tableN As String
    Dim tableK As String
    Dim tableF As String
    Dim varCo() As String
    Dim curVal As String
    Dim tagVal As String
    Dim db As DAO.Database
    Dim tb As DAO.Recordset
    Dim lstPos As Long
    Dim rFld As String
    With lstView
        If Len(lstClear) = 0 Then
            .ListItems.Clear
        End If
        .Sorted = False
        .View = 3
        .GridLines = True
        .FullRowSelect = True
    End With
    Set db = DAO.OpenDatabase(DataSource)
    Set tb = db.OpenRecordset(RecordSource)
    tb.MoveLast
    recTot = tb.RecordCount
    'if headers are specified, get them
    If Len(Headings) > 0 Then
        LstViewMakeHeadings lstView, Headings, lstClear
    Else
        Headings = TbFldNames(tb)
        LstViewMakeHeadings lstView, Headings, lstClear
    End If
    If recTot = 0 Then
        tb.Close
        db.Close
        Err.Clear
        Exit Sub
    End If
    FldTot = tb.Fields.Count - 1
    Call ProgBarInit(MeForm, progBar, recTot)
    tb.MoveFirst
    For recCnt = 1 To recTot
        recStr = ""
        Call UpdateProgress(MeForm, recCnt, progBar, "loading records to list view...")
        For FldCnt = 0 To FldTot
            rFld = StringRemNull(tb(FldCnt))
            If MvSearch(ShowDay, tb(FldCnt).Name, ",") > 0 Then
                rFld = Format$(rFld, "dd/mm/yyyy dddd")
            End If
            If MvSearch(AmountFlds, tb(FldCnt).Name, ",") > 0 Then
                rFld = ProperAmount(rFld)
            End If
            Select Case FldCnt
            Case FldTot
                recStr = StringsConcat(recStr, rFld)
            Case Else
                recStr = StringsConcat(recStr, rFld, KM)
            End Select
            Err.Clear
        Next
        Call StringParse(fldArray, recStr, KM)
        ' correlation exists for some of the values
        If IsMissing(varCorrelate) = False Then
            ' how many correlated columns are there
            varTot = UBound(varCorrelate)
            For varCnt = 1 To varTot
                ' read properties of  the correlation
                varStr = varCorrelate(varCnt)
                Call StringParse(varCo, varStr, ",")
                ' the correlation should meet requirements 1. columnpos, 2 table, 3 table key, 4 field to read
                ReDim Preserve varCo(4)
                tableP = varCo(1)
                tableN = varCo(2)
                tableK = varCo(3)
                tableF = varCo(4)
                curVal = fldArray(Val(tableP))
                ' check table names for particular table names
                Select Case LCase$(tableN)
                Case "priority": fldArray(Val(tableP)) = StringPriority(Val(curVal))
                Case "sensitivity": fldArray(Val(tableP)) = StringSensitivity(Val(curVal))
                Case "status": fldArray(Val(tableP)) = StringStatus(Val(curVal))
                Case "remind": fldArray(Val(tableP)) = StringFlag(Val(curVal))
                Case "confirmed": fldArray(Val(tableP)) = StringFlag(Val(curVal))
                Case "dealtwith": fldArray(Val(tableP)) = StringFlag(Val(curVal))
                Case "flag": fldArray(Val(tableP)) = StringFlag(Val(curVal))
                Case "deliver": fldArray(Val(tableP)) = StringFlag(Val(curVal))
                Case "branch": fldArray(Val(tableP)) = BranchString(curVal)
                Case "calltype": fldArray(Val(tableP)) = CallTypeString(Val(curVal))
                Case "sent": fldArray(Val(tableP)) = StringFlag(Val(curVal))
                Case "action": fldArray(Val(tableP)) = ActionName(Val(curVal))
                Case "rulefield": fldArray(Val(tableP)) = RuleFieldName(Val(curVal))
                Case Else
                    fldArray(Val(tableP)) = dbCorrelate(db, tableN, tableK, curVal, tableF)
                End Select
                Err.Clear
            Next
        End If
        If Len(TagFldName) > 0 Then
            tagVal = StringRemNull(tb.Fields(TagFldName))
        Else
            tagVal = ""
        End If
        If UpperCase = True Then
            ArrayToUpperCase fldArray
        End If
        lstPos = LstViewUpdate(fldArray, lstView, "")
        lstView.ListItems(lstPos).Tag = tagVal
        If Len(strImage) > 0 Then
            lstView.ListItems(lstPos).Icon = strImage
        End If
        If Len(strSelImage) > 0 Then
            lstView.ListItems(lstPos).SmallIcon = strSelImage
        End If
        tb.MoveNext
        Err.Clear
    Next
    tb.Close
    db.Close
    Set tb = Nothing
    Set db = Nothing
    ProgBarClose MeForm, progBar
    LstViewAutoResize lstView
    Err.Clear
End Sub
Function TreeViewRemoveChecked(MeForm As Form, progBar As ProgressBar, treeDms As TreeView, Optional ByVal NodeChecked As Boolean = False) As String
    On Error Resume Next
    Dim treeTot As Long
    Dim treeCnt As Long
    Dim treeChk As Boolean
    Dim strOut As String
    strOut = ""
    treeTot = treeDms.Nodes.Count
    Call ProgBarInit(MeForm, progBar, treeTot)
    For treeCnt = treeTot To 1 Step -1
        Call UpdateProgress(MeForm, treeCnt, progBar, "Removing tree items...")
        treeChk = treeDms.Nodes(treeCnt).Checked
        If treeChk <> NodeChecked Then
            GoTo nextLine
        End If
        strOut = strOut & treeDms.Nodes(treeCnt).Fullpath & vbNewLine
        treeDms.Nodes.Remove treeCnt
nextLine:
        Err.Clear
    Next
    ProgBarClose MeForm, progBar
    TreeViewRemoveChecked = StringRemoveDelim(strOut, vbNewLine)
    Err.Clear
End Function
Public Function TreeViewTopicPosition(MeForm As Form, progBar As ProgressBar, TreeD As TreeView, ByVal Fullpath As String) As Long
    On Error Resume Next
    Fullpath = Replace$(Fullpath, " : ", "\")
    TreeViewTopicPosition = TreeViewSearchPathPosNew(MeForm, progBar, TreeD, Fullpath)
    Err.Clear
End Function
Public Sub TreeViewCheckChildren(treeDms As TreeView, ByVal nodeParent As Long, Optional ByVal boolCheck As Boolean = False)
    On Error Resume Next
    Dim nodeChild As Node
    ' Get the parent node's first child
    Set nodeChild = treeDms.Nodes(nodeParent).Child
    ' Now walk through the current parent node's children
    Do While Not (nodeChild Is Nothing)
        ' If the current child node has it's own children...
        treeDms.Nodes(nodeChild.Index).Checked = boolCheck
        If nodeChild.Children Then
            ' Recursively call this proc searching for the target string.
            ' If found, return the index and exit.
            Call TreeViewCheckChildren(treeDms, nodeChild.Index, boolCheck)
        End If
        ' Get the current child node's next sibling
        Set nodeChild = nodeChild.Next
        Err.Clear
    Loop
    Set nodeChild = Nothing
    Err.Clear
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' PrintRTF - Prints the contents of a RichTextBox control using the
'            provided margins
'
' RTF - A RichTextBox control to print
'
' LeftMarginWidth - Width of desired left margin in twips
'
' TopMarginHeight - Height of desired top margin in twips
'
' RightMarginWidth - Width of desired right margin in twips
'
' BottomMarginHeight - Height of desired bottom margin in twips
'
' Notes - If you are also using WYSIWYG_RTF() on the provided RTF
'         parameter you should specify the same LeftMarginWidth and
'         RightMarginWidth that you used to call WYSIWYG_RTF()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub PrintRTF(RTF As RichTextBox, LeftMarginWidth As Long, TopMarginHeight As Long, RightMarginWidth As Long, BottomMarginHeight As Long)
    On Error Resume Next
    Dim LeftOffset As Long
    Dim TopOffset As Long
    Dim LeftMargin As Long
    Dim TopMargin As Long
    Dim RightMargin As Long
    Dim BottomMargin As Long
    Dim fr As FormatRange
    Dim rcDrawTo As RECT
    Dim rcPage As RECT
    Dim TextLength As Long
    Dim NextCharPosition As Long
    Dim R As Long
    ' Start a print job to get a valid Printer.hDC
    Printer.Print Space$(1)
    Printer.ScaleMode = vbTwips
    ' Get the offsett to the printable area on the page in twips
    LeftOffset = Printer.ScaleX(GetDeviceCaps(Printer.hdc, PHYSICALOFFSETX), vbPixels, vbTwips)
    TopOffset = Printer.ScaleY(GetDeviceCaps(Printer.hdc, PHYSICALOFFSETY), vbPixels, vbTwips)
    ' Calculate the Left, Top, Right, and Bottom margins
    LeftMargin = LeftMarginWidth - LeftOffset
    TopMargin = TopMarginHeight - TopOffset
    RightMargin = (Printer.Width - RightMarginWidth) - LeftOffset
    BottomMargin = (Printer.Height - BottomMarginHeight) - TopOffset
    ' Set printable area rect
    rcPage.Left = 0
    rcPage.Top = 0
    rcPage.Right = Printer.ScaleWidth
    rcPage.Bottom = Printer.ScaleHeight
    ' Set rect in which to print (relative to printable area)
    rcDrawTo.Left = LeftMargin
    rcDrawTo.Top = TopMargin
    rcDrawTo.Right = RightMargin
    rcDrawTo.Bottom = BottomMargin
    ' Set up the print instructions
    fr.hdc = Printer.hdc   ' Use the same DC for measuring and rendering
    fr.hdcTarget = Printer.hdc  ' Point at printer hDC
    fr.rc = rcDrawTo            ' Indicate the area on page to draw to
    fr.rcPage = rcPage          ' Indicate entire size of page
    fr.chrg.cpMin = 0           ' Indicate start of text through
    fr.chrg.cpMax = -1          ' end of the text
    ' Get length of text in RTF
    TextLength = Len(RTF.Text)
    ' Loop printing each page until done
    Do
        ' Print the page by sending EM_FORMATRANGE message
        NextCharPosition = SendMessage(RTF.hWnd, EM_FORMATRANGE, True, fr)
        If NextCharPosition >= TextLength Then
            'If done then exit
            Exit Do
        End If
        fr.chrg.cpMin = NextCharPosition ' Starting position for next page
        Printer.NewPage                  ' Move on to next page
        Printer.Print Space$(1) ' Re-initialize hDC
        fr.hdc = Printer.hdc
        fr.hdcTarget = Printer.hdc
        Err.Clear
    Loop
    ' Commit the print job
    Printer.EndDoc
    ' Allow the RTF to free up memory
    R = SendMessage(RTF.hWnd, EM_FORMATRANGE, False, ByVal CLng(0))
    Err.Clear
End Sub
Public Sub LstBoxClearAPI(lstBox As Variant)
    On Error Resume Next
    Select Case TypeName(lstBox)
    Case "ListBox"
        Call SendMessage(lstBox.hWnd, LB_RESETCONTENT, 0, ByVal 0&)
    Case "ComboBox"
        Call SendMessage(lstBox.hWnd, CB_RESETCONTENT, 0, ByVal 0&)
    End Select
    Err.Clear
End Sub
Public Function StringProperCase(ByVal StrString As String, Optional Delim As String = "\") As String
    On Error Resume Next
    Dim spItems() As String
    Dim spSubs() As String
    Dim spTot As Long
    Dim spCnt As Long
    StrString = Trim$(StrString)
    spTot = StringParse(spItems, StrString, Delim)
    For spCnt = 1 To spTot
        spItems(spCnt) = StrConv(spItems(spCnt), vbProperCase)
        Err.Clear
    Next
    StringProperCase = MvFromArray(spItems, Delim)
    Erase spItems
    Erase spSubs
    Err.Clear
End Function
Function TreeViewAddPath(TreeV As TreeView, ByVal sPath As String, Optional ByVal Image As String = "", Optional ByVal SelectedImage As String = "", Optional ByVal Tag As String = "", Optional Delim As String = "\") As Long
    On Error Resume Next
    Dim prevP As String
    Dim currP As String
    Dim lngC As Long
    Dim lngT As Long
    Dim pStr() As String
    Dim currN As String
    Dim nodeN As Node
    Dim pKey As String
    Dim cLoc As Long
    sPath = StringProperCase(sPath)
    Call StringParse(pStr, sPath, Delim)
    lngT = UBound(pStr)
    For lngC = 1 To lngT
        prevP = MvFromMv(sPath, 1, lngC - 1, Delim)
        currP = MvFromMv(sPath, 1, lngC, Delim)
        currN = pStr(lngC)
        If prevP = "" Then
            ' this is the root node
            cLoc = TreeViewPathLocation(TreeV, currP)
            If cLoc = 0 Then
                Set nodeN = TreeV.Nodes.Add(, , currP, StringProperCase(currP))
                If Len(Image) > 0 Then nodeN.Image = Image
                If Len(SelectedImage) > 0 Then nodeN.SelectedImage = SelectedImage
                nodeN.Tag = Tag
            Else
                Set nodeN = TreeV.Nodes(cLoc)
            End If
        Else
            ' this is the second, third etc node
            cLoc = TreeViewPathLocation(TreeV, currP)
            If cLoc = 0 Then
                Set nodeN = TreeV.Nodes.Add(pKey, tvwChild, currP, StringProperCase(currN))
                If Len(Image) > 0 Then nodeN.Image = Image
                If Len(SelectedImage) > 0 Then nodeN.SelectedImage = SelectedImage
                nodeN.Tag = Tag
            Else
                Set nodeN = TreeV.Nodes(cLoc)
            End If
        End If
        pKey = nodeN.Key
        If lngC = lngT Then
            TreeViewAddPath = nodeN.Index
            Exit For
        End If
        Err.Clear
    Next
    Err.Clear
End Function
Public Function dbNextOpenSequence(ByVal Dbase As DAO.Database, ByVal TbName As String, ByVal KeyName As String) As String
    On Error Resume Next
    Dim lngStart As Long
    Dim tb As DAO.Recordset
    Set tb = Dbase.OpenRecordset("select max(" & KeyName & ") as Sequence from [" & TbName & "];", dbOpenDynaset, dbSeeChanges, dbPessimistic)
    With tb
        .MoveLast
        lngStart = Val(StringRemNull(!Sequence)) + 1
    End With
    tb.Close
    Set tb = Nothing
    dbNextOpenSequence = CStr(lngStart)
    Err.Clear
End Function
Public Function boolIsBlank(ObjectName As Variant, ByVal FldName As String) As Boolean
    On Error Resume Next
    Dim StrM As String
    Dim StrT As String
    Dim strO As String
    Dim strK As String
    boolIsBlank = False
    If TypeOf ObjectName Is TextBox Then
        If Len(Trim$(ObjectName.Text)) = 0 Then
            strO = "type"
            GoSub CompileError
            Call intError(StrT, StrM)
            boolIsBlank = True
        End If
    ElseIf TypeOf ObjectName Is ComboBox Then
        If Len(Trim$(ObjectName.Text)) = 0 Then
            strO = "select"
            GoSub CompileError
            Call intError(StrT, StrM)
            boolIsBlank = True
        End If
    ElseIf TypeOf ObjectName Is ImageCombo Then
        If Len(Trim$(ObjectName.Text)) = 0 Then
            strO = "select"
            GoSub CompileError
            Call intError(StrT, StrM)
            boolIsBlank = True
        End If
    ElseIf TypeOf ObjectName Is CheckBox Then
        If ObjectName.Value = 0 Then
            strO = "select"
            GoSub CompileError
            Call intError(StrT, StrM)
            boolIsBlank = True
        End If
    ElseIf TypeOf ObjectName Is ListBox Then
        If (ObjectName.ListCount - 1) = -1 Then
            strO = "select"
            GoSub CompileError
            Call intError(StrT, StrM)
            boolIsBlank = True
        End If
    ElseIf TypeOf ObjectName Is OptionButton Then
        If ObjectName.Value = False Then
            strO = "select"
            GoSub CompileError
            Call intError(StrT, StrM)
            boolIsBlank = True
        End If
    ElseIf TypeOf ObjectName Is Label Then
        If Len(Trim$(ObjectName.Caption)) = 0 Then
            strO = "specify"
            GoSub CompileError
            Call intError(StrT, StrM)
            boolIsBlank = True
        End If
    ElseIf TypeOf ObjectName Is MaskEdBox Then
        strK = ObjectName.Mask
        strK = Replace$(strK, "#", "_")
        If Trim$(ObjectName.Text = strK) Then
            strO = "enter"
            GoSub CompileError
            Call intError(StrT, StrM)
            boolIsBlank = True
        End If
    End If
    Err.Clear
    Exit Function
CompileError:
    StrM = "The " & LCase$(FldName) & " cannot be left blank. Please " & strO & " the " & LCase$(FldName) & "."
    StrT = StringProperCase(FldName & " error")
    Err.Clear
    Return
    Err.Clear
End Function
Public Function ReadIniApi(ByVal sKey As String, Optional ByVal sSection As String = "account", Optional ByVal sIniFile As String = "") As String
    On Error Resume Next
    Dim retStr As String
    Dim RetLng As Long
    retStr = String$(255, Chr$(0))
    If Len(sIniFile) = 0 Then
        sIniFile = StringsConcat(App.Path, "\", AppTitle, ".ini")
    End If
    RetLng = GetPrivateProfileString(sSection, sKey, "", retStr, 255, sIniFile)
    ReadIniApi = Left$(retStr, RetLng)
    Err.Clear
End Function
Public Sub SaveIniApi(ByVal sKey As String, ByVal sValue As String, Optional ByVal sSection As String = "account", Optional ByVal sIniFile As String = "")
    On Error Resume Next
    Dim RetLng As Long
    If Len(sIniFile) = 0 Then
        sIniFile = StringsConcat(App.Path, "\", AppTitle, ".ini")
    End If
    sValue = Replace$(sValue, vbCr, vbNullChar)
    sValue = Replace$(sValue, vbLf, vbNullChar)
    sSection = Replace$(sSection, VM, "\")
    RetLng = WritePrivateProfileString(sSection, sKey, sValue, sIniFile)
    Err.Clear
End Sub
Public Function StringRemNull(Fromvariable As DAO.Field, Optional ClearQuotes As Boolean = False) As String
    On Error Resume Next
    Dim sResult As String
    sResult = Trim$(Fromvariable.Value & "")
    If ClearQuotes = True Then
        sResult = Trim$(Replace$(sResult, Quote, ""))
    End If
    StringRemNull = sResult
    Err.Clear
End Function
Public Sub FileUpdate(ByVal filName As String, ByVal filLines As String, Optional ByVal Wora As String = "")
    On Error Resume Next
    Dim iFileNum As Integer
    Wora = UCase$(Trim$(Wora))
    If Len(Wora) = 0 Then
        Wora = "W"
    End If
    iFileNum = FreeFile
    Select Case Wora
    Case "W"
        Open filName For Output As #iFileNum
        Case "A"
            Open filName For Append As #iFileNum
            End Select
            Print #iFileNum, filLines
        Close #iFileNum
        Err.Clear
End Sub
Function MvRemoveDuplicates(ByVal StrMvString As String, Optional ByVal Delim As String = ";") As String
    On Error Resume Next
    Dim spData() As String
    Dim spTot As Long
    Dim spCnt As Long
    Dim xCol As New Collection
    Call StringParse(spData, StrMvString, Delim)
    spTot = UBound(spData)
    For spCnt = 1 To spTot
        xCol.Add spData(spCnt), spData(spCnt)
        Err.Clear
    Next
    MvRemoveDuplicates = MvFromCollection(xCol, Delim)
    Err.Clear
End Function
Public Function MvCount(ByVal StringMv As String, Optional ByVal Delim As String = "") As Long
    On Error Resume Next
    Dim xNew() As String
    If Len(Delim) = 0 Then
        Delim = VM
    End If
    Call StringParse(xNew, StringMv, Delim)
    MvCount = UBound(xNew)
    Err.Clear
End Function
Public Function StringRemoveDelim(ByVal strData As String, Optional ByVal Delim As String = "") As String
    On Error Resume Next
    Dim intDataSize As Long
    Dim intDelimSize As Long
    Dim strLast As String
    If Len(Delim) = 0 Then Delim = VM
    intDataSize = Len(strData)
    intDelimSize = Len(Delim)
    strLast = Right$(strData, intDelimSize)
    Select Case LCase$(strLast)
    Case LCase$(Delim)
        StringRemoveDelim = Left$(strData, (intDataSize - intDelimSize))
    Case Else
        StringRemoveDelim = strData
    End Select
    Err.Clear
End Function
Sub CreateNestedDirectory(ByVal StrCompletePath As String)
    On Error Resume Next
    Dim spPaths() As String
    Dim spTot As Long
    Dim spCnt As Long
    Dim curPath As String
    Call StringParse(spPaths, StrCompletePath, "\")
    spTot = UBound(spPaths)
    For spCnt = 1 To spTot
        curPath = MvFromMv(StrCompletePath, 1, spCnt, "\")
        If boolDirExists(curPath) = False Then
            MkDir curPath
        End If
        Err.Clear
    Next
    Err.Clear
End Sub
Function MyPrompt(ByVal StrMsg As String, Optional ByVal strButton As String = "o", Optional ByVal StrIcon As String = "e", Optional ByVal StrHeading As String = "") As Long
    On Error Resume Next
    ' button can be any of
    ' ync - yesnocancel, c - cancel, o - ok, oc - okcancel, rc - retrycancel and yn - yesno
    ' and ari - abortretryignore, bc - backclose, bnc - backnextclose
    ' bns - backnextsnooze, nc - nextclose, sc - searchclose, toc - tipsoptionsclose, yanc - yesallnocancel
    ' icon can be any of
    ' i - information, w - warning, c - critical, t - tip, q - query
    ' mode can be any of
    ' ad - autodown, ma - modal, me - modeless
    Dim isCheck As Long
    Dim Mode As Long
    Dim Button As Long
    Dim Icon As Long
    ' see if excel is already running
    If Len(StrHeading) = 0 Then
        StrHeading = AppTitle
    End If
    isCheck = 0
    Mode = vbApplicationModal
    Select Case LCase$(strButton)
    Case "ync"
        Button = vbYesNoCancel
    Case "c"
        Button = vbCancel
    Case "o"
        Button = vbOKOnly
    Case "oc"
        Button = vbOKCancel
    Case "rc"
        Button = vbRetryCancel
    Case "yn"
        Button = vbYesNo
    Case "ari"
        Button = vbAbortRetryIgnore
    End Select
    Select Case LCase$(StrIcon)
    Case "i", "t"
        Icon = vbInformation
    Case "w", "e"
        Icon = vbExclamation
    Case "c"
        Icon = vbCritical
    Case "q"
        Icon = vbQuestion
    End Select
    MyPrompt = MsgBox(StrMsg, Button + Icon + Mode, StrHeading)
    Err.Clear
End Function
Private Sub TreeViewWriteToTable(MeForm As Form, progBar As ProgressBar, mRs As DAO.Recordset, TreeV As TreeView, Optional ByVal TbType As String = "n")
    On Error Resume Next
    'Writes the Node information from the TreeView into a table
    Dim iTmp As Long
    Dim iIndex As Long
    mnIndex = 1
    ProgBarInit MeForm, progBar, TreeV.Nodes.Count
    'get the index of the root node that is at the top of the treeview
    iIndex = TreeV.Nodes(mnIndex).FirstSibling.Index
    iTmp = iIndex
    cntNodes = 1
    mRs.AddNew
    mRs!Parent = 0               'this is a root node
    mRs!Key = TreeV.Nodes(iIndex).Index
    If LCase$(TbType) = "n" Then
        mRs!Text = Left$(TreeV.Nodes(iIndex).Text, 255)
    Else
        mRs!Text = TreeV.Nodes(iIndex).Text
    End If
    mRs!Sequence = cntNodes
    mRs!Image = TreeV.Nodes(iIndex).Image
    mRs!SelectedImage = TreeV.Nodes(iIndex).SelectedImage
    mRs!Tag = TreeV.Nodes(iIndex).Tag
    mRs!Checked = IIf((TreeV.Nodes(iIndex).Checked = True), 1, 0)
    mRs!Bold = IIf((TreeV.Nodes(iIndex).Bold = True), 1, 0)
    mRs!ForeColor = TreeV.Nodes(iIndex).ForeColor
    mRs!Fullpath = TreeV.Nodes(iIndex).Fullpath
    mRs.Update
    'If the Node has Children call the sub that writes the children
    If TreeV.Nodes(iIndex).Children > 0 Then
        TreeViewWriteChild MeForm, progBar, iIndex, mRs, TreeV, TbType
    End If
    Do While iIndex <> TreeV.Nodes(iTmp).LastSibling.Index
        'loop through all the root nodes
        cntNodes = cntNodes + 1
        Call UpdateProgress(MeForm, cntNodes, progBar, "Saving tree view to table...")
        mRs.AddNew
        mRs!Parent = 0                           'this is a root node
        mRs!Key = TreeV.Nodes(iIndex).Next.Index
        If LCase$(TbType) = "n" Then
            mRs!Text = Left$(TreeV.Nodes(iIndex).Next.Text, 255)
        Else
            mRs!Text = TreeV.Nodes(iIndex).Next.Text
        End If
        mRs!Sequence = cntNodes
        mRs!Image = TreeV.Nodes(iIndex).Next.Image
        mRs!SelectedImage = TreeV.Nodes(iIndex).Next.SelectedImage
        mRs!Tag = TreeV.Nodes(iIndex).Next.Tag
        mRs!Checked = IIf((TreeV.Nodes(iIndex).Next.Checked = True), 1, 0)
        mRs!Bold = IIf((TreeV.Nodes(iIndex).Next.Bold = True), 1, 0)
        mRs!ForeColor = TreeV.Nodes(iIndex).Next.ForeColor
        mRs!Fullpath = TreeV.Nodes(iIndex).Next.Fullpath
        mRs.Update
        'If the Node has Children call the sub that writes the children
        If TreeV.Nodes(iIndex).Next.Children > 0 Then
            TreeViewWriteChild MeForm, progBar, TreeV.Nodes(iIndex).Next.Index, mRs, TreeV, TbType
        End If
        ' Move to the Next root Node
        iIndex = TreeV.Nodes(iIndex).Next.Index
        Err.Clear
    Loop
    ProgBarClose MeForm, progBar
    Err.Clear
End Sub
Public Function StringIconv(ByVal sValue As String, Optional ByVal sFormat As String = "") As String
    On Error Resume Next
    Dim sRslt As String
    Dim i As Long
    Dim ch As String
    Dim L As Long
    Dim sN As String
    sRslt = sValue
    Select Case UCase$(sFormat)
    Case ""
        sRslt = Replace$(sRslt, ",", "")
        sRslt = Replace$(sRslt, "/", "")
        sRslt = Replace$(sRslt, ".", "")
        sRslt = Replace$(sRslt, "(", "")
        sRslt = Replace$(sRslt, ")", "")
        sRslt = Replace$(sRslt, "~", "")
        sRslt = Replace$(sRslt, "!", "")
        sRslt = Replace$(sRslt, "@", "")
        sRslt = Replace$(sRslt, "#", "")
        sRslt = Replace$(sRslt, "$", "")
        sRslt = Replace$(sRslt, "%", "")
        sRslt = Replace$(sRslt, "^", "")
        sRslt = Replace$(sRslt, "&", "")
        sRslt = Replace$(sRslt, "*", "")
        sRslt = Replace$(sRslt, "_", "")
        sRslt = Replace$(sRslt, "-", "")
        sRslt = Replace$(sRslt, "=", "")
        sRslt = Replace$(sRslt, "|", "")
        sRslt = Replace$(sRslt, "\", "")
        sRslt = Replace$(sRslt, ":", "")
        sRslt = Replace$(sRslt, ";", "")
        sRslt = Replace$(sRslt, "<", "")
        sRslt = Replace$(sRslt, ">", "")
        sRslt = Replace$(sRslt, "?", "")
        sRslt = Replace$(sRslt, "/", "")
        sRslt = Replace$(sRslt, "'", "")
        sRslt = Replace$(sRslt, "`", "")
        sRslt = Replace$(sRslt, "+", "")
        sRslt = Replace$(sRslt, "{", "")
        sRslt = Replace$(sRslt, "}", "")
        sRslt = Replace$(sRslt, "[", "")
        sRslt = Replace$(sRslt, "]", "")
        sRslt = Replace$(sRslt, Quote, "")
    Case "Q"
        sRslt = Replace$(sRslt, "''", "")
        sRslt = Replace$(sRslt, "'", "")
    Case "F"
        sRslt = Replace$(sRslt, "/", "%")
        sRslt = Replace$(sRslt, "\", "%")
        sRslt = Replace$(sRslt, "|", "%")
    Case "C"
        sRslt = Replace$(sRslt, ",", "")
    Case "M"
        sRslt = Replace$(sRslt, ",", "")
        sRslt = Replace$(sRslt, ".", "")
    Case "S"
        L = Len(sRslt)
        sRslt = sRslt
        If L = 0 Then
            Err.Clear
            Exit Function
        End If
        sN = ""
        For i = 1 To L
            ch = Mid$(sRslt, i, 1)
            If ch = " " Then
                sN = StringConcat(sN, ch)
            End If
            If ch >= "a" Then
                If ch <= "z" Then
                    sN = StringConcat(sN, ch)
                End If
            End If
            If ch >= "A" Then
                If ch <= "Z" Then
                    sN = StringConcat(sN, ch)
                End If
            End If
            Err.Clear
        Next
        sRslt = sN
    Case "T"
        sRslt = Replace$(sRslt, "!", "")
        sRslt = Replace$(sRslt, "[", "")
        sRslt = Replace$(sRslt, "]", "")
        sRslt = Replace$(sRslt, ".", "")
        sRslt = Replace$(sRslt, Quote, "")
        sRslt = Replace$(sRslt, "`", "")
        sRslt = Replace$(sRslt, "'", "")
        sRslt = Replace$(sRslt, ",", "")
    End Select
    StringIconv = sRslt
    Err.Clear
End Function
Function MvRemoveBlanks(ByVal strValue As String, Optional ByVal Delim As String = "") As String
    On Error Resume Next
    Dim xData() As String
    Dim xTot As Long
    Dim xCnt As Long
    Dim xRslt As String
    If Len(Delim) = 0 Then
        Delim = VM
    End If
    xRslt = ""
    Call StringParse(xData, strValue, Delim)
    xTot = UBound(xData)
    For xCnt = 1 To xTot
        If Len(Trim$(xData(xCnt))) > 0 Then
            xRslt = StringsConcat(xRslt, xData(xCnt), Delim)
        End If
        Err.Clear
    Next
    xRslt = StringRemoveDelim(xRslt, Delim)
    MvRemoveBlanks = xRslt
    Err.Clear
End Function
Public Function StringParse(retarray() As String, ByVal strText As String, Optional ByVal Delim As String = "") As Long
    On Error Resume Next
    Dim varArray() As String
    Dim varCnt As Long
    Dim VarS As Long
    Dim VarE As Long
    Dim VarA As Long
    If Len(Delim) = 0 Then
        Delim = Chr$(253)
    End If
    varArray = Split(strText, Delim)
    VarS = LBound(varArray)
    VarE = UBound(varArray)
    VarA = VarE + 1
    ReDim retarray(VarA)
    For varCnt = VarS To VarE
        VarA = varCnt + 1
        retarray(VarA) = varArray(varCnt)
        Err.Clear
    Next
    StringParse = UBound(retarray)
    Err.Clear
End Function
Sub ArrayTrimItems(varArray() As String)
    On Error Resume Next
    Dim uArray As Long
    Dim cArray As Long
    Dim lArray As Long
    uArray = UBound(varArray)
    lArray = LBound(varArray)
    For cArray = lArray To uArray
        varArray(cArray) = Trim$(varArray(cArray))
        Err.Clear
    Next
    Err.Clear
End Sub
Public Function dbType(ByVal StrType As String) As Integer
    On Error Resume Next
    Dim StrTp As String
    StrTp = LCase$(Trim$(StrType))
    Select Case StrTp
    Case "big", "bigint": dbType = dbBigInt
    Case "bi", "bin", "binary": dbType = dbLongBinary
    Case "cha", "char": dbType = dbChar
    Case "dec", "decimal": dbType = dbDecimal
    Case "flo", "float": dbType = dbFloat
    Case "gui", "guid": dbType = dbGUID
    Case "tim", "time": dbType = dbTime
    Case "tis", "timestamp": dbType = dbTimeStamp
    Case "num", "numeric": dbType = dbNumeric
    Case "var", "varbinary": dbType = dbVarBinary
    Case "bo", "boo", "boolean": dbType = dbBoolean
    Case "by", "byt", "byte": dbType = dbByte
    Case "in", "int", "integer": dbType = dbInteger
    Case "lo", "lon", "long": dbType = dbLong
    Case "cu", "cur", "currency": dbType = dbCurrency
    Case "si", "sin", "single": dbType = dbSingle
    Case "do", "dou", "double": dbType = dbDouble
    Case "da", "dat", "date": dbType = dbDate
    Case "te", "tex", "text": dbType = dbText
    Case "lob", "longbinary", "long binary": dbType = dbLongBinary
    Case "ole", "object": dbType = dbLongBinary
    Case "me", "mem", "memo": dbType = dbMemo
    Case "st", "str", "string": dbType = dbText
    End Select
    Err.Clear
End Function
Public Function MvSearch(ByVal StringMv As String, ByVal StrLookFor As String, Optional ByVal Delim As String = "", Optional TrimItems As Boolean = False) As Long
    On Error Resume Next
    Dim TheFields() As String
    MvSearch = 0
    If Len(StringMv) = 0 Then
        MvSearch = 0
        Err.Clear
        Exit Function
    End If
    If Len(Delim) = 0 Then
        Delim = Chr$(253)
    End If
    Call StringParse(TheFields, StringMv, Delim)
    If TrimItems = True Then ArrayTrimItems TheFields
    MvSearch = ArraySearch(TheFields, StrLookFor)
    Err.Clear
End Function
Public Sub TreeViewClearAPI(TreeV As TreeView)
    On Error Resume Next
    Dim lNodeHandle As Long
    Dim tvHwnd As Long
    tvHwnd = TreeV.hWnd
    ' Turn off redrawing on the Treeview for more speed improvements
    Do
        lNodeHandle = SendMessageLONG(tvHwnd, TVM_GETNEXTITEM, TVGN_ROOT, 0&)
        If lNodeHandle > 0 Then
            SendMessageLONG tvHwnd, TVM_DELETEITEM, 0, lNodeHandle
        Else
            Exit Do
        End If
        Err.Clear
    Loop
    Err.Clear
End Sub
Public Sub ApplicationOnTop(ByVal wHandle As Long)
    On Error Resume Next
    IsOnTop = SetWindowPos(wHandle, HWND_TOPMOST, 0, 0, 0, 0, WindowFlags)
    Err.Clear
End Sub
Public Sub Pause(ByVal nSecond As Double)
    On Error Resume Next
    ' call pause(2)      delay for 2 seconds
    Dim t0 As Double
    t0 = Timer
    Do While Timer - t0 < nSecond
        DoEvents
        ' if we cross midnight, back up one day
        If Timer < t0 Then
            t0 = t0 - CLng(24) * CLng(60) * CLng(60)
        End If
        Err.Clear
    Loop
    Err.Clear
End Sub
Public Sub TreeViewGetChildren(tvw As TreeView, nodN As Node, nodP As Node)
    On Error Resume Next
    Dim nodC As Node
    Dim nodT As Node
    Dim i As Integer
    Dim i_Tot As Integer
    With tvw
        'For each children in the tree
        i_Tot = nodN.Children
        For i = 1 To i_Tot
            'If it's the first child:
            If i = 1 Then
                'Add the node:
                Set nodC = .Nodes.Add(nodP.Index, tvwChild, , nodN.Child.Text, nodN.Child.Image, nodN.Child.SelectedImage)
                nodC.Tag = nodN.Child.Tag
                'Set us up for the next child:
                Set nodT = nodN.Child.Next
                'Get the added nodes children:
                If nodN.Child.Children <> 0 Then
                    TreeViewGetChildren tvw, nodN.Child, nodC
                End If
                'It's not the first child:
            Else
                'Add the node:
                Set nodC = .Nodes.Add(nodP.Index, tvwChild, , nodT.Text, nodT.Image, nodT.SelectedImage)
                nodC.Tag = nodT.Tag
                'Get the added nodes children:
                If nodT.Children <> 0 Then
                    TreeViewGetChildren tvw, nodT, nodC
                End If
                'Set us up again:
                Set nodT = nodT.Next
            End If
            Err.Clear
        Next
    End With
    Err.Clear
End Sub
Public Sub LstViewMakeHeadings(lstView As ListView, ByVal strHeads As String, Optional ByVal ClearItems As String = "")
    On Error Resume Next
    Dim FldCnt As Integer
    Dim FldHead() As String
    Dim FldTot As Integer
    Dim ColX As ColumnHeader
    Dim cPos As Long
    Call StringParse(FldHead, strHeads, ",")
    FldTot = UBound(FldHead)
    lstView.ColumnHeaders.Clear
    If ClearItems = "" Then
        lstView.ListItems.Clear
    End If
    ' first column should be left aligned
    Set ColX = lstView.ColumnHeaders.Add(, , StringProperCase(FldHead(1)), 1440)
    For FldCnt = 2 To FldTot
        With lstView.ColumnHeaders
            FldHead(FldCnt) = StringProperCase(FldHead(FldCnt))
            cPos = ArraySearch(ViewHeadings, FldHead(FldCnt))
            Select Case cPos
            Case 0
                Set ColX = .Add(, , FldHead(FldCnt), 1440)
            Case Else
                Set ColX = .Add(, , FldHead(FldCnt), 1440, vbRightJustify)
            End Select
        End With
        Err.Clear
    Next
    lstView.View = lvwReport
    lstView.Checkboxes = True
    lstView.GridLines = True
    lstView.FullRowSelect = True
    lstView.Refresh
    Err.Clear
End Sub
Public Function TbFldNames(ByVal dbRs As DAO.Recordset, Optional ByVal Delim As String = ",") As String
    On Error Resume Next
    Dim fL As String
    Dim fC As Integer
    Dim fT As Integer
    Dim fN As String
    fL = ""
    If Len(Delim) = 0 Then
        Delim = ","
    End If
    fT = dbRs.Fields.Count - 1
    For fC = 0 To fT
        fN = dbRs.Fields(fC).Name
        fL = StringsConcat(fL, fN, Delim)
        Err.Clear
    Next
    TbFldNames = StringRemoveDelim(fL, Delim)
    Err.Clear
End Function
Function NoCommas(ByVal strValue As String) As String
    On Error Resume Next
    NoCommas = Replace$(strValue, ",", "")
    Err.Clear
End Function
Function ProperAmount(ByVal strValue As String) As String
    On Error Resume Next
    strValue = MakeMoney(strValue)
    ProperAmount = NoCommas(strValue)
    Err.Clear
End Function
Public Function MakeMoney(ByVal strValue As String) As String
    On Error Resume Next
    strValue = Trim$(strValue)
    If Len(strValue) = 0 Then strValue = "0"
    MakeMoney = Format$(strValue, "###,###,###,###,###,###,###,###,###,###,##0.00")
    Err.Clear
End Function
Public Function StringsConcat(ParamArray Items()) As String
    On Error Resume Next
    Dim Item As Variant
    Dim NewString As String
    NewString = ""
    For Each Item In Items
        NewString = StringConcat(NewString, CStr(Item))
        Err.Clear
    Next
    StringsConcat = NewString
    Set Item = Nothing
    Err.Clear
End Function
Function StringPriority(ByVal valCur As Integer) As String
    On Error Resume Next
    Select Case valCur
    Case 0: StringPriority = "Low"
    Case 1: StringPriority = "Normal"
    Case 2: StringPriority = "High"
    End Select
    Err.Clear
End Function
Function StringSensitivity(ByVal valCur As Integer) As String
    On Error Resume Next
    Select Case valCur
    Case 0: StringSensitivity = "Normal"
    Case 1: StringSensitivity = "Personal"
    Case 2: StringSensitivity = "Private"
    Case 3: StringSensitivity = "Confidential"
    End Select
    Err.Clear
End Function
Function StringStatus(ByVal valCur As Integer) As String
    On Error Resume Next
    Select Case valCur
    Case 0: StringStatus = "Not Started"
    Case 1: StringStatus = "In Progress"
    Case 2: StringStatus = "Completed"
    Case 3: StringStatus = "Waiting On Someone"
    Case 4: StringStatus = "Deferred"
    End Select
    Err.Clear
End Function
Function StringFlag(ByVal intValue As Integer) As String
    On Error Resume Next
    StringFlag = IIf((intValue = 0), "N", "Y")
    Err.Clear
End Function
Public Function BranchString(strValue As String) As String
    On Error Resume Next
    BranchString = IIf(LCase$(strValue) = "a", "Other", "Branch")
    Err.Clear
End Function
Function CallTypeString(optCallType As Integer) As String
    On Error Resume Next
    Select Case optCallType
    Case 0: CallTypeString = "Telephoned"
    Case 1: CallTypeString = "Please Phone Back"
    Case 2: CallTypeString = "Returned Your Call"
    Case 3: CallTypeString = "Called To See You"
    Case 4: CallTypeString = "Desires Appointment"
    Case 5: CallTypeString = "Will Phone Back"
    End Select
    Err.Clear
End Function
Function ActionName(ByVal intValue As Integer) As String
    On Error Resume Next
    Select Case intValue
    Case 0:        ActionName = "Erase"
    Case 1:        ActionName = "Rename"
    Case 2:        ActionName = "Replace"
    Case 3:        ActionName = "Split"
    End Select
    Err.Clear
End Function
Function RuleFieldName(ByVal intValue As Integer) As String
    On Error Resume Next
    Select Case intValue
    Case 0:        RuleFieldName = "In From"
    Case 1:        RuleFieldName = "Out To"
    Case 2:        RuleFieldName = "In From,Out To"
    End Select
    Err.Clear
End Function
Public Function dbCorrelate(ByVal Ddatabase As DAO.Database, ByVal TableName As String, ByVal TableKey As String, ByVal ValuetoSeek As String, ByVal Fieldtoread As String) As String
    On Error Resume Next
    Dim MaDbTable As DAO.Recordset
    Set MaDbTable = Ddatabase.OpenRecordset(TableName)
    dbCorrelate = ""
    With MaDbTable
        .Index = TableKey
        .Seek "=", ValuetoSeek
        If .NoMatch = False Then
            dbCorrelate = StringRemNull(.Fields(Fieldtoread))
        End If
    End With
    MaDbTable.Close
    Set MaDbTable = Nothing
    Err.Clear
End Function
Public Sub ArrayToUpperCase(varArray As Variant)
    On Error Resume Next
    Dim aTot As Long
    Dim aCnt As Long
    Dim cCnt As Long
    aCnt = LBound(varArray)
    aTot = UBound(varArray)
    For cCnt = aCnt To aTot
        varArray(cCnt) = UCase$(varArray(cCnt))
        Err.Clear
    Next
    Err.Clear
End Sub
Public Function LstViewUpdate(Arrfields() As String, lstView As ListView, Optional ByVal lstIndex As String = "") As Long
    On Error Resume Next
    Dim itmX As ListItem
    Dim FldCnt As Integer
    Dim sStr As String
    Dim wCnt As Integer
    sStr = CStr(Val(lstIndex))
    Select Case sStr
    Case "0"
        Set itmX = lstView.ListItems.Add()
    Case Else
        Set itmX = lstView.ListItems(Val(lstIndex))
    End Select
    wCnt = UBound(Arrfields) - 1
    With itmX
        .Text = Arrfields(1)
        For FldCnt = 1 To wCnt
            .SubItems(FldCnt) = Arrfields(FldCnt + 1)
            Err.Clear
        Next
    End With
    LstViewUpdate = itmX.Index
    Set itmX = Nothing
    Err.Clear
End Function
Public Sub LstViewAutoResize(lstView As ListView)
    On Error Resume Next
    'Size each column based on the maximum of
    'EITHER the columnheader text width, or,
    'if the items below it are wider, the
    'widest list item in the column
    Dim col2adjust As Long
    Dim col2adjust_Tot As Long
    If lstView.ListItems.Count = 0 Then
        Err.Clear
        Exit Sub
    End If
    col2adjust_Tot = lstView.ColumnHeaders.Count - 1
    For col2adjust = 0 To col2adjust_Tot
        Call SendMessage(lstView.hWnd, LVM_SETCOLUMNWIDTH, col2adjust, ByVal LVSCW_AUTOSIZE_USEHEADER)
        Err.Clear
    Next
    LstViewResizeMax lstView
    Err.Clear
End Sub
Public Function TreeViewSearchPathPosNew(MeForm As Form, progBar As ProgressBar, TreeV As TreeView, ByVal StrSearch As String, Optional ByVal TreeRefresh As Boolean = False) As Long
    On Error Resume Next
    StrSearch = UCase$(Trim$(StrSearch))
    Dim pNode As Node
    Set pNode = TreeV.Nodes(StrSearch)
    If Err.Number = 35601 Then
        ' only search when current path is not found
        If TreeRefresh = True Then
            TreeViewRefreshKeys MeForm, progBar, TreeV, False
            Set pNode = TreeV.Nodes(StrSearch)
            If Err.Number = 35601 Then
                TreeViewSearchPathPosNew = 0
            Else
                TreeViewSearchPathPosNew = pNode.Index
            End If
        Else
            TreeViewSearchPathPosNew = 0
        End If
    Else
        TreeViewSearchPathPosNew = pNode.Index
    End If
    Err.Clear
End Function
Public Function MvFromArray(vArray() As String, Optional ByVal Delim As String = "", Optional StartingAt As Long = 1, Optional TrimItem As Boolean = True) As String
    On Error Resume Next
    If Len(Delim) = 0 Then Delim = VM
    Dim i As Long
    Dim BldStr As String
    Dim strL As String
    Dim TotArray As Long
    TotArray = UBound(vArray)
    For i = StartingAt To TotArray
        strL = vArray(i)
        If TrimItem = True Then strL = Trim$(strL)
        BldStr = BldStr & strL & Delim
        Err.Clear
    Next
    MvFromArray = StringRemoveDelim(BldStr, Delim)
    Err.Clear
End Function
Public Function MvFromMv(ByVal strOriginalMv As String, ByVal startPos As Long, Optional ByVal NumOfItems As Long = -1, Optional ByVal Delim As String = "") As String
    On Error Resume Next
    Dim sporiginal() As String
    Dim spTot As Long
    Dim spCnt As Long
    Dim sLine As String
    Dim endPos As Long
    sLine = ""
    If Len(Delim) = 0 Then
        Delim = VM
    End If
    Call StringParse(sporiginal, strOriginalMv, Delim)
    spTot = UBound(sporiginal)
    If NumOfItems = -1 Then
        endPos = spTot
    Else
        endPos = (startPos + NumOfItems) - 1
    End If
    For spCnt = startPos To endPos
        If spCnt = endPos Then
            sLine = StringsConcat(sLine, sporiginal(spCnt))
        Else
            sLine = StringsConcat(sLine, sporiginal(spCnt), Delim)
        End If
        Err.Clear
    Next
    MvFromMv = sLine
    Err.Clear
End Function
Function TreeViewPathLocation(treeDms As TreeView, ByVal SearchPath As String) As Long
    On Error Resume Next
    Dim myNode As Node
    TreeViewPathLocation = 0
    For Each myNode In treeDms.Nodes
        If LCase$(myNode.Fullpath) = LCase$(SearchPath) Then
            TreeViewPathLocation = myNode.Index
            Exit For
        End If
        Err.Clear
    Next
    Err.Clear
End Function
Public Function intError(ByVal strTitle As String, ByVal strmessage As String) As Integer
    On Error Resume Next
    intError = MyPrompt(strmessage, "o", "w", StringProperCase(strTitle))
    Err.Clear
End Function
Public Function MvFromCollection(objCollection As Collection, Optional ByVal Delim As String = "") As String
    On Error Resume Next
    Dim xTot As Long
    Dim xCnt As Long
    Dim sRet As String
    sRet = ""
    If Delim = "" Then
        Delim = VM
    End If
    xTot = objCollection.Count
    For xCnt = 1 To xTot
        sRet = sRet & objCollection.Item(xCnt) & Delim
        Err.Clear
    Next
    MvFromCollection = StringRemoveDelim(sRet, Delim)
    Err.Clear
End Function
Private Sub TreeViewWriteChild(MeForm As Form, progBar As ProgressBar, ByVal iNodeIndex As Long, ByVal mRs As DAO.Recordset, TreeV As TreeView, Optional ByVal TbType As String = "n")
    On Error Resume Next
    ' Write the child nodes to the table. This sub uses recursion
    ' to loop through the child nodes. It receives the Index of
    ' the node that has the children
    Dim i As Long
    Dim iTempIndex As Long
    Dim iTempChild As Long
    iTempIndex = TreeV.Nodes(iNodeIndex).Child.FirstSibling.Index
    iTempChild = TreeV.Nodes(iNodeIndex).Children
    'Loop through all a Parents Child Nodes
    For i = 1 To iTempChild
        cntNodes = cntNodes + 1
        Call UpdateProgress(MeForm, cntNodes, progBar, "Saving tree view to table...")
        mRs.AddNew
        mRs!Parent = TreeV.Nodes(iTempIndex).Parent.Index
        mRs!Key = TreeV.Nodes(iTempIndex).Index
        If LCase$(TbType) = "n" Then
            mRs!Text = Left$(TreeV.Nodes(iTempIndex).Text, 255)
        Else
            mRs!Text = TreeV.Nodes(iTempIndex).Text
        End If
        mRs!Sequence = cntNodes
        mRs!Image = TreeV.Nodes(iTempIndex).Image
        mRs!SelectedImage = TreeV.Nodes(iTempIndex).SelectedImage
        mRs!Tag = TreeV.Nodes(iTempIndex).Tag
        mRs!Checked = IIf((TreeV.Nodes(iTempIndex).Checked = True), 1, 0)
        mRs!Bold = IIf((TreeV.Nodes(iTempIndex).Bold = True), 1, 0)
        mRs!ForeColor = TreeV.Nodes(iTempIndex).ForeColor
        mRs!Fullpath = TreeV.Nodes(iTempIndex).Fullpath
        mRs.Update
        ' If the Node we are on has a child call the Sub again
        If TreeV.Nodes(iTempIndex).Children > 0 Then
            TreeViewWriteChild MeForm, progBar, iTempIndex, mRs, TreeV, TbType
        End If
        ' If we are not on the last child move to the next child Node
        If i <> TreeV.Nodes(iNodeIndex).Children Then
            iTempIndex = TreeV.Nodes(iTempIndex).Next.Index
        End If
        Err.Clear
    Next
    Err.Clear
End Sub
Public Function StringConcat(ByVal dest As String, ByVal Source As String) As String
    On Error Resume Next
    Dim sL As Long
    Dim dL As Long
    Dim NL As Long
    Dim sN As String
    Const cI As Long = 50000
    sN = dest
    sL = Len(Source)
    dL = Len(dest)
    NL = dL + sL
    Select Case NL
    Case Is >= dL
        Select Case sL
        Case Is > cI
            sN = sN & Space$(sL)
        Case Else
            sN = sN & Space$(sL + 1)
        End Select
    End Select
    Mid$(sN, dL + 1, sL) = Source
    StringConcat = Left$(sN, NL)
    Err.Clear
End Function
Public Function ArraySearch(varArray() As String, ByVal StrSearch As String) As Long
    On Error Resume Next
    ArraySearch = 0
    Dim ArrayTot As Long
    Dim arrayCnt As Long
    Dim strCur As String
    StrSearch = LCase$(Trim$(StrSearch))
    ArrayTot = UBound(varArray)
    If ArrayTot = 0 Then
        Err.Clear
        Exit Function
    End If
    For arrayCnt = 1 To ArrayTot
        strCur = varArray(arrayCnt)
        strCur = LCase$(Trim$(strCur))
        Select Case strCur
        Case StrSearch
            ArraySearch = arrayCnt
            Exit For
        End Select
        Err.Clear
    Next
    Err.Clear
End Function
Public Sub LstViewResizeMax(lstView As ListView)
    On Error Resume Next
    'Because applying the LVSCW_AUTOSIZE_USEHEADER
    'message to the last column in the control always
    'sets its width to the maximum remaining control
    'space, calling SendMessage passing the last column
    'will cause the listview data to utilize the full
    'control width space. For example, if a four-column
    'listview had a total width of 2000, and the first
    'three columns each had individual widths of 250,
    'calling this will cause the last column to widen
    'to cover the remaining 1250.
    'For this message to (visually) work as expected,
    'all columns should be within the viewing rect of the
    'listview control; if the last column is wider than
    'the control the message works, but the columns
    'remain wider than the control.
    Dim col2adjust As Long
    col2adjust = lstView.ColumnHeaders.Count - 1
    Call SendMessage(lstView.hWnd, LVM_SETCOLUMNWIDTH, col2adjust, ByVal LVSCW_AUTOSIZE_USEHEADER)
    Err.Clear
End Sub
Sub TreeViewRefreshKeys(MeForm As Form, progBar As ProgressBar, TreeV As TreeView, Optional ByVal RenumberKeys As Boolean = True)
    On Error Resume Next
    Dim treeT As Long
    Dim treeC As Long
    treeT = TreeV.Nodes.Count
    Call ProgBarInit(MeForm, progBar, treeT)
    For treeC = 1 To treeT
        Call UpdateProgress(MeForm, treeC, progBar, "Refreshing tree view keys...")
        TreeV.Nodes(treeC).Key = StringProperCase("\" & TreeV.Nodes(treeC).Fullpath)
        Err.Clear
    Next
    ProgBarClose MeForm, progBar
    If RenumberKeys = False Then
        Err.Clear
        Exit Sub
    End If
    Call ProgBarInit(MeForm, progBar, treeT)
    For treeC = 1 To treeT
        Call UpdateProgress(MeForm, treeC, progBar, "Renumbering tree view keys...")
        TreeV.Nodes(treeC).Key = "K" & CStr(treeC)
        Err.Clear
    Next
    ProgBarClose MeForm, progBar
    Err.Clear
End Sub
Public Sub LstViewRowsToMV(lstView As ListView, Retrows() As String, Optional ByVal Delim As String = "")
    On Error Resume Next
    Dim clsRowTot As Long
    Dim clsRowCnt As Long
    clsRowTot = lstView.ListItems.Count
    If Len(Delim) = 0 Then
        Delim = VM
    End If
    ReDim Retrows(clsRowTot)
    For clsRowCnt = 1 To clsRowTot
        Retrows(clsRowCnt) = LstViewRowToMv(lstView, clsRowCnt, Delim)
        Err.Clear
    Next
    Err.Clear
End Sub
Public Function LstViewRowToMv(lstView As ListView, ByVal rowPos As Long, Optional ByVal Delim As String = "") As String
    On Error Resume Next
    Dim lRow() As String
    lRow = LstViewGetRow(lstView, rowPos)
    LstViewRowToMv = MvFromArray(lRow, Delim)
    Err.Clear
End Function
Public Function LstViewGetRow(lstView As ListView, ByVal idx As Long) As Variant
    On Error Resume Next
    Dim retarray() As String
    Dim clsColTot As Long
    Dim clsColCnt As Long
    clsColTot = lstView.ColumnHeaders.Count
    ReDim retarray(clsColTot)
    retarray(1) = lstView.ListItems(idx).Text
    clsColTot = clsColTot - 1
    For clsColCnt = 1 To clsColTot
        retarray(clsColCnt + 1) = lstView.ListItems(idx).SubItems(clsColCnt)
        Err.Clear
    Next
    LstViewGetRow = retarray
    Err.Clear
End Function
Public Function TreeViewSearchPath(objTree As TreeView, ByVal StrSearch As String) As Long
    On Error Resume Next
    Dim iTmp As Long
    Dim iIndex As Long
    Dim mnIndex As Long
    Dim sCur As String
    TreeViewSearchPath = 0
    If objTree.Nodes.Count = 0 Then
        Err.Clear
        Exit Function
    End If
    mnIndex = 1
    'get the index of the root node that is at the top of the treeview
    iIndex = objTree.Nodes(mnIndex).FirstSibling.Index
    iTmp = iIndex
    sCur = UCase$(Trim$(objTree.Nodes(iIndex).Fullpath))
    StrSearch = UCase$(Trim$(StrSearch))
    Select Case StrSearch
    Case sCur
        TreeViewSearchPath = objTree.Nodes(iIndex).Index
        Err.Clear
        Exit Function
    End Select
    If objTree.Nodes(iIndex).Children > 0 Then
        TreeViewSearchPath = TreeViewSearchChildPath(iIndex, objTree, StrSearch)
        If TreeViewSearchPath >= 1 Then
            Err.Clear
            Exit Function
        End If
    End If
    While iIndex <> objTree.Nodes(iTmp).LastSibling.Index
        'loop through all the root nodes
        sCur = UCase$(Trim$(objTree.Nodes(iIndex).Next.Fullpath))
        Select Case StrSearch
        Case sCur
            TreeViewSearchPath = objTree.Nodes(iIndex).Next.Index
            Err.Clear
            Exit Function
        End Select
        If objTree.Nodes(iIndex).Next.Children > 0 Then
            TreeViewSearchPath = TreeViewSearchChildPath(objTree.Nodes(iIndex).Next.Index, objTree, StrSearch)
            If TreeViewSearchPath >= 1 Then
                Err.Clear
                Exit Function
            End If
        End If
        ' Move to the Next root Node
        iIndex = objTree.Nodes(iIndex).Next.Index
        Err.Clear
    Wend
    Err.Clear
End Function
Private Function TreeViewSearchChildPath(ByVal iNodeIndex As Long, ByVal objTree As TreeView, ByVal StrSearch As String) As Long
    On Error Resume Next
    Dim i As Long
    Dim iTempIndex As Long
    Dim lngChild As Long
    Dim sCur As String
    TreeViewSearchChildPath = 0
    StrSearch = UCase$(Trim$(StrSearch))
    iTempIndex = objTree.Nodes(iNodeIndex).Child.FirstSibling.Index
    'Loop through all a Parents Child Nodes
    lngChild = objTree.Nodes(iNodeIndex).Children
    For i = 1 To lngChild
        sCur = UCase$(Trim$(objTree.Nodes(iTempIndex).Fullpath))
        Select Case StrSearch
        Case sCur
            TreeViewSearchChildPath = objTree.Nodes(iTempIndex).Index
            Exit For
        End Select
        ' If the Node we are on has a child call the Sub again
        If objTree.Nodes(iTempIndex).Children > 0 Then
            TreeViewSearchChildPath = TreeViewSearchChildPath(iTempIndex, objTree, StrSearch)
            If TreeViewSearchChildPath >= 1 Then
                Exit For
            End If
        End If
        ' If we are not on the last child move to the next child Node
        If i <> objTree.Nodes(iNodeIndex).Children Then
            iTempIndex = objTree.Nodes(iTempIndex).Next.Index
        End If
        Err.Clear
    Next
    Err.Clear
End Function
Public Sub LstViewRemoveChecked(lstView As ListView, Optional ByVal bCheckedStatus As Boolean = True)
    On Error Resume Next
    Dim bOp As Boolean
    Dim lstTot As Long
    Dim lstCnt As Long
    lstTot = lstView.ListItems.Count
    For lstCnt = lstTot To 1 Step -1
        bOp = lstView.ListItems(lstCnt).Checked
        If bOp = bCheckedStatus Then
            lstView.ListItems.Remove lstCnt
        End If
        Err.Clear
    Next
    Err.Clear
End Sub
Public Function LstViewCheckedToMV(lstView As ListView, ByVal lngPos As Long, Optional ByVal Delim As String = "", Optional bRemoveDuplicates As Boolean = False, Optional bRemoveBlanks As Boolean = False, Optional bRemoveStars As Boolean = True, Optional bRemoveTotals As Boolean = True) As String
    On Error Resume Next
    Dim lstTot As Long
    Dim lstCnt As Long
    Dim bOp As Boolean
    Dim lstStr() As String
    Dim retStr As String
    retStr = ""
    If Len(Delim) = 0 Then
        Delim = VM
    End If
    lstTot = lstView.ListItems.Count
    For lstCnt = 1 To lstTot
        bOp = lstView.ListItems(lstCnt).Checked
        Select Case bOp
        Case True
            lstStr = LstViewGetRow(lstView, lstCnt)
            retStr = StringsConcat(retStr, lstStr(lngPos), Delim)
        End Select
        Err.Clear
    Next
    retStr = StringRemoveDelim(retStr, Delim)
    If bRemoveTotals = True Then
        retStr = Replace$(retStr, "Totals", "")
    End If
    If bRemoveStars = True Then retStr = Replace$(retStr, "*", "")
    If bRemoveDuplicates = True Then
        retStr = MvRemoveDuplicates(retStr, Delim)
    End If
    If bRemoveBlanks = True Then
        retStr = MvRemoveBlanks(retStr, Delim)
    End If
    LstViewCheckedToMV = retStr
    Err.Clear
End Function
Public Function Dao_ReadRecordToArray(ByVal Dbase As String, ByVal TableName As String, ByVal TableKey As String, ByVal ValuetoSeek As String, FieldsToRead As Variant) As Variant
    On Error Resume Next
    ' reads an array of fields and puts results in an array
    Dim adoC As DAO.Database
    Dim adoRs As DAO.Recordset
    Dim spTot As Long
    Dim spCnt As Long
    Dim spFld As String
    Dim spVal As String
    Dim spRec() As String
    Dao_ReadRecordToArray = spRec
    Set adoC = DAO.OpenDatabase(Dbase)
    Set adoRs = adoC.OpenRecordset(TableName)
    adoRs.Index = TableKey
    adoRs.Seek "=", ValuetoSeek
    Select Case adoRs.NoMatch
    Case False
        spTot = UBound(FieldsToRead)
        ReDim spRec(spTot)
        For spCnt = 1 To spTot
            spFld = FieldsToRead(spCnt)
            spVal = adoRs.Fields(spFld).Value & ""
            spRec(spCnt) = spVal
            Err.Clear
        Next
    End Select
    Dao_ReadRecordToArray = spRec
    adoRs.Close
    adoC.Close
    Set adoC = Nothing
    Set adoRs = Nothing
    Err.Clear
End Function
Public Function Dao_RecordExists(ByVal Dbase As String, ByVal TableName As String, ByVal TableKey As String, ByVal ValuetoSeek As String) As Boolean
    On Error Resume Next
    ' reads an array of fields and puts results in an array
    Dim adoC As DAO.Database
    Dim adoRs As DAO.Recordset
    Set adoC = DAO.OpenDatabase(Dbase)
    Set adoRs = adoC.OpenRecordset(TableName)
    adoRs.Index = TableKey
    adoRs.Seek "=", ValuetoSeek
    Dao_RecordExists = Not adoRs.NoMatch
    adoRs.Close
    adoC.Close
    Set adoC = Nothing
    Set adoRs = Nothing
    Err.Clear
End Function
Public Sub LstViewFromMv(lstView As ListView, ByVal StringMv As String, Optional ByVal Delim As String = "", Optional ByVal lstClear As String = "")
    On Error Resume Next
    Dim spLine() As String
    Dim spTot As Long
    Dim spCnt As Long
    If Len(Delim) = 0 Then
        Delim = VM
    End If
    Call StringParse(spLine, StringMv, Delim)
    spTot = UBound(spLine)
    If Len(lstClear) = 0 Then
        lstView.ListItems.Clear
    End If
    For spCnt = 1 To spTot
        Call lstView.ListItems.Add(, , spLine(spCnt))
        Err.Clear
    Next
    LstViewAutoResize lstView
    Err.Clear
End Sub
Public Sub dbConvertValue(ByVal FldName As DAO.Field, FldValue As Variant)
    On Error Resume Next
    Select Case FldName.Type
    Case dbText, dbMemo, dbChar
        FldName.Value = CStr(FldValue)
    Case dbBoolean
        FldName.Value = CBool(FldValue)
    Case dbByte
        FldName.Value = CByte(FldValue)
    Case dbDecimal
        FldName.Value = CDec(FldValue)
    Case dbInteger
        FldName.Value = CInt(FldValue)
    Case dbDouble
        FldName.Value = CDbl(FldValue)
    Case dbCurrency
        FldName.Value = CCur(FldValue)
    Case dbLong
        FldName.Value = CLng(FldValue)
    Case dbSingle
        FldName.Value = CSng(FldValue)
    Case dbCurrency
        FldName.Value = CCur(FldValue)
    Case dbDouble
        FldName.Value = CDbl(FldValue)
    Case dbDate
        If IsDate(FldValue) = True Then
            FldName.Value = CDate(FldValue)
        Else
            FldName.Value = Null
        End If
    Case Else
        FldName.Value = FldValue
    End Select
    Err.Clear
End Sub
Public Sub TextBoxHiLite(TxtBox As Variant)
    On Error Resume Next
    With TxtBox
        .SelStart = 0
        .SelLength = Len(TxtBox.Text)
    End With
    Err.Clear
End Sub
Public Function StringSpellNumber(ByVal Mynumber As String) As String
    On Error Resume Next
    Dim Dollars As String
    Dim Cents As String
    Dim temp As String
    Dim DecimalPlace As Integer
    Dim Count As Integer
    ReDim Place(9) As String
    Place(2) = " Thousand "
    Place(3) = " Million "
    Place(4) = " Billion "
    Place(5) = " Trillion "
    ' String representation of amount.
    Mynumber = Trim$(Mynumber)
    ' Position of decimal place 0 if none.
    DecimalPlace = InStr(Mynumber, ".")
    ' Convert cents and set MyNumber to dollar amount.
    If DecimalPlace > 0 Then
        Cents = StringGetTens(Left$(Mid$(Mynumber, DecimalPlace + 1) & "00", 2))
        Mynumber = Trim$(Left$(Mynumber, DecimalPlace - 1))
    End If
    Count = 1
    Do While Mynumber <> ""
        temp = StringGetHundreds(Right$(Mynumber, 3))
        If temp <> "" Then
            Dollars = temp & Place(Count) & Dollars
        End If
        If Len(Mynumber) > 3 Then
            Mynumber = Left$(Mynumber, Len(Mynumber) - 3)
        Else
            Mynumber = ""
        End If
        Count = Count + 1
        Err.Clear
    Loop
    Select Case Dollars
    Case ""
        Dollars = ""        '"No Dollars"
    Case "One"
        Dollars = "One"     '"One Dollar"
    Case Else
        Dollars = Dollars   ' & " Dollars"
    End Select
    Select Case Cents
    Case ""
        Cents = ""   '" and No Cents"
    Case "One"
        Cents = " One"  '" and One Cent"
    Case Else
        Cents = " " & Cents '  " and " & Cents & " Cents"
    End Select
    StringSpellNumber = Dollars & Cents
    Err.Clear
End Function
Private Function StringGetTens(ByVal Tenstext As String) As String
    On Error Resume Next
    Dim Result As String
    Result = ""           ' Null out the temporary function value.
    If Val(Left$(Tenstext, 1)) = 1 Then
        ' If value between 10-19...
        Select Case Val(Tenstext)
        Case 10: Result = "Ten"
        Case 11: Result = "Eleven"
        Case 12: Result = "Twelve"
        Case 13: Result = "Thirteen"
        Case 14: Result = "Fourteen"
        Case 15: Result = "Fifteen"
        Case 16: Result = "Sixteen"
        Case 17: Result = "Seventeen"
        Case 18: Result = "Eighteen"
        Case 19: Result = "Nineteen"
        Case Else
        End Select
    Else                                 ' If value between 20-99...
        Select Case Val(Left$(Tenstext, 1))
        Case 2: Result = "Twenty "
        Case 3: Result = "Thirty "
        Case 4: Result = "Forty "
        Case 5: Result = "Fifty "
        Case 6: Result = "Sixty "
        Case 7: Result = "Seventy "
        Case 8: Result = "Eighty "
        Case 9: Result = "Ninety "
        Case Else
        End Select
        Result = Result & StringGetDigit(Right$(Tenstext, 1))        ' Retrieve ones place.
    End If
    StringGetTens = Result
    Err.Clear
End Function
Private Function StringGetHundreds(ByVal Mynumber As String) As String
    On Error Resume Next
    Dim Result As String
    Dim resTmp As String
    If Val(Mynumber) = 0 Then
        Err.Clear
        Exit Function
    End If
    Mynumber = Right$("000" & Mynumber, 3)
    ' Convert the hundreds place.
    If Mid$(Mynumber, 1, 1) <> "0" Then
        Result = StringGetDigit(Mid$(Mynumber, 1, 1)) & " Hundred "
    End If
    ' Convert the tens and ones place.
    If Mid$(Mynumber, 2, 1) <> "0" Then
        resTmp = StringGetTens(Mid$(Mynumber, 2))
        Select Case resTmp
        Case ""
            Result = Result & resTmp
        Case Else
            If Result <> "" Then
                Result = Result & "and " & resTmp
            Else
                Result = Result & resTmp
            End If
        End Select
        'Result = Result & GetTens(mid$(MyNumber, 2))
    Else
        Result = Result & StringGetDigit(Mid$(Mynumber, 3))
    End If
    StringGetHundreds = Result
    Err.Clear
End Function
Private Function StringGetDigit(ByVal Digit As String) As String
    On Error Resume Next
    Select Case Val(Digit)
    Case 1: StringGetDigit = "One"
    Case 2: StringGetDigit = "Two"
    Case 3: StringGetDigit = "Three"
    Case 4: StringGetDigit = "Four"
    Case 5: StringGetDigit = "Five"
    Case 6: StringGetDigit = "Six"
    Case 7: StringGetDigit = "Seven"
    Case 8: StringGetDigit = "Eight"
    Case 9: StringGetDigit = "Nine"
    Case Else: StringGetDigit = ""
    End Select
    Err.Clear
End Function
Public Function StringFileFilters() As String
    On Error Resume Next
    Dim s_result As String
    s_result = "All Files (*.*)|*.*|Template File (*.tem)|*.tem"
    s_result = s_result & "|Temporal File (*.tmp)|*.tmp|Transaction File (*.trn)|*.trn"
    s_result = s_result & "|Data File (*.dat)|*.dat|Settings File (*.ini)|*.ini"
    s_result = s_result & "|Wave File (*.wav)|*.wav|Mpeg 3 File (*.mp3)|*.mp3"
    s_result = s_result & "|Help File Creator (*.hfc)|*.hfc|Bible File (*.bib)|*.bib"
    s_result = s_result & "|Dictionary File (*.dic)|*.dic|Topic Note File (*.top)|*.top"
    s_result = s_result & "|Study Node File (*.stu)|*.stu|Commentary File (*.com)|*.com"
    s_result = s_result & "|Graphics File (*.gra)|*.gra|Audio File (*.aud)|*.aud"
    s_result = s_result & "|Comma-Separated Values (*.csv)|*.csv|Sequential Access (*.seq)|*.seq"
    s_result = s_result & "|Excel (*.xls)|*.xls|Lotus 123 (*.wks)|*.wks"
    s_result = s_result & "|Rich Text Format (*.rtf)|*.rtf|Text (*.txt)|*.txt"
    s_result = s_result & "|Word for Windows (*.doc)|*.doc|Microsoft Access (*.mdb)|*.mdb"
    s_result = s_result & "|Adobe Acrobat (*.pdf)|*.pdf"
    s_result = s_result & "|Tree File (*.tree)|*.tree|Dictionary File (*.dict)|*.dict"
    s_result = s_result & "|Visual Basic Project File (*.vbp)|*.vbp|Visual Basic Project Group File (*.vbg)|*.vbg"
    s_result = s_result & "|Visual Basic Mak File (*.mak)|*.mak|Visual Basic Form File(*.frm)|*.frm"
    s_result = s_result & "|Visual Basic Module File (*.mod)|*.mod|Visual Basic Class Module (*.cls)|*.cls"
    s_result = s_result & "|Bitmap File (*.bmp)|*.bmp|Tif File (*.tif)|*.tif"
    s_result = s_result & "|Tiff File (*.tif)|*.tif|Jpeg File (*.jpg)|*.jpg"
    s_result = s_result & "|Gif File (*.gif)|*.gif|Png File (*.png)|*.png"
    s_result = s_result & "|Batch File (*.bat)|*.bat|Executable File (*.exe)|*.exe"
    s_result = s_result & "|Icon File (*.ico|*.ico|Configuration (*.cfg)|*.cfg"
    s_result = s_result & "|Visual Basic Setup File (*.lst)|*.lst|Inno Setup Script (*.iss)|*.iss"
    s_result = s_result & "|Spell Check Log (*.spl)|*.spl|Document Tracking System (*.dts)|*.dts"
    s_result = s_result & "|Archive File (*.arc)|*.arc|List View File (*.lvf)|*.lvf"
    s_result = s_result & "|Executable File (*.exe)|*.exe|Dynamic Link Library File (*.dll)|*.dll"
    StringFileFilters = s_result
    Err.Clear
End Function
Public Function DialogOpen(ByVal Filter As String, Optional ByVal Title As String = "", Optional ByVal InitDir As String = "", Optional ByVal DefaultExt As String = "") As String
    On Error GoTo ErrHandler
    Dim filName As String
    Dim filCnt As Integer
    Dim spFilter() As String
    filCnt = 0
    If Len(DefaultExt) = 0 Then
        DefaultExt = "*.*"
    End If
    If Len(Title) = 0 Then
        Title = "Open an existing file"
    End If
    If Len(InitDir) = 0 Then
        InitDir = ReadReg("lastpath")
    End If
    If Len(InitDir) = 0 Then
        InitDir = "..."
    End If
    Call StringParse(spFilter, Filter, "|")
    filCnt = ArraySearch(spFilter, DefaultExt)
    If filCnt <> 0 Then
        filCnt = filCnt / 2
    End If
    filName = ""
    With frmKB.cd1
        .CancelError = True
        '.flags = OFN_HIDEREADONLY Or OFN_FILEMUSTEXIST
        .Filter = Filter
        .DialogTitle = Title
        .InitDir = InitDir
        .FileName = ""
        .DefaultExt = DefaultExt
        .FilterIndex = filCnt
        .ShowOpen
        filName = .FileName
        If Len(filName) = 0 Then
            Err.Clear
            Exit Function
        End If
    End With
    DialogOpen = filName
    SaveReg "lastpath", StringGetFileToken(filName, "p")
    Err.Clear
    Exit Function
ErrHandler:
    Err.Clear
    Exit Function
    Err.Clear
End Function
Private Sub IniHeadings()
    On Error Resume Next
    ViewHeadings(1) = "Amount"
    ViewHeadings(2) = "Amt"
    ViewHeadings(3) = "Receipt"
    ViewHeadings(4) = "Pay"
    ViewHeadings(5) = "January"
    ViewHeadings(6) = "February"
    ViewHeadings(7) = "March"
    ViewHeadings(8) = "April"
    ViewHeadings(9) = "May"
    ViewHeadings(10) = "June"
    ViewHeadings(11) = "July"
    ViewHeadings(12) = "August"
    ViewHeadings(13) = "September"
    ViewHeadings(14) = "October"
    ViewHeadings(15) = "November"
    ViewHeadings(16) = "December"
    ViewHeadings(17) = "Tot"
    ViewHeadings(18) = "Members"
    ViewHeadings(19) = "Supposed"
    ViewHeadings(20) = "Actual"
    ViewHeadings(21) = "Active"
    ViewHeadings(22) = "Deceased"
    ViewHeadings(23) = "Good"
    ViewHeadings(24) = "Bad"
    ViewHeadings(25) = "Married"
    ViewHeadings(26) = "Divorced"
    ViewHeadings(27) = "Single"
    ViewHeadings(28) = "Widowed"
    ViewHeadings(29) = "Amount"
    ViewHeadings(30) = "Registered"
    ViewHeadings(31) = "Temporal"
    ViewHeadings(32) = "Suspended"
    ViewHeadings(33) = "Size"
    ViewHeadings(34) = "Receipt"
    ViewHeadings(35) = "Payee"
    ViewHeadings(36) = "Paypoint"
    ViewHeadings(37) = "Tax"
    ViewHeadings(38) = "Sales"
    ViewHeadings(39) = "Cash"
    ViewHeadings(40) = "Debit"
    ViewHeadings(41) = "Credit"
    ViewHeadings(42) = "Exclusive"
    ViewHeadings(43) = "Inclusive"
    ViewHeadings(44) = "Start"
    ViewHeadings(45) = "End"
    ViewHeadings(46) = "Current"
    ViewHeadings(47) = "Difference"
    ViewHeadings(48) = "Number"
    ViewHeadings(49) = "Membership Number"
    ViewHeadings(50) = "Member #"
    ViewHeadings(51) = "Membership #"
    ViewHeadings(52) = "Age"
    ViewHeadings(53) = "Average"
    ViewHeadings(54) = "Premium"
    ViewHeadings(55) = "Regions"
    ViewHeadings(56) = "Id"
    ViewHeadings(57) = "Id Number"
    ViewHeadings(58) = "Id #"
    ViewHeadings(59) = "Id No."
    ViewHeadings(60) = "Id No"
    ViewHeadings(61) = "Incomes"
    ViewHeadings(62) = "Expenses"
    ViewHeadings(63) = "Totals"
    ViewHeadings(64) = "Real Receipts"
    ViewHeadings(65) = "Actual Receipts"
    ViewHeadings(66) = "Supposed Receipts"
    ViewHeadings(67) = "Qty"
    ViewHeadings(68) = "Quantity"
    ViewHeadings(69) = "Cash Sales"
    ViewHeadings(70) = "Member Sales"
    ViewHeadings(71) = "Sales"
    ViewHeadings(72) = "Price"
    ViewHeadings(73) = "Discount"
    ViewHeadings(74) = "Tax %"
    ViewHeadings(75) = "Disc %"
    ViewHeadings(76) = "Tax Amount"
    ViewHeadings(77) = "Discount Amount"
    ViewHeadings(78) = "Target"
    ViewHeadings(79) = "Granted"
    ViewHeadings(80) = "Recorded Expenses"
    ViewHeadings(81) = "Recorded Difference"
    ViewHeadings(82) = "Amount Required"
    ViewHeadings(83) = "Amount Received"
    Err.Clear
End Sub
Sub AddStatusBar(Sobject As StatusBar, progBar As ProgressBar)
    On Error Resume Next
    Dim pnlA As Panel
    Dim RowCounter As Integer
    For RowCounter = 1 To 5
        Set pnlA = Sobject.Panels.Add()
        Err.Clear
    Next
    ' set the style of each panel
    With Sobject.Panels
        .Item(1).Style = 0
        .Item(1).Width = 3850
        .Item(1).Bevel = sbrInset
        .Item(2).Style = 0
        .Item(2).Width = 3000
        .Item(2).Bevel = sbrInset
        .Item(3).Style = 0
        .Item(3).Width = 3000
        .Item(3).Bevel = sbrInset
        .Item(4).Style = 0
        .Item(4).Width = 3000
        .Item(4).Bevel = sbrInset
        .Item(5).Width = 1000
        .Item(5).Bevel = sbrInset
    End With
    Sobject.Refresh
    PutProgressBarInStatusBar Sobject, progBar, 4
    Set pnlA = Nothing
    Err.Clear
End Sub
Sub ResizeStatusBar(objForm As Form, objStatusBar As StatusBar, progBar As ProgressBar)
    On Error Resume Next
    Dim lngSum As Long
    With objStatusBar.Panels
        lngSum = .Item(1).Width + .Item(2).Width + .Item(3).Width + .Item(4).Width
        .Item(5).Width = objForm.Width - lngSum
    End With
    objStatusBar.Refresh
    PutProgressBarInStatusBar objStatusBar, progBar, 4
    Err.Clear
End Sub
Public Sub PutProgressBarInStatusBar(objStatusBar As StatusBar, objProgressBar As ProgressBar, PnlNumber As Integer)
    On Error Resume Next
    Dim R As RECT
    SetParent objProgressBar.hWnd, objStatusBar.hWnd
    SendMessage objStatusBar.hWnd, SB_GETRECT, PnlNumber - 1, R
    MoveWindow objProgressBar.hWnd, R.Left + 1, R.Top + 1, R.Right - R.Left - 2, R.Bottom - R.Top - 2, True
    Err.Clear
End Sub
Public Sub ProgBarClose(MeForm As Form, progBar As ProgressBar)
    On Error Resume Next
    progBar.Value = 0
    StatusMessage MeForm
    Err.Clear
End Sub
Sub UpdateProgress(objForm As Form, ByVal minValue As Long, progBar As ProgressBar, Optional ByVal Note As String = "")
    On Error Resume Next
    progBar.Value = minValue
    StatusMessage objForm, Note
    DoEvents
    Err.Clear
End Sub
Public Function Keywords_Validate(ByVal strFullPath As String) As String
    On Error Resume Next
    strFullPath = MvFromMv(strFullPath, 2, , "\")
    strFullPath = Replace$(strFullPath, "\", FM)
    Keywords_Validate = Replace$(strFullPath, " ", FM)
    Err.Clear
End Function
Public Sub RemoveUselessTopics(MeForm As Form, progBar As ProgressBar)
    On Error Resume Next
    Dim rsCnt As Long
    Dim rsTot As Long
    Dim db As DAO.Database
    Dim tbP As DAO.Recordset
    Dim tbT As DAO.Recordset
    Dim tbD As DAO.Recordset
    Dim sFullPath As String
    StatusMessage MeForm, "Selecting topics, please wait..."
    Set db = DAO.OpenDatabase(sProjDb)
    Set tbP = db.OpenRecordset("Properties")
    Set tbT = db.OpenRecordset(sProject)
    'Set tbD = db.OpenRecordset("DataFiles")
    tbT.Index = "FullPath"
    tbP.MoveLast
    rsTot = tbP.RecordCount
    ProgBarInit MeForm, progBar, rsTot
    For rsCnt = rsTot To 1 Step -1
        sFullPath = tbP!Fullpath.Value & ""
        UpdateProgress MeForm, rsCnt, progBar, "Checking topic validity: 1/2" & sFullPath
        tbT.Seek "=", sFullPath
        If tbT.NoMatch = True Then tbP.Delete
        tbP.MovePrevious
        Err.Clear
    Next
    ''
    '    tbD.MoveLast
    '    rsTot = tbD.RecordCount
    '    ProgBarInit MeForm, progBar, rsTot
    '    For rsCnt = rsTot To 1 Step -1
    '        sFullPath = tbD!Fullpath.Value & ""
    '        UpdateProgress MeForm, rsCnt, progBar, "Checking topic validity: 2/2" & sFullPath
    '        tbT.Seek "=", sFullPath
    '        If tbT.NoMatch = True Then tbD.Delete
    '        tbD.MovePrevious
    '    Next
    '
    ProgBarClose MeForm, progBar
    tbT.Close
    tbP.Close
    '    tbD.Close
    db.Close
    Set tbT = Nothing
    Set tbP = Nothing
    '    Set tbD = Nothing
    Set db = Nothing
    Err.Clear
End Sub
Public Function FileData(ByVal FileName As String) As String
    On Error Resume Next
    ' return contents of file
    Dim sLen As Long
    Dim fileNum As Long
    Dim Size As Long
    fileNum = FreeFile
    Size = FileLen(FileName)
    Open FileName For Input Access Read As #fileNum
        sLen = LOF(fileNum)
        FileData = Input(sLen, #fileNum)
    Close #fileNum
    Err.Clear
End Function
Public Function MvField(ByVal strData As String, Optional ByVal fldPos As Long = 1, Optional ByVal Delim As String = ";") As String
    On Error Resume Next
    Dim spData() As String
    Dim spCnt As Long
    MvField = ""
    If Len(Delim) = 0 Then
        Delim = VM
    End If
    If Len(strData) = 0 Then
        Err.Clear
        Exit Function
    End If
    Call StringParse(spData, strData, Delim)
    spCnt = UBound(spData)
    Select Case fldPos
    Case -1
        MvField = Trim$(spData(spCnt))
    Case -2
        MvField = Trim$(spData(spCnt - 1))
    Case Else
        If fldPos <= spCnt Then
            MvField = Trim$(spData(fldPos))
        End If
    End Select
    Err.Clear
End Function
Public Sub TOC_Fix(MeForm As Form, progBar As ProgressBar)
    On Error Resume Next
    Dim rsCnt As Long
    Dim rsTot As Long
    Dim OriginalFile() As String
    Dim correctFile() As String
    Dim strContents As String
    Dim strPart As String
    Dim lngTopic As Long
    Dim strLine As String
    Dim strNLine As String
    Dim topicNum As Long
    Dim topicNNum As Long
    Dim topicTot As Long
    Dim strLinesToFix As String
    Dim posLoc As Long
    Dim lngPartLen As Integer
    strLinesToFix = ""
    lngTopic = 0
    StatusMessage MeForm, "Reading contents file contents, please wait..."
    strContents = Trim$(FileData(sProjCnt))
    strContents = MvRemoveBlanks(strContents, vbNewLine) & vbNewLine
    rsTot = StringParse(OriginalFile, strContents, vbNewLine)
    ReDim correctFile(rsTot)
    ProgBarInit MeForm, progBar, rsTot
    For rsCnt = 1 To rsTot
        UpdateProgress MeForm, rsCnt, progBar, "Fixing contents file (1/2), please be patient..."
        ' current line
        strLine = Trim$(OriginalFile(rsCnt))
        If Len(strLine) = 0 Then GoTo NextTopic
        ' next line
        strNLine = Trim$(OriginalFile(rsCnt + 1))
        ' determine starting part of the line
        strPart = Left$(strLine, 1)
        ' check the part of current line
        Select Case strPart
        Case ":"
            ' configuration settings
            lngTopic = lngTopic + 1
            ReDim Preserve correctFile(lngTopic)
            correctFile(lngTopic) = strLine
        Case "1", "2", "3", "4", "5", "6", "7", "8", "9"
            ' the topic starts with a number
            topicNum = Val(MvField(strLine, 1, " "))
            ' the next topic should start with a number
            topicNNum = Val(MvField(strNLine, 1, " "))
            If topicNNum < topicNum Then
                ' add current topic
                lngTopic = lngTopic + 1
                ReDim Preserve correctFile(lngTopic)
                correctFile(lngTopic) = strLine
                ' is the next topic a book or a leaf, leaf topics have =
                ' if its a book, its ok
                topicTot = MvCount(strNLine, "=")
                If topicTot = 1 Then
                    ' this is a book
                Else
                    ' this is a leaf, add a parent book
                    lngTopic = lngTopic + 1
                    ReDim Preserve correctFile(lngTopic)
                    correctFile(lngTopic) = MvField(strNLine, 1, "=")
                    ' store location of topic to fix
                    strLinesToFix = strLinesToFix & CStr(lngTopic + 1) & ";"
                End If
            Else
                ' add current topic
                lngTopic = lngTopic + 1
                ReDim Preserve correctFile(lngTopic)
                correctFile(lngTopic) = strLine
            End If
        End Select
NextTopic:
        Err.Clear
    Next
    ' increment lines to fix by one
    rsTot = UBound(correctFile)
    ProgBarInit MeForm, progBar, rsTot
    For rsCnt = 1 To rsTot
        UpdateProgress MeForm, rsCnt, progBar, "Fixing contents file (2/2), please be patient..."
        ' is this topic part of those to fix
        posLoc = MvSearch(strLinesToFix, CStr(rsCnt), ";")
        If posLoc > 0 Then
            ' read previous topic
            strLine = correctFile(rsCnt - 1)
            strNLine = correctFile(rsCnt)
            ' read previous topic level
            strPart = MvField(strLine, 1, " ")
            ' how long is this part
            lngPartLen = Len(strPart)
            ' the fixed line should be incremented by one
            strNLine = Val(strPart) + 1 & Mid$(strNLine, lngPartLen + 1)
            correctFile(rsCnt) = strNLine
        End If
        Err.Clear
    Next
    strContents = MvFromArray(correctFile, vbNewLine)
    FileUpdate sProjCnt, strContents, "w"
    ProgBarClose MeForm, progBar
    Err.Clear
End Sub
Public Function TOC_Errors(MeForm As Form, progBar As ProgressBar) As Long
    On Error Resume Next
    Dim rsCnt As Long
    Dim rsTot As Long
    Dim OriginalFile() As String
    Dim correctFile() As String
    Dim strContents As String
    Dim strPart As String
    Dim lngTopic As Long
    Dim strLine As String
    Dim strNLine As String
    Dim topicNum As Long
    Dim topicNNum As Long
    Dim topicTot As Long
    Dim strLinesToFix As String
    strLinesToFix = ""
    lngTopic = 0
    StatusMessage MeForm, "Reading contents file contents, please wait..."
    strContents = Trim$(FileData(sProjCnt))
    strContents = MvRemoveBlanks(strContents, vbNewLine)
    rsTot = StringParse(OriginalFile, strContents, vbNewLine)
    ReDim correctFile(rsTot)
    ProgBarInit MeForm, progBar, rsTot
    For rsCnt = 1 To rsTot - 1
        UpdateProgress MeForm, rsCnt, progBar, "Fixing contents file (1/2), please be patient..."
        ' current line
        strLine = Trim$(OriginalFile(rsCnt))
        If Len(strLine) = 0 Then GoTo NextTopic
        ' next line
        strNLine = Trim$(OriginalFile(rsCnt + 1))
        ' determine starting part of the line
        strPart = Left$(strLine, 1)
        ' check the part of current line
        Select Case strPart
        Case ":"
            ' configuration settings
            lngTopic = lngTopic + 1
            ReDim Preserve correctFile(lngTopic)
            correctFile(lngTopic) = strLine
        Case "1", "2", "3", "4", "5", "6", "7", "8", "9"
            ' the topic starts with a number
            topicNum = Val(MvField(strLine, 1, " "))
            ' the next topic should start with a number
            topicNNum = Val(MvField(strNLine, 1, " "))
            If topicNNum < topicNum Then
                ' add current topic
                lngTopic = lngTopic + 1
                ReDim Preserve correctFile(lngTopic)
                correctFile(lngTopic) = strLine
                ' is the next topic a book or a leaf, leaf topics have =
                ' if its a book, its ok
                topicTot = MvCount(strNLine, "=")
                If topicTot = 1 Then
                    ' this is a book
                Else
                    ' this is a leaf, add a parent book
                    lngTopic = lngTopic + 1
                    ReDim Preserve correctFile(lngTopic)
                    correctFile(lngTopic) = MvField(strNLine, 1, "=")
                    ' store location of topic to fix
                    strLinesToFix = strLinesToFix & CStr(lngTopic + 1) & ";"
                End If
            Else
                ' add current topic
                lngTopic = lngTopic + 1
                ReDim Preserve correctFile(lngTopic)
                correctFile(lngTopic) = strLine
            End If
        End Select
NextTopic:
        Err.Clear
    Next
    ProgBarClose MeForm, progBar
    strLinesToFix = StringRemoveDelim(strLinesToFix, ";")
    TOC_Errors = MvCount(strLinesToFix, ";")
    Err.Clear
End Function
Public Function Jump_Propercase(ByVal strValue As String) As String
    On Error Resume Next
    Dim strBefore As String
    Dim strAfter As String
    Dim uldbPos As Long
    Dim strBetween As String
    Dim rsCnt As Long
    Dim rsTot As Long
    Dim rsStr As String
    Dim backPos As Long
    Dim strNew As String
    Dim ulPos As Long
    strValue = Replace$(strValue, vbNewLine, VM)
    strBefore = ""
    strBetween = ""
    uldbPos = InStr(1, strValue, "\uldb", vbTextCompare)
    ulPos = InStr(1, strValue, "\ul", vbTextCompare)
    If uldbPos > 0 Then
        ' find next space
        strBefore = Left$(strValue, uldbPos - 1)
        strAfter = Mid$(strValue, uldbPos + 5)
        rsTot = Len(strAfter)
        For rsCnt = 1 To rsTot
            rsStr = Mid$(strAfter, rsCnt, 1)
            Select Case rsStr
            Case " "
                Exit For
            Case Else
                strBetween = strBetween & rsStr
            End Select
            Err.Clear
        Next
        strBefore = strBefore & "\uldb" & strBetween
        strAfter = Mid$(strAfter, Len(strBetween) + 1)
        backPos = InStr(1, strAfter, "\", vbTextCompare)
        If backPos = 0 Then
            strBetween = StringProperCase(strAfter)
            strNew = strBefore & " " & strBetween
            Jump_Propercase = Replace$(strNew, VM, vbNewLine)
        Else
            strBetween = Left$(strAfter, backPos - 1)
            strBetween = StringProperCase(strBetween)
            strAfter = Mid$(strAfter, backPos)
            strNew = strBefore & " " & strBetween & strAfter
            Jump_Propercase = Replace$(strNew, VM, vbNewLine)
        End If
    ElseIf ulPos > 0 Then
        ' find next space
        strBefore = Left$(strValue, ulPos - 1)
        strAfter = Mid$(strValue, ulPos + 3)
        rsTot = Len(strAfter)
        For rsCnt = 1 To rsTot
            rsStr = Mid$(strAfter, rsCnt, 1)
            Select Case rsStr
            Case " "
                Exit For
            Case Else
                strBetween = strBetween & rsStr
            End Select
            Err.Clear
        Next
        strBefore = strBefore & "\ul" & strBetween
        strAfter = Mid$(strAfter, Len(strBetween) + 1)
        backPos = InStr(1, strAfter, "\", vbTextCompare)
        If backPos = 0 Then
            strBetween = StringProperCase(strAfter)
            strNew = strBefore & " " & strBetween
            Jump_Propercase = Replace$(strNew, VM, vbNewLine)
        Else
            strBetween = Left$(strAfter, backPos - 1)
            strBetween = StringProperCase(strBetween)
            strAfter = Mid$(strAfter, backPos)
            strNew = strBefore & " " & strBetween & strAfter
            Jump_Propercase = Replace$(strNew, VM, vbNewLine)
        End If
    Else
        Jump_Propercase = Replace$(StringProperCase(strValue), VM, vbNewLine)
    End If
    Err.Clear
End Function
Public Function Jump_UpperCase(ByVal strValue As String) As String
    On Error Resume Next
    Dim strBefore As String
    Dim strAfter As String
    Dim uldbPos As Long
    Dim strBetween As String
    Dim rsCnt As Long
    Dim rsTot As Long
    Dim rsStr As String
    Dim backPos As Long
    Dim strNew As String
    Dim ulPos As Long
    strValue = Replace$(strValue, vbNewLine, VM)
    strBefore = ""
    strBetween = ""
    uldbPos = InStr(1, strValue, "\uldb", vbTextCompare)
    ulPos = InStr(1, strValue, "\ul", vbTextCompare)
    If uldbPos > 0 Then
        ' find next space
        strBefore = Left$(strValue, uldbPos - 1)
        strAfter = Mid$(strValue, uldbPos + 5)
        rsTot = Len(strAfter)
        For rsCnt = 1 To rsTot
            rsStr = Mid$(strAfter, rsCnt, 1)
            Select Case rsStr
            Case " "
                Exit For
            Case Else
                strBetween = strBetween & rsStr
            End Select
            Err.Clear
        Next
        strBefore = strBefore & "\uldb" & strBetween
        strAfter = Mid$(strAfter, Len(strBetween) + 1)
        backPos = InStr(1, strAfter, "\", vbTextCompare)
        If backPos = 0 Then
            strBetween = UCase$(strAfter)
            strNew = strBefore & " " & strBetween
            Jump_UpperCase = Replace$(strNew, VM, vbNewLine)
        Else
            strBetween = Left$(strAfter, backPos - 1)
            strBetween = UCase$(strBetween)
            strAfter = Mid$(strAfter, backPos)
            strNew = strBefore & " " & strBetween & strAfter
            Jump_UpperCase = Replace$(strNew, VM, vbNewLine)
        End If
    ElseIf ulPos > 0 Then
        ' find next space
        strBefore = Left$(strValue, ulPos - 1)
        strAfter = Mid$(strValue, ulPos + 3)
        rsTot = Len(strAfter)
        For rsCnt = 1 To rsTot
            rsStr = Mid$(strAfter, rsCnt, 1)
            Select Case rsStr
            Case " "
                Exit For
            Case Else
                strBetween = strBetween & rsStr
            End Select
            Err.Clear
        Next
        strBefore = strBefore & "\ul" & strBetween
        strAfter = Mid$(strAfter, Len(strBetween) + 1)
        backPos = InStr(1, strAfter, "\", vbTextCompare)
        If backPos = 0 Then
            strBetween = UCase$(strAfter)
            strNew = strBefore & " " & strBetween
            Jump_UpperCase = Replace$(strNew, VM, vbNewLine)
        Else
            strBetween = Left$(strAfter, backPos - 1)
            strBetween = UCase$(strBetween)
            strAfter = Mid$(strAfter, backPos)
            strNew = strBefore & " " & strBetween & strAfter
            Jump_UpperCase = Replace$(strNew, VM, vbNewLine)
        End If
    Else
        Jump_UpperCase = Replace$(UCase$(strValue), VM, vbNewLine)
    End If
    Err.Clear
End Function
Function StringNextFile(ByVal OriginalFile As String, Optional ByVal NewPath As String = "", Optional ByVal IncludeCopy As Boolean = True) As String
    On Error Resume Next
    Dim fCount As Long
    Dim strPath As String
    Dim strFile As String
    Dim fExist As Boolean
    Dim newFile As String
    Dim strFileOnly As String
    Dim strExtension As String
    fCount = 0
    strPath = StringGetFileToken(OriginalFile, "p")
    If Len(NewPath) > 0 Then
        strPath = NewPath
    End If
    strFile = StringGetFileToken(OriginalFile, "f")
    strFileOnly = StringGetFileToken(OriginalFile, "fo")
    strExtension = StringGetFileToken(OriginalFile, "e")
    CreateNestedDirectory NewPath
    newFile = strPath & "\" & strFile
    fExist = boolFileExists(newFile)
    Do Until fExist = False
        fCount = fCount + 1
        newFile = strPath & "\Copy " & CStr(fCount) & " Of " & strFile
        If IncludeCopy = False Then
            newFile = strPath & "\" & strFileOnly & " " & CStr(fCount) & "." & strExtension
        End If
        fExist = boolFileExists(newFile)
        Err.Clear
    Loop
    StringNextFile = newFile
    Err.Clear
End Function
Sub Main()
    On Error Resume Next
    If App.PrevInstance Then
        MsgBox "Another instance of MyHelp is running!", vbOKOnly + vbInformation + vbApplicationModal, "MyHelp Running"
        End
    End If
    App.HelpFile = App.Path & "\myhelp.hlp"
    IniHeadings
    Compression = 32
    Author = "Made with MyHelp " & App.Major & "." & App.Minor
    HeadlineColor = 0
    TextColor = 0
    HeadlineBackColor = 12648447
    TextBackColor = 12648447
    FontHeadline = "Tahoma"
    FontHeadlineSize = 12
    FontText = "Tahoma"
    FontTextSize = 8
    FontHeadlineBold = 1
    PictureHeight = 3615
    PictureWidth = 6375
    CompilerLocation = App.Path & "\Help\hcw.exe"
    Quote = Chr$(34)
    FM = Chr$(254)
    VM = Chr$(253)
    KM = Chr$(193)
    AppTitle = App.Title
    Load frmSplash
    frmSplash.Show
    frmSplash.Refresh
    If blnWord97Loaded = False Then RunMsWord
    Load frmKB
    frmKB.Show
    frmKB.Refresh
    Unload frmSplash
    Err.Clear
End Sub


Private Function RemDelim(ByVal Dataobj As String, ByVal Delimiter As String) As String
    On Error Resume Next
    Dim intDataSize As Long
    Dim intDelimSize As Long
    Dim strLast As String
    intDataSize = Len(Dataobj)
    intDelimSize = Len(Delimiter)
    strLast = Right$(Dataobj, intDelimSize)
    Select Case strLast
    Case Delimiter
        RemDelim = Left$(Dataobj, (intDataSize - intDelimSize))
    Case Else
        RemDelim = Dataobj
    End Select
    Err.Clear
End Function



Private Sub UpdateTopicDataFiles(ByVal TopicPath As String, ByVal FileName As String)
    On Error Resume Next
    Dim sExisting As String
    Dim db As DAO.Database
    Dim tb As DAO.Recordset
    Set db = DAO.OpenDatabase(sProjDb)
    Set tb = db.OpenRecordset("DataFiles")
    tb.Index = "Fullpath"
    tb.Seek "=", TopicPath
    If tb.NoMatch = False Then
        tb.Edit
        sExisting = tb!FileNames.Value & "" & TopicPath & Chr$(253)
    Else
        tb.AddNew
        tb!Fullpath = TopicPath
        sExisting = FileName
    End If
    sExisting = RemDelim(sExisting, Chr$(253))
    tb!FileNames.Value = sExisting
    tb.Update
    tb.Close
    db.Close
    Set tb = Nothing
    Set db = Nothing
    Err.Clear
End Sub


Public Function Dao_DatabaseCompress(ByVal Datab As String) As Boolean
    On Error GoTo Compact_Repair_Error
    If Len(Datab) = 0 Then Exit Function
    Dim RepairDb As String
    Dim TemporDb As String
    Dim TestDb As DAO.Database
    Dim Path As String
    Path = StringGetFileToken(Datab, "p")
    RepairDb = Datab
    TemporDb = ExactPath(Path) & "\tmp.mdb"
    If boolFileExists(TemporDb) = True Then
        Kill TemporDb
    End If
    Set TestDb = DAO.OpenDatabase(RepairDb, True, False)  ' open exclusive, read write
    TestDb.Close
    Set TestDb = Nothing
    DAO.DBEngine.RepairDatabase RepairDb
    DAO.DBEngine.CompactDatabase RepairDb, TemporDb
    FileCopy TemporDb, RepairDb
    Kill TemporDb
    Dao_DatabaseCompress = True
    Err.Clear
    Exit Function
Compact_Repair_Error:
    Select Case Err
    Case 401
        Resume Next
    Case Else
        Dao_DatabaseCompress = False
        Set TestDb = Nothing
        Err.Clear
        Exit Function
    End Select
    Err.Clear
End Function

Public Sub LstBoxRemoveItemAPI(lstBox As Variant, ParamArray cboItems())
    On Error Resume Next
    Dim cboItem As Variant
    Dim cboPos As Long
    Dim cboStr As String
    Select Case TypeName(lstBox)
    Case "ListBox"
        For Each cboItem In cboItems
            cboStr = CStr(cboItem)
            cboPos = LstBoxFindExactItemAPI(lstBox, cboStr$)
            If cboPos <> -1 Then
                Call SendMessage(lstBox.hWnd, LB_DELETESTRING, cboPos, ByVal 0&)
            End If
            Err.Clear
        Next
    Case "ComboBox"
        For Each cboItem In cboItems
            cboStr = CStr(cboItem)
            cboPos = LstBoxFindExactItemAPI(lstBox, cboStr$)
            If cboPos <> -1 Then
                Call SendMessage(lstBox.hWnd, CB_DELETESTRING, cboPos, ByVal 0&)
            End If
            Err.Clear
        Next
    End Select
    Err.Clear
End Sub

Public Function LstBoxFindExactItemAPI(lstBox As Variant, ByVal sSearch As String) As Long
    On Error Resume Next
    Select Case TypeName(lstBox)
    Case "ListBox"
        LstBoxFindExactItemAPI = SendMessage(lstBox.hWnd, LB_FINDSTRINGEXACT, 0&, ByVal sSearch$)
    Case "ComboBox"
        LstBoxFindExactItemAPI = SendMessage(lstBox.hWnd, CB_FINDSTRINGEXACT, 0&, ByVal sSearch$)
    End Select
    Err.Clear
End Function

