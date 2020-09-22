Attribute VB_Name = "Md_SeekFiles"
'
'___________________________________________________________________________
' Program name      : KhompAllMDB.
' Description       : An easy way to gain some space on your HD,
'                     by compacting Access Databases.
' Company           : MELANTECH
' Authors           : Weitten Pascal
'___________________________________________________________________________
'
' Date              : (c) 2005.01.19
' Version N°        : V0.1
' Customer          : Internal stuff.
'
' Last Modification : 2005.01.19
'___________________________________________________________________________
' TODO :
'       - .
'       -
'___________________________________________________________________________
'
' Module description:
' Scans all the directories and sub-directories and creates a collection
' containing all those directories to ease the search of MDB files.
'___________________________________________________________________________
'

Global strFilePath, strFileCurrrentName
Global Const AccessExtension00 = "mdb"
Dim LastRow As Long

Sub GetFileInfoFromDirectory(MonChemin)
    Dim dblMDBOriginalSize As Double
    
    On Error Resume Next
    strFilePath = MonChemin
    strFileCurrrentName = Dir(strFilePath, vbNormal)  'Starts on current HD unit.

    Do While strFileCurrrentName <> ""
        'Ignores current directory.
        If strFileCurrrentName <> "." And strFileCurrrentName <> ".." Then
            'Bit by bit comparison to see if we the current object is a file.
            If (GetAttr(strFilePath & strFileCurrrentName) And vbNormal) = vbNormal Then
                'Tests the file's extension.
                'To avoid any troubles we test the extension using CAPITAL chars.
                If UCase(Right(strFileCurrrentName, 3)) = UCase(AccessExtension00) Then
                    With Fm_Main.MSF_MDBGrid
                        .Redraw = False
                        .AddRow LastRow + 1
                        .CellText(LastRow + 1, 1) = strFilePath
                        .CellTextAlign(LastRow + 1, 1) = DT_LEFT
                        .CellText(LastRow + 1, 2) = strFileCurrrentName
                        .CellTextAlign(LastRow + 1, 2) = DT_LEFT
                        .CellText(LastRow + 1, 3) = GetLastModifcationDate(strFilePath + strFileCurrrentName)
                        .CellTextAlign(LastRow + 1, 3) = DT_CENTER
                        dblMDBOriginalSize = FileLen(strFilePath + strFileCurrrentName)
                        .CellText(LastRow + 1, 4) = dblMDBOriginalSize
                        .CellTextAlign(LastRow + 1, 4) = DT_RIGHT
                        .Redraw = True
                        LastRow = LastRow + 1
                    End With
                End If
            End If
        End If
        strFileCurrrentName = Dir    'Step to next file.
    Loop
End Sub

Sub ScanAllDirectories(UnitéDisque As String)
    Dim colAllDirectories As New Collection
    Dim intNext_Directory As Integer
    Dim strDirectoryName As String
    Dim strSubDirectory As String
    Dim i As Integer

    On Error Resume Next
        LastRow = 0
        
        MousePointer = vbHourglass
        DoEvents
        
        intNext_Directory = 1
        
        'HD unit on which we are working.
        colAllDirectories.Add Left(UnitéDisque, "2")
        Do While intNext_Directory <= colAllDirectories.Count
            'Scans next directory.
            strDirectoryName = colAllDirectories(intNext_Directory)
            intNext_Directory = intNext_Directory + 1
            
            'Reads directories using strDirectoryName.
            strSubDirectory = Dir$(strDirectoryName & "\*", vbDirectory)
            Do While strSubDirectory <> ""
                'Adds the name to the collection colAllDirectories if
                'this is a directory.
                If UCase$(strSubDirectory) <> "PAGEFILE.SYS" And strSubDirectory <> "." And strSubDirectory <> ".." Then
                    strSubDirectory = strDirectoryName & "\" & strSubDirectory
                    On Error Resume Next
                    If GetAttr(strSubDirectory) And vbDirectory Then
                        colAllDirectories.Add strSubDirectory
                        Fm_Main.StatusBar1.Panels.Item(1) = "Scanning: " + strSubDirectory
                    End If
                End If
                strSubDirectory = Dir$(, vbDirectory)
            Loop
        Loop
        
        'Seeks for Access MDB files for each directory.
        '"\" is needed at the end of the directory name to ease up
        'the use of the GetFileInfoFromDirectory function.
        Fm_Main.StatusBar1.Panels.Item(2) = "Scanned dirs: " + CStr(colAllDirectories.Count)
        For i = 1 To colAllDirectories.Count
            Call GetFileInfoFromDirectory(colAllDirectories(i) + "\")
        Next i
        Fm_Main.StatusBar1.Panels.Item(1) = "Scan finished ..."
        Fm_Main.StatusBar1.Panels.Item(2) = "MDB found: " + CStr(Fm_Main.MSF_MDBGrid.Rows)
End Sub

Function GetLastModifcationDate(strFile As String) As Date
    Dim FSO, F
    Dim FileName As String
    
    GetLastModifcationDate = Date       'By default, in case there should be a problem.
    FileName = strFile
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set F = FSO.GetFile(FileName)
    GetLastModifcationDate = F.DateLastModified
End Function

