VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{DE8CE233-DD83-481D-844C-C07B96589D3A}#1.1#0"; "VBALSGRID6.OCX"
Begin VB.Form Fm_Main 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "..."
   ClientHeight    =   8085
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11730
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8085
   ScaleWidth      =   11730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Cmd_Quit 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Quit"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7080
      Width           =   2640
   End
   Begin VB.CommandButton Cmd_PerformOP 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Perform Operation(s)"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7080
      Width           =   2640
   End
   Begin VB.Frame Frm_Operation 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Select an operation to perform ..."
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5040
      TabIndex        =   11
      Top             =   1080
      Width           =   6615
      Begin VB.CheckBox Chk_Compact 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Compact Database"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   360
         Value           =   1  'Checked
         Width           =   3135
      End
      Begin VB.CheckBox Chk_Repair 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Repair Database"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   360
         Value           =   1  'Checked
         Width           =   3135
      End
   End
   Begin VB.CommandButton Cmd_SearchMDB 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Search MDB files"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7080
      Width           =   2640
   End
   Begin VB.Frame Frm_HDUnit 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Select a HD unit ..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   4815
      Begin VB.DriveListBox Drive1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   0
         Top             =   360
         Width           =   4455
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   7710
      Width           =   11730
      _ExtentX        =   20690
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   16193
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   4410
            MinWidth        =   4410
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin vbAcceleratorSGrid6.vbalGrid MSF_MDBGrid 
      Height          =   4695
      Left            =   120
      TabIndex        =   3
      Top             =   2040
      Width           =   11475
      _ExtentX        =   20241
      _ExtentY        =   8281
      RowMode         =   -1  'True
      NoHorizontalGridLines=   -1  'True
      NoVerticalGridLines=   -1  'True
      BackgroundPictureHeight=   768
      BackgroundPictureWidth=   1024
      BackColor       =   5260340
      ForeColor       =   16777215
      HighlightBackColor=   11115663
      HighlightForeColor=   0
      AlternateRowBackColor=   7365722
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HeaderDragReorderColumns=   0   'False
      HeaderHotTrack  =   0   'False
      HeaderFlat      =   -1  'True
      BorderStyle     =   0
      ScrollBarStyle  =   2
      DisableIcons    =   -1  'True
      SelectionOutline=   -1  'True
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   11520
      Y1              =   6840
      Y2              =   6840
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   750
      Left            =   240
      Picture         =   "Fm_Main.frx":0000
      Top             =   120
      Width           =   750
   End
   Begin VB.Label Lbl_Title 
      BackStyle       =   0  'Transparent
      Caption         =   "KhompactAllMDB "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   1080
      TabIndex        =   10
      Top             =   120
      Width           =   7215
   End
   Begin VB.Label Lbl_Title 
      BackStyle       =   0  'Transparent
      Caption         =   "This tutorial will show you how to scan a omplete HD partition, find the Access Databases, and compact them."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   1080
      TabIndex        =   9
      Top             =   480
      Width           =   10575
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Left            =   0
      Top             =   0
      Width           =   11775
   End
End
Attribute VB_Name = "Fm_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
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
' Don't forget to put a reference to Microsoft DAO 3.51 Object library if working
' with MS Office 97 or DAO 3.6 for versions of office over Ms Office 97.
' If you have no idea, then just try DAO 3.51

Private Sub Cmd_FinMDB_Click()

End Sub

Private Sub Cmd_PerformOP_Click()
    Dim i As Double
    Dim dblCompactedSize As Double
    Dim dblMDBOriginalSize As Double
    Dim dblAllOriginalFileSize As Double, dblAllCompactedFileSize As Double
    On Error Resume Next
    
    If Chk_Repair.Value = vbChecked Then
        For i = 1 To MSF_MDBGrid.Rows
            'subRepairDataBase=1 -> Success.
            'subRepairDataBase=0 -> Error.
            If subRepairDataBase(MSF_MDBGrid.cell(i, 1).Text + MSF_MDBGrid.cell(i, 2).Text) Then
                MSF_MDBGrid.Redraw = False
                MSF_MDBGrid.CellText(i, 7) = "Y"
                MSF_MDBGrid.CellTextAlign(i, 7) = DT_CENTER
                MSF_MDBGrid.Redraw = True
            Else
                MSF_MDBGrid.Redraw = False
                MSF_MDBGrid.CellText(i, 7) = "ERR"
                MSF_MDBGrid.CellTextAlign(i, 7) = DT_CENTER
                MSF_MDBGrid.Redraw = True
            End If
        Next i
    End If
    
    
    If Chk_Compact.Value = vbChecked Then
        'We'll use next variables to calculate the space gained.
        dblAllOriginalFileSize = 0
        dblAllCompactedFileSize = 0
        
        For i = 1 To MSF_MDBGrid.Rows
        
            
            'Get the original Size of the MDB file.
            dblMDBOriginalSize = CDbl(MSF_MDBGrid.cell(i, 4).Text)
            dblAllOriginalFileSize = dblAllOriginalFileSize + dblMDBOriginalSize
            'Retrieves the compacted filelen.
            dblCompactedSize = funcCompactDataBase(MSF_MDBGrid.cell(i, 1).Text + MSF_MDBGrid.cell(i, 2).Text)
            dblAllCompactedFileSize = dblAllCompactedFileSize + dblCompactedSize
            
            'Did the compact succeed?
            If dblCompactedSize <> -1 Then
                MSF_MDBGrid.Redraw = False
                MSF_MDBGrid.CellText(i, 5) = dblCompactedSize
                MSF_MDBGrid.CellTextAlign(i, 5) = DT_RIGHT
                MSF_MDBGrid.CellText(i, 6) = Format(100 - (dblCompactedSize * 100) / dblMDBOriginalSize, "00.00") + " %"
                MSF_MDBGrid.CellTextAlign(i, 6) = DT_CENTER
                MSF_MDBGrid.CellText(i, 8) = "Y"
                MSF_MDBGrid.CellTextAlign(i, 8) = DT_CENTER
                MSF_MDBGrid.Redraw = True
            Else
                MSF_MDBGrid.Redraw = False
                MSF_MDBGrid.CellText(i, 5) = "ERR"
                MSF_MDBGrid.CellTextAlign(i, 5) = DT_RIGHT
                MSF_MDBGrid.CellText(i, 6) = "ERR"
                MSF_MDBGrid.CellTextAlign(i, 6) = DT_CENTER
                MSF_MDBGrid.CellText(i, 8) = "N"
                MSF_MDBGrid.CellTextAlign(i, 8) = DT_CENTER
                MSF_MDBGrid.Redraw = True
            End If
        Next i
        
        'Show the space gained by the compact operation.
        StatusBar1.Panels.Item(1) = "Size before operation: " + CStr(dblAllOriginalFileSize) + "  -  Size after operation: " + CStr(dblAllCompactedFileSize)
        StatusBar1.Panels.Item(2) = "Gained:" + CStr(dblAllOriginalFileSize - dblAllCompactedFileSize)
    End If
End Sub

Private Sub Cmd_Quit_Click()
    End
End Sub

Private Sub Cmd_SearchMDB_Click()
    
    On Error Resume Next
    Call ValidateButtons(False)
    Screen.MousePointer = vbHourglass
        Fm_Nag.Lbl_Info.Caption = vbCrLf + "Parsing all the directories and sub-directories on Unit:" + vbCrLf + CStr(Drive1.Drive) + vbCrLf + vbCrLf + "Please be patient ..."
        Fm_Nag.Show
        
        'Clear the grid if there are existing rows.
        If MSF_MDBGrid.Rows > 0 Then
            MSF_MDBGrid.Clear
        End If
        
        Call ScanAllDirectories(Drive1.Drive)
        Unload Fm_Nag
    Screen.MousePointer = vbNormal
    Call ValidateButtons(True)
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Me.Caption = "KhompactAllMDB - V " & App.Major & "." & App.Minor & "." & App.Revision & " (C) 2005 Melantech. "
    Call Setup_MDBGrid
End Sub

Sub Setup_MDBGrid()
    'Setup the VBAccelerator Grid.
    With MSF_MDBGrid
        .AddColumn "Directory", "Directory", ecgHdrTextALignLeft, , 140, True, True, eSortType:=CCLSortString
        .AddColumn "FileName", "File Name", ecgHdrTextALignLeft, , 130, True, True, eSortType:=CCLSortString
        .AddColumn "LastMod", "Last Modified", ecgHdrTextALignCentre, , 100, True, True, eSortType:=CCLSortDate
        .AddColumn "SizeUncompresses", "Size Uncomp.", ecgHdrTextALignRight, , 80, True, True, eSortType:=CCLSortNumeric
        .AddColumn "SizeCompressed", "Size Comp.", ecgHdrTextALignRight, , 80, True, True, eSortType:=CCLSortNumeric
        .AddColumn "Ratio", "Ratio", ecgHdrTextALignCentre, , 80, True, True, eSortType:=CCLSortString
        .AddColumn "Repaired", "Repaired", ecgHdrTextALignCentre, , 70, True, True, eSortType:=CCLSortString
        .AddColumn "Compacted", "Compacted", ecgHdrTextALignCentre, , 70, True, True, eSortType:=CCLSortString
      
        .OwnerDrawImpl = Me
        .RowMode = True
        .GridLines = True
        
        .SelectionAlphaBlend = True
        .SelectionOutline = True
        .DrawFocusRectangle = False
        .HotTrack = True
    End With
End Sub

Private Sub MSF_MDBGrid_ColumnClick(ByVal lCol As Long)
   With MSF_MDBGrid
      If (.ColumnSortOrder(lCol) = CCLOrderAscending) Then
         .ColumnSortOrder(lCol) = CCLOrderDescending
      Else
         .ColumnSortOrder(lCol) = CCLOrderAscending
      End If
      With .SortObject
         .Clear
         .SortColumn(1) = lCol
         .SortOrder(1) = MSF_MDBGrid.ColumnSortOrder(lCol)
         .SortType(1) = MSF_MDBGrid.ColumnSortType(lCol)
      End With
      .Sort
   End With
End Sub

Private Sub MSF_MDBGrid_SelectionChange(ByVal lRow As Long, ByVal lCol As Long)
    StatusBar1.Panels.Item(1) = MSF_MDBGrid.cell(lRow, 1).Text + MSF_MDBGrid.cell(lRow, 2).Text
End Sub

Sub ValidateButtons(boolValidation As Boolean)
    Frm_Operation.Enabled = boolValidation
    Cmd_PerformOP.Enabled = boolValidation
End Sub
