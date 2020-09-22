VERSION 5.00
Begin VB.Form Fm_Nag 
   BorderStyle     =   0  'None
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3990
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   3990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Lbl_Info 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   3735
   End
   Begin VB.Label Lbl_Title 
      BackStyle       =   0  'Transparent
      Caption         =   "This window display the current operations status."
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
      TabIndex        =   1
      Top             =   480
      Width           =   2655
   End
   Begin VB.Label Lbl_Title 
      BackStyle       =   0  'Transparent
      Caption         =   "Operations Status"
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
      TabIndex        =   0
      Top             =   120
      Width           =   7215
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   750
      Left            =   240
      Picture         =   "Fm_Nag.frx":0000
      Top             =   120
      Width           =   750
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Left            =   0
      Top             =   0
      Width           =   3975
   End
End
Attribute VB_Name = "Fm_Nag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
