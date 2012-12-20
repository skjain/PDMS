VERSION 5.00
Object = "{2BF72F7D-D367-4712-9146-5533EF3E691A}#1.2#0"; "FraPlus1.ocx"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WELCHBUTTON.OCX"
Begin VB.Form frmDatabase 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8085
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3420
   ScaleWidth      =   8085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picBackup 
      BackColor       =   &H80000009&
      Height          =   2535
      Left            =   480
      Picture         =   "frmDatabase.frx":0000
      ScaleHeight     =   2475
      ScaleWidth      =   1875
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   1935
   End
   Begin FramePlusCtl.FramePlus FramePlus1 
      Height          =   3495
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   6165
      BackgroundPictureAlignment=   5
      BorderStyle     =   10
      BackColorGradient=   16761024
      BackColor       =   16777215
      Style           =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   " "
      Begin VB.PictureBox picNew 
         BackColor       =   &H80000009&
         Height          =   2535
         Left            =   480
         Picture         =   "frmDatabase.frx":0EB1
         ScaleHeight     =   2475
         ScaleWidth      =   1875
         TabIndex        =   6
         Top             =   120
         Width           =   1935
      End
      Begin VB.PictureBox picRestore 
         BackColor       =   &H80000009&
         Height          =   2535
         Left            =   3000
         Picture         =   "frmDatabase.frx":1AE6
         ScaleHeight     =   2475
         ScaleWidth      =   1875
         TabIndex        =   5
         Top             =   120
         Width           =   1935
      End
      Begin VB.PictureBox picSetup 
         BackColor       =   &H80000009&
         Height          =   2535
         Left            =   5640
         Picture         =   "frmDatabase.frx":2A14
         ScaleHeight     =   2475
         ScaleWidth      =   1875
         TabIndex        =   4
         Top             =   120
         Width           =   1935
      End
      Begin WelchButton.lvButtons_H cmdNew 
         Height          =   495
         Left            =   240
         TabIndex        =   0
         Top             =   2760
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   873
         Caption         =   "&New Database"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   8421504
         cFHover         =   8421504
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdRestore 
         Height          =   495
         Left            =   2760
         TabIndex        =   1
         Top             =   2760
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   873
         Caption         =   "&Restore Database"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   255
         cFHover         =   255
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdNodeSetup 
         Height          =   495
         Left            =   5400
         TabIndex        =   2
         Top             =   2760
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   873
         Caption         =   "&Node Setup"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   8388736
         cFHover         =   8388736
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdCancel 
         Height          =   495
         Left            =   7680
         TabIndex        =   8
         Top             =   0
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   873
         Caption         =   "                 &x"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         Image           =   "frmDatabase.frx":399F
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdBackup 
         Height          =   495
         Left            =   240
         TabIndex        =   9
         Top             =   2760
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   873
         Caption         =   "&Backup Database"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   32768
         cFHover         =   32768
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
   End
End
Attribute VB_Name = "frmDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBackup_Click()
    DBTYP = "BACKUP"
    Unload Me
End Sub

Private Sub cmdCancel_Click()
  End
End Sub

Private Sub cmdNodeSetup_Click()
    DBTYP = "SETUP"
    Unload Me
End Sub

Private Sub cmdNew_Click()
    DBTYP = "NEW"
    Unload Me
End Sub

Private Sub cmdRestore_Click()
    DBTYP = "RESTORE"
    Unload Me
End Sub

Private Sub Form_Load()
Dim flag As Boolean
flag = False

If DBTYP = "MENU" Then
   flag = False
Else
   flag = True
End If

  picNew.Visible = flag
  cmdNew.Visible = flag
  picBackup.Visible = Not flag
  Cmdbackup.Visible = Not flag

DBTYP = ""
End Sub
