VERSION 5.00
Object = "{2BF72F7D-D367-4712-9146-5533EF3E691A}#1.2#0"; "FraPlus1.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WELCHBUTTON.OCX"
Begin VB.Form frmPartyChange 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Billing A/c Party Change Before Auditing"
   ClientHeight    =   7785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8070
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7785
   ScaleWidth      =   8070
   Begin VB.TextBox txtChgRate 
      Height          =   285
      Left            =   6600
      TabIndex        =   41
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Frame FRMBTRM 
      Height          =   2295
      Left            =   7800
      TabIndex        =   33
      Top             =   5400
      Visible         =   0   'False
      Width           =   3975
      Begin VB.TextBox TXTBNET 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   405
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   36
         Text            =   "0.00"
         Top             =   1800
         Width           =   1905
      End
      Begin VB.TextBox txtBEdit 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1440
         TabIndex        =   35
         Top             =   1320
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.TextBox TXTADLS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2040
         TabIndex        =   34
         Text            =   "0.00"
         Top             =   1320
         Visible         =   0   'False
         Width           =   1575
      End
      Begin MSFlexGridLib.MSFlexGrid flexBTRM 
         Height          =   1635
         Left            =   120
         TabIndex        =   37
         Top             =   120
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   2884
         _Version        =   393216
         Rows            =   0
         Cols            =   3
         FixedRows       =   0
         Appearance      =   0
      End
      Begin VB.Label LBLNET 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Net Amount"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   38
         Top             =   1920
         Width           =   1305
      End
   End
   Begin FramePlusCtl.FramePlus FramePlus1 
      Height          =   7995
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   14102
      BackgroundPictureAlignment=   5
      BorderStyle     =   10
      BackColorGradient=   8438015
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
      Begin VB.TextBox TXTCARAT 
         Height          =   285
         Left            =   6600
         Locked          =   -1  'True
         TabIndex        =   49
         Top             =   6360
         Width           =   1215
      End
      Begin VB.TextBox TXTARAT 
         Height          =   285
         Left            =   6600
         Locked          =   -1  'True
         TabIndex        =   47
         Top             =   3960
         Width           =   1215
      End
      Begin VB.TextBox TXTTQTY 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5760
         TabIndex        =   46
         Top             =   6960
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox TXTTPCS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6000
         TabIndex        =   45
         Top             =   7200
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox TXTITOT 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5400
         TabIndex        =   44
         Text            =   "0.00"
         Top             =   6720
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox txtRate 
         Height          =   285
         Left            =   6600
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   3600
         Width           =   1215
      End
      Begin VB.ComboBox M_RTTX 
         Height          =   315
         ItemData        =   "frmPartyChange.frx":0000
         Left            =   1680
         List            =   "frmPartyChange.frx":000A
         TabIndex        =   12
         Text            =   "M_RTTX"
         Top             =   6360
         Width           =   2295
      End
      Begin VB.TextBox TXTCHGBRNM 
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   6000
         Width           =   3975
      End
      Begin VB.TextBox TXTRETTAX 
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   3960
         Width           =   3975
      End
      Begin VB.TextBox M_BRNM 
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   3600
         Width           =   3975
      End
      Begin VB.TextBox TXTCHGDADD 
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   10
         Top             =   5640
         Width           =   6135
      End
      Begin VB.TextBox TXTADDRESS 
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   3240
         Width           =   6135
      End
      Begin VB.TextBox TXTCHGDPTY 
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   5280
         Width           =   6135
      End
      Begin VB.TextBox TXTCHGPTY 
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   4920
         Width           =   6135
      End
      Begin VB.TextBox txtAccoutParty 
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   2520
         Width           =   6135
      End
      Begin VB.TextBox txtConsignee 
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   2880
         Width           =   6135
      End
      Begin VB.TextBox TXTVBNO 
         Height          =   285
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   1
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox TXTSALTYP 
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   1080
         Width           =   4485
      End
      Begin VB.PictureBox Search 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   300
         Left            =   3720
         Picture         =   "frmPartyChange.frx":002B
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   2
         Top             =   1440
         Width           =   300
      End
      Begin WelchButton.lvButtons_H cmdCancel 
         Height          =   495
         Left            =   1800
         TabIndex        =   15
         Top             =   6840
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   873
         Caption         =   "&Cancel"
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
         Image           =   "frmPartyChange.frx":05B5
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdExit 
         Height          =   495
         Left            =   3000
         TabIndex        =   16
         Top             =   6840
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   873
         Caption         =   "E&xit"
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
         Image           =   "frmPartyChange.frx":0A07
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdChange 
         Height          =   495
         Left            =   600
         TabIndex        =   13
         Top             =   6840
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   873
         Caption         =   "&Change"
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
         Image           =   "frmPartyChange.frx":0FA1
         cBack           =   -2147483633
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Ass. Rate"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   5760
         TabIndex        =   50
         Top             =   6360
         Width           =   855
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Ass. Rate"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   5760
         TabIndex        =   48
         Top             =   3960
         Width           =   855
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Rate"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   5760
         TabIndex        =   43
         Top             =   6000
         Width           =   735
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   6840
         TabIndex        =   42
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Rate"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   5760
         TabIndex        =   40
         Top             =   3600
         Width           =   615
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Note : User can change details of invoice, if it is made by without order."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   240
         TabIndex        =   32
         Tag             =   "S"
         Top             =   7440
         Width           =   7335
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Retail/Tax"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   600
         TabIndex        =   31
         Tag             =   "S"
         Top             =   6360
         Width           =   975
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Agent "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   600
         TabIndex        =   30
         Top             =   6000
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "A/C Party"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   600
         TabIndex        =   29
         Tag             =   "S"
         Top             =   4920
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Consignee "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   600
         TabIndex        =   28
         Tag             =   "S"
         Top             =   5280
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   600
         TabIndex        =   27
         Tag             =   "S"
         Top             =   5640
         Width           =   855
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Retail/Tax"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   600
         TabIndex        =   26
         Tag             =   "S"
         Top             =   3960
         Width           =   975
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Agent "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   600
         TabIndex        =   25
         Top             =   3600
         Width           =   1095
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H008080FF&
         Caption         =   "Change Invoice Detail"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   2520
         TabIndex        =   24
         Tag             =   "S"
         Top             =   4560
         Width           =   2415
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H008080FF&
         Caption         =   "Existing Invoice Detail"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   2520
         TabIndex        =   23
         Tag             =   "S"
         Top             =   2160
         Width           =   2295
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   600
         TabIndex        =   22
         Tag             =   "S"
         Top             =   3240
         Width           =   855
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         Height          =   2055
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   4680
         Width           =   7815
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice &Type"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   21
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Consignee "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   600
         TabIndex        =   20
         Tag             =   "S"
         Top             =   2880
         Width           =   975
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         Height          =   2055
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   2280
         Width           =   7815
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "A/C Party"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   600
         TabIndex        =   19
         Tag             =   "S"
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         Height          =   1095
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   840
         Width           =   7815
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Sale &Bill No."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   18
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label lblAlert 
         BackStyle       =   0  'Transparent
         Caption         =   "Billing A/c Party Change"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   2280
         TabIndex        =   17
         Top             =   240
         Width           =   4095
      End
   End
End
Attribute VB_Name = "frmPartyChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim M_OPER(0 To 15) As String
Dim M_PERC(0 To 15) As Double
Dim M_POSTCOD(0 To 15) As String
Dim M_NICK(0 To 15) As String
Dim M_POSTYESNO(0 To 15) As String
Dim M_FMLA(0 To 15) As String
Dim M_RDOF(0 To 15) As String
Dim TXCD As String
Dim M_QTY As Double
Dim M_TXNM As String

Private Sub cmdChange_Click()
  On Error GoTo LAST
  Dim BILLRS As ADODB.Recordset
  Set BILLRS = New ADODB.Recordset
  
  Dim TMPRS As ADODB.Recordset
  Set TMPRS = New ADODB.Recordset
  
  Dim PCOD As String, DCOD As String, DADD As String, BRCD As String
         
      'FIND SALE TYPE CODE
      If TMPRS.State = 1 Then TMPRS.Close
      TMPRS.Open "SELECT CODE FROM SERIALMASTER WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
      "' AND NAME='" & TXTSALTYP & "' AND VTYP='SAL' AND FYCD='" & FYCD & "'", CN, adOpenDynamic, adLockOptimistic
      
      If TMPRS.EOF Then
        MsgBox "Invalid Sale Type", vbCritical: TXTSALTYP.SetFocus: Exit Sub
      Else
        TXTSALTYP.Tag = TMPRS!CODE
      End If
      TMPRS.Close
      '--------------------
      
      'CHECK VALLID BILL IS READY TO DELETE
      If InvalidBill Then Exit Sub
      '-------------------------------------
      
      'PARTY
      If TMPRS.State = 1 Then TMPRS.Close
      TMPRS.Open "SELECT CODE FROM ACCMST WHERE NAME='" & TXTCHGPTY & "' ", CN, adOpenDynamic, adLockOptimistic
      If Not TMPRS.EOF Then
         PCOD = Trim(TMPRS!CODE & "")
      Else
         MsgBox "Party Can't be Empty", vbCritical
         If TXTCHGPTY.Enabled Then TXTCHGPTY.SetFocus
         Exit Sub
      End If
      TMPRS.Close
      
      'Consignee
      If TMPRS.State = 1 Then TMPRS.Close
      TMPRS.Open "SELECT CODE,SRNO FROM PADDMST WHERE NAME ='" & TXTCHGDPTY & "' AND ADDR ='" & TXTCHGDADD & "' ", CN, adOpenDynamic, adLockOptimistic
      If Not TMPRS.EOF Then
         DCOD = Trim(TMPRS!CODE & "")
         DADD = Trim(TMPRS!SRNO & "")
      Else
         MsgBox "Delivery Party Can't be Empty", vbCritical
         If TXTCHGDPTY.Enabled Then TXTCHGDPTY.SetFocus
         Exit Sub
      End If
      TMPRS.Close
      
      'Agent
      If TMPRS.State = 1 Then TMPRS.Close
      TMPRS.Open "SELECT CODE FROM REFMST WHERE NAME='" & TXTCHGBRNM & "' ", CN, adOpenDynamic, adLockOptimistic
      If Not TMPRS.EOF Then
         BRCD = Trim(TMPRS!CODE & "")
      Else
         MsgBox "Broker Can't be Empty", vbCritical
         If TXTCHGBRNM.Enabled Then TXTCHGBRNM.SetFocus
         Exit Sub
      End If
      TMPRS.Close
      
      'USE CONFIRMATION
      Dim AYS
      AYS = MsgBox("Are You sure to Change Detail of selected Invoice ? ", vbYesNo)
      If AYS = VBNO Then Exit Sub
      '------------------
      'OPERATION BEGIN
      CN.BeginTrans
                        
      'DELETE INVOICE
'      SQL = "UPDATE BILLMAIN SET PCOD = '" & PCOD & "',DRAC = '" & PCOD & "',DCOD= '" & DCOD & _
'            "',ADDRESS = '" & DADD & "',BRCD = '" & BRCD & "',TTYP = '" & M_RTTX & "' WHERE COMP='" & compPth & _
'            "' AND UNIT='" & UNCD & "' AND VTYP='SAL' AND DBCD='" & TXTSALTYP.Tag & _
'            "' AND VBNO='" & TXTVBNO.Text & "' AND RECSTAT<>'D'"
'
'      CN.Execute SQL
      
'      SQL = "UPDATE EGPMAN SET DRAC = '" & PCOD & "',DCOD = '" & DCOD & "',BRCD = '" & BRCD & _
'      "',RORT = '" & M_RTTX & "' WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
'      "' AND VTYP='SAL' AND DBCD='" & TXTSALTYP.Tag & "' AND VBNO='" & TXTVBNO.Text & "' AND RECSTAT<>'D'"
'
'      CN.Execute SQL
      
      Dim TEMPRS As New ADODB.Recordset
      If TEMPRS.State = 1 Then TEMPRS.Close
      TEMPRS.Open "SELECT *FROM BILLMAIN WHERE COMP='" & compPth & _
            "' AND UNIT='" & UNCD & "' AND VTYP='SAL' AND DBCD='" & TXTSALTYP.Tag & _
            "' AND VBNO='" & TXTVBNO.Text & "' AND RECSTAT<>'D'", CN, adOpenDynamic, adLockOptimistic
      If Not TEMPRS.EOF Then
            TEMPRS!PCOD = Trim(PCOD & "")
            TEMPRS!DRAC = Trim(PCOD & "")
            TEMPRS!DCOD = Trim(DCOD & "")
            TEMPRS!ADDRESS = Trim(DADD & "")
            TEMPRS!BRCD = Trim(BRCD & "")
            TEMPRS!TTYP = Trim(M_RTTX & "")
            TEMPRS!ITOT = Val(TXTITOT)
            TEMPRS!BADJ = Val(TXTBNET) - Val(TXTITOT)
            TEMPRS!BNET = Val(TXTBNET)
            
            i = 0
            For i = 0 To flexBTRM.Rows - 1
              J = 0
              For J = 0 To TEMPRS.Fields.COUNT - 1
                If Trim(TEMPRS.Fields(J).NAME) = Trim(flexBTRM.TextMatrix(i, 0)) Then
                  TEMPRS.Fields(J).Value = Val(flexBTRM.TextMatrix(i, 2))
                End If
                If Trim(TEMPRS.Fields(J).NAME) = "PER" & Trim(flexBTRM.TextMatrix(i, 0)) Then
                  TEMPRS.Fields(J).Value = Val(flexBTRM.TextMatrix(i, 1))
                End If
              Next
            Next
              
            TEMPRS.Update
      End If
      
      If TEMPRS.State = 1 Then TEMPRS.Close
      TEMPRS.Open "SELECT *FROM EGPMAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
            "' AND VTYP='SAL' AND DBCD='" & TXTSALTYP.Tag & "' AND VBNO='" & _
            TXTVBNO.Text & "' AND RECSTAT<>'D'", CN, adOpenDynamic, adLockOptimistic
      If Not TEMPRS.EOF Then
            TEMPRS!DRAC = Trim(PCOD & "")
            TEMPRS!DCOD = Trim(DCOD & "")
            TEMPRS!BRCD = Trim(BRCD & "")
            TEMPRS!RORT = Trim(M_RTTX & "")
            TEMPRS!AMNT = Val(TXTITOT)
            TEMPRS!ITOT = Val(TXTITOT)
            TEMPRS!BADJ = Val(TXTBNET) - Val(TXTITOT)
            TEMPRS!BNET = Val(TXTBNET)
            
            i = 0
            For i = 0 To flexBTRM.Rows - 1
              J = 0
              For J = 0 To TEMPRS.Fields.COUNT - 1
                If Trim(TEMPRS.Fields(J).NAME) = Trim(flexBTRM.TextMatrix(i, 0)) Then
                  TEMPRS.Fields(J).Value = Val(flexBTRM.TextMatrix(i, 2))
                End If
                If Trim(TEMPRS.Fields(J).NAME) = "PER" & Trim(flexBTRM.TextMatrix(i, 0)) Then
                  TEMPRS.Fields(J).Value = Val(flexBTRM.TextMatrix(i, 1))
                End If
              Next
            Next
              
            TEMPRS.Update
      End If
      
      SQL = "UPDATE SPTRAN SET DRAC = '" & PCOD & "',PCOD = '" & PCOD & "',DCOD = '" & DCOD & "',ADDRESS= '" & DADD & _
            "',BRCD = '" & BRCD & "',TXRT = '" & M_RTTX & "',RATE=" & GetReverseRate(Trim(M_BRNM), Trim(M_TXNM), Val(txtChgRate), 0) & _
            ",ARAT=" & Val(txtChgRate) & ",AMNT=" & Val(TXTITOT) & _
            " WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND VTYP='SAL' " & _
            " AND DBCD='" & TXTSALTYP.Tag & "' AND VBNO='" & TXTVBNO.Text & "' AND RECSTAT<>'D'"
            
      CN.Execute SQL
      
      SQL = "UPDATE STORETRAN SET DRAC = '" & PCOD & "',PCOD = '" & PCOD & "',DCOD = '" & DCOD & _
            "',RATE=" & Val(TXTCARAT) & ",AMNT=" & Val(TXTITOT) & " WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
            "' AND VTYP='SAL' AND DBCD='" & TXTSALTYP.Tag & "' AND VBNO='" & TXTVBNO.Text & "' AND RECSTAT<>'D'"
            
      CN.Execute SQL
          
      'challan
      SQL = "UPDATE SPTRAN SET DRAC = '" & PCOD & "',PCOD = '" & PCOD & "',DCOD = '" & DCOD & "',ADDRESS= '" & DADD & _
            "',BRCD = '" & BRCD & "',TXRT = '" & M_RTTX & "',RATE=" & Val(txtChgRate) & _
            ",AMNT=" & Val(M_QTY) * Val(txtChgRate) & " WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND RTYP='SAL' " & _
            " AND SDBC='" & TXTSALTYP.Tag & "' AND SVBN='" & TXTVBNO.Text & "' AND RECSTAT<>'D'"
            
      CN.Execute SQL
                            
      'OPERATION FINISH SUCCESSFULLY
      If RS.State = 1 Then RS.Close
      RS.Open "SELECT * FROM DAILYSTAT WHERE 1=2", CN, adOpenDynamic, adLockOptimistic
      RS.AddNew
      RS!COMP = compPth
      RS!VTYP = "SAL"
      RS!PCOD = ""
      RS!SRNO = "Change"
      RS!VBNO = Trim(TXTVBNO)
      RS!AMNT = 0
      RS!CUSR = cUName
      RS!QNTY = 0
      RS!ACTN = "D"
      RS!DTTM = Format(Now, "YYYY/MM/DD HH:MM:SS")
      RS.Update
      
      CN.CommitTrans
            
      MsgBox "Bill Detail Changed Successfuly "
      
  Call cmdCancel_Click
  Exit Sub
  
LAST:

  MsgBox ERR.Description
  CN.RollbackTrans
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
   Call CenterChild(frm_Main, Me)
   If M_RTTX.ListCount > 0 Then M_RTTX.ListIndex = 0
End Sub

Private Sub M_RTTX_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDelete Then
     KeyCode = 0
  End If
End Sub

Private Sub M_RTTX_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub Search_Click()
   Call GenDetails
End Sub

Private Sub TXTCHGBRNM_GotFocus()
TXTCHGBRNM.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTCHGBRNM_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF2 Then
        NEW_VISIBLE = False
        CANCEL_VISIBLE = True
        M_DESC = Empty
        TXTCHGBRNM = SearchList1("Select  TOP 20 Code,Name From REFMST WHERE CATA='B'", 0, Empty, "Select Dealer From List")
        TXTCHGBRNM.Tag = Key
    ElseIf KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
        TXTCHGBRNM = Empty
        TXTCHGBRNM.Tag = Empty
    End If
End Sub

Private Sub TXTCHGBRNM_LostFocus()
  TXTCHGBRNM.BackColor = vbWhite
End Sub

Private Sub TXTCHGDADD_GotFocus()
  TXTCHGDADD.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTCHGDADD_KeyDown(KeyCode As Integer, Shift As Integer)
   TXTCHGDADD.FontSize = 8
   If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
    TXTCHGDADD = Empty
   ElseIf KeyCode = vbKeyF2 Then
    TXTCHGDADD = SearchAddress("Select SRNO,ADDR From PADDMST WHERE NAME='" & TXTCHGDPTY & "'", 0, Empty, "Select A/c Party Filtered by Party Group")
   End If
End Sub

Private Sub TXTCHGDADD_LostFocus()
  TXTCHGDADD.BackColor = vbWhite
End Sub

Private Sub TXTCHGDPTY_Change()
  TXTCHGDADD = Empty
End Sub

Private Sub TXTCHGDPTY_GotFocus()
  TXTCHGDPTY.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTCHGDPTY_KeyDown(KeyCode As Integer, Shift As Integer)
 Me.KeyPreview = False
  If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
     TXTCHGDPTY = Empty
  ElseIf KeyCode = vbKeyF2 Then
     M_DESC = Empty:   NEW_VISIBLE = False
     TXTCHGDPTY = SearchList1("Select DISTINCT CODE,NAME From PADDMST WHERE RECSTAT='A'", 0, Empty, "Select Consignee Name ")
  End If
 Me.KeyPreview = True
End Sub

Private Sub TXTCHGDPTY_LostFocus()
  TXTCHGDPTY.BackColor = vbWhite
End Sub

Private Sub TXTCHGPTY_GotFocus()
TXTCHGPTY.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTCHGPTY_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        NEW_VISIBLE = False: CANCEL_VISIBLE = True:  M_DESC = Empty
        TXTCHGPTY = SearchList1("Select  TOP 20 Code,Name From ACCMST", 0, Empty, "Select A/c Party From List")
        TXTCHGPTY.Tag = Key
    ElseIf KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
        TXTCHGPTY = Empty
        TXTCHGPTY.Tag = Empty
    End If
End Sub

Private Sub TXTCHGPTY_LostFocus()
  TXTCHGPTY.BackColor = vbWhite
End Sub

Private Sub TXTCHGRETTAX_GotFocus()
  TXTCHGRETTAX.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTCHGRETTAX_LostFocus()
  TXTCHGRETTAX.BackColor = vbWhite
End Sub

Private Sub txtChgRate_Change()
     If Val(txtChgRate) > 0 Then
        TXTCARAT = GetReverseRate(Trim(M_BRNM), Trim(M_TXNM), Val(txtChgRate), 0)
        TXTITOT = Format((Val(M_QTY) * Val(TXTCARAT)), "##########.00")
        'Call FIL_Billingterm
        calBTRM 0
        Call calADLS
     End If
End Sub

Private Sub txtChgRate_LostFocus()
    If Val(txtChgRate) > 0 Then
        TXTCARAT = GetReverseRate(Trim(M_BRNM), Trim(M_TXNM), Val(txtChgRate), 0)
        TXTITOT = Format(FormatNumber(Val(M_QTY) * Val(TXTCARAT), 0), "##########.00")
        'Call FIL_Billingterm
        calBTRM 0
        Call calADLS
     End If
End Sub

Private Sub TXTSALTYP_GotFocus()
 TXTSALTYP.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTSALTYP_KeyDown(KeyCode As Integer, Shift As Integer)
  If (KeyCode = vbKeyReturn And Trim(TXTSALTYP) = Empty) Or KeyCode = vbKeyF2 Then
    TXTSALTYP.Text = SearchList1("SELECT CODE,NAME FROM SERIALMASTER WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
                     "' AND VTYP='SAL' AND FYCD='" & FYCD & "' AND ACTIVE='Y'", 0, TXTSALTYP.Text, "SELECT INVOICE TYPE FROM LIST")
    TXTSALTYP.Tag = Key
    Call FindLastBill
  End If
End Sub

Private Sub TXTSALTYP_LostFocus()
 TXTSALTYP.BackColor = vbWhite
End Sub

Private Sub TXTVBNO_GotFocus()
 TXTVBNO.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtVBNO_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And TXTVBNO <> Empty And Len(TXTVBNO) = 10 Then
    Call GenDetails
    If TXTCHGPTY.Enabled Then TXTCHGPTY.SetFocus
End If
End Sub

Private Sub TXTVBNO_LostFocus()
 TXTVBNO.BackColor = vbWhite
End Sub

Private Sub GenDetails()
  On Error GoTo LAST
  Dim M_DCOD As String, M_ADDR As String
  Dim ORDEXC As Boolean
  ORDEXC = False
  Dim BILLRS As New ADODB.Recordset
  Set BILLRS = New ADODB.Recordset
  
  SQL = "SELECT * FROM BILLMAIN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND VTYP='SAL' " & _
        " AND DBCD='" & TXTSALTYP.Tag & "' AND VBNO='" & TXTVBNO.Text & "' AND RECSTAT<>'D' AND DVCD<>'000001'"
  
  If BILLRS.State = 1 Then BILLRS.Close
  BILLRS.Open SQL, CN, adOpenDynamic, adLockOptimistic
  
  If BILLRS.EOF Then
    MsgBox "Invalid Bill No."
    TXTVBNO.SetFocus
    Exit Sub
  End If
          
  If RS.State = 1 Then RS.Close
  RS.Open "SELECT * FROM TAXMST WHERE CODE='" & BILLRS!TXCD & "'", CN, adOpenDynamic, adLockOptimistic
  If Not RS.EOF Then
     TXCD = RS!CODE
     M_TXNM = Trim(RS!NAME & "")
  End If
  RS.Close

  If Not BILLRS.EOF Then
        TXTITOT = Format(BILLRS!ITOT, "########.00")
        TXTBNET = Format(BILLRS!BNET, "########.00")
        TXTTQTY = Format(BILLRS!TQTY, "########.000")
        TXTTPCS = Format(BILLRS!TPCS, "########")
        M_QTY = Format(BILLRS!TQTY, "########.000")
        Call FIL_Billingterm
    
        Dim i As Double
        Dim J As Double
        i = 0
        For i = 0 To flexBTRM.Rows - 1
            J = 0
            For J = 0 To BILLRS.Fields.COUNT - 1
                If Trim(BILLRS.Fields(J).NAME) = Trim(flexBTRM.TextMatrix(i, 0)) Then
                    flexBTRM.TextMatrix(i, 2) = Format(BILLRS.Fields(J).Value, "#########.00")
                End If
                If Trim(BILLRS.Fields(J).NAME) = "PER" & Trim(flexBTRM.TextMatrix(i, 0)) Then
                    flexBTRM.TextMatrix(i, 1) = Format(BILLRS.Fields(J).Value, "######.00")
                End If
            Next
        Next
  End If

  Dim SPRS As New ADODB.Recordset
  If SPRS.State = 1 Then SPRS.Close
  SPRS.Open "SELECT * FROM SPTRAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND VTYP='SAL' " & _
        " AND DBCD='" & TXTSALTYP.Tag & "' AND VBNO='" & TXTVBNO.Text & "' AND RECSTAT<>'D' AND DVCD<>'000001'", CN, adOpenDynamic, adLockOptimistic
  If Not SPRS.EOF Then
     TXTRATE = Format(SPRS!ARAT, "#########.000")
     txtChgRate = Format(SPRS!ARAT, "#########.000")
     TXTARAT = Format(SPRS!RATE, "#########.000")
     TXTCARAT = Format(SPRS!RATE, "#########.000")
  End If
  SPRS.Close
  
  If RS.State = 1 Then RS.Close
  RS.Open "SELECT NAME FROM ACCMST WHERE CODE='" & BILLRS!PCOD & "'", CN, adOpenDynamic, adLockOptimistic
  If Not RS.EOF Then
     txtAccoutParty = Trim(RS!NAME & "")
     TXTCHGPTY = Trim(RS!NAME & "")
     TXTCHGPTY.Tag = Trim(BILLRS!PCOD & "")
  End If
  RS.Close
  
  If RS.State = 1 Then RS.Close
  RS.Open "SELECT NAME FROM REFMST WHERE CODE='" & BILLRS!BRCD & "'", CN, adOpenDynamic, adLockOptimistic
  If Not RS.EOF Then
      TXTCHGBRNM.Tag = Trim(BILLRS!BRCD & "")
     M_BRNM = Trim(RS!NAME & "")
     TXTCHGBRNM = Trim(RS!NAME & "")
  End If
  RS.Close
  
  TXTRETTAX = Trim(BILLRS!TTYP & "")
  M_RTTX = Trim(BILLRS!TTYP & "")
  
  If RS.State = 1 Then RS.Close
  RS.Open "SELECT DCOD,ADDRESS,ISNULL(EXTRA1,'') AS ORDN FROM SPTRAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND VTYP='SAL' " & _
          " AND DBCD='" & TXTSALTYP.Tag & "' AND VBNO='" & TXTVBNO.Text & "' AND RECSTAT<>'D' AND DVCD<>'000001'", CN, adOpenDynamic, adLockOptimistic
  If Not RS.EOF Then
     M_DCOD = Trim(RS!DCOD & "")
     M_ADDR = Trim(RS!ADDRESS & "")
     
     If Trim(RS!ORDN & "") <> Empty Then
        ORDEXC = True
     End If
     
     If RS.State = 1 Then RS.Close
     RS.Open "SELECT NAME,ADDR FROM PADDMST WHERE CODE='" & BILLRS!DCOD & "' AND SRNO='" & M_ADDR & "'", CN, adOpenDynamic, adLockOptimistic
     If Not RS.EOF Then
        txtConsignee = Trim(RS!NAME & "")
        TXTADDRESS = Trim(RS!ADDR & "")
        
        TXTCHGDPTY = Trim(RS!NAME & "")
        TXTCHGDADD = Trim(RS!ADDR & "")
     End If
     RS.Close
  End If
  
  If ORDEXC Then
     MsgBox "Current Sale Bill made by Sales Order Contract", vbCritical, "No Permission to Change"
     Call cmdCancel_Click
  End If
    
Exit Sub
LAST:

  MsgBox ERR.Description
End Sub

Private Sub FindLastBill()
On Error GoTo LAST
Dim FINDRS As ADODB.Recordset
Set FINDRS = New ADODB.Recordset

SQL = "SELECT SRNO FROM SERIALMASTER WHERE COMP='" & compPth & "' AND UNIT = '" & UNCD & "' AND VTYP='SAL' " & _
      "AND CODE='" & TXTSALTYP.Tag & "' AND FYCD='" & FYCD & "'"

If FINDRS.State = 1 Then FINDRS.Close
FINDRS.Open SQL, CN, adOpenDynamic, adLockOptimistic
If Not FINDRS.EOF Then
   TXTVBNO = FINDRS!SRNO & ""
Else
   TXTVBNO = Empty
End If
Exit Sub
LAST:
MsgBox ERR.Description
End Sub

Private Sub cmdCancel_Click()
    'TXTSALTYP = Empty: TXTSALTYP.Tag = Empty: txtAccoutParty = Empty: txtConsignee = Empty
    'TXTADDRESS = Empty: TXTVBNO = Empty: TXTSALTYP.SetFocus
    'Call CLEARDATA
    Call ClsData(Me)
    If TXTSALTYP.Enabled Then TXTSALTYP.SetFocus
End Sub

Private Sub CMDEXIT_Click()
  Unload Me
End Sub

Private Function InvalidBill() As Boolean
InvalidBill = False 'CONSIDER IT IS A VALID BILL

Dim VALIDRS As ADODB.Recordset
 Set VALIDRS = New ADODB.Recordset

 'CASE :1 : EXIST BILL AND NOT HAVING ANY PAYMENT RECEIPT ENTRY
 SQL = "SELECT * FROM BILLMAIN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
 "' AND VTYP='SAL' AND DBCD='" & TXTSALTYP.Tag & "' AND VBNO='" & TXTVBNO.Text & "' AND RECSTAT<>'D'"
 
 If VALIDRS.State = 1 Then VALIDRS.Close
 VALIDRS.Open SQL, CN, adOpenDynamic, adLockOptimistic
 If VALIDRS.EOF Then
    InvalidBill = True  'PROVE INVALID BILL
    MsgBox "Invoice not Audited till date", vbCritical
    TXTVBNO.SetFocus
    Exit Function
  Else
    
    If UCase(Trim(VALIDRS!BSTS & "")) = "A" Then
       InvalidBill = True
       MsgBox "Further Entry Exist", vbCritical, "Sale Bill is Audited"
       TXTVBNO.SetFocus
       Exit Function
    End If
    
    If RS.State = 1 Then RS.Close
    RS.Open "SELECT COMP FROM RPTRAN WHERE COMP='" & VALIDRS!COMP & "' AND BSR1='" & VALIDRS!VTYP & _
            "' AND SVBN='" & VALIDRS!VBNO & "' AND SDBC='" & VALIDRS!dbcd & _
            "' AND UNIT='" & VALIDRS!unit & "'", CN, adOpenDynamic, adLockOptimistic
    If RS.EOF = False Then
       InvalidBill = True
       MsgBox "Further Entry Exist Can not modified"
       TXTVBNO.SetFocus
       Exit Function
    End If
 End If
End Function
 
Private Sub calADLS()
    Dim P As Integer
    TXTADLS.Text = Empty
    For P = 0 To flexBTRM.Rows - 1
        If M_OPER(P) = "-" Then
            TXTADLS.Text = Format(Val(TXTADLS.Text) - Val(flexBTRM.TextMatrix(P, 2)), "############.00")
        Else
            TXTADLS.Text = Format(Val(TXTADLS.Text) + Val(flexBTRM.TextMatrix(P, 2)), "############.00")
        End If
    Next P
    If M_BILRDOF = "Y" Then
        TXTBNET.Text = Format(FormatNumber(Val(TXTITOT.Text) + Val(TXTADLS.Text), 0), "##########.00")
    Else
        TXTBNET.Text = Format(Val(TXTITOT.Text) + Val(TXTADLS.Text), "##########.00")
    End If
End Sub


Private Sub calBTRM(ByVal ICTR As Integer)
    Dim J As Integer, iFMLA(20) As Double, subTot As Double
    Dim c_FMLA(20) As String
    Dim L As Integer
    Dim m As Integer
    Dim B() As String
    subTot = 0
    Dim a() As String, K As Integer
    J = 0
    If flexBTRM.Rows = 0 Then
      Exit Sub
    End If
    For J = flexBTRM.ROW To flexBTRM.Rows - 1
        If Val(flexBTRM.TextMatrix(J, 1)) <> 0 Then flexBTRM.TextMatrix(J, 2) = 0
    Next J

    For J = flexBTRM.ROW To flexBTRM.Rows - 1
        c_FMLA(J) = Trim(M_FMLA(J))
        If Len(c_FMLA(J)) <= 6 Then
            Select Case c_FMLA(J)
                Case "M_STOT"
                    If Val(flexBTRM.TextMatrix(J, 1)) <> 0 Then
                        flexBTRM.TextMatrix(J, 2) = Format((Val(flexBTRM.TextMatrix(J, 1)) * Val(TXTITOT.Text)) / 100, "##########.000")
                    Else
                        flexBTRM.TextMatrix(J, 2) = Format(flexBTRM.TextMatrix(J, 2), "#.000")
                    End If
                Case "M_TQTY"
                    If Val(flexBTRM.TextMatrix(J, 1)) <> 0 Then
                        flexBTRM.TextMatrix(J, 2) = Format((Val(flexBTRM.TextMatrix(J, 1)) * Val(TXTTQTY.Text)), "##########.000")
                    Else
                        flexBTRM.TextMatrix(J, 2) = Format(flexBTRM.TextMatrix(J, 2), "#.000")
                    End If
                Case "M_TPCS"
                    If Val(flexBTRM.TextMatrix(J, 1)) <> 0 Then
                        flexBTRM.TextMatrix(J, 2) = Format((Val(flexBTRM.TextMatrix(J, 1)) * Val(TXTTPCS.Text)), "##########.000")
                    Else
                        flexBTRM.TextMatrix(J, 2) = Format(flexBTRM.TextMatrix(J, 2), "#.000")
                    End If
                Case "AMT_01"
                    If Val(flexBTRM.TextMatrix(J, 1)) <> 0 Then
                        flexBTRM.TextMatrix(J, 2) = Format((Val(flexBTRM.TextMatrix(J, 1)) * Val(flexBTRM.TextMatrix(0, 2))) / 100, "##########.000")
                    Else
                        flexBTRM.TextMatrix(J, 2) = Format(flexBTRM.TextMatrix(J, 2), "#.000")
                    End If
                Case "AMT_02"
                    If Val(flexBTRM.TextMatrix(J, 1)) <> 0 Then
                        flexBTRM.TextMatrix(J, 2) = Format((Val(flexBTRM.TextMatrix(J, 1)) * Val(flexBTRM.TextMatrix(1, 2))) / 100, "##########.000")
                    Else
                        flexBTRM.TextMatrix(J, 2) = Format(flexBTRM.TextMatrix(J, 2), "#.000")
                    End If
                Case "AMT_03"
                    If Val(flexBTRM.TextMatrix(J, 1)) <> 0 Then
                        flexBTRM.TextMatrix(J, 2) = Format((Val(flexBTRM.TextMatrix(J, 1)) * Val(flexBTRM.TextMatrix(2, 2))) / 100, "##########.000")
                    Else
                        flexBTRM.TextMatrix(J, 2) = Format(flexBTRM.TextMatrix(J, 2), "#.000")
                    End If
                Case "AMT_04"
                    If Val(flexBTRM.TextMatrix(J, 1)) <> 0 Then
                        flexBTRM.TextMatrix(J, 2) = Format((Val(flexBTRM.TextMatrix(J, 1)) * Val(flexBTRM.TextMatrix(3, 2))) / 100, "##########.000")
                    Else
                        flexBTRM.TextMatrix(J, 2) = Format(flexBTRM.TextMatrix(J, 2), "#.000")
                    End If
                Case "AMT_05"
                    If Val(flexBTRM.TextMatrix(J, 1)) <> 0 Then
                        flexBTRM.TextMatrix(J, 2) = Format((Val(flexBTRM.TextMatrix(J, 1)) * Val(flexBTRM.TextMatrix(4, 2))) / 100, "##########.000")
                    Else
                        flexBTRM.TextMatrix(J, 2) = Format(flexBTRM.TextMatrix(J, 2), "#.000")
                    End If
                Case "AMT_06"
                    If Val(flexBTRM.TextMatrix(J, 1)) <> 0 Then
                        flexBTRM.TextMatrix(J, 2) = Format((Val(flexBTRM.TextMatrix(J, 1)) * Val(flexBTRM.TextMatrix(5, 2))) / 100, "##########.000")
                    Else
                        flexBTRM.TextMatrix(J, 2) = Format(flexBTRM.TextMatrix(J, 2), "#.000")
                    End If
                Case "AMT_07"
                    If Val(flexBTRM.TextMatrix(J, 1)) <> 0 Then
                        flexBTRM.TextMatrix(J, 2) = Format((Val(flexBTRM.TextMatrix(J, 1)) * Val(flexBTRM.TextMatrix(6, 2))) / 100, "##########.000")
                    Else
                        flexBTRM.TextMatrix(J, 2) = Format(flexBTRM.TextMatrix(J, 2), "#.000")
                    End If
                Case "AMT_08"
                    If Val(flexBTRM.TextMatrix(J, 1)) <> 0 Then
                        flexBTRM.TextMatrix(J, 2) = Format((Val(flexBTRM.TextMatrix(J, 1)) * Val(flexBTRM.TextMatrix(7, 2))) / 100, "##########.000")
                    Else
                        flexBTRM.TextMatrix(J, 2) = Format(flexBTRM.TextMatrix(J, 2), "#.000")
                    End If
                Case "AMT_09"
                    If Val(flexBTRM.TextMatrix(J, 1)) <> 0 Then
                        flexBTRM.TextMatrix(J, 2) = Format((Val(flexBTRM.TextMatrix(J, 1)) * Val(flexBTRM.TextMatrix(8, 2))) / 100, "##########.000")
                    Else
                        flexBTRM.TextMatrix(J, 2) = Format(flexBTRM.TextMatrix(J, 2), "#.000")
                    End If
            End Select
                
            If M_RDOF(J) = "Y" Then
                flexBTRM.TextMatrix(J, 2) = Format(FormatNumber(Val(flexBTRM.TextMatrix(J, 2)), 0), "############.00")
            Else
                flexBTRM.TextMatrix(J, 2) = Format(flexBTRM.TextMatrix(J, 2), "############.00")
            End If
        Else
            c_FMLA(J) = Replace(c_FMLA(J), "M_STOT", Val(TXTITOT.Text))
            c_FMLA(J) = Replace(c_FMLA(J), "M_TQTY", Val(TXTTQTY.Text))
            c_FMLA(J) = Replace(c_FMLA(J), "M_TPCS", Val(TXTTPCS.Text))
            For K = 0 To J
                c_FMLA(J) = Replace(c_FMLA(J), "AMT_0" & K + 1, Format(flexBTRM.TextMatrix(K, 2), "##########.00"))
            Next K
            c_FMLA(J) = c_FMLA(J)
            a() = Split(c_FMLA(J), " ")
            
            Dim Y As Double
            
            Y = 0
            For K = 0 To UBound(a)
             Y = Y + Val(a(K))
            Next
                
            If Val(flexBTRM.TextMatrix(J, 1)) <> 0 Then
            
              flexBTRM.TextMatrix(J, 2) = Abs(Y)
              
            End If
            
            If M_RDOF(J) = "N" Then
                If Val(flexBTRM.TextMatrix(J, 1)) <> 0 Then
                  flexBTRM.TextMatrix(J, 2) = Format((Val(flexBTRM.TextMatrix(J, 2)) * Val(flexBTRM.TextMatrix(J, 1))) / 100, "##########.00")
                  
                End If
            Else
                If Val(flexBTRM.TextMatrix(J, 1)) <> 0 Then
                  flexBTRM.TextMatrix(J, 2) = Format(FormatNumber(Val(flexBTRM.TextMatrix(J, 2)) * Val(flexBTRM.TextMatrix(J, 1)) / 100, 0), "##########.00")
                End If
            End If
            
        End If
MsubTot:
        If M_OPER(J) = "+" Then
            subTot = subTot + Val(flexBTRM.TextMatrix(J, 2))
        Else
            subTot = subTot - Val(flexBTRM.TextMatrix(J, 2))
        End If
        TXTBNET.Text = Val(TXTITOT.Text) + subTot
    Next J
End Sub

Private Sub FIL_Billingterm()
Dim CNTR As Byte
  flexBTRM.Clear
  flexBTRM.Rows = 0
  M_OPER(0) = ""
  M_OPER(1) = ""
  M_OPER(2) = ""
  M_OPER(3) = ""
  M_OPER(4) = ""
  M_OPER(5) = ""
  M_OPER(6) = ""
  M_OPER(7) = ""
  M_OPER(8) = ""
  M_OPER(9) = ""
  M_PERC(0) = 0
  M_PERC(1) = 0
  M_PERC(2) = 0
  M_PERC(3) = 0
  M_PERC(4) = 0
  M_PERC(5) = 0
  M_PERC(6) = 0
  M_PERC(7) = 0
  M_PERC(8) = 0
  M_PERC(9) = 0
  M_POSTCOD(0) = ""
  M_POSTCOD(1) = ""
  M_POSTCOD(2) = ""
  M_POSTCOD(3) = ""
  M_POSTCOD(4) = ""
  M_POSTCOD(5) = ""
  M_POSTCOD(6) = ""
  M_POSTCOD(7) = ""
  M_POSTCOD(8) = ""
  M_POSTCOD(9) = ""
  M_NICK(0) = ""
  M_NICK(1) = ""
  M_NICK(2) = ""
  M_NICK(3) = ""
  M_NICK(4) = ""
  M_NICK(5) = ""
  M_NICK(6) = ""
  M_NICK(7) = ""
  M_NICK(8) = ""
  M_NICK(9) = ""
  M_POSTYESNO(0) = ""
  M_POSTYESNO(1) = ""
  M_POSTYESNO(2) = ""
  M_POSTYESNO(3) = ""
  M_POSTYESNO(4) = ""
  M_POSTYESNO(5) = ""
  M_POSTYESNO(6) = ""
  M_POSTYESNO(7) = ""
  M_POSTYESNO(8) = ""
  M_POSTYESNO(9) = ""
  M_FMLA(0) = ""
  M_FMLA(1) = ""
  M_FMLA(2) = ""
  M_FMLA(3) = ""
  M_FMLA(4) = ""
  M_FMLA(5) = ""
  M_FMLA(6) = ""
  M_FMLA(7) = ""
  M_FMLA(8) = ""
  M_FMLA(9) = ""
  M_RDOF(0) = ""
  M_RDOF(1) = ""
  M_RDOF(2) = ""
  M_RDOF(3) = ""
  M_RDOF(4) = ""
  M_RDOF(5) = ""
  M_RDOF(6) = ""
  M_RDOF(7) = ""
  M_RDOF(8) = ""
  M_RDOF(9) = ""
  Set RS = New ADODB.Recordset
  If RS.State = 1 Then RS.Close
  RS.Open "select * from config where comp='" & compPth & "' and vtyp='SAL' AND DBCD='" & TXCD & "'  AND UNIT='" & UNCD & "' order by srch", CN, adOpenKeyset, adLockPessimistic
  CNTR = 0
  Do While Not RS.EOF
   flexBTRM.Rows = flexBTRM.Rows + 1
   flexBTRM.TextMatrix(CNTR, 0) = RS!NICK & ""
   flexBTRM.TextMatrix(CNTR, 1) = Format(RS!PERC, "#######.00")
   M_OPER(CNTR) = Trim(RS!OPER)
   M_PERC(CNTR) = RS!PERC
   M_POSTCOD(CNTR) = Trim(RS!CODE)
   M_NICK(CNTR) = Trim(RS!NICK)
   M_POSTYESNO(CNTR) = Trim(RS!post)
   M_FMLA(CNTR) = Trim(RS!FMLA)
   M_RDOF(CNTR) = Trim(RS!rdof)
   RS.MoveNext
   CNTR = CNTR + 1
  Loop
  Dim TMP_FMLA(0 To 15) As String
  CNTR = 0
  For CNTR = 0 To 9
    
    M_FMLA(CNTR) = Replace(M_FMLA(CNTR), "GROSS TOTAL", "M_STOT ")
    M_FMLA(CNTR) = Replace(M_FMLA(CNTR), "TOTAL QUANTITY", "M_TQTY ")
    M_FMLA(CNTR) = Replace(M_FMLA(CNTR), "TOTAL PCS", "M_TPCS ")
    If M_NICK(0) <> "" Then
        If M_OPER(0) = "+" Then
          M_FMLA(CNTR) = Replace(M_FMLA(CNTR), M_NICK(0), "AMT_01 ")
         Else
          M_FMLA(CNTR) = Replace(M_FMLA(CNTR), M_NICK(0), " -AMT_01")
        End If
    End If
    If M_NICK(1) <> "" Then
        If M_OPER(1) = "+" Then
          M_FMLA(CNTR) = Replace(M_FMLA(CNTR), "+" + M_NICK(1), " +AMT_02")
         Else
          M_FMLA(CNTR) = Replace(M_FMLA(CNTR), "-" + M_NICK(1), " -AMT_02")
        End If
    End If
    If M_NICK(2) <> "" Then
        If M_OPER(2) = "+" Then
          M_FMLA(CNTR) = Replace(M_FMLA(CNTR), "+" + M_NICK(2), " +AMT_03")
         Else
          M_FMLA(CNTR) = Replace(M_FMLA(CNTR), "-" + M_NICK(2), " -AMT_03")
        End If
    End If
    
    If M_NICK(3) <> "" Then
        If M_OPER(3) = "+" Then
          M_FMLA(CNTR) = Replace(M_FMLA(CNTR), "+" + M_NICK(3), " +AMT_04")
         Else
          M_FMLA(CNTR) = Replace(M_FMLA(CNTR), "-" + M_NICK(3), " -AMT_04")
        End If
    End If
    
    If M_NICK(4) <> "" Then
        If M_OPER(4) = "+" Then
          M_FMLA(CNTR) = Replace(M_FMLA(CNTR), "+" + M_NICK(4), " +AMT_05")
         Else
          M_FMLA(CNTR) = Replace(M_FMLA(CNTR), "-" + M_NICK(4), " -AMT_05")
        End If
    End If
    
    If M_NICK(5) <> "" Then
        If M_OPER(5) = "+" Then
          M_FMLA(CNTR) = Replace(M_FMLA(CNTR), "+" + M_NICK(5), " +AMT_06")
        Else
          M_FMLA(CNTR) = Replace(M_FMLA(CNTR), "-" + M_NICK(5), " -AMT_06")
        End If
    End If
    
    If M_NICK(6) <> "" Then
        If M_OPER(6) = "+" Then
          M_FMLA(CNTR) = Replace(M_FMLA(CNTR), "+" + M_NICK(6), " +AMT_07")
        Else
          M_FMLA(CNTR) = Replace(M_FMLA(CNTR), "-" + M_NICK(6), " -AMT_07")
        End If
    End If
    
    If M_NICK(7) <> "" Then
        If M_OPER(7) = "+" Then
          M_FMLA(CNTR) = Replace(M_FMLA(CNTR), "+" + M_NICK(7), " +AMT_08")
        Else
          M_FMLA(CNTR) = Replace(M_FMLA(CNTR), "-" + M_NICK(7), " -AMT_08")
        End If
    End If
    
    If M_NICK(8) <> "" Then
        If M_OPER(8) = "+" Then
          M_FMLA(CNTR) = Replace(M_FMLA(CNTR), "+" + M_NICK(8), " +AMT_09")
        Else
          M_FMLA(CNTR) = Replace(M_FMLA(CNTR), "-" + M_NICK(8), " -AMT_09")
        End If
    End If
  Next
  If flexBTRM.Rows > 0 Then
    'O.k
   Else
    flexBTRM.Enabled = False
  End If
End Sub

Private Sub FillChargesPercent()
Dim FINDRS As ADODB.Recordset
Set FINDRS = New ADODB.Recordset
If FINDRS.State = 1 Then FINDRS.Close
FINDRS.Open "SELECT * FROM TAXMST WHERE NAME ='" & M_TXNM & "'", CN, adOpenDynamic, adLockOptimistic
If Not FINDRS.EOF Then
  i = 0
  For i = 1 To flexBTRM.Rows
    J = 0
    For J = 0 To FINDRS.Fields.COUNT - 1
      If Trim(FINDRS.Fields(J).NAME) = Trim(flexBTRM.TextMatrix(i - 1, 0)) Then
         flexBTRM.TextMatrix(i - 1, 1) = FINDRS.Fields(J).Value
      End If
    Next
  Next
End If
FINDRS.Close
calBTRM 0
End Sub

