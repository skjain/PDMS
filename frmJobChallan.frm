VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{2BF72F7D-D367-4712-9146-5533EF3E691A}#1.2#0"; "FraPlus1.ocx"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WelchButton.ocx"
Begin VB.Form frmJobChallan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lumpsum Challan (Without DO)"
   ClientHeight    =   4995
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11385
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   11385
   Begin VB.PictureBox picToolTip 
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      FillColor       =   &H00404040&
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   720
      ScaleHeight     =   255
      ScaleWidth      =   4815
      TabIndex        =   40
      Top             =   5160
      Visible         =   0   'False
      Width           =   4815
      Begin VB.Timer tmrTool 
         Enabled         =   0   'False
         Interval        =   5000
         Left            =   1920
         Top             =   0
      End
      Begin VB.Label lblToolTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   240
         Left            =   120
         TabIndex        =   41
         Top             =   0
         Width           =   120
      End
   End
   Begin VB.Timer TimerBillNo 
      Interval        =   100
      Left            =   0
      Top             =   5040
   End
   Begin FramePlusCtl.FramePlus FramePlus1 
      Height          =   4995
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   8811
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
      Begin VB.TextBox TXTSTKQTY 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   9240
         Locked          =   -1  'True
         TabIndex        =   17
         Tag             =   "0"
         Top             =   2880
         Width           =   1215
      End
      Begin VB.ComboBox cmbPackingType 
         BackColor       =   &H0080C0FF&
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
         Height          =   315
         ItemData        =   "frmJobChallan.frx":0000
         Left            =   2040
         List            =   "frmJobChallan.frx":0002
         TabIndex        =   9
         Tag             =   "0"
         Text            =   "Select Type of Dispatch"
         Top             =   960
         Width           =   2655
      End
      Begin VB.TextBox TXTPCS 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   600
         TabIndex        =   18
         Top             =   3720
         Width           =   1215
      End
      Begin VB.TextBox TXTAMNT 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5400
         TabIndex        =   21
         Top             =   3720
         Width           =   1455
      End
      Begin VB.TextBox TXTITM 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   2880
         Width           =   3015
      End
      Begin VB.TextBox TXTGRAD 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5640
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   2880
         Width           =   975
      End
      Begin VB.TextBox txtDCOD 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   1920
         Width           =   3135
      End
      Begin VB.TextBox txtQTY 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2160
         TabIndex        =   19
         Top             =   3720
         Width           =   1215
      End
      Begin VB.TextBox TXTRATE 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3840
         MaxLength       =   200
         TabIndex        =   20
         Top             =   3720
         Width           =   1215
      End
      Begin VB.TextBox TXTSUBGRD 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   7080
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   2880
         Width           =   1215
      End
      Begin VB.TextBox txtLTNO 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   2880
         Width           =   1575
      End
      Begin VB.TextBox BRMK 
         Height          =   285
         Left            =   7080
         MaxLength       =   50
         TabIndex        =   22
         Top             =   3720
         Width           =   3975
      End
      Begin VB.TextBox txtConsinee 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1920
         Width           =   3735
      End
      Begin VB.TextBox TXTADDRESS 
         Height          =   525
         Left            =   7560
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   12
         Top             =   1800
         Width           =   3615
      End
      Begin VB.TextBox TXTDVNM 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Tag             =   "0"
         Top             =   480
         Width           =   3255
      End
      Begin MSComCtl2.DTPicker TXTVBDT 
         Height          =   315
         Left            =   9120
         TabIndex        =   8
         Top             =   960
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   16449537
         CurrentDate     =   39383
      End
      Begin WelchButton.lvButtons_H cmdAdd 
         Height          =   495
         Left            =   2040
         TabIndex        =   0
         Top             =   4320
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   873
         Caption         =   "&Add"
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
         Image           =   "frmJobChallan.frx":0004
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdEdit 
         Height          =   495
         Left            =   5640
         TabIndex        =   3
         Top             =   4320
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   873
         Caption         =   "E&dit"
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
         Image           =   "frmJobChallan.frx":039E
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdDelete 
         Height          =   495
         Left            =   6840
         TabIndex        =   4
         Top             =   4320
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   873
         Caption         =   "&Delete"
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
         Image           =   "frmJobChallan.frx":0738
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdSave 
         Height          =   495
         Left            =   3240
         TabIndex        =   1
         Top             =   4320
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   873
         Caption         =   "&Save"
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
         Image           =   "frmJobChallan.frx":0AD2
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdCancel 
         Height          =   495
         Left            =   4440
         TabIndex        =   2
         Top             =   4320
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
         Image           =   "frmJobChallan.frx":185C
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdExit 
         Height          =   495
         Left            =   8040
         TabIndex        =   5
         Top             =   4320
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
         Image           =   "frmJobChallan.frx":1CAE
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdSavePrint 
         Height          =   495
         Left            =   9840
         TabIndex        =   6
         Top             =   4320
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
         Caption         =   "Save/&Print"
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
         Image           =   "frmJobChallan.frx":2100
         cBack           =   -2147483633
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Stock Quantity"
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
         Left            =   9240
         TabIndex        =   43
         Tag             =   "S"
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Type of Dispatch "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   42
         Top             =   960
         Width           =   1815
      End
      Begin VB.Line Line13 
         BorderColor     =   &H00000080&
         X1              =   6960
         X2              =   6960
         Y1              =   2400
         Y2              =   4080
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Pcs"
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
         TabIndex        =   39
         Tag             =   "S"
         Top             =   3360
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
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
         Left            =   5520
         TabIndex        =   38
         Tag             =   "S"
         Top             =   3360
         Width           =   1095
      End
      Begin VB.Line Line12 
         BorderColor     =   &H00000080&
         X1              =   120
         X2              =   11280
         Y1              =   3600
         Y2              =   3600
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
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
         Left            =   7920
         TabIndex        =   37
         Tag             =   "S"
         Top             =   3360
         Width           =   1935
      End
      Begin VB.Line Line11 
         BorderColor     =   &H00000080&
         X1              =   1920
         X2              =   1920
         Y1              =   2400
         Y2              =   4080
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00000080&
         X1              =   8520
         X2              =   8520
         Y1              =   2400
         Y2              =   3240
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00000080&
         X1              =   5280
         X2              =   5280
         Y1              =   2400
         Y2              =   4080
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00000080&
         X1              =   3600
         X2              =   3600
         Y1              =   3240
         Y2              =   4080
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00000080&
         X1              =   120
         X2              =   11280
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Item Description"
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
         Left            =   2640
         TabIndex        =   36
         Tag             =   "S"
         Top             =   2520
         Width           =   1935
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Grade"
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
         Left            =   5880
         TabIndex        =   35
         Tag             =   "S"
         Top             =   2520
         Width           =   615
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Qnty."
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
         Left            =   2280
         TabIndex        =   34
         Tag             =   "S"
         Top             =   3360
         Width           =   1095
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
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
         Left            =   4080
         TabIndex        =   33
         Tag             =   "S"
         Top             =   3360
         Width           =   735
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "SubGrade"
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
         Left            =   7320
         TabIndex        =   32
         Tag             =   "S"
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "LotNo."
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
         Left            =   360
         TabIndex        =   31
         Tag             =   "S"
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000080&
         X1              =   4080
         X2              =   4080
         Y1              =   1440
         Y2              =   2400
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00000080&
         X1              =   120
         X2              =   7440
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
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
         Left            =   1200
         TabIndex        =   30
         Tag             =   "S"
         Top             =   1515
         Width           =   1455
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Consignee Address"
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
         Left            =   7920
         TabIndex        =   29
         Tag             =   "S"
         Top             =   1515
         Width           =   2175
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00000080&
         X1              =   120
         X2              =   11280
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         Height          =   2655
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   1440
         Width           =   11175
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00000080&
         X1              =   120
         X2              =   11280
         Y1              =   3240
         Y2              =   3240
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Consignee Name"
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
         Left            =   4680
         TabIndex        =   28
         Tag             =   "S"
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00000080&
         X1              =   7440
         X2              =   7440
         Y1              =   1440
         Y2              =   2400
      End
      Begin VB.Label lblAlert 
         BackStyle       =   0  'Transparent
         Caption         =   "           Challan No.   :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   6960
         TabIndex        =   27
         Top             =   480
         Width           =   2055
      End
      Begin VB.Shape BORDER 
         BorderColor     =   &H80000002&
         Height          =   300
         Left            =   6840
         Shape           =   4  'Rounded Rectangle
         Top             =   480
         Width           =   3975
      End
      Begin VB.Label lblBill 
         BackStyle       =   0  'Transparent
         Caption         =   "XXXXXXXXXX"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   9120
         TabIndex        =   26
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Division Name"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   480
         Width           =   1695
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         Height          =   1335
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   11055
      End
      Begin VB.Label LBLCHDT 
         BackStyle       =   0  'Transparent
         Caption         =   "            Challan Date :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   6960
         TabIndex        =   24
         Top             =   960
         Width           =   2175
      End
   End
End
Attribute VB_Name = "frmJobChallan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim LOAD As String
Dim DIVCODE As String
Dim DIVNAME As String
Dim SQL As String
Dim SAVEFLAG As Boolean
Dim M_DBCD As String
Dim SPARTY As String
Dim SCONSINEE As String
Dim SITEM As String
Dim SADD As String
Dim SGRD As String
Dim SUBGRD As String
Dim ALLOWEDITDEL As Boolean
Public CHALLAN As String

Private Sub cmbPackingType_Click()

Call FindStock

If InStr(1, UCase(cmbPackingType.Text), "JOB") <> 0 Then
   Me.Caption = "Box Dispatch (Job Challan) "
   lblAlert = "Job Challan No ."
   LBLCHDT = "Job Challan Date :"
ElseIf InStr(1, UCase(cmbPackingType.Text), "EXPORT") <> 0 Then
   Me.Caption = "Box Dispatch (Export Challan) "
   lblAlert = "Export Challan No ."
   LBLCHDT = "Export Challan Date :"
Else
   Me.Caption = "Box Dispatch (Sale Challan) "
   lblAlert = "    Challan No ."
   LBLCHDT = "    Challan Date :"
End If
   
   TXTVBDT = Now
   Call SetInternal
   lblBill.Caption = GenDPFVNO("DPF", M_DBCD, DIVCODE)
End Sub

Private Sub cmbPackingType_KeyDown(KeyCode As Integer, Shift As Integer)
  KeyCode = 0
End Sub

Private Sub cmbPackingType_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub cmdCancel_Click()
  TXTDVNM.Tag = TXTDVNM
  Call ClsData(Me)
  Call SetPackingType
  cmbPackingType.Locked = False
  txtQty.Tag = 0
  TXTDVNM = TXTDVNM.Tag
  SAVEFLAG = True
    Call btn_sts(True)
    If zoomflag = True Then
       Call cmdExit_Click
       Exit Sub
    End If
    
  lblBill.Caption = GenDPFVNO("DPF", M_DBCD, DIVCODE)
End Sub

Private Sub cmdDelete_Click()
  ALLOWEDITDEL = True
  SAVEFLAG = False
  CHALLAN = Empty
  btn_sts (True)
  frmJobChallanList.Show 1
  
  If ALLOWEDITDEL = False Then
    MsgBox "Sale Bill has been made can not edit/delete ", vbInformation
   Else
    If Not CHALLAN = Empty Then
      Dim AYS
      AYS = MsgBox("Are you sure to delete this Job Challan ", vbYesNo)
      If AYS = vbYes Then
        CN.BeginTrans
        CN.Execute "UPDATE SPTRAN SET RECSTAT='D' WHERE COMP='" & compPth & "' AND UNIT ='" & UNCD & "' AND VTYP='DPF' AND DBCD='" & M_DBCD & "' AND VBNO='" & CHALLAN & "'"
        CN.Execute "UPDATE JOBIN SET RECSTAT='D' WHERE COMP='" & compPth & "' AND UNIT ='" & UNCD & "' AND VTYP='DPF' AND DBCD='" & M_DBCD & "' AND VBNO='" & CHALLAN & "'"
        CN.Execute "UPDATE PKGMAN SET RECSTAT='D' WHERE COMP='" & compPth & "' AND UNIT ='" & UNCD & "' AND VTYP='DPF' AND DBCD='" & M_DBCD & "' AND SLIPNO='" & CHALLAN & "'"
        Call DAILYSTATUS("DPF", GetCode("ACCMST", txtCONSINEE, "NAME", "CODE"), M_DBCD, Val(txtQty), lblBill, Val(TXTAMNT), cUName, "D", Now, TXTVBDT)
        CN.CommitTrans
        MsgBox "Data Successfully Deleted."
      End If
    End If
  End If
  Call cmdCancel_Click
  If cmdAdd.Enabled Then cmdAdd.SetFocus
End Sub

Private Sub cmdEdit_Click()
    SAVEFLAG = False
    frmJobChallanList.DIVCODE = DIVCODE
    frmJobChallanList.M_DBCD = M_DBCD
    CHALLAN = Empty
    frmJobChallanList.Show 1
    If CHALLAN = Empty Or CHALLAN = "" Then
       btn_sts (True)
       cmdAdd.Enabled = True
       SAVEFLAG = True
       cmdAdd.SetFocus
       cmbPackingType.Locked = False
       txtQty.Tag = 0
    Else
       btn_sts (False)
       txtCONSINEE.Enabled = True
       txtCONSINEE.SetFocus
       cmbPackingType.Locked = True
       Call FindStock
    End If
End Sub

Private Sub cmdSave_Click()
On Error GoTo LAST
Dim Index As Long
Dim FLAG As Boolean
Dim SLIP As String
Dim COPS As Double
Dim PCS As Double

If INVALIDDATA Then Exit Sub

Call SetInternal

If Val(txtQty) > TXTSTKQTY Then
   MsgBox "Challan Quantity Exceed From Stock Quantity."
   txtQty.Enabled = True: txtQty.SetFocus: Exit Sub
End If

Dim NSQL As String
Dim MSGS As String: MSGS = "Unit"

If SAVEFLAG Then
   SLIP = GenDPFVNO("DPF", M_DBCD, DIVCODE)
   
   NSQL = "SELECT * FROM SPTRAN WHERE COMP='" & compPth & "'AND UNIT='" & UNCD & _
           "' AND VTYP='DPF' AND DBCD='" & M_DBCD & "' AND VBNO = '" & SLIP & "' "
   
   If UNT_DIVSERIES_REQ = "Y" Then
      NSQL = NSQL & " AND DVCD='" & DIVCODE & "' "
      MSGS = "Division"
   End If
   
   If RS.State Then RS.Close
   RS.Open NSQL, CN, adOpenDynamic, adLockOptimistic
   If Not RS.EOF Then
      MsgBox "Challan No. " & SLIP & " Already Exist. Check Last No. In " & MSGS & " Configuration", vbCritical
      Exit Sub
   End If
   RS.Close
Else
   SLIP = CHALLAN
End If

CN.BeginTrans

If SAVEFLAG = True Then

SLIP = GenDPFVNO("DPF", M_DBCD, DIVCODE)

SQL = "INSERT INTO SPTRAN(COMP,UNIT,DVCD,VTYP,DBCD,VBNO,CHLN,CHDT,DATE,CRAC,DRAC,PCOD,"
SQL = SQL & "DCOD,ADDRESS,LTNO,ICOD,GRAD,SUBGRD,PCES,QNTY,RATE,AMNT,QORP,[SYSR],[USER],OPER,"
SQL = SQL & "RECSTAT,COPS,EXTRA1)VALUES('" & compPth & "','" & UNCD & "','" & DIVCODE & _
"','DPF','" & M_DBCD & "','" & SLIP & "','" & SLIP & "','" & Format(TXTVBDT, "YYYY/MM/DD") & _
"','" & Format(TXTVBDT, "YYYY/MM/DD") & "','XXXXXX','" & SPARTY & "','" & SPARTY & "','" & SCONSINEE & _
"','" & SADD & "','" & txtLTNo & "','" & SITEM & "','" & SGRD & _
"','" & SUBGRD & "','" & TXTPCS & "','" & txtQty & "'," & TXTRATE & "," & TXTAMNT & _
",'Q','N','" & cUName & "','-','A','" & TXTPCS & "','" & Trim(BRMK) & "')"

CN.Execute SQL

SQL = "INSERT INTO PKGMAN (COMP,UNIT,DVCD,DBCD,VTYP,PCOD,SRNO,SRCH,DATE,SLIPNO,PKG_STCOD,"
SQL = SQL & "LOTNO,FINITMCOD,GRAD,SUBGRAD,QNTY,SYSR,[USER],OPER,RECSTAT) VALUES "
SQL = SQL & "('" & compPth & "','" & UNCD & "','" & DIVCODE & "','" & M_DBCD & "','DPF',"
SQL = SQL & "'" & SPARTY & "','1','1','" & Format(TXTVBDT, "YYYY/MM/DD") & "','" & SLIP & "','000000',"
SQL = SQL & "'" & txtLTNo & "','" & SITEM & "','" & SGRD & "','" & SUBGRD & "','" & txtQty & _
"','N','" & cUName & "','-','A')"

CN.Execute SQL

If InStr(1, UCase(cmbPackingType.Text), "JOB") <> 0 Then
SQL = "INSERT INTO JOBIN(COMP,UNIT,DVCD,DBCD,VTYP,VBNO,CHLN,CHDT,DATE,CRAC,DRAC,PCOD,"
SQL = SQL & "DCOD,ADDRESS,LTNO,ICOD,GRAD,SUBGRD,PCES,QNTY,RATE,AMNT,QORP,[SYSR],[USER],OPER,"
SQL = SQL & "RECSTAT,COPS,EXTRA1) VALUES('" & compPth & "','" & UNCD & "','" & DIVCODE & _
"','" & M_DBCD & "','DPF','" & SLIP & "','" & SLIP & "','" & Format(TXTVBDT, "YYYY/MM/DD") & _
"','" & Format(TXTVBDT, "YYYY/MM/DD") & "','XXXXXX','" & SPARTY & "','" & SPARTY & _
"','" & SCONSINEE & "','" & SADD & "','" & txtLTNo & "','" & SITEM & "','" & SGRD & _
"','" & SUBGRD & "','" & TXTPCS & "','" & txtQty & "'," & TXTRATE & "," & TXTAMNT & _
",'Q','N','" & cUName & "','-','A','" & TXTPCS & "','" & BRMK & "')"

CN.Execute SQL
End If

Dim UPSQL As String
UPSQL = "UPDATE SERIALMASTER SET [SRNO]='" & SLIP & "' WHERE COMP='" & compPth & "' AND UNIT = '" & UNCD & _
         "' AND VTYP='DPF' AND CODE='" & M_DBCD & "' AND FYCD='" & FYCD & "' "

If UNT_DIVSERIES_REQ = "Y" Then
   UPSQL = UPSQL & " AND DVCD='" & DIVCODE & "' "
End If
 
CN.Execute UPSQL

Else

SLIP = lblBill.Caption

SQL = "UPDATE SPTRAN SET DATE ='" & Format(TXTVBDT, "YYYY/MM/DD") & "',DRAC='" & SPARTY & "',PCOD='" & SPARTY & "',DCOD='" & SCONSINEE & _
"',ADDRESS='" & SADD & "',ICOD='" & SITEM & "',PCES='" & TXTPCS & "',QNTY='" & txtQty & "',RATE=" & TXTRATE & _
",AMNT=" & TXTAMNT & ",GRAD='" & SGRD & "',LTNO='" & txtLTNo & "',COPS='" & TXTPCS & _
"' WHERE COMP='" & compPth & "'AND UNIT='" & UNCD & "' AND DVCD='" & DIVCODE & _
"' AND VTYP='DPF' AND DBCD='" & M_DBCD & "' AND VBNO = '" & SLIP & "'"
   
CN.Execute SQL

SQL = "UPDATE PKGMAN SET DATE ='" & Format(TXTVBDT, "YYYY/MM/DD") & "',PCOD='" & SPARTY & _
"',LOTNO='" & txtLTNo & "',FINITMCOD='" & SITEM & "',GRAD='" & SGRD & "',SUBGRAD='" & SUBGRD & _
"',NOB='" & TXTPCS & "',QNTY='" & txtQty & "' WHERE COMP='" & compPth & "'AND UNIT='" & UNCD & "' AND DVCD='" & DIVCODE & _
"' AND VTYP='DPF' AND DBCD='" & M_DBCD & "' AND SLIPNO = '" & SLIP & "'"

CN.Execute SQL

If InStr(1, UCase(cmbPackingType.Text), "JOB") <> 0 Then
   SQL = "UPDATE JOBIN SET DATE ='" & Format(TXTVBDT, "YYYY/MM/DD") & "',DRAC='" & SPARTY & "',PCOD='" & SPARTY & "',DCOD='" & SCONSINEE & _
   "',ADDRESS='" & SADD & "',ICOD='" & SITEM & "',PCES='" & TXTPCS & "',QNTY='" & txtQty & "',RATE=" & TXTRATE & _
   ",AMNT=" & TXTAMNT & ",GRAD='" & SGRD & "',SUBGRD='" & SUBGRD & "',LTNO='" & txtLTNo & "',EXTRA1='" & BRMK & "',COPS='" & TXTPCS & _
   "' WHERE COMP='" & compPth & "'AND UNIT='" & UNCD & "' AND DVCD='" & DIVCODE & _
   "' AND VTYP='DPF' AND DBCD='" & M_DBCD & "' AND VBNO = '" & SLIP & "'"
   
   CN.Execute SQL
End If

End If
If SAVEFLAG = True Then
  Call DAILYSTATUS("DPF", SPARTY, M_DBCD, Val(txtQty), SLIP, Val(TXTAMNT), cUName, "N", Now, TXTVBDT)
  Else
  Call DAILYSTATUS("DPF", SPARTY, M_DBCD, Val(txtQty), SLIP, Val(TXTAMNT), cUName, "M", Now, TXTVBDT)
End If

CN.CommitTrans

If SAVEFLAG Then
  MsgBox "Your Challan No. is : " & SLIP
Else
   MsgBox "Challan No. : " & SLIP & " is Successfully Edited."
End If

Call cmdCancel_Click

Exit Sub
LAST:
MsgBox ERR.Description
Exit Sub
End Sub

Private Sub cmdSavePrint_Click()
  Call cmdSave_Click
End Sub

Private Sub Form_Activate()
'If LOAD = "N" Then Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  SendKeys "{TAB}"
End If
End Sub

Private Sub Form_Load()
Call ColorComponent(Me)
   
  M_DESC = Empty: Key = Empty: NEW_VISIBLE = False
  DIVCODE = Empty: DIVNAME = Empty
  
  If DIVCODE = Empty Then
    DIVNAME = SearchList1("SELECT  TOP 20 CODE,NAME FROM DIVMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND RECSTAT='A'  AND CODE<>'000001'", 0, "", "SELECT DIVISION MASTER")
    DIVCODE = Key
  End If
  TXTDVNM = DIVNAME
  
 If PackingType(Key) = "C" Then MsgBox "Division Not Allowed Lumpsum Packing.Check Configuration": LOAD = "N": GoTo JUMP
  
  TXTVBDT = Now
  TXTVBDT = GetMinDate
  TXTVBDT = GetMaxDate
    
    If zoomflag = True Then
        btn_sts (False)
        SAVEFLAG = False
    Else
        btn_sts (True)
    End If
    
 SAVEFLAG = True
 Call SetPackingType
 Call SetInternal
 lblBill.Caption = GenDPFVNO("DPF", M_DBCD, DIVCODE)
JUMP:
End Sub

Private Sub cmdadd_Click()
    zoomflag = False
    btn_sts (False)
    lblBill.Caption = GenDPFVNO("DPF", M_DBCD, DIVCODE)
    txtCONSINEE.SetFocus
    SAVEFLAG = True
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub TXTADDRESS_KeyDown(KeyCode As Integer, Shift As Integer)
   TXTADDRESS.FontSize = 8
   If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
    TXTADDRESS = Empty
   ElseIf KeyCode = vbKeyF2 Or (TXTADDRESS = Empty And KeyCode = vbKeyReturn) Then
    TXTADDRESS = SearchAddress("Select SRNO,ADDR From PADDMST WHERE NAME='" & TXTDCOD & "'", 0, Empty, "Select A/c Party Filtered by Party Group")
   End If
End Sub

Private Sub txtCONSINEE_KeyDown(KeyCode As Integer, Shift As Integer)
  Me.KeyPreview = False
    
  If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
      txtCONSINEE = Empty
  ElseIf KeyCode = vbKeyF2 Then
     M_DESC = Empty:   NEW_VISIBLE = False
     txtCONSINEE = SearchList1("Select TOP 20 Code,Name From ACCMST", 0, Empty, "Select A/c Party ")
  End If
      
  Me.KeyPreview = True
  
End Sub

Private Sub txtConsinee_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And txtCONSINEE = Empty Then
        Call txtCONSINEE_KeyDown(vbKeyF2, 0)
    End If
End Sub

Private Sub TXTDCOD_KeyDown(KeyCode As Integer, Shift As Integer)
 Me.KeyPreview = False
    
  If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
     TXTDCOD = Empty
  ElseIf KeyCode = vbKeyF2 Or TXTDCOD = Empty Then
     M_DESC = Empty:   NEW_VISIBLE = False
     TXTDCOD = SearchList1("Select DISTINCT CODE,NAME From PADDMST", 0, Empty, "Select Consinee Name ")
  End If
  
 Me.KeyPreview = True
End Sub

Private Sub TXTGRAD_Change()
Call FindStock
End Sub

Private Sub TXTGRAD_KeyDown(KeyCode As Integer, Shift As Integer)
  If (TXTGRAD = Empty And KeyCode = vbKeyReturn) Or KeyCode = vbKeyF2 Then
    NEW_VISIBLE = True
    TXTGRAD = SearchList1("SELECT DISTINCT GRAD AS GRD,GRAD FROM GRDMST", 0, TXTGRAD, "SELECT MAIN GRAD FROM LIST")
      If key_PressNew = True Then
          M_DESC = ""
          TXTGRAD = Empty
          FRM_GRDMST.Show
      End If
  End If
End Sub

Private Sub TXTITM_Change()
 Call FindStock
End Sub

Private Sub txtLTNO_Change()
  Call FindStock
End Sub

Private Sub txtltno_KeyDown(KeyCode As Integer, Shift As Integer)
Me.KeyPreview = False

If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
     txtLTNo = Empty
ElseIf KeyCode = vbKeyF2 Or (KeyCode = vbKeyReturn And txtLTNo = Empty) Then
   M_DESC = Empty:   NEW_VISIBLE = False
   SQL = "SELECT DISTINCT LTNO,LTNO FROM TXULOT WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCODE & "' AND ACTIVE='Y' "
   txtLTNo = SearchList(SQL)
End If
   TXTITM = FindItem
Me.KeyPreview = True
End Sub

Private Sub TXTPCS_KeyPress(KeyAscii As Integer)
 If CheckNumericKey(KeyAscii, TXTPCS, Me) = 0 Then KeyAscii = 0
End Sub

Private Sub txtQTY_KeyPress(KeyAscii As Integer)
 If CheckNumericKey(KeyAscii, txtQty, Me) = 0 Then KeyAscii = 0
End Sub

Private Sub txtRate_KeyPress(KeyAscii As Integer)
 If CheckNumericKey(KeyAscii, TXTRATE, Me) = 0 Then KeyAscii = 0
 TXTAMNT = Val(txtQty) * Val(TXTRATE)
 TXTAMNT = nstr(TXTAMNT, 12, 2)
End Sub

Private Sub TXTSUBGRD_Change()
 Call FindStock
End Sub

Private Sub TXTSUBGRD_KeyDown(KeyCode As Integer, Shift As Integer)
Me.KeyPreview = False

If TXTGRAD = Empty Then TXTGRAD.Enabled = True: TXTGRAD.SetFocus: Exit Sub

If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
     TXTSUBGRD = Empty
ElseIf KeyCode = vbKeyF2 Or (TXTSUBGRD = Empty And KeyCode = vbKeyReturn) Then
   M_DESC = Empty:   NEW_VISIBLE = False
   SQL = "SELECT DISTINCT RDIFF,NAME FROM SUBGRDMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCODE & "' AND GRAD='" & GetCode("GRDMST", TXTGRAD, "GRAD", "CODE") & "'"
   TXTSUBGRD = SearchList1(SQL, 0, Empty)
End If
Me.KeyPreview = True
End Sub

Private Function GetGroupPartyCode(PNAM As String) As String
Dim GRRS As ADODB.Recordset
Set GRRS = New ADODB.Recordset

If GRRS.State = 1 Then GRRS.Close
GRRS.Open "SELECT CPCD FROM ACCMST WHERE NAME ='" & PNAM & "'", CN, adOpenDynamic, adLockOptimistic
If Not GRRS.EOF Then
   GetGroupPartyCode = Trim(GRRS!CDCD & "")
Else
   GetGroupPartyCode = Empty
End If

GRRS.Close
End Function

Private Sub TXTAMNT_LostFocus()
TXTAMNT.BackColor = vbWhite
End Sub

Private Sub TXTPCS_LostFocus()
TXTPCS.BackColor = vbWhite
End Sub

Private Sub TXTQTY_LostFocus()
txtQty.BackColor = vbWhite
End Sub

Private Sub txtRate_LostFocus()
TXTRATE.BackColor = vbWhite
End Sub


Private Sub TXTSUBGRD_LostFocus()
TXTSUBGRD.BackColor = vbWhite
 picToolTip.Visible = False
End Sub

Private Sub BRMK_GotFocus()
BRMK.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTADDRESS_GotFocus()
   TXTADDRESS.BackColor = RGB(BRED, BGREEN, BBLUE)
   SendKeys "{HOME}+{END}"
  
  If TXTADDRESS = Empty Then
      ToolTip Me, "Press {F2} / {Enter} For Address Master Help", "", TXTADDRESS.Left - 400, TXTADDRESS.Top + TXTADDRESS.Height + 100
  Else
      ToolTip Me, "Press {F2} For Address Master Help", "", TXTADDRESS.Left - 400, TXTADDRESS.Top + TXTADDRESS.Height + 100
  End If
   
End Sub

Private Sub TXTADDRESS_LostFocus()
   TXTADDRESS.BackColor = vbWhite
   picToolTip.Visible = False
End Sub

Private Sub txtAmnt_GotFocus()
   TXTAMNT.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub


Private Sub txtConsinee_GotFocus()
  txtCONSINEE.BackColor = RGB(BRED, BGREEN, BBLUE)
  SendKeys "{HOME}+{END}"
  
  If txtCONSINEE = Empty Then
      ToolTip Me, "Press {F2} / {Enter} For A/C Party Master Help", "", txtCONSINEE.Left, txtCONSINEE.Top + txtCONSINEE.Height + 100
  Else
      ToolTip Me, "Press {F2} For A/C Party Master Help", "", txtCONSINEE.Left, txtCONSINEE.Top + txtCONSINEE.Height + 100
  End If
End Sub

Private Sub txtConsinee_LostFocus()
 txtCONSINEE.BackColor = vbWhite
 picToolTip.Visible = False
 
    If SAVEFLAG Then
     Dim GETRS As ADODB.Recordset
     Set GETRS = New ADODB.Recordset
  
     If GETRS.State = 1 Then GETRS.Close
     GETRS.Open "SELECT RCOD FROM ACCMST WHERE NAME='" & txtCONSINEE & "' ", CN, adOpenDynamic, adLockOptimistic
     If Not GETRS.EOF Then
        TXTDCOD = GetCode("PADDMST", GETRS!RCOD & "", "CODE", "NAME")
        TXTADDRESS = GetCode("PADDMST", GETRS!RCOD & "", "CODE", "ADDR")
     End If
  End If
 
End Sub

Private Sub TXTDCOD_GotFocus()
 TXTDCOD.BackColor = RGB(BRED, BGREEN, BBLUE)
 SendKeys "{HOME}+{END}"
  
  If TXTDCOD = Empty Then
      ToolTip Me, "Press {F2} / {Enter} For Consinee Master Help", "", TXTDCOD.Left, TXTDCOD.Top + TXTDCOD.Height + 100
  Else
      ToolTip Me, "Press {F2} For Consinee Master Help", "", TXTDCOD.Left, TXTDCOD.Top + TXTDCOD.Height + 100
  End If
End Sub

Private Sub TXTDCOD_LostFocus()
TXTDCOD.BackColor = vbWhite
 picToolTip.Visible = False
End Sub

Private Sub TXTGRAD_GotFocus()
 TXTGRAD.BackColor = RGB(BRED, BGREEN, BBLUE)
 SendKeys "{HOME}+{END}"
  
  If TXTGRAD = Empty Then
      ToolTip Me, "Press {F2} / {Enter} For Grade Master Help", "", TXTGRAD.Left - 1900, TXTGRAD.Top + TXTGRAD.Height + 100
  Else
      ToolTip Me, "Press {F2} For Grade Master Help", "", TXTGRAD.Left - 1900, TXTGRAD.Top + TXTGRAD.Height + 100
  End If
End Sub

Private Sub TXTGRAD_LostFocus()
TXTGRAD.BackColor = vbWhite
 picToolTip.Visible = False
End Sub

Private Sub TXTITM_GotFocus()
TXTITM.BackColor = RGB(BRED, BGREEN, BBLUE)
SendKeys "{HOME}+{END}"
ToolTip Me, "Item Defined in Lot", "", TXTITM.Left, TXTITM.Top + TXTITM.Height + 100
End Sub

Private Sub TXTITM_LostFocus()
TXTITM.BackColor = vbWhite
 picToolTip.Visible = False
End Sub

Private Sub txtltno_GotFocus()
  txtLTNo.BackColor = RGB(BRED, BGREEN, BBLUE)
  SendKeys "{HOME}+{END}"
  
  If txtLTNo = Empty Then
      ToolTip Me, "Press {F2} / {Enter} For Lot Master Help", "", txtLTNo.Left, txtLTNo.Top + txtLTNo.Height + 100
  Else
      ToolTip Me, "Press {F2} For Lot Master Help", "", txtLTNo.Left, txtLTNo.Top + txtLTNo.Height + 100
  End If
End Sub

Private Sub txtltno_LostFocus()
  txtLTNo.BackColor = vbWhite
   picToolTip.Visible = False
End Sub

Private Sub TXTPCS_GotFocus()
TXTPCS.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTQTY_GotFocus()
 txtQty.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtRate_GotFocus()
  TXTRATE.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub BRMK_LostFocus()
  BRMK.BackColor = vbWhite
End Sub

Private Sub TXTSUBGRD_GotFocus()
  TXTSUBGRD.BackColor = RGB(BRED, BGREEN, BBLUE)
   
  If TXTSUBGRD = Empty Then
      ToolTip Me, "Press {F2} / {Enter} For SubGrade Master Help", "", TXTSUBGRD.Left - 3800, TXTSUBGRD.Top + TXTSUBGRD.Height + 100
  Else
      ToolTip Me, "Press {F2} For SubGrade Master Help", "", TXTSUBGRD.Left - 3800, TXTSUBGRD.Top + TXTSUBGRD.Height + 100
  End If
End Sub

Public Sub btn_sts(Yes As Boolean)
    cmdSave.Enabled = Not Yes
    cmdCancel.Enabled = Not Yes
    cmdAdd.Enabled = Yes
    cmdEdit.Enabled = Yes
    cmdDelete.Enabled = Yes
    txtCONSINEE.Enabled = Not Yes
    TXTVBDT.Enabled = Not Yes
    TXTDCOD.Enabled = Not Yes
    TXTADDRESS.Enabled = Not Yes
    TXTSTKQTY.Enabled = Not Yes
    txtLTNo.Enabled = Not Yes
    TXTITM.Enabled = Not Yes
    TXTGRAD.Enabled = Not Yes
    TXTSUBGRD.Enabled = Not Yes
    TXTPCS.Enabled = Not Yes
    txtQty.Enabled = Not Yes
    TXTRATE.Enabled = Not Yes
    TXTAMNT.Enabled = Not Yes
    BRMK.Enabled = Not Yes
End Sub

Private Sub TimerBillNo1_Timer()
    Static ctr As Integer
    
    If ctr Mod 45 = 0 And ctr <= 45 Then
        lblAlert.ForeColor = RGB(LBLRED, LBLGREEN, LBLBLUE)
        BORDER.BorderColor = RGB(LBLRED, LBLGREEN, LBLBLUE)
        lblBill.ForeColor = RGB(LBLRED, LBLGREEN, LBLBLUE)
    ElseIf ctr Mod 75 = 0 And ctr <= 75 Then
        lblAlert.ForeColor = vbRed
        BORDER.BorderColor = vbRed
        lblBill.ForeColor = vbRed
    ElseIf ctr Mod 105 = 0 And ctr <= 105 Then
        lblAlert.ForeColor = vbBlue
        BORDER.BorderColor = vbBlue
        lblBill.ForeColor = vbBlue
        ctr = 0
    End If
    
    ctr = ctr + 15
End Sub

Private Function FindItem() As String
Dim FICD As String
Dim FINDRS As ADODB.Recordset
Set FINDRS = New ADODB.Recordset


If FINDRS.State = 1 Then FINDRS.Close
FINDRS.Open "SELECT FICD FROM TXULOT WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
"' AND DVCD='" & DIVCODE & "' AND LTNO='" & txtLTNo & "'", CN, adOpenDynamic, adLockOptimistic
If Not FINDRS.EOF Then
   FICD = Trim(FINDRS!FICD & "")
Else
   FICD = Empty
   Exit Function
End If
FINDRS.Close

If FINDRS.State = 1 Then FINDRS.Close
FINDRS.Open "SELECT NAME FROM FINITMMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
"' AND DVCD='" & DIVCODE & "' AND CODE='" & FICD & "'", CN, adOpenDynamic, adLockOptimistic
If Not FINDRS.EOF Then
   FindItem = Trim(FINDRS!NAME & "")
Else
   FindItem = Empty
   Exit Function
End If
FINDRS.Close

End Function

Private Sub SetInternal()

'PART:1
Dim GRRS As ADODB.Recordset
Set GRRS = New ADODB.Recordset

If GRRS.State = 1 Then GRRS.Close
GRRS.Open "SELECT CODE FROM SERIALMASTER WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
"' AND VTYP='DPF' AND NAME = '" & cmbPackingType.Text & "'", CN, adOpenDynamic, adLockOptimistic

If Not GRRS.EOF Then
   M_DBCD = Trim(GRRS!CODE & "")
Else
   M_DBCD = Empty
End If

'PART:2
If txtLTNo = Empty Or TXTITM = Empty Or TXTGRAD = Empty Or TXTSUBGRD = Empty Then Exit Sub
SGRD = GetCode("GRDMST", TXTGRAD, "GRAD", "CODE")

If GRRS.State = 1 Then GRRS.Close
GRRS.Open "SELECT * FROM SUBGRDMST WHERE COMP='" & compPth & "' AND UNIT ='" & UNCD & _
"' AND DVCD = '" & DIVCODE & "' AND GRAD = '" & SGRD & "' AND NAME = '" & TXTSUBGRD & "'", CN, adOpenDynamic, adLockOptimistic
If Not GRRS.EOF Then
   SUBGRD = Trim(GRRS!SUBGRD & "")
End If
GRRS.Close

If GRRS.State = 1 Then GRRS.Close
GRRS.Open "SELECT CODE FROM FINITMMST WHERE COMP='" & compPth & "' AND UNIT ='" & UNCD & _
"' AND DVCD = '" & DIVCODE & "' AND NAME = '" & TXTITM & "'", CN, adOpenDynamic, adLockOptimistic
If Not GRRS.EOF Then
   SITEM = Trim(GRRS!CODE & "")
End If
GRRS.Close

'PART:3
If txtCONSINEE = Empty Or TXTDCOD = Empty Or TXTADDRESS = Empty Then Exit Sub

SPARTY = GetCode("ACCMST", txtCONSINEE, "NAME", "CODE")
SCONSINEE = GetCode("PADDMST", TXTDCOD, "NAME", "CODE")

If GRRS.State = 1 Then GRRS.Close
GRRS.Open "SELECT SRNO FROM PADDMST WHERE CODE='" & SCONSINEE & "' AND ADDR='" & TXTADDRESS & "'", CN, adOpenDynamic, adLockOptimistic
If Not GRRS.EOF Then
  SADD = GRRS!SRNO
Else
  SADD = Empty
End If
GRRS.Close

End Sub

Private Sub UPDATEDELSTATUS()
  Dim DLYSTA As New ADODB.Recordset
  If DLYSTA.State = adStateOpen Then DLYSTA.Close
  DLYSTA.Open "SELECT * FROM DAILYSTAT WHERE 1=2", CN, adOpenKeyset, adLockPessimistic
  DLYSTA.AddNew
  DLYSTA!COMP = compPth & ""
  DLYSTA!VTYP = "DPF"
  DLYSTA!PCOD = txtCONSINEE
  DLYSTA!dbcd = M_DBCD
  DLYSTA!QNTY = Val(txtQty)
  DLYSTA!VBNO = lblBill & ""
  DLYSTA!AMNT = Val(TXTAMNT)
  DLYSTA!CUSR = cUName
  DLYSTA!ACTN = "D"
  DLYSTA!DTTM = Format(Now, "YYYY/MM/DD HH:MM:SS AMPM")
  DLYSTA.Update
  DLYSTA.Close
End Sub

Private Sub FindStock()

If txtLTNo = Empty Or TXTITM = Empty Or TXTGRAD = Empty Or TXTSUBGRD = Empty Then Exit Sub

Call SetInternal

Dim PACKEDQTY As Double: PACKEDQTY = 0
Dim DISPATCHEDQTY As Double: DISPATCHEDQTY = 0

SQL = "SELECT SUM(ISNULL(QNTY,0)) AS PACKED FROM PKGMAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
"' AND DVCD='" & DIVCODE & "' AND VTYP='PPF' AND LOTNO='" & txtLTNo & "' AND FINITMCOD='" & SITEM & _
"' AND GRAD='" & SGRD & "' AND SUBGRAD='" & SUBGRD & "' AND OPER='+' AND RECSTAT='A' "

If InStr(1, UCase(cmbPackingType.Text), "JOB") <> 0 Then
   SQL = SQL & "AND DBCD='000005' AND PCOD='" & SPARTY & "' "
ElseIf InStr(1, UCase(cmbPackingType.Text), "CAPTIVE") <> 0 Then
   SQL = SQL & "AND DBCD='000001' "
ElseIf InStr(1, UCase(cmbPackingType.Text), "EXPORT") <> 0 Then
   SQL = SQL & "AND DBCD='000002' "
Else
   SQL = SQL & "AND DBCD IN ('000003','000004') "
End If

Dim CHKRS As ADODB.Recordset
Set CHKRS = New ADODB.Recordset
If CHKRS.State = 1 Then CHKRS.Close
CHKRS.Open SQL, CN, adOpenDynamic, adLockOptimistic
If Not CHKRS.EOF Then
 PACKEDQTY = Val(Trim(CHKRS!PACKED & ""))
End If
CHKRS.Close

If CHKRS.State = 1 Then CHKRS.Close
CHKRS.Open "SELECT SUM(ISNULL(QNTY,0)) AS DISPACHED FROM PKGMAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
"' AND DVCD='" & DIVCODE & "' AND VTYP='DPF' AND DBCD='" & M_DBCD & "' AND LOTNO='" & txtLTNo & _
"' AND FINITMCOD='" & SITEM & "' AND GRAD='" & SGRD & _
"' AND SUBGRAD='" & SUBGRD & "' AND OPER='-' AND RECSTAT='A' AND PCOD='" & SPARTY & "'", CN, adOpenDynamic, adLockOptimistic

If Not CHKRS.EOF Then
 DISPATCHEDQTY = Val(Trim(CHKRS!DISPACHED & ""))
End If
CHKRS.Close

TXTSTKQTY = PACKEDQTY - DISPATCHEDQTY

If Trim(TXTITM) = Trim(TXTITM.Tag) And Trim(TXTGRAD) = Trim(TXTGRAD.Tag) And Trim(TXTSUBGRD) = Trim(TXTSUBGRD.Tag) And Trim(txtLTNo.Tag) = Trim(txtLTNo) And Not SAVEFLAG Then
   TXTSTKQTY = TXTSTKQTY + Val(txtQty.Tag)
End If

TXTSTKQTY = nstr(TXTSTKQTY, 12, 3)
TXTSTKQTY = Trim(TXTSTKQTY)
End Sub

Private Function INVALIDDATA() As Boolean

If txtCONSINEE = Empty Then
  If txtCONSINEE.Enabled Then txtCONSINEE.SetFocus
  INVALIDDATA = True
  Exit Function
End If

If TXTDCOD = Empty Then
  If TXTDCOD.Enabled Then TXTDCOD.SetFocus
  INVALIDDATA = True
  Exit Function
End If

If TXTADDRESS = Empty Then
  If TXTADDRESS.Enabled Then TXTADDRESS.SetFocus
  INVALIDDATA = True
  Exit Function
End If

If TXTITM = Empty Then
  If TXTITM.Enabled Then TXTITM.SetFocus
  INVALIDDATA = True
  Exit Function
End If

If txtLTNo = Empty Then
  If txtLTNo.Enabled Then txtLTNo.SetFocus
  INVALIDDATA = True
  Exit Function
End If

If TXTGRAD = Empty Then
  If TXTGRAD.Enabled Then TXTGRAD.SetFocus
  INVALIDDATA = True
  Exit Function
End If

If TXTSUBGRD = Empty Then
  If TXTSUBGRD.Enabled Then TXTSUBGRD.SetFocus
  INVALIDDATA = True
  Exit Function
End If

If TXTPCS = Empty Then
  If TXTPCS.Enabled Then TXTPCS.SetFocus
  INVALIDDATA = True
  Exit Function
End If

If txtQty = Empty Or Val(txtQty) = 0 Then
  If txtQty.Enabled Then txtQty.SetFocus
  INVALIDDATA = True
  Exit Function
End If

If TXTRATE = Empty Then
  If TXTRATE.Enabled Then TXTRATE.SetFocus
  INVALIDDATA = True
  Exit Function
End If

If TXTAMNT = Empty Then
  If TXTAMNT.Enabled Then TXTAMNT.SetFocus
  INVALIDDATA = True
  Exit Function
End If
End Function


Private Sub SetPackingType()

cmbPackingType.Clear
Dim PKTYPRS As ADODB.Recordset
Set PKTYPRS = New ADODB.Recordset
Dim SQL As String
SQL = "SELECT DISTINCT CODE,NAME FROM SERIALMASTER WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
      "' AND VTYP='DPF' AND NAME NOT LIKE '%CAPTIVE%' AND NAME NOT LIKE '%WASTAGE%'  AND NAME<>''"

If PKTYPRS.State = 1 Then PKTYPRS.Close
PKTYPRS.Open SQL, CN, adOpenDynamic, adLockOptimistic

If Not PKTYPRS.EOF Then M_DBCD = Trim(PKTYPRS!CODE)

Do While Not PKTYPRS.EOF
 cmbPackingType.AddItem Trim(PKTYPRS!NAME)
 PKTYPRS.MoveNext
Loop
If cmbPackingType.ListCount > 0 Then cmbPackingType.ListIndex = 0
End Sub
