VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{2BF72F7D-D367-4712-9146-5533EF3E691A}#1.2#0"; "FraPlus1.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WELCHBUTTON.OCX"
Begin VB.Form frmBoxPacking 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Carton / Pallet Packing "
   ClientHeight    =   7215
   ClientLeft      =   375
   ClientTop       =   1110
   ClientWidth     =   11355
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Packing"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7246.647
   ScaleMode       =   0  'User
   ScaleWidth      =   11385.24
   Begin VB.PictureBox picToolTip 
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      FillColor       =   &H00404040&
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   1320
      ScaleHeight     =   255
      ScaleWidth      =   4815
      TabIndex        =   24
      Top             =   7440
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
         ForeColor       =   &H00000080&
         Height          =   240
         Left            =   120
         TabIndex        =   25
         Top             =   0
         Width           =   120
      End
   End
   Begin VB.Timer TimerBillNo 
      Interval        =   100
      Left            =   0
      Top             =   7440
   End
   Begin FramePlusCtl.FramePlus FrmAutoConsumption 
      Height          =   7275
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   12832
      BackgroundPictureAlignment=   5
      BorderStyle     =   10
      BackColorGradient=   12632319
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
      Begin VB.TextBox TXTSHADE 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   2040
         Visible         =   0   'False
         Width           =   2175
      End
      Begin Crystal.CrystalReport CRPT 
         Left            =   5280
         Top             =   4560
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.CheckBox GRPLAT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0FF&
         Caption         =   "Pallet Complete ?"
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
         Left            =   360
         TabIndex        =   16
         Top             =   4920
         Width           =   2130
      End
      Begin VB.TextBox TXTMCCD 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9840
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1140
         Width           =   1455
      End
      Begin VB.TextBox TXTPCOD 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   1200
         Width           =   3015
      End
      Begin VB.TextBox NETWGT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   285
         Left            =   9720
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   54
         Tag             =   "0"
         Top             =   6840
         Width           =   1335
      End
      Begin VB.TextBox NETCOPS 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   285
         Left            =   8640
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   53
         Tag             =   "0"
         Top             =   6840
         Width           =   975
      End
      Begin VB.TextBox NETBOXES 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   285
         Left            =   7320
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   52
         Tag             =   "0"
         Top             =   6840
         Width           =   975
      End
      Begin VB.TextBox TXTRMRK 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         MaxLength       =   148
         TabIndex        =   17
         Top             =   5400
         Width           =   2415
      End
      Begin VB.TextBox TXTNTWT 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   15
         Tag             =   "0"
         Top             =   4320
         Width           =   1335
      End
      Begin VB.TextBox TXTGRWT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         MaxLength       =   9
         TabIndex        =   13
         Top             =   3600
         Width           =   1335
      End
      Begin VB.TextBox TXTTRWT 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   14
         Tag             =   "0"
         Top             =   3960
         Width           =   1335
      End
      Begin VB.TextBox TXTCTWT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         MaxLength       =   9
         TabIndex        =   11
         Top             =   2760
         Width           =   1335
      End
      Begin VB.TextBox TXTCPWT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         MaxLength       =   9
         TabIndex        =   12
         Text            =   "0."
         Top             =   3060
         Width           =   1335
      End
      Begin VB.TextBox TXTTWIST 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         MaxLength       =   1
         TabIndex        =   8
         Text            =   "S"
         Top             =   2040
         Width           =   495
      End
      Begin VB.TextBox TXTCOP 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         MaxLength       =   5
         TabIndex        =   10
         Top             =   2460
         Width           =   975
      End
      Begin VB.TextBox TXTGRAD 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9840
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   800
         Width           =   1455
      End
      Begin VB.TextBox TXTLOC 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6960
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   1140
         Width           =   1935
      End
      Begin VB.ComboBox cmbPackaging 
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
         ItemData        =   "frmBoxPacking.frx":0000
         Left            =   2040
         List            =   "frmBoxPacking.frx":0002
         Sorted          =   -1  'True
         TabIndex        =   1
         Tag             =   "0"
         Text            =   "Select Type of Packaging"
         Top             =   840
         Width           =   3015
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
         ItemData        =   "frmBoxPacking.frx":0004
         Left            =   2040
         List            =   "frmBoxPacking.frx":0006
         TabIndex        =   0
         Tag             =   "0"
         Text            =   "cmbPackingType"
         Top             =   480
         Width           =   3015
      End
      Begin VB.TextBox TXTDENI 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6960
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   780
         Width           =   1935
      End
      Begin VB.TextBox TXTLTNO 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6960
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   480
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker TXTVBDT 
         Height          =   285
         Left            =   9840
         TabIndex        =   35
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Format          =   21757953
         CurrentDate     =   39347
      End
      Begin WelchButton.lvButtons_H cmdSave 
         Height          =   375
         Left            =   240
         TabIndex        =   20
         Top             =   6720
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
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
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdExit 
         Height          =   375
         Left            =   2640
         TabIndex        =   21
         Top             =   6750
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
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
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdSavePrint 
         Height          =   375
         Left            =   1440
         TabIndex        =   19
         Top             =   6750
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         Caption         =   "S&ave/Print"
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
         cBack           =   -2147483633
      End
      Begin MSFlexGridLib.MSFlexGrid Flex 
         Height          =   4890
         Left            =   3795
         TabIndex        =   49
         Top             =   1680
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   8625
         _Version        =   393216
         Cols            =   12
         FixedCols       =   0
         BackColor       =   -2147483628
         BackColorBkg    =   15786495
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid FLEXPLY 
         Height          =   885
         Left            =   240
         TabIndex        =   18
         Top             =   5760
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   1561
         _Version        =   393216
         Cols            =   50
         BackColor       =   -2147483628
         BackColorBkg    =   15786495
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Shape Shape1 
         Height          =   495
         Left            =   240
         Shape           =   4  'Rounded Rectangle
         Top             =   4800
         Width           =   2415
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Back Date Packing not Allowed."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   3960
         TabIndex        =   59
         Top             =   6840
         Width           =   3135
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "M/c No."
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
         Left            =   9000
         TabIndex        =   58
         Top             =   1155
         Width           =   855
      End
      Begin VB.Label LBLPCOD 
         BackStyle       =   0  'Transparent
         Caption         =   "Job Party Name :"
         Enabled         =   0   'False
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
         TabIndex        =   57
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Net Weight"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   9960
         TabIndex        =   56
         Top             =   6600
         Width           =   1065
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total BOXES"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   7200
         TabIndex        =   55
         Top             =   6600
         Width           =   1170
      End
      Begin VB.Line Line2 
         BorderStyle     =   5  'Dash-Dot-Dot
         X1              =   3720
         X2              =   3720
         Y1              =   1560
         Y2              =   7200
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Double Click On Box No For Edit."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   3960
         TabIndex        =   51
         Top             =   6600
         Width           =   3135
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total COPS"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   8640
         TabIndex        =   50
         Top             =   6600
         Width           =   1050
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks."
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
         TabIndex        =   48
         Top             =   5400
         Width           =   1095
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H80000002&
         BorderStyle     =   5  'Dash-Dot-Dot
         DrawMode        =   1  'Blackness
         FillColor       =   &H00000080&
         Height          =   1275
         Left            =   120
         Top             =   3480
         Width           =   3615
      End
      Begin VB.Label LBLSZO 
         BackStyle       =   0  'Transparent
         Caption         =   "{S/Z/0}"
         Enabled         =   0   'False
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
         Left            =   2040
         TabIndex        =   47
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label BOXNO 
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
         Left            =   1440
         TabIndex        =   46
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Tare Wgt.                            {-}"
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
         TabIndex        =   45
         Top             =   3960
         Width           =   3135
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Net Wgt.                              {=}"
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
         TabIndex        =   44
         Top             =   4320
         Width           =   3015
      End
      Begin VB.Label LBLBOXWGT 
         BackStyle       =   0  'Transparent
         Caption         =   "Box Wgt. "
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
         TabIndex        =   43
         Top             =   2760
         Width           =   1575
      End
      Begin VB.Label LBLCOPSWGT 
         BackStyle       =   0  'Transparent
         Caption         =   "Cops Wgt. "
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
         Top             =   3075
         Width           =   1575
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Gross Wgt.                          {+}"
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
         TabIndex        =   41
         Top             =   3600
         Width           =   3135
      End
      Begin VB.Label LBLBOXNO 
         BackStyle       =   0  'Transparent
         Caption         =   "Box No. :"
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
         TabIndex        =   40
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label LBLTWST 
         BackStyle       =   0  'Transparent
         Caption         =   "Twist "
         Enabled         =   0   'False
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
         TabIndex        =   39
         Top             =   2040
         Width           =   1815
      End
      Begin VB.Label LBLNOCOPS 
         BackStyle       =   0  'Transparent
         Caption         =   "No.of Cops "
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
         TabIndex        =   38
         Top             =   2430
         Width           =   1455
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H80000002&
         BorderStyle     =   5  'Dash-Dot-Dot
         DrawMode        =   1  'Blackness
         FillColor       =   &H00000080&
         Height          =   5625
         Left            =   120
         Top             =   1560
         Width           =   11175
      End
      Begin VB.Label LBLLOT 
         BackStyle       =   0  'Transparent
         Caption         =   "LotNo:"
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
         Left            =   5880
         TabIndex        =   37
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Storage Location :"
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
         Left            =   5160
         TabIndex        =   36
         Top             =   1125
         Width           =   1815
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Packaging Type :"
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
         TabIndex        =   34
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Type of Packing :"
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
         TabIndex        =   33
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label LBLCFG 
         BackStyle       =   0  'Transparent
         Caption         =   "Grade:"
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
         TabIndex        =   32
         Top             =   795
         Width           =   735
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Da&te :"
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
         Left            =   9240
         TabIndex        =   31
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Item Name  :"
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
         Left            =   5640
         TabIndex        =   30
         Top             =   795
         Width           =   1335
      End
      Begin VB.Label LBLDESC1 
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
         Left            =   1920
         TabIndex        =   29
         Top             =   120
         Width           =   3375
      End
      Begin VB.Shape BORDER1 
         BorderColor     =   &H80000002&
         Height          =   300
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   5295
      End
      Begin VB.Label LBLHEADING1 
         BackStyle       =   0  'Transparent
         Caption         =   "Division Name :"
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
         TabIndex        =   28
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label LBLDESC2 
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
         Left            =   7560
         TabIndex        =   27
         Top             =   120
         Width           =   3615
      End
      Begin VB.Shape BORDER2 
         BorderColor     =   &H80000002&
         Height          =   300
         Left            =   5640
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   5535
      End
      Begin VB.Label LBLHEADING2 
         BackStyle       =   0  'Transparent
         Caption         =   "Packing Station :"
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
         Left            =   5760
         TabIndex        =   26
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Carton Name"
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
         Height          =   375
         Left            =   3600
         TabIndex        =   23
         Tag             =   "S"
         Top             =   -2040
         Width           =   1335
      End
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   480
      Top             =   7320
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
End
Attribute VB_Name = "frmBoxPacking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim IsShadeReq As Boolean
Private Const chEventStart = "+"
Dim LOT_MC_CHANGE_OCCUR As Boolean
Dim CallExit As Boolean
Dim SubGradename As String
Dim TWSTREQ As String
Dim PICK_WT As Boolean
Dim ERROROCCUR As Boolean
Dim LOAD As String
Dim DIVCODE As String
Dim LSPKGCOD As String
Dim M_DBCD As String
Dim PKGNGCD As String
Dim MCCD As String
Dim LOCCOD As String
Dim RETURNABLE As String
Dim GRADE As String
Dim SUBGRADE As String
Dim CHALLAN As String
Dim PALETNO As String
Dim SUBPKG As String
Dim SUBPKGCODE As String
Dim INFORS As New ADODB.Recordset
Dim CFGTYP As String
Dim bauardrate As String
Dim COMPORTX As Integer
Dim Returnstring As String
Dim SAVEFLAG As Boolean
Dim ROWNO As Long
Dim SWITCH As Boolean
Dim SQL As String
Dim COUNTER As Long
Dim M_PCOD As String
Dim LASTBOXN As String
Dim FINITMCOD As String
Dim REQNOCOPS As Boolean
Dim REQCOPSWGT As Boolean
Dim REQBOXWGT As Boolean
Dim REQPALLET As Boolean
Dim REQONLP As Boolean

Private Sub cmbPackaging_Click()
  Call SETPLYLIMIT
End Sub

Private Sub cmbPackaging_KeyDown(KeyCode As Integer, Shift As Integer)
   Call SETPLYLIMIT
   KeyCode = 0
End Sub

Private Sub cmbPackingType_KeyDown(KeyCode As Integer, Shift As Integer)
  KeyCode = 0
End Sub

Private Sub cmdExit_Click()
  CallExit = True
  Unload Me
End Sub

Private Sub cmdSave_Click()
On Error GoTo LAST

Dim i As Long, J As Long
Dim RSTMP As ADODB.Recordset
Set RSTMP = New ADODB.Recordset

ERROROCCUR = False

Dim TABLENAME As String

If BOXNO = "XXXXXXXXXX" Then Exit Sub

If CheckData(ROWNO) Then Exit Sub

If Not SAVEFLAG Then
If IsDispatchExist(DIVCODE, LSPKGCOD, BOXNO) Then
   MsgBox "Boxno. " & BOXNO & " has been dispatched."
   BOXNO.Caption = GenPackSlipNo(LSPKGCOD)
   Call CLEARDATA
   SAVEFLAG = True
   SWITCH = False
   TXTVBDT.Enabled = True
   Exit Sub
End If
End If

Call SetGlobal
RETURNABLE = "Y"
CN.BeginTrans

If SAVEFLAG Then

BOXNO.Caption = GenPackSlipNo(LSPKGCOD)

COUNTER = COUNTER + 1

If IsBoxExistInUnit(BOXNO.Caption) Then
   MsgBox "BoxNo. " & BOXNO.Caption & " Already Exist."
   CN.RollbackTrans
   Exit Sub
End If

SQL = "INSERT INTO BOXREGISTER(COMP,UNIT,DVCD,DBCD,VTYP,VBNO,PLTNO,VBDT,CHLN,PKG_STCOD,PKGNG_COD,"
SQL = SQL & "LOCCOD,PCOD,ISRETURNABLE,LOTNO,ICOD,GRAD,SUBGRD,MCCD,COPS,BOXWGT,COPSWGT,GRSWGT,TRWGT,"
SQL = SQL & "NTWGT,PACKER,RMRK,RECSTAT)VALUES('" & compPth & _
"','" & UNCD & "','" & DIVCODE & "','" & M_DBCD & "','PPF','" & BOXNO & "','" & PALETNO & _
"','" & Format(TXTVBDT, "MM/DD/YYYY") & "','" & CHALLAN & _
"','" & LSPKGCOD & "','" & PKGNGCD & "','" & LOCCOD & "','" & M_PCOD & "','" & RETURNABLE & "','" & txtLTNo & _
"','" & FINITMCOD & "','" & GRADE & "','" & FindSubGradeCode & "','" & MCCD & "','" & Val(TXTCOP) & _
"','" & Val(TXTCTWT) & "','" & Val(TXTCPWT) & "','" & Val(TXTGRWT) & "','" & Val(TXTTRWT) & _
"','" & Val(TXTNTWT) & "','" & cUName & "','" & TXTRMRK & "','A')"

CN.Execute SQL

Call SetRawMaterial
If ERROROCCUR Then Exit Sub

'UPDATE VOUCHER TYPE MASTER
CN.Execute "UPDATE PCKMST SET [LBNO]='" & BOXNO & "',LSTPCKDT = '" & Format(TXTVBDT, "MM/DD/YYYY") & _
           "' WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND CODE='" & LSPKGCOD & "'"
TXTVBDT.MinDate = Format(TXTVBDT, "DD/MM/YYYY")
Call FillDetail
If ERROROCCUR Then Exit Sub

Else

Call SetRawMaterial

SQL = "UPDATE BOXREGISTER SET VBDT='" & Format(TXTVBDT, "MM/DD/YYYY") & "',PKGNG_COD='" & PKGNGCD & _
"',LOCCOD='" & LOCCOD & "',MCCD='" & MCCD & "',ISRETURNABLE='" & RETURNABLE & "',LOTNO='" & txtLTNo & _
"',ICOD='" & FINITMCOD & "',GRAD='" & GRADE & "',SUBGRD='" & FindSubGradeCode & _
"',COPS='" & TXTCOP & "',BOXWGT='" & TXTCTWT & "',COPSWGT='" & TXTCPWT & "',GRSWGT='" & TXTGRWT & _
"',TRWGT='" & TXTTRWT & "',NTWGT='" & TXTNTWT & "',RMRK='" & TXTRMRK & "',CHLN = '" & CHALLAN & "' WHERE COMP='" & compPth & _
"' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCODE & "' AND DBCD='" & M_DBCD & _
"' AND VTYP='PPF' AND VBNO='" & BOXNO & "' AND PKG_STCOD='" & LSPKGCOD & "' AND RECSTAT<>'D'"

CN.Execute SQL

If ERROROCCUR Then Exit Sub
Call FillDetail
If ERROROCCUR Then Exit Sub
End If

'GR SUMMARY ENTRY : GOODS RETURN
If InStr(1, cmbPackingType.Text, "GR") <> 0 Then   'GR Case
   Call SetGRPacking
End If

'PLY UPDATION COMMON FOR BOTH SAVE AND EDIT
If FLEXPLY.Enabled Then
If RSTMP.State = 1 Then RSTMP.Close
RSTMP.Open "SELECT * FROM BOXREGISTER WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
"' AND DVCD='" & DIVCODE & "' AND DBCD='" & M_DBCD & _
"' AND VTYP='PPF' AND VBNO='" & BOXNO & "' AND PKG_STCOD='" & LSPKGCOD & "' AND RECSTAT<>'D'", CN, adOpenDynamic, adLockOptimistic

If Not RSTMP.EOF Then
If FLEXPLY.Cols > 1 Then RSTMP![Top] = 1
If FLEXPLY.Cols > 1 Then RSTMP!Bottom = 1

i = 0
  For i = 1 To FLEXPLY.Cols - 1
    J = 0
    For J = 0 To RSTMP.Fields.COUNT - 1
      If Trim(RSTMP.Fields(J).NAME) = Trim(FLEXPLY.TextMatrix(0, i)) Then
         RSTMP.Fields(J).Value = Val(FLEXPLY.TextMatrix(1, i))
      End If
    Next
  Next
RSTMP.Update
End If
End If
'-------------------------------------------------

'FOR PALLET
 If GRPLAT.Value = 1 Then
    CN.Execute "UPDATE PCKMST SET [LPNO]='" & PALETNO & "' WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
              "' AND CODE='" & LSPKGCOD & "'"
    PALETNO = GenPackSlipNo(LSPKGCOD, "LPNO")
    GRPLAT.Value = 0
 End If
'-------------

LASTBOXN = Trim(BOXNO.Caption)
BOXNO.Caption = GenPackSlipNo(LSPKGCOD)

Call CLEARDATA

SAVEFLAG = True
LOT_MC_CHANGE_OCCUR = False
txtLTNo.Tag = txtLTNo
TXTMCCD.Tag = TXTMCCD

SWITCH = False
TXTVBDT.Enabled = True
BOXNO.Caption = GenPackSlipNo(LSPKGCOD)
'-------------------------------------------------

'DAILYSTATUS ENTRY
 If SAVEFLAG = True Then
     Call DAILYSTATUS("PPF", FINITMCOD, M_DBCD, Val(NETWGT), txtLTNo, 0, cUName, "N", Now, TXTVBDT)
      Else
     Call DAILYSTATUS("PPF", FINITMCOD, M_DBCD, Val(NETWGT), txtLTNo, 0, cUName, "M", Now, TXTVBDT)
 End If
'----------------------------
 
CN.CommitTrans

If REQONLP = True Then
    If Dir("C:\DOSPRINT", vbDirectory) = Empty Then MkDir ("C:\DOSPRINT")
    Close #1
    Open "C:\DOSPRINT\" & ComputerName & "SLIP.TXT" For Output As #1
     
     If M_COMPBILL = "STR" Then
        Call frm_PackingSlip.PrintBoxDetail_STR(LSPKGCOD, LASTBOXN)
     ElseIf M_COMPBILL = "CHK" Then
        Call frm_PackingSlip.PrintBoxDetail_CHK(LSPKGCOD, LASTBOXN)
     ElseIf M_COMPBILL = "MAH" Then
        Call frm_PackingSlip.PrintBoxDetail_MAH(LSPKGCOD, LASTBOXN)
     ElseIf M_COMPBILL = "SHL" Then
        Call frm_PackingSlip.PrintBoxDetail_SHL(LSPKGCOD, LASTBOXN)
     ElseIf M_COMPBILL = "NIR" Then
        Call frm_PackingSlip.PrintBoxDetail_NIR(LSPKGCOD, LASTBOXN)
     ElseIf M_COMPBILL = "LKN" Then
        Call frm_PackingSlip.PrintBoxDetail_LKN(LSPKGCOD, LASTBOXN)
     ElseIf M_COMPBILL = "TEX" Then
        Call frm_PackingSlip.PrintBoxDetail_TEX(LSPKGCOD, LASTBOXN)
     ElseIf M_COMPBILL = "MCS" Then
        Call frm_PackingSlip.PrintBoxDetail_MCS(LSPKGCOD, LASTBOXN)
     ElseIf M_COMPBILL = "MCK" Then
        Call frm_PackingSlip.PrintBoxDetail_MCK(LSPKGCOD, LASTBOXN)
     Else
       '2 copy required
       Call PACKINGSLIP_GENERAL
       Call PACKINGSLIP_GENERAL
     End If
     
  Close #1
  Shell App.PATH & "\Reports\PRINTDOC.BAT " & "C:\DOSPRINT\" & ComputerName & "SLIP.TXT", vbHide
End If

Exit Sub
LAST:
MsgBox ERR.Description
CN.RollbackTrans
End Sub

Private Sub cmdSavePrint_Click()
  Call cmdSave_Click
  
  If Dir("C:\DOSPRINT", vbDirectory) = Empty Then MkDir ("C:\DOSPRINT")
  Close #1
  Open "C:\DOSPRINT\" & ComputerName & "SLIP.TXT" For Output As #1
     
     If M_COMPBILL = "STR" Then
        Call frm_PackingSlip.PrintBoxDetail_STR(LSPKGCOD, LASTBOXN)
     ElseIf M_COMPBILL = "CHK" Then
        Call frm_PackingSlip.PrintBoxDetail_CHK(LSPKGCOD, LASTBOXN)
     ElseIf M_COMPBILL = "MAH" Then
        Call frm_PackingSlip.PrintBoxDetail_MAH(LSPKGCOD, LASTBOXN)
     ElseIf M_COMPBILL = "SHL" Then
        Call frm_PackingSlip.PrintBoxDetail_SHL(LSPKGCOD, LASTBOXN)
     ElseIf M_COMPBILL = "NIR" Then
        Call frm_PackingSlip.PrintBoxDetail_NIR(LSPKGCOD, LASTBOXN)
     ElseIf M_COMPBILL = "LKN" Then
        Call frm_PackingSlip.PrintBoxDetail_LKN(LSPKGCOD, LASTBOXN)
     ElseIf M_COMPBILL = "TEX" Then
        Call frm_PackingSlip.PrintBoxDetail_TEX(LSPKGCOD, LASTBOXN)
     ElseIf M_COMPBILL = "MCS" Then
        Call frm_PackingSlip.PrintBoxDetail_MCS(LSPKGCOD, LASTBOXN)
     ElseIf M_COMPBILL = "MCK" Then
        Call frm_PackingSlip.PrintBoxDetail_MCK(LSPKGCOD, LASTBOXN)
     Else
       '2 copy required
       Call PACKINGSLIP_GENERAL
       Call PACKINGSLIP_GENERAL
     End If
     
  Close #1
  Shell App.PATH & "\Reports\PRINTDOC.BAT " & "C:\DOSPRINT\" & ComputerName & "SLIP.TXT", vbHide
End Sub

Private Sub FLEX_DblClick()
Dim i As Long, J As Long
   If Trim(Flex.TextMatrix(Flex.ROW, 0)) = Empty Then
      Exit Sub
   End If
   
   SAVEFLAG = False
   If Flex.Rows > 1 And Flex.TextMatrix(Flex.ROW, 1) <> Empty Then
    ROWNO = Flex.ROW
    BOXNO = Flex.TextMatrix(ROWNO, 0)
    
    Call GetBoxInfo(BOXNO)
    
    If IsShadeReq Then
       TXTSHADE = Flex.TextMatrix(ROWNO, 1)
    ElseIf TXTTWIST.Enabled = True Then
       TXTTWIST = Flex.TextMatrix(ROWNO, 1)
    End If
    
    TXTCOP = Flex.TextMatrix(ROWNO, 2)
    TXTCOP.Tag = Flex.TextMatrix(ROWNO, 2)
    TXTCTWT = Flex.TextMatrix(ROWNO, 3)
    
    If Val(Flex.TextMatrix(ROWNO, 4)) < 1 Then
       TXTCPWT = "0" & Flex.TextMatrix(ROWNO, 4)
    Else
       TXTCPWT = Flex.TextMatrix(ROWNO, 4)
    End If
    
    TXTGRWT = Flex.TextMatrix(ROWNO, 5)
    TXTTRWT = Flex.TextMatrix(ROWNO, 6)
    TXTNTWT = Flex.TextMatrix(ROWNO, 7)
    TXTNTWT.Tag = Flex.TextMatrix(ROWNO, 7)
    'TXTRMRK = FLEX.TextMatrix(ROWNO, 8)
    
    J = 0
    i = 8
    Do While (i < Flex.Cols - 1)
      i = i + 1
      J = J + 1
      FLEXPLY.TextMatrix(1, J) = Flex.TextMatrix(ROWNO, i)
    Loop
       
    SWITCH = True
    TXTVBDT.Enabled = False
    
  End If
    
  If Val(Flex.ROW) > 0 Then
     If TXTTWIST.Enabled Then
        TXTTWIST.SetFocus
     ElseIf TXTCOP.Enabled Then
        TXTCOP.SetFocus
     ElseIf TXTGRWT.Enabled Then
        TXTGRWT.SetFocus
     End If
  End If

End Sub

Private Sub FLEXPLY_EnterCell()
  If FLEXPLY.ROW = 0 Then Exit Sub
  FLEXPLY.CellBackColor = RGB(BRED, BGREEN, BBLUE)
  FLEXPLY.ColWidth(FLEXPLY.COL - 1) = 155 * Len(FLEXPLY.TextMatrix(FLEXPLY.ROW - 1, FLEXPLY.COL - 1))
  FLEXPLY.ColWidth(FLEXPLY.COL) = 155 * Len(FLEXPLY.TextMatrix(FLEXPLY.ROW - 1, FLEXPLY.COL))
  If FLEXPLY.COL + 1 < FLEXPLY.Cols Then
    FLEXPLY.ColWidth(FLEXPLY.COL + 1) = 155 * Len(FLEXPLY.TextMatrix(FLEXPLY.ROW - 1, FLEXPLY.COL + 1))
  End If
End Sub

Private Sub FLEXPLY_KeyPress(KeyAscii As Integer)
  If FLEXPLY.COL = 0 And (FLEXPLY.ROW = 1 Or FLEXPLY.ROW = 2) Then Exit Sub
  
  On Error GoTo LAST
  Dim ALLOW_KEY As Boolean
  Dim FWD_COL As Boolean
  Dim ENTER_PRESS As Boolean
  Dim MSTDAT As New ADODB.Recordset
  
  Set MSTDAT = New ADODB.Recordset
  
  FWD_COL = False
  ALLOW_KEY = False
  
  If FLEXPLY.ROW > 0 Then
    If InStr(1, FLEXPLY.TextMatrix(FLEXPLY.ROW, FLEXPLY.COL), ".") > 0 And KeyAscii = 46 Then
        KeyAscii = 0
        Exit Sub
    End If
  End If
  
  If FLEXPLY.ROW > 0 Then    'START ------------>>>
  
  'CASE1
  If KeyAscii >= 48 And KeyAscii <= 57 Then             ' 0- 9
      ALLOW_KEY = True
  Else
      ALLOW_KEY = False
  End If
  
  'CASE2
  If KeyAscii = vbKeyReturn Then
    ENTER_PRESS = True
   Else
    ENTER_PRESS = False
  End If
  
  'CASE3
  If KeyAscii = 8 Then
    Dim lnth As Double
    lnth = Len(FLEXPLY.TextMatrix(FLEXPLY.ROW, FLEXPLY.COL))
    If lnth > 0 Then
      FLEXPLY.TextMatrix(FLEXPLY.ROW, FLEXPLY.COL) = Mid(FLEXPLY.TextMatrix(FLEXPLY.ROW, FLEXPLY.COL), 1, lnth - 1)
      Exit Sub
    End If
  End If
  
  'CASE4
  If ALLOW_KEY = False Then
    If ENTER_PRESS = True Then
       If FLEXPLY.COL > 2 Then FLEXPLY.ColWidth(FLEXPLY.COL - 2) = 0
    Else
      KeyAscii = 0
      Exit Sub
    End If
  End If
  
  'ACTION 1
  If ALLOW_KEY = True Then
    If ENTER_PRESS = False Then
      FLEXPLY.TextMatrix(FLEXPLY.ROW, FLEXPLY.COL) = FLEXPLY.TextMatrix(FLEXPLY.ROW, FLEXPLY.COL) + Chr(KeyAscii)
    End If
  End If
  
  'ACTION 2
  If ENTER_PRESS = True And FLEXPLY.COL + 1 < FLEXPLY.Cols Then
     FLEXPLY.COL = FLEXPLY.COL + 1
  ElseIf ENTER_PRESS = True Then
     cmdSave.SetFocus
  End If
  
  End If   'END<<-----------------------
      
  Exit Sub
LAST:
  MsgBox ERR.Description
  FLEXPLY.SetFocus
  Exit Sub
End Sub

Private Sub FLEXPLY_LeaveCell()
FLEXPLY.CellBackColor = vbWhite
End Sub

Private Sub Form_Activate()
  
  If DIVCODE = Empty Or Trim(LBLDESC1.Caption) = "XXXXXXXXXX" Then
     MsgBox "Select Division For Packing."
     Unload Me
  End If
  
  If LSPKGCOD = Empty Or Trim(LBLDESC2.Caption) = "XXXXXXXXXX" Then
     MsgBox "Select Packing Station For Packing."
     Unload Me
  End If
  
 'For Raw Material Consumption Slip
  If CHALLAN = Empty Or PALETNO = Empty Then
     Unload Me
  End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
 If TypeOf ActiveControl Is TextBox Then If ActiveControl.Text = Empty Then Exit Sub
 If UCase(ActiveControl.NAME) = "FLEXPLY" Then Exit Sub
  
  If UCase(ActiveControl.NAME) = "TXTGRWT" Then Exit Sub
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
  
  txtLTNo = Empty: txtLTNo.Tag = Empty
  TXTMCCD = Empty: TXTMCCD.Tag = Empty
  
  Call CenterChild(frm_Main, Me): Call ColorComponent(Me)
  
  ERROROCCUR = False
  
  Me.Left = 50: Me.KeyPreview = True
  SAVEFLAG = True
'-------DIVISION NAME
  M_DESC = Empty: Key = Empty:  NEW_VISIBLE = False:  DIVCODE = Empty
  LBLDESC1.Caption = Empty
  If DIVCODE = Empty Then
    LBLDESC1 = SearchList1("SELECT  TOP 20 CODE,NAME FROM DIVMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND RECSTAT='A' AND CODE<>'000001'", 0, "", "SELECT DIVISION MASTER")
    If LBLDESC1 <> Empty Then DIVCODE = Key Else LBLDESC1 = "???????": Unload Me
  End If
       
  LBLCFG.Caption = LabelDisplay(DIVCODE & "", UNCD)
  
  IsShadeReq = False
  
  If IsTwistReq(DIVCODE) = "Y" Then
     TWSTREQ = "Y"
     LBLTWST.Enabled = True: LBLSZO.Enabled = True: TXTTWIST.Enabled = True
     Flex.TextMatrix(0, 1) = "T#"
  ElseIf SetIsShadeReq(DIVCODE) = "Y" Then
     IsShadeReq = True
     LBLTWST.Caption = "Shade"
     LBLTWST.Enabled = True
     TXTTWIST.Enabled = False
     TXTTWIST.Visible = False
     TXTSHADE.Enabled = True
     TXTSHADE.Visible = True
  Else
     Flex.TextMatrix(0, 1) = "SubGrd"
  End If
  
 If PackingType(Key) = "L" Then MsgBox "Division Not Allowed Carton Packing.Check Configuration": LOAD = "N": GoTo JUMP
 '------------------------------------------------------------------------
 'SUB PACKAGING TYPE
 If INFORS.State = 1 Then INFORS.Close
    INFORS.Open "SELECT * FROM DIVMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND CODE = '" & DIVCODE & "' AND RECSTAT='A' ", CN, adOpenDynamic, adLockOptimistic
 If Not INFORS.EOF Then
    CFGTYP = INFORS!CFGTYP & ""
 End If
 
'-------PACKING STATION MASTER
M_DESC = Empty:  Key = Empty:  NEW_VISIBLE = False: LSPKGCOD = Empty
LBLDESC2 = SearchList1("SELECT TOP 20 CODE,NAME FROM PCKMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "'", 0, "", "SELECT PACKING STATION FROM MASTER LIST")
If Key = Empty Then Exit Sub
LSPKGCOD = Key

If RS.State = 1 Then RS.Close
RS.Open "SELECT * FROM PCKMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND CODE ='" & LSPKGCOD & "'", CN, adOpenDynamic, adLockOptimistic
If Not RS.EOF Then
   PICK_WT = IIf(Trim(RS!WSCALE) = "Y", True, False)
   If PICK_WT = True Then
      MSComm1.Settings = LCase(Trim(RS!Settings & ""))      'bauardrate = Trim(RS!BAURDRATE) + ",N,8,1"
      MSComm1.CommPort = Val(Trim(RS!COMPORTX & ""))      'COMPORTX = Val(Trim(RS!COMPORTX))
      MSComm1.Handshaking = Val(Trim(RS!FLOW & ""))
      Call CompPortConnect
   End If
   
   REQNOCOPS = IIf(Trim(RS!REQNOCOPS) = "Y", True, False)
   REQCOPSWGT = IIf(Trim(RS!REQCOPSWGT) = "Y", True, False)
   REQBOXWGT = IIf(Trim(RS!REQBOXWGT) = "Y", True, False)
   REQPALLET = IIf(Trim(RS!REQPALLET) = "Y", True, False)
   REQONLP = IIf(Trim(RS!ONLP) = "Y", True, False)
End If

'---------------------------
'For Raw Material Consumption Slip
CHALLAN = GenPackSlipNo(LSPKGCOD, "LCNO")
'FOR PALLET NO.
PALETNO = GenPackSlipNo(LSPKGCOD, "LPNO")
'For Box No.
BOXNO.Caption = GenPackSlipNo(LSPKGCOD)

COUNTER = 0

TXTVBDT.MinDate = FSDT
TXTVBDT.MaxDate = FEDT

TXTVBDT.Value = Now

Call SetLastDateForPacking

Call SetPackingType
Call setHeading

If cmbPackingType.ListCount > 1 Then cmbPackingType.ListIndex = 2
If cmbPackaging.ListCount > 0 Then cmbPackaging.ListIndex = 0

JUMP:
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  
  If COUNTER > 0 Then
    CN.Execute "UPDATE PCKMST SET [LCNO]='" & CHALLAN & "' WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
                "' AND CODE='" & LSPKGCOD & "'"
  End If
  
  CallExit = True
  
  If MSComm1.PortOpen Then
     MSComm1.PortOpen = False
  End If
  
  Unload Me
End Sub

Private Sub ProcessEvent(stEvent As String)
    TXTGRWT.Text = Val(stEvent)
End Sub

Private Sub GRPLAT_GotFocus()
 If Not REQPALLET Then
     GRPLAT.Enabled = False
     Exit Sub
  End If
End Sub

Private Sub MSComm1_OnComm()
Static stEvent             As String                       'storage for an Identifier event
    Dim stComChar               As String * 1                   'temporary storage for received comm port data
    Select Case MSComm1.CommEvent
        Case comEvReceive                                      ' Received RThreshold # of chars.
          '----------------------------------------------------------------------------------------------
          'The following illustrates how the Identifier is designed
          'to make authoring software easy as '123' for developers:
          '1) Look for a "+" character which indicates the beginning of an event
          '2) Save subsequent characters until you detect a carriage return
          '3) Process the Event
          '----------------------------------------------------------------------------------------------
            Do
                stComChar = MSComm1.Input                         'read 1 character .Inputlen = 1
                Select Case stComChar
                Case chEventStart                           'Beginning of Identifier event
                     stEvent = ""
                Case vbLf                                   'Ignore linefeeds
                Case vbCr                                   'The CR indicates the end of the Identifier Event
                     ProcessEvent stEvent                    'Process the Identifier event
                Case Else
                     stEvent = stEvent + stComChar           'Save everything between the + and CR
                End Select
            Loop While MSComm1.InBufferCount                      'Loop until all characters in receive buffer are processed
    End Select
End Sub

Private Sub txtCop_KeyPress(KeyAscii As Integer)

If KeyAscii < 48 Or KeyAscii > 57 Then             ' 0- 9
   If KeyAscii <> 8 Then KeyAscii = 0
End If
End Sub

Private Sub TXTCPWT_Change()
 TXTTRWT = nstr((Val(TXTCOP) * Val(TXTCPWT)) + Val(TXTCTWT), 9, 3)
 TXTTRWT = Trim(TXTTRWT)
 TXTNTWT = nstr(Val(TXTGRWT) - Val(TXTTRWT), 9, 3)
 TXTNTWT = Trim(TXTNTWT)
End Sub

Private Sub TXTGRAD_Change()
'For Raw Material Consumption Slip
CHALLAN = GenPackSlipNo(LSPKGCOD, "LCNO")
End Sub

Private Sub TXTGRAD_KeyDown(KeyCode As Integer, Shift As Integer)
'If COUNTER > 0 Then Exit Sub
  If Trim(TXTGRAD.Text) = Empty Or KeyCode = vbKeyF2 Then
    NEW_VISIBLE = False: Key = Empty
    TXTGRAD.Text = SearchList1("SELECT TOP 20 CODE,GRAD FROM GRDMST", 0, TXTGRAD, "SELECT " & LBLCFG.Caption)
    TXTGRAD.Tag = Key
  End If
End Sub

Private Sub TXTGRWT_Change()
TXTNTWT = Val(TXTGRWT) - Val((Val(TXTCOP) * Val(TXTCPWT)) + Val(TXTCTWT))
TXTNTWT = nstr(TXTNTWT, 9, 3)
TXTNTWT = Trim(TXTNTWT)
End Sub

Private Sub TXTGRWT_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
If MSComm1.PortOpen = True Then
   If Val(TXTGRWT) > 0 Then
      MSComm1.PortOpen = False
   End If
End If
End Sub

Private Sub TXTLOC_KeyDown(KeyCode As Integer, Shift As Integer)
  Key = Empty
  If (KeyCode = vbKeyReturn And Trim(txtLoc.Text) = Empty) Or KeyCode = vbKeyF2 Then
    txtLoc.Text = SearchList1("SELECT  TOP 20 CODE,NAME FROM LOCMST", 0, txtLoc, "SELECT LOCATION FROM MASTER")
  End If
End Sub

Private Sub txtLTNO_Change()
'For Raw Material Consumption Slip
CHALLAN = GenPackSlipNo(LSPKGCOD, "LCNO")
If txtLTNo <> Empty Then FindFinishItem
End Sub

Private Sub txtltno_GotFocus()
  txtLTNo.BackColor = RGB(BRED, BGREEN, BBLUE)
  SendKeys "{HOME}+{END}"
  'If COUNTER > 0 Then Exit Sub
  ToolTip Me, "Press {F2} / {Enter} For Lot Master Help", "", txtLTNo.Left - 50, txtLTNo.Top + txtLTNo.Height + 100
End Sub

Private Sub txtDENI_GotFocus()
  TXTDENI.BackColor = RGB(BRED, BGREEN, BBLUE)
  SendKeys "{HOME}+{END}"
  If COUNTER > 0 Then Exit Sub
  ToolTip Me, "Finish Item of Lot : " & txtLTNo, "", TXTDENI.Left - 120, TXTDENI.Top + TXTDENI.Height + 100
End Sub

Private Sub TXTGRAD_GotFocus()
  TXTGRAD.BackColor = RGB(BRED, BGREEN, BBLUE)
  SendKeys "{HOME}+{END}"
  'If COUNTER > 0 Then Exit Sub
  ToolTip Me, "Press {F2} / {Enter} For Grade Master Help", "", TXTGRAD.Left - 3820, TXTGRAD.Top + TXTGRAD.Height + 100
End Sub

Private Sub TXTLOC_GotFocus()
  txtLoc.BackColor = RGB(BRED, BGREEN, BBLUE)
  SendKeys "{HOME}+{END}"
  ToolTip Me, "Press {F2} / {Enter} For Location Master Help", "", txtLoc.Left - 620, txtLoc.Top + txtLoc.Height + 100
End Sub

Private Sub txtltno_KeyDown(KeyCode As Integer, Shift As Integer)
'If COUNTER > 0 Then Exit Sub
Dim SQL As String: Me.KeyPreview = False
If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
     txtLTNo = Empty
ElseIf KeyCode = vbKeyF2 Or (KeyCode = vbKeyReturn And txtLTNo = Empty) Then
   M_DESC = Empty:   NEW_VISIBLE = False: Key = Empty
   SQL = "SELECT DISTINCT LTNO,LTNO FROM TXULOT WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCODE & "' AND ACTIVE = 'Y' "
   txtLTNo = SearchList(SQL)
End If

If txtLTNo <> Empty Then FindFinishItem

If SAVEFLAG Then
   txtLTNo.Tag = txtLTNo
   LOT_MC_CHANGE_OCCUR = False
Else
   If txtLTNo.Tag <> txtLTNo Then
      LOT_MC_CHANGE_OCCUR = True
   End If
End If

Me.KeyPreview = True
End Sub

Private Sub TXTMCCD_Change()
CHALLAN = GenPackSlipNo(LSPKGCOD, "LCNO")
End Sub

Private Sub TXTMCCD_GotFocus()
  TXTMCCD.BackColor = RGB(BRED, BGREEN, BBLUE)
  SendKeys "{HOME}+{END}"
  'If COUNTER > 0 Then Exit Sub
  ToolTip Me, "Press {F2} / {Enter} For Machine Master Help", "", TXTMCCD.Left - 620, TXTMCCD.Top + TXTMCCD.Height + 100
End Sub

Private Sub TXTMCCD_KeyDown(KeyCode As Integer, Shift As Integer)
'If COUNTER > 0 Then Exit Sub
Me.KeyPreview = False
   If KeyCode = vbKeyF2 Or (Trim(TXTMCCD.Text) = Empty And KeyCode = 13) Then
        NEW_VISIBLE = False:  M_DESC = Empty:   Key = Empty
        TXTMCCD.Text = SearchList1("SELECT  TOP 20 CODE,NAME FROM MACMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCODE & "'", 0, TXTMCCD, "List of Machine Name")
   ElseIf KeyCode = vbKeyDelete Then
        TXTMCCD = Empty
   End If
   
   If SAVEFLAG Then
      TXTMCCD.Tag = TXTMCCD
      LOT_MC_CHANGE_OCCUR = False
   Else
      If TXTMCCD.Tag <> TXTMCCD Then
         LOT_MC_CHANGE_OCCUR = True
      End If
   End If
   
Me.KeyPreview = True
End Sub

Private Sub txtPCOD_GotFocus()
  TXTPCOD.BackColor = RGB(BRED, BGREEN, BBLUE)
  SendKeys "{HOME}+{END}"
  If COUNTER > 0 Then Exit Sub
  ToolTip Me, "Press {F2} / {Enter} For Party Master Help", "", TXTPCOD.Left - 620, TXTPCOD.Top + TXTPCOD.Height + 100
End Sub

Private Sub TXTPCOD_KeyDown(KeyCode As Integer, Shift As Integer)
If COUNTER > 0 Then Exit Sub
 Me.KeyPreview = False
   If KeyCode = vbKeyF2 Or Trim(TXTPCOD.Text) = Empty Then
        NEW_VISIBLE = False
        M_DESC = Empty
        Key = Empty
        TXTPCOD.Text = SearchList1("SELECT  TOP 20 CODE,NAME FROM ACCMST", 0, TXTPCOD, "List of Job Party A/c")
   ElseIf KeyCode = vbKeyDelete Then
        TXTPCOD = Empty
    End If
 Me.KeyPreview = True
End Sub

Private Sub TXTRMRK_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And FLEXPLY.Cols > 1 Then
  FLEXPLY.ROW = 1
  FLEXPLY.COL = 1
End If
End Sub

Private Sub TXTSHADE_GotFocus()
  TXTSHADE.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTSHADE_KeyDown(KeyCode As Integer, Shift As Integer)
If (Trim(TXTSHADE.Text) = Empty And KeyCode = 13) Or KeyCode = vbKeyF2 Then
    NEW_VISIBLE = False: Key = Empty
    TXTSHADE.Text = SearchList1("SELECT DISTINCT SUBGRD,NAME FROM SUBGRDMST", 0, TXTSHADE, "SELECT SHADE")
    TXTSHADE.Tag = Key
  End If
End Sub

Private Sub TXTSHADE_LostFocus()
  TXTSHADE.BackColor = vbWhite
End Sub

Private Sub txtTwist_KeyPress(KeyAscii As Integer)
 Select Case KeyAscii
        Case Asc("s"), Asc("S")
            TXTTWIST = Empty
            KeyAscii = Asc("S")
        Case Asc("z"), Asc("Z")
            TXTTWIST = Empty
            KeyAscii = Asc("Z")
        Case Asc("0")
            TXTTWIST = Empty
            KeyAscii = Asc("O")
        Case Else
            KeyAscii = 0
 End Select
End Sub

Private Sub TXTCPWT_GotFocus()
    
    If Not REQCOPSWGT Then
      LBLCOPSWGT.Enabled = False
      'SendKeys "{TAB}"
      TXTCPWT = Empty
      TXTCPWT.Enabled = False
      Exit Sub
    End If
    
    TXTCPWT.BackColor = RGB(BRED, BGREEN, BBLUE)
    TXTCPWT.SelStart = 2
    If Len(TXTCPWT) > 2 Then TXTCPWT.SelLength = Len(TXTCPWT) - 2
End Sub

Private Sub txtTwist_GotFocus(): TXTTWIST.BackColor = RGB(BRED, BGREEN, BBLUE): SendKeys "{HOME}+{END}": End Sub
Private Sub txtCTWT_GotFocus()
  If Not REQBOXWGT Then
     LBLBOXWGT.Enabled = False
     TXTCTWT.Enabled = False
     'SendKeys "{TAB}"
     Exit Sub
  End If

  TXTCTWT.BackColor = RGB(BRED, BGREEN, BBLUE)
  SendKeys "{HOME}+{END}"

End Sub

Private Sub txtCop_GotFocus()
  If Not REQNOCOPS Then
     LBLNOCOPS.Enabled = False
     TXTCOP.Enabled = False
     'SendKeys "{TAB}"
     Exit Sub
  End If
  
  TXTCOP.BackColor = RGB(BRED, BGREEN, BBLUE)
  SendKeys "{HOME}+{END}"
End Sub

Private Sub txtGRWT_GotFocus()
On Error GoTo LAST

  If PICK_WT Then
     TXTGRWT = Round(Val(TXTGRWT), 2)
     MSComm1.Output = "ATX,5,7500" + vbCr
  End If
  
  TXTGRWT.BackColor = RGB(BRED, BGREEN, BBLUE)
  SendKeys "{HOME}+{END}"
Exit Sub
LAST:
MsgBox "UNABLE TO CONNECT WITH WEIGHING SCALE"
End Sub

Private Sub TXTRMRK_GotFocus(): TXTRMRK.BackColor = RGB(BRED, BGREEN, BBLUE): SendKeys "{HOME}+{END}": End Sub

Private Sub TXTVBDT_Change()
  CHALLAN = GenPackSlipNo(LSPKGCOD, "LCNO")
End Sub

Private Sub TXTVBDT_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
  If txtLTNo.Enabled Then
     txtLTNo.SetFocus
  End If
End If
End Sub

Private Sub TXTVBDT_KeyPress(KeyAscii As Integer)
If COUNTER > 0 Then KeyAscii = 0
End Sub

Private Sub txtTwist_LostFocus(): TXTTWIST.BackColor = vbWhite: End Sub
Private Sub txtCTWT_LostFocus(): TXTCTWT.BackColor = vbWhite: End Sub
Private Sub txtCop_LostFocus(): TXTCOP.BackColor = vbWhite: End Sub
Private Sub TXTCPWT_LostFocus(): TXTCPWT.BackColor = vbWhite: End Sub
Private Sub txtGRWT_LostFocus(): TXTGRWT.BackColor = vbWhite: End Sub
Private Sub TXTRMRK_LostFocus(): TXTRMRK.BackColor = vbWhite: End Sub
Private Sub TXTGRAD_LostFocus(): TXTGRAD.BackColor = vbWhite: picToolTip.Visible = False: End Sub
Private Sub txtDENI_LostFocus(): TXTDENI.BackColor = vbWhite: picToolTip.Visible = False: End Sub
Private Sub txtltno_LostFocus(): txtLTNo.BackColor = vbWhite: picToolTip.Visible = False: End Sub
Private Sub TXTLOC_LostFocus(): txtLoc.BackColor = vbWhite: picToolTip.Visible = False: End Sub
Private Sub txtPCOD_LostFocus(): TXTPCOD.BackColor = vbWhite: picToolTip.Visible = False: End Sub
Private Sub TXTMCCD_LostFocus(): TXTMCCD.BackColor = vbWhite: picToolTip.Visible = False: End Sub
Private Sub TXTCOP_Change():  Call TXTCPWT_Change: End Sub
Private Sub TXTCTWT_Change(): Call TXTCPWT_Change: End Sub

Private Sub cmbPackaging_KeyPress(KeyAscii As Integer): KeyAscii = 0: End Sub
Private Sub cmbPackingType_KeyPress(KeyAscii As Integer): KeyAscii = 0: End Sub

Private Sub cmbPackingType_Click()
  SendKeys "{HOME}"
  If InStr(1, cmbPackingType.Text, "JobWork") <> 0 Or InStr(1, cmbPackingType.Text, "GR ") <> 0 Then    'PARTY REQUIRED in CASE OF JOB
    TXTPCOD.Enabled = True
    LBLPCOD.Enabled = True
  Else
    TXTPCOD = Empty
    TXTPCOD.Enabled = False
    TXTPCOD.BackColor = vbWhite
    LBLPCOD.Enabled = False
  End If
End Sub

Private Sub cmbPackingType_GotFocus()
 If COUNTER > 0 Then cmbPackingType.Locked = True
 SendKeys "{HOME}+{END}"
End Sub

Private Sub txtCTWT_KeyPress(KeyAscii As Integer)
  If CheckNumericKey(KeyAscii, TXTCTWT, Me) = 0 Then KeyAscii = 0
End Sub

Private Sub txtCPWT_KeyPress(KeyAscii As Integer)
  If CheckNumericKey(KeyAscii, TXTCPWT, Me) = 0 Then KeyAscii = 0
End Sub
Private Sub txtGRWT_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 And Val(TXTGRWT) > 0 And TXTNTWT.Enabled And Val(TXTGRWT) > Val(TXTTRWT) Then
     If GRPLAT.Enabled Then
        GRPLAT.SetFocus
     ElseIf TXTRMRK.Enabled Then
        TXTRMRK.SetFocus
     End If
     Exit Sub
  End If
  If CheckNumericKey(KeyAscii, TXTGRWT, Me) = 0 Then KeyAscii = 0
End Sub

Private Sub txtTRWT_KeyPress(KeyAscii As Integer)
If CheckNumericKey(KeyAscii, TXTTRWT, Me) = 0 Then KeyAscii = 0
End Sub

Private Sub txtNTWT_KeyPress(KeyAscii As Integer)
If CheckNumericKey(KeyAscii, TXTNTWT, Me) = 0 Then KeyAscii = 0
End Sub

Private Sub TimerBillNo1_Timer()
Static ctr As Integer
If ctr Mod 45 = 0 And ctr <= 45 Then
   LBLHEADING1.ForeColor = RGB(LBLRED, LBLGREEN, LBLBLUE): LBLHEADING2.ForeColor = RGB(LBLRED, LBLGREEN, LBLBLUE): BORDER1.BorderColor = RGB(LBLRED, LBLGREEN, LBLBLUE)
   BORDER2.BorderColor = RGB(LBLRED, LBLGREEN, LBLBLUE): LBLDESC1.ForeColor = RGB(LBLRED, LBLGREEN, LBLBLUE): LBLDESC2.ForeColor = RGB(LBLRED, LBLGREEN, LBLBLUE): BOXNO.ForeColor = RGB(LBLRED, LBLGREEN, LBLBLUE)
ElseIf ctr Mod 75 = 0 And ctr <= 75 Then
   LBLHEADING1.ForeColor = vbRed: LBLHEADING2.ForeColor = vbRed: BORDER1.BorderColor = vbRed: BORDER2.BorderColor = vbRed
   LBLDESC1.ForeColor = vbRed: LBLDESC2.ForeColor = vbRed: BOXNO.ForeColor = vbRed
ElseIf ctr Mod 105 = 0 And ctr <= 105 Then
   LBLHEADING1.ForeColor = vbBlue: LBLHEADING2.ForeColor = vbBlue: BORDER1.BorderColor = vbBlue: BORDER2.BorderColor = vbBlue
   LBLDESC1.ForeColor = vbBlue: LBLDESC2.ForeColor = vbBlue: BOXNO.ForeColor = vbBlue
   ctr = 0
End If
ctr = ctr + 15
End Sub

Private Sub FindFinishItem()
Dim RSITM As ADODB.Recordset: Set RSITM = New ADODB.Recordset
Dim FICD As String

If RSITM.State = 1 Then RSITM.Close
RSITM.Open "SELECT FICD FROM TXULOT WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCODE & "' AND LTNO='" & txtLTNo & "'", CN, adOpenDynamic, adLockOptimistic
If Not RSITM.EOF Then FICD = RSITM!FICD
RSITM.Close

If FICD <> Empty Then
  If RSITM.State = 1 Then RSITM.Close
  RSITM.Open "SELECT NAME FROM FINITMMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCODE & "' AND CODE='" & FICD & "'", CN, adOpenDynamic, adLockOptimistic
  If Not RSITM.EOF Then
     TXTDENI = RSITM!NAME
  Else
     TXTDENI = Empty
  End If
  RSITM.Close
End If

End Sub

Private Sub SetPackingType()
Dim PKTYPRS As ADODB.Recordset
Set PKTYPRS = New ADODB.Recordset
If PKTYPRS.State = 1 Then PKTYPRS.Close
PKTYPRS.Open "SELECT NAME FROM SERIALMASTER WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
"' AND VTYP='PPF' AND NAME NOT LIKE '%WASTAGE%' AND NAME NOT LIKE '%GR PACKING%' AND FYCD='" & FYCD & "'", CN, adOpenDynamic, adLockOptimistic
Do While Not PKTYPRS.EOF
    cmbPackingType.AddItem Trim(PKTYPRS!NAME)
    PKTYPRS.MoveNext
Loop

If PKTYPRS.State = 1 Then PKTYPRS.Close
PKTYPRS.Open "SELECT NAME FROM PKGNGMST WHERE STATUS='A' AND RECSTAT='A'", CN, adOpenDynamic, adLockOptimistic
Do While Not PKTYPRS.EOF
 cmbPackaging.AddItem Trim(PKTYPRS!NAME)
PKTYPRS.MoveNext
Loop
End Sub

Private Sub setHeading()
Dim i As Long, J As Long
With FLEXPLY
    .TextMatrix(0, 0) = "PlyName"
    .TextMatrix(1, 0) = "No.ofPly"
    .ColWidth(0) = 850
    .ColWidth(0) = 850
End With

Dim COUNT As Long
Dim GETRS As ADODB.Recordset
Set GETRS = New ADODB.Recordset

If GETRS.State = 1 Then GETRS.Close
GETRS.Open "SELECT * FROM PLYMST WHERE RECSTAT='A'", CN, adOpenDynamic, adLockOptimistic
If Not GETRS.EOF Then
   FLEXPLY.Cols = GETRS.RecordCount + 1
End If

Do While Not GETRS.EOF
    COUNT = COUNT + 1
    FLEXPLY.TextMatrix(0, COUNT) = Trim(GETRS!NAME & "")
    FLEXPLY.ColWidth(COUNT) = 155 * Len(Trim(GETRS!NAME & ""))
GETRS.MoveNext
Loop
GETRS.Close

FLEXPLY.Cols = COUNT + 1

With Flex
 .Cols = 9 + COUNT
 
 .TextMatrix(0, 0) = "Box No.": .TextMatrix(0, 1) = "T#": .TextMatrix(0, 2) = "Cops"
 .TextMatrix(0, 3) = "Box Wgt.": .TextMatrix(0, 4) = "Cops Wgt.": .TextMatrix(0, 5) = "Gross Wgt."
 .TextMatrix(0, 6) = "Tare Wgt.": .TextMatrix(0, 7) = "Net Wgt.": .TextMatrix(0, 8) = "Rmrk"
 
 If IsShadeReq Then
    .TextMatrix(0, 1) = "Shade"
 End If
 
 J = 8
 For i = 1 To FLEXPLY.Cols - 1
    J = J + 1
    Flex.TextMatrix(0, J) = FLEXPLY.TextMatrix(0, i)
    .ColWidth(J) = 0
 Next
    
 .ColWidth(0) = 1250: .ColWidth(1) = 750: .ColWidth(2) = 500: .ColWidth(3) = 550
 .ColWidth(4) = 550: .ColWidth(5) = 1000: .ColWidth(6) = 950: .ColWidth(7) = 950: .ColWidth(8) = 950
  
End With

 If REQNOCOPS = False And REQCOPSWGT = False And REQBOXWGT = False Then
    LBLBOXNO.Caption = "Roll No."
    Flex.TextMatrix(0, 0) = "Roll No."
 End If
 
 If SetIsShadeReq(DIVCODE) = "Y" Then
    Flex.ColWidth(1) = 1250
    Flex.ColAlignment(1) = 1
 End If
 
 Flex.ColAlignment(8) = vbLeftJustify
 
 If REQNOCOPS = False Then Flex.ColWidth(2) = 0
 If REQBOXWGT = False Then Flex.ColWidth(3) = 0
 If REQCOPSWGT = False Then Flex.ColWidth(4) = 0
 
End Sub

Private Sub SetGlobal()
Dim DBCDRS As ADODB.Recordset
Set DBCDRS = New ADODB.Recordset

If DBCDRS.State = 1 Then DBCDRS.Close
DBCDRS.Open "SELECT CODE FROM SERIALMASTER WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
"' AND VTYP='PPF' AND NAME = '" & cmbPackingType.Text & "' AND FYCD='" & FYCD & "'", CN, adOpenDynamic, adLockOptimistic
If Not DBCDRS.EOF Then
   M_DBCD = Trim(DBCDRS!CODE & "")
Else
   M_DBCD = Empty
End If
DBCDRS.Close

If DBCDRS.State = 1 Then DBCDRS.Close
DBCDRS.Open "SELECT CODE FROM PKGNGMST WHERE NAME='" & cmbPackaging.Text & "' AND RECSTAT='A'", CN, adOpenDynamic, adLockOptimistic
If Not DBCDRS.EOF Then
   PKGNGCD = Trim(DBCDRS!CODE & "")
Else
   PKGNGCD = Empty
End If
DBCDRS.Close



If DBCDRS.State = 1 Then DBCDRS.Close
DBCDRS.Open "SELECT CODE FROM LOCMST WHERE NAME='" & txtLoc.Text & "'", CN, adOpenDynamic, adLockOptimistic
If Not DBCDRS.EOF Then
   LOCCOD = Trim(DBCDRS!CODE & "")
Else
   LOCCOD = Empty
End If
DBCDRS.Close

If DBCDRS.State = 1 Then DBCDRS.Close
DBCDRS.Open "SELECT CODE FROM MACMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
"' AND DVCD = '" & DIVCODE & "' AND NAME='" & TXTMCCD.Text & "'", CN, adOpenDynamic, adLockOptimistic
If Not DBCDRS.EOF Then
   MCCD = Trim(DBCDRS!CODE & "")
Else
   MCCD = Empty
End If
DBCDRS.Close

GRADE = GetCode("GRDMST", TXTGRAD, "GRAD", "CODE")

FINITMCOD = FindFinItemCode

End Sub

Private Function FindFinItemCode() As String
Dim GRRS As ADODB.Recordset
Set GRRS = New ADODB.Recordset

If GRRS.State = 1 Then GRRS.Close
GRRS.Open "SELECT * FROM FINITMMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCODE & "' AND NAME ='" & TXTDENI & "'", CN, adOpenDynamic, adLockOptimistic
If Not GRRS.EOF Then
   FindFinItemCode = GRRS!CODE
   RETURNABLE = Trim(GRRS!ISRETURNABLE & "")
Else
   FindFinItemCode = Empty
   RETURNABLE = "N"
End If
GRRS.Close
End Function

Private Function FindSubGradeCode() As String
SubGradename = ""

Dim LOTRS As ADODB.Recordset
Set LOTRS = New ADODB.Recordset
Dim GRRS As ADODB.Recordset
Set GRRS = New ADODB.Recordset
Dim COPSWGT As Double

If IsShadeReq Then
   If GRRS.State = 1 Then GRRS.Close
   GRRS.Open "SELECT SUBGRD FROM SUBGRDMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
             "' AND NAME='" & TXTSHADE & "'", CN, adOpenDynamic, adLockOptimistic
   If Not GRRS.EOF Then
      FindSubGradeCode = Trim(GRRS!SUBGRD & "")
      Exit Function
   End If
   GRRS.Close
End If

If TWSTREQ = "Y" Then
   FindSubGradeCode = Trim(TXTTWIST)
   Exit Function
End If

If CFGTYP = "SG" Then
If LOTRS.State = 1 Then LOTRS.Close
LOTRS.Open "SELECT * FROM TXULOT WHERE COMP = '" & compPth & "' AND UNIT = '" & UNCD & "' AND DVCD = '" & DIVCODE & _
           "' AND LTNO = '" & Trim(txtLTNo) & "'", CN, adOpenDynamic, adLockOptimistic
If Not LOTRS.EOF Then
       SUBPKGCODE = Trim(LOTRS!SUBPKGCODE & "")
End If
End If

If Val(TXTCOP) > 0 Then
   COPSWGT = Val(TXTNTWT) / Val(TXTCOP)
Else
   COPSWGT = Val(TXTNTWT)
End If

If GRRS.State = 1 Then GRRS.Close
If CFGTYP = "SG" And SUBPKGCODE <> "" Then
GRRS.Open "SELECT NAME,SUBGRD FROM SUBGRDMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
"' AND  DVCD='" & DIVCODE & "' AND GRAD='" & GRADE & "' AND SWGT <= " & COPSWGT & _
" AND EWGT >= " & COPSWGT & " AND SUBPKGCODE = '" & SUBPKGCODE & "'", CN, adOpenDynamic, adLockOptimistic

Else
GRRS.Open "SELECT NAME,SUBGRD FROM SUBGRDMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
"' AND  DVCD='" & DIVCODE & "' AND GRAD='" & GRADE & "' AND SWGT <= " & COPSWGT & _
" AND EWGT >= " & COPSWGT & "", CN, adOpenDynamic, adLockOptimistic
End If

If Not GRRS.EOF Then
   FindSubGradeCode = Trim(GRRS!SUBGRD & "")
   SubGradename = Trim(GRRS!NAME & "")
Else
   Msg "Sub Grade Not Properly Defined"
   FindSubGradeCode = Trim(TXTTWIST)
   SubGradename = Trim(TXTTWIST)
End If

GRRS.Close
End Function

Private Sub FillDetail()
On Error GoTo ERRFLEX
 Dim i As Long, J As Long
 Dim INDEX As Long
 
 If Not SWITCH Then
    ROWNO = Flex.Rows - 1
 End If
    
    Flex.TextMatrix(ROWNO, 0) = BOXNO
    
    If IsShadeReq Then
       Flex.TextMatrix(ROWNO, 1) = TXTSHADE
    ElseIf TXTTWIST.Enabled = False Then
       Flex.TextMatrix(ROWNO, 1) = SubGradename
    Else
       Flex.TextMatrix(ROWNO, 1) = Trim(TXTTWIST)
    End If
    
    
    Flex.TextMatrix(ROWNO, 2) = Trim(TXTCOP)
    Flex.TextMatrix(ROWNO, 3) = Trim(nstr(Val(TXTCTWT), 12, 3))
    Flex.TextMatrix(ROWNO, 4) = Trim(nstr(Val(TXTCPWT), 12, 3))
    Flex.TextMatrix(ROWNO, 5) = Trim(nstr(Val(TXTGRWT), 12, 3))
    Flex.TextMatrix(ROWNO, 6) = Trim(nstr(Val(TXTTRWT), 12, 3))
    Flex.TextMatrix(ROWNO, 7) = Trim(nstr(Val(TXTNTWT), 12, 3))
    Flex.TextMatrix(ROWNO, 8) = Trim(TXTRMRK) & " " & Trim(TXTDENI)
        
    J = 8
    For i = 1 To FLEXPLY.Cols - 1
      J = J + 1
      Flex.TextMatrix(ROWNO, J) = FLEXPLY.TextMatrix(1, i)
    Next
    
    If Not SWITCH Then
        If COUNTER > 1 Then Call HighlightRow(ROWNO - 1, vbWhite)
        Call HighlightRow(ROWNO, RGB(214, 218, 254))
        NETBOXES = Flex.Rows - 1
    Else
        Call HighlightRow(ROWNO, RGB(255, 255, 218))
    End If
    
    If Flex.TextMatrix(Flex.Rows - 1, 1) <> "" Then
       Flex.Rows = Flex.Rows + 1
    End If
    If Flex.Rows > 17 Then Flex.TopRow = Flex.TopRow + 1
        
    If TXTTWIST.Enabled Then
       TXTTWIST.SetFocus
    ElseIf TXTCOP.Enabled Then
       TXTCOP.SetFocus
    ElseIf TXTGRWT.Enabled Then
       TXTGRWT.SetFocus
    End If
               
    'REMOVE BELOW COMMENT BLOCK WHEN ITEMS PROCESS ARE GOING TO MULTIPLE
     SWITCH = False
     TXTVBDT.Enabled = True
     
Exit Sub
ERRFLEX:
CN.RollbackTrans
MsgBox "ERROR IN FLEX"
ERROROCCUR = True
End Sub

Private Function CheckData(RNO As Long) As Boolean
Dim INDEX As Long
 If Val(TXTCOP) <= 0 And TXTCOP.Enabled Then MsgBox "Please Enter No.of Cops !!", vbInformation: TXTCOP.SetFocus: CheckData = True: Exit Function
 If Val(TXTCTWT) <= 0 And TXTCTWT.Enabled Then MsgBox "Please Enter Carton Weight !!", vbInformation, "Weight Missing !!": TXTCTWT.SetFocus: CheckData = True: Exit Function
 If Val(TXTGRWT) <= 0 And TXTGRWT.Enabled Then MsgBox "Please Enter Gross Weight !!", vbInformation, "Weight Missing !!": TXTGRWT.SetFocus: CheckData = True: Exit Function
 
 If Val(TXTNTWT) <= 0 Then MsgBox "Net Weight Can't be Negative or Zero!!", vbInformation, "Weight Missing !!": TXTGRWT.SetFocus: CheckData = True: Exit Function
 
 If TXTDENI = Empty Then MsgBox "Please Select Proper Lot !!", vbInformation, "Item Missing !!": txtLTNo.SetFocus: CheckData = True: Exit Function
 If txtLTNo = Empty Then MsgBox "Please Select Proper Lot !!", vbInformation, "Lot Missing !!": txtLTNo.SetFocus: CheckData = True: Exit Function
 If TXTGRAD = Empty Then MsgBox "Please Select Grade !!", vbInformation, "Grade Missing !!": TXTGRAD.SetFocus: CheckData = True: Exit Function
 If TXTMCCD = Empty Then MsgBox "Please Select Machine !!", vbInformation, "Machine Missing !!": TXTMCCD.SetFocus: CheckData = True: Exit Function
 
 M_PCOD = Empty
 If (InStr(1, UCase(cmbPackingType.Text), "JOBWORK") <> 0 Or InStr(1, UCase(cmbPackingType.Text), "GR") <> 0) And TXTPCOD = Empty Then    'PARTY REQUIRED in CASE OF JOB
   MsgBox "Party Required in case of JobWork/GR PACKING !!", vbCritical
   If TXTPCOD.Enabled Then TXTPCOD.SetFocus
   CheckData = True
   Exit Function
Else
    M_PCOD = GetCode("ACCMST", TXTPCOD, "NAME", "CODE")
End If

If FLEXPLY.Enabled Then
Dim i As Long, TOTPLY As Long: TOTPLY = 0
    For i = 1 To FLEXPLY.Cols - 1
        TOTPLY = TOTPLY + Val(FLEXPLY.TextMatrix(1, i))
    Next i
    
    If TOTPLY <> Val(FLEXPLY.Tag) Then
       MsgBox "Please Enter Exact No. of Ply (" & CStr(Val(FLEXPLY.Tag)) & ") that r defined in Packaging Master!!", vbInformation
       FLEXPLY.SetFocus
       If FLEXPLY.Rows > 1 Then FLEXPLY.ROW = 1
       If FLEXPLY.Cols > 1 Then FLEXPLY.COL = 1
       CheckData = True
       Exit Function
    End If
End If

End Function

Private Sub CLEARDATA()
 Dim i As Long, J As Long
 TXTGRWT = Empty: TXTNTWT = Empty
 For i = 1 To FLEXPLY.Cols - 1
  FLEXPLY.TextMatrix(1, i) = ""
 Next
End Sub

Private Sub HighlightRow(StartRowNumber As Long, Optional RowColor As Long = vbYellow)

 Dim SaveRow As Long
 Dim SaveCol As Long
 Dim SaveFillStyle As Long
 With Flex
 SaveRow = .ROW
 SaveCol = .COL
 SaveFillStyle = .FillStyle
 ' Set the range to be highlighted...
 ' Row and Col must be set before RowSel and ColSel
 .COL = .FixedCols
 .ROW = StartRowNumber
 ' Set the rest of the range to highlight
 .RowSel = StartRowNumber
 .ColSel = .Cols - 1
 ' Force change to all selected cells
 .FillStyle = flexFillRepeat
 ' Cell properties
 .CellBackColor = RowColor
 .ROW = SaveRow
 .COL = SaveCol
 .FillStyle = SaveFillStyle
 End With
End Sub

Private Sub SetRawMaterial()

If InStr(1, cmbPackingType.Text, "GR") <> 0 Then   'GR Case
   Exit Sub
End If

If LOT_MC_CHANGE_OCCUR Then
   Call SetRaw
   Exit Sub
End If

On Error GoTo LAST
Dim ITMCODE As String
Dim ITMQTY As Double, ITMRATE As Double, ITMEDITQTY As Double

Dim MAINRS As ADODB.Recordset
Set MAINRS = New ADODB.Recordset
Dim TEMPRS As ADODB.Recordset
Set TEMPRS = New ADODB.Recordset

Dim TABLE As String, pCode As String, FIELD As String

'CASE : 1
'EITHER INSERT IN STORETRAN(INTERNAL PACKING) OR JOBIN(PACKING FOR JOB WORK)
'CASE : 2
'ONE TIME INSERT AND ALL TIME UPDATE ON BASIS OF PACKING STATION CHALLAN SLIP NUMBER

If InStr(1, cmbPackingType.Text, "JobWork") <> 0 Then    'PARTY REQUIRED in CASE OF JOB
    TABLE = "JOBIN"
    pCode = GetCode("ACCMST", TXTPCOD, "NAME", "CODE")
    FIELD = "JOBQ"
Else
    TABLE = "STORETRAN"
    pCode = MCCD
    FIELD = "BALQ"
End If

Dim FLAG As Boolean 'FOR UPDATE CHALLAN ON PACKING STATION
FLAG = False

If TEMPRS.State = 1 Then TEMPRS.Close
TEMPRS.Open "SELECT RICD,PERC,SRCH FROM TXULOT WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
            "' AND DVCD='" & DIVCODE & "' AND LTNO='" & txtLTNo & "' ORDER BY SRCH", CN, adOpenDynamic, adLockOptimistic
            
Do While Not TEMPRS.EOF
   ITMCODE = Trim(TEMPRS!RICD & "")
   ITMQTY = Val(nstr((Val(TEMPRS!PERC) * TXTNTWT) / 100, 10, 3))
   ITMEDITQTY = Val(nstr((Val(TEMPRS!PERC) * Val(TXTNTWT.Tag)) / 100, 10, 3))
   ITMRATE = Val(GetCode("ITMMST", TEMPRS!RICD, "CODE", "PURR"))
      
   If MAINRS.State = 1 Then MAINRS.Close
   MAINRS.Open "SELECT * FROM STORETRAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
               "' AND VTYP='PPF' AND DBCD='" & M_DBCD & "' AND CHLN='" & CHALLAN & _
               "' AND SRNO = '" & LSPKGCOD & "' AND SRCH='" & Trim(TEMPRS!SRCH & "") & "'", CN, adOpenDynamic, adLockOptimistic
   
   If MAINRS.EOF Then
      MAINRS.AddNew
      FLAG = True
   End If
   
   MAINRS!COMP = compPth
   MAINRS!unit = UNCD
   MAINRS!DVCD = DIVCODE
   MAINRS!VTYP = "PPF"
   MAINRS!dbcd = M_DBCD
   MAINRS!SRNO = LSPKGCOD
   MAINRS!SRCH = Trim(TEMPRS!SRCH & "")
   MAINRS!VBNO = CHALLAN
   MAINRS!chln = CHALLAN
   MAINRS!Date = Format(TXTVBDT, "YYYY/MM/DD")
   MAINRS!CHDT = Format(TXTVBDT, "YYYY/MM/DD")
   MAINRS!PCOD = MCCD
   MAINRS!ICOD = ITMCODE
   MAINRS!PCES = 1
   MAINRS!QNTY = Val(MAINRS!QNTY) + ITMQTY - ITMEDITQTY
   MAINRS!GWGT = 0
   MAINRS!TWGT = 0
   MAINRS!RATE = ITMRATE
   MAINRS!AMNT = Val(MAINRS!QNTY) * Val(MAINRS!RATE)
   MAINRS!QORP = "Q"
   MAINRS![User] = cUName
   MAINRS![SYSR] = "N"
   MAINRS!OPER = "-"
   MAINRS!grad = GRADE
   MAINRS!ltno = txtLTNo
   MAINRS!SUBGRD = FindSubGradeCode
   MAINRS!COPS = 0
   MAINRS!TWST = ""
   MAINRS!RECSTAT = "A"
   MAINRS.Update
TEMPRS.MoveNext
Loop
TEMPRS.Close

If FLAG Then
   CN.Execute "UPDATE PCKMST SET [LCNO]='" & CHALLAN & "' WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
              "' AND CODE='" & LSPKGCOD & "'"
End If

NETWGT = Val(NETWGT) + Val(TXTNTWT) - Val(TXTNTWT.Tag)
NETWGT = nstr(NETWGT, 9, 3): NETWGT = Trim(NETWGT)
NETCOPS = Val(NETCOPS) + Val(TXTCOP) - Val(TXTCOP.Tag)
TXTNTWT.Tag = "0"
TXTCOP.Tag = "0"
 
Exit Sub
LAST:
ERROROCCUR = True
MsgBox ERR.Description
CN.RollbackTrans
End Sub

Private Sub SetRaw()

If InStr(1, cmbPackingType.Text, "GR") <> 0 Then   'GR Case
   Exit Sub
End If

On Error GoTo LAST
Dim ITMCODE As String
Dim ITMQTY As Double, ITMRATE As Double, ITMEDITQTY As Double

Dim MAINRS As ADODB.Recordset
Set MAINRS = New ADODB.Recordset
Dim TEMPRS As ADODB.Recordset
Set TEMPRS = New ADODB.Recordset

Dim TABLE As String, pCode As String, FIELD As String

'CASE : 1
'EITHER INSERT IN STORETRAN(INTERNAL PACKING) OR JOBIN(PACKING FOR JOB WORK)
'CASE : 2
'ONE TIME INSERT AND ALL TIME UPDATE ON BASIS OF PACKING STATION CHALLAN SLIP NUMBER

If InStr(1, cmbPackingType.Text, "JobWork") <> 0 Then    'PARTY REQUIRED in CASE OF JOB
    TABLE = "JOBIN"
    pCode = GetCode("ACCMST", TXTPCOD, "NAME", "CODE")
    FIELD = "JOBQ"
Else
    TABLE = "STORETRAN"
    pCode = MCCD
    FIELD = "BALQ"
End If

If TEMPRS.State = 1 Then TEMPRS.Close
TEMPRS.Open "SELECT RICD,PERC,SRCH FROM TXULOT WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
            "' AND DVCD='" & DIVCODE & "' AND LTNO='" & txtLTNo.Tag & "' ORDER BY SRCH", CN, adOpenDynamic, adLockOptimistic
            
Do While Not TEMPRS.EOF
   ITMCODE = Trim(TEMPRS!RICD & "")
   ITMQTY = Val(nstr((Val(TEMPRS!PERC) * TXTNTWT) / 100, 10, 3))
   ITMEDITQTY = Val(nstr((Val(TEMPRS!PERC) * Val(TXTNTWT.Tag)) / 100, 10, 3))
   ITMRATE = 0
      
   If MAINRS.State = 1 Then MAINRS.Close
   MAINRS.Open "SELECT * FROM STORETRAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
               "' AND VTYP='PPF' AND DBCD='" & M_DBCD & "' AND CHLN='" & CHALLAN & _
               "' AND SRNO = '" & LSPKGCOD & "' AND SRCH='" & Trim(TEMPRS!SRCH & "") & "'", CN, adOpenDynamic, adLockOptimistic
   
   If Not MAINRS.EOF Then
      MAINRS!QNTY = Val(MAINRS!QNTY) - ITMEDITQTY
      MAINRS!AMNT = Val(MAINRS!QNTY) * Val(MAINRS!RATE)
      MAINRS!ltno = txtLTNo.Tag
      MAINRS!PCOD = GetMachineCode(DIVCODE, TXTMCCD.Tag)
      MAINRS.Update
   End If
   
TEMPRS.MoveNext
Loop
TEMPRS.Close

'For Raw Material Consumption Slip
CHALLAN = GenPackSlipNo(LSPKGCOD, "LCNO")

If TEMPRS.State = 1 Then TEMPRS.Close
TEMPRS.Open "SELECT RICD,PERC,SRCH FROM TXULOT WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
            "' AND DVCD='" & DIVCODE & "' AND LTNO='" & txtLTNo & "' ORDER BY SRCH", CN, adOpenDynamic, adLockOptimistic
            
Do While Not TEMPRS.EOF
   ITMCODE = Trim(TEMPRS!RICD & "")
   ITMQTY = Val(nstr((Val(TEMPRS!PERC) * TXTNTWT) / 100, 10, 3))
   ITMEDITQTY = Val(nstr((Val(TEMPRS!PERC) * Val(TXTNTWT.Tag)) / 100, 10, 3))
   ITMRATE = 0
      
   If MAINRS.State = 1 Then MAINRS.Close
   MAINRS.Open "SELECT * FROM STORETRAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
               "' AND VTYP='PPF' AND DBCD='" & M_DBCD & "' AND CHLN='" & CHALLAN & _
               "' AND SRNO = '" & LSPKGCOD & "' AND SRCH='" & Trim(TEMPRS!SRCH & "") & "'", CN, adOpenDynamic, adLockOptimistic
   
   If MAINRS.EOF Then
      MAINRS.AddNew
      MAINRS!COMP = compPth
      MAINRS!unit = UNCD
      MAINRS!DVCD = DIVCODE
      MAINRS!VTYP = "PPF"
      MAINRS!dbcd = M_DBCD
      MAINRS!SRNO = LSPKGCOD
      MAINRS!SRCH = Trim(TEMPRS!SRCH & "")
      MAINRS!VBNO = CHALLAN
      MAINRS!chln = CHALLAN
      MAINRS!Date = Format(TXTVBDT, "YYYY/MM/DD")
      MAINRS!CHDT = Format(TXTVBDT, "YYYY/MM/DD")
      MAINRS!PCOD = MCCD
      MAINRS!ICOD = ITMCODE
      MAINRS!PCES = 1
      MAINRS!QNTY = Val(MAINRS!QNTY) + ITMQTY
      MAINRS!GWGT = 0
      MAINRS!TWGT = 0
      MAINRS!RATE = ITMRATE
      MAINRS!AMNT = Val(MAINRS!QNTY) * Val(MAINRS!RATE)
      MAINRS!QORP = "Q"
      MAINRS![User] = cUName
      MAINRS![SYSR] = "N"
      MAINRS!OPER = "-"
      MAINRS!grad = GRADE
      MAINRS!ltno = txtLTNo
      MAINRS!SUBGRD = ""
      MAINRS!COPS = 0
      MAINRS!TWST = ""
      MAINRS!RECSTAT = "A"
      MAINRS.Update
   
   End If
   
TEMPRS.MoveNext
Loop
TEMPRS.Close

NETWGT = Val(NETWGT) + Val(TXTNTWT) - Val(TXTNTWT.Tag)
NETWGT = nstr(NETWGT, 9, 3): NETWGT = Trim(NETWGT)
NETCOPS = Val(NETCOPS) + Val(TXTCOP) - Val(TXTCOP.Tag)
TXTNTWT.Tag = "0"
TXTCOP.Tag = "0"

CN.Execute "UPDATE PCKMST SET [LCNO]='" & CHALLAN & "' WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
           "' AND CODE='" & LSPKGCOD & "'"
 
Exit Sub
LAST:
ERROROCCUR = True
MsgBox ERR.Description
CN.RollbackTrans
End Sub

Private Sub SETPLYLIMIT()
Dim LIMITRS As ADODB.Recordset
Set LIMITRS = New ADODB.Recordset

If LIMITRS.State = 1 Then LIMITRS.Close
LIMITRS.Open "SELECT * FROM PKGNGMST WHERE STATUS='A' AND RECSTAT='A' AND NAME='" & cmbPackaging.Text & "' AND PALLET='Y'", CN, adOpenDynamic, adLockOptimistic
If Not LIMITRS.EOF Then
 FLEXPLY.Enabled = True
 FLEXPLY.Tag = Val(Trim(LIMITRS!NOPLY & ""))
Else
 FLEXPLY.Enabled = False
End If

LIMITRS.Close

End Sub

Private Sub SetLastDateForPacking()
Dim DTRS As ADODB.Recordset
Set DTRS = New ADODB.Recordset

If DTRS.State = 1 Then DTRS.Close
DTRS.Open "SELECT IsNull(LSTPCKDT,'" & FSDT & "') AS LSTPCKDT FROM PCKMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND CODE='" & LSPKGCOD & "'", CN, adOpenDynamic, adLockOptimistic
If Not DTRS.EOF Then
   TXTVBDT.MinDate = Format(DTRS!LSTPCKDT, "DD/MM/YYYY")
End If
DTRS.Close

If DTRS.State = 1 Then DTRS.Close
DTRS.Open "SELECT IsNull(LSTPCKDT,'" & FEDT & "') AS LSTPCKDT FROM PCKMST WHERE COMP='" & compPth & _
          "' AND UNIT='" & UNCD & "' AND CODE='" & LSPKGCOD & _
          "' AND LSTPCKDT <= '" & Format(TXTVBDT.Value, "MM/DD/YYYY") & "'", CN, adOpenDynamic, adLockOptimistic
If Not DTRS.EOF Then
   TXTVBDT.MaxDate = Format(TXTVBDT, "DD/MM/YYYY")
Else
   TXTVBDT.MaxDate = Format(FEDT, "DD/MM/YYYY")
End If
DTRS.Close

End Sub

Private Function IsBoxExistInUnit(BOXNUM As String) As Boolean
IsBoxExistInUnit = False

Dim CHKRS As ADODB.Recordset
Set CHKRS = New ADODB.Recordset
If CHKRS.State = 1 Then CHKRS.Close
CHKRS.Open "SELECT * FROM BOXREGISTER WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
           "' AND VBNO='" & BOXNUM & "' AND RECSTAT='A'", CN, adOpenDynamic, adLockOptimistic
If Not CHKRS.EOF Then
   IsBoxExistInUnit = True
End If
CHKRS.Close
End Function

Private Sub GetBoxInfo(BOXNUM As String)
Dim RSDATA As ADODB.Recordset
Set RSDATA = New ADODB.Recordset

If RSDATA.State = 1 Then RSDATA.Close
RSDATA.Open "SELECT MACMST.NAME AS MACHINE,* FROM BOXREGISTER " & _
"INNER JOIN MACMST ON MACMST.COMP=BOXREGISTER.COMP AND MACMST.UNIT=BOXREGISTER.UNIT " & _
"AND MACMST.DVCD=BOXREGISTER.DVCD AND MACMST.CODE=BOXREGISTER.MCCD " & _
"WHERE BOXREGISTER.COMP='" & compPth & "' AND BOXREGISTER.UNIT='" & UNCD & _
"' AND BOXREGISTER.DVCD='" & DIVCODE & "' AND (VTYP='PPF' OR VTYP='OPN') AND BOXREGISTER.VBNO='" & BOXNUM & _
"' AND BOXREGISTER.PKG_STCOD='" & LSPKGCOD & "' AND RECSTAT<>'D'", CN, adOpenDynamic, adLockOptimistic
     
If Not RSDATA.EOF Then
 txtLTNo = Trim(RSDATA!LOTNO & "")
 txtLTNo.Tag = Trim(RSDATA!LOTNO & "")
 TXTGRAD = GetCode("GRDMST", Trim(RSDATA!grad & ""), "CODE", "GRAD")
 CHALLAN = Trim(RSDATA!chln & "")
 TXTMCCD = Trim(RSDATA!MACHINE & "")
 TXTMCCD.Tag = Trim(RSDATA!MACHINE & "")
 LOT_MC_CHANGE_OCCUR = False
 TXTRMRK = Trim(RSDATA!RMRK & "")
End If
End Sub

Private Sub PACKINGSLIP_GENERAL()
            CRPT.Reset
            crptConnect CRPT
            
            ReportName = App.PATH & "\Reports\rpt_PackSlip_" & M_COMPBILL & ".rpt"
            
            If Dir(ReportName, vbNormal) = Empty Then
                ReportErrorMessage 1001
                Exit Sub
            End If
            
            CRPT.ReportFileName = ReportName
            rptsql = Empty
            Dim i As Double
            Dim M_BOXN As String
            
            M_BOXN = LASTBOXN
            If M_BOXN = Empty Then
                MsgBox "No Item Selected !!", vbInformation, "Carton Missing !!"
                Exit Sub
            End If
            
            rptsql = "{BOXREG.COMP}='" & compPth & "' AND {BOXREG.UNIT}='" & UNCD & "' AND {BOXREG.VBNO}='" & M_BOXN & "' AND {BOXREG.RECSTAT}<>'D'"
            
            CRPT.ReplaceSelectionFormula rptsql
            RPTN = "Packing Slip"
            
            With CRPT
                RPTN = RPTN + Space(5) + ReportName
                .DiscardSavedData = True
                .WindowTitle = RPTN
                
                
                .Destination = crptToPrinter
                '.Destination = crptToWindow
                
                
                .WindowState = crptMaximized
                .WindowShowProgressCtls = True
                .WindowShowPrintBtn = True
                .WindowShowPrintSetupBtn = True
                .WindowShowRefreshBtn = True
                .WindowShowSearchBtn = True
                .PageLast
                .PageFirst
                .ACTION = 1
                
            End With
            
End Sub

Private Sub SetGRPacking()
On Error GoTo LAST
     
    Dim MAINRS As ADODB.Recordset
    Set MAINRS = New ADODB.Recordset
       
    Dim SRCH As String: SRCH = 0
    
    'IF ENTRY EXIST THEN DELETE RECORD
    CN.Execute "DELETE FROM GRPACKING WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCODE & _
               "' AND PKG_STCOD='" & LSPKGCOD & "' AND VBNO='" & CHALLAN & "' "
    
    'FORM QUERY FROM BOXREGISTER DIRECTLY THROUGH GROUPING
    SQL = "SELECT COMP,UNIT,DVCD,PKG_STCOD,CHLN AS VBNO,PCOD,LOTNO,ICOD,GRAD,SUBGRD," & _
          "ISNULL(SUM(GRSWGT),0) AS GRSWGT,ISNULL(SUM(TRWGT),0) AS TRWGT,ISNULL(SUM(NTWGT),0) AS NTWGT " & _
          ",COUNT(NTWGT) AS BOXES,ISNULL(SUM(COPS),0) AS COPS FROM BOXREGISTER " & _
          "WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCODE & _
          "' AND VTYP IN ('PPF','OPN') AND PKG_STCOD='" & LSPKGCOD & "' AND RECSTAT<>'D' AND CHLN='" & CHALLAN & _
          "' GROUP BY COMP,UNIT,DVCD,PKG_STCOD,CHLN,PCOD,LOTNO,ICOD,GRAD,SUBGRD"

    If MAINRS.State = 1 Then MAINRS.Close
    MAINRS.Open SQL, CN, adOpenDynamic, adLockOptimistic
    Do While Not MAINRS.EOF
    
       SRCH = SRCH + 1
       CN.Execute "INSERT INTO GRPACKING([COMP],[UNIT],[DVCD],[PKG_STCOD],[VBNO],SRCH,[VBDT],[PCOD],[LOTNO]," & _
                  "[ICOD],[GRAD],[SUBGRD],[BOXES],[COPS],[GRSWGT],[TRWGT],[NETWGT],[RECSTAT]) VALUES('" & compPth & _
                  "','" & UNCD & "','" & DIVCODE & "','" & LSPKGCOD & "','" & CHALLAN & _
                  "','" & SRCH & "','" & Format(TXTVBDT, "MM/DD/YYYY") & _
                  "','" & MAINRS!PCOD & "','" & MAINRS!LOTNO & _
                  "','" & MAINRS!ICOD & "','" & MAINRS!grad & "','" & MAINRS!SUBGRD & _
                  "','" & MAINRS!BOXES & "','" & MAINRS!COPS & "','" & MAINRS!GRSWGT & _
                  "','" & MAINRS!TRWGT & "','" & MAINRS!NTWGT & "','A') "
                  
    MAINRS.MoveNext
    Loop
    MAINRS.Close
    
            
Exit Sub
LAST:
MsgBox ERR.Description
CN.RollbackTrans
End Sub

Private Sub CompPortConnect()
On Error GoTo INITERR
  If Not MSComm1.PortOpen Then                              ' Open the comm port if not already open
     MSComm1.PortOpen = True
  End If

  If Not MSComm1.PortOpen Then                              ' if there is a problem opening the port
     MsgBox "Cannot open comm port " & MSComm1.CommPort    ' display an error first
     End                                                 ' bail out of the program
  End If

  ' Initialize communications and update app UI
  MSComm1.RThreshold = 1                                    ' Generate a receive event on every character received
  MSComm1.InputLen = 1                                      ' Read the receive buffer 1 char at a time
  MSComm1.Output = vbCr + "ATSN" + vbCr                     ' Send command to put Identifier in event mode and receive serial number
  MSComm1.Output = "ATX,5,7500" + vbCr             'Set DTMF timeout to 7.5 seconds
  
Exit Sub
INITERR:
MsgBox "When Port is Not Open"
End Sub

