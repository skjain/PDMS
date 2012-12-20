VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{2BF72F7D-D367-4712-9146-5533EF3E691A}#1.2#0"; "FraPlus1.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WELCHBUTTON.OCX"
Begin VB.Form frmOpnBoxPacking 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Packing / Opening Box Packing"
   ClientHeight    =   6855
   ClientLeft      =   375
   ClientTop       =   1110
   ClientWidth     =   11445
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6885.067
   ScaleMode       =   0  'User
   ScaleWidth      =   12478.35
   Begin VB.Timer TimerBillNo 
      Interval        =   100
      Left            =   0
      Top             =   7920
   End
   Begin VB.PictureBox picToolTip 
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      FillColor       =   &H00404040&
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   1200
      ScaleHeight     =   255
      ScaleWidth      =   4815
      TabIndex        =   54
      Top             =   7920
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
         TabIndex        =   55
         Top             =   0
         Width           =   120
      End
   End
   Begin FramePlusCtl.FramePlus FrmAutoConsumption 
      Height          =   7155
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   12621
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
         TabIndex        =   15
         Top             =   2040
         Visible         =   0   'False
         Width           =   2175
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
         Left            =   1470
         TabIndex        =   22
         Top             =   4605
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
         TabIndex        =   12
         Top             =   1140
         Width           =   1455
      End
      Begin VB.TextBox TXTPCOD 
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
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1200
         Width           =   3015
      End
      Begin VB.TextBox BOXNO 
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
         MaxLength       =   10
         TabIndex        =   13
         Top             =   1680
         Width           =   1695
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
         TabIndex        =   5
         Top             =   480
         Width           =   1815
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
         TabIndex        =   7
         Top             =   780
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
         ItemData        =   "frmOpnBoxPacking.frx":0000
         Left            =   2040
         List            =   "frmOpnBoxPacking.frx":0002
         Sorted          =   -1  'True
         TabIndex        =   1
         Tag             =   "0"
         Text            =   "cmbPackaging"
         Top             =   760
         Width           =   3015
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
         TabIndex        =   11
         Top             =   1140
         Width           =   1935
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
         TabIndex        =   9
         Top             =   800
         Width           =   1455
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
         TabIndex        =   16
         Top             =   2340
         Width           =   975
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
         TabIndex        =   14
         Text            =   "S"
         Top             =   2040
         Width           =   495
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
         TabIndex        =   18
         Text            =   "0."
         Top             =   2940
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
         TabIndex        =   17
         Top             =   2640
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
         TabIndex        =   20
         Tag             =   "0"
         Top             =   3750
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
         TabIndex        =   19
         Top             =   3450
         Width           =   1335
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
         TabIndex        =   21
         Tag             =   "0"
         Top             =   4050
         Width           =   1335
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
         Left            =   1440
         MaxLength       =   149
         TabIndex        =   23
         Top             =   5040
         Width           =   2295
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
         Left            =   7440
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   31
         Tag             =   "0"
         Top             =   6360
         Width           =   855
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
         Left            =   8760
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   30
         Tag             =   "0"
         Top             =   6360
         Width           =   855
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
         Left            =   9960
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   29
         Tag             =   "0"
         Top             =   6360
         Width           =   975
      End
      Begin MSComCtl2.DTPicker TXTVBDT 
         Height          =   285
         Left            =   9840
         TabIndex        =   28
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
         Format          =   53411841
         CurrentDate     =   39347
      End
      Begin WelchButton.lvButtons_H cmdSave 
         Height          =   375
         Left            =   720
         TabIndex        =   25
         Top             =   6315
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
      Begin WelchButton.lvButtons_H cmdCancel 
         Height          =   375
         Left            =   2280
         TabIndex        =   26
         Top             =   6315
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         Caption         =   "&Clear"
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
         Enabled         =   0   'False
         cBack           =   -2147483633
      End
      Begin MSFlexGridLib.MSFlexGrid Flex 
         Height          =   4410
         Left            =   4080
         TabIndex        =   32
         Top             =   1680
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   7779
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
         TabIndex        =   24
         Top             =   5385
         Width           =   3495
         _ExtentX        =   6165
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
         Height          =   375
         Left            =   1440
         Shape           =   4  'Rounded Rectangle
         Top             =   4560
         Width           =   2295
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
         Left            =   9070
         TabIndex        =   56
         Top             =   1155
         Width           =   855
      End
      Begin VB.Label LBLPCOD 
         BackStyle       =   0  'Transparent
         Caption         =   "Party Name :"
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
         TabIndex        =   2
         Top             =   1095
         Width           =   1575
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
         TabIndex        =   53
         Tag             =   "S"
         Top             =   -2040
         Width           =   1335
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
         TabIndex        =   52
         Top             =   120
         Width           =   1695
      End
      Begin VB.Shape BORDER2 
         BorderColor     =   &H80000002&
         Height          =   300
         Left            =   5640
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   5655
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
         TabIndex        =   51
         Top             =   120
         Width           =   3615
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
         TabIndex        =   50
         Top             =   120
         Width           =   1695
      End
      Begin VB.Shape BORDER1 
         BorderColor     =   &H80000002&
         Height          =   300
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   5295
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
         TabIndex        =   49
         Top             =   120
         Width           =   3375
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "ItemName :"
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
         TabIndex        =   6
         Top             =   795
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Date   :"
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
         Left            =   9075
         TabIndex        =   48
         Top             =   480
         Width           =   735
      End
      Begin VB.Label LBLCFG 
         BackStyle       =   0  'Transparent
         Caption         =   "Grade :"
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
         Left            =   9070
         TabIndex        =   8
         Top             =   795
         Width           =   855
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Type of Packing : Fresh Packing"
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
         TabIndex        =   47
         Top             =   445
         Width           =   4575
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
         TabIndex        =   0
         Top             =   760
         Width           =   1815
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
         TabIndex        =   10
         Top             =   1125
         Width           =   1815
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Lot No.      :"
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
         TabIndex        =   4
         Top             =   480
         Width           =   1215
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H80000002&
         BorderStyle     =   5  'Dash-Dot-Dot
         DrawMode        =   1  'Blackness
         FillColor       =   &H00000080&
         Height          =   5265
         Left            =   120
         Top             =   1560
         Width           =   11295
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
         TabIndex        =   46
         Top             =   2310
         Width           =   1455
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
         TabIndex        =   45
         Top             =   1995
         Width           =   1815
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
         TabIndex        =   44
         Top             =   1680
         Width           =   1095
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
         TabIndex        =   43
         Top             =   3435
         Width           =   3135
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
         Top             =   2955
         Width           =   1575
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
         TabIndex        =   41
         Top             =   2640
         Width           =   1575
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
         TabIndex        =   40
         Top             =   4035
         Width           =   3015
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
         TabIndex        =   39
         Top             =   3720
         Width           =   3135
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
         TabIndex        =   38
         Top             =   2040
         Width           =   735
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H80000002&
         BorderStyle     =   5  'Dash-Dot-Dot
         DrawMode        =   1  'Blackness
         FillColor       =   &H00000080&
         Height          =   1065
         Left            =   120
         Top             =   3360
         Width           =   3735
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
         TabIndex        =   37
         Top             =   5040
         Width           =   1215
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
         TabIndex        =   36
         Top             =   6195
         Width           =   1050
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Click On Box No For Edit."
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
         Left            =   4080
         TabIndex        =   35
         Top             =   6240
         Width           =   2535
      End
      Begin VB.Line Line2 
         BorderStyle     =   5  'Dash-Dot-Dot
         X1              =   3960
         X2              =   3960
         Y1              =   1560
         Y2              =   6840
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
         Left            =   7320
         TabIndex        =   34
         Top             =   6195
         Width           =   1170
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
         TabIndex        =   33
         Top             =   6195
         Width           =   1065
      End
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   480
      Top             =   7920
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
End
Attribute VB_Name = "frmOpnBoxPacking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ISGRBOX As Boolean
Dim SubGradename As String
Dim IsShadeReq As Boolean
Dim LOT_MC_CHANGE_OCCUR As Boolean
Dim TWSTREQ As String
Dim ERROROCCUR As Boolean
Dim DIVCODE As String
Dim LSPKGCOD As String
Dim M_DBCD As String
Dim PKGNGCD As String
Dim LOCCOD As String
Dim MCCD As String
Dim RETURNABLE As String
Dim GRADE As String
Dim SUBGRADE As String
Dim SUBPKGCODE As String
Dim CHALLAN As String
Dim PALETNO As String
Dim FLAG As Boolean
Dim CFGTYP As String
Dim INFORS As New ADODB.Recordset
'---
Dim SAVEFLAG As Boolean
Dim ROWNO As Long
Dim SWITCH As Boolean
Dim SQL As String
Dim COUNTER As Long
Dim M_PCOD As String
Dim TABLE As String
Dim FIELD As String
Dim TTYP As String
Dim FINITMCOD As String
Dim REQNOCOPS As Boolean
Dim REQCOPSWGT As Boolean
Dim REQBOXWGT As Boolean
Dim REQPALLET As Boolean

Private Sub BOXNO_Change()
 If Len(Trim(BOXNO)) = 10 Then   '1
   Call BOXNO_KeyPress(13)
 End If
End Sub

Private Sub BOXNO_KeyPress(KeyAscii As Integer)
ISGRBOX = False
Select Case KeyAscii
Case 48 To 58
Case 97 To 122
     KeyAscii = KeyAscii - 32
Case 65 To 90
Case 8
Case 13
Case Else
     KeyAscii = 0
     Exit Sub
End Select

If KeyAscii = 32 Or KeyAscii = 95 Then KeyAscii = 0: Exit Sub

If Len(BOXNO) = 10 And KeyAscii = 13 Then   '1

'CHECK BOX IN ANOTHER PACKING STATION
Dim STATION As String
STATION = ExistInAnotherPackingStation
If STATION <> Empty Then
   MsgBox "BOX NO. " & BOXNO & " EXIST IN PACKING STATION : " & STATION
   BOXNO = Empty
   If BOXNO.Enabled Then BOXNO.SetFocus
   Exit Sub
End If
'----------------------------------------

If IsDispatchExist(DIVCODE, LSPKGCOD, BOXNO) Then
   MsgBox "Boxno. " & BOXNO & " has been Dispatched."
   BOXNO = Empty
   If BOXNO.Enabled Then BOXNO.SetFocus
   Exit Sub
End If
     
Dim PKTYPRS As ADODB.Recordset
Set PKTYPRS = New ADODB.Recordset
Dim RSDATA As ADODB.Recordset
Set RSDATA = New ADODB.Recordset

If RSDATA.State = 1 Then RSDATA.Close
RSDATA.Open "SELECT * FROM BOXREGISTER WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
"' AND DVCD='" & DIVCODE & "' AND (VTYP='PPF' OR VTYP='OPN') AND VBNO='" & BOXNO & _
"' AND PKG_STCOD='" & LSPKGCOD & "' AND RECSTAT<>'D'", CN, adOpenDynamic, adLockOptimistic
     
If Not RSDATA.EOF Then '2
 M_DBCD = Trim(RSDATA!dbcd & "")
 
 'in case of wastage production
 If M_DBCD = "000006" Then
    MsgBox "This No. " & BOXNO & " Alloted to Wastage Production Module", vbCritical, "Can't Edit Wastage Entry from Box Packing"
    If BOXNO.Enabled Then BOXNO.SetFocus: BOXNO = Empty
    Exit Sub
 End If
 '================================
 
 TTYP = Trim(RSDATA!VTYP & "")
 SWITCH = False
 TXTVBDT.Enabled = True
 FLAG = True
 SAVEFLAG = False
 cmbPackaging.Text = GetCode("PKGNGMST", Trim(RSDATA!PKGNG_COD & ""), "CODE", "NAME")
 ISGRBOX = IIf(Trim(RSDATA!PCOD & "") = "GRPACK", True, False)
 txtLTNo = Trim(RSDATA!LOTNO & "")
 txtLTNo.Tag = Trim(RSDATA!LOTNO & "")
 TXTDENI = FindFinItemCode("CODE", Trim(RSDATA!ICOD & ""))
 TXTGRAD = GetCode("GRDMST", Trim(RSDATA!grad & ""), "CODE", "GRAD")
 TXTLOC = GetCode("LOCMST", Trim(RSDATA!LOCCOD & ""), "CODE", "NAME")
 TXTPCOD = GetCode("ACCMST", Trim(RSDATA!PCOD & ""), "CODE", "NAME")
 TXTMCCD = GetMachineName(DIVCODE, Trim(RSDATA!MCCD & ""))
 TXTMCCD.Tag = TXTMCCD
 LOT_MC_CHANGE_OCCUR = False
 TXTVBDT = Format(RSDATA!VBDT, "DD/MM/YYYY")
  
 If Trim(TXTPCOD) <> Empty Then    'PARTY REQUIRED in CASE OF JOB
    TXTPCOD.Enabled = True: LBLPCOD.Enabled = True
    TABLE = "JOBIN"
 Else
    TXTPCOD.Enabled = False: LBLPCOD.Enabled = False
    TABLE = "STORETRAN"
 End If
     
'RETURNABLE COPS : BHAIJI
'-------------------
 
If Trim(RSDATA!SUBGRD) = "S" Or Trim(RSDATA!SUBGRD) = "Z" Or Trim(RSDATA!SUBGRD) = "O" Then
   TXTTWIST = Trim(RSDATA!SUBGRD)
Else
   Call SetShadeName(Trim(RSDATA!SUBGRD))
End If
        
    TXTCOP = Trim(RSDATA!COPS)
    TXTCOP.Tag = Trim(RSDATA!COPS)
    TXTCTWT = Trim(RSDATA!BOXWGT)
    TXTCPWT = Trim(RSDATA!COPSWGT)
    TXTGRWT = Trim(RSDATA!GRSWGT)
    TXTTRWT = Trim(RSDATA!TRWGT)
    TXTNTWT = Trim(RSDATA!NTWGT)
    TXTNTWT.Tag = Trim(RSDATA!NTWGT)
    TXTRMRK = Trim(RSDATA!RMRK & "")
    CHALLAN = Trim(RSDATA!chln & "")
    
   Dim i As Double, J As Double
   i = 0
   For i = 1 To FLEXPLY.Cols - 1
      J = 0
      For J = 0 To RSDATA.Fields.COUNT - 1
        If Trim(RSDATA.Fields(J).NAME) = Trim(FLEXPLY.TextMatrix(0, i)) Then
            FLEXPLY.TextMatrix(1, i) = Val(RSDATA.Fields(J).Value)
        End If
      Next
   Next
   
    SWITCH = True   'PROBLEM IF EXISTING BOX EDIT AND OPENING
    If TTYP <> "OPN" Then TXTVBDT.Enabled = False
Else
FLAG = False
End If ' 2
SendKeys "{TAB}"
End If ' 1
End Sub

Private Sub cmbPackaging_Click()
  Call SETPLYLIMIT
End Sub

Private Sub cmbPackaging_KeyDown(KeyCode As Integer, Shift As Integer)
  Call SETPLYLIMIT
  KeyCode = 0
End Sub

Private Sub cmdSave_Click()
ERROROCCUR = False
On Error GoTo LAST
Dim i As Long, J As Long
Dim RSTMP As ADODB.Recordset
Set RSTMP = New ADODB.Recordset

Dim TABLENAME As String

If BOXNO = "XXXXXXXXXX" Then Exit Sub

If Len(Trim(BOXNO)) <> 10 Then
  MsgBox "Length of Box Should Be 10 digit "
  If BOXNO.Enabled Then BOXNO.SetFocus
  Exit Sub
End If

If IsDispatchExist(DIVCODE, LSPKGCOD, BOXNO) Then
   MsgBox "Boxno. " & BOXNO & " has been Dispatched."
   BOXNO = Empty
   If BOXNO.Enabled Then BOXNO.SetFocus
   Exit Sub
End If

If CheckData(ROWNO) Then Exit Sub

Call SetGlobal

CN.BeginTrans

If Not isAllowGRPacking Then
   CN.RollbackTrans
   Exit Sub
End If

If SAVEFLAG Then

 Dim TEMPRS As ADODB.Recordset
 Set TEMPRS = New ADODB.Recordset
 If TEMPRS.State = 1 Then TEMPRS.Close
 TEMPRS.Open "SELECT * FROM BOXREGISTER WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
 "' AND PKG_STCOD = '" & LSPKGCOD & "' AND VTYP='OPN' AND VBNO = '" & BOXNO & "'AND RECSTAT='A'", CN, adOpenDynamic, adLockOptimistic
 If Not TEMPRS.EOF Then
    MsgBox "Box Already Exist !!", vbInformation, "Box Missing !!"
    BOXNO.SetFocus
    CN.RollbackTrans
    Exit Sub
 End If

COUNTER = COUNTER + 1

If IsBoxExistInUnit(Trim(BOXNO)) Then
   MsgBox "BoxNo. " & BOXNO & " Already Exist."
   CN.RollbackTrans
   Exit Sub
End If

SQL = "INSERT INTO BOXREGISTER(COMP,UNIT,DVCD,DBCD,VTYP,VBNO,PLTNO,VBDT,CHLN,PKG_STCOD,PKGNG_COD,"
SQL = SQL & "LOCCOD,MCCD,ISRETURNABLE,LOTNO,ICOD,GRAD,SUBGRD,COPS,BOXWGT,COPSWGT,GRSWGT,TRWGT,"
SQL = SQL & "NTWGT,PACKER,RMRK,RECSTAT,PVTYP)VALUES('" & compPth & _
"','" & UNCD & "','" & DIVCODE & "','" & M_DBCD & "','OPN','" & BOXNO & "','" & PALETNO & _
"','" & Format(TXTVBDT, "MM/DD/YYYY") & "','" & CHALLAN & _
"','" & LSPKGCOD & "','" & PKGNGCD & "','" & LOCCOD & "','" & MCCD & "','" & RETURNABLE & "','" & txtLTNo & _
"','" & FINITMCOD & "','" & GRADE & "','" & FindSubGradeCode & "','" & Val(TXTCOP) & _
"','" & Val(TXTCTWT) & "','" & Val(TXTCPWT) & "','" & Val(TXTGRWT) & "','" & Val(TXTTRWT) & _
"','" & Val(TXTNTWT) & "','" & cUName & "','" & TXTRMRK & "','A','OPN')"
CN.Execute SQL

Call FillDetail
If ERROROCCUR Then Exit Sub

Call FindTotal

Else

If FLAG = True Then
   If TTYP <> "OPN" Then
       Call SetRawMaterial
       'Call SetGRPacking
       If ERROROCCUR Then Exit Sub
   End If
   If TTYP = "OPN" Then
      Call FillDetail
      If ERROROCCUR Then Exit Sub
   End If
End If

SQL = "UPDATE BOXREGISTER SET MCCD='" & MCCD & "',PKGNG_COD='" & PKGNGCD & "',LOCCOD='" & LOCCOD & _
"',ISRETURNABLE='" & RETURNABLE & "', LOTNO='" & txtLTNo & "',ICOD='" & FINITMCOD & _
"',GRAD='" & GRADE & "',SUBGRD='" & FindSubGradeCode & "',COPS='" & Val(TXTCOP) & _
"',BOXWGT='" & Val(TXTCTWT) & "',COPSWGT='" & Val(TXTCPWT) & "',GRSWGT='" & Val(TXTGRWT) & _
"',TRWGT='" & Val(TXTTRWT) & "',NTWGT='" & Val(TXTNTWT) & "',RMRK='" & TXTRMRK & _
"',VBDT='" & Format(TXTVBDT, "MM/DD/YYYY") & "',CHLN = '" & CHALLAN & _
"',PCOD = '" & M_PCOD & "',PACKER = '" & cUName & "' WHERE COMP='" & compPth & _
"' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCODE & "' AND VBNO='" & BOXNO & _
"' AND PKG_STCOD='" & LSPKGCOD & "' AND RECSTAT<>'D' AND (VTYP='PPF' OR VTYP='OPN')"

CN.Execute SQL

If FLAG = True Then
   If TTYP <> "OPN" Then
      'Call SetGRPacking
      If ERROROCCUR Then Exit Sub
   End If
End If

Call FindTotal

End If
TXTNTWT.Tag = 0
TXTCOP.Tag = 0

'PLY UPDATION COMMON FOR BOTH SAVE AND EDIT
If RSTMP.State = 1 Then RSTMP.Close
RSTMP.Open "SELECT * FROM BOXREGISTER WHERE COMP='" & compPth & _
"' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCODE & "' AND VBNO='" & BOXNO & _
"' AND PKG_STCOD='" & LSPKGCOD & "' AND RECSTAT<>'D' AND (VTYP='PPF' OR VTYP='OPN')", CN, adOpenDynamic, adLockOptimistic

If Not RSTMP.EOF Then
    If FLEXPLY.Enabled = True Then
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
        RSTMP.Update
      Next
    End If
RSTMP.Close
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

Call CLEARDATA

TTYP = Empty
SAVEFLAG = True
txtLTNo.Tag = txtLTNo
TXTMCCD.Tag = TXTMCCD
LOT_MC_CHANGE_OCCUR = False

FLAG = False
SWITCH = False
TXTVBDT.Enabled = True
TXTPCOD = Empty: TXTPCOD.Enabled = False: LBLPCOD.Enabled = False
BOXNO.SetFocus
CN.CommitTrans

Exit Sub
LAST:
MsgBox ERR.Description
CN.RollbackTrans
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
    TXTTWIST = Flex.TextMatrix(ROWNO, 1)
    TXTCOP = Flex.TextMatrix(ROWNO, 2)
    TXTCOP.Tag = Flex.TextMatrix(ROWNO, 2)
    TXTCTWT = Flex.TextMatrix(ROWNO, 3)
    
    If IsShadeReq Then
       TXTSHADE = Flex.TextMatrix(ROWNO, 1)
    ElseIf TXTTWIST.Enabled = True Then
       TXTTWIST = Flex.TextMatrix(ROWNO, 1)
    End If
    
    If Val(Flex.TextMatrix(ROWNO, 4)) < 1 Then
       TXTCPWT = "0" & Flex.TextMatrix(ROWNO, 4)
    Else
       TXTCPWT = Flex.TextMatrix(ROWNO, 4)
    End If
    
    TXTGRWT = Flex.TextMatrix(ROWNO, 5)
    TXTTRWT = Flex.TextMatrix(ROWNO, 6)
    TXTNTWT = Flex.TextMatrix(ROWNO, 7)
    TXTNTWT.Tag = Flex.TextMatrix(ROWNO, 7)
    TXTRMRK = Flex.TextMatrix(ROWNO, 8)
    
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
     End If
  End If

End Sub

Private Sub Form_Activate()
  
  If DIVCODE = Empty Or Trim(LBLDESC1.Caption) = "XXXXXXXXXX" Then
     MsgBox "Select Division For Packing."
     Unload Me
     Exit Sub
  End If
  
  If LSPKGCOD = Empty Or Trim(LBLDESC2.Caption) = "XXXXXXXXXX" Then
     MsgBox "Select Packing Station For Packing."
     Unload Me
     Exit Sub
  End If
    
  'For Raw Material Consumption Slip
  If PALETNO = Empty Then
     Unload Me
  End If
    
  If cmbPackaging.ListCount > 0 Then cmbPackaging.ListIndex = 0
  
  
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If TypeOf ActiveControl Is TextBox And UCase(ActiveControl.NAME) <> "TXTRMRK" Then If ActiveControl.Text = Empty Then Exit Sub
  If UCase(ActiveControl.NAME) = "FLEXPLY" Then Exit Sub
  If UCase(ActiveControl.NAME) = "TXTGRWT" Then Exit Sub
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
  Call CenterChild(frm_Main, Me): Call ColorComponent(Me)
  
  SWITCH = False
  
  ERROROCCUR = False
  
  Me.Left = 50: Me.KeyPreview = True
  M_DBCD = "000003"
  SAVEFLAG = True
'-------DIVISION NAME
  M_DESC = Empty: Key = Empty:  NEW_VISIBLE = False:  DIVCODE = Empty
  LBLDESC1.Caption = Empty
  If DIVCODE = Empty Then
    LBLDESC1 = SearchList1("SELECT  TOP 20 CODE,NAME FROM DIVMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND RECSTAT='A' AND CODE<>'000001'", 0, "", "SELECT DIVISION MASTER")
    If LBLDESC1 <> Empty Then DIVCODE = Key Else LBLDESC1 = "???????": Unload Me
  End If
  
  LBLCFG.Caption = LabelDisplay(DIVCODE & "", UNCD)
  
  If IsTwistReq(DIVCODE) = "Y" Then
    TWSTREQ = "Y"
    LBLTWST.Enabled = True: LBLSZO.Enabled = True: TXTTWIST.Enabled = True
  End If
  
  '{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{
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
  '}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}

 '------------------------------------------------------------------------
 'SUB PACKAGING TYPE
 If INFORS.State = 1 Then INFORS.Close
    INFORS.Open "SELECT * FROM DIVMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND CODE = '" & DIVCODE & "' AND RECSTAT='A' ", CN, adOpenDynamic, adLockOptimistic
 If Not INFORS.EOF Then
    CFGTYP = INFORS!CFGTYP & ""
 End If
'------------------------------------------------------------------------

'-------PACKING STATION MASTER
M_DESC = Empty:  Key = Empty:  NEW_VISIBLE = False: LSPKGCOD = Empty
LBLDESC2 = SearchList1("SELECT TOP 20 CODE,NAME FROM PCKMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "'", 0, "", "SELECT PACKING STATION FROM MASTER LIST")
If Key = Empty Then Exit Sub
LSPKGCOD = Key

If RS.State = 1 Then RS.Close
RS.Open "SELECT * FROM PCKMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND CODE ='" & LSPKGCOD & "'", CN, adOpenDynamic, adLockOptimistic
If Not RS.EOF Then
   REQNOCOPS = IIf(Trim(RS!REQNOCOPS) = "Y", True, False)
   If Not REQNOCOPS Then
      LBLNOCOPS.Enabled = False
      TXTCOP.Enabled = False
      LBLBOXWGT.Enabled = False
      TXTCTWT.Enabled = False
      LBLCOPSWGT.Enabled = False
      TXTCPWT.Enabled = False
   End If
   
   REQCOPSWGT = IIf(Trim(RS!REQCOPSWGT) = "Y", True, False)
   REQBOXWGT = IIf(Trim(RS!REQBOXWGT) = "Y", True, False)
   REQPALLET = IIf(Trim(RS!REQPALLET) = "Y", True, False)
End If
'---------------------------

COUNTER = 0
       
TXTVBDT.Value = Now

'FOR PALLET NO.
PALETNO = GenPackSlipNo(LSPKGCOD, "LPNO")

Call SetPackingType
Call setHeading
End Sub

Private Sub GRPLAT_GotFocus()
 If Not REQPALLET Then
     GRPLAT.Enabled = False
     Exit Sub
  End If
End Sub

Private Sub TXTMCCD_GotFocus()
  TXTMCCD.BackColor = RGB(BRED, BGREEN, BBLUE)
  SendKeys "{HOME}+{END}"
  ToolTip Me, "Press {F2} / {Enter} For Machine Master Help", "", TXTMCCD.Left - 620, TXTMCCD.Top + TXTMCCD.Height + 100
End Sub

Private Sub TXTMCCD_KeyDown(KeyCode As Integer, Shift As Integer)
Me.KeyPreview = False
   If KeyCode = vbKeyF2 Or Trim(TXTMCCD.Text) = Empty Then
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

Private Sub TXTMCCD_LostFocus(): TXTMCCD.BackColor = vbWhite: picToolTip.Visible = False: End Sub

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

Private Sub TXTGRAD_KeyDown(KeyCode As Integer, Shift As Integer)
  If Trim(TXTGRAD.Text) = Empty Or KeyCode = vbKeyF2 Then
    TXTGRAD.Text = SearchList1("select TOP 20 grad as grade,grad from grdmst", 0, TXTGRAD, "SELECT " & LBLCFG.Caption)
  End If
End Sub

Private Sub TXTGRWT_Change()
TXTNTWT = Val(TXTGRWT) - Val((Val(TXTCOP) * Val(TXTCPWT)) + Val(TXTCTWT))
TXTNTWT = nstr(TXTNTWT, 9, 3)
TXTNTWT = Trim(TXTNTWT)
End Sub

Private Sub TXTLOC_KeyDown(KeyCode As Integer, Shift As Integer)
  If (KeyCode = vbKeyReturn And Trim(TXTLOC.Text) = Empty) Or KeyCode = vbKeyF2 Then
    TXTLOC.Text = SearchList1("SELECT  TOP 20 CODE,NAME FROM LOCMST", 0, TXTLOC, "SELECT LOCATION FROM MASTER")
  End If
End Sub

Private Sub txtltno_GotFocus()
  txtLTNo.BackColor = RGB(BRED, BGREEN, BBLUE)
  SendKeys "{HOME}+{END}"
  ToolTip Me, "Press {F2} / {Enter} For Lot Master Help", "", txtLTNo.Left - 50, txtLTNo.Top + txtLTNo.Height + 100
End Sub

Private Sub txtDENI_GotFocus()
  TXTDENI.BackColor = RGB(BRED, BGREEN, BBLUE)
  SendKeys "{HOME}+{END}"
  ToolTip Me, "Finish Item of Lot : " & txtLTNo, "", TXTDENI.Left - 120, TXTDENI.Top + TXTDENI.Height + 100
End Sub

Private Sub TXTGRAD_GotFocus()
  TXTGRAD.BackColor = RGB(BRED, BGREEN, BBLUE)
  SendKeys "{HOME}+{END}"
  ToolTip Me, "Press {F2} / {Enter} For Grade Master Help", "", TXTGRAD.Left - 3820, TXTGRAD.Top + TXTGRAD.Height + 100
End Sub

Private Sub TXTLOC_GotFocus()
  TXTLOC.BackColor = RGB(BRED, BGREEN, BBLUE)
  SendKeys "{HOME}+{END}"
  ToolTip Me, "Press {F2} / {Enter} For Location Master Help", "", TXTLOC.Left - 620, TXTLOC.Top + TXTLOC.Height + 100
End Sub

Private Sub txtltno_KeyDown(KeyCode As Integer, Shift As Integer)
Dim SQL As String: Me.KeyPreview = False
If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
     txtLTNo = Empty
ElseIf KeyCode = vbKeyF2 Or (KeyCode = vbKeyReturn And txtLTNo = Empty) Then
   M_DESC = Empty:   NEW_VISIBLE = False
   SQL = "SELECT DISTINCT LTNO,LTNO FROM TXULOT WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCODE & "'"
   Key = Empty
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
  If FLEXPLY.Cols > 1 Then
    FLEXPLY.ROW = 1
    FLEXPLY.COL = 1
  End If
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

Private Sub txtGRWT_GotFocus(): TXTGRWT.BackColor = RGB(BRED, BGREEN, BBLUE): SendKeys "{HOME}+{END}": End Sub
Private Sub TXTRMRK_GotFocus(): TXTRMRK.BackColor = RGB(BRED, BGREEN, BBLUE): SendKeys "{HOME}+{END}": End Sub
Private Sub txtPCOD_LostFocus(): TXTPCOD.BackColor = vbWhite: picToolTip.Visible = False: End Sub

Private Sub TXTVBDT_GotFocus()
 'ToolTip Me, "Date Can't Be Modified.  ", "", TXTVBDT.Left - 50, TXTVBDT.Top + TXTVBDT.Height + 100
End Sub

Private Sub TXTVBDT_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub TXTVBDT_KeyPress(KeyAscii As Integer)
 KeyAscii = 0
End Sub

Private Sub TXTVBDT_LostFocus()
picToolTip.Visible = False
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
Private Sub TXTLOC_LostFocus(): TXTLOC.BackColor = vbWhite: picToolTip.Visible = False: End Sub

Private Sub TXTCOP_Change():  Call TXTCPWT_Change: End Sub
Private Sub TXTCTWT_Change(): Call TXTCPWT_Change: End Sub

Private Sub cmbPackaging_KeyPress(KeyAscii As Integer): KeyAscii = 0: End Sub

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
   BORDER2.BorderColor = RGB(LBLRED, LBLGREEN, LBLBLUE): LBLDESC1.ForeColor = RGB(LBLRED, LBLGREEN, LBLBLUE): LBLDESC2.ForeColor = RGB(LBLRED, LBLGREEN, LBLBLUE)
ElseIf ctr Mod 75 = 0 And ctr <= 75 Then
   LBLHEADING1.ForeColor = vbRed: LBLHEADING2.ForeColor = vbRed: BORDER1.BorderColor = vbRed: BORDER2.BorderColor = vbRed
   LBLDESC1.ForeColor = vbRed: LBLDESC2.ForeColor = vbRed
ElseIf ctr Mod 105 = 0 And ctr <= 105 Then
   LBLHEADING1.ForeColor = vbBlue: LBLHEADING2.ForeColor = vbBlue: BORDER1.BorderColor = vbBlue: BORDER2.BorderColor = vbBlue
   LBLDESC1.ForeColor = vbBlue: LBLDESC2.ForeColor = vbBlue
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
  RSITM.Open "SELECT NAME,ISRETURNABLE FROM FINITMMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCODE & "' AND CODE='" & FICD & "'", CN, adOpenDynamic, adLockOptimistic
  If Not RSITM.EOF Then
     TXTDENI = RSITM!NAME
     RETURNABLE = Trim(RSITM!ISRETURNABLE & "")
  Else
     RETURNABLE = "N"
     TXTDENI = Empty
  End If
  
  RSITM.Close
End If
End Sub

Private Sub SetPackingType()
Dim PKTYPRS As ADODB.Recordset
Set PKTYPRS = New ADODB.Recordset
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
 
 J = 8
 For i = 1 To FLEXPLY.Cols - 1
    J = J + 1
    Flex.TextMatrix(0, J) = FLEXPLY.TextMatrix(0, i)
    .ColWidth(J) = 0
 Next
    
 .ColWidth(0) = 1250: .ColWidth(1) = 350: .ColWidth(2) = 600: .ColWidth(3) = 950
 .ColWidth(4) = 950: .ColWidth(5) = 1000: .ColWidth(6) = 950: .ColWidth(7) = 950: .ColWidth(8) = 0
End With

 If REQNOCOPS = False And REQCOPSWGT = False And REQBOXWGT = False Then
    LBLBOXNO.Caption = "Roll No."
    Flex.TextMatrix(0, 0) = "Roll No."
 End If
 
 If IsTwistReq(DIVCODE) <> "Y" Then Flex.ColWidth(1) = 0
 If REQNOCOPS = False Then Flex.ColWidth(2) = 0
 If REQBOXWGT = False Then Flex.ColWidth(3) = 0
 If REQCOPSWGT = False Then Flex.ColWidth(4) = 0

End Sub

Private Sub SetGlobal()
Dim DBCDRS As ADODB.Recordset
Set DBCDRS = New ADODB.Recordset

If DBCDRS.State = 1 Then DBCDRS.Close
DBCDRS.Open "SELECT CODE FROM PKGNGMST WHERE NAME='" & cmbPackaging.Text & "' AND RECSTAT='A'", CN, adOpenDynamic, adLockOptimistic
If Not DBCDRS.EOF Then
   PKGNGCD = Trim(DBCDRS!CODE & "")
Else
   PKGNGCD = Empty
End If
DBCDRS.Close

If DBCDRS.State = 1 Then DBCDRS.Close
DBCDRS.Open "SELECT CODE FROM ACCMST WHERE NAME='" & TXTPCOD.Text & "'", CN, adOpenDynamic, adLockOptimistic
If Not DBCDRS.EOF Then
   M_PCOD = Trim(DBCDRS!CODE & "")
Else
   M_PCOD = Empty
End If
DBCDRS.Close

If DBCDRS.State = 1 Then DBCDRS.Close
DBCDRS.Open "SELECT CODE FROM LOCMST WHERE NAME='" & TXTLOC.Text & "'", CN, adOpenDynamic, adLockOptimistic
If Not DBCDRS.EOF Then
   LOCCOD = Trim(DBCDRS!CODE & "")
Else
   LOCCOD = Empty
End If
DBCDRS.Close

If DBCDRS.State = 1 Then DBCDRS.Close
DBCDRS.Open "SELECT CODE FROM MACMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
            "' AND DVCD='" & DIVCODE & "' AND NAME='" & TXTMCCD.Text & "'", CN, adOpenDynamic, adLockOptimistic
If Not DBCDRS.EOF Then
   MCCD = Trim(DBCDRS!CODE & "")
Else
   MCCD = Empty
End If
DBCDRS.Close

FINITMCOD = FindFinItemCode

GRADE = GetCode("GRDMST", TXTGRAD, "GRAD", "CODE")

End Sub

Private Function FindFinItemCode(Optional FIELD As String = "NAME", Optional CODE As String) As String
Dim GRRS As ADODB.Recordset
Set GRRS = New ADODB.Recordset
Dim QUERY As String
QUERY = "SELECT * FROM FINITMMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
"' AND DVCD='" & DIVCODE & "' "
If FIELD = "NAME" Then
  QUERY = QUERY & "AND NAME ='" & TXTDENI & "'"
ElseIf FIELD = "CODE" Then
  QUERY = QUERY & "AND CODE ='" & CODE & "'"
End If

If GRRS.State = 1 Then GRRS.Close
GRRS.Open QUERY, CN, adOpenDynamic, adLockOptimistic
If Not GRRS.EOF Then
  If FIELD = "NAME" Then
     FindFinItemCode = GRRS!CODE
  ElseIf FIELD = "CODE" Then
     FindFinItemCode = GRRS!NAME
   End If
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
On Error GoTo FLEXERR
 Dim i As Long, J As Long
 Dim INDEX As Long
 
 If ROWNO = 0 And Flex.Rows = 2 Then
    If Flex.TextMatrix(1, 0) = Empty Then
       ROWNO = 1
    End If
 End If
 
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
    
    'Flex.TextMatrix(ROWNO, 1) = Trim(TXTTWIST)
    
    Flex.TextMatrix(ROWNO, 2) = Trim(TXTCOP)
    Flex.TextMatrix(ROWNO, 3) = Trim(nstr(Val(TXTCTWT), 12, 3))
    Flex.TextMatrix(ROWNO, 4) = Trim(nstr(Val(TXTCPWT), 12, 3))
    Flex.TextMatrix(ROWNO, 5) = Trim(nstr(Val(TXTGRWT), 12, 3))
    Flex.TextMatrix(ROWNO, 6) = Trim(nstr(Val(TXTTRWT), 12, 3))
    Flex.TextMatrix(ROWNO, 7) = Trim(nstr(Val(TXTNTWT), 12, 3))
    Flex.TextMatrix(ROWNO, 8) = TXTRMRK
            
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
    End If
               
    'REMOVE BELOW COMMENT BLOCK WHEN ITEMS PROCESS ARE GOING TO MULTIPLE
    
    SWITCH = False
    TXTVBDT.Enabled = True
    
 Exit Sub
FLEXERR:
MsgBox "Invalid Data on Flex"
ERROROCCUR = True
End Sub

Private Function CheckData(RNO As Long) As Boolean
On Error GoTo CHECKERR

Dim INDEX As Long
 If Val(TXTCOP) <= 0 And TXTCOP.Enabled Then MsgBox "Please Enter No.of Cops !!", vbInformation: TXTCOP.SetFocus: CheckData = True: Exit Function
 If Val(TXTCTWT) <= 0 And TXTCTWT.Enabled Then MsgBox "Please Enter Carton Weight !!", vbInformation, "Weight Missing !!": TXTCTWT.SetFocus: CheckData = True: Exit Function
 If Val(TXTGRWT) <= 0 And TXTGRWT.Enabled Then MsgBox "Please Enter Gross Weight !!", vbInformation, "Weight Missing !!": TXTGRWT.SetFocus: CheckData = True: Exit Function
 If Val(TXTNTWT) <= 0 Then MsgBox "Net Weight Can't be Negative or Zero!!", vbInformation, "Weight Missing !!": TXTGRWT.SetFocus: CheckData = True: Exit Function
 If TXTDENI = Empty Then MsgBox "Please Select Proper Lot !!", vbInformation, "Item Missing !!": txtLTNo.SetFocus: CheckData = True: Exit Function
 If txtLTNo = Empty Then MsgBox "Please Select Proper Lot !!", vbInformation, "Lot Missing !!": txtLTNo.SetFocus: CheckData = True: Exit Function
 If TXTMCCD = Empty Then MsgBox "Please Select Machine !!", vbInformation, "Machine Missing !!": TXTMCCD.SetFocus: CheckData = True: Exit Function
 If TXTGRAD = Empty Then MsgBox "Please Select Grade !!", vbInformation, "Grade Missing !!": TXTGRAD.SetFocus: CheckData = True: Exit Function
 
 
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

Exit Function
CHECKERR:
MsgBox "Invalid Checking"
ERROROCCUR = True
End Function

Private Sub CLEARDATA()
 Dim i As Long, J As Long
 TXTGRWT = Empty: TXTNTWT = Empty: TXTSHADE = Empty
 BOXNO = Empty
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
On Error GoTo LAST
  Dim AI As String
  Dim BQ As Double
Dim COUNT As Long: COUNT = 0
Dim ITMCODE As String
Dim TOTALQTY As Double, ITMQTY As Double, ITMRATE As Double, ITMAMT As Double, ITMEDITQTY As Double
Dim TEMPRS As ADODB.Recordset
Set TEMPRS = New ADODB.Recordset
Dim pCode As String, FIELD As String

If LOT_MC_CHANGE_OCCUR Then
   Call SetRaw
   Exit Sub
End If

If TABLE = "JOBIN" Then    'PARTY REQUIRED in CASE OF JOB
   pCode = GetCode("ACCMST", TXTPCOD, "NAME", "CODE")
   FIELD = "JOBQ"
Else
   TABLE = "STORETRAN"
   pCode = MCCD
   FIELD = "BALQ"
End If

If TEMPRS.State = 1 Then TEMPRS.Close
TEMPRS.Open "SELECT RICD,PERC FROM TXULOT WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCODE & "' AND LTNO='" & txtLTNo & "'", CN, adOpenDynamic, adLockOptimistic
Do While Not TEMPRS.EOF
 ITMCODE = Trim(TEMPRS!RICD & "")
 ITMQTY = Val(TXTNTWT)
 ITMQTY = Val((Val(TEMPRS!PERC) * ITMQTY) / 100)
 ITMEDITQTY = Val((Val(TEMPRS!PERC) * Val(TXTNTWT.Tag)) / 100)
 
 CN.Execute "UPDATE " & TABLE & " SET QNTY = QNTY  + " & ITMQTY & " - " & ITMEDITQTY & _
 ",AMNT =((QNTY + " & ITMQTY & ") * RATE),PCOD='" & pCode & _
 "' WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND VTYP='PPF' AND DBCD='" & M_DBCD & _
 "' AND VBNO='" & CHALLAN & "' AND ICOD='" & ITMCODE & "'"
  
  
TEMPRS.MoveNext
Loop
 NETWGT = Val(NETWGT) + Val(TXTNTWT) - Val(TXTNTWT.Tag)
 NETWGT = nstr(NETWGT, 9, 3): NETWGT = Trim(NETWGT)
 NETCOPS = Val(NETCOPS) + Val(TXTCOP) - Val(TXTCOP.Tag)
 TXTNTWT.Tag = "0"
 TXTCOP.Tag = "0"
 M_DBCD = "000003"
Exit Sub

LAST:
MsgBox ERR.Description
ERROROCCUR = True
CN.RollbackTrans
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
  Resume
  FLEXPLY.SetFocus
  Exit Sub
End Sub

Private Sub FLEXPLY_LeaveCell()
FLEXPLY.CellBackColor = vbWhite
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


Private Sub FindTotal()
NETWGT = 0
NETCOPS = 0
Dim i As Long

For i = 1 To Flex.Rows - 1
    NETCOPS = Val(NETCOPS) + Val(Flex.TextMatrix(i, 3))
    NETWGT = Val(NETWGT) + Val(Flex.TextMatrix(i, 1))
Next i

End Sub

'FIND IN {YES OR NO} : IF YES THEN GET NAME
Private Function ExistInAnotherPackingStation() As String
'default
ExistInAnotherPackingStation = Empty
'-----------------------------------
Dim CHKRS As ADODB.Recordset
Set CHKRS = New ADODB.Recordset
Dim SQL As String

  'CODE TO CHECK SALE BILL EXIST
  SQL = "SELECT TOP 1 PKG_STCOD FROM BOXREGISTER WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
        "' AND DVCD='" & DIVCODE & "' AND (VTYP='PPF' OR VTYP='OPN') AND VBNO='" & BOXNO & _
        "' AND PKG_STCOD <>'" & LSPKGCOD & "' AND RECSTAT<>'D'"
   
  If CHKRS.State = 1 Then CHKRS.Close
  CHKRS.Open SQL, CN, adOpenDynamic, adLockOptimistic
  If Not CHKRS.EOF Then
     ExistInAnotherPackingStation = GetPackingStation(Trim(CHKRS!PKG_STCOD)) 'FIND NAME
  End If
  '---------------------------------
End Function

Private Function GetPackingStation(PCKCOD) As String
'default
GetPackingStation = Empty
Dim GETRS As ADODB.Recordset
Set GETRS = New ADODB.Recordset
Dim SQL As String

SQL = "SELECT NAME FROM PCKMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
      "' AND CODE='" & PCKCOD & "'"
      
  If GETRS.State = 1 Then GETRS.Close
  GETRS.Open SQL, CN, adOpenDynamic, adLockOptimistic
  If Not GETRS.EOF Then
     GetPackingStation = Trim(GETRS!NAME)
  End If
  
End Function


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

Private Sub SetRaw()
On Error GoTo LAST
Dim ITMCODE As String
Dim TOTALQTY As Double, ITMQTY As Double, ITMRATE As Double, ITMAMT As Double, ITMEDITQTY As Double
Dim TEMPRS As ADODB.Recordset
Set TEMPRS = New ADODB.Recordset
Dim pCode As String, FIELD As String

If TABLE = "JOBIN" Then    'PARTY REQUIRED in CASE OF JOB
   pCode = GetCode("ACCMST", TXTPCOD, "NAME", "CODE")
   FIELD = "JOBQ"
Else
   TABLE = "STORETRAN"
   pCode = MCCD
   FIELD = "BALQ"
End If

If TEMPRS.State = 1 Then TEMPRS.Close
TEMPRS.Open "SELECT RICD,PERC FROM TXULOT WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCODE & "' AND LTNO='" & txtLTNo.Tag & "'", CN, adOpenDynamic, adLockOptimistic
Do While Not TEMPRS.EOF
 ITMCODE = Trim(TEMPRS!RICD & "")
 ITMQTY = Val(TXTNTWT)
 ITMQTY = Val(nstr((Val(TEMPRS!PERC) * ITMQTY) / 100, 9, 3))
 ITMEDITQTY = Val(nstr((Val(TEMPRS!PERC) * Val(TXTNTWT.Tag)) / 100, 9, 3))
 
 CN.Execute "UPDATE " & TABLE & " SET QNTY = QNTY  - " & ITMEDITQTY & _
 ",PCOD='" & GetMachineCode(DIVCODE, TXTMCCD.Tag) & "',LTNO='" & txtLTNo.Tag & "' WHERE COMP='" & compPth & _
 "' AND UNIT='" & UNCD & "' AND DVCD= '" & DIVCODE & "' AND VTYP='PPF' AND DBCD='" & M_DBCD & _
 "' AND VBNO='" & CHALLAN & "' AND ICOD='" & ITMCODE & "'"
  
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
   ITMQTY = Val(nstr((Val(TEMPRS!PERC) * Val(TXTNTWT)) / 100, 10, 3))
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
 M_DBCD = "000003"
 
 CN.Execute "UPDATE PCKMST SET [LCNO]='" & CHALLAN & "' WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
            "' AND CODE='" & LSPKGCOD & "'"
 
Exit Sub

LAST:
Resume
MsgBox ERR.Description
ERROROCCUR = True
CN.RollbackTrans
End Sub

Private Sub SetShadeName(SUBGRD As String)
Dim GRRS As ADODB.Recordset
Set GRRS = New ADODB.Recordset

   If GRRS.State = 1 Then GRRS.Close
   GRRS.Open "SELECT NAME FROM SUBGRDMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
             "' AND SUBGRD='" & SUBGRD & "'", CN, adOpenDynamic, adLockOptimistic
   If Not GRRS.EOF Then
      TXTSHADE = Trim(GRRS!NAME & "")
      Exit Sub
   End If
   GRRS.Close

End Sub

Private Function isAllowGRPacking() As Boolean
isAllowGRPacking = True

If SAVEFLAG Then TXTNTWT.Tag = 0

'REVERSE IF EDITING
'FIFO----------------------
Dim INDEX As Long
Dim BALQNTY As Double, TMPQTY As Double, NETWGT As Double
Dim ITMCODE As String

Dim TEMPRS As ADODB.Recordset: Set TEMPRS = New ADODB.Recordset
Dim FIFORS As ADODB.Recordset: Set FIFORS = New ADODB.Recordset

If TEMPRS.State = 1 Then TEMPRS.Close
TEMPRS.Open "SELECT NTWGT FROM BOXREGISTER " & _
            "WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCODE & _
            "' AND VBNO='" & BOXNO & "' AND PCOD='GRPACK'", CN, adOpenDynamic, adLockOptimistic
If TEMPRS.EOF Then
   Exit Function
End If

If Not TEMPRS.EOF Then
    M_PCOD = "GRPACK"
    NETWGT = Val(TEMPRS!NTWGT)
    BALQNTY = Val(TEMPRS!NTWGT)
    'FIND GRPACKING
    If FIFORS.State = 1 Then FIFORS.Close
    FIFORS.Open "SELECT NETWGT AS NTWGT,FRESH,WASTAGE FROM GRPACKING " & _
                "WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCODE & _
                "' AND ICOD='" & FINITMCOD & "' AND FRESH > 0 AND RECSTAT<>'D' ORDER BY VBDT DESC", CN, adOpenDynamic, adLockOptimistic
            
    Do While Not FIFORS.EOF
   
        TMPQTY = Val(FIFORS!FRESH)
            
        If BALQNTY > TMPQTY Then
           FIFORS!FRESH = 0
           BALQNTY = BALQNTY - TMPQTY
           FIFORS.Update
        ElseIf BALQNTY > 0 Or BALQNTY = TMPQTY Then
           FIFORS!FRESH = Val(FIFORS!FRESH) - BALQNTY
           FIFORS.Update
           BALQNTY = 0
           Exit Do
        End If
                
    FIFORS.MoveNext
    Loop
    FIFORS.Close
    
End If
TEMPRS.Close

'================++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Dim CHKRS As ADODB.Recordset
Set CHKRS = New ADODB.Recordset

If CHKRS.State = 1 Then CHKRS.Close
CHKRS.Open "SELECT ISNULL(SUM(NETWGT-FRESH-WASTAGE),0) AS BALWGT FROM GRPACKING WHERE COMP='" & compPth & _
           "' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCODE & "' AND ICOD='" & FINITMCOD & _
           "' AND RECSTAT<>'D' AND ((NETWGT-FRESH-WASTAGE) > 0) ", CN, adOpenDynamic, adLockOptimistic
If CHKRS.EOF Then
   MsgBox "Goods Return Stock Not Support", vbInformation
   isAllowGRPacking = False
   Exit Function
Else
   If Val(CHKRS!BALWGT) <= 0 Then
      MsgBox "Goods Return Stock Not Support", vbInformation
      isAllowGRPacking = False
      Exit Function
   End If
   
   If Val(CHKRS!BALWGT) < Val(TXTNTWT) Then
      MsgBox "Goods Return Stock Not Support,balance upto " & CStr(Val(CHKRS!BALWGT)), vbInformation
      If TXTGRWT.Enabled Then TXTGRWT.SetFocus
      isAllowGRPacking = False
      Exit Function
   End If
   
   'Reconcile with GRPacking module
   Dim DSPQTY As Double: DSPQTY = Val(TXTNTWT)
   Dim RSFIRST As ADODB.Recordset
   Set RSFIRST = New ADODB.Recordset
   Dim RSSECOND As ADODB.Recordset
   Set RSSECOND = New ADODB.Recordset

   If RSFIRST.State = 1 Then RSFIRST.Close
   RSFIRST.Open "SELECT * FROM GRPACKING WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCODE & _
                "' AND ICOD='" & FINITMCOD & "' AND ((NETWGT-FRESH-WASTAGE) > 0)  AND RECSTAT<>'D' ORDER BY VBDT", CN, adOpenDynamic, adLockOptimistic
   Do While Not RSFIRST.EOF
      
      TMPQTY = Val(RSFIRST!NETWGT) - Val(RSFIRST!FRESH) - Val(RSFIRST!WASTAGE)
        
        If DSPQTY >= TMPQTY Then
           RSFIRST!FRESH = RSFIRST!FRESH + TMPQTY
           DSPQTY = DSPQTY - TMPQTY
        ElseIf DSPQTY > 0 Then
           RSFIRST!FRESH = RSFIRST!FRESH + DSPQTY
           DSPQTY = 0
        End If
   
  RSFIRST.MoveNext
  Loop
  RSFIRST.Close

End If
 
End Function

