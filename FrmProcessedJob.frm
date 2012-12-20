VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WELCHBUTTON.OCX"
Begin VB.Form FrmProcessedJob 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GRN Entry For Processed Job"
   ClientHeight    =   7140
   ClientLeft      =   375
   ClientTop       =   1110
   ClientWidth     =   11580
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7140
   ScaleWidth      =   11580
   Begin VB.Frame FRMBTRM 
      Height          =   2655
      Left            =   7920
      TabIndex        =   48
      Top             =   4440
      Width           =   3615
      Begin VB.TextBox TXTBNET 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
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
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   51
         Text            =   "0.00"
         Top             =   2160
         Width           =   1905
      End
      Begin VB.TextBox txtBEdit 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1440
         TabIndex        =   50
         Top             =   1320
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.TextBox TXTADLS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
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
         Left            =   1800
         TabIndex        =   49
         Text            =   "0.00"
         Top             =   1320
         Visible         =   0   'False
         Width           =   1575
      End
      Begin MSFlexGridLib.MSFlexGrid flexBTRM 
         Height          =   1995
         Left            =   120
         TabIndex        =   52
         Top             =   120
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   3519
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
         Left            =   120
         TabIndex        =   53
         Top             =   2280
         Width           =   1305
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      TabIndex        =   41
      Top             =   120
      Width           =   11415
      Begin VB.OptionButton optRGP 
         BackColor       =   &H00C0C0FF&
         Caption         =   "&RETURNABLE RECEIVED"
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
         Left            =   6960
         TabIndex        =   8
         Top             =   240
         Width           =   3495
      End
      Begin VB.OptionButton optJob 
         BackColor       =   &H00C0C0FF&
         Caption         =   "&JOB WORK RECEIVED"
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
         Left            =   2760
         TabIndex        =   7
         Top             =   240
         Value           =   -1  'True
         Width           =   3255
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Type of &Transaction    (A)                                                                (B)    "
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
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   6735
      End
   End
   Begin VB.Frame frm_head 
      Height          =   1335
      Left            =   120
      TabIndex        =   40
      Top             =   720
      Width           =   11415
      Begin VB.TextBox TXTSBILLNO 
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   20
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox TXTSCHLN 
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   16
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox TXTVBNO 
         Height          =   285
         Left            =   9360
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox TXTDBAC 
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   240
         Width           =   6495
      End
      Begin MSComCtl2.DTPicker TXTSCHLNDT 
         Height          =   285
         Left            =   5880
         TabIndex        =   18
         Top             =   600
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
         Format          =   24248321
         CurrentDate     =   39347
      End
      Begin MSComCtl2.DTPicker TXTVBDT 
         Height          =   315
         Left            =   9360
         TabIndex        =   13
         Top             =   600
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
         Format          =   24248321
         CurrentDate     =   39347
      End
      Begin MSComCtl2.DTPicker TXTSBILLDATE 
         Height          =   285
         Left            =   5880
         TabIndex        =   22
         Top             =   960
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
         Format          =   24248321
         CurrentDate     =   39347
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier &Bill No."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Bill Da&te :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4680
         TabIndex        =   21
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Chln &Date :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4680
         TabIndex        =   17
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier C&hln No."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label LBLBILLDATE 
         BackStyle       =   0  'Transparent
         Caption         =   "G&RN Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8400
         TabIndex        =   12
         Top             =   600
         Width           =   975
      End
      Begin VB.Label LBLBILLNO 
         BackStyle       =   0  'Transparent
         Caption         =   "&GRN No."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8400
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
      Begin VB.Label LBLDRAC 
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier &Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame ITMFRM 
      Height          =   2295
      Left            =   120
      TabIndex        =   38
      Top             =   2040
      Width           =   11415
      Begin VB.TextBox TXTTAMT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
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
         Left            =   9480
         TabIndex        =   59
         Text            =   "0.00"
         Top             =   1920
         Width           =   1575
      End
      Begin VB.CommandButton cmddelitm 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Remove Item"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   1920
         Width           =   1455
      End
      Begin VB.TextBox TXTTPCS 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
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
         Left            =   3840
         TabIndex        =   35
         Top             =   1920
         Width           =   855
      End
      Begin VB.TextBox TXTTQTY 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
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
         Left            =   6360
         TabIndex        =   37
         Top             =   1920
         Width           =   1215
      End
      Begin MSFlexGridLib.MSFlexGrid FLEX 
         Height          =   1695
         Left            =   0
         TabIndex        =   23
         Top             =   120
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   2990
         _Version        =   393216
         Cols            =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   "Item Value"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8160
         TabIndex        =   60
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label LBLGRS 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9000
         TabIndex        =   39
         Top             =   1920
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Total Pcs"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2760
         TabIndex        =   34
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Total Quantity"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4800
         TabIndex        =   36
         Top             =   1920
         Width           =   1455
      End
   End
   Begin WelchButton.lvButtons_H cmdAdd 
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   6540
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
      Image           =   "FrmProcessedJob.frx":0000
      cBack           =   -2147483633
   End
   Begin WelchButton.lvButtons_H cmdEdit 
      Height          =   495
      Left            =   5280
      TabIndex        =   4
      Top             =   6540
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
      Image           =   "FrmProcessedJob.frx":039A
      cBack           =   -2147483633
   End
   Begin WelchButton.lvButtons_H cmdDelete 
      Height          =   495
      Left            =   4080
      TabIndex        =   3
      Top             =   6540
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
      Image           =   "FrmProcessedJob.frx":0734
      cBack           =   -2147483633
   End
   Begin WelchButton.lvButtons_H cmdSave 
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   6540
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
      Image           =   "FrmProcessedJob.frx":0ACE
      cBack           =   -2147483633
   End
   Begin WelchButton.lvButtons_H cmdCancel 
      Height          =   495
      Left            =   2880
      TabIndex        =   2
      Top             =   6540
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
      Image           =   "FrmProcessedJob.frx":1858
      cBack           =   -2147483633
   End
   Begin WelchButton.lvButtons_H cmdExit 
      Height          =   495
      Left            =   6480
      TabIndex        =   5
      Top             =   6540
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
      Image           =   "FrmProcessedJob.frx":1CAA
      cBack           =   -2147483633
   End
   Begin TabDlg.SSTab FRMLRDTL 
      Height          =   1935
      Left            =   120
      TabIndex        =   43
      Top             =   4440
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   3413
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      ForeColor       =   128
      TabCaption(0)   =   "Description"
      TabPicture(0)   =   "FrmProcessedJob.frx":20FC
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label15"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label12"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label10"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label14"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "TXTSDESC"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "TXTSAMT"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "chkReturnable"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "TXTMDESC"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "TXTITOT"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "   Transport Detail"
      TabPicture(1)   =   "FrmProcessedJob.frx":2118
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "TXTTRNM"
      Tab(1).Control(1)=   "TXTVHCL"
      Tab(1).Control(2)=   "TXTLRNO"
      Tab(1).Control(3)=   "TXTRMRK"
      Tab(1).Control(4)=   "TXTLRDT"
      Tab(1).Control(5)=   "LBLTRNM"
      Tab(1).Control(6)=   "LBLVHCL"
      Tab(1).Control(7)=   "LBLLRNO"
      Tab(1).Control(8)=   "LBLLRDT"
      Tab(1).Control(9)=   "LBLRMRK"
      Tab(1).ControlCount=   10
      Begin VB.TextBox TXTITOT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   1560
         TabIndex        =   26
         Text            =   "0.00"
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox TXTTRNM 
         Height          =   285
         Left            =   -73800
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   480
         Width           =   5775
      End
      Begin VB.TextBox TXTVHCL 
         Height          =   285
         Left            =   -73800
         MaxLength       =   20
         TabIndex        =   30
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox TXTLRNO 
         Height          =   285
         Left            =   -73800
         MaxLength       =   20
         TabIndex        =   31
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox TXTRMRK 
         Height          =   285
         Left            =   -73800
         MaxLength       =   250
         TabIndex        =   33
         Top             =   1560
         Width           =   5775
      End
      Begin VB.TextBox TXTMDESC 
         Height          =   285
         Left            =   1560
         MaxLength       =   250
         TabIndex        =   27
         Top             =   1560
         Width           =   5535
      End
      Begin VB.CheckBox chkReturnable 
         Caption         =   "Transport Detail Required"
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
         TabIndex        =   28
         Top             =   60
         Width           =   2655
      End
      Begin VB.TextBox TXTSAMT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   1560
         TabIndex        =   24
         Text            =   "0.00"
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox TXTSDESC 
         Height          =   285
         Left            =   1560
         MaxLength       =   250
         TabIndex        =   25
         Top             =   840
         Width           =   5535
      End
      Begin MSComCtl2.DTPicker TXTLRDT 
         Height          =   300
         Left            =   -71520
         TabIndex        =   32
         Top             =   1200
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   529
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
         Format          =   24248321
         CurrentDate     =   39347
      End
      Begin VB.Label LBLTRNM 
         Caption         =   "Transport :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   58
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label LBLVHCL 
         Caption         =   "Vehicle No."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   57
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label LBLLRNO 
         Caption         =   "L.R.No."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   56
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label LBLLRDT 
         Caption         =   "L.R.Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -72360
         TabIndex        =   55
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label LBLRMRK 
         Caption         =   "Remarks :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   54
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label14 
         Caption         =   "Material Desc. : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "Material Cost :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label12 
         Caption         =   "Service Amt."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label15 
         Caption         =   "Service Desc. : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   840
         Width           =   1455
      End
   End
   Begin VB.Shape Shape1 
      Height          =   615
      Left            =   120
      Top             =   6480
      Width           =   7695
   End
End
Attribute VB_Name = "FrmProcessedJob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'FIFO
Public M_ISSUE As String

Public TABLENAME As String
Public SUMMARYTABLE As String
Public M_DBCD_DIRIVR As String
Public M_DBCD As String
Dim R_VTYP As String
Public ALLOWEDITDEL As Boolean
Public SAVEFLAG As Boolean
Dim CHK_FLX As Boolean
Dim FLXROW As Double
Dim FLXCOL As Double
Dim Emptycell As Boolean
Dim MRGN As String
Dim SPECI As String
'BILLING TERMS
Dim M_OPER(0 To 20) As String
Dim M_PERC(0 To 20) As Double
Dim M_POSTCOD(0 To 20) As String
Dim M_NICK(0 To 20) As String
Dim M_POSTYESNO(0 To 20) As String
Dim M_FMLA(0 To 20) As String
Dim M_RDOF(0 To 20) As String
Dim M_BILRDOF As String
Dim ITEMVAT As Double
Dim DISCOUNT_ROW As String

Dim chgFlag As Boolean
Dim calbtm As Boolean
'----------------------------------------------

Private Sub chkReturnable_Click()
If chkReturnable.Value = 1 Then
   FRMLRDTL.Tab = 1
   If TXTTRNM.Enabled Then TXTTRNM.SetFocus
Else
   FRMLRDTL.Tab = 0
   TXTVHCL = Empty: TXTTRNM = Empty: TXTLRNO = Empty: TXTRMRK = Empty
   If cmdSave.Enabled Then cmdSave.SetFocus
End If
End Sub

Private Sub chkReturnable_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   If cmdSave.Enabled Then cmdSave.SetFocus
End If
End Sub

Private Sub cmddelitm_Click()
  If FLEX.ROW > 1 Then
    FLEX.RemoveItem (FLEX.ROW)
    TXTTPCS.Text = 0
    TXTTQTY.Text = 0
    TXTTAMT.Text = 0
    Dim i As Double
    i = 1
    For i = 1 To FLEX.Rows - 1
      FLEX.TextMatrix(i, 0) = i
      TXTTPCS.Text = Format(Val(TXTTPCS.Text) + Val(FLEX.TextMatrix(i, 5)), "######")
      TXTTQTY.Text = Format(Val(TXTTQTY.Text) + Val(FLEX.TextMatrix(i, 6)), "########.000")
      TXTTAMT.Text = Format(Val(TXTTAMT.Text) + Val(FLEX.TextMatrix(i, 8)), "########.00")
    Next
    FLEX.Refresh
    FLEX.ROW = FLEX.Rows - 1
    FLEX.COL = 8
    FLEX.SetFocus
  End If
  cmddelitm.Enabled = False
End Sub

Private Sub cmddelitm_LostFocus()
  cmddelitm.Enabled = False
End Sub

Private Sub Flex_Click()
  cmddelitm.Enabled = True
End Sub

Private Sub Form_Activate()
 Call ColorComponent(Me)
 Me.BackColor = RGB(RED, GREEN, BLUE)
End Sub

Private Sub Form_Load()
'TEMP : FIFO
  FIFOREQ = "Y"
  '----------
  Call ColorComponent(Me)
  Call CenterChild(frm_Main, Me)
  Me.Left = 130
  Emptycell = True
  
  TXTSBILLDATE = Now
  TXTVBDT = Now: TXTLRDT = Now
  TXTSCHLNDT = Now
  TXTVBDT.MinDate = FSDT
  TXTVBDT.MaxDate = FEDT
   
  Call setflexhead
  Call btn_sts(True)
  
  If optJob.Value = True Then M_DBCD = "000003": R_VTYP = "ANX" Else M_DBCD = "000004": R_VTYP = "RGP"
  TXTVBNO = GenVNO("IVR", M_DBCD)
  
  Call FIL_Billingterm
  Me.KeyPreview = True
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If UCase(ActiveControl.NAME) = "TXTSCHLN" And TXTSCHLN = Empty Then Exit Sub
  If (ActiveControl.NAME = "TXTDBAC" Or UCase(ActiveControl.NAME) = "TXTRECNO" Or UCase(ActiveControl.NAME) = "FLEX" Or UCase(ActiveControl.NAME) = "TXTRMRK") Then Exit Sub
  
  If UCase(ActiveControl.NAME) = "CHKRETURNABLE" Then
     Exit Sub
  End If
  
  If UCase(ActiveControl.NAME) = "CMDSAVE" Or UCase(ActiveControl.NAME) = "CMDCANCEL" Or UCase(ActiveControl.NAME) = "CMDEDIT" Then
     Exit Sub
  End If
  
   If UCase(ActiveControl.NAME) = "CHKRETURNABLE" Then
     Exit Sub
  End If
  
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
  End If
End Sub

Public Sub cmdAdd_Click()
  SAVEFLAG = True
  btn_sts (False)
  cmddelitm.Enabled = False
   
  If optJob.Value = True Then M_DBCD = "000003": R_VTYP = "ANX" Else M_DBCD = "000004": R_VTYP = "RGP"
  TXTVBNO = GenVNO("IVR", M_DBCD)
  
  If TXTDBAC.Enabled = True Then
    TXTDBAC.SetFocus
  End If
  
End Sub

Private Sub cmdCancel_Click()
  ClsData (Me)
  FLEX.Rows = 1
  FLEX.Rows = 2
  btn_sts (True)
  cmdAdd.SetFocus
  optJob.Enabled = True
  optRGP.Enabled = True
End Sub

Private Sub cmdDelete_Click()
  If M_USRSECLEVL = "1" Then
    If ReadConfigMaster("000037", 6, "M") = False Then ModuleDeniedMessage: Exit Sub
  End If
  ALLOWEDITDEL = True
  SAVEFLAG = False
  btn_sts (False)
  FrmProcessedJobList.Show 1
  
  If TXTVBNO = Empty Then Call cmdCancel_Click: Exit Sub
    
  If ALLOWEDITDEL = False Then
    'MsgBox "Purchase of this GRN have been made can not edit/delete ", vbInformation
   Else
     Dim AYS
     AYS = MsgBox("Are you sure to delete the invoice ", vbYesNo)
     If AYS = vbYes Then
        CN.BeginTrans
        'Call ResetAction
        
        CN.Execute "UPDATE JOBOUT SET RECSTAT='D' WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
        "' and VTYP = 'IVR' AND DBCD='" & M_DBCD & "' AND VBNO='" & TXTVBNO & "' AND RECSTAT<>'D'"
  
        CN.Execute "UPDATE STORETRAN SET RECSTAT='D' WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
        "' and VTYP = 'IVR' AND DBCD='" & M_DBCD & "' AND VBNO='" & TXTVBNO & "' AND RECSTAT<>'D'"
        
        Call DAILYSTATUS("IVR", GetCode("ACCMST", TXTDBAC, "NAME", "CODE"), M_DBCD, Val(TXTTQTY), TXTVBNO, Val(TXTITOT), cUName, "D", Now, TXTVBDT)
       
        CN.CommitTrans
      End If
  End If
  Call cmdCancel_Click
End Sub

Private Sub cmdEdit_Click()
  If M_USRSECLEVL = "1" Then
     If ReadConfigMaster("000037", 5, "M") = False Then ModuleDeniedMessage: Exit Sub
  End If
  ALLOWEDITDEL = True
  SAVEFLAG = False
  
  
  M_ISSUE = "N"
  
  FrmProcessedJobList.Show 1
  
  If M_DBCD <> Empty Then
     If M_DBCD = "000003" Then
        optRGP.Enabled = False
     Else
        optJob.Enabled = False
     End If
  Else
    Call ClsData(Me)
    btn_sts (True)
    cmdAdd.SetFocus
    Exit Sub
  End If
  
   If ALLOWEDITDEL = False Then
        'MsgBox "Purchase of this GRN have been made can not edit/delete ", vbInformation
        'Call ClsData(Me)
        'btn_sts (True)
        'cmdAdd.SetFocus
   Else
        'Check for Receipt and Payment Entires
        'If M_SRNO = Empty Then
        '  Exit Sub
        'End If
        btn_sts (False)
        TXTDBAC.SetFocus
  End If
  
  'FIFO
    If Trim(M_ISSUE) = "Y" Then
        MsgBox "You Can Not Edit GRN Detail!! Issue Entry Exist!!", vbExclamation, "Access Denied"
        cmdDelete.Enabled = False
        cmdSave.Enabled = False
        Exit Sub
    End If
  '------------
End Sub

Private Sub CMDEXIT_Click()
  Unload Me
End Sub

Private Sub FLEX_EnterCell()
  FLEX.CellBackColor = RGB(BRED, BGREEN, BBLUE)
  Emptycell = True
End Sub

Private Sub FLEX_GotFocus()
  Me.KeyPreview = False
  FLEX.TextMatrix(FLEX.ROW, 0) = FLEX.ROW
End Sub

Private Sub Flex_LeaveCell()
  Dim i As Double
    
  FLEX.CellBackColor = vbWhite
  FLEX.TextMatrix(FLEX.ROW, 8) = Format(Val(FLEX.TextMatrix(FLEX.ROW, 6)) * Val(FLEX.TextMatrix(FLEX.ROW, 7)), "#########.00")
  
  TXTTPCS.Text = 0
  TXTTQTY.Text = 0
  TXTTAMT.Text = 0
  
  For i = 1 To FLEX.Rows - 1
    TXTTPCS.Text = Val(Format(Val(TXTTPCS.Text) + Val(FLEX.TextMatrix(i, 5)), "######"))
    TXTTQTY.Text = Format(Val(TXTTQTY.Text) + Val(FLEX.TextMatrix(i, 6)), "########.000")
    TXTTAMT.Text = Format(Val(TXTTAMT.Text) + Val(FLEX.TextMatrix(i, 8)), "########.00")
  Next
  
End Sub

Private Sub FLEX_LostFocus()
  Dim i As Double
    
  FLEX.CellBackColor = vbWhite
  FLEX.TextMatrix(FLEX.ROW, 8) = Format(Val(FLEX.TextMatrix(FLEX.ROW, 6)) * Val(FLEX.TextMatrix(FLEX.ROW, 7)), "#########.00")
  
  TXTTPCS.Text = 0
  TXTTQTY.Text = 0
  TXTTAMT.Text = 0
  
  For i = 1 To FLEX.Rows - 1
    TXTTPCS.Text = Val(Format(Val(TXTTPCS.Text) + Val(FLEX.TextMatrix(i, 5)), "######"))
    TXTTQTY.Text = Format(Val(TXTTQTY.Text) + Val(FLEX.TextMatrix(i, 6)), "########.000")
    TXTTAMT.Text = Format(Val(TXTTAMT.Text) + Val(FLEX.TextMatrix(i, 8)), "########.00")
  Next
  
  Me.KeyPreview = True
End Sub

Private Sub setflexhead()

    FLEX.TextMatrix(0, 0) = "Sr."
    FLEX.TextMatrix(0, 1) = "Annexture No."
    FLEX.TextMatrix(0, 2) = "Item Name"
    FLEX.TextMatrix(0, 3) = "Lot No."
    FLEX.TextMatrix(0, 4) = "Cops"
    FLEX.TextMatrix(0, 5) = "Pcs"
    FLEX.TextMatrix(0, 6) = "Qnty"
    FLEX.TextMatrix(0, 7) = "Rate"
    FLEX.TextMatrix(0, 8) = "Amount"
    FLEX.TextMatrix(0, 9) = "ICOD"
    
    FLEX.ColWidth(0) = 350
    FLEX.ColWidth(1) = 1500
    FLEX.ColWidth(2) = 2800
    FLEX.ColWidth(3) = 1200
    FLEX.ColWidth(4) = 700
    FLEX.ColWidth(5) = 600
    FLEX.ColWidth(6) = 1300
    FLEX.ColWidth(7) = 1300
    FLEX.ColWidth(8) = 1600
    FLEX.ColWidth(9) = 0
        
    FLEX.ColAlignment(2) = 1
    FLEX.ColAlignment(4) = 1
    FLEX.ColAlignment(5) = 1
    FLEX.ColAlignment(6) = 1
    FLEX.ColAlignment(7) = 1
    
End Sub

Public Sub btn_sts(Yes As Boolean)
    cmdSave.Enabled = Not Yes
    cmdCancel.Enabled = Not Yes
    cmdAdd.Enabled = Yes
    cmdEdit.Enabled = Yes
    cmdDelete.Enabled = Yes
    frm_head.Enabled = Not Yes
    ITMFRM.Enabled = Not Yes
    FRMLRDTL.Enabled = Not Yes
End Sub

Private Sub LBLGRS_Change()
  TXTTAMT = Format(LBLGRS, "#########.00")
End Sub

Private Sub optJob_Click()
  FLEX.TextMatrix(0, 1) = "Annexture No."
End Sub

Private Sub optRGP_Click()
  FLEX.TextMatrix(0, 1) = "Issue Chln No."
End Sub



Private Sub TXTDBAC_GotFocus()
TXTDBAC.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTDBAC_KeyDown(KeyCode As Integer, Shift As Integer)
   Me.KeyPreview = False
   If KeyCode = vbKeyF2 Or Trim(TXTDBAC.Text) = Empty Then
        NEW_VISIBLE = False
        M_DESC = Empty
        Key = Empty
        TXTDBAC.Text = SearchList1("SELECT  TOP 20 CODE,NAME FROM ACCMST", 0, TXTDBAC, "List of Debit A/c")
    ElseIf KeyCode = vbKeyDelete Then
        TXTDBAC = Empty
    End If
    Me.KeyPreview = True
End Sub

Private Sub TXTDBAC_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn And (Not Trim(TXTDBAC.Text) = Empty) Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub TXTDBAC_LostFocus()
 TXTDBAC.BackColor = vbWhite
End Sub

Private Sub TXTITOT_Change()
If flexBTRM.Rows > 0 Then
    flexBTRM.COL = 0
    flexBTRM.ROW = 0
  End If
  calBTRM 0
  Call calADLS
End Sub

Private Sub TXTMDESC_GotFocus()
 SendKeys "{HOME}+{END}"
 TXTMDESC.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTMDESC_LostFocus()
  TXTMDESC.BackColor = vbWhite
End Sub

Private Sub TXTRMRK_KeyPress(KeyAscii As Integer)
 If KeyAscii = vbKeyReturn And cmdSave.Enabled Then cmdSave.SetFocus
End Sub

Private Sub TXTSAMT_Change()
If flexBTRM.Rows > 0 Then
    flexBTRM.COL = 0
    flexBTRM.ROW = 0
  End If
  calBTRM 0
  Call calADLS
End Sub

Private Sub TXTSAMT_GotFocus()
  SendKeys "{HOME}+{END}"
  TXTSAMT.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTSAMT_KeyPress(KeyAscii As Integer)
 If KeyAscii = 46 And InStr(1, TXTSAMT, ".", vbTextCompare) > 0 Then KeyAscii = 0
 If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 46 Then KeyAscii = 0
End Sub

Private Sub TXTSAMT_LostFocus()
  TXTSAMT.BackColor = vbWhite
End Sub

Private Sub TXTITOT_GotFocus()
  SendKeys "{HOME}+{END}"
  TXTITOT.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTITOT_KeyPress(KeyAscii As Integer)
 If KeyAscii = 46 And InStr(1, TXTITOT, ".", vbTextCompare) > 0 Then KeyAscii = 0
 If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 46 Then KeyAscii = 0
End Sub

Private Sub TXTITOT_LostFocus()
  TXTITOT.BackColor = vbWhite
End Sub

Private Sub TXTSDESC_GotFocus()
 SendKeys "{HOME}+{END}"
 TXTSDESC.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTSDESC_LostFocus()
TXTSDESC.BackColor = vbWhite
End Sub

Private Sub TXTSBILLDATE_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TXTSBILLNO_GotFocus()
 TXTSBILLNO.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTSBILLNO_LostFocus()
 TXTSBILLNO.BackColor = vbWhite
End Sub

Private Sub TXTSCHLNDT_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    TXTVBDT.SetFocus
  End If
End Sub

Private Sub TXTSCHLN_GotFocus()
TXTSCHLN.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTSCHLN_LostFocus()
TXTSCHLN.BackColor = vbWhite
End Sub

Private Sub txtLRDT_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
  End If
End Sub
Private Sub TXTLRNO_GotFocus()
  TXTLRNO.BackColor = RGB(BRED, BGREEN, BBLUE)
  Me.KeyPreview = True
  TXTLRNO.SelStart = 0
  TXTLRNO.SelLength = Len(Trim(TXTLRNO))
End Sub

Private Sub TXTLRNO_LostFocus()
 TXTLRNO.BackColor = vbWhite
End Sub

Private Sub TXTRMRK_GotFocus()
  TXTRMRK.BackColor = RGB(BRED, BGREEN, BBLUE)
  TXTRMRK.SelStart = 0
  TXTRMRK.SelLength = Len(Trim(TXTRMRK))
End Sub

Private Sub TXTRMRK_LostFocus()
  TXTRMRK.BackColor = vbWhite
End Sub

Private Sub TXTTRNM_GotFocus()
 TXTTRNM.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTTRNM_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Or Trim(TXTTRNM.Text) = Empty Then
        NEW_VISIBLE = True: M_DESC = Empty:  Key = Empty
        TXTTRNM.Text = SearchList1("SELECT  TOP 20 CODE,NAME FROM TRANSPORTMST", 0, TXTTRNM, "List of Transporter")
        If key_PressNew = True Then
            M_DESC = ""
            Key = ""
            TXTTRNM.Text = ""
            frmTransportMaster.Show
        End If
    ElseIf KeyCode = vbKeyDelete Then
        TXTTRNM = Empty
    End If
End Sub

Private Sub TXTTRNM_LostFocus()
 TXTTRNM.BackColor = vbWhite
End Sub

Private Sub TXTVBDT_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
     If TXTSBILLNO.Enabled Then TXTSBILLNO.SetFocus
  End If
End Sub

Private Function CHKSAVEDATA() As Boolean
  Dim CHKRS As New ADODB.Recordset
  Set CHKRS = New ADODB.Recordset
  
  'Party  A/c Code
  If CHKRS.State = 1 Then CHKRS.Close
  CHKRS.Open "SELECT * from ACCMST WHERE NAME='" & TXTDBAC.Text & "'", CN, adOpenKeyset, adLockPessimistic
  If CHKRS.EOF Then
     MsgBox "Debit A/c Name Not Define ", vbCritical
     CHKSAVEDATA = False
     Exit Function
  End If
    
  If CHKRS.State = 1 Then CHKRS.Close
  CHKRS.Open "SELECT VBNO FROM JOBOUT WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
  "' and VTYP = 'IVR' AND DBCD='" & M_DBCD & "' AND VBNO='" & TXTVBNO & _
  "' AND RECSTAT<>'D'", CN, adOpenDynamic, adLockOptimistic
  
  If Not CHKRS.EOF Then
    If CHKRS!VBNO <> TXTVBNO Then
      MsgBox "Duplicate GRN No. !!!! ", vbCritical
      CHKSAVEDATA = False
      Exit Function
    End If
  End If
  
  CHKSAVEDATA = True
End Function

Private Sub cmdSave_Click()
  On Error GoTo LAST
  If CHKSAVEDATA = False Then
    Exit Sub
  End If
  
  Call CHKFLEX
  
  If Not CHK_FLX Then
    MsgBox "Invalid Item Detail"
    FLEX.ROW = FLXROW
    FLEX.COL = FLXCOL
    Exit Sub
  End If
     
  If optJob.Value = True Then M_DBCD = "000003": R_VTYP = "ANX" Else M_DBCD = "000004": R_VTYP = "RGP"
  
  If SAVEFLAG = True Then
        TXTVBNO = GenVNO("IVR", M_DBCD)
        
        Dim SAVDAT As ADODB.Recordset
        Set SAVDAT = New ADODB.Recordset
        If SAVDAT.State = 1 Then SAVDAT.Close
        SAVDAT.Open "SELECT * FROM JOBOUT WHERE COMP='" & compPth & "' AND VTYP='IVR' AND UNIT='" & UNCD & _
                    "' AND  VBNO='" & TXTVBNO & "' AND DBCD='" & M_DBCD & "' ", CN, adOpenDynamic, adLockOptimistic
          If Not SAVDAT.EOF Then
            MsgBox "Duplicate GRN No. Make Change in Configuration for GRN No."
            cmdSave.SetFocus
            Exit Sub
          End If
  End If
     
  Call SAVERECIVR
  
  If SAVEFLAG = True Then
    MsgBox "Your GRN No. is " + TXTVBNO.Text
  End If
  
  Call cmdCancel_Click
  Exit Sub
LAST:
  MsgBox ERR.Description
End Sub

Private Sub SAVERECIVR()
  
  On Error GoTo LAST
  Dim SQL As String
  
  Dim SAVDAT As New ADODB.Recordset 'USE
  Dim MSTDAT As New ADODB.Recordset
  
  Dim M_DRAC As String 'USE
  Dim M_TRCD As String 'USE
    
  Dim i As Double
  Dim J As Double
  Set SAVDAT = New ADODB.Recordset 'USE
  Set MSTDAT = New ADODB.Recordset
  
  'Party A/c
  If SAVDAT.State = 1 Then SAVDAT.Close
  SAVDAT.Open "SELECT CODE FROM ACCMST WHERE NAME='" & TXTDBAC.Text & "'", CN, adOpenDynamic, adLockOptimistic
  If Not SAVDAT.EOF Then
   M_DRAC = SAVDAT!CODE & ""
  End If
  SAVDAT.Close
  
  'TRANSPORT
  If SAVDAT.State = 1 Then SAVDAT.Close
  SAVDAT.Open "SELECT CODE FROM TRANSPORTMST WHERE NAME='" & TXTTRNM.Text & "'", CN, adOpenDynamic, adLockOptimistic
  If Not SAVDAT.EOF Then
     M_TRCD = SAVDAT!CODE & ""
  End If
  SAVDAT.Close
  
  CN.BeginTrans
  
  Call DELETEIVR
  
  SQL = Empty
  If SAVDAT.State = 1 Then SAVDAT.Close
  
  SAVDAT.Open "SELECT * FROM JOBOUT WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' and VTYP = 'IVR' " & _
  "AND DBCD='" & M_DBCD & "' AND VBNO='" & TXTVBNO & "' AND RECSTAT<>'D'", CN, adOpenDynamic, adLockOptimistic
         
  i = 1
  For i = 1 To FLEX.Rows - 1
    If FLEX.TextMatrix(i, 1) <> Empty Then
        SAVDAT.AddNew
        SAVDAT!COMP = compPth
        SAVDAT!unit = UNCD
        SAVDAT!VTYP = "IVR"
        SAVDAT!dbcd = M_DBCD
        SAVDAT!VBNO = TXTVBNO
        SAVDAT!SRCH = i
        SAVDAT!Date = Format(TXTVBDT, "YYYY/MM/DD")
        SAVDAT!PCOD = M_DRAC
        SAVDAT!ICOD = Trim(FLEX.TextMatrix(i, 9))
        
        SAVDAT!chln = TXTSCHLN
        SAVDAT!CHDT = Format(TXTSCHLNDT, "YYYY/MM/DD")
        SAVDAT!BILNO = TXTSBILLNO
        SAVDAT!BILDT = Format(TXTSBILLDATE, "YYYY/MM/DD")
        
        SAVDAT!LRNO = TXTLRNO
        SAVDAT!LRDT = Format(TXTLRDT, "YYYY/MM/DD")
        SAVDAT!TRCD = M_TRCD
        SAVDAT!VHCLNO = Trim(TXTVHCL)
                
        SAVDAT!IDNO = ""
        SAVDAT!COPS = Val(FLEX.TextMatrix(i, 4))
        SAVDAT!PCES = Val(FLEX.TextMatrix(i, 5))
        SAVDAT!QNTY = Val(FLEX.TextMatrix(i, 6))
        SAVDAT!RATE = Val(FLEX.TextMatrix(i, 7))
        SAVDAT!AMNT = Val(FLEX.TextMatrix(i, 8))
        SAVDAT!OPER = "-"
        SAVDAT!RECNO = Trim(FLEX.TextMatrix(i, 1))
        'SAVDAT!REC_QNTY = FindRecQnty(Trim(FLEX.TextMatrix(i, 9)), Val(FLEX.TextMatrix(i, 6)))
        SAVDAT!ltno = Trim(FLEX.TextMatrix(i, 3))
                       
        SAVDAT!XDAYS = 0
        SAVDAT!RMRK = Trim(TXTRMRK)
        SAVDAT!User = cUName
        SAVDAT!SYSR = "N"
        SAVDAT!RECSTAT = "A"
        SAVDAT!Mode = IIf(chkReturnable.Value = 1, "Y", "N")
        
        SAVDAT.Update
   End If
 Next
    
 '===================================================================================================
 'STORETRAN (ADD)
 '===================================================================================================
        
 If SAVDAT.State = 1 Then SAVDAT.Close
  SAVDAT.Open "SELECT * FROM STORETRAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND VTYP='IVR' AND DBCD='" & M_DBCD & "' AND VBNO='" & TXTVBNO & "' AND RECSTAT='A'", CN, adOpenDynamic, adLockOptimistic
  
  Dim AI As String
  Dim BQ As Double
  Dim CR As Double
  
  i = 1
  For i = 1 To FLEX.Rows - 1
    SAVDAT.AddNew
    SAVDAT!COMP = compPth
    SAVDAT!VTYP = "IVR"
    SAVDAT!SRNO = ""
    SAVDAT!SRCH = i
    SAVDAT!VBNO = TXTVBNO.Text
    SAVDAT!chln = TXTSCHLN
    SAVDAT!CHDT = Format(TXTSCHLNDT, "YYYY/MM/DD")
    SAVDAT!Date = Format(TXTVBDT, "YYYY/MM/DD")
    SAVDAT!dbcd = M_DBCD
    SAVDAT!CRAC = ""
    SAVDAT!DRAC = M_DRAC
    SAVDAT!PCOD = M_DRAC
    SAVDAT!DCOD = ""
    SAVDAT!ICOD = FLEX.TextMatrix(i, 9): AI = FLEX.TextMatrix(i, 9)
    SAVDAT!COPS = Val(FLEX.TextMatrix(i, 4))
    SAVDAT!PCES = Val(FLEX.TextMatrix(i, 5))
    SAVDAT!QNTY = Val(FLEX.TextMatrix(i, 6)): BQ = Val(FLEX.TextMatrix(i, 6))
    SAVDAT!RATE = Val(FLEX.TextMatrix(i, 7)): CR = Val(FLEX.TextMatrix(i, 7))
    SAVDAT!AMNT = Val(FLEX.TextMatrix(i, 8))
    SAVDAT!QORP = "Q"
    SAVDAT![User] = cUName
    If SAVEFLAG = True Then
      SAVDAT!SYSR = "N"
     Else
      SAVDAT!SYSR = "U"
    End If
    SAVDAT!OPER = "+"
    SAVDAT!DVCD = "000001"
    SAVDAT!unit = UNCD
    SAVDAT!RECSTAT = "A"
    SAVDAT!ltno = Trim(FLEX.TextMatrix(i, 3))
    SAVDAT!MRGN = Trim(FLEX.TextMatrix(i, 3))
    SAVDAT!SPECIFICATION = GetSpeci(FLEX.TextMatrix(i, 9))
    
    SAVDAT.Update
    
    If MSTDAT.State = 1 Then MSTDAT.Close
    MSTDAT.Open "SELECT * FROM MRGMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND MRGN='" & FLEX.TextMatrix(i, 3) & "' AND ICOD = '" & FLEX.TextMatrix(i, 9) & "'", CN, adOpenDynamic, adLockOptimistic
    If MSTDAT.EOF Then
      MSTDAT.AddNew
      MSTDAT!COMP = compPth
      MSTDAT!unit = UNCD
      MSTDAT!MRGN = Mid(UCase(FLEX.TextMatrix(i, 3)), 1, 19)
      MSTDAT!PCOD = Trim(M_DRAC)
      MSTDAT!ICOD = FLEX.TextMatrix(i, 9)
      MSTDAT.Update
    End If
   
  
    Call SetItemRate(AI, CR, BQ)
    Call SetItemBalQty("BALQ", AI, BQ, "+")
    
 Next
  
'------------FIFO----------------------
   If FIFOREQ = "Y" Then
      Call SetItemInfo
   End If
'======================================

  '===================================================================================================
'AUTOMATION ENTRY FOR SERVICE TAX
'===================================================================================================
  SQL = Empty
  If SAVDAT.State = 1 Then SAVDAT.Close
  SAVDAT.Open "SELECT * FROM GRN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
              "' AND VTYP='IVR' AND DBCD ='" & M_DBCD & "' AND VBNO = '" & Trim(TXTVBNO) & _
              "' AND RECSTAT<>'D'", CN, adOpenDynamic, adLockOptimistic
  If SAVDAT.EOF Then
      SAVDAT.AddNew
  End If
  
  SAVDAT!COMP = compPth
  SAVDAT!unit = UNCD
  SAVDAT!DVCD = "000001"
  SAVDAT!VTYP = "IVR"
  SAVDAT!SRNO = ""
  SAVDAT!SRCH = 1
  SAVDAT!dbcd = M_DBCD
  SAVDAT!Date = Format(TXTVBDT.Value, "YYYY/MM/DD")
  SAVDAT!VBNO = Trim(TXTVBNO)
  SAVDAT!DRAC = M_DRAC
  SAVDAT!PCOD = M_DRAC
  SAVDAT!TPCS = Val(TXTTPCS)
  SAVDAT!TQTY = Val(TXTTQTY)
  SAVDAT!SDESC = Trim(TXTSDESC)
  SAVDAT!SAMT = Val(TXTSAMT)
  SAVDAT!MDESC = Trim(TXTMDESC)
  SAVDAT!ITOT = Val(TXTITOT)
  SAVDAT!BADJ = Val(TXTITOT) + Val(TXTSAMT) - Val(TXTBNET)
  SAVDAT!BNET = Val(TXTBNET)
  SAVDAT!ACEFFECT = "Y"
  
  SAVDAT!LRNO = TXTLRNO
  SAVDAT!LRDT = Format(TXTLRDT, "YYYY/MM/DD")
  SAVDAT!VHCL = Trim(TXTVHCL)
  SAVDAT!TRCD = M_TRCD
  SAVDAT!BRMK = Trim(TXTRMRK)
  
  If SAVEFLAG = True Then
    SAVDAT!SYSR = "N"
   Else
    SAVDAT!SYSR = "U"
  End If
  SAVDAT![User] = cUName & ""
  SAVDAT!chln = Trim(TXTSCHLN)
  SAVDAT!CHDT = Format(TXTSCHLNDT, "YYYY/MM/DD")
  SAVDAT!GATD = Format(TXTSBILLDATE.Value, "YYYY/MM/DD")
  SAVDAT!CVBN = TXTSBILLNO
  SAVDAT!RECSTAT = "A"
  SAVDAT!BRMK = Trim(TXTSDESC.Text)
  
  i = 0
  For i = 0 To flexBTRM.Rows - 1
    J = 0
    For J = 0 To SAVDAT.Fields.COUNT - 1
      If Trim(SAVDAT.Fields(J).NAME) = Trim(flexBTRM.TextMatrix(i, 0)) Then
        SAVDAT.Fields(J).Value = Val(flexBTRM.TextMatrix(i, 2))
      End If
      If Trim(SAVDAT.Fields(J).NAME) = "PER" & Trim(flexBTRM.TextMatrix(i, 0)) Then
        SAVDAT.Fields(J).Value = Val(flexBTRM.TextMatrix(i, 1))
      End If
    Next
  Next
  
  SAVDAT.Update
    
  Dim EXCISE As ADODB.Recordset
  Set EXCISE = New ADODB.Recordset
  If EXCISE.State = 1 Then EXCISE.Close
  EXCISE.Open "SELECT * FROM EGPMAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
              "' AND VTYP='IVR'  AND DBCD ='" & M_DBCD & "' AND VBNO='" & TXTVBNO & "'", CN, adOpenDynamic, adLockOptimistic
  If EXCISE.EOF Then
     EXCISE.AddNew
  End If
  EXCISE!COMP = compPth
  EXCISE!unit = UNCD
  EXCISE!dbcd = M_DBCD
  EXCISE!VTYP = "IVR"
  EXCISE!VBNO = TXTVBNO
  EXCISE!Date = Format(TXTVBDT, "YYYY/MM/DD")
  EXCISE!SRNO = ""
  EXCISE!SRCH = 1
  EXCISE!chln = Trim(TXTSCHLN)
  EXCISE!CHDT = Format(TXTSCHLNDT, "YYYY/MM/DD")
  EXCISE!DRAC = M_DRAC & ""
  EXCISE!PCES = 0
  EXCISE!QNTY = 0
  EXCISE!AMNT = Val(TXTITOT) + Val(TXTSAMT)
  EXCISE!ITOT = Val(TXTITOT)
  EXCISE!BADJ = Val(TXTITOT) + Val(TXTSAMT) - Val(TXTBNET)
  EXCISE!BNET = Val(TXTBNET)
  EXCISE!TTYP = "NONE"
  EXCISE!RECSTAT = "A"
  
  i = 0
  For i = 0 To flexBTRM.Rows - 1
    J = 0
    For J = 0 To EXCISE.Fields.COUNT - 1
      If Trim(EXCISE.Fields(J).NAME) = Trim(flexBTRM.TextMatrix(i, 0)) Then
        EXCISE.Fields(J).Value = Val(flexBTRM.TextMatrix(i, 2))
      End If
      If Trim(EXCISE.Fields(J).NAME) = "PER" & Trim(flexBTRM.TextMatrix(i, 0)) Then
        EXCISE.Fields(J).Value = Val(flexBTRM.TextMatrix(i, 1))
      End If
    Next
  Next
  
  EXCISE.Update
  
'===================================================================================================
    
  'UPDATE VOUCHER TYPE MASTER
  If SAVEFLAG = True Then
    Call SetSRNO(TXTVBNO, "IVR", M_DBCD)
  End If
 '-------------------------
 'DAILYSTATUS ENTRY
  If SAVEFLAG = True Then
   Call DAILYSTATUS("IVR", M_DRAC, M_DBCD, Val(TXTTQTY), TXTVBNO, Val(TXTITOT), cUName, "N", Now, TXTVBDT)
  Else
   Call DAILYSTATUS("IVR", M_DRAC, M_DBCD, Val(TXTTQTY), TXTVBNO, Val(TXTITOT), cUName, "M", Now, TXTVBDT)
  End If
 '-------------------------
  CN.CommitTrans
  Exit Sub
  
LAST:
 MsgBox ERR.Description
 If SAVDAT.State = 1 Then
   SAVDAT.CancelUpdate
   SAVDAT.Close
 End If
 CN.RollbackTrans
End Sub

Private Sub DELETEIVR()
  On Error GoTo LAST
    
  CN.Execute "DELETE FROM JOBOUT WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' and VTYP = 'IVR' " & _
  "AND DBCD='" & M_DBCD & "' AND VBNO='" & TXTVBNO & "' AND RECSTAT<>'D'"
  
  CN.Execute "DELETE FROM STORETRAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' and VTYP = 'IVR' " & _
  "AND DBCD='" & M_DBCD & "' AND VBNO='" & TXTVBNO & "' AND RECSTAT<>'D'"
  
  CN.Execute "DELETE FROM GRNTRAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' and VTYP = 'IVR' " & _
  "AND DBCD='" & M_DBCD & "' AND VBNO='" & TXTVBNO & "'"
  
  Exit Sub
LAST:
  MsgBox ERR.Description
  Exit Sub
End Sub

Private Sub TXTVBNO_GotFocus()
 TXTVBNO.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTVBNO_LostFocus()
 TXTVBNO.BackColor = vbWhite
End Sub

Private Sub TXTVHCL_GotFocus()
  TXTVHCL.BackColor = RGB(BRED, BGREEN, BBLUE)
  TXTVHCL.SelStart = 0
  TXTVHCL.SelLength = Len(Trim(TXTVHCL))
End Sub

Private Sub flex_KeyPress(KeyAscii As Integer)
  On Error GoTo LAST
  FLEX.TextMatrix(FLEX.ROW, 0) = FLEX.ROW
  
  'CONTROL VARIABLE-----------------------------------------------------
  Dim ALLOW_KEY As Boolean, FWD_COL As Boolean, ENTER_PRESS As Boolean
  'DEFAULT VALUE
  FWD_COL = False: ALLOW_KEY = False
  '---------------------------------------------------------------------
    
  'LOCAL RECORD SET
  Dim MSTDAT As New ADODB.Recordset
  Set MSTDAT = New ADODB.Recordset
  '--------------------------------
    
  'USER DOESN'T ENTERED MORE THAN ONE DECIMAL:4-PCS
  If FLEX.COL = 4 Or FLEX.COL = 5 Or FLEX.COL = 6 Then
    If InStr(1, FLEX.TextMatrix(FLEX.ROW, FLEX.COL), ".") > 0 And KeyAscii = 46 Then
      KeyAscii = 0
      Exit Sub
    End If
  End If
  '--------------------------------------------
  
  'NO IDEA
  If Emptycell = True And (Not KeyAscii = 13) Then
     FLEX.TextMatrix(FLEX.ROW, FLEX.COL) = Empty
     Emptycell = False
  End If
  
  '----------------------------------------------
  'COLUMN WISE ENTERED TEXT CHECKING : ALLOW KEY
  '-----------------------------------------------
  Select Case FLEX.COL
  Case 2
    NEW_VISIBLE = True
    M_DESC = Empty
    Key = Empty
    FLEX.TextMatrix(FLEX.ROW, 2) = SearchList1("select TOP 20 code,name from itmmst", 0, FLEX.TextMatrix(FLEX.ROW, 2), "SELECT ITEM FROM LIST")
    FLEX.TextMatrix(FLEX.ROW, 9) = Key
    If key_PressNew = True Then
       M_DESC = ""
       Key = ""
       FLEX.TextMatrix(FLEX.ROW, 2) = ""
       frm_Item.Show
    End If
    ALLOW_KEY = True
    
   Case 1, 4, 5 'SIMPLE NUMBER WITHOUT DECIMAL
    If (KeyAscii >= 48 And KeyAscii <= 57) Then             ' 0- 9
      ALLOW_KEY = True
    Else
      ALLOW_KEY = False
    End If
   
   Case 6, 7 'DECIMAL NUMBER
    If KeyAscii >= 48 And KeyAscii <= 57 Then             ' 0- 9
      ALLOW_KEY = True
    ElseIf KeyAscii = 46 Then                              '.
      ALLOW_KEY = True
    Else
      ALLOW_KEY = False
    End If
   Case 7      'AMOUNT
    ALLOW_KEY = False
    
   Case 3  'MAY BE CHARACTER OR DECIMAL
   If KeyAscii >= 48 And KeyAscii <= 57 Then             ' 0- 9
      ALLOW_KEY = True
    ElseIf KeyAscii >= 65 And KeyAscii <= 90 Then         ' A-Z
      ALLOW_KEY = True
    ElseIf KeyAscii >= 97 And KeyAscii <= 122 Then         'a-z
      ALLOW_KEY = True
    ElseIf KeyAscii = 46 Then                              '.
      ALLOW_KEY = True
    ElseIf KeyAscii = 47 Then                              '/
      ALLOW_KEY = True
    ElseIf KeyAscii = 45 Then
      ALLOW_KEY = True
    Else
      ALLOW_KEY = False
    End If
  End Select
  
  '-------------------------------------------------------------------------------------------------
  'ENTER PRESS
  '-------------------------------------------------------------------------------------------------
  If KeyAscii = vbKeyReturn Then
    ENTER_PRESS = True
  Else
    ENTER_PRESS = False
  End If
  
  '-------------------------------------------------------------------------------------------------
  'BACK SPACE : COMES FIRST THEN KEYASCII = 0
  '-------------------------------------------------------------------------------------------------
  If KeyAscii = 8 Then
    Dim lnth As Double
    lnth = Len(FLEX.TextMatrix(FLEX.ROW, FLEX.COL))
    If lnth > 0 Then
      FLEX.TextMatrix(FLEX.ROW, FLEX.COL) = Mid(FLEX.TextMatrix(FLEX.ROW, FLEX.COL), 1, lnth - 1)
      Exit Sub
    End If
  End If
  
  '-------------------------------------------------------------------------------------------------
  'RESULT OF ALLOW KEY AND ENTER PRESS
  '-------------------------------------------------------------------------------------------------
  If ENTER_PRESS = False Then
     If ALLOW_KEY = True Then
        FLEX.TextMatrix(FLEX.ROW, FLEX.COL) = Trim(FLEX.TextMatrix(FLEX.ROW, FLEX.COL)) + Chr(KeyAscii)
     Else
        KeyAscii = 0
        Exit Sub
     End If
  End If
    
  '=================================================================================================
  'FORWARD MOVE FROM ONE COLUMN TO ANOTHER : IS TRUE OR FALSE ?? ON BASIS OF ENTER PRESS
  '=================================================================================================
   
  FWD_COL = False
  
  If ENTER_PRESS = True Then '-------------------------MAIN
    Select Case FLEX.COL
    Case 4, 5, 6, 7, 8
         If IsNumeric(FLEX.TextMatrix(FLEX.ROW, FLEX.COL)) Then
            FWD_COL = True
         End If
    Case 2, 3
            FWD_COL = True
    Case 1
         If Len(FLEX.TextMatrix(FLEX.ROW, FLEX.COL)) = 10 Then
            FWD_COL = True
         End If
    End Select
  
  '-----------------------------------------------SUB
  ' RESULT OF FORWARD COLUMN
  '-----------------------------------------------SUB
  
  '-------------------------------------------------
  'CHECKING FOR MERGE NO. AND SPECIFICATION
  
  '1. FOR MERGE NO. REQUIRED OR NOT
    If FLEX.COL = 3 Then
        Call MrgnReq
        If MRGN = "Y" Then
           If FLEX.TextMatrix(FLEX.ROW, 3) = Empty Then
           MsgBox "Lot No. Empty", vbOKOnly
           FLEX.ROW = FLEX.ROW
           FLEX.COL = FLEX.COL
           FLEX.SetFocus
           Exit Sub
        End If
      End If
      End If
      
 '2 . IF SPECIFICATION BOX + QUANTITY
     If FLEX.COL = 5 Then
        SPECI = GetSpeci(FLEX.TextMatrix(FLEX.ROW, 9))
        If SPECI = 0 Then
           If Val(FLEX.TextMatrix(FLEX.ROW, 5)) <= 0 And Val(FLEX.TextMatrix(FLEX.ROW, 6)) <= 0 Then
              MsgBox "BOX/Pcs. Empty ", vbOKOnly
              FLEX.ROW = FLEX.ROW
              FLEX.COL = FLEX.COL
              FLEX.SetFocus
              Exit Sub
          End If
        End If
        End If
        
  '--------------------------------------------------------------------------------
  '3 IF SPECIFICATION COPS + QUANTITY
  ' I.)
  If FLEX.COL = 4 Then
          SPECI = GetSpeci(FLEX.TextMatrix(FLEX.ROW, 9))
        If SPECI = 3 Then
           If Val(FLEX.TextMatrix(FLEX.ROW, 4)) <= 0 Then
              MsgBox "Please Enter Cops", vbOKOnly
              FLEX.ROW = FLEX.ROW
              FLEX.COL = FLEX.COL
              FLEX.SetFocus
              Exit Sub
          End If
          End If
          End If
          
  'II.)
  If FLEX.COL = 6 Then
          SPECI = GetSpeci(FLEX.TextMatrix(FLEX.ROW, 9))
        If SPECI = 3 Then
           If Val(FLEX.TextMatrix(FLEX.ROW, 6)) <= 0 Then
              MsgBox "Please Enter Quantity ", vbOKOnly
              FLEX.ROW = FLEX.ROW
              FLEX.COL = FLEX.COL
              FLEX.SetFocus
              Exit Sub
          End If
          End If
          End If
  '-----------------------------------------------------------
  ' SPECIFICATION ON QUANTITY
  If FLEX.COL = 6 Then
          SPECI = GetSpeci(FLEX.TextMatrix(FLEX.ROW, 9))
        If SPECI = 1 Then
           If Val(FLEX.TextMatrix(FLEX.ROW, 6)) <= 0 Then
              MsgBox "Please Enter Quantity ", vbOKOnly
              FLEX.ROW = FLEX.ROW
              FLEX.COL = FLEX.COL
              FLEX.SetFocus
              Exit Sub
          End If
          End If
          End If
        
'-------------------------------------------------
  
  If FWD_COL = True Then
     If FLEX.COL = 8 Then
        Dim AYS
        AYS = MsgBox("Want to Add More Item ", vbYesNo)
        If AYS = vbYes Then
          FLEX.Rows = FLEX.Rows + 1
          FLEX.ROW = FLEX.Rows - 1
          FLEX.COL = 1
          If FLEX.ROW > 1 Then
             FLEX.TextMatrix(FLEX.ROW, 1) = FLEX.TextMatrix(FLEX.ROW - 1, 1)
          End If
          FLEX.SetFocus
         Else
          If TXTSAMT.Enabled Then TXTSAMT.SetFocus Else SendKeys "{TAB}"
         End If
        Exit Sub
     ElseIf FLEX.COL = 7 Then
          FLEX.TextMatrix(FLEX.ROW, 8) = Format(Val(FLEX.TextMatrix(FLEX.ROW, 7)) * Val(FLEX.TextMatrix(FLEX.ROW, 6)), "#########.00")
          FLEX.COL = FLEX.COL + 1
     Else
          FLEX.COL = FLEX.COL + 1
     End If
  End If
  '-------------------------------------------------------SUB
  
  Emptycell = True
  End If
  
  '-------------------------------------------------------MAIN
  Exit Sub
LAST:
  MsgBox "Error In Item Detail"
  FLEX.SetFocus
  Exit Sub
End Sub

Private Sub CHKFLEX()
  CHK_FLX = True
  Dim MSTDAT As New ADODB.Recordset
  Set MSTDAT = New ADODB.Recordset
  Dim chkitm As String
  Dim FLXR As Double
  
  For FLXR = 1 To FLEX.Rows - 1
    chkitm = FLEX.TextMatrix(FLXR, 9)
    
    If MSTDAT.State = 1 Then MSTDAT.Close
    MSTDAT.Open "SELECT * FROM ITMMST WHERE CODE='" & chkitm & "'", CN, adOpenDynamic, adLockOptimistic
    If MSTDAT.EOF Then
       CHK_FLX = False
       FLXROW = FLXR
       FLXCOL = 2
       Exit For
    End If
        
    If Len(Trim(FLEX.TextMatrix(FLXR, 1))) <> 10 Then  'length
       CHK_FLX = False
       FLXROW = FLXR
       FLXCOL = 1
       Exit For
    ElseIf Not ValidRecNo(Trim(FLEX.TextMatrix(FLXR, 1))) Then   'valid
       CHK_FLX = False
       FLXROW = FLXR
       FLXCOL = 1
       Exit For
    End If
    
    If Not IsNumeric(FLEX.TextMatrix(FLXR, 4)) Then
       CHK_FLX = False
       FLXROW = FLXR
       FLXCOL = 4
       Exit For
    End If
    
    If Not IsNumeric(FLEX.TextMatrix(FLXR, 5)) Then
       CHK_FLX = False
       FLXROW = FLXR
       FLXCOL = 5
       Exit For
    End If
    If Not IsNumeric(FLEX.TextMatrix(FLXR, 6)) Then
       CHK_FLX = False
       FLXROW = FLXR
       FLXCOL = 6
       Exit For
    End If
    If Not IsNumeric(FLEX.TextMatrix(FLXR, 7)) Then
       CHK_FLX = False
       FLXROW = FLXR
       FLXCOL = 7
       Exit For
    End If
  Next
End Sub

Private Sub UPDATESTATUS()
  Dim DLYSTA As New ADODB.Recordset
  If DLYSTA.State = adStateOpen Then DLYSTA.Close
  DLYSTA.Open "SELECT * FROM DAILYSTAT WHERE CUSR='JNMOPAWCGBDSXZAS'", CN, adOpenKeyset, adLockPessimistic
  DLYSTA.AddNew
  DLYSTA!COMP = compPth & ""
  DLYSTA!VTYP = "IVR"
  DLYSTA!PCOD = TXTDBAC
  DLYSTA!dbcd = M_DBCD
  DLYSTA!QNTY = Val(TXTTQTY)
  DLYSTA!VBNO = TXTVBNO & ""
  DLYSTA!AMNT = Val(TXTITOT)
  DLYSTA!CUSR = cUName
  If SAVEFLAG = True Then
    DLYSTA!ACTN = "E"
   Else
    DLYSTA!ACTN = "M"
  End If
  DLYSTA!DTTM = Format(Now, "YYYY/MM/DD HH:MM:SS AMPM")
  DLYSTA.Update
  DLYSTA.Close
End Sub

Private Sub UPDATEDELSTATUS()
  Dim DLYSTA As New ADODB.Recordset
  If DLYSTA.State = adStateOpen Then DLYSTA.Close
  DLYSTA.Open "SELECT * FROM DAILYSTAT WHERE CUSR='JNMOPAWCGBDSXZAS'", CN, adOpenKeyset, adLockPessimistic
  DLYSTA.AddNew
  DLYSTA!COMP = compPth & ""
  DLYSTA!VTYP = "IVR"
  DLYSTA!PCOD = TXTDBAC
  DLYSTA!dbcd = M_DBCD
  DLYSTA!QNTY = Val(TXTTQTY)
  DLYSTA!VBNO = TXTVBNO & ""
  'DLYSTA!AMNT = Val(TXTBNET)
  DLYSTA!CUSR = cUName
  DLYSTA!ACTN = "D"
  DLYSTA!DTTM = Format(Now, "YYYY/MM/DD HH:MM:SS AMPM")
  DLYSTA.Update
  DLYSTA.Close
End Sub

Private Sub TXTVHCL_LostFocus()
TXTVHCL.BackColor = vbWhite
End Sub

Private Sub SetItemRate(ICOD As String, RATE As Double, QTY As Double)
Dim L As Long
Dim STKQNTY As Double
Dim WGTRATE As Double
Dim WGTRS As ADODB.Recordset
Set WGTRS = New ADODB.Recordset

If WGTRS.State = 1 Then WGTRS.Close
WGTRS.Open "SELECT WEIGHTEDRATE,BALQ,LAST_PURDATE FROM ITMMST WHERE CODE = '" & ICOD & "'", CN, adOpenDynamic, adLockOptimistic
If Not WGTRS.EOF Then
   WGTRATE = Val(WGTRS!WEIGHTEDRATE)
   STKQNTY = Val(WGTRS!BALQ)
End If

WGTRATE = ((STKQNTY * WGTRATE) + (QTY * RATE)) / (STKQNTY + QTY)

If Trim(WGTRS!LAST_PURDATE) <> Format(Now, "DD/MM/YYYY") And Not SAVEFLAG Then
   CN.Execute "UPDATE ITMMST SET WEIGHTEDRATE = " & WGTRATE & " WHERE CODE = '" & ICOD & "'", L
   Exit Sub
End If
'22/05/2010
CN.Execute "UPDATE ITMMST SET WEIGHTEDRATE = " & WGTRATE & " ,PURR = " & RATE & ",LAST_PURDATE = '" & Format(TXTVBDT.Value, "YYYY/MM/DD") & "' WHERE CODE = '" & ICOD & "'", L
  
End Sub

Private Sub ReSetRate(ICOD As String, RATE As Double, QTY As Double)
Dim STKQNTY As Double
Dim WGTRATE As Double
Dim WGTRS As ADODB.Recordset
Set WGTRS = New ADODB.Recordset

If WGTRS.State = 1 Then WGTRS.Close
WGTRS.Open "SELECT WEIGHTEDRATE,BALQ FROM ITMMST WHERE CODE = '" & ICOD & "'", CN, adOpenDynamic, adLockOptimistic
If Not WGTRS.EOF Then
   WGTRATE = Val(WGTRS!WEIGHTEDRATE)
   STKQNTY = Val(WGTRS!BALQ)
End If
WGTRS.Close

If (STKQNTY - QTY) <> 0 Then
  WGTRATE = ((STKQNTY * WGTRATE) - (QTY * RATE)) / (STKQNTY - QTY)
Else
  WGTRATE = ((STKQNTY * WGTRATE) - (QTY * RATE)) / 1
End If

CN.Execute "UPDATE ITMMST SET WEIGHTEDRATE = " & WGTRATE & " WHERE CODE = '" & ICOD & "'"
  
End Sub

Private Function ValidRecNo(ISSNO As String) As Boolean
ValidRecNo = True

On Error GoTo ERRREC
Dim CHKRS As ADODB.Recordset
Set CHKRS = New ADODB.Recordset
If CHKRS.State = 1 Then CHKRS.Close
CHKRS.Open "SELECT CLRSTATUS FROM JOBOUT WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' " & _
" AND VBNO='" & ISSNO & "' AND RECSTAT='A'", CN, adOpenDynamic, adLockOptimistic
If CHKRS.EOF Then
   MsgBox FLEX.TextMatrix(0, 1) & " doesn't exist"
   ValidRecNo = False
Else
   If Trim(CHKRS!CLRSTATUS & "") = "Y" Then
     MsgBox FLEX.TextMatrix(0, 1) & " has been cleared."
     ValidRecNo = False
   End If
End If

Exit Function
ERRREC:
MsgBox ERR.Description
End Function

Private Sub SetItemInfo()
On Error GoTo LAST
Dim INDEX As Long
Dim SQL As String
Dim RATE As Double

With FLEX
 
For INDEX = 1 To .Rows - 1
    
    SQL = "INSERT INTO GRNTRAN([COMP],[UNIT],[VTYP],[VBNO],[DBCD],[SRCH],DATE,[ICOD],[RATE],[GRN_QNTY],[NETRATE],[BAL_QNTY],[MRGN])"
    SQL = SQL & " VALUES('" & compPth & "','" & UNCD & "','IVR','" & TXTVBNO & _
    "','" & M_DBCD & "','" & INDEX & "','" & Format(TXTVBDT, "yyyy-MM-dd hh:mm:ss") & _
    "','" & Trim(.TextMatrix(INDEX, 9)) & "','" & Val(.TextMatrix(INDEX, 7)) & "','" & Val(.TextMatrix(INDEX, 6)) & _
    "','" & Val(.TextMatrix(INDEX, 7)) & "','" & Val(.TextMatrix(INDEX, 6)) & "','" & Trim(.TextMatrix(INDEX, 3)) & "')"
  
CN.Execute SQL
  
Next INDEX
 
End With
Exit Sub
LAST:
MsgBox ERR.Description
End Sub


Private Function GetSpeci(ICOD) As String
GetSpeci = ""
Dim SPECI As String

Dim IGCOD As String
Dim GETRS As ADODB.Recordset
Set GETRS = New ADODB.Recordset
Dim SPRS As ADODB.Recordset
Set SPRS = New ADODB.Recordset
If GETRS.State = 1 Then GETRS.Close
GETRS.Open "SELECT * FROM ITMMST WHERE CODE = '" & ICOD & "'", CN, adOpenDynamic, adLockOptimistic
If Not GETRS.EOF Then
   IGCOD = GETRS!igcd
End If

If SPRS.State = 1 Then SPRS.Close
SPRS.Open "SELECT * FROM IGMMST WHERE CODE = '" & Trim(IGCOD) & "'", CN, adOpenDynamic, adLockOptimistic
If Not SPRS.EOF Then
MRGN = SPRS!MERGE
GetSpeci = SPRS!SPECIFICATION
End If
End Function

Private Sub MrgnReq()
Dim SPECI As String

Dim IGCOD As String
Dim GETRS As ADODB.Recordset
Set GETRS = New ADODB.Recordset
Dim SPRS As ADODB.Recordset
Set SPRS = New ADODB.Recordset
If GETRS.State = 1 Then GETRS.Close
Dim i As Long

GETRS.Open "SELECT * FROM ITMMST WHERE CODE = '" & FLEX.TextMatrix(FLEX.ROW, 9) & "'", CN, adOpenDynamic, adLockOptimistic
If Not GETRS.EOF Then
   IGCOD = GETRS!igcd
End If

If SPRS.State = 1 Then SPRS.Close
SPRS.Open "SELECT * FROM IGMMST WHERE CODE = '" & Trim(IGCOD) & "'", CN, adOpenDynamic, adLockOptimistic
If Not SPRS.EOF Then
MRGN = SPRS!MERGE
End If
End Sub

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
        TXTBNET.Text = Format(FormatNumber(Val(TXTSAMT.Text) + Val(TXTITOT.Text) + Val(TXTADLS.Text), 0), "##########.00")
    Else
        TXTBNET.Text = Format(Val(TXTITOT.Text) + Val(TXTSAMT.Text) + Val(TXTADLS.Text), "##########.00")
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
                        flexBTRM.TextMatrix(J, 2) = Format((Val(flexBTRM.TextMatrix(J, 1))), "##########.000")
                    Else
                        flexBTRM.TextMatrix(J, 2) = Format(flexBTRM.TextMatrix(J, 2), "#.000")
                    End If
                Case "M_TPCS"
                    If Val(flexBTRM.TextMatrix(J, 1)) <> 0 Then
                        flexBTRM.TextMatrix(J, 2) = Format((Val(flexBTRM.TextMatrix(J, 1))), "##########.000")
                    Else
                        flexBTRM.TextMatrix(J, 2) = Format(flexBTRM.TextMatrix(J, 2), "#.000")
                    End If
                Case "M_SAMT"
                    If Val(flexBTRM.TextMatrix(J, 1)) <> 0 Then
                        flexBTRM.TextMatrix(J, 2) = Format((Val(flexBTRM.TextMatrix(J, 1)) * Val(TXTSAMT.Text)) / 100, "##########.000")
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
            c_FMLA(J) = Replace(c_FMLA(J), "M_TQTY", 0)
            c_FMLA(J) = Replace(c_FMLA(J), "M_TPCS", 0)
            c_FMLA(J) = Replace(c_FMLA(J), "M_SAMT", Val(TXTSAMT.Text))
            
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
        TXTBNET.Text = Val(TXTITOT.Text) + Val(TXTITOT.Text) + subTot
    Next J

    
End Sub

Private Sub EditKeyCode(MSHFlexGrid As Control, Edt As Control, KeyCode As Integer, Shift As Integer)
   
    Dim ANS As String
    chgFlag = True
    'Standard edit control processing.
   Select Case KeyCode
    
   Case 27   ' ESC: hide, return focus to MSHFlexGrid.
      Edt.Visible = False
      MSHFlexGrid.SetFocus
    
   Case 9    ' TAB return focus to mshflexgrid.
        'If Flex.COL - 1 <> 7 And Flex.COL - 1 <> 0 Then Flex.TextMatrix(Flex.ROW, Flex.COL - 1) = 0
   Case 13    ' ENTER return focus to MSHFlexGrid.
         MSHFlexGrid.SetFocus
         If MSHFlexGrid.COL = 2 Then
            If MSHFlexGrid.ROW < MSHFlexGrid.Rows - 1 Then
               MSHFlexGrid.ROW = MSHFlexGrid.ROW + 1
               MSHFlexGrid.COL = 1
            End If
         Else
            MSHFlexGrid.COL = MSHFlexGrid.COL + 1
        End If
   Case 38      ' Up.
      MSHFlexGrid.SetFocus
      DoEvents
      If MSHFlexGrid.ROW > MSHFlexGrid.FixedRows Then
         MSHFlexGrid.ROW = MSHFlexGrid.ROW - 1
      End If

   Case 40      ' Down.
      MSHFlexGrid.SetFocus
      DoEvents
      If MSHFlexGrid.ROW < MSHFlexGrid.Rows - 1 Then
         MSHFlexGrid.ROW = MSHFlexGrid.ROW + 1
      End If
   End Select
   chgFlag = False
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
  RS.Open "SELECT * FROM CONFIG WHERE COMP='" & compPth & "' and vtyp='IVR' AND DBCD='" & M_DBCD & "' " & _
          "AND UNIT='" & UNCD & "' ORDER BY SRCH", CN, adOpenKeyset, adLockPessimistic
  CNTR = 0
  Do While Not RS.EOF
   flexBTRM.Rows = flexBTRM.Rows + 1
   flexBTRM.TextMatrix(CNTR, 0) = RS!NICK & ""
   flexBTRM.ColWidth(0) = 1450
   flexBTRM.ColWidth(1) = 850
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
    'PLUS
    M_FMLA(CNTR) = Replace(M_FMLA(CNTR), "SERVICE AMT.", "M_SAMT")
    
    
    If M_NICK(0) <> "" Then
        If M_OPER(0) = "+" Then
          M_FMLA(CNTR) = Replace(M_FMLA(CNTR), M_NICK(0), "AMT_01 ")
         Else
          M_FMLA(CNTR) = Replace(M_FMLA(CNTR), M_NICK(0), " -AMT_01")
        End If
    End If
    If M_NICK(1) <> "" Then
        If M_OPER(1) = "+" Then
          M_FMLA(CNTR) = Replace(M_FMLA(CNTR), M_NICK(1), " +AMT_02")
         Else
          M_FMLA(CNTR) = Replace(M_FMLA(CNTR), M_NICK(1), " -AMT_02")
        End If
    End If
    If M_NICK(2) <> "" Then
        If M_OPER(2) = "+" Then
          M_FMLA(CNTR) = Replace(M_FMLA(CNTR), M_NICK(2), " +AMT_03")
         Else
          M_FMLA(CNTR) = Replace(M_FMLA(CNTR), M_NICK(2), " -AMT_03")
        End If
    End If
    
    If M_NICK(3) <> "" Then
        If M_OPER(3) = "+" Then
          M_FMLA(CNTR) = Replace(M_FMLA(CNTR), M_NICK(3), " +AMT_04")
         Else
          M_FMLA(CNTR) = Replace(M_FMLA(CNTR), M_NICK(3), " -AMT_04")
        End If
    End If
    
    If M_NICK(4) <> "" Then
        If M_OPER(4) = "+" Then
          M_FMLA(CNTR) = Replace(M_FMLA(CNTR), M_NICK(4), " +AMT_05")
         Else
          M_FMLA(CNTR) = Replace(M_FMLA(CNTR), M_NICK(4), " -AMT_05")
        End If
    End If
    
    If M_NICK(5) <> "" Then
        If M_OPER(5) = "+" Then
          M_FMLA(CNTR) = Replace(M_FMLA(CNTR), M_NICK(5), " +AMT_06")
        Else
          M_FMLA(CNTR) = Replace(M_FMLA(CNTR), M_NICK(5), " -AMT_06")
        End If
    End If
    
    If M_NICK(6) <> "" Then
        If M_OPER(6) = "+" Then
          M_FMLA(CNTR) = Replace(M_FMLA(CNTR), M_NICK(6), " +AMT_07")
        Else
          M_FMLA(CNTR) = Replace(M_FMLA(CNTR), M_NICK(6), " -AMT_07")
        End If
    End If
    
    If M_NICK(7) <> "" Then
        If M_OPER(7) = "+" Then
          M_FMLA(CNTR) = Replace(M_FMLA(CNTR), M_NICK(7), " +AMT_08")
        Else
          M_FMLA(CNTR) = Replace(M_FMLA(CNTR), M_NICK(7), " -AMT_08")
        End If
    End If
    
    If M_NICK(8) <> "" Then
        If M_OPER(8) = "+" Then
          M_FMLA(CNTR) = Replace(M_FMLA(CNTR), M_NICK(8), " +AMT_09")
        Else
          M_FMLA(CNTR) = Replace(M_FMLA(CNTR), M_NICK(8), " -AMT_09")
        End If
    End If
  Next
  If flexBTRM.Rows > 0 Then
    'O.k
   Else
    flexBTRM.Enabled = False
  End If
End Sub

Private Sub MSHFlexGridEdit(MSHFlexGrid As Control, Edt As Control, KeyAscii As Integer)
    chgFlag = True
    ' Use the character that was typed.
   Select Case KeyAscii
   ' A space means edit the current text.
   Case 0 To 12
      Edt = MSHFlexGrid
      Edt.SelStart = 1000
   Case 14 To 26
      Edt = MSHFlexGrid
      Edt.SelStart = 1000
   Case 13
      If MSHFlexGrid.COL = 2 Then
            If MSHFlexGrid.Rows <> MSHFlexGrid.ROW + 1 Then
                MSHFlexGrid.ROW = MSHFlexGrid.ROW + 1
            Else
                cmdSave.SetFocus
            End If
            MSHFlexGrid.COL = 1
            Exit Sub
        Else
            
            MSHFlexGrid.COL = MSHFlexGrid.COL + 1
            Exit Sub
      End If
   Case 28 To 32
      Edt = MSHFlexGrid
      Edt.SelStart = 1000
   Case 27
        Edt.Text = Empty
        Exit Sub
   ' Anything else means replace the current text.
   Case Else
      Edt = Chr(KeyAscii)
      Edt.SelStart = 1
   End Select

   ' Show Edt at the right place.
   Edt.MOVE MSHFlexGrid.Left + MSHFlexGrid.CellLeft, _
      MSHFlexGrid.Top + MSHFlexGrid.CellTop, _
      MSHFlexGrid.CellWidth - 8, _
      MSHFlexGrid.CellHeight - 8
   Edt.Visible = True

   ' And make it work.
   Edt.SetFocus
   chgFlag = False
End Sub

Private Sub flexBTRM_DblClick()
    calbtm = True
    MSHFlexGridEdit flexBTRM, txtBEdit, 32 ' Simulate a space.
End Sub

Private Sub flexBTRM_GotFocus()
    Me.KeyPreview = False
    Msg "Billing Terms"
    If flexBTRM.Rows > 0 Then
      flexBTRM.COL = 1
      flexBTRM.TopRow = 0
      flexBTRM.LeftCol = 1
     Else
      TXTBNET = Val(TXTITOT) + Val(TXTSAMT)
    End If
End Sub
Private Sub flexBTRM_KeyPress(KeyAscii As Integer)
    If flexBTRM.COL = 2 And flexBTRM.ROW + 1 = flexBTRM.Rows Then calbtm = False Else calbtm = True
    MSHFlexGridEdit flexBTRM, txtBEdit, KeyAscii
    If KeyAscii = vbKeyReturn Then
      If flexBTRM.ROW Mod 4 = 0 And flexBTRM.COL = 2 And flexBTRM.ROW > 0 Then
         'SendKeys "{Down}"
         flexBTRM.TopRow = flexBTRM.ROW - 1
         'flexBTRM.Col = 1
      End If
    End If
End Sub

Private Sub flexBTRM_EnterCell()
If flexBTRM.COL <> 0 Then flexBTRM.CellBackColor = RGB(BRED, BGREEN, BBLUE)
End Sub


Private Sub flexBTRM_LeaveCell()
If flexBTRM.COL <> 0 Then flexBTRM.CellBackColor = vbWhite
    If txtBEdit.Visible = False Then Exit Sub
    flexBTRM = txtBEdit
    txtBEdit.Visible = False
End Sub

Private Sub flexBTRM_RowColChange()
    If flexBTRM.COL = 1 Then
        If calbtm = True Then
            calBTRM 0
        End If
    End If
    If flexBTRM.Rows > 7 Then
        If flexBTRM.ROW Mod 5 = 0 And flexBTRM.ROW <> 0 Then
            flexBTRM.TopRow = 5
        End If
    End If
    calADLS
End Sub
