VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WELCHB~1.OCX"
Begin VB.Form FrmGRNEntry 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GRN Entry"
   ClientHeight    =   7245
   ClientLeft      =   375
   ClientTop       =   1110
   ClientWidth     =   11385
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   11385
   Begin VB.Frame frm_head 
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   120
      TabIndex        =   42
      Top             =   -120
      Width           =   11175
      Begin VB.TextBox txtdty 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   9840
         MaxLength       =   7
         TabIndex        =   19
         Top             =   1680
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.TextBox txtfrt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   7320
         MaxLength       =   7
         TabIndex        =   18
         Top             =   1680
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.TextBox txtcha 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   4560
         MaxLength       =   7
         TabIndex        =   17
         Top             =   1680
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.TextBox txtEXRate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   2760
         MaxLength       =   7
         TabIndex        =   16
         Top             =   1680
         Width           =   1125
      End
      Begin VB.TextBox txtCURNCY 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1200
         MaxLength       =   7
         TabIndex        =   15
         Top             =   1680
         Width           =   525
      End
      Begin VB.CheckBox chkAcEffect 
         Caption         =   "GRN With Non A/c Effect"
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
         Left            =   8280
         TabIndex        =   12
         Top             =   960
         Width           =   3495
      End
      Begin VB.TextBox TXTSBILLNO 
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   10
         Top             =   960
         Width           =   1815
      End
      Begin VB.ComboBox cmbSelection 
         BackColor       =   &H00C0C0FF&
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
         ItemData        =   "FrmGRNEntry.frx":0000
         Left            =   9360
         List            =   "FrmGRNEntry.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Tag             =   "0"
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox TXTRTORTAX 
         Height          =   285
         Left            =   5880
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1320
         Width           =   2295
      End
      Begin VB.TextBox TXTTAXNAM 
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   1320
         Width           =   2655
      End
      Begin VB.TextBox TXTSCHLN 
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   5
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox TXTVBNO 
         Height          =   285
         Left            =   9360
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox TXTDBAC 
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   6495
      End
      Begin MSComCtl2.DTPicker TXTSCHLNDT 
         Height          =   285
         Left            =   5880
         TabIndex        =   7
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
         Format          =   51838977
         CurrentDate     =   39347
      End
      Begin MSComCtl2.DTPicker TXTVBDT 
         Height          =   315
         Left            =   9360
         TabIndex        =   9
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
         Format          =   51838977
         CurrentDate     =   39347
      End
      Begin MSComCtl2.DTPicker TXTSBILLDATE 
         Height          =   285
         Left            =   5880
         TabIndex        =   11
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
         Format          =   51838977
         CurrentDate     =   39347
      End
      Begin VB.Label lblnote 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Note: Please Exclude the modvat Amount from duty and taxes "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   3960
         TabIndex        =   74
         Top             =   1995
         Visible         =   0   'False
         Width           =   7005
      End
      Begin VB.Label lbldty 
         BackStyle       =   0  'Transparent
         Caption         =   "Duty && Taxes :"
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
         Left            =   8520
         TabIndex        =   73
         Tag             =   "S"
         Top             =   1725
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label lblfrt 
         BackStyle       =   0  'Transparent
         Caption         =   "Freight && Insurance :"
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
         Left            =   5640
         TabIndex        =   72
         Tag             =   "S"
         Top             =   1725
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Shape ShapeCHA 
         BorderColor     =   &H80000002&
         Height          =   300
         Left            =   3960
         Shape           =   4  'Rounded Rectangle
         Top             =   1680
         Visible         =   0   'False
         Width           =   6975
      End
      Begin VB.Label lblcha 
         BackStyle       =   0  'Transparent
         Caption         =   "CHA :"
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
         TabIndex        =   71
         Tag             =   "S"
         Top             =   1725
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label LBLEXRATE 
         BackStyle       =   0  'Transparent
         Caption         =   "Ex. Rate :"
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
         Left            =   1800
         TabIndex        =   22
         Tag             =   "S"
         Top             =   1725
         Width           =   975
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Currency :"
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
         TabIndex        =   21
         Tag             =   "S"
         Top             =   1725
         Width           =   975
      End
      Begin VB.Shape Shape5 
         BorderColor     =   &H80000002&
         Height          =   300
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   1680
         Width           =   3855
      End
      Begin VB.Label Label8 
         Caption         =   "Supplier Bill No."
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
         TabIndex        =   69
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label7 
         Caption         =   "Bill Date :"
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
         TabIndex        =   68
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "TAX Type"
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
         TabIndex        =   67
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label LBLRTORTX 
         Caption         =   "Retail/Tax Inv."
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
         Left            =   4440
         TabIndex        =   66
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label LBLTAXNAM 
         Caption         =   "Tax Reference"
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
         TabIndex        =   65
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "Chln Date :"
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
         TabIndex        =   6
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Supplier Chln No."
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
         TabIndex        =   4
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label LBLBILLDATE 
         Caption         =   "GRN Date"
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
         TabIndex        =   24
         Top             =   600
         Width           =   975
      End
      Begin VB.Label LBLBILLNO 
         Caption         =   "GRN No."
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
         TabIndex        =   8
         Top             =   240
         Width           =   855
      End
      Begin VB.Label LBLDRAC 
         Caption         =   "Supplier Name"
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
         TabIndex        =   2
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame FRMBTRM 
      Height          =   2460
      Left            =   7440
      TabIndex        =   1
      Top             =   4560
      Width           =   3855
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
         Height          =   420
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   32
         Text            =   "0.00"
         Top             =   1960
         Width           =   1905
      End
      Begin VB.TextBox txtBEdit 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1440
         TabIndex        =   39
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
         Left            =   2040
         TabIndex        =   38
         Text            =   "0.00"
         Top             =   1320
         Visible         =   0   'False
         Width           =   1575
      End
      Begin MSFlexGridLib.MSFlexGrid flexBTRM 
         Height          =   1815
         Left            =   120
         TabIndex        =   45
         Top             =   120
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   3201
         _Version        =   393216
         Rows            =   0
         Cols            =   5
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
         TabIndex        =   31
         Top             =   2040
         Width           =   1305
      End
   End
   Begin VB.Frame ITMFRM 
      Height          =   2295
      Left            =   120
      TabIndex        =   40
      Top             =   2040
      Width           =   11175
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
         TabIndex        =   43
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
         Left            =   3480
         TabIndex        =   26
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
         Left            =   6120
         TabIndex        =   28
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox TXTITOT 
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
         Left            =   9000
         TabIndex        =   30
         Text            =   "0.00"
         Top             =   1920
         Width           =   1575
      End
      Begin MSFlexGridLib.MSFlexGrid FLEX 
         Height          =   1695
         Left            =   0
         TabIndex        =   44
         Top             =   120
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   2990
         _Version        =   393216
         Cols            =   24
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
      Begin VB.Label Label1 
         Caption         =   "Gross Amount "
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
         Left            =   7440
         TabIndex        =   29
         Top             =   1920
         Width           =   1575
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
         TabIndex        =   41
         Top             =   1920
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Total Carton"
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
         Left            =   2040
         TabIndex        =   25
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
         Left            =   4560
         TabIndex        =   27
         Top             =   1920
         Width           =   1455
      End
   End
   Begin WelchButton.lvButtons_H cmdAdd 
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   6600
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
      Image           =   "FrmGRNEntry.frx":0004
      cBack           =   -2147483633
   End
   Begin WelchButton.lvButtons_H cmdEdit 
      Height          =   495
      Left            =   5040
      TabIndex        =   36
      Top             =   6600
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
      Image           =   "FrmGRNEntry.frx":039E
      cBack           =   -2147483633
   End
   Begin WelchButton.lvButtons_H cmdDelete 
      Height          =   495
      Left            =   3840
      TabIndex        =   35
      Top             =   6600
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
      Image           =   "FrmGRNEntry.frx":0738
      cBack           =   -2147483633
   End
   Begin WelchButton.lvButtons_H cmdSave 
      Height          =   495
      Left            =   1440
      TabIndex        =   33
      Top             =   6600
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
      Image           =   "FrmGRNEntry.frx":0AD2
      cBack           =   -2147483633
   End
   Begin WelchButton.lvButtons_H cmdCancel 
      Height          =   495
      Left            =   2640
      TabIndex        =   34
      Top             =   6600
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
      Image           =   "FrmGRNEntry.frx":185C
      cBack           =   -2147483633
   End
   Begin WelchButton.lvButtons_H cmdExit 
      Height          =   495
      Left            =   6225
      TabIndex        =   37
      Top             =   6600
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
      Image           =   "FrmGRNEntry.frx":1CAE
      cBack           =   -2147483633
   End
   Begin TabDlg.SSTab FRMLRDTL 
      Height          =   2055
      Left            =   120
      TabIndex        =   70
      Top             =   4440
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   3625
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      ForeColor       =   128
      TabCaption(0)   =   "Transport Details"
      TabPicture(0)   =   "FrmGRNEntry.frx":2100
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "LBLRMRK"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "LBLVHCL"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "LBLTRNM"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "LBLLRDT"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "LBLLRNO"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label11"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "TXTLRDT"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "TXTRMRK"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "TXTVHCL"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "TXTTRNM"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "TXTLRNO"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "chkReturnable"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "TXTGDN"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).ControlCount=   13
      TabCaption(1)   =   "   Returnable Pallet/Cops"
      TabPicture(1)   =   "FrmGRNEntry.frx":211C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label9(0)"
      Tab(1).Control(1)=   "Label9(1)"
      Tab(1).Control(2)=   "Label9(2)"
      Tab(1).Control(3)=   "FLEXPLY"
      Tab(1).Control(4)=   "TxtPallet"
      Tab(1).Control(5)=   "txtCops"
      Tab(1).Control(6)=   "txtPly"
      Tab(1).ControlCount=   7
      Begin VB.TextBox TXTGDN 
         Height          =   285
         Left            =   960
         MaxLength       =   250
         TabIndex        =   54
         Top             =   1200
         Width           =   6135
      End
      Begin VB.CheckBox chkReturnable 
         Caption         =   "Returnable Pallet/Cops Entry Required"
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
         Left            =   2460
         TabIndex        =   57
         Top             =   40
         Width           =   3855
      End
      Begin VB.TextBox txtPly 
         Height          =   285
         Left            =   -69360
         MaxLength       =   20
         TabIndex        =   63
         Top             =   450
         Width           =   855
      End
      Begin VB.TextBox txtCops 
         Height          =   285
         Left            =   -71400
         MaxLength       =   20
         TabIndex        =   61
         Top             =   450
         Width           =   855
      End
      Begin VB.TextBox TxtPallet 
         Height          =   285
         Left            =   -73560
         MaxLength       =   20
         TabIndex        =   59
         Top             =   450
         Width           =   855
      End
      Begin VB.TextBox TXTLRNO 
         Height          =   285
         Left            =   960
         MaxLength       =   20
         TabIndex        =   47
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox TXTTRNM 
         Height          =   285
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   49
         Top             =   480
         Width           =   2775
      End
      Begin VB.TextBox TXTVHCL 
         Height          =   285
         Left            =   4320
         MaxLength       =   20
         TabIndex        =   53
         Top             =   840
         Width           =   2775
      End
      Begin VB.TextBox TXTRMRK 
         Height          =   285
         Left            =   960
         MaxLength       =   250
         TabIndex        =   55
         Top             =   1560
         Width           =   6135
      End
      Begin MSComCtl2.DTPicker TXTLRDT 
         Height          =   300
         Left            =   960
         TabIndex        =   51
         Top             =   840
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   529
         _Version        =   393216
         Format          =   51838977
         CurrentDate     =   39347
      End
      Begin MSFlexGridLib.MSFlexGrid FLEXPLY 
         Height          =   765
         Left            =   -74880
         TabIndex        =   64
         Top             =   960
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   1349
         _Version        =   393216
         Cols            =   5
         BackColor       =   -2147483628
         BackColorBkg    =   -2147483633
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
      Begin VB.Label Label11 
         Caption         =   "Godown"
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
         TabIndex        =   76
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "No. of Ply"
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
         Index           =   2
         Left            =   -70320
         TabIndex        =   62
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "No. of Cops"
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
         Index           =   1
         Left            =   -72600
         TabIndex        =   60
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "No. of Pallets"
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
         Index           =   0
         Left            =   -74760
         TabIndex        =   58
         Top             =   480
         Width           =   1215
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
         Left            =   120
         TabIndex        =   46
         Top             =   480
         Width           =   615
      End
      Begin VB.Label LBLLRDT 
         Caption         =   "L.R.Dt."
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
         TabIndex        =   50
         Top             =   840
         Width           =   855
      End
      Begin VB.Label LBLTRNM 
         Caption         =   "Name of Transport"
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
         Left            =   2520
         TabIndex        =   48
         Top             =   480
         Width           =   1695
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
         Left            =   2520
         TabIndex        =   52
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label LBLRMRK 
         Caption         =   "Remark "
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
         TabIndex        =   56
         Top             =   1560
         Width           =   735
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00000080&
         Height          =   2655
         Left            =   -72840
         Shape           =   4  'Rounded Rectangle
         Top             =   720
         Width           =   3975
      End
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Effect In Cost Y:Yes/N:No"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   8880
      TabIndex        =   75
      Top             =   4395
      Width           =   2295
   End
   Begin VB.Shape Shape1 
      Height          =   615
      Left            =   120
      Top             =   6525
      Width           =   7215
   End
End
Attribute VB_Name = "FrmGRNEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'FIFO
Public M_ISSUE As String
Public TABLENAME As String
Public SUMMARYTABLE As String
Public IVRBOK_DIRIVR As String
Public M_DBCD_DIRIVR As String
Public ALLOWEDITDEL As Boolean
Dim M_OPER(0 To 15) As String
Dim M_PERC(0 To 15) As Double
Dim M_POSTCOD(0 To 15) As String
Dim M_NICK(0 To 15) As String
Dim M_POSTYESNO(0 To 15) As String
Dim M_FMLA(0 To 15) As String
Dim M_RDOF(0 To 15) As String
Dim M_BILRDOF As String
Public SAVEFLAG As Boolean
Public MIN_DAT As Date
Dim calbtm As Boolean
Dim chgFlag As Boolean
Dim FRMOPER As String
Public EFFGRN As String
Dim sav_srfx  As String
Dim M_EXCISABLE As String
Dim CHK_FLX As Boolean
Dim FLXROW As Double
Dim FLXCOL As Double
Dim Emptycell As Boolean
Dim EXCISE As ADODB.Recordset
Dim MRGN As String
Dim SPECI As Double

Private Sub chkAcEffect_GotFocus()
  chkAcEffect.FontSize = 10
End Sub

Private Sub chkAcEffect_LostFocus()
 chkAcEffect.FontSize = 8
End Sub

Private Sub chkReturnable_Click()
If chkReturnable.Value = 1 Then
   FRMLRDTL.Tab = 1
   If TxtPallet.Enabled Then TxtPallet.SetFocus
Else
   FRMLRDTL.Tab = 0
   TxtPallet = Empty
   txtCops = Empty
   txtPly = Empty
End If
End Sub

Private Sub chkReturnable_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 And chkReturnable.Value = 0 Then
   cmdSave.SetFocus
ElseIf KeyCode = 32 Then
   Call chkReturnable_Click
End If
End Sub

Private Sub cmbSelection_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn And cmbSelection.ListIndex <> -1 Then
   FLEX.SetFocus
   FLEX.ROW = 1
   FLEX.COL = 3
End If
End Sub

Private Sub cmbSelection_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub Form_Activate()
 Call ColorComponent(Me)
 Me.BackColor = RGB(RED, GREEN, BLUE)
 TXTBNET.FontSize = 14
 TXTBNET.FontBold = True
  DIVCOD = Me.Tag
  Dim divisionmaster As New ADODB.Recordset
  Set divisionmaster = New ADODB.Recordset
  If divisionmaster.State = 1 Then divisionmaster.Close
  divisionmaster.Open "select * from DIVMST where code='" & DIVCOD & "' AND UNIT='" & UNCD & "' and comp='" & compPth & "'", CN
  If Not divisionmaster.EOF Then
    DIVNAM = divisionmaster!NAME
    DIVCOD = divisionmaster!CODE
   Else
    DIVNAM = "??????"
  End If
  If DIVNAM = "??????" Then
    Unload Me
  End If
  IVRBOK_DIRIVR = Me.Caption
  FRMPARA = "IVR"
  If DIVNAM = "??????" Or IVRBOK_DIRIVR = "??????" Or Trim(IVRBOK_DIRIVR) = "" Then
    Unload Me
  End If
  
  If cmbSelection.ListCount > 1 Then cmbSelection.ListIndex = 0
End Sub

Private Sub Form_Load()
  Call ColorComponent(Me)
  Call CenterChild(frm_Main, Me)
  'TEMP : FIFO
  FIFOREQ = "Y"
  '----------
  Me.Left = 130
  Emptycell = True
  flexBTRM.ColWidth(0) = 1300
  flexBTRM.ColWidth(1) = 600
  flexBTRM.ColWidth(2) = 900
  flexBTRM.ColWidth(3) = 400
  flexBTRM.ColWidth(4) = 0
  
  TXTSBILLDATE = Now
  TXTVBDT = Now
  TXTLRDT = Now
  TXTSCHLNDT = Now
  TXTVBDT.MinDate = FSDT
  TXTVBDT.MaxDate = FEDT
  
  If TABLENAME = "STORETRAN" Then
     M_DBCD_DIRIVR = "000001"
     Me.Caption = "GRN ENTRY FOR RAW MATERIAL"
  ElseIf TABLENAME = "JOBIN" Then
     M_DBCD_DIRIVR = "000002"
     Me.Caption = "GRN ENTRY FOR JOBWORK"
  End If
  
  FRMPARA = "IVR"
  Call setflexhead
  Call setHeading
  
  M_DESC = Empty
  Key = Empty
  NEW_VISIBLE = False
  DIVCOD = "000001"
  DIVNAM = Empty
  'If DIVCOD = Empty Then
  '  DIVNAM = SearchList1("SELECT  TOP 20 CODE,NAME FROM DIVMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND CODE='000001'", 0, "", "SELECT DIVISION MASTER")
  'End If
  DIVNAM = GETDIVNAME(DIVCOD)
  
  If RS.State = 1 Then RS.Close
  RS.Open "SELECT * FROM DIVMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND NAME='" & DIVNAM & "'", CN, adOpenKeyset, adLockPessimistic
  If Not RS.EOF Then
     DIVCOD = RS!CODE
     DIVNAM = RS!NAME
     Me.Tag = DIVCOD
  End If
      
  Call FIL_Billingterm
  Call btn_sts(True)
  
  cmbSelection.Clear
  cmbSelection.AddItem ("NONE")
  cmbSelection.AddItem ("RG23-A")
  cmbSelection.AddItem ("RG23-C")
  
  Call SetLastDate(Me, "TXTVBDT", "IVR", M_DBCD_DIRIVR)
  
  Me.KeyPreview = True
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If UCase(ActiveControl.NAME) = "TXTSCHLN" Then
     If TXTSCHLN = Empty Then Exit Sub
  End If
  If UCase(ActiveControl.NAME) = "FLEXPLY" Then Exit Sub
  If (ActiveControl.NAME = "TXTCRAC" Or ActiveControl.NAME = "TXTDBAC" Or ActiveControl.NAME = "TXTDLPTY" Or ActiveControl.NAME = "TXTBRNM" Or ActiveControl.NAME = "TXTTAXNAM" Or UCase(ActiveControl.NAME) = "CMBSELECTION" Or UCase(ActiveControl.NAME) = "FLEX") Then Exit Sub
  If UCase(ActiveControl.NAME) = "CMDSAVE" Then Exit Sub
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
  End If
End Sub

Public Sub cmdAdd_Click()
  FRMOPER = "*"
  SAVEFLAG = True
  btn_sts (False)
  cmddelitm.Enabled = False
  
  TXTVBNO = GenVNO("IVR", M_DBCD_DIRIVR)
  If TXTDBAC.Enabled = True Then
    TXTDBAC.SetFocus
  End If
  txtCURNCY = Trim(M_Currency)
End Sub

Private Sub cmdCancel_Click()
  ClsData (Me)
  FLEX.Clear
  FLEX.Rows = 2
  btn_sts (True)
  Call setflexhead
  cmdAdd.SetFocus
  Dim i As Integer
  For i = 0 To flexBTRM.Rows - 1
    flexBTRM.TextMatrix(i, 2) = "0.00"
  Next
  FLEXPLY.Clear
  Call setHeading
  
  Call SetLastDate(Me, "TXTVBDT", "IVR", M_DBCD_DIRIVR)
  
  TXTBNET.Text = "0.00"
  chkReturnable.Visible = True
  chkReturnable.Value = 0
  chkAcEffect.Value = 0
End Sub

Private Sub cmdDelete_Click()
    If M_USRSECLEVL = "1" Then
       If ReadConfigMaster("000049", 6, "M") = False Then ModuleDeniedMessage: Exit Sub
    End If

  If M_USRSECLEVL = "1" Then
     If FrmGRNEntry.TABLENAME = "STORETRAN" And FrmGRNEntry.SUMMARYTABLE = "GRN" Then
        If ReadConfigMaster("000035", 6, "M") = False Then ModuleDeniedMessage: Exit Sub
     End If
     If FrmGRNEntry.TABLENAME = "JOBIN" And FrmGRNEntry.SUMMARYTABLE = "JOBGRN" Then
        If ReadConfigMaster("000036", 6, "M") = False Then ModuleDeniedMessage: Exit Sub
     End If
  End If
  
  ALLOWEDITDEL = True
  SAVEFLAG = False
  btn_sts (False)
    
  FrmGRNEntryList1.Show 1
  FRMLRDTL.Tab = 0
  
  If ALLOWEDITDEL = False Then
    MsgBox "Purchase of this GRN have been made can not edit/delete ", vbInformation
   Else
    'Check for Receipt and Payment Entires
    If Not TXTVBNO = Empty Then
      Dim AYS
      AYS = MsgBox("Are you sure to delete the invoice ", vbYesNo)
      If AYS = vbYes Then
        CN.BeginTrans
        CN.Execute "UPDATE " & TABLENAME & " set recstat='D' WHERE COMP='" & compPth & "' AND UNIT ='" & UNCD & _
        "' AND VTYP='IVR' AND DBCD='" & M_DBCD_DIRIVR & "' AND VBNO='" & TXTVBNO & "'"
        
        CN.Execute "UPDATE " & SUMMARYTABLE & " SET RECSTAT='D' WHERE COMP='" & compPth & "' AND UNIT ='" & UNCD & _
        "' AND VTYP='IVR' AND DBCD='" & M_DBCD_DIRIVR & "' AND VBNO='" & TXTVBNO & "'"
        
        CN.Execute "UPDATE EGPMAN SET RECSTAT = 'D' WHERE COMP='" & compPth & "' AND UNIT ='" & UNCD & _
                   "' AND VTYP='IVR' AND LTRIM(RTRIM(VBNO))='" & TXTVBNO & "' AND DBCD='" & M_DBCD_DIRIVR & "'"
                   
        CN.Execute "UPDATE EGPMAN SET RECSTAT = 'D' WHERE  COMP='" & compPth & "' AND UNIT='" & UNCD & _
                   "' AND DBCD='RG23-C' AND VTYP='IVR' AND VBNO='" & TXTVBNO & "'"
        
        CN.Execute "UPDATE PKGSTK SET RECSTAT='D' WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
                   "' AND VTYP='IVR' AND DBCD='" & M_DBCD_DIRIVR & "' AND CHLN='" & Trim(TXTVBNO) & "'"
        
        Call DAILYSTATUS("IVR", GetCode("ACCMST", TXTDBAC, "NAME", "CODE"), M_DBCD_DIRIVR, Val(TXTTQTY), TXTVBNO, Val(TXTITOT), cUName, "D", Now, TXTVBDT)
        CN.CommitTrans
      End If
    End If
  End If
  Call cmdCancel_Click
End Sub

Private Sub cmddelitm_Click()
  If FLEX.ROW > 1 Then
    FLEX.RemoveItem (FLEX.ROW)
    TXTTPCS.Text = 0
    TXTTQTY.Text = 0
    TXTITOT.Text = 0
    Dim i As Double
    i = 1
    For i = 1 To FLEX.Rows - 1
      FLEX.TextMatrix(i, 0) = i
      TXTTPCS.Text = Format(Val(TXTTPCS.Text) + Val(FLEX.TextMatrix(i, 7)), "######")
      TXTTQTY.Text = Format(Val(TXTTQTY.Text) + Val(FLEX.TextMatrix(i, 8)), "########.000")
      TXTITOT.Text = Format(Val(TXTITOT.Text) + Val(FLEX.TextMatrix(i, 10)), "########.00")
    Next
    FLEX.Refresh
    FLEX.ROW = FLEX.Rows - 1
    FLEX.COL = 16
    FLEX.SetFocus
  End If
  cmddelitm.Enabled = False
  
End Sub

Private Sub cmddelitm_LostFocus()
  cmddelitm.Enabled = False
End Sub

Private Sub cmdEdit_Click()
   If M_USRSECLEVL = "1" Then
    If ReadConfigMaster("000049", 6, "M") = False Then ModuleDeniedMessage: Exit Sub
  End If

  If M_USRSECLEVL = "1" Then
  If FrmGRNEntry.TABLENAME = "STORETRAN" And FrmGRNEntry.SUMMARYTABLE = "GRN" Then
     If ReadConfigMaster("000035", 5, "M") = False Then ModuleDeniedMessage: Exit Sub
  End If
  If FrmGRNEntry.TABLENAME = "JOBIN" And FrmGRNEntry.SUMMARYTABLE = "JOBGRN" Then
     If ReadConfigMaster("000036", 5, "M") = False Then ModuleDeniedMessage: Exit Sub
  End If
  End If
  ALLOWEDITDEL = True
  SAVEFLAG = False
  TXTVBNO = Empty
  
  M_ISSUE = "N"
  
  TXTVBDT.MinDate = Format(FSDT, "DD/MM/YYYY")
  FrmGRNEntryList1.Show 1
  FRMLRDTL.Tab = 0
  
    If ALLOWEDITDEL = False Then
       MsgBox "Purchase of this GRN have been made can not edit/delete ", vbInformation
       Call ClsData(Me)
       btn_sts (True)
       Call SetLastDate(Me, "TXTVBDT", "IVR", M_DBCD_DIRIVR)
       cmdAdd.SetFocus
   Else
    'Check for Receipt and Payment Entires
   If TXTVBNO = Empty Then
      Call SetLastDate(Me, "TXTVBDT", "IVR", M_DBCD_DIRIVR)
      Exit Sub
   Else
      TXTBNET.Text = Format(FormatNumber(Val(TXTBNET.Text), 0), "##########.00")
   End If
     
   'FIFO
    If Trim(M_ISSUE) = "Y" Then
        Call SetLastDate(Me, "TXTVBDT", "IVR", M_DBCD_DIRIVR)
        MsgBox "You Can Not Edit GRN Detail!! Issue Entry Exist!!", vbExclamation, "Access Denied"
        Call cmdCancel_Click
        'cmdDelete.Enabled = False
        'cmdSave.Enabled = False
        Exit Sub
    End If
    '------------
    TXTVBDT.Enabled = False
    btn_sts (False)
    TXTDBAC.SetFocus
  End If
End Sub

Private Sub CMDEXIT_Click()
  Unload Me
End Sub

Private Sub Flex_Click()
  cmddelitm.Enabled = True
End Sub

Private Sub FLEX_EnterCell()
  FLEX.CellBackColor = RGB(BRED, BGREEN, BBLUE)
  Emptycell = True
  
  If FLEX.COL = 9 Then
        If UCase(Trim(M_Currency)) <> UCase(Trim(txtCURNCY)) Then
            If Val(FLEX.TextMatrix(FLEX.ROW, 9)) > 0 Then
            '    Flex.TextMatrix(Flex.ROW, 9) = Format(Val(Flex.TextMatrix(Flex.ROW, 9)) / Val(TXTEXRATE), "####.000")
            Else
            '   Flex.TextMatrix(Flex.ROW, 9) = 0
            End If
        End If
  End If
End Sub

Private Sub FLEX_GotFocus()
  Me.KeyPreview = False
  FLEX.TextMatrix(FLEX.ROW, 0) = FLEX.ROW
  FLEX.TextMatrix(1, 1) = TXTSCHLN
  FLEX.TextMatrix(1, 2) = Format(TXTSCHLNDT.Value, "DD/MM/YYYY")
  
End Sub

Private Sub Flex_LeaveCell()
  Dim FLEXROW As Double
  Dim FLEXCOL As Double
  Dim i As Double
  If M_COMPBILL = "VFL" Then
    FLEX.TextMatrix(FLEX.ROW, 4) = "0000"
    FLEX.TextMatrix(FLEX.ROW, 5) = "IST"
  End If
  If M_COMPBILL = "VF1" Then
    FLEX.TextMatrix(FLEX.ROW, 5) = "IST"
  End If
  FLEX.CellBackColor = vbWhite
  FLEX.TextMatrix(FLEX.ROW, 10) = Format(Val(FLEX.TextMatrix(FLEX.ROW, 9)) * Val(FLEX.TextMatrix(FLEX.ROW, 8)), "#########.00")
    
  If UCase(Trim(M_Currency)) <> UCase(Trim(txtCURNCY)) Then
     If Val(txtEXRate) > 0 Then
        FLEX.TextMatrix(FLEX.ROW, 10) = Format(Val(FLEX.TextMatrix(FLEX.ROW, 9)) * Val(txtEXRate) * Val(FLEX.TextMatrix(FLEX.ROW, 8)), "####.00")
     End If
  End If
  
  FLEXROW = FLEX.ROW
  FLEXCOL = FLEX.COL
  TXTTPCS.Text = 0
  TXTTQTY.Text = 0
  TXTITOT.Text = 0
  For i = 1 To FLEX.Rows - 1
    TXTTPCS.Text = Val(Format(Val(TXTTPCS.Text) + Val(FLEX.TextMatrix(i, 7)), "######"))
    TXTTQTY.Text = Format(Val(TXTTQTY.Text) + Val(FLEX.TextMatrix(i, 8)), "########.000")
    TXTITOT.Text = Format(Val(TXTITOT.Text) + Val(FLEX.TextMatrix(i, 10)), "########.00")
  Next
  
  FLEX.ROW = FLEXROW
  FLEX.COL = FLEXCOL
  
  FLEX.SetFocus
End Sub

Private Sub FLEX_LostFocus()
  Dim FLEXROW As Double
  Dim FLEXCOL As Double
  Dim i As Double
  FLEX.CellBackColor = vbWhite
  FLEX.TextMatrix(FLEX.ROW, 10) = Format(Val(FLEX.TextMatrix(FLEX.ROW, 9)) * Val(FLEX.TextMatrix(FLEX.ROW, 8)), "#########.00")
  
  If UCase(Trim(M_Currency)) <> UCase(Trim(txtCURNCY)) Then
     If Val(txtEXRate) > 0 Then
        FLEX.TextMatrix(FLEX.ROW, 10) = Format(Val(FLEX.TextMatrix(FLEX.ROW, 9)) * Val(txtEXRate) * Val(FLEX.TextMatrix(FLEX.ROW, 8)), "####.00")
     End If
  End If
  
  
  FLEXROW = FLEX.ROW
  FLEXCOL = FLEX.COL
  TXTTPCS.Text = 0
  TXTTQTY.Text = 0
  TXTITOT.Text = 0
  For i = 1 To FLEX.Rows - 1
    TXTTPCS.Text = Format(Val(TXTTPCS.Text) + Val(FLEX.TextMatrix(i, 7)), "######")
    TXTTQTY.Text = Format(Val(TXTTQTY.Text) + Val(FLEX.TextMatrix(i, 8)), "########.000")
    TXTITOT.Text = Format(Val(TXTITOT.Text) + Val(FLEX.TextMatrix(i, 10)), "########.00")
  Next
  FLEX.ROW = FLEXROW
  FLEX.COL = FLEXCOL
  
  'FLEX.SetFocus
End Sub

Private Sub flexBTRM_EnterCell()
If flexBTRM.COL <> 0 Then flexBTRM.CellBackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub setflexhead()
    FLEX.TextMatrix(0, 0) = "Sr."
    FLEX.TextMatrix(0, 1) = "ChallanNo."
    FLEX.TextMatrix(0, 2) = "ChallanDt."
    FLEX.TextMatrix(0, 3) = "Item Name"
    FLEX.TextMatrix(0, 4) = "Merge No."
    FLEX.TextMatrix(0, 5) = "Grade"
    FLEX.TextMatrix(0, 6) = "Cops"
    FLEX.TextMatrix(0, 7) = "Pieces"
    FLEX.TextMatrix(0, 8) = "Quantity"
    FLEX.TextMatrix(0, 9) = "Rate"
    FLEX.TextMatrix(0, 10) = "Amount"
    FLEX.TextMatrix(0, 11) = "ICOD"
    FLEX.TextMatrix(0, 12) = "RTYP"
    FLEX.TextMatrix(0, 13) = "RSRN"
    FLEX.TextMatrix(0, 14) = "ORDN"
    FLEX.TextMatrix(0, 15) = "ORDRATE"
    FLEX.TextMatrix(0, 16) = "TWST"
    FLEX.TextMatrix(0, 17) = ""
    FLEX.TextMatrix(0, 18) = ""
    FLEX.TextMatrix(0, 19) = ""
    FLEX.TextMatrix(0, 20) = "Disc %"
    FLEX.TextMatrix(0, 21) = "Disc Amnt"
    FLEX.TextMatrix(0, 22) = "Vat %"
    FLEX.TextMatrix(0, 23) = "VAT Amnt"
    FLEX.ColWidth(0) = 350
    FLEX.ColWidth(1) = 0
    FLEX.ColWidth(2) = 0
    FLEX.ColWidth(3) = 2750
    FLEX.ColWidth(4) = 1300
    FLEX.ColWidth(5) = 900
    FLEX.ColWidth(6) = 1000
    FLEX.ColWidth(7) = 1000
    FLEX.ColWidth(8) = 1100
    FLEX.ColWidth(9) = 1000
    FLEX.ColWidth(10) = 1400
    FLEX.ColWidth(11) = 0
    FLEX.ColWidth(12) = 0
    FLEX.ColWidth(13) = 0
    FLEX.ColWidth(14) = 0
    FLEX.ColWidth(15) = 0
    FLEX.ColWidth(16) = 0
    FLEX.ColWidth(17) = 0
    FLEX.ColWidth(18) = 0
    FLEX.ColWidth(19) = 0
    FLEX.ColWidth(20) = 0
    FLEX.ColWidth(21) = 0
    FLEX.ColWidth(22) = 0
    FLEX.ColWidth(23) = 0
    FLEX.ColAlignment(3) = 0
    FLEX.ColAlignment(2) = 0
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
    FRMBTRM.Enabled = Not Yes
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
  RS.Open "select * from config where comp='" & compPth & "' and vtyp='" & FRMPARA & "' AND DBCD='" & M_DBCD_DIRIVR & "'  AND UNIT='" & UNCD & "' order by srch", CN, adOpenKeyset, adLockPessimistic
  CNTR = 0
  Do While Not RS.EOF
   flexBTRM.Rows = flexBTRM.Rows + 1
   flexBTRM.TextMatrix(CNTR, 0) = RS!NICK & ""
   flexBTRM.TextMatrix(CNTR, 1) = Format(RS!PERC, "#######.00")
   flexBTRM.TextMatrix(CNTR, 3) = "N"
   
   Dim EXPRS As ADODB.Recordset: Set EXPRS = New ADODB.Recordset
   If EXPRS.State = 1 Then EXPRS.Close
   EXPRS.Open "SELECT DIRECTEXP FROM CHRGMST WHERE NAME='" & flexBTRM.TextMatrix(CNTR, 0) & "'", CN, adOpenDynamic, adLockOptimistic
   If Not EXPRS.EOF Then
      flexBTRM.TextMatrix(CNTR, 4) = Mid(Trim(EXPRS!DIRECTEXP & "") + "N", 1, 1)
   Else
      flexBTRM.TextMatrix(CNTR, 4) = "N"
   End If
      
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

Private Sub FRMLRDTL_Click(PreviousTab As Integer)
   If FRMLRDTL.Tab = 0 Then
     If Val(TxtPallet) = 0 And Val(txtCops) = 0 And Val(txtPly) = 0 And chkReturnable = 1 Then
        chkReturnable.Value = 0
     End If
   ElseIf FRMLRDTL.Tab = 1 Then
     chkReturnable.Value = 1
   End If
End Sub

Private Sub LBLGRS_Change()
  TXTITOT = Format(LBLGRS, "#########.00")
End Sub



Private Sub txtBEdit_GotFocus()
 txtBEdit.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtBEdit_LostFocus()
 txtBEdit.BackColor = vbWhite
End Sub

Private Sub txtCURNCY_Change()
    'If Not SAVEFLAG And Flex.Rows <= 1 Then Exit Sub
    
    If SAVEFLAG Then
        If UCase(Trim(txtCURNCY)) <> UCase(Trim(M_Currency)) Then
            LBLEXRATE.Enabled = True
            txtEXRate.Enabled = True
        Else
            LBLEXRATE.Enabled = False
            txtEXRate.Enabled = False
            txtEXRate = Empty
        End If
    ElseIf Not SAVEFLAG Then
        If UCase(Trim(txtCURNCY)) = UCase(Trim(M_Currency)) And Val(txtEXRate) <> 0 Then
            txtEXRate.Enabled = False
            LBLEXRATE.Enabled = False
            txtEXRate = Empty
        ElseIf UCase(Trim(txtCURNCY)) <> UCase(Trim(M_Currency)) Then
            LBLEXRATE.Enabled = True
            txtEXRate.Enabled = True
            txtEXRate = txtEXRate.Tag
        End If
    End If
    
    
    If UCase(Trim(txtCURNCY)) <> UCase(Trim(M_Currency)) Then
       ShapeCHA.Visible = True
       lblcha.Visible = True
       lblfrt.Visible = True
       lbldty.Visible = True
       txtcha.Visible = True
       txtfrt.Visible = True
       txtdty.Visible = True
       lblnote.Visible = True
   Else
       ShapeCHA.Visible = False
       lblcha.Visible = False
       lblfrt.Visible = False
       lbldty.Visible = False
       txtcha.Visible = False
       txtfrt.Visible = False
       txtdty.Visible = False
       lblnote.Visible = False
   End If
    
    
    If txtCURNCY = Empty Then txtCURNCY = Trim(M_Currency)
End Sub

Private Sub txtCURNCY_GotFocus()
    txtCURNCY.BackColor = RGB(BRED, BGREEN, BBLUE)
    SendKeys "{HOME}+{END}"
End Sub

Private Sub txtCURNCY_LostFocus()
    txtCURNCY.BackColor = vbWhite
End Sub

Private Sub TXTDBAC_Change()
  If RS.State = 1 Then RS.Close
  RS.Open "SELECT * FROM ACCMST WHERE NAME='" & TXTDBAC.Text & "'", CN, adOpenDynamic, adLockPessimistic
  If Not RS.EOF Then
    TXTRTORTAX.Text = RS!TTYP & ""
  End If
End Sub

Private Sub TXTDBAC_GotFocus()
TXTDBAC.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTDBAC_KeyDown(KeyCode As Integer, Shift As Integer)
   Me.KeyPreview = False
   If KeyCode = vbKeyF2 Or Trim(TXTDBAC.Text) = Empty Then
        
        NEW_VISIBLE = True
        M_DESC = Empty
        Key = Empty
        TXTDBAC.Text = SearchList1("SELECT  TOP 20 CODE,NAME FROM ACCMST", 0, TXTDBAC, "List of Debit A/c")
        If key_PressNew = True Then
            M_DESC = ""
            Key = ""
            TXTDBAC.Text = ""
            frm_Acc.Show
            
        End If
    ElseIf KeyCode = vbKeyDelete Then
        TXTDBAC = Empty
    End If
    Me.KeyPreview = True
    'If KeyCode = vbKeyReturn Then
    '  SendKeys "{TAB}"
    'End If
    Dim M_BRCD
    Dim MSTDAT As New ADODB.Recordset
    Set MSTDAT = New ADODB.Recordset
    If MSTDAT.State = 1 Then MSTDAT.Close
    MSTDAT.Open "SELECT * FROM ACCMST WHERE NAME='" & TXTDBAC.Text & "'", CN, adOpenDynamic, adLockOptimistic
    If Not MSTDAT.EOF Then
      M_BRCD = MSTDAT!BRCD & ""
    End If
End Sub

Private Sub TXTDBAC_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn And (Not Trim(TXTDBAC.Text) = Empty) Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub TXTDBAC_LostFocus()
  TXTDBAC.BackColor = vbWhite
  
  If SAVEFLAG Then
     Dim GETRS As ADODB.Recordset
     Set GETRS = New ADODB.Recordset
  
     If GETRS.State = 1 Then GETRS.Close
     GETRS.Open "SELECT TXCD,TTYP FROM ACCMST WHERE NAME='" & TXTDBAC & "' ", CN, adOpenDynamic, adLockOptimistic
     If Not GETRS.EOF Then
        TXTTAXNAM = GetCode("TAXMST", GETRS!TXCD & "", "CODE", "NAME")
        TXTRTORTAX = Trim(GETRS!TTYP & "")
     End If
 End If
 
End Sub

Private Sub txtEXRate_Change()
 If UCase(Trim(M_Currency)) = UCase(Trim(txtCURNCY)) Then Exit Sub
 If Val(txtEXRate) <= 0 Then Exit Sub
 
 Dim i As Long
 For i = 1 To FLEX.Rows - 1
     FLEX.TextMatrix(i, 10) = Format(Val(FLEX.TextMatrix(i, 9)) * Val(FLEX.TextMatrix(i, 8)) * Val(txtEXRate), "#######.00")
 Next i
End Sub

Private Sub TXTEXRATE_GotFocus()
   txtEXRate.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTEXRATE_KeyPress(KeyAscii As Integer)
   If CheckNumericKey(KeyAscii, txtEXRate, Me) = 0 Then KeyAscii = 0
End Sub

Private Sub TXTEXRATE_LostFocus()
  txtEXRate.BackColor = vbWhite
End Sub

Private Sub txtEXRate_Validate(Cancel As Boolean)
    If Val(txtEXRate) = 0 And UCase(Trim(txtCURNCY)) <> UCase(Trim(M_Currency)) Then
        MsgBox "Please Enter Exchange Rate !!", vbInformation, "Exchange Rate Required !!"
        Cancel = True
    End If
End Sub

Private Sub TXTGDN_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF2 Or Trim(TXTGDN.Text) = Empty Then
        NEW_VISIBLE = True
        M_DESC = Empty
        Key = Empty
        TXTGDN.Text = SearchList1("SELECT  TOP 20 CODE,NAME FROM LOCMST", 0, TXTGDN, "List of GODOWN")
        TXTGDN.Tag = Key
        If key_PressNew = True Then
           M_DESC = "": Key = "": TXTGDN.Text = ""
           frm_mstlocation.Show
        End If
    ElseIf KeyCode = vbKeyDelete Then
        TXTGDN = Empty
    End If
    
    If KeyCode = vbKeyReturn Then
       SendKeys "{TAB}"
    End If
End Sub

Private Sub TxtPallet_KeyPress(KeyAscii As Integer)
  If CheckNumericKey(KeyAscii, TxtPallet, Me) = 0 Then KeyAscii = 0
End Sub

Private Sub TXTRTORTAX_GotFocus()
TXTRTORTAX.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTRTORTAX_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Or Trim(TXTRTORTAX) = Empty Then
    TXTRTORTAX = SearchList1("SELECT DISTINCT TTYP AS CODE,TTYP AS NAME FROM accmst where NOT (TTYP='' OR TTYP IS NULL)", 0, TXTRTORTAX, "Select Sale Tax Type")
  End If
End Sub

Private Sub TXTRTORTAX_LostFocus()
TXTRTORTAX.BackColor = vbWhite
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
    If TXTVBDT.Enabled Then
       TXTVBDT.SetFocus
    Else
       SendKeys "{TAB}"
    End If
  End If
End Sub

Private Sub TXTSCHLN_GotFocus()
TXTSCHLN.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTSCHLN_LostFocus()
TXTSCHLN.BackColor = vbWhite
End Sub

Private Sub TXTITOT_Change()
  If flexBTRM.Rows > 0 Then
    flexBTRM.COL = 0
    flexBTRM.ROW = 0
  End If
  calBTRM 0
  Call calADLS
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
  TXTLRNO.SelLength = Len(TXTLRNO)
End Sub


Private Sub TXTLRNO_LostFocus()
 TXTLRNO.BackColor = vbWhite
End Sub

Private Sub TXTRMRK_GotFocus()
  TXTRMRK.BackColor = RGB(BRED, BGREEN, BBLUE)
  TXTRMRK.SelStart = 0
  TXTRMRK.SelLength = Len(TXTRMRK)
End Sub

Private Sub TXTRMRK_LostFocus()
  TXTRMRK.BackColor = vbWhite
End Sub

Private Sub TXTTAXNAM_GotFocus()
TXTTAXNAM.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTTAXNAM_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Or Trim(TXTTAXNAM.Text) = Empty Then
        NEW_VISIBLE = True
        M_DESC = Empty
        Key = Empty
      
        TXTTAXNAM.Text = SearchList1("SELECT  TOP 20 CODE,NAME FROM TAXMST WHERE RECSTAT='A'", 0, TXTTAXNAM, "List of Tax Catagoery")
        If key_PressNew = True Then
            M_DESC = ""
            Key = ""
            Ref_Cat = "T"
            TXTTAXNAM.Text = ""
            FrmSaleTaxMaster.Show
        End If
    ElseIf KeyCode = vbKeyDelete Then
        TXTTAXNAM = Empty
    End If
    If KeyCode = vbKeyReturn Then
      SendKeys "{TAB}"
    End If
End Sub

Private Sub TXTTAXNAM_LostFocus()
TXTTAXNAM.BackColor = vbWhite
End Sub

Private Sub TXTTPCS_Change()
  calBTRM 0
End Sub

Private Sub TXTTQTY_Change()
  calBTRM 0
End Sub

Private Sub TXTTRNM_GotFocus()
 TXTTRNM.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTTRNM_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Or Trim(TXTTRNM.Text) = Empty Then
        NEW_VISIBLE = True
        M_DESC = Empty
        Key = Empty
        TXTTRNM.Text = SearchList1("SELECT  TOP 20 CODE,NAME FROM TRANSPORTMST", 0, TXTTRNM, "List of Transporter")
        If key_PressNew = True Then
           M_DESC = "": Key = "": TXTTRNM.Text = ""
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
        If Val(flexBTRM.TextMatrix(J, 1)) <> 0 Then
              flexBTRM.TextMatrix(J, 2) = 0
        End If
    Next J

    For J = flexBTRM.ROW To flexBTRM.Rows - 1
        c_FMLA(J) = Trim(M_FMLA(J))
        If Len(c_FMLA(J)) <= 6 Then
            Select Case c_FMLA(J)
                Case "M_STOT"
                    If Val(flexBTRM.TextMatrix(J, 1)) <> 0 Then
                        flexBTRM.TextMatrix(J, 2) = Format((Val(flexBTRM.TextMatrix(J, 1)) * Val(TXTITOT.Text)) / 100, "##########.00")
                    Else
                        flexBTRM.TextMatrix(J, 2) = Format(flexBTRM.TextMatrix(J, 2), "#.00")
                    End If
                Case "M_TQTY"
                    If Val(flexBTRM.TextMatrix(J, 1)) <> 0 Then
                        flexBTRM.TextMatrix(J, 2) = Format((Val(flexBTRM.TextMatrix(J, 1)) * Val(TXTTQTY.Text)), "##########.00")
                    Else
                        flexBTRM.TextMatrix(J, 2) = Format(flexBTRM.TextMatrix(J, 2), "#.00")
                    End If
                Case "M_TPCS"
                    If Val(flexBTRM.TextMatrix(J, 1)) <> 0 Then
                        flexBTRM.TextMatrix(J, 2) = Format((Val(flexBTRM.TextMatrix(J, 1)) * Val(TXTTPCS.Text)), "##########.00")
                    Else
                        flexBTRM.TextMatrix(J, 2) = Format(flexBTRM.TextMatrix(J, 2), "#.00")
                    End If
                Case "AMT_01"
                    If Val(flexBTRM.TextMatrix(J, 1)) <> 0 Then
                        flexBTRM.TextMatrix(J, 2) = Format((Val(flexBTRM.TextMatrix(J, 1)) * Val(flexBTRM.TextMatrix(0, 2))) / 100, "##########.00")
                    Else
                        flexBTRM.TextMatrix(J, 2) = Format(flexBTRM.TextMatrix(J, 2), "#.00")
                    End If
                Case "AMT_02"
                    If Val(flexBTRM.TextMatrix(J, 1)) <> 0 Then
                        flexBTRM.TextMatrix(J, 2) = Format((Val(flexBTRM.TextMatrix(J, 1)) * Val(flexBTRM.TextMatrix(1, 2))) / 100, "##########.00")
                    Else
                        flexBTRM.TextMatrix(J, 2) = Format(flexBTRM.TextMatrix(J, 2), "#.00")
                    End If
                Case "AMT_03"
                    If Val(flexBTRM.TextMatrix(J, 1)) <> 0 Then
                        flexBTRM.TextMatrix(J, 2) = Format((Val(flexBTRM.TextMatrix(J, 1)) * Val(flexBTRM.TextMatrix(2, 2))) / 100, "##########.00")
                    Else
                        flexBTRM.TextMatrix(J, 2) = Format(flexBTRM.TextMatrix(J, 2), "#.00")
                    End If
                Case "AMT_04"
                    If Val(flexBTRM.TextMatrix(J, 1)) <> 0 Then
                        flexBTRM.TextMatrix(J, 2) = Format((Val(flexBTRM.TextMatrix(J, 1)) * Val(flexBTRM.TextMatrix(3, 2))) / 100, "##########.00")
                    Else
                        flexBTRM.TextMatrix(J, 2) = Format(flexBTRM.TextMatrix(J, 2), "#.00")
                    End If
                Case "AMT_05"
                    If Val(flexBTRM.TextMatrix(J, 1)) <> 0 Then
                        flexBTRM.TextMatrix(J, 2) = Format((Val(flexBTRM.TextMatrix(J, 1)) * Val(flexBTRM.TextMatrix(4, 2))) / 100, "##########.00")
                    Else
                        flexBTRM.TextMatrix(J, 2) = Format(flexBTRM.TextMatrix(J, 2), "#.00")
                    End If
                Case "AMT_06"
                    If Val(flexBTRM.TextMatrix(J, 1)) <> 0 Then
                        flexBTRM.TextMatrix(J, 2) = Format((Val(flexBTRM.TextMatrix(J, 1)) * Val(flexBTRM.TextMatrix(5, 2))) / 100, "##########.00")
                    Else
                        flexBTRM.TextMatrix(J, 2) = Format(flexBTRM.TextMatrix(J, 2), "#.00")
                    End If
                Case "AMT_07"
                    If Val(flexBTRM.TextMatrix(J, 1)) <> 0 Then
                        flexBTRM.TextMatrix(J, 2) = Format((Val(flexBTRM.TextMatrix(J, 1)) * Val(flexBTRM.TextMatrix(6, 2))) / 100, "##########.00")
                    Else
                        flexBTRM.TextMatrix(J, 2) = Format(flexBTRM.TextMatrix(J, 2), "#.00")
                    End If
                Case "AMT_08"
                    If Val(flexBTRM.TextMatrix(J, 1)) <> 0 Then
                        flexBTRM.TextMatrix(J, 2) = Format((Val(flexBTRM.TextMatrix(J, 1)) * Val(flexBTRM.TextMatrix(7, 2))) / 100, "##########.00")
                    Else
                        flexBTRM.TextMatrix(J, 2) = Format(flexBTRM.TextMatrix(J, 2), "#.00")
                    End If
                Case "AMT_09"
                    If Val(flexBTRM.TextMatrix(J, 1)) <> 0 Then
                        flexBTRM.TextMatrix(J, 2) = Format((Val(flexBTRM.TextMatrix(J, 1)) * Val(flexBTRM.TextMatrix(8, 2))) / 100, "##########.00")
                    Else
                        flexBTRM.TextMatrix(J, 2) = Format(flexBTRM.TextMatrix(J, 2), "#.00")
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

TXTBNET.Text = Format(FormatNumber(Val(TXTBNET.Text), 0), "##########.00")
    
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
        If FLEX.COL - 1 <> 7 And FLEX.COL - 1 <> 0 Then FLEX.TextMatrix(FLEX.ROW, FLEX.COL - 1) = 0
   Case 13    ' ENTER return focus to MSHFlexGrid.
         MSHFlexGrid.SetFocus
         If MSHFlexGrid.COL = 3 Then
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
      TXTBNET = TXTITOT
    End If
    
    TXTBNET.Text = Format(FormatNumber(Val(TXTBNET.Text), 0), "##########.00")
    
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
      If MSHFlexGrid.COL = 3 Then
            If MSHFlexGrid.Rows <> MSHFlexGrid.ROW + 1 Then
                MSHFlexGrid.ROW = MSHFlexGrid.ROW + 1
            Else
              If TXTLRNO.Enabled And TXTLRNO.Visible Then
                 TXTLRNO.SetFocus
              Else
                 cmdSave.SetFocus
              End If
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
      If KeyAscii <> 45 Then
      Edt = Chr(KeyAscii)
      End If
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
    
    TXTBNET.Text = Format(FormatNumber(Val(TXTBNET.Text), 0), "##########.00")
    
End Sub

Private Sub txtBEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    'If flexBTRM.COL = 2 Then
    
    '    If flexBTRM.TextMatrix(flexBTRM.ROW, 4) = "Y" And KeyCode = vbKeyReturn Then
    '        Dim PREV_VALUE As Double
    '        PREV_VALUE = Val(flexBTRM.TextMatrix(flexBTRM.ROW, 2))
    '
    '         If (Val(txtBEdit) > PREV_VALUE + 1) Or (Val(txtBEdit) < PREV_VALUE - 1) Then
    '           txtBEdit = PREV_VALUE
    '        End If
    '
    '    End If
         
    ' End If
    
     EditKeyCode flexBTRM, txtBEdit, KeyCode, Shift
    
End Sub

Private Sub txtBEdit_KeyPress(KeyAscii As Integer)
   If KeyAscii = Asc(vbCr) Then KeyAscii = 0
   'If flexBTRM.TextMatrix(flexBTRM.ROW, 4) = "Y" And flexBTRM.COL <> 2 Then KeyAscii = 0
   
   If CheckNumericKey(KeyAscii, txtBEdit, Me) = 0 Then KeyAscii = 0
    
   If flexBTRM.COL = 3 Then
      
      If KeyAscii = 8 And Len(Trim(txtBEdit)) > 0 Then
         txtBEdit = IIf(Trim(txtBEdit) = "Y", "N", "Y")
      ElseIf KeyAscii = 8 Then
         KeyAscii = 0
      End If
      
      If Chr(KeyAscii) <> "Y" And Chr(KeyAscii) <> "N" Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub flexBTRM_LeaveCell()
If flexBTRM.COL <> 0 Then flexBTRM.CellBackColor = vbWhite
    If txtBEdit.Visible = False Then Exit Sub
    flexBTRM = txtBEdit
    txtBEdit.Visible = False
End Sub

Private Sub flexBTRM_RowColChange()
    If flexBTRM.COL = 1 Then
        'If flexBTRM.ROW = 0 Then txtDUTY.Text = Format(Val(txtDUTY.Text), "#########.00")
        If calbtm = True Then
            calBTRM FLEX.ROW
        End If
    End If
    If flexBTRM.Rows > 7 Then
        If flexBTRM.ROW Mod 5 = 0 And flexBTRM.ROW <> 0 Then
            flexBTRM.TopRow = 5
        End If
    End If
    calADLS
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
  
  'Godown Code
  If CHKRS.State = 1 Then CHKRS.Close
  CHKRS.Open "SELECT * from LOCMST WHERE NAME='" & TXTGDN.Text & "'", CN, adOpenKeyset, adLockPessimistic
  If CHKRS.EOF Then
    ' MsgBox "Godown Name Not Define ", vbCritical
    ' TXTGDN.SetFocus
    ' CHKSAVEDATA = False
    ' Exit Function
  End If
    
  Dim i As Double
  Dim IVR_icod As String
  For i = 1 To FLEX.Rows - 1
     IVR_icod = FLEX.TextMatrix(i, 11)
     If CHKRS.State = 1 Then CHKRS.Close
     CHKRS.Open "select * from itmmst where code='" & IVR_icod & "'", CN, adOpenKeyset, adLockPessimistic
     If CHKRS.EOF Then
        MsgBox "Item Missing From Master !!! ", vbCritical
        CHKSAVEDATA = False
        Exit Function
     End If
     'If SAVEFLAG = True Then
     If CHKRS.State = 1 Then CHKRS.Close
     CHKRS.Open "SELECT * FROM MRGMST WHERE COMP = '" & compPth & "' AND UNIT = '" & UNCD & "' AND  ICOD = '" & FLEX.TextMatrix(i, 11) & "' AND MRGN = '" & FLEX.TextMatrix(i, 4) & "'", CN, adOpenDynamic, adLockOptimistic
     If Not CHKRS.EOF Then
        If GetCode("ACCMST", TXTDBAC, "NAME", "CODE") <> CHKRS!PCOD & "" Then
        MsgBox "This Merge No. Already Exist !!! ", vbCritical
        CHKSAVEDATA = False
        Exit Function
        End If
     End If
     'End If
  Next
  If CHKRS.State = 1 Then CHKRS.Close
  CHKRS.Open "SELECT * FROM " & SUMMARYTABLE & " WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
             "' AND VTYP = 'IVR' AND DBCD='" & M_DBCD_DIRIVR & "' AND VBNO='" & TXTVBNO & "'", CN, adOpenDynamic, adLockOptimistic
  If Not CHKRS.EOF Then
    If CHKRS!VBNO = TXTVBNO Then
      'O.k
     Else
      MsgBox "Duplicate GRN No. !!!! ", vbCritical
      CHKSAVEDATA = False
      Exit Function
    End If
  End If
  
  'CONSIDER DBCD BECAUSE SAME CHALLAN MAY ENTER FOR JOB ALSO
  If CHKRS.State = 1 Then CHKRS.Close
  CHKRS.Open "SELECT * FROM " & SUMMARYTABLE & " WHERE COMP ='" & compPth & "' AND UNIT ='" & UNCD & _
            "' AND VTYP = 'IVR' AND DBCD='" & M_DBCD_DIRIVR & "' AND PCOD ='" & GetCode("ACCMST", TXTDBAC, "NAME", "CODE") & _
            "' AND CHLN ='" & Trim(TXTSCHLN.Text) & _
            "' AND RECSTAT <> 'D' AND DATE>='" & Format(FSDT, "MM/dd/yyyy") & _
            "' AND DATE<='" & Format(FEDT, "MM/dd/yyyy") & "'", CN, adOpenDynamic, adLockOptimistic
    
  If CHKRS.EOF = False And SAVEFLAG Then
     MsgBox "Duplicate Party Challan No. ", vbCritical
     If TXTSCHLN.Enabled Then TXTSCHLN.SetFocus
     CHKSAVEDATA = False
     Exit Function
  End If
  CHKRS.Close
  
  CHKSAVEDATA = True
End Function

Private Sub cmdSave_Click()
  On Error GoTo LAST
  If CHKSAVEDATA = False Then
    Exit Sub
  End If
  
  If txtCURNCY = Empty Then txtCURNCY = Trim(M_Currency)
  
  Call CHKFLEX
  If Not CHK_FLX Then
    MsgBox "Invalid Item Detail"
    FLEX.ROW = FLXROW
    FLEX.COL = FLXCOL
    Exit Sub
  End If
  
  If SAVEFLAG = True Then
    TXTVBNO = GenVNO("IVR", M_DBCD_DIRIVR)
  End If
  
  Dim SAVDAT As ADODB.Recordset
  Set SAVDAT = New ADODB.Recordset
  If SAVDAT.State = 1 Then SAVDAT.Close
  SAVDAT.Open "SELECT * FROM " & SUMMARYTABLE & " WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
              "' AND VTYP='IVR' AND VBNO='" & TXTVBNO & "' AND DBCD='" & M_DBCD_DIRIVR & "'", CN, adOpenDynamic, adLockOptimistic
  If Not SAVDAT.EOF Then
    If SAVDAT!VBNO = TXTVBNO Then
     Else
      MsgBox "Duplicate GRN No.Make Change in Configuration for Invoice No."
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
  Resume
  MsgBox ERR.Description
End Sub

Private Sub SAVERECIVR()
  
  On Error GoTo LAST
  Dim SQL As String
  Dim GDNCODE As String
  Dim SAVDAT As New ADODB.Recordset
  Dim MSTDAT As New ADODB.Recordset
  Dim M_CRAC As String
  Dim M_DRAC As String
  Dim M_PCOD As String
  Dim M_DCOD As String
  Dim M_CPCD As String
  Dim M_ARCD As String
  Dim M_TRCD As String
  Dim M_TXCD As String
  Dim M_BRCD As String
  Dim i As Double
  Dim J As Double
  Set SAVDAT = New ADODB.Recordset
  Set MSTDAT = New ADODB.Recordset
  'Party A/c
  If SAVDAT.State = 1 Then SAVDAT.Close
  SAVDAT.Open "SELECT * FROM ACCMST WHERE NAME='" & TXTDBAC.Text & "'", CN, adOpenDynamic, adLockOptimistic
  M_DRAC = SAVDAT!CODE & ""
  M_PCOD = SAVDAT!CODE & ""
  M_CPCD = SAVDAT!CPCD & ""
  M_ARCD = SAVDAT!ARCD & ""
  M_BRCD = SAVDAT!BRCD & ""
       
  'Retail / Tax Invoice
  If SAVDAT.State = 1 Then SAVDAT.Close
  SAVDAT.Open "SELECT * FROM TAXMST WHERE NAME='" & TXTTAXNAM.Text & "'", CN, adOpenDynamic, adLockOptimistic
  If Not SAVDAT.EOF Then
     M_TXCD = SAVDAT!CODE & ""
  End If
  SAVDAT.Close
  
  'TRANSPORT
  If SAVDAT.State = 1 Then SAVDAT.Close
  SAVDAT.Open "SELECT * FROM TRANSPORTMST WHERE NAME='" & TXTTRNM.Text & "'", CN, adOpenDynamic, adLockOptimistic
  If Not SAVDAT.EOF Then
     M_TRCD = SAVDAT!CODE & ""
  End If
  SAVDAT.Close
  
 
  If SAVDAT.State = 1 Then SAVDAT.Close
  SAVDAT.Open "SELECT * FROM LOCMST WHERE NAME='" & TXTGDN.Text & "'", CN, adOpenDynamic, adLockOptimistic
  If Not SAVDAT.EOF Then
     GDNCODE = SAVDAT!CODE & ""
  End If
  SAVDAT.Close

  Dim excperc As Double
  excperc = 100
  If cmbSelection = "RG23-C" Then
    If RS.State = 1 Then RS.Close
    RS.Open "select exccperc from untcfg where comp='" & compPth & "' and unit='" & UNCD & "'", CN, adOpenDynamic, adLockOptimistic
    If Not RS.EOF Then
      excperc = RS!exccperc
    End If
  End If
  
  CN.BeginTrans
  
  'If SAVEFLAG = False And M_ISSUE = "Y" Then
  '   CN.Execute "UPDATE " & SUMMARYTABLE & " SET "
  '   Exit Sub
  'End If
  
  Call DELETEIVR
  SQL = Empty
  If SAVDAT.State = 1 Then SAVDAT.Close
  SAVDAT.Open "SELECT * FROM " & SUMMARYTABLE & " WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
  "' AND VTYP='IVR' AND DBCD='" & M_DBCD_DIRIVR & "' AND VBNO='" & TXTVBNO & "' ", CN, adOpenDynamic, adLockOptimistic
  If SAVDAT.EOF Then
      SAVDAT.AddNew
  End If
  
  SAVDAT!COMP = compPth
  SAVDAT!VTYP = "IVR"
  SAVDAT!SRNO = TXTVBNO
  SAVDAT!SRCH = 1
  SAVDAT!dbcd = M_DBCD_DIRIVR
  SAVDAT!Date = Format(TXTVBDT.Value, "YYYY/MM/DD")
  SAVDAT!VBNO = Trim(TXTVBNO.Text)
  SAVDAT!chln = Trim(TXTSCHLN.Text)
  SAVDAT!CHDT = Format(TXTSCHLNDT.Value, "YYYY/MM/DD")
  SAVDAT!CRAC = M_CRAC
  
  If chkAcEffect.Value = 1 Then
     SAVDAT!ACEFFECT = "Y"
  Else
     SAVDAT!ACEFFECT = "N"
  End If
  
  SAVDAT!DRAC = M_DRAC
  SAVDAT!PCOD = M_PCOD
  SAVDAT!DCOD = M_DCOD
  SAVDAT!BRCD = M_BRCD
  SAVDAT!CPCD = M_CPCD
  SAVDAT!ARCD = M_ARCD
  SAVDAT!TXCD = M_TXCD
  SAVDAT!TPCS = 0
  SAVDAT!TQTY = 0
  SAVDAT!ITOT = Val(TXTITOT.Text)
  SAVDAT!BADJ = Val(TXTBNET.Text) - Val(TXTITOT.Text)
  SAVDAT!BNET = Val(TXTBNET.Text)
  If SAVEFLAG = True Then
    SAVDAT!SYSR = "N"
   Else
    SAVDAT!SYSR = "U"
  End If
  SAVDAT![User] = cUName & ""
  SAVDAT!DVCD = DIVCOD
  SAVDAT!unit = UNCD
  SAVDAT!TRCD = M_TRCD
  SAVDAT!LRNO = Trim(TXTLRNO.Text)
  SAVDAT!LRDT = Format(TXTLRDT.Value, "YYYY/MM/DD")
  SAVDAT!CVBN = Trim(TXTSBILLNO.Text)
  SAVDAT!GATD = Format(TXTSBILLDATE.Value, "YYYY/MM/DD")
  SAVDAT!VHCL = Trim(TXTVHCL)
  SAVDAT!RECSTAT = "A"
  SAVDAT!RORT = Trim(TXTRTORTAX.Text)
  SAVDAT!TTYP = Trim(cmbSelection.Text)
  SAVDAT!BRMK = Trim(TXTRMRK.Text)
  If FrmGRNEntry.SUMMARYTABLE = "GRN" Then
     SAVDAT!GDNCOD = Trim(GDNCODE & "")
  End If
  'IMPORT CASE
  SAVDAT!Currency = Trim(txtCURNCY)
  SAVDAT!EXRATE = Val(txtEXRate)
  SAVDAT!CHAVALUE = Val(txtcha)
  SAVDAT!FRTVALUE = Val(txtfrt)
  SAVDAT!DTYVALUE = Val(txtdty)
  '============================
  
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
      If Trim(SAVDAT.Fields(J).NAME) = Trim(flexBTRM.TextMatrix(i, 0)) & "_INCOST" Then
        SAVDAT.Fields(J).Value = Trim(flexBTRM.TextMatrix(i, 3))
      End If
    Next
  Next
  'SAVDAT!CVBN = ""
  Dim K As Double
  K = 1
  SAVDAT.Update
     
  If SAVDAT.State = 1 Then SAVDAT.Close
  SAVDAT.Open "SELECT * FROM " & TABLENAME & " WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
              "' AND VTYP='IVR' AND DBCD='" & M_DBCD_DIRIVR & "' AND VBNO='" & TXTVBNO & "'", CN, adOpenDynamic, adLockOptimistic
  
  Dim AI As String
  Dim BQ As Double
  Dim CR As Double
  
  i = 1
  For i = 1 To FLEX.Rows - 1
    SAVDAT.AddNew
    SAVDAT!COMP = compPth
    SAVDAT!VTYP = "IVR"
    SAVDAT!SRNO = TXTVBNO
    SAVDAT!SRCH = i
    SAVDAT!VBNO = TXTVBNO.Text
    SAVDAT!chln = Trim(FLEX.TextMatrix(i, 1))
    If Not Trim(FLEX.TextMatrix(i, 2)) = "" Then
      SAVDAT!CHDT = Format(FLEX.TextMatrix(i, 2), "YYYY/MM/DD")
    End If
    SAVDAT!Date = Format(TXTVBDT, "YYYY/MM/DD")
    SAVDAT!dbcd = M_DBCD_DIRIVR
    SAVDAT!CRAC = M_CRAC
    SAVDAT!DRAC = M_DRAC
    SAVDAT!PCOD = M_PCOD
    SAVDAT!DCOD = M_DCOD
    SAVDAT!ICOD = FLEX.TextMatrix(i, 11): AI = FLEX.TextMatrix(i, 11)
    SAVDAT!PCES = Val(FLEX.TextMatrix(i, 7))
    SAVDAT!QNTY = Val(FLEX.TextMatrix(i, 8))
    SAVDAT!GWGT = Val(FLEX.TextMatrix(i, 9))   'BASIC RATE
    
     If UCase(Trim(M_Currency)) <> UCase(Trim(txtCURNCY)) Then
        If Val(txtEXRate) > 0 Then
           SAVDAT!GWGT = Val(FLEX.TextMatrix(i, 9)) * Val(txtEXRate)
        End If
     End If
        
    'FOR NET RATE=============================================================================================
     Dim BASIC_AMT As Double, NET_RAT As Double, GROSS_AMT As Double, QUANTITY As Double, BASIC_RATE As Double
                
     BASIC_AMT = Val(SAVDAT!QNTY) * Val(SAVDAT!GWGT)
     BASIC_RATE = Val(FLEX.TextMatrix(i, 9))
     
     If UCase(Trim(M_Currency)) <> UCase(Trim(txtCURNCY)) Then
        If Val(txtEXRate) > 0 Then
           BASIC_AMT = Val(SAVDAT!QNTY) * Val(FLEX.TextMatrix(i, 9)) * Val(txtEXRate)
           BASIC_RATE = Val(FLEX.TextMatrix(i, 9)) * Val(txtEXRate)
        End If
     End If
     
      GROSS_AMT = Val(TXTITOT)
      QUANTITY = Val(FLEX.TextMatrix(i, 8))
      NET_RAT = 0
      NET_RAT = CALNETRATE(BASIC_AMT, GROSS_AMT, BASIC_RATE, QUANTITY)
     '=============================================================================================
      SAVDAT!RATE = NET_RAT
      SAVDAT!AMNT = Val(SAVDAT!RATE) * Val(FLEX.TextMatrix(i, 8))
    '==============================================================================================
    
     SAVDAT!QORP = "Q"
     SAVDAT![User] = cUName
     
     If SAVEFLAG = True Then
        SAVDAT!SYSR = "N"
     Else
        SAVDAT!SYSR = "U"
     End If
     
     SAVDAT!OPER = "+"
     SAVDAT!DVCD = DIVCOD
     SAVDAT!unit = UNCD
     SAVDAT!grad = Mid(FLEX.TextMatrix(i, 5), 1, 5)
     
     SAVDAT!ltno = Mid(Trim(FLEX.TextMatrix(i, 4)), 1, 19)
     SAVDAT!MRGN = Mid(Trim(FLEX.TextMatrix(i, 4)), 1, 19)
     
     SAVDAT!COPS = Val(FLEX.TextMatrix(i, 6))
     SAVDAT!TWST = Mid(Trim(FLEX.TextMatrix(i, 16)), 1, 1)
     SAVDAT!RTYP = FLEX.TextMatrix(i, 12)
     SAVDAT!RSRN = FLEX.TextMatrix(i, 13)
     SAVDAT!RSRC = Val(FLEX.TextMatrix(i, 17))
     SAVDAT!RECSTAT = "A"
     SAVDAT!GDNCOD = Trim(GDNCODE)
     SAVDAT!SPECIFICATION = GetSpeci(FLEX.TextMatrix(i, 11))
     SAVDAT.Update
            
    Call MrgnReq
    If MRGN = "Y" Then
            
    If MSTDAT.State = 1 Then MSTDAT.Close
    MSTDAT.Open "SELECT * FROM MRGMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND MRGN='" & FLEX.TextMatrix(i, 4) & "' AND ICOD = '" & Trim(FLEX.TextMatrix(i, 11)) & "'", CN, adOpenDynamic, adLockOptimistic
    If MSTDAT.EOF Then
      MSTDAT.AddNew
      MSTDAT!COMP = compPth
      MSTDAT!unit = UNCD
      MSTDAT!MRGN = Mid(UCase(FLEX.TextMatrix(i, 4)), 1, 19)
      MSTDAT!PCOD = Trim(M_PCOD)
      MSTDAT!ICOD = FLEX.TextMatrix(i, 11)
      MSTDAT.Update
    End If
    End If
  Next
 
  '------------FIFO----------------------
  If TABLENAME = "STORETRAN" Then
     Call SetItemInfo
  End If
  '======================================
    
 'EXCISE DETAILS=================================================================================================================
    
  Set EXCISE = New ADODB.Recordset
  If EXCISE.State = 1 Then EXCISE.Close
  
  CN.Execute "DELETE FROM EGPMAN WHERE  COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DBCD='" & _
               M_DBCD_DIRIVR & "' AND VTYP='IVR' AND VBNO='" & TXTVBNO & "' AND RECSTAT<>'D'"
  
  CN.Execute "DELETE FROM EGPMAN WHERE  COMP='" & compPth & "' AND UNIT='" & UNCD & _
             "' AND DBCD='RG23-C' AND VTYP='IVR' AND VBNO='" & TXTVBNO & "' AND RECSTAT<>'D'"
  
  EXCISE.Open "SELECT * FROM EGPMAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DBCD='" & _
               M_DBCD_DIRIVR & "' AND VTYP='IVR' AND VBNO='" & TXTVBNO & "' AND RECSTAT<>'D'", CN, adOpenDynamic, adLockOptimistic
               
  If EXCISE.EOF Then
     EXCISE.AddNew
  End If
  
  EXCISE!COMP = compPth
  EXCISE!unit = UNCD
  EXCISE!dbcd = M_DBCD_DIRIVR
  EXCISE!VTYP = "IVR"
  EXCISE!VBNO = Trim(TXTVBNO.Text)
  EXCISE!SRNO = TXTVBNO
  EXCISE!SRCH = 1
  EXCISE!Date = Format(TXTVBDT, "YYYY/MM/DD")
  EXCISE!CRAC = M_CRAC & ""
  EXCISE!DRAC = M_DRAC & ""
  EXCISE!ICOD = FLEX.TextMatrix(1, 11)
  
  ' EXCISE!VBNO = TXTVBNO
  EXCISE!chln = Trim(TXTSCHLN)
  EXCISE!CHDT = Format(TXTSCHLNDT, "YYYY/MM/DD")
  
  EXCISE!PCES = Val(TXTTPCS)
  EXCISE!QNTY = Val(TXTTQTY)
  EXCISE!AMNT = Val(TXTITOT)
  EXCISE!ITOT = Val(TXTITOT)
  EXCISE!BADJ = Val(TXTBNET) - Val(TXTITOT)
  EXCISE!BNET = Val(TXTBNET)
  EXCISE!TTYP = cmbSelection.Text
  EXCISE!RECSTAT = "A"
  EXCISE!EXTRA3 = "True"
  
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
  
  If cmbSelection = "RG23-C" Then
     If excperc > 0 Then
        EXCISE!CENVAT = Round((EXCISE!CENVAT * excperc) / 100, 2)
        EXCISE!EDUCESS = Round((EXCISE!EDUCESS * excperc) / 100, 2)
        EXCISE!H_ED_CESS = Round((EXCISE!H_ED_CESS * excperc) / 100, 2)
     End If
  End If
  
  EXCISE.Update
    
  If cmbSelection = "RG23-C" Then
    If EXCISE.State = 1 Then EXCISE.Close
    EXCISE.Open "SELECT * FROM EGPMAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DBCD='RG23-C' AND VTYP='IVR' AND VBNO='" & TXTVBNO & "' AND RECSTAT<>'D'", CN, adOpenDynamic, adLockOptimistic
               
    If EXCISE.EOF Then
        EXCISE.AddNew
     End If
  
     EXCISE!COMP = compPth
     EXCISE!unit = UNCD
     EXCISE!dbcd = "RG23-C"
     EXCISE!VTYP = "IVR"
     EXCISE!VBNO = TXTVBNO
     EXCISE!SRNO = TXTVBNO
     EXCISE!SRCH = 1
     EXCISE!Date = Format(FEDT + 1, "YYYY/MM/DD")
       
     EXCISE!CRAC = M_CRAC & ""
     EXCISE!DRAC = M_DRAC & ""
     EXCISE!ICOD = FLEX.TextMatrix(1, 11)
  
    ' EXCISE!VBNO = TXTVBNO
     EXCISE!chln = Trim(TXTSCHLN)
     EXCISE!CHDT = Format(FEDT + 1, "YYYY/MM/DD")
  
     EXCISE!PCES = Val(TXTTPCS)
     EXCISE!QNTY = Val(TXTTQTY)
     EXCISE!AMNT = Val(TXTITOT)
     EXCISE!ITOT = Val(TXTITOT)
     EXCISE!BADJ = Val(TXTBNET) - Val(TXTITOT)
     EXCISE!BNET = Val(TXTBNET)
     EXCISE!TTYP = cmbSelection.Text
     EXCISE!RECSTAT = "A"
     EXCISE!EXTRA3 = "True"
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
     
     If excperc > 0 Then
        EXCISE!CENVAT = Round((EXCISE!CENVAT * (100 - excperc)) / 100, 2)
        EXCISE!EDUCESS = Round((EXCISE!EDUCESS * (100 - excperc)) / 100, 2)
        EXCISE!H_ED_CESS = Round((EXCISE!H_ED_CESS * (100 - excperc)) / 100, 2)
     End If
  
     EXCISE.Update
  End If
    
  '======================================================================================================

  
  If MSTDAT.State = 1 Then MSTDAT.Close
  MSTDAT.Open "SELECT ISNULL(SUM(PCES),0) AS TPCS,ISNULL(SUM(QNTY),0) AS TQTY FROM " & TABLENAME & _
              " WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DBCD = '" & M_DBCD_DIRIVR & _
              "' AND VTYP='IVR' AND VBNO='" & TXTVBNO & "'", CN, adOpenDynamic, adLockOptimistic
  
  If Not MSTDAT.EOF Then
    CN.Execute "UPDATE " & SUMMARYTABLE & " SET TPCS='" & MSTDAT!TPCS & "', TQTY='" & MSTDAT!TQTY & _
               "' WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DBCD = '" & M_DBCD_DIRIVR & _
               "' AND VTYP='IVR' AND VBNO='" & TXTVBNO & "'"
  End If
  
  If chkReturnable.Value = 1 Then
    Call SetReturnableEntry(M_PCOD)
  End If
  
  'UPDATE VOUCHER TYPE MASTER
  If SAVEFLAG = True Then
     Call SetSRNO(TXTVBNO, "IVR", M_DBCD_DIRIVR)
     
     'UPDATE VOUCHER TYPE MASTER
      CN.Execute "UPDATE SERIALMASTER SET LEDT = '" & Format(TXTVBDT, "MM/DD/YYYY") & _
                 "' WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
                 "' AND CODE='" & M_DBCD_DIRIVR & "' AND FYCD='" & FYCD & "' AND VTYP='IVR' "
                 
      TXTVBDT.MinDate = Format(TXTVBDT, "DD/MM/YYYY")

  End If
  
  '-----------------------
  'DAILYSTAUS ENTRY
  If SAVEFLAG = True Then
  Call DAILYSTATUS("IVR", M_DRAC, M_DBCD_DIRIVR, Val(TXTTQTY), TXTVBNO, Val(TXTITOT), cUName, "N", Now, TXTVBDT)
  Else
  Call DAILYSTATUS("IVR", M_DRAC, M_DBCD_DIRIVR, Val(TXTTQTY), TXTVBNO, Val(TXTITOT), cUName, "M", Now, TXTVBDT)
  End If
 '------------------------
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
  Dim m_rtyp As String
  Dim m_rsrn As String
  
  CN.Execute "DELETE FROM EGPMAN WHERE COMP='" & compPth & "' AND UNIT ='" & UNCD & "' AND VTYP='IVR' AND DBCD='" & M_DBCD_DIRIVR & "' AND LTRIM(RTRIM(VBNO))='" & TXTVBNO & "'"
  CN.Execute "DELETE FROM " & TABLENAME & " WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND VTYP='IVR'  AND DBCD='" & M_DBCD_DIRIVR & "' AND VBNO='" & TXTVBNO & "'"
  CN.Execute "DELETE FROM " & SUMMARYTABLE & " WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND VTYP='IVR' AND DBCD='" & M_DBCD_DIRIVR & "' AND VBNO='" & TXTVBNO & "'"
  CN.Execute "DELETE FROM PKGSTK WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
             "' AND VTYP='IVR' AND CHLN='" & TXTVBNO & "' AND DBCD='" & M_DBCD_DIRIVR & "' "
   
  'FIFO
  If TABLENAME = "STORETRAN" Then
     CN.Execute "DELETE FROM GRNTRAN WHERE COMP='" & compPth & "' AND UNIT ='" & UNCD & _
     "' AND DBCD='" & M_DBCD_DIRIVR & "' AND VBNO='" & TXTVBNO & "'"
  End If
  '-----------------------------------------------------
     
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
  TXTVHCL.SelLength = Len(TXTVHCL)
End Sub

Private Sub flex_KeyPress(KeyAscii As Integer)
  On Error GoTo LAST
  FLEX.TextMatrix(FLEX.ROW, 0) = FLEX.ROW
  Dim ALLOW_KEY As Boolean
  Dim FWD_COL As Boolean
  Dim ENTER_PRESS As Boolean
  Dim MSTDAT As New ADODB.Recordset
  Set MSTDAT = New ADODB.Recordset
  FWD_COL = False
  ALLOW_KEY = False
  
  If FLEX.COL = 6 Or FLEX.COL = 7 Or FLEX.COL = 8 Or FLEX.COL = 9 Then
    If InStr(1, FLEX.TextMatrix(FLEX.ROW, FLEX.COL), ".") > 0 And KeyAscii = 46 Then
      KeyAscii = 0
      Exit Sub
    End If
  End If
  
  If FLEX.COL = 8 Then
    If InStr(1, FLEX.TextMatrix(FLEX.ROW, FLEX.COL), "-") > 0 And KeyAscii = 45 Then
      KeyAscii = 0
      Exit Sub
    End If
  End If
  
  If Emptycell = True And (Not KeyAscii = 13) Then
     FLEX.TextMatrix(FLEX.ROW, FLEX.COL) = Empty
     Emptycell = False
  End If
  
  Select Case FLEX.COL
   Case 1
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
    Else
      ALLOW_KEY = False
    End If
   Case 2
    If KeyAscii >= 48 And KeyAscii <= 57 Then             ' 0- 9
      ALLOW_KEY = True
    ElseIf KeyAscii = 47 Then                              '/
      ALLOW_KEY = True
    Else
      ALLOW_KEY = False
    End If
   Case 3
    If KeyAscii = vbKeyF2 Or (KeyAscii = vbKeyReturn And Trim(FLEX.TextMatrix(FLEX.ROW, 3)) = Empty) Then
        NEW_VISIBLE = True
        M_DESC = Empty
        Key = Empty
        FLEX.TextMatrix(FLEX.ROW, 3) = SearchList1("select TOP 20 code,name from itmmst", 0, FLEX.TextMatrix(FLEX.ROW, 3), "SELECT ITEM FROM LIST")
        If key_PressNew = True Then
           M_DESC = ""
           Key = ""
           FLEX.TextMatrix(FLEX.ROW, 3) = ""
           frm_Item.Show
        End If
        ALLOW_KEY = True
     End If
   Case 4
   
    If KeyAscii >= 48 And KeyAscii <= 57 Then             ' 0- 9
      ALLOW_KEY = True
    ElseIf KeyAscii >= 65 And KeyAscii <= 90 Then         ' A-Z
      ALLOW_KEY = True
    ElseIf KeyAscii >= 97 And KeyAscii <= 122 Then         'a-z
      ALLOW_KEY = True
    ElseIf KeyAscii = 46 Then                              '.
      ALLOW_KEY = True
    ElseIf KeyAscii = 45 Then
      ALLOW_KEY = True
    ElseIf KeyAscii = 47 Then                              '/
      ALLOW_KEY = True
    Else
      ALLOW_KEY = False
    End If
    If M_COMPBILL = "VFL" Then
      ALLOW_KEY = False
    End If
   Case 5
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
    Else
      ALLOW_KEY = False
    End If
   Case 6
    If KeyAscii >= 48 And KeyAscii <= 57 Then             ' 0- 9
      ALLOW_KEY = True
    ElseIf KeyAscii = 46 Then                              '.
      ALLOW_KEY = True
    Else
      ALLOW_KEY = False
    End If
   Case 7
    If KeyAscii >= 48 And KeyAscii <= 57 Then             ' 0- 9
      ALLOW_KEY = True
    Else
      ALLOW_KEY = False
    End If
   Case 8
    If KeyAscii >= 48 And KeyAscii <= 57 Then             ' 0- 9
      ALLOW_KEY = True
    ElseIf KeyAscii = 46 Then                              '.
      ALLOW_KEY = True
    ElseIf KeyAscii = 45 Then                              '-
      ALLOW_KEY = True
    Else
      ALLOW_KEY = False
    End If
   Case 9
    If KeyAscii >= 48 And KeyAscii <= 57 Then             ' 0- 9
      ALLOW_KEY = True
    ElseIf KeyAscii = 46 Then                              '.
      ALLOW_KEY = True
    Else
      ALLOW_KEY = False
    End If
   Case 10
    ALLOW_KEY = False
   Case 16
    If Chr(KeyAscii) = "S" Or Chr(KeyAscii) = "Z" Or Chr(KeyAscii) = "0" Or Chr(KeyAscii) = " " Then
      ALLOW_KEY = True
     Else
      ALLOW_KEY = False
    End If
  End Select
  If KeyAscii = vbKeyReturn Then
    ENTER_PRESS = True
   Else
    ENTER_PRESS = False
  End If
  If KeyAscii = 8 Then
    Dim lnth As Double
    lnth = Len(FLEX.TextMatrix(FLEX.ROW, FLEX.COL))
    If lnth > 0 Then
      FLEX.TextMatrix(FLEX.ROW, FLEX.COL) = Mid(FLEX.TextMatrix(FLEX.ROW, FLEX.COL), 1, lnth - 1)
      Exit Sub
    End If
  End If
  If ALLOW_KEY = False Then
    If ENTER_PRESS = True Then
     Else
      KeyAscii = 0
      Exit Sub
    End If
  End If
  
  If ALLOW_KEY = True Then
    If ENTER_PRESS = False Then
      FLEX.TextMatrix(FLEX.ROW, FLEX.COL) = Trim(FLEX.TextMatrix(FLEX.ROW, FLEX.COL)) + Chr(KeyAscii)
      
      Select Case FLEX.COL
      Case 9
          FLEX.TextMatrix(FLEX.ROW, 10) = Format(Val(FLEX.TextMatrix(FLEX.ROW, 9)) * Val(FLEX.TextMatrix(FLEX.ROW, 8)), "#########.00")
          
          If UCase(Trim(M_Currency)) <> UCase(Trim(txtCURNCY)) Then
             If Val(txtEXRate) > 0 Then
                FLEX.TextMatrix(FLEX.ROW, 10) = Format(Val(FLEX.TextMatrix(FLEX.ROW, 9)) * Val(txtEXRate) * Val(FLEX.TextMatrix(FLEX.ROW, 8)), "####.00")
             End If
          End If
          
      End Select
      
    End If
  End If
  FWD_COL = False
  If ENTER_PRESS = True Then
    Select Case FLEX.COL
     Case 1
      FWD_COL = True
     Case 2
      If Len(FLEX.TextMatrix(FLEX.ROW, FLEX.COL)) = 10 Then
        If IsDate(CDate(FLEX.TextMatrix(FLEX.ROW, FLEX.COL))) Then
          FWD_COL = True
         Else
          FWD_COL = False
        End If
       Else
        FWD_COL = False
      End If
     Case 3
      If MSTDAT.State = 1 Then MSTDAT.Close
      MSTDAT.Open "select * from itmmst where name='" & FLEX.TextMatrix(FLEX.ROW, FLEX.COL) & "'", CN, adOpenDynamic, adLockOptimistic
      If MSTDAT.EOF Then
        FWD_COL = False
       Else
        FLEX.TextMatrix(FLEX.ROW, 11) = MSTDAT!CODE
        FWD_COL = True
      End If
     Case 4
      FWD_COL = True
     Case 5
      FWD_COL = True
     Case 6
      If IsNumeric(FLEX.TextMatrix(FLEX.ROW, FLEX.COL)) Then
        FWD_COL = True
       Else
        FWD_COL = False
      End If
     Case 7
      If IsNumeric(FLEX.TextMatrix(FLEX.ROW, FLEX.COL)) Then
        FWD_COL = True
       Else
        FWD_COL = False
      End If
     Case 8
      If IsNumeric(FLEX.TextMatrix(FLEX.ROW, FLEX.COL)) Then
        FWD_COL = True
       Else
        FWD_COL = False
      End If
     Case 9
      If IsNumeric(FLEX.TextMatrix(FLEX.ROW, FLEX.COL)) Then
        FWD_COL = True
       Else
        FWD_COL = False
      End If
     Case 10
      If IsNumeric(FLEX.TextMatrix(FLEX.ROW, FLEX.COL)) Then
        FWD_COL = True
       Else
        FWD_COL = False
      End If
     Case 16
      FWD_COL = True
    End Select
    If FWD_COL = True Then
'-------------------------------------------------------------------------------
'1. FOR MERGE NO. REQUIRED OR NOT
    If FLEX.COL = 4 Then
        Call MrgnReq
        If MRGN = "Y" Then
           If FLEX.TextMatrix(FLEX.ROW, 4) = Empty Then
           MsgBox "Lot No. Empty", vbOKOnly
           FLEX.ROW = FLEX.ROW
           FLEX.COL = FLEX.COL
           FLEX.SetFocus
           Exit Sub
        End If
      End If
      End If
'2 . IF SPECIFICATION BOX + QUANTITY
     If FLEX.COL = 7 Then
        SPECI = GetSpeci(FLEX.TextMatrix(FLEX.ROW, 11))
        If SPECI = 0 Then
           If Val(FLEX.TextMatrix(FLEX.ROW, 7)) <= 0 And Val(FLEX.TextMatrix(FLEX.ROW, 8)) <= 0 Then
              MsgBox "BOX/Pcs. Empty ", vbOKOnly
              FLEX.ROW = FLEX.ROW
              FLEX.COL = FLEX.COL
              FLEX.SetFocus
              Exit Sub
          End If
        End If
        End If
        
 '------------------------------------------------------------------------------------------------
 '3 . IF SPECIFICATION BOX + COPS + QUANTITY
      If FLEX.COL = 6 Then
      SPECI = GetSpeci(FLEX.TextMatrix(FLEX.ROW, 11))
        If SPECI = 2 Then
           If Val(FLEX.TextMatrix(FLEX.ROW, 6)) <= 0 Then
              MsgBox "Please Enter Cops ", vbOKOnly
              FLEX.ROW = FLEX.ROW
              FLEX.COL = FLEX.COL
              FLEX.SetFocus
              Exit Sub
          End If
          End If
          End If
          
    If FLEX.COL = 7 Then
      SPECI = GetSpeci(FLEX.TextMatrix(FLEX.ROW, 11))
        If SPECI = 2 Then
           If Val(FLEX.TextMatrix(FLEX.ROW, 7)) <= 0 Then
              MsgBox "Please Enter Box/Pieces ", vbOKOnly
              FLEX.ROW = FLEX.ROW
              FLEX.COL = FLEX.COL
              FLEX.SetFocus
              Exit Sub
          End If
          End If
          End If
          
   If FLEX.COL = 8 Then
      SPECI = GetSpeci(FLEX.TextMatrix(FLEX.ROW, 11))
        If SPECI = 2 Then
           If Val(FLEX.TextMatrix(FLEX.ROW, 8)) <= 0 Then
              MsgBox "Please Enter Quantity", vbOKOnly
              FLEX.ROW = FLEX.ROW
              FLEX.COL = FLEX.COL
              FLEX.SetFocus
              Exit Sub
          End If
          End If
          End If
          
  '-----------------------------------------------------------------------------------------------------------------------------------
  '4 IF SPECIFICATION COPS + QUANTITY
       If FLEX.COL = 6 Then
          SPECI = GetSpeci(FLEX.TextMatrix(FLEX.ROW, 11))
        If SPECI = 3 Then
           If Val(FLEX.TextMatrix(FLEX.ROW, 6)) <= 0 Then
              MsgBox "Please Enter Cops", vbOKOnly
              FLEX.ROW = FLEX.ROW
              FLEX.COL = FLEX.COL
              FLEX.SetFocus
              Exit Sub
          End If
          End If
          End If
          
          
          
     If FLEX.COL = 8 Then
          SPECI = GetSpeci(FLEX.TextMatrix(FLEX.ROW, 11))
        If SPECI = 3 Then
           If Val(FLEX.TextMatrix(FLEX.ROW, 8)) <= 0 Then
              MsgBox "Please Enter Quantity ", vbOKOnly
              FLEX.ROW = FLEX.ROW
              FLEX.COL = FLEX.COL
              FLEX.SetFocus
              Exit Sub
          End If
          End If
          End If
               
          
    '5 IF SPECIFICATION QUANTITY
       If FLEX.COL = 8 Then
          SPECI = GetSpeci(FLEX.TextMatrix(FLEX.ROW, 11))
        If SPECI = 1 Then
           If Val(FLEX.TextMatrix(FLEX.ROW, 8)) <= 0 Then
              MsgBox "Quantity is Empty ", vbOKOnly
              FLEX.ROW = FLEX.ROW
              FLEX.COL = FLEX.COL
              FLEX.SetFocus
              Exit Sub
          End If
          End If
          End If
    '------------------------------------------------------------------------------------

      If FLEX.COL = 16 Then
        'Allowed to add row with msgbox
        'Check all the cell are filled
        Call CHKFLEX
        If Not CHK_FLX Then
          MsgBox "Invalid Data in item details "
          FLEX.ROW = FLXROW
          FLEX.COL = FLXCOL
          FLEX.SetFocus
          Exit Sub
        End If
        'FRM_ITMDTLGRN.Show 1
        Dim AYS
        AYS = MsgBox("Want to Add More Item ", vbYesNo)
        If AYS = vbYes Then
          FLEX.Rows = FLEX.Rows + 1
          FLEX.ROW = FLEX.Rows - 1
          FLEX.COL = 1
          FLEX.TextMatrix(FLEX.ROW, 1) = FLEX.TextMatrix(FLEX.ROW - 1, 1)
          FLEX.TextMatrix(FLEX.ROW, 2) = FLEX.TextMatrix(FLEX.ROW - 1, 2)
         Else
          If flexBTRM.Enabled = True Then
            flexBTRM.SetFocus
            
           Else
            Call calADLS
            If TXTLRNO.Visible And TXTLRNO.Enabled Then
               TXTLRNO.SetFocus
            Else
               cmdSave.SetFocus
            End If
          End If
          Exit Sub
        End If
       Else
        If FLEX.COL = 9 Then
          FLEX.TextMatrix(FLEX.ROW, 10) = Format(Val(FLEX.TextMatrix(FLEX.ROW, 9)) * Val(FLEX.TextMatrix(FLEX.ROW, 8)), "#########.00")
          
          If UCase(Trim(M_Currency)) <> UCase(Trim(txtCURNCY)) Then
             If Val(txtEXRate) > 0 Then
                FLEX.TextMatrix(FLEX.ROW, 10) = Format(Val(FLEX.TextMatrix(FLEX.ROW, 9)) * Val(txtEXRate) * Val(FLEX.TextMatrix(FLEX.ROW, 8)), "####.00")
             End If
          End If
          
          FLEX.COL = 16
          
         Else
          FLEX.COL = FLEX.COL + 1
        End If
      End If
      Emptycell = True
    End If
  End If
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
    chkitm = FLEX.TextMatrix(FLXR, 11)
    SPECI = GetSpeci(chkitm)
    If MSTDAT.State = 1 Then MSTDAT.Close
    MSTDAT.Open "SELECT * FROM ITMMST WHERE CODE='" & chkitm & "'", CN, adOpenDynamic, adLockOptimistic
    If MSTDAT.EOF Then
       CHK_FLX = False
       FLXROW = FLXR
       FLXCOL = 3
       Exit For
    End If
    Call MrgnReq
    If MRGN = "Y" Then
    If FLEX.TextMatrix(FLXR, 4) = Empty Then
      MsgBox "Please Enter Lot No.", vbOKOnly
      CHK_FLX = False
      FLXROW = FLXR
      FLXCOL = 4
      Exit For
      End If
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
    If Not IsNumeric(FLEX.TextMatrix(FLXR, 8)) Then
       CHK_FLX = False
       FLXROW = FLXR
       FLXCOL = 8
       Exit For
    End If
    If Not IsNumeric(FLEX.TextMatrix(FLXR, 9)) Then
       CHK_FLX = False
       FLXROW = FLXR
       FLXCOL = 9
       Exit For
    End If
    If Not IsNumeric(FLEX.TextMatrix(FLXR, 10)) Then
       CHK_FLX = False
       FLXROW = FLXR
       FLXCOL = 10
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
  DLYSTA!dbcd = ""
  DLYSTA!QNTY = Val(TXTTQTY)
  DLYSTA!VBNO = TXTVBNO & ""
  DLYSTA!AMNT = Val(TXTBNET)
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
  DLYSTA!dbcd = ""
  DLYSTA!QNTY = Val(TXTTQTY)
  DLYSTA!VBNO = TXTVBNO & ""
  DLYSTA!AMNT = Val(TXTBNET)
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

Dim NUMERATOR As Double
NUMERATOR = (STKQNTY * WGTRATE) + (QTY * RATE)
If NUMERATOR <> 0 Then
   WGTRATE = NUMERATOR / (STKQNTY + QTY)
Else
   WGTRATE = 0
End If

If Trim(WGTRS!LAST_PURDATE) <> Format(Now, "DD/MM/YYYY") And Not SAVEFLAG Then
   CN.Execute "UPDATE ITMMST SET WEIGHTEDRATE = " & WGTRATE & " WHERE CODE = '" & ICOD & "'", L
   Exit Sub
End If
'22/05/2010
CN.Execute "UPDATE ITMMST SET WEIGHTEDRATE = " & WGTRATE & " ,PURR = " & RATE & ",LAST_PURDATE = '" & Format(TXTVBDT.Value, "YYYY/MM/DD") & "' WHERE CODE = '" & ICOD & "'", L
  
End Sub

Private Sub SetItemInfo()
On Error GoTo LAST
Dim INDEX As Long
Dim SQL As String
Dim RATE As Double

With FLEX
 
For INDEX = 1 To .Rows - 1

    'FOR NET RATE=============================================================================================
     Dim BASIC_AMT As Double, NET_RAT As Double, GROSS_AMT As Double, QUANTITY As Double, BASIC_RATE As Double
                
     BASIC_AMT = Val(FLEX.TextMatrix(INDEX, 8)) * Val(FLEX.TextMatrix(INDEX, 9))
     BASIC_RATE = Val(FLEX.TextMatrix(INDEX, 9))
     
     If UCase(Trim(M_Currency)) <> UCase(Trim(txtCURNCY)) Then
        If Val(txtEXRate) > 0 Then
           BASIC_AMT = Val(FLEX.TextMatrix(INDEX, 8)) * Val(FLEX.TextMatrix(INDEX, 9)) * Val(txtEXRate)
           BASIC_RATE = Val(FLEX.TextMatrix(INDEX, 9)) * Val(txtEXRate)
        End If
     End If
     
     
     
     GROSS_AMT = Val(TXTITOT)
     QUANTITY = Val(FLEX.TextMatrix(INDEX, 8))
     NET_RAT = 0
     NET_RAT = CALNETRATE(BASIC_AMT, GROSS_AMT, BASIC_RATE, QUANTITY)
    '==========================================================================================================
    
    SQL = "INSERT INTO GRNTRAN([COMP],[UNIT],[VTYP],[VBNO],[DBCD],[SRCH],DATE,[ICOD],[RATE],[GRN_QNTY],[NETRATE],[BAL_QNTY],MRGN)"
    SQL = SQL & " VALUES('" & compPth & "','" & UNCD & "','IVR','" & TXTVBNO & _
    "','" & M_DBCD_DIRIVR & "','" & INDEX & "','" & Format(TXTVBDT, "yyyy-MM-dd hh:mm:ss") & _
    "','" & Trim(.TextMatrix(INDEX, 11)) & "','" & NET_RAT & "','" & Val(.TextMatrix(INDEX, 8)) & _
    "','" & NET_RAT & "','" & Val(.TextMatrix(INDEX, 8)) & "','" & Trim(.TextMatrix(INDEX, 4)) & "')"
    
    CN.Execute SQL
    Next INDEX
 
End With
Exit Sub
LAST:
MsgBox ERR.Description
End Sub


Private Sub TXTPALLET_GotFocus(): TxtPallet.BackColor = RGB(BRED, BGREEN, BBLUE): End Sub
Private Sub TxtPallet_LostFocus(): TxtPallet.BackColor = vbWhite: End Sub

Private Sub txtCops_GotFocus(): txtCops.BackColor = RGB(BRED, BGREEN, BBLUE): End Sub
Private Sub txtCops_LostFocus(): txtCops.BackColor = vbWhite: End Sub

Private Sub txtPly_GotFocus(): txtPly.BackColor = RGB(BRED, BGREEN, BBLUE): End Sub
Private Sub txtPly_LostFocus(): txtPly.BackColor = vbWhite: End Sub

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

FLEXPLY.Cols = 5

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
       'If FLEXPLY.COL > 2 Then FLEXPLY.ColWidth(FLEXPLY.COL - 2) = 0
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
     FLEXPLY.CellBackColor = vbWhite
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


Private Sub SetReturnableEntry(PCOD As String)
  '----------------------------------
  'RETURNABLE COPS,PALLET & PLY
  '----------------------------------
  Dim RETRS As ADODB.Recordset
  Set RETRS = New ADODB.Recordset
  If RETRS.State = 1 Then RETRS.Close
  If RETRS.State = 1 Then RETRS.Close
  RETRS.Open "SELECT * FROM PKGSTK WHERE COMP='" & compPth & "' AND VTYP='IVR' AND UNIT='" & UNCD & _
              "'  AND CHLN='" & TXTSCHLN & "' AND RECSTAT='A'", CN, adOpenDynamic, adLockOptimistic
  If RETRS.EOF Then
     RETRS.AddNew
  End If
  
  
  RETRS!COMP = compPth
  RETRS!unit = UNCD
  RETRS!DVCD = DIVCOD
  RETRS!dbcd = "000001"
  RETRS!VTYP = "IVR"
  RETRS!chln = TXTVBNO
  RETRS!Date = Format(TXTVBDT, "YYYY/MM/DD")
  RETRS!PCHLN = TXTSCHLN
  RETRS!PCHDT = Format(TXTSCHLNDT, "YYYY/MM/DD")
  RETRS!PCOD = PCOD
  RETRS!OPER = "+"
  RETRS!TOPPLY = Val(TxtPallet)
  RETRS!BOTTOMPLY = Val(TxtPallet)
  RETRS!QNTY = Val(txtCops)
  RETRS!BRMK = TXTRMRK
  RETRS!RECSTAT = "A"
  
'PLY UPDATION COMMON FOR BOTH SAVE AND EDIT


 Dim i As Long, J As Long
  i = 0
  For i = 1 To FLEXPLY.Cols - 1
    J = 0
    For J = 0 To RETRS.Fields.COUNT - 1
      If Trim(RETRS.Fields(J).NAME) = Trim(FLEXPLY.TextMatrix(0, i)) Then
         RETRS.Fields(J).Value = Val(FLEXPLY.TextMatrix(1, i))
      End If
    Next
 Next
  
'--------------------------------------------------
 RETRS.Update
'--------------------------------------------------
End Sub

Private Function CALNETRATE(BASIC_AMT As Double, GROSS_AMT As Double, BASIC_RATE As Double, QUANTITY As Double) As Double
  Dim ITMRATIO As Double
  CALNETRATE = 0
  CALNETRATE = BASIC_RATE
  If GROSS_AMT > 0 And BASIC_AMT > 0 Then
    ITMRATIO = (BASIC_AMT / GROSS_AMT) * 100
   Else
    ITMRATIO = 0
  End If
  
  Dim IRW As Double
  IRW = 0
  For IRW = 1 To flexBTRM.Rows - 1
    If flexBTRM.TextMatrix(IRW, 3) = "Y" Then
      If ITMRATIO > 0 And Val(txtcha.Text) + Val(txtfrt.Text) + Val(txtdty.Text) + Val(flexBTRM.TextMatrix(IRW, 2)) > 0 Then
        CALNETRATE = CALNETRATE + ((Val(txtcha.Text) + Val(txtfrt.Text) + Val(txtdty.Text) + (Val(flexBTRM.TextMatrix(IRW, 2))) * ITMRATIO) / 100) / QUANTITY
       Else
        CALNETRATE = CALNETRATE
      End If
    End If
  Next
End Function

Private Function GetSpeci(ICOD) As String
GetSpeci = ""
Dim SPECI As String
Dim MRGN As String
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


GETRS.Open "SELECT * FROM ITMMST WHERE CODE = '" & FLEX.TextMatrix(FLEX.ROW, 11) & "'", CN, adOpenDynamic, adLockOptimistic
If Not GETRS.EOF Then
   IGCOD = GETRS!igcd
End If

If SPRS.State = 1 Then SPRS.Close
SPRS.Open "SELECT * FROM IGMMST WHERE CODE = '" & Trim(IGCOD) & "'", CN, adOpenDynamic, adLockOptimistic
If Not SPRS.EOF Then
MRGN = SPRS!MERGE
End If

End Sub

