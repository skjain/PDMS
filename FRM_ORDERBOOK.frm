VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WELCHBUTTON.OCX"
Begin VB.Form frm_ORDERBOOK 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Order Booking"
   ClientHeight    =   6795
   ClientLeft      =   300
   ClientTop       =   1575
   ClientWidth     =   10995
   ControlBox      =   0   'False
   Icon            =   "FRM_ORDERBOOK.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   10995
   Begin VB.TextBox TXTDVCD 
      Height          =   285
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   480
      Width           =   3975
   End
   Begin VB.TextBox TXTUNIT 
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   480
      Width           =   3855
   End
   Begin VB.Frame FRM3 
      Enabled         =   0   'False
      Height          =   735
      Left            =   120
      TabIndex        =   36
      Top             =   5160
      Width           =   10815
      Begin VB.TextBox M_RMRK 
         Height          =   285
         Left            =   5640
         MaxLength       =   250
         TabIndex        =   41
         Top             =   240
         Width           =   2775
      End
      Begin VB.TextBox M_CRDS 
         Height          =   285
         Left            =   1560
         MaxLength       =   3
         TabIndex        =   38
         Top             =   240
         Width           =   495
      End
      Begin WelchButton.lvButtons_H cmdExport 
         Height          =   495
         Left            =   8880
         TabIndex        =   42
         Top             =   120
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   873
         Caption         =   " Expor&t Contract Detail"
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
         Image           =   "FRM_ORDERBOOK.frx":000C
         cBack           =   -2147483633
      End
      Begin VB.Label Label7 
         Caption         =   "Schedule Days"
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
         TabIndex        =   37
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label LBLDUDT 
         Caption         =   "Delivery Date dd/mm/yyyy"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   2280
         TabIndex        =   39
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label9 
         Caption         =   "Remark"
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
         Left            =   4920
         TabIndex        =   40
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame FRM4 
      Height          =   855
      Left            =   3600
      TabIndex        =   0
      Top             =   5880
      Width           =   7335
      Begin WelchButton.lvButtons_H cmdAdd 
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   240
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
         Image           =   "FRM_ORDERBOOK.frx":05A6
         Enabled         =   0   'False
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdEdit 
         Height          =   495
         Left            =   3720
         TabIndex        =   45
         Top             =   240
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
         Image           =   "FRM_ORDERBOOK.frx":0940
         Enabled         =   0   'False
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdDelete 
         Height          =   495
         Left            =   4920
         TabIndex        =   46
         Top             =   240
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
         Image           =   "FRM_ORDERBOOK.frx":0CDA
         Enabled         =   0   'False
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdSave 
         Height          =   495
         Left            =   1320
         TabIndex        =   43
         Top             =   240
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
         Image           =   "FRM_ORDERBOOK.frx":1274
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdCancel 
         Height          =   495
         Left            =   2520
         TabIndex        =   44
         Top             =   240
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
         Image           =   "FRM_ORDERBOOK.frx":180E
         Enabled         =   0   'False
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdExit 
         Height          =   495
         Left            =   6120
         TabIndex        =   47
         Top             =   240
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
         Image           =   "FRM_ORDERBOOK.frx":1C60
         cBack           =   -2147483633
      End
   End
   Begin VB.Frame FRM2 
      Enabled         =   0   'False
      Height          =   3135
      Left            =   120
      TabIndex        =   19
      Top             =   2040
      Width           =   10815
      Begin VB.TextBox txtgrad 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2925
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox M_ARAT 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   7080
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   31
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox RMRK 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   8400
         MaxLength       =   200
         TabIndex        =   33
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox M_RATE 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5760
         MaxLength       =   12
         TabIndex        =   29
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox M_QNTY 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4560
         MaxLength       =   12
         TabIndex        =   27
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox M_INAM 
         Height          =   285
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox M_SRCH 
         Height          =   285
         Left            =   120
         MaxLength       =   2
         TabIndex        =   21
         Top             =   720
         Width           =   555
      End
      Begin MSFlexGridLib.MSFlexGrid ITMFLEX 
         Height          =   1575
         Left            =   120
         TabIndex        =   35
         Top             =   1440
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   2778
         _Version        =   393216
         Cols            =   8
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
      Begin WelchButton.lvButtons_H CMDOK 
         Height          =   375
         Left            =   9720
         TabIndex        =   34
         Top             =   300
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
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
         Image           =   "FRM_ORDERBOOK.frx":21FA
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H CMDITMDEL 
         Height          =   375
         Left            =   9720
         TabIndex        =   49
         Top             =   760
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         Caption         =   "&Drop"
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
         Image           =   "FRM_ORDERBOOK.frx":2794
         cBack           =   -2147483633
      End
      Begin VB.Label LBLCFG 
         Alignment       =   2  'Center
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
         Left            =   3240
         TabIndex        =   24
         Tag             =   "S"
         Top             =   360
         Width           =   615
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00000080&
         X1              =   4440
         X2              =   4440
         Y1              =   240
         Y2              =   1200
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00000080&
         X1              =   6960
         X2              =   6960
         Y1              =   240
         Y2              =   1200
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00000080&
         X1              =   75
         X2              =   9600
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
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
         Left            =   7080
         TabIndex        =   30
         Tag             =   "S"
         Top             =   360
         Width           =   1095
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00000080&
         X1              =   9600
         X2              =   9600
         Y1              =   240
         Y2              =   1200
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00000080&
         X1              =   8280
         X2              =   8280
         Y1              =   240
         Y2              =   1200
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00000080&
         X1              =   5640
         X2              =   5640
         Y1              =   240
         Y2              =   1200
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00000080&
         X1              =   2880
         X2              =   2880
         Y1              =   240
         Y2              =   1200
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00000080&
         X1              =   720
         X2              =   720
         Y1              =   240
         Y2              =   1200
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
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
         Left            =   8400
         TabIndex        =   32
         Tag             =   "S"
         Top             =   360
         Width           =   1095
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00000080&
         Height          =   975
         Left            =   75
         Top             =   240
         Width           =   10695
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Caption         =   "Net Rate"
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
         TabIndex        =   28
         Tag             =   "S"
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Caption         =   "Quantity"
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
         Left            =   4440
         TabIndex        =   26
         Tag             =   "S"
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Caption         =   "Item Name"
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
         Left            =   960
         TabIndex        =   22
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Caption         =   "Sr No."
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
         Left            =   120
         TabIndex        =   20
         Tag             =   "S"
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame FRM1 
      Enabled         =   0   'False
      Height          =   1335
      Left            =   120
      TabIndex        =   48
      Top             =   720
      Width           =   10815
      Begin VB.TextBox TXTFREIGHT 
         Height          =   285
         Left            =   6840
         MaxLength       =   7
         TabIndex        =   12
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox M_TXNM 
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   960
         Width           =   4215
      End
      Begin VB.TextBox M_PORD 
         Height          =   285
         Left            =   9360
         MaxLength       =   30
         TabIndex        =   18
         Top             =   960
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker M_ORDT 
         Height          =   315
         Left            =   9360
         TabIndex        =   16
         Top             =   600
         Width           =   1365
         _ExtentX        =   2408
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
         Format          =   53215233
         CurrentDate     =   39339
      End
      Begin VB.TextBox M_ORDN 
         Height          =   285
         Left            =   9360
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox M_BRNM 
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   600
         Width           =   6615
      End
      Begin VB.TextBox M_DNAM 
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   3000
         Visible         =   0   'False
         Width           =   6495
      End
      Begin VB.TextBox M_PNAM 
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   240
         Width           =   6615
      End
      Begin VB.Label Label8 
         Caption         =   "Freight / KG :"
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
         Left            =   5640
         TabIndex        =   53
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label16 
         Caption         =   "Tax Category"
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
         TabIndex        =   10
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Party Order No."
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
         Left            =   8040
         TabIndex        =   17
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Order Date"
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
         Left            =   8040
         TabIndex        =   15
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Order No."
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
         Left            =   8040
         TabIndex        =   13
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Agent Name"
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
         TabIndex        =   8
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Delivery Party"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   3000
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "A/c Party"
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
         Width           =   1695
      End
   End
   Begin VB.Label Label19 
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1800
      TabIndex        =   62
      Top             =   6480
      Width           =   135
   End
   Begin VB.Label Label18 
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1800
      TabIndex        =   61
      Top             =   6240
      Width           =   135
   End
   Begin VB.Label Label17 
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1800
      TabIndex        =   60
      Top             =   6000
      Width           =   135
   End
   Begin VB.Label LBLLEDBAL 
      Caption         =   "Ledger Balance    "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   59
      Top             =   6480
      Width           =   1815
   End
   Begin VB.Label LBLPNDDO 
      Caption         =   "Pending Order Amt "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   58
      Top             =   6240
      Width           =   1815
   End
   Begin VB.Label LBLLEDBALVAL 
      Alignment       =   1  'Right Justify
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2040
      TabIndex        =   57
      Top             =   6480
      Width           =   1215
   End
   Begin VB.Label LBLPNDDOVAL 
      Alignment       =   1  'Right Justify
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2040
      TabIndex        =   56
      Top             =   6255
      Width           =   1215
   End
   Begin VB.Label LBLCRLIMITVAL 
      Alignment       =   1  'Right Justify
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1920
      TabIndex        =   55
      Top             =   6000
      Width           =   1335
   End
   Begin VB.Label LBLCRLIMIT 
      Caption         =   "Party Credit Limit "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   54
      Top             =   6000
      Width           =   1695
   End
   Begin VB.Label LBLDIVISION 
      Caption         =   "Division Name :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   5520
      TabIndex        =   52
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label LBLUNIT 
      Caption         =   "Unit Name :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   240
      TabIndex        =   51
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label LBLHEAD 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      Caption         =   "SALE ORDER BOOKING MODULE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6480
      TabIndex        =   50
      Top             =   30
      Width           =   4455
   End
End
Attribute VB_Name = "frm_ORDERBOOK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public DIVCODE As String
Public DIVNAME As String
Dim SAVEFLAG As Boolean
Public M_PCOD As String
Public EDIT_ORDQTY As Double

'ASSESABLE RATE
Dim RATECOD As String
Dim ASSRATE As Double
Dim M_CD As Double
Dim BROKERAGE_REQ As Boolean
Dim FREIGHT As Double
Dim VAT As Double
Dim PERVAT As Double
Dim CST As Double
Dim PERCST As Double
Dim EXCISE As Double
Dim PEREXCISE As Double
Dim M_DCOD As String
Dim M_BRCD As String
Public ORDBOK As String
Public ORDDBCD As String
Public RATMASTERREQ As String
Dim M_TRCD As String
Dim M_TXCD As String
Dim M_RATECOD As String
Dim FIL_IGCD As String
Dim FIL_IGNM As String
Dim M_BRKPERC As Double
Dim ROWNO As Long
Dim SWITCH As Boolean

Private Sub cmdAdd_Click()
    Dim Ctrl As Control
    For Each Ctrl In Me
        If TypeOf Ctrl Is TextBox Then
            Ctrl = Replace(Ctrl, "'", "", 1)
        End If
    Next
    Call ClsTxt
    Call btn_sts(True)

    Frm1.Enabled = True
    FRM2.Enabled = True
    FRM3.Enabled = True
    M_SRCH.Enabled = True
    M_PNAM.Enabled = True
    M_PNAM.SetFocus
        
    
    SAVEFLAG = True
    cmdCancel.Cancel = True
    Call GENORDN
    If M_COMPBILL = "SIL" Then
     Else
      M_RMRK = Trim(M_BRNM.Text) + " " + M_ORDN
    End If
End Sub

Public Sub cmdCancel_Click()
   
   cmdExit.Cancel = True
   Call btn_sts(False)
   Call ClsTxt
   Frm1.Enabled = False
   FRM2.Enabled = False
   FRM3.Enabled = False
   cmdAdd.Enabled = False
   cmdEdit.Enabled = False
   cmdDelete.Enabled = False
   
   txtUNIT.Enabled = True
   txtDVCD.Enabled = True
   If txtUNIT.Enabled Then txtUNIT.SetFocus
     
   SAVEFLAG = True
   ClsData (frm_ORDERBOOK)
   ROWNO = ITMFLEX.Rows - 1
   cmdOk.Caption = "&Add"
   SWITCH = False
End Sub

Private Sub cmdDelete_Click()
  If M_USRSECLEVL = "1" Then
    If ReadConfigMaster("000025", 6, "M") = False Then ModuleDeniedMessage: Exit Sub
  End If
  
  On Error GoTo LAST
    
    If M_ORDN = Empty Then
        Call cmdEdit_Click
    Else
    
    End If
    
    If M_ORDN = Empty Then Exit Sub
    
    If cmdSave.Enabled = False Then
        MsgBox "Record Can Not Be Edited / Deleted !!", vbCritical
        Exit Sub
    End If
    
    Set RS = New ADODB.Recordset
    If RS.State = 1 Then RS.Close
    RS.Open "SELECT SUM(DOQTY+DISPATCHQTY+CANCELQTY) AS USEDQTY FROM ORDMAN WHERE COMP='" & compPth & _
            "' AND UNIT='" & FindUnit & "' AND ORDN='" & M_ORDN & "' AND RECSTAT<> 'D'", CN, adOpenKeyset, adLockPessimistic
    If Not RS.EOF Then
        If Val(RS!USEDQTY) > 0 Then
           cmdSave.Enabled = False
           MsgBox "Further Transaction Exists. Record Can Not Be Deleted.", vbCritical
           btn_sts (False)
           Call ClsTxt
           cmdAdd.SetFocus
           Exit Sub
        End If
    End If
    
    Dim AYS
    
    AYS = MsgBox("Are You Sure To Delete the Data ", vbQuestion + vbYesNo, "Remove This ?")
    
    If AYS = vbYes Then
        CN.BeginTrans
        CN.Execute "UPDATE ordman SET RECSTAT='D' where  COMP='" & compPth & "' AND UNIT='" & FindUnit & "' AND DBCD='" & ORDDBCD & "' AND ORDN='" & M_ORDN & "'"
        CN.Execute "INSERT INTO DAILYSTAT (COMP,VTYP,SRNO,PCOD,DBCD,VBNO,QNTY,AMNT,CUSR,DTTM,ACTN) VALUES('" & compPth & "','SOB','XXXXXXXXXXXXX','" & M_PNAM & "',NULL,'" & M_ORDN & "',0,0,'" & cUName & "','" & Format(Now, "MM/dd/yyyy HH:MM:SS AMPM") & "','D')"
        CN.CommitTrans
    End If
    
    btn_sts (False)
    
    Call ClsTxt
    
    cmdAdd.SetFocus
  
  Exit Sub
LAST:
  
  MsgBox ERR.Description
  CN.RollbackTrans
End Sub

Private Sub cmdEdit_Click()
  If M_USRSECLEVL = "1" Then
    If ReadConfigMaster("000025", 5, "M") = False Then ModuleDeniedMessage: Exit Sub
  End If
  
  On Error GoTo LAST
  SAVEFLAG = False
  
  EDIT_ORDQTY = 0
  If EXP_REQ Then
    frm_ORDLIST.EXPORTREQ = "Y"
  Else
    frm_ORDLIST.EXPORTREQ = "N"
  End If
  
  frm_ORDLIST.Show 1
  
  If ITMFLEX.TextMatrix(1, 2) <> Empty Then Call ITMFLEX_Click
  If M_PNAM.Enabled = True Then
    M_PNAM.SetFocus
  End If
  If cmdSave.Enabled Then
    cmdDelete.Enabled = True
  End If
  If ITMFLEX.TextMatrix(1, 2) <> Empty Then Call ITMFLEX_Click
  Exit Sub

LAST:
  ErrNumber = ERR.Number
  ErrMessage = ERR.Description
  frm_ErrorHandler.Show vbModal
End Sub
Private Sub CMDEXIT_Click()
Dim ANS As Integer
    
    If cmdSave.Enabled Then
        ANS = vbYes
        'ANS = MsgBox("Are You Sure ? Want To Discard Changes ?", vbQuestion + vbYesNo, "Exit Without Changes ?")
        If ANS = vbYes Then
            Unload Me
        ElseIf ANS = VBNO Then
            Exit Sub
        End If
    End If
    
    Unload Me
    
End Sub

Private Sub cmdexport_Click()
  FRM_TRNEXPORD.Show 1
  If cmdSave.Enabled Then cmdSave.SetFocus
End Sub

Private Sub CMDITMDEL_Click()

'============================================================================
'In Case of Export : User Can Change Shade But Can't Drop Any Row Of Flex
'Because Ordtrn already Set the OSRC for each Row
If SAVEFLAG = False And EXP_REQ Then
   MsgBox "In Export Order, You Can Change Shade/Quantity But Can not Remove Row.", vbCritical
   Exit Sub
End If
'=============================================================================
Dim CURSOR As Long
Dim J As Long

For J = ROWNO To ITMFLEX.Rows - 2
 ITMFLEX.TextMatrix(J, 1) = ITMFLEX.TextMatrix(J + 1, 1)
 ITMFLEX.TextMatrix(J, 2) = ITMFLEX.TextMatrix(J + 1, 2)
 ITMFLEX.TextMatrix(J, 3) = ITMFLEX.TextMatrix(J + 1, 3)
 ITMFLEX.TextMatrix(J, 4) = ITMFLEX.TextMatrix(J + 1, 4)
 ITMFLEX.TextMatrix(J, 5) = ITMFLEX.TextMatrix(J + 1, 5)
 ITMFLEX.TextMatrix(J, 6) = ITMFLEX.TextMatrix(J + 1, 6)
Next J

If ITMFLEX.Rows > 2 Then
   ITMFLEX.Rows = ITMFLEX.Rows - 1
Else
   ITMFLEX.Rows = ITMFLEX.Rows - 1
   ITMFLEX.Rows = 2
End If

Call CLEARDATA

If MsgBox("Want to Add More Item ", vbYesNo + vbDefaultButton2) = vbYes Then
          If ITMFLEX.TextMatrix(ITMFLEX.Rows - 1, 1) <> "" Then
             ITMFLEX.Rows = ITMFLEX.Rows + 1
          End If
           If ITMFLEX.Rows > 6 Then ITMFLEX.TopRow = ITMFLEX.TopRow + 2
        M_INAM.SetFocus
    Else
        M_CRDS.Enabled = True: M_CRDS.SetFocus
    End If
    
    If ITMFLEX.Rows - 1 <= 9 Then
       M_SRCH = "0" & CStr(ITMFLEX.Rows - 1)
    Else
       M_SRCH = CStr(ITMFLEX.Rows - 1)
    End If

SWITCH = False
M_INAM.SetFocus
cmdOk.Caption = "&Add"
CMDITMDEL.Enabled = False

End Sub

Private Sub CMDOK_Click()
 Dim INDEX As Long
 
 If Not SWITCH Then
      ROWNO = ITMFLEX.Rows - 1
 End If
 
 RMRK = Replace(RMRK, "'", "", 1)
 
 If CheckData(ROWNO) Then Exit Sub
 
    ITMFLEX.TextMatrix(ROWNO, 0) = Trim(M_SRCH)
    ITMFLEX.TextMatrix(ROWNO, 1) = Trim(M_INAM)
    ITMFLEX.TextMatrix(ROWNO, 2) = Trim(TXTGRAD)
    ITMFLEX.TextMatrix(ROWNO, 3) = Trim(nstr(Val(M_QNTY), 12, 3))
    ITMFLEX.TextMatrix(ROWNO, 4) = Trim(nstr(Val(M_ARAT), 12, 4))
    ITMFLEX.TextMatrix(ROWNO, 5) = Trim(nstr(Val(M_RATE), 12, 4))
    ITMFLEX.TextMatrix(ROWNO, 6) = RMRK
           
    If MsgBox("Want to Add More Item ", vbYesNo + vbDefaultButton2) = vbYes Then
          If ITMFLEX.TextMatrix(ITMFLEX.Rows - 1, 1) <> "" Then
             ITMFLEX.Rows = ITMFLEX.Rows + 1
          End If
           If ITMFLEX.Rows > 6 Then ITMFLEX.TopRow = ITMFLEX.TopRow + 2
        M_INAM.SetFocus
    Else
        M_CRDS.Enabled = True: M_CRDS.SetFocus
    End If
    
    If ITMFLEX.Rows - 1 <= 9 Then
       M_SRCH = "0" & CStr(ITMFLEX.Rows - 1)
    Else
       M_SRCH = CStr(ITMFLEX.Rows - 1)
    End If
    
    'REMOVE BELOW COMMENT BLOCK WHEN ITEMS PROCESS ARE GOING TO MULTIPLE
    Call CLEARDATA
    cmdOk.Caption = "&Add"
    SWITCH = False

End Sub

Private Sub Form_Activate()
  If Trim(ORDBOK) = "" Then
     Unload Me
     Exit Sub
  End If
  Call ColorComponent(Me)
  lblUnit.ForeColor = &HFF0000
  LBLDIVISION.ForeColor = &HFF0000
  Me.Caption = "SALE ORDER BOOKING BY : " + ORDBOK
  
  If EXP_REQ Then
     cmdExport.Enabled = True
  Else
     cmdExport.Enabled = False
  End If
  
  LBLCRLIMIT.ForeColor = vbRed
  LBLCRLIMITVAL.ForeColor = vbRed
  LBLPNDDO.ForeColor = vbRed
  LBLPNDDOVAL.ForeColor = vbRed
  LBLLEDBAL.ForeColor = vbRed
  LBLLEDBALVAL.ForeColor = vbRed
  
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If UCase(ActiveControl.NAME) = "TXTDVCD" Then Exit Sub
  
  If ActiveControl = Empty And UCase(ActiveControl.NAME) <> "RMRK" And UCase(ActiveControl.NAME) <> "M_PORD" And UCase(ActiveControl.NAME) <> "TXTUNIT" And UCase(ActiveControl.NAME) <> "TXTDVCD" Then Exit Sub
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub Form_Load()
Call CenterChild(frm_Main, Me)
Me.KeyPreview = True

'-------SALESMAN MASTER
M_DESC = Empty:  Key = Empty:  NEW_VISIBLE = False: CANCEL_VISIBLE = True: key_PressNew = True
ORDBOK = Empty: ORDDBCD = Empty
ORDBOK = SearchList1("SELECT TOP 20 CODE,NAME FROM SALMANMST WHERE RECSTAT='A'", 0, ORDBOK, "SELECT SALESMAN FROM LIST")
If Key = Empty Then Exit Sub

ORDDBCD = Key
RATMASTERREQ = GetUnitMaster("EXTRA1")
        
M_ORDT.Value = Now
M_ORDT.MinDate = FSDT
M_ORDT.MaxDate = FEDT

End Sub

Private Sub ITMFLEX_Click()
   If ITMFLEX.Rows > 1 And ITMFLEX.TextMatrix(ITMFLEX.ROW, 1) <> Empty Then
   
   If Val(ITMFLEX.TextMatrix(ITMFLEX.ROW, 7)) > 0 Then
      'MsgBox "Further Entry Exist(Packing/Dispatch/Cancellation)", vbCritical
      'Exit Sub
   End If
   
    cmdOk.Caption = "Upd&ate"
    CMDITMDEL.Enabled = True
    ROWNO = ITMFLEX.ROW
    M_SRCH = ITMFLEX.TextMatrix(ROWNO, 0)
    M_INAM = ITMFLEX.TextMatrix(ROWNO, 1)
    TXTGRAD = ITMFLEX.TextMatrix(ROWNO, 2)
    M_QNTY = ITMFLEX.TextMatrix(ROWNO, 3)
    M_ARAT = ITMFLEX.TextMatrix(ROWNO, 4)
    M_RATE = ITMFLEX.TextMatrix(ROWNO, 5)
    RMRK = ITMFLEX.TextMatrix(ROWNO, 6)
    SWITCH = True
  End If
    
   If Val(ITMFLEX.ROW) > 0 Then
     M_SRCH = ITMFLEX.TextMatrix(ITMFLEX.ROW, 0)
     Call M_SRCH_LostFocus
     M_INAM.SetFocus
   End If
   
End Sub

Private Sub ITMFLEX_GotFocus()
ITMFLEX.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub ITMFLEX_LostFocus()
 ITMFLEX.BackColor = vbWhite
End Sub


Private Sub M_PNAM_Change()
   Call SetPartyHelp
End Sub

Private Sub RMRK_GotFocus()
 RMRK.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub RMRK_LostFocus()
RMRK.BackColor = vbWhite
End Sub

Private Sub M_ARAT_GotFocus()
  M_ARAT.BackColor = RGB(BRED, BGREEN, BBLUE)
  M_ARAT.SelStart = 0
  M_ARAT.SelLength = Len(M_ARAT)
End Sub

Private Sub M_ARAT_LostFocus()
M_ARAT.BackColor = vbWhite
End Sub

Private Sub M_BRNM_GotFocus()
M_BRNM.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub M_BRNM_KeyDown(KeyCode As Integer, Shift As Integer)
    Me.KeyPreview = False
    
    If KeyCode = vbKeyF2 Or (Trim(M_BRNM) = Empty And KeyCode = vbKeyReturn) Then
        NEW_VISIBLE = True
        M_DESC = Empty
        Key = Empty
        M_BRNM.Text = SearchList1("SELECT TOP 20 CODE, NAME FROM REFMST WHERE CATA='B'", 0, M_BRNM.Text, "SELECT AGENT FROM LIST")
        If key_PressNew = True Then
            M_DESC = ""
            Key = ""
            Ref_Cat = "B"
            M_BRNM.Text = ""
            Frm_Ref_FAS.Show
        Else
            M_BRCD = Key
        End If
    End If
    
    Me.KeyPreview = True
    Call M_RATE_Change
End Sub

Private Sub M_BRNM_LostFocus()
 M_BRNM.BackColor = vbWhite
End Sub

Private Sub M_CRDS_Change()
  LBLDUDT.Caption = "Delivery Date " + Format((M_ORDT.Value + Val(M_CRDS)), "DD/MM/YYYY")
End Sub

Private Sub M_CRDS_GotFocus()
M_CRDS.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub M_CRDS_KeyPress(KeyAscii As Integer)
  If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub M_CRDS_LostFocus()
 M_CRDS.BackColor = vbWhite
End Sub

Private Sub M_DNAM_KeyDown(KeyCode As Integer, Shift As Integer)
    Me.KeyPreview = False
    If KeyCode = vbKeyF2 Or (Trim(M_DNAM) = Empty And KeyCode = vbKeyReturn) Then
        NEW_VISIBLE = True
        M_DESC = Empty
        Key = Empty
        M_DNAM.Text = SearchList1("SELECT TOP 20 CODE, NAME FROM REFMST WHERE CATA='Y'", 0, M_DNAM.Text, "SELECT DELIVERY PARTY")
        If key_PressNew = True Then
            M_DESC = ""
            Key = ""
            Ref_Cat = "Y"
            M_DNAM.Text = ""
            Frm_Ref_FAS.Show
        Else
            M_DCOD = Key
        End If
    End If
    Me.KeyPreview = True
End Sub

Private Sub M_INAM_GotFocus()
M_INAM.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub M_INAM_KeyDown(KeyCode As Integer, Shift As Integer)
    Me.KeyPreview = False
    
    If KeyCode = vbKeyF2 Or (Trim(M_INAM) = Empty And KeyCode = vbKeyReturn) Then
        NEW_VISIBLE = True
        M_DESC = Empty
        Key = Empty
        M_INAM.Text = SearchList1("SELECT TOP 20 CODE, NAME FROM FINITMMST WHERE COMP='" & compPth & "' AND UNIT='" & FindUnit & "' AND DVCD='" & DIVCODE & "'", 0, M_INAM.Text, "SELECT FINISH ITEM FROM LIST")
        
        If key_PressNew = True Then
          M_DESC = ""
          DIVCOD = DIVCODE
          DIVNAM = DIVNAME
          frm_FinItmMst.ONLINEITEM = True
          M_INAM = Empty
          frm_FinItmMst.Show
        End If
    End If
    Me.KeyPreview = True
End Sub

Private Sub M_INAM_LostFocus()
 M_INAM.BackColor = vbWhite
End Sub

Private Sub M_ORDN_GotFocus()
  M_ORDN.BackColor = RGB(BRED, BGREEN, BBLUE)
  M_RMRK = Trim(M_BRNM.Text) + " " + M_ORDN
End Sub

Private Sub M_ORDN_LostFocus()
 M_ORDN.BackColor = vbWhite
End Sub

Private Sub M_ORDT_KeyDown(KeyCode As Integer, Shift As Integer)
  If M_COMPBILL = "SIL" Then
   Else
    M_RMRK = Trim(M_BRNM.Text) + " " + M_ORDN
  End If
    
  If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
  End If
End Sub
Public Sub btn_sts(bool As Boolean)
    cmdSave.Enabled = bool
    cmdCancel.Enabled = bool
    cmdAdd.Enabled = Not bool
    cmdEdit.Enabled = Not bool
    cmdDelete.Enabled = Not bool
End Sub

Private Sub ClsTxt()
  
  LBLCRLIMITVAL = "0.00"
  LBLLEDBALVAL = "0.00"
  LBLPNDDOVAL = "0.00"
  
  M_PNAM = Empty
  M_DNAM = Empty
  M_BRNM = Empty
  M_ORDN = Empty
  M_PORD = Empty
  ITMFLEX.Clear
  ITMFLEX.Rows = 2
  M_SRCH = Empty
  M_INAM = Empty
  M_QNTY = Empty
  M_RATE = Empty
  M_ARAT = Empty
  RMRK = Empty
  M_CRDS = Empty
  
  M_RMRK = Empty
  M_PCOD = Empty
  M_DCOD = Empty
  M_BRCD = Empty
  M_TRCD = Empty
  M_TXNM = Empty
  TXTGRAD = Empty
  TXTFREIGHT = Empty
    
  ITMFLEX.Clear
  ITMFLEX.ColWidth(0) = 400
  ITMFLEX.ColWidth(1) = 1700
  ITMFLEX.ColWidth(2) = 1900
  ITMFLEX.ColWidth(3) = 1250
  ITMFLEX.ColWidth(4) = 1000
  ITMFLEX.ColWidth(5) = 1000
  ITMFLEX.ColWidth(6) = 1600
  ITMFLEX.ColWidth(7) = 0    'Specially in editing of export
  ITMFLEX.Clear
  ITMFLEX.TextMatrix(0, 0) = "Sr."
  ITMFLEX.TextMatrix(0, 1) = "Item Name"
  ITMFLEX.TextMatrix(0, 2) = LBLCFG.Caption
  ITMFLEX.TextMatrix(0, 3) = "Quantity"
  ITMFLEX.TextMatrix(0, 4) = "Ass.Rate"
  ITMFLEX.TextMatrix(0, 5) = "Net Rate"
  ITMFLEX.TextMatrix(0, 6) = "Remarks"
  ITMFLEX.TextMatrix(0, 7) = "UsedQty" 'Specially in editing of export
  
  ITMFLEX.ColAlignment(0) = vbLeftJustify
  ITMFLEX.ColAlignment(1) = vbLeftJustify
  ITMFLEX.ColAlignment(2) = vbRightJustify
  ITMFLEX.ColAlignment(3) = vbRightJustify
  ITMFLEX.ColAlignment(4) = vbRightJustify
End Sub

Private Sub M_PNAM_GotFocus()
 M_PNAM.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub M_PNAM_KeyDown(KeyCode As Integer, Shift As Integer)
    Me.KeyPreview = False
    
    If KeyCode = vbKeyF2 Or (Trim(M_PNAM) = Empty And KeyCode = vbKeyReturn) Then
        NEW_VISIBLE = True
        M_DESC = Empty
        Key = Empty
        M_PNAM.Text = SearchList1("SELECT TOP 20 CODE, NAME FROM ACCMST WHERE DRCR='D'", 0, M_PNAM.Text, "SELECT A/C PARTY")
        If key_PressNew = True Then
            M_DESC = ""
            Key = ""
            M_PNAM.Text = ""
            frm_Acc.Show
        Else
            M_PNAM.Tag = Key
        End If
    End If
    
    Me.KeyPreview = True
End Sub

Private Sub GENORDN()
   Dim NO As Double, Prfx As String
     
   If RS.State = 1 Then RS.Close
   RS.Open "SELECT ISNULL(PRFX,'A') AS PRFX,SUBSTRING(LSRNO,2,5) AS SRNO FROM SALMANMST WHERE CODE='" & ORDDBCD & "'", CN, adOpenDynamic, adLockOptimistic
   If RS.EOF Then Exit Sub
   
     Prfx = Trim(RS!Prfx & "")
     If Prfx = Empty Then Prfx = "O"
     NO = Val(RS!SRNO)
     NO = NO + 1
   
   RS.Close
        
   If NO < 10 Then
     M_ORDN = "0000" + Trim(nstr(NO, 1, 0))
   ElseIf NO < 100 Then
     M_ORDN = "000" + Trim(nstr(NO, 1, 0))
   ElseIf NO < 1000 Then
     M_ORDN = "00" + Trim(nstr(NO, 1, 0))
   ElseIf NO < 10000 Then
     M_ORDN = "0" + Trim(nstr(NO, 1, 0))
   ElseIf NO < 100000 Then
     M_ORDN = Trim(nstr(NO, 1, 0))
   End If
      
   'M_ORDN = Prfx & M_ORDN & Mid$(CStr(FSDT), 9, 10) & Mid$(CStr(FEDT), 9, 10)
   M_ORDN = Prfx & M_ORDN & FYCD
End Sub

Private Sub M_PNAM_LostFocus()
 M_PNAM.BackColor = vbWhite
 
    If SAVEFLAG Then
     Dim GETRS As ADODB.Recordset
     Set GETRS = New ADODB.Recordset
  
     If GETRS.State = 1 Then GETRS.Close
     GETRS.Open "SELECT BRCD,RCOD,TXCD,TTYP FROM ACCMST WHERE NAME='" & M_PNAM & "' ", CN, adOpenDynamic, adLockOptimistic
     If Not GETRS.EOF Then
        M_BRNM = GetCode("REFMST", GETRS!BRCD & "", "CODE", "NAME")
        M_BRCD = Trim(GETRS!BRCD & "")
        M_TXNM = GetCode("TAXMST", GETRS!TXCD & "", "CODE", "NAME")
        M_TXCD = Trim(GETRS!TXCD & "")
        
     End If
  End If
End Sub

Private Sub M_PORD_GotFocus()
 M_PORD.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub M_PORD_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub M_PORD_LostFocus()
   M_PORD.BackColor = vbWhite
   M_INAM.Enabled = True
   If M_INAM.Enabled Then M_INAM.SetFocus
End Sub

Private Sub M_QNTY_GotFocus()
  M_QNTY.BackColor = RGB(BRED, BGREEN, BBLUE)
  M_QNTY.SelStart = 0
  M_QNTY.SelLength = Len(M_QNTY)
End Sub

Private Sub M_QNTY_KeyPress(KeyAscii As Integer)
 If CheckNumericKey(KeyAscii, M_QNTY, Me) = 0 Then KeyAscii = 0
End Sub

Private Sub M_QNTY_LostFocus()
 M_QNTY.BackColor = vbWhite
End Sub

Private Sub M_RATE_Change()
  If Val(M_RATE) <= 0 Or M_TXNM = Empty Or M_BRNM = Empty Then Exit Sub
  M_ARAT = GetOrderReverseRate(M_BRNM, M_TXNM, M_RATE, Val(TXTFREIGHT))
End Sub

Private Sub M_RATE_GotFocus()
  M_RATE.BackColor = RGB(BRED, BGREEN, BBLUE)
  M_RATE.SelStart = 0
  M_RATE.SelLength = Len(M_RATE)
End Sub

Private Sub M_RATE_KeyPress(KeyAscii As Integer)
 'If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 46 Then KeyAscii = 0
 If CheckNumericKey(KeyAscii, M_RATE, Me) = 0 Then KeyAscii = 0
End Sub

Private Sub M_RATE_LostFocus()
M_RATE.BackColor = vbWhite
End Sub

Private Sub M_RMRK_GotFocus()
  M_RMRK.BackColor = RGB(BRED, BGREEN, BBLUE)
  If SAVEFLAG Then
     M_RMRK = Trim(M_BRNM.Text) + " --- O.No. :" + M_ORDN
  End If
End Sub

Private Sub M_RMRK_LostFocus()
 M_RMRK.BackColor = vbWhite
End Sub

Private Sub M_SRCH_GotFocus()
  M_SRCH.BackColor = RGB(BRED, BGREEN, BBLUE)
  
  M_SRCH.SelStart = 0
  M_SRCH.SelLength = Len(M_SRCH)
  
  If Not DataEntered Then
    M_SRCH = "1"
    M_SRCH.Enabled = False
  End If
  
End Sub

Private Sub M_SRCH_LostFocus()
  M_SRCH.BackColor = vbWhite
  If Val(M_SRCH) > 1 And DataEntered Then
 
    M_CRDS.SetFocus
       Exit Sub
  End If
  
  If Val(M_SRCH) = 0 And DataEntered Then
    M_CRDS.SetFocus
    Exit Sub
  End If
  
  Dim NO As Double
   
  NO = Val(M_SRCH)
     
  If NO <= 9 Then
    M_SRCH = "0" + nstr(NO, 1, 0)
   Else
    M_SRCH = nstr(NO, 2, 0)
  End If
  
End Sub


Private Sub cmdSave_Click()
On Error GoTo LAST

If chkdata Then Exit Sub

If Not IsValidOrderAmt Then Exit Sub

If Val(M_CRDS) = 0 Then
  If (MsgBox("Is Zero Allowed for Scheduled Days ?", vbYesNo) = VBNO) Then
      M_CRDS.SetFocus
      Exit Sub
  End If
End If

M_PORD = Replace(M_PORD, "'", "", 1)
M_RMRK = Replace(M_RMRK, "'", "", 1)
    
If Not DataEntered Then
   MsgBox "No Data Found To Save Record !! Can Not Save Record !!", vbInformation, "Cancelled !!"
   Exit Sub
End If
    
M_PCOD = GetCode("ACCMST", M_PNAM, "NAME", "CODE")
If M_PCOD = Empty Then M_PNAM.SetFocus: Exit Sub
         
    If RS.State = 1 Then RS.Close
    RS.Open "SELECT * FROM ORDMAN WHERE COMP='" & compPth & "' AND UNIT='" & FindUnit & "' AND ORDN='" & M_ORDN & "' AND DBCD='" & ORDDBCD & "'", CN, adOpenDynamic, adLockOptimistic
    If Not RS.EOF And SAVEFLAG = True Then
        MsgBox "Duplicate Order No."
        Exit Sub
    End If
    M_DCOD = Empty
  
    If RS.State = 1 Then RS.Close
    RS.Open "SELECT * FROM REFMST WHERE NAME='" & M_BRNM & "' AND CATA='B'", CN, adOpenKeyset, adLockPessimistic
    If RS.EOF Then
        M_BRNM.SetFocus
        Exit Sub
    Else
        M_BRCD = RS!CODE
    End If
    
    If RS.State = 1 Then RS.Close
    RS.Open "SELECT CODE,RATE_CODE FROM TAXMST WHERE NAME='" & M_TXNM & "'", CN, adOpenKeyset, adLockPessimistic
    If RS.EOF Then
        M_TXNM.SetFocus
        Exit Sub
    Else
        M_TXCD = Trim(RS!CODE & "")
        RATECOD = Trim(RS!RATE_CODE & "")
    End If
    
    'ASSUMING USER WILL NOT SAVE ORDER IMPROPER OTHERWISE EDITING OF ORDER TAKEN FIXED QNTY TO EDITED
    'CHECK DATA FOR EXPORT DATA IN EDIT MODE
     
     If EXP_REQ And Not SAVEFLAG And IsFinalQty Then
        Dim FNLQTY As Double: FNLQTY = 0
        Dim IND As Long
        For IND = 1 To ITMFLEX.Rows - 1
            FNLQTY = FNLQTY + Val(ITMFLEX.TextMatrix(IND, 3))
        Next
        
        If FNLQTY <> EDIT_ORDQTY Then MsgBox "Order Quantity must be same with edit Qty.", vbCritical: Exit Sub
     End If
      
  Dim M_ICOD As String
  Dim M_GRAD As String
  Dim i As Long
  i = 1
  
  CN.BeginTrans
  
  For i = 1 To ITMFLEX.Rows - 1                  'Export Case
    If Not Val(ITMFLEX.TextMatrix(i, 0)) = 0 And Val(ITMFLEX.TextMatrix(i, 7)) = 0 Then
                       
        If RS.State = 1 Then RS.Close
        RS.Open "SELECT * FROM FINITMMST WHERE COMP='" & compPth & "' AND UNIT='" & FindUnit & "' AND DVCD='" & DIVCODE & "' AND NAME='" & ITMFLEX.TextMatrix(i, 1) & "'", CN, adOpenKeyset, adLockPessimistic
        If Not RS.EOF Then
            M_ICOD = RS!CODE
        Else
            M_ICOD = Empty
        End If
        
        If RS.State = 1 Then RS.Close
        RS.Open "SELECT * FROM GRDMST WHERE GRAD='" & ITMFLEX.TextMatrix(i, 2) & "'", CN, adOpenKeyset, adLockPessimistic
        If Not RS.EOF Then
            M_GRAD = Trim(RS!CODE & "")
        Else
            M_GRAD = Empty
        End If
        
            If SAVEFLAG = True Then
              Call GENORDN
            End If
            
QUERY = "INSERT INTO ORDMAN(COMP,UNIT,ORDN,OSRC,PCOD,DCOD,BRCD,ICOD,QNTY,RATE,ARAT,AMNT,PORD,"
QUERY = QUERY & "CRDS,RMRK,TRCD,ordt,TXCD,DBCD,FREIGHT_PERKG,FREIGHT_FACTOR,RTCD,ORDRMRK) VALUES ('" & compPth & "','" & FindUnit & _
"','" & M_ORDN & "','" & ITMFLEX.TextMatrix(i, 0) & "','" & M_PCOD & "','" & DIVCODE & _
"','" & M_BRCD & "','" & M_ICOD & "','" & Val(ITMFLEX.TextMatrix(i, 3)) & _
"','" & Val(ITMFLEX.TextMatrix(i, 5)) & "','" & Val(ITMFLEX.TextMatrix(i, 4)) & _
"','" & Val(ITMFLEX.TextMatrix(i, 3)) * Val(ITMFLEX.TextMatrix(i, 5)) & "','" & M_PORD & "','" & Val(M_CRDS) & _
"','" & ITMFLEX.TextMatrix(i, 6) & "','" & M_GRAD & "','" & Format(M_ORDT, "mm/dd/yyyy") & _
"','" & M_TXCD & "','" & ORDDBCD & "'," & Val(TXTFREIGHT) & "," & FREIGHT & ",'" & RATECOD & "','" & M_RMRK & "')"
            
CN.Execute "DELETE FROM ORDMAN WHERE COMP='" & compPth & "' AND DBCD='" & ORDDBCD & _
           "' AND ORDN='" & M_ORDN & "' AND OSRC='" & ITMFLEX.TextMatrix(i, 0) & "'"
CN.Execute QUERY

ElseIf Not Val(ITMFLEX.TextMatrix(i, 0)) = 0 And Val(ITMFLEX.TextMatrix(i, 7)) > 0 Then 'FOR EDITING AFTER DISPATCH

   CN.Execute "UPDATE ORDMAN SET ORDT='" & Format(M_ORDT, "mm/dd/yyyy") & "',QNTY ='" & Val(ITMFLEX.TextMatrix(i, 3)) & _
              "',RATE ='" & Val(ITMFLEX.TextMatrix(i, 5)) & "',ARAT = '" & Val(ITMFLEX.TextMatrix(i, 4)) & _
              "',AMNT = '" & Val(ITMFLEX.TextMatrix(i, 3)) * Val(ITMFLEX.TextMatrix(i, 5)) & _
              "',ORDRMRK ='" & M_RMRK & "' WHERE COMP='" & compPth & "' AND DBCD='" & ORDDBCD & "' AND ORDN='" & M_ORDN & _
              "' AND OSRC='" & ITMFLEX.TextMatrix(i, 0) & "'"

End If
Next
  
If SAVEFLAG Then
 ' CN.Execute "INSERT INTO DAILYSTAT (COMP,VTYP,SRNO,PCOD,DBCD,VBNO,QNTY,AMNT,CUSR,DTTM,ACTN) VALUES('" & compPth & "','SOB','XXXXXXXXXXXXX','" & M_PNAM & "',NULL,'" & M_ORDN & "'," & Val(ITMFLEX.TextMatrix(I, 2)) & "," & Val(ITMFLEX.TextMatrix(I, 4)) & ",'" & cUName & "','" & Format(Now, "MM/dd/yyyy HH:MM:SS AMPM") & "','N')"
Else
  'CN.Execute "INSERT INTO DAILYSTAT (COMP,VTYP,SRNO,PCOD,DBCD,VBNO,QNTY,AMNT,CUSR,DTTM,ACTN) VALUES('" & compPth & "','SOB','XXXXXXXXXXXXX','" & M_PNAM & "',NULL,'" & M_ORDN & "'," & Val(ITMFLEX.TextMatrix(I, 2)) & "," & Val(ITMFLEX.TextMatrix(I, 4)) & ",'" & cUName & "','" & Format(Now, "MM/dd/yyyy HH:MM:SS AMPM") & "','M')"
End If

Call EXPORT_DETAIL
CN.CommitTrans

MsgBox "Your Order No. is " + M_ORDN.Text

            
If SAVEFLAG = True Then
   If RS.State = 1 Then RS.Close
   RS.Open "SELECT * FROM SALMANMST WHERE CODE='" & ORDDBCD & "'", CN, adOpenDynamic, adLockOptimistic
   If Not RS.EOF Then
      RS!LSRNO = Trim(M_ORDN)
      RS.Update
   End If
End If
    
  Call ClsTxt
  
  btn_sts (False)
  SAVEFLAG = True
  cmdAdd.SetFocus
  
  Exit Sub
  
LAST:
  ErrNumber = ERR.Number
  ErrMessage = ERR.Description
  frm_ErrorHandler.Show vbModal
  On Error GoTo 0
  CN.RollbackTrans
0:
End Sub

Private Sub M_TXNM_GotFocus()
M_TXNM.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub M_TXNM_KeyDown(KeyCode As Integer, Shift As Integer)
    Me.KeyPreview = False
    
    If KeyCode = vbKeyF2 Or (Trim(M_TXNM) = Empty And KeyCode = vbKeyReturn) Then
        NEW_VISIBLE = True
        M_DESC = Empty
        Key = Empty
        M_TXNM.Text = SearchList1("SELECT TOP 20 CODE, NAME FROM TAXMST WHERE RECSTAT='A'", 0, M_TXNM.Text, "SELECT TAX FROM LIST")
        If key_PressNew = True Then
            M_DESC = ""
            Key = ""
            Ref_Cat = "T"
            M_TXNM.Text = ""
            FrmSaleTaxMaster.Show
        Else
            M_TXCD = Key
        End If
    End If
    Me.KeyPreview = True
End Sub

Private Sub M_TXNM_LostFocus()
  M_TXNM.BackColor = vbWhite
End Sub

Private Sub TXTDVCD_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn And txtDVCD <> Empty Then
    txtDVCD.Enabled = False
    txtUNIT.Enabled = False
    Call ClsTxt
    btn_sts (False)
    cmdAdd.SetFocus
End If
End Sub

Private Sub txtFREIGHT_GotFocus()
TXTFREIGHT.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTFREIGHT_KeyPress(KeyAscii As Integer)
 If CheckNumericKey(KeyAscii, TXTFREIGHT, Me) = 0 Then KeyAscii = 0
End Sub

Private Sub txtFREIGHT_LostFocus()
TXTFREIGHT.BackColor = vbWhite
End Sub

Private Sub TXTGRAD_GotFocus()
 TXTGRAD.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTGRAD_KeyDown(KeyCode As Integer, Shift As Integer)
  If (TXTGRAD = Empty And KeyCode = vbKeyReturn) Or KeyCode = vbKeyF2 Then
    NEW_VISIBLE = True
    TXTGRAD = SearchList1("SELECT DISTINCT GRAD AS GRD,GRAD FROM GRDMST", 0, TXTGRAD, "SELECT " & LBLCFG.Caption)
      If key_PressNew = True Then
          M_DESC = ""
          TXTGRAD = Empty
          FRM_GRDMST.Show
      End If
  End If
End Sub

Private Sub TXTGRAD_LostFocus()
 TXTGRAD.BackColor = vbWhite
End Sub

Private Function CheckData(RNO As Long) As Boolean
Dim INDEX As Long
    If Trim(M_INAM) = Empty Then
        MsgBox "Please Select Items From List !!", vbInformation
        M_INAM.SetFocus
        CheckData = True
        Exit Function
    End If
    
    If Trim(TXTGRAD) = Empty Then
        MsgBox "Please Select From List !!", vbInformation
        TXTGRAD.SetFocus
        CheckData = True
        Exit Function
    End If
    
    If Val(M_QNTY) < 1 Then
        If MsgBox("Want to Made Zero Qty Order", vbYesNo + vbDefaultButton2) = VBNO Then
           M_QNTY.Enabled = True: M_QNTY.SetFocus
           CheckData = True
           Exit Function
        End If
    End If
    
    If Val(M_RATE) = 0 Then
        MsgBox "Please Enter Valid Rate Value !!", vbInformation, "Rate Is Missing"
        M_RATE.SetFocus
        CheckData = True
        Exit Function
    End If
        
    If Val(M_ARAT) = 0 Then
        MsgBox "Please Enter Valid Ass. Rate Value !!", vbInformation, "Rate Is Missing"
        M_ARAT.SetFocus
        CheckData = True
        Exit Function
    End If

    For INDEX = 1 To ITMFLEX.Rows - 1
        If SWITCH And INDEX = RNO Then
           If Val(M_QNTY) < Val(ITMFLEX.TextMatrix(INDEX, 7)) Then
              MsgBox "Quantity Can be Greater or Equal to " & ITMFLEX.TextMatrix(INDEX, 7)
              M_QNTY.SetFocus
              CheckData = True
              Exit Function
           End If
        End If
        
        If (Trim(ITMFLEX.TextMatrix(INDEX, 1)) & Trim(ITMFLEX.TextMatrix(INDEX, 2)) = M_INAM & TXTGRAD) And (Not SWITCH Or (SWITCH And INDEX <> RNO)) Then
           MsgBox "Invalid Item Detail"
           M_INAM.SetFocus
           CheckData = True
           Exit Function
        End If
    Next INDEX
    
End Function

Private Sub CLEARDATA()
        M_INAM = Empty
        TXTGRAD = Empty
        M_QNTY = Empty
        M_ARAT = Empty
        M_RATE = Empty
        RMRK = Empty
End Sub

Private Function DataEntered() As Boolean
Dim i As Integer
    DataEntered = False
    For i = 1 To ITMFLEX.Rows - 1
        If Trim(ITMFLEX.TextMatrix(i, 1)) <> "" Then
            DataEntered = True
            Exit Function
        End If
    Next
End Function

Private Sub txtUnit_KeyDown(KeyCode As Integer, Shift As Integer)
    Me.KeyPreview = False
    
    If KeyCode = vbKeyF2 Or (Trim(txtUNIT) = Empty And KeyCode = vbKeyReturn) Then
        NEW_VISIBLE = False
        M_DESC = Empty
        Key = Empty
        txtUNIT.Text = SearchList1("SELECT TOP 20 CODE, NAME FROM UNTMST WHERE COMP='" & compPth & "'", 0, txtUNIT.Text, "SELECT UNIT FROM LIST")
        txtUNIT.Tag = Key
    End If
    
      Me.KeyPreview = True
End Sub

Private Sub txtDVCD_KeyDown(KeyCode As Integer, Shift As Integer)
    Me.KeyPreview = False
    
    If txtUNIT = Empty Then txtUNIT.Enabled = True: txtUNIT.SetFocus: Exit Sub
    
    If KeyCode = vbKeyF2 Or (Trim(txtDVCD) = Empty And KeyCode = vbKeyReturn) Then
        NEW_VISIBLE = False
        M_DESC = Empty
        Key = Empty
        txtDVCD.Text = SearchList1("SELECT TOP 20 CODE, NAME FROM DIVMST WHERE COMP='" & compPth & "' AND UNIT='" & FindUnit & "' AND CODE<>'000001' AND RECSTAT='A'", 0, txtDVCD.Text, "SELECT DIVISION FROM LIST")
        txtDVCD.Tag = Key
        
        DIVNAME = txtDVCD
        DIVCODE = Key
        
        LBLCFG.Caption = LabelDisplay(txtDVCD.Tag, FindUnit)
    End If
    
    Me.KeyPreview = True
End Sub

Private Function chkdata() As Boolean
If txtUNIT = Empty Then
  MsgBox "Enter Unit then Save"
  chkdata = True
  txtUNIT.Enabled = True
  txtUNIT.SetFocus
  Exit Function
End If

If txtDVCD = Empty Then
  MsgBox "Enter Division then Save"
  chkdata = True
  txtDVCD.Enabled = True
  txtDVCD.SetFocus
  Exit Function
End If

Dim LUNIT As String, i As Long
LUNIT = FindUnit

For i = 1 To ITMFLEX.Rows - 1
    If Not Val(ITMFLEX.TextMatrix(i, 0)) = 0 Then
                       
       If RS.State = 1 Then RS.Close
       RS.Open "SELECT * FROM FINITMMST WHERE COMP='" & compPth & "' AND UNIT='" & LUNIT & _
               "' AND DVCD='" & DIVCODE & "' AND NAME='" & ITMFLEX.TextMatrix(i, 1) & "'", CN, adOpenKeyset, adLockPessimistic
       If RS.EOF Then
          MsgBox "Invalid Item In Selected List, Check in Master", vbCritical
          chkdata = True
          ITMFLEX.ROW = i
          ITMFLEX.COL = 1
          ITMFLEX.SetFocus
          Exit Function
       End If
        
       If RS.State = 1 Then RS.Close
       RS.Open "SELECT * FROM GRDMST WHERE GRAD='" & ITMFLEX.TextMatrix(i, 2) & "'", CN, adOpenKeyset, adLockPessimistic
       If RS.EOF Then
          MsgBox "Invalid Grade In Selected List, Check in Master", vbCritical
          chkdata = True
          ITMFLEX.ROW = i
          ITMFLEX.COL = 2
          ITMFLEX.SetFocus
          Exit Function
       End If
       
     End If
     
 Next i

End Function

Public Function FindUnit() As String
Dim UNTRS As ADODB.Recordset
Set UNTRS = New ADODB.Recordset
If UNTRS.State = 1 Then UNTRS.Close
UNTRS.Open "SELECT * FROM UNTMST WHERE COMP='" & compPth & "' AND NAME ='" & txtUNIT & "'", CN, adOpenDynamic, adLockOptimistic
If Not UNTRS.EOF Then
   FindUnit = UNTRS!CODE
Else
   FindUnit = Empty
End If
UNTRS.Close
End Function

Private Sub EXPORT_DETAIL()
Dim EXPRS As ADODB.Recordset
Set EXPRS = New ADODB.Recordset

If EXP_REQ Then

    Dim PKGCOD As String
    Dim BNKCOD As String
    
    If EXPRS.State = 1 Then EXPRS.Close
    EXPRS.Open "SELECT CODE,NAME FROM PKGNGMST WHERE NAME='" & FRM_TRNEXPORD.TXTPKGTYP & "' AND STATUS='A' AND RECSTAT='A'", CN, adOpenDynamic, adLockOptimistic
    If Not EXPRS.EOF Then
      PKGCOD = EXPRS!CODE
     Else
      PKGCOD = Empty
    End If
    
    If EXPRS.State = 1 Then EXPRS.Close
    EXPRS.Open "SELECT CODE,NAME FROM REFMST WHERE NAME='" & FRM_TRNEXPORD.TXTBANKDTL & "' AND CATA='L'", CN, adOpenDynamic, adLockOptimistic
    If Not EXPRS.EOF Then
      BNKCOD = EXPRS!CODE
     Else
      BNKCOD = Empty
    End If
    
    If EXPRS.State = 1 Then EXPRS.Close
    EXPRS.Open "SELECT * FROM EXPORD WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DBCD='" & ORDDBCD & "' AND ORDN='" & M_ORDN & "'", CN, adOpenDynamic, adLockOptimistic
    If EXPRS.EOF Then
      EXPRS.AddNew
    End If
    
    EXPRS!COMP = compPth
    EXPRS!unit = UNCD
    EXPRS!dbcd = ORDDBCD
    EXPRS!ORDN = M_ORDN
    EXPRS!EXPORTREFNO = Trim(Mid(Trim(FRM_TRNEXPORD.TXTEXPORTREF), 1, 50))
    EXPRS!CNTRYOFORGIN = Trim(Mid(Trim(FRM_TRNEXPORD.TXTCNTRYOFORIGIN), 1, 50))
    EXPRS!CNTRYOFFINALDES = Trim(Mid(Trim(FRM_TRNEXPORD.TXTCNTRYFNLDES), 1, 50))
    EXPRS!TRMSOFDLRY = Trim(Mid(Trim(FRM_TRNEXPORD.TXTTERMS), 1, 50))
    EXPRS!TRMSOFPYMT = Trim(Mid(Trim(FRM_TRNEXPORD.TXTPAYMENT), 1, 50))
    EXPRS!PRECARIGBY = Trim(Mid(Trim(FRM_TRNEXPORD.TXTPRECARIAGE), 1, 50))
    EXPRS!PLACEOFRCPT = Trim(Mid(Trim(FRM_TRNEXPORD.TXTPLACEOFRCPT), 1, 50))
    EXPRS!VSLFLTNO = Trim(Mid(Trim(FRM_TRNEXPORD.TXTVSLNO), 1, 50))
    EXPRS!PORTOFLOAD = Trim(Mid(Trim(FRM_TRNEXPORD.TXTPORTOFLOD), 1, 50))
    EXPRS!PORTOFDISCHARG = Trim(Mid(Trim(FRM_TRNEXPORD.TXTPORTOFDIS), 1, 50))
    EXPRS!FINALDEST = Trim(Mid(Trim(FRM_TRNEXPORD.TXTFNLDES), 1, 50))
    EXPRS!REMARK1 = Trim(Mid(Trim(FRM_TRNEXPORD.TXTREMARK1), 1, 50))
    EXPRS!REMARK2 = Trim(Mid(Trim(FRM_TRNEXPORD.TXTREMARK2), 1, 50))
    EXPRS!REMARK3 = Trim(Mid(Trim(FRM_TRNEXPORD.TXTREMARK3), 1, 50))
    EXPRS!MARKNO = Trim(Mid(Trim(FRM_TRNEXPORD.TXTMARKS), 1, 50))
    EXPRS!PKGTYPE = Trim(PKGCOD)
    EXPRS!PKGDESC = Trim(Mid(Trim(FRM_TRNEXPORD.TXTPKGTYP), 1, 50))
    EXPRS!BANKCODE = Trim(BNKCOD)
    EXPRS!CIFFOB = Trim(Mid(FRM_TRNEXPORD.TXTCIFFOB, 1, 50))
    EXPRS!PAYMENTBYIRRLC = Trim(Mid(Trim(FRM_TRNEXPORD.TXTPAYMENTIRR), 1, 50))
    EXPRS!EXRAT = Val(FRM_TRNEXPORD.txtEXRate)
    EXPRS.Update
End If
  
End Sub

Private Function EXP_REQ() As Boolean
EXP_REQ = False
Dim ISEXPORT As String
Dim EXPRS As ADODB.Recordset
Set EXPRS = New ADODB.Recordset

If EXPRS.State = 1 Then EXPRS.Close
EXPRS.Open "SELECT * FROM SALMANMST WHERE CODE='" & ORDDBCD & "'", CN, adOpenDynamic, adLockOptimistic
If Not EXPRS.EOF Then
   ISEXPORT = Trim(EXPRS!ISEXPORTORDER & "")
End If

If ISEXPORT = "1" Then
  EXP_REQ = True
Else
  EXP_REQ = False
End If
End Function

Private Function IsFinalQty() As Boolean
IsFinalQty = False

Dim FNLRS As ADODB.Recordset
Set FNLRS = New ADODB.Recordset

If FNLRS.State = 1 Then FNLRS.Close
FNLRS.Open "SELECT * FROM ORDMAN WHERE COMP='" & compPth & "' AND UNIT='" & FindUnit & "' AND ORDN='" & M_ORDN & _
           "' AND DBCD='" & ORDDBCD & "' AND (ORDMAN.DOQTY <> 0 or ORDMAN.DISPATCHQTY <> 0 or ORDMAN.CANCELQTY <> 0)", CN, adOpenDynamic, adLockOptimistic
                                 
If Not FNLRS.EOF Then
   IsFinalQty = True
End If
FNLRS.Close
End Function

Private Sub SetPartyHelp()
If M_PNAM = Empty Then
   Exit Sub
End If

Dim HLPRS As ADODB.Recordset
Set HLPRS = New ADODB.Recordset

If HLPRS.State = 1 Then HLPRS.Close
HLPRS.Open "SELECT LIMB FROM ACCMST WHERE NAME='" & M_PNAM & "'", CN, adOpenDynamic, adLockOptimistic
If Not HLPRS.EOF Then
    LBLCRLIMITVAL = nstr(Val(HLPRS!LIMB), 12, 2)
End If
HLPRS.Close

If HLPRS.State = 1 Then HLPRS.Close
HLPRS.Open "SELECT BALN FROM ONLINEBAL WHERE COMP='" & compPth & "' AND UNIT='" & FindUnit & _
           "' AND NAME ='" & M_PNAM & "'", CN, adOpenDynamic, adLockOptimistic
If Not HLPRS.EOF Then
    LBLLEDBALVAL = nstr(Val(HLPRS!BALN), 12, 2)
End If
HLPRS.Close

If HLPRS.State = 1 Then HLPRS.Close
HLPRS.Open "SELECT ISNULL(SUM((QNTY - DOQTY - DISPATCHQTY - CANCELQTY) * RATE),0) AS PNDORDVAL FROM ORDMAN " & _
           "WHERE COMP='" & compPth & "' AND UNIT='" & FindUnit & _
           "' AND PCOD ='" & GetCode("ACCMST", M_PNAM, "NAME", "CODE") & _
           "' AND (QNTY - DOQTY - DISPATCHQTY - CANCELQTY) > 0 AND ORDN<>'" & M_ORDN & "'", CN, adOpenDynamic, adLockOptimistic
If Not HLPRS.EOF Then
    LBLPNDDOVAL = nstr(Val(HLPRS!PNDORDVAL), 12, 2)
End If
HLPRS.Close

End Sub

Private Function IsValidOrderAmt() As Boolean
IsValidOrderAmt = True

Call SetPartyHelp

If Val(LBLCRLIMITVAL) = 0 Then Exit Function

Dim CURRENTORDERAMT As Double: CURRENTORDERAMT = 0
Dim CURORDAMT As String
Dim BALCRLIMIT As String
Dim i As Long

For i = 1 To ITMFLEX.Rows - 1
    If Not Val(ITMFLEX.TextMatrix(i, 0)) = 0 Then
       CURRENTORDERAMT = CURRENTORDERAMT + Val(ITMFLEX.TextMatrix(i, 3)) * Val(ITMFLEX.TextMatrix(i, 5))
    End If
Next

CURORDAMT = Trim(nstr(CURRENTORDERAMT, 12, 2))

If (Val(LBLCRLIMITVAL) - Val(LBLLEDBALVAL) - Val(LBLPNDDOVAL)) < CURRENTORDERAMT Then
   BALCRLIMIT = Trim(nstr(Val(LBLCRLIMITVAL) - Val(LBLLEDBALVAL) - Val(LBLPNDDOVAL), 12, 2))
   MsgBox "A/c Balance Credit Limit is " & BALCRLIMIT & " and Your Current Order Amount is " & CURORDAMT, vbCritical, "Credit Limit Exceed"
   IsValidOrderAmt = False
   Exit Function
End If

End Function


Public Function GetOrderReverseRate(BRNM As String, TAXNAM As String, RATE As Double, FREIGHT As Double) As Double
'INITIALISE BASIC RATE IS REVERSE RATE
 GetOrderReverseRate = RATE
'-----------------------------------

  Dim RATECOD As String
  Dim REVRS As ADODB.Recordset
  Set REVRS = New ADODB.Recordset
          
  'FIND RATE CODE FROM TAXCODE
  If REVRS.State = 1 Then REVRS.Close
  REVRS.Open "SELECT RATE_CODE,REVERSERATEREQ FROM TAXMST WHERE NAME='" & TAXNAM & "'", CN, adOpenDynamic, adLockOptimistic
  If Not REVRS.EOF Then
     RATECOD = REVRS!RATE_CODE
  End If
  REVRS.Close
  '----------------------------------------
  
  Dim VAT As Double, CST As Double, EXCISE As Double, TCESS As Double
  Dim PERVAT As Double, PERCST As Double, PEREXCISE As Double
  Dim BROKERAGE_REQ As Boolean
  Dim M_BRKPERC As Double, M_CD As Double, ASSRATE As Double
  Dim MID_RATE As Double, BSC_RATE As Double, BRK_AMNT As Double, RATE_FACTOR As Double
  BROKERAGE_REQ = False:  M_BRKPERC = 0:  M_CD = 0: VAT = 0: CST = 0: EXCISE = 0
  
  'BROKERAGE PERCENTAGE--------------------
  Set REVRS = New ADODB.Recordset
  If REVRS.State = 1 Then REVRS.Close
  REVRS.Open "SELECT * FROM REFMST WHERE NAME='" & BRNM & "'", CN, adOpenDynamic, adLockOptimistic
  If Not REVRS.EOF Then M_BRKPERC = REVRS!PERC
  REVRS.Close
  '----------------------------------------
  
  'FIND REQUIRED INFO FROM RATE MASTER
  If REVRS.State = 1 Then REVRS.Close
  REVRS.Open "SELECT * FROM RATEMST WHERE CODE='" & RATECOD & "' AND RECSTAT<>'D'", CN, adOpenDynamic, adLockOptimistic
  If Not REVRS.EOF Then
  
    BROKERAGE_REQ = Trim(REVRS!BROKERAGE & "")
    M_CD = REVRS!CD
    RATE_FACTOR = REVRS!RATE_FACTOR
        
    If Val(REVRS!PERTCESS & "") > 0 Then
       TCESS = (1 + Round(Val(REVRS!PERTCESS & "") / 100, 5))
    End If
    
    If Val(REVRS!PERVAT & "") > 0 Then
       VAT = (1 + Round(Val(REVRS!PERVAT & "") / 100, 5))
    End If
    If Val(REVRS!PERCST & "") > 0 Then
       CST = (1 + Round(Val(REVRS!PERCST & "") / 100, 5))
    End If
    If Val(REVRS!PEREXCISE & "") > 0 Then
       EXCISE = (1 + Round(Val(REVRS!PEREXCISE & "") / 100, 5))
    End If
    
  End If
  REVRS.Close
  '-----------------------------------------
  
  MID_RATE = RATE
  
  If BROKERAGE_REQ Then
    MID_RATE = RATE - M_CD             'NET RATE REVISE CD LESS
    Dim FAC1 As Double:   FAC1 = 0
    FAC1 = (1 + Round(M_BRKPERC / 100, 5))
    If FAC1 > 0 Then
       MID_RATE = 1 + (MID_RATE / FAC1)
    Else
       MsgBox "Invalid Brokerage in Broker Master", vbCritical
       Exit Function
    End If
  End If
     
     If TCESS > 0 Then
        MID_RATE = MID_RATE / TCESS
     End If
     
     If Val(VAT + CST) > 0 Then
        FREIGHT = Val(FREIGHT) / (VAT + CST)
        MID_RATE = MID_RATE / (VAT + CST)
     Else
        FREIGHT = Val(FREIGHT)
        MID_RATE = MID_RATE
     End If
     
     
     MID_RATE = MID_RATE - FREIGHT
     
     If Val(EXCISE) > 0 Then
       ASSRATE = MID_RATE / EXCISE
     Else
       ASSRATE = MID_RATE
     End If
     
     GetOrderReverseRate = nstr(ASSRATE, 12, 4)
     
End Function

