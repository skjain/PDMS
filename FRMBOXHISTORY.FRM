VERSION 5.00
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WELCHBUTTON.OCX"
Begin VB.Form FRMBOXHISTORY 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Box History"
   ClientHeight    =   7515
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9105
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7515
   ScaleWidth      =   9105
   Begin VB.Frame txtbill 
      Caption         =   "Invoice Detail"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   62
      Top             =   6720
      Width           =   8895
      Begin VB.TextBox txtINVNO 
         Height          =   285
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   64
         Top             =   240
         Width           =   1695
      End
      Begin WelchButton.lvButtons_H cmdExit 
         Height          =   375
         Left            =   7560
         TabIndex        =   65
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
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
         Image           =   "FRMBOXHISTORY.frx":0000
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdCLR 
         Height          =   375
         Left            =   6360
         TabIndex        =   66
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
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
         Image           =   "FRMBOXHISTORY.frx":059A
         cBack           =   -2147483633
      End
      Begin VB.Label Label30 
         Caption         =   "Invoice No.   "
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
         Left            =   1920
         TabIndex        =   63
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Dispatch Detail :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   120
      TabIndex        =   4
      Top             =   4200
      Width           =   8895
      Begin VB.TextBox DISPRETTYP 
         Height          =   285
         Left            =   6000
         Locked          =   -1  'True
         TabIndex        =   61
         Top             =   1560
         Width           =   2535
      End
      Begin VB.TextBox TXTADDRESS 
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   59
         Top             =   1920
         Width           =   6375
      End
      Begin VB.TextBox TXTDISPTYP 
         Height          =   285
         Left            =   6000
         Locked          =   -1  'True
         TabIndex        =   57
         Top             =   840
         Width           =   2535
      End
      Begin VB.TextBox TXTCONSINEE 
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   55
         Top             =   1560
         Width           =   3135
      End
      Begin VB.TextBox TXTAGENT 
         Height          =   285
         Left            =   6000
         Locked          =   -1  'True
         TabIndex        =   53
         Top             =   1200
         Width           =   2535
      End
      Begin VB.TextBox TXTPARTY 
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   52
         Top             =   1200
         Width           =   3135
      End
      Begin VB.TextBox TXTDONO 
         Height          =   285
         Left            =   6960
         Locked          =   -1  'True
         TabIndex        =   50
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox TXTORDN 
         Height          =   285
         Left            =   4440
         Locked          =   -1  'True
         TabIndex        =   48
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox TXTCHLN 
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   45
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox TXTCHLNDT 
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   44
         Top             =   480
         Width           =   1695
      End
      Begin VB.Shape Shape3 
         Height          =   1935
         Left            =   120
         Top             =   360
         Width           =   8655
      End
      Begin VB.Label Label29 
         Caption         =   "Returnable        "
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
         Left            =   4800
         TabIndex        =   60
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label28 
         Caption         =   "Consignee  Address"
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
         Left            =   360
         TabIndex        =   58
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label Label27 
         Caption         =   "Type of Dispatch        "
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
         Left            =   4320
         TabIndex        =   56
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label26 
         Caption         =   "Consignee      "
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
         Left            =   360
         TabIndex        =   54
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label25 
         Caption         =   "Agent Name          "
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
         Left            =   4800
         TabIndex        =   51
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label24 
         Caption         =   "D.O.No.           "
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
         Left            =   6120
         TabIndex        =   49
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label23 
         Caption         =   "Order No.           "
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
         Left            =   3360
         TabIndex        =   47
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label22 
         Caption         =   "Party Name          "
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
         Left            =   360
         TabIndex        =   46
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label21 
         Caption         =   "Challan No.            "
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
         Left            =   360
         TabIndex        =   43
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label20 
         Caption         =   "Challan Date             "
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
         Left            =   360
         TabIndex        =   42
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Packing Detail :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   8895
      Begin VB.TextBox TXTRMRK 
         Height          =   285
         Left            =   6480
         Locked          =   -1  'True
         TabIndex        =   41
         Top             =   2760
         Width           =   2175
      End
      Begin VB.TextBox TXTPLY 
         Height          =   285
         Left            =   6480
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   2400
         Width           =   2175
      End
      Begin VB.TextBox TXTRET 
         Height          =   285
         Left            =   6480
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   2040
         Width           =   2175
      End
      Begin VB.TextBox TXTNTWT 
         Height          =   285
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   2760
         Width           =   1215
      End
      Begin VB.TextBox TXTTRWT 
         Height          =   285
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox TXTGRWT 
         Height          =   285
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox TXTCPWT 
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   2760
         Width           =   1215
      End
      Begin VB.TextBox TXTCTWT 
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox TXTCOP 
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox TXTLTNO 
         Height          =   285
         Left            =   6360
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   360
         Width           =   2295
      End
      Begin VB.TextBox TXTGRAD 
         Height          =   285
         Left            =   6360
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox TXTSUBGRD 
         Height          =   285
         Left            =   6360
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   1080
         Width           =   2295
      End
      Begin VB.TextBox TXTFINITM 
         Height          =   285
         Left            =   6360
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   1440
         Width           =   2295
      End
      Begin VB.TextBox TXTDIVISION 
         Height          =   285
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   360
         Width           =   2535
      End
      Begin VB.TextBox TXTPACKST 
         Height          =   285
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   720
         Width           =   2535
      End
      Begin VB.TextBox TXTPACKTYP 
         Height          =   285
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   1080
         Width           =   2535
      End
      Begin VB.TextBox TXTPKGTYP 
         Height          =   285
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1440
         Width           =   2535
      End
      Begin VB.Shape Shape2 
         Height          =   1575
         Left            =   120
         Top             =   240
         Width           =   8655
      End
      Begin VB.Shape Shape1 
         Height          =   1335
         Left            =   120
         Top             =   1920
         Width           =   8655
      End
      Begin VB.Label Label19 
         Caption         =   "Remark          "
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
         Left            =   5400
         TabIndex        =   38
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label Label18 
         Caption         =   "Ply Details          "
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
         Left            =   5400
         TabIndex        =   37
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label17 
         Caption         =   "Returnable             "
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
         Left            =   5400
         TabIndex        =   36
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label16 
         Caption         =   "Net Wt.              "
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
         Left            =   3000
         TabIndex        =   31
         Top             =   2760
         Width           =   855
      End
      Begin VB.Label Label15 
         Caption         =   "Tare Wt.              "
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
         Left            =   3000
         TabIndex        =   30
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label Label14 
         Caption         =   "Gross Wt.              "
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
         Left            =   3000
         TabIndex        =   29
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label13 
         Caption         =   "Cops Wt.                "
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
         Left            =   360
         TabIndex        =   26
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label Label12 
         Caption         =   "Box Wt.                "
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
         Left            =   360
         TabIndex        =   25
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label Label11 
         Caption         =   "No. of Cops               "
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
         Left            =   360
         TabIndex        =   23
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "Finish Item     "
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
         Left            =   5160
         TabIndex        =   17
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "Sub Grade         "
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
         Left            =   5160
         TabIndex        =   16
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Grade          "
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
         Left            =   5160
         TabIndex        =   15
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Lot No              "
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
         Left            =   5160
         TabIndex        =   14
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Type of Packaging "
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
         Left            =   480
         TabIndex        =   9
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Type of Packing   "
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
         Left            =   480
         TabIndex        =   8
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Packing Station"
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
         Left            =   480
         TabIndex        =   7
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Division                "
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
         Left            =   480
         TabIndex        =   6
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8895
      Begin VB.TextBox txtvbdt 
         Height          =   285
         Left            =   6960
         TabIndex        =   2
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox txtBox 
         Height          =   285
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   1
         Top             =   240
         Width           =   2535
      End
      Begin WelchButton.lvButtons_H cmdGo 
         Height          =   375
         Left            =   3960
         TabIndex        =   33
         Top             =   240
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
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
         Image           =   "FRMBOXHISTORY.frx":0B34
         cBack           =   -2147483633
      End
      Begin VB.Label Label10 
         Caption         =   "Packing Date :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5280
         TabIndex        =   22
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Box No. :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
   End
End
Attribute VB_Name = "FRMBOXHISTORY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdGo_Click()

  Dim RS1 As New ADODB.Recordset
  Dim UNITCOD As String
  Dim DIVICOD As String
  Dim ITMCOD As String
  Dim PCOD As String
  Dim DTYP As String
  Dim DSRN As String
  Dim STCOD As String
  Dim GRD As Integer
  Dim SUBGRD As String
  Dim PKGNG As String
  Dim dbcd As String
  Dim DISPTYP As String
  Dim PARTY As String
  Dim CONSINEE As String
  Dim ADD As String
  Dim BROKER As String
  Dim DCOD As String
  Dim RDBC As String
  
Set RS = New ADODB.Recordset
  If RS.State = 1 Then RS.Close
  RS.Open "SELECT * FROM BOXREGISTER WHERE VBNO ='" & txtBox & "' AND RECSTAT <>'D'", CN, adOpenDynamic, adLockOptimistic
  If RS.EOF Then
    MsgBox "Box No. Not Exist"
    txtBox.SetFocus
    Exit Sub
  End If

  TXTVBDT = Format(RS!VBDT, "DD/MM/YYYY")
  RDBC = Trim(RS!RDBC & "")
  dbcd = Trim(RS!dbcd)
  PKGNG = Trim(RS!PKGNG_COD)
  
  If Trim(RS!dbcd) <> "000006" Then
      TXTGRAD = RS!grad & ""
  Else
       TXTGRAD = "WASTE"
  End If
  
  SUBGRD = Trim(RS!SUBGRD & "")
  
  
  GRD = Trim(RS!grad)
  STCOD = Trim(RS!PKG_STCOD)
  
  If (RS!ISRETURNABLE) = "Y" Then
  TXTRET.Text = "YES"
  ElseIf RS!ISRETURNABLE = "N" Then
  TXTRET = "NO"
  End If
  
  txtLTNo = RS!LOTNO & ""
    

  TXTCHLN = RS!RVBNO & ""
  PCOD = RS!PCOD & ""
  UNITCOD = RS!unit & ""
  DIVICOD = RS!DVCD & ""
  DTYP = RS!RVTYP & ""
 
  ITMCOD = RS!ICOD & ""
  
  TXTCOP = RS!COPS
  TXTCTWT = Format(RS!BOXWGT, ".000")
  TXTCPWT = Format(RS!COPSWGT, ".000")
  
  TXTGRWT = Format(RS!GRSWGT, ".000")
  TXTTRWT = Format(RS!TRWGT, ".000")
  txtNTWT = Format(RS!NTWGT, ".000")
  TXTRMRK = Trim(Mid(RS!RMRK & "" + Space(20), 1, 10))
  
  'FINISH ITEM MASTER
  
  If RS.State = 1 Then RS.Close
  RS.Open "SELECT * FROM FINITMMST WHERE COMP = '" & compPth & "' AND UNIT = '" & UNITCOD & "' AND DVCD = '" & DIVICOD & "' AND CODE='" & ITMCOD & "'", CN, adOpenDynamic, adLockOptimistic
  If Not RS.EOF Then
    TXTFINITM = RS!NAME
  End If
  
  If RS.State = 1 Then RS.Close
  RS.Open "SELECT * FROM ACCMST WHERE CODE='" & PCOD & "'", CN, adOpenDynamic, adLockOptimistic
  If Not RS.EOF Then
    txtParty = RS!NAME & ""
  End If
  
  'PACKING TYPE
  
  If RS.State = 1 Then RS.Close
  RS.Open "SELECT * FROM PCKMST WHERE CODE='" & STCOD & "'", CN, adOpenDynamic, adLockOptimistic
  If Not RS.EOF Then
    TXTPACKST = RS!NAME & ""
  End If
  
  'PACKAGING TYPE
  
  If RS.State = 1 Then RS.Close
  RS.Open "SELECT * FROM PKGNGMST WHERE CODE='" & PKGNG & "'", CN, adOpenDynamic, adLockOptimistic
  If Not RS.EOF Then
    TXTPKGTYP = RS!NAME & ""
  End If
  
  If dbcd <> "000006" Then
  'GRADE MASTER
  If RS.State = 1 Then RS.Close
  RS.Open "SELECT * FROM GRDMST WHERE CODE='" & GRD & "'", CN, adOpenDynamic, adLockOptimistic
  If Not RS.EOF Then
    TXTGRAD = RS!grad & ""
  End If
  Else
     TXTGRAD = "WASTE"
  End If
  
  
  If dbcd <> "000006" Then
  'SUBGRADE MASTER
  If RS.State = 1 Then RS.Close
  RS.Open "SELECT * FROM SUBGRDMST WHERE COMP = '" & compPth & "' AND UNIT = '" & UNITCOD & "' AND DVCD = '" & DIVICOD & "' AND GRAD = '" & GRD & "' AND SUBGRD ='" & SUBGRD & "'", CN, adOpenDynamic, adLockOptimistic
  If Not RS.EOF Then
    TXTSUBGRD = RS!NAME & ""
  End If
  
  Else
    TXTSUBGRD = "WASTE"
  End If
  
  'DIVISION MASTER
  
  If RS.State = 1 Then RS.Close
  RS.Open "SELECT * FROM DIVMST WHERE COMP = '" & compPth & "' AND UNIT = '" & UNITCOD & "' AND  CODE='" & DIVICOD & "'", CN, adOpenDynamic, adLockOptimistic
  If Not RS.EOF Then
    TXTDIVISION = RS!NAME & ""
  End If
  
  'SERIAL MASTER
  
  If RS.State = 1 Then RS.Close
  RS.Open "SELECT * FROM SERIALMASTER WHERE COMP = '" & compPth & "' AND UNIT = '" & UNITCOD & "' AND CODE='" & dbcd & "' AND VTYP IN ('PPF','OPN')", CN, adOpenDynamic, adLockOptimistic
  If Not RS.EOF Then
    TXTPACKTYP = RS!NAME & ""
  End If
  
  'DISPATCH DETAILS
  
  If RS1.State = 1 Then RS1.Close
  RS1.Open "SELECT * FROM SPTRAN WHERE COMP='" & compPth & "' AND unit = '" & UNITCOD & _
  "' and dvcd='" & DIVICOD & "' and vtyp='DPF' AND DBCD = '" & RDBC & "' and chln = '" & TXTCHLN & "' and RECSTAT <> 'D'", CN, adOpenDynamic, adLockOptimistic
  If Not RS1.EOF Then
    txtDONO = RS1!EXTRA2 & ""
    TXTORDN = RS1!EXTRA1 & ""
    DISPTYP = RS1!dbcd & ""
    PARTY = RS1!PCOD & ""
    ADD = RS1!ADDRESS & ""
    BROKER = RS1!BRCD & ""
    TXTCHLN = RS1!chln & ""
    TXTCHLNDT = RS1!CHDT & ""
    txtINVNO = RS1!SVBN & ""
    DCOD = RS1!DCOD & ""
End If
  
  'SERIAL MASTER
  
  If RS1.State = 1 Then RS1.Close
  RS1.Open "SELECT * FROM SERIALMASTER WHERE COMP = '" & compPth & "' AND UNIT  = '" & UNITCOD & "' AND  CODE='" & DISPTYP & "' AND VTYP IN ('DPF')", CN, adOpenDynamic, adLockOptimistic
  If Not RS1.EOF Then
    TXTDISPTYP = RS1!NAME & ""
  End If
  
  'FOR ADDRESS
  
  If RS1.State = 1 Then RS1.Close
  RS1.Open "SELECT * FROM PADDMST WHERE CODE = '" & DCOD & "' AND SRNO = '" & ADD & "' ", CN, adOpenDynamic, adLockOptimistic
  If Not RS1.EOF Then
    TXTADDRESS = RS1!ADDR & ""
  End If
  
  'PARTY
  
  If RS1.State = 1 Then RS1.Close
  RS1.Open "SELECT * FROM ACCMST WHERE CODE= '" & PARTY & "' ", CN, adOpenDynamic, adLockOptimistic
  If Not RS1.EOF Then
    txtParty = RS1!NAME & ""
    txtCONSINEE = RS1!NAME & ""
  End If
  
  'SALESMAN MASTER AND ORDER NO.
  
  If RS1.State = 1 Then RS1.Close
  RS1.Open "SELECT * FROM ORDTRN WHERE COMP = '" & compPth & "' AND UNIT = '" & UNITCOD & "' AND DONO = '" & txtDONO & "' ", CN, adOpenDynamic, adLockOptimistic
  If Not RS1.EOF Then
     If RS1!ISRETURNABLE = "Y" Then
    DISPRETTYP = "YES"
    Else
    DISPRETTYP = "NO"
    End If
  End If
  
  ' CATA B
  ' BROKER MASTER
  
  If RS1.State = 1 Then RS1.Close
  RS1.Open "SELECT * FROM REFMST WHERE CODE = '" & BROKER & "' AND CATA = 'B'", CN, adOpenDynamic, adLockOptimistic
  If Not RS1.EOF Then
    TXTAGENT = RS1!NAME & ""
  End If
End Sub

Private Sub Form_Load()
Call CenterChild(frm_Main, Me)
Call ColorComponent(Me)
End Sub

Private Sub txtBox_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdGo.SetFocus
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdclr_Click()
Call ClsData(FRMBOXHISTORY)
End Sub
