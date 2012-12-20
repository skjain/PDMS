VERSION 5.00
Object = "{A19332D7-D707-4A30-9F38-796D120AF5B3}#1.2#0"; "BtnPlus1.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{2BF72F7D-D367-4712-9146-5533EF3E691A}#1.2#0"; "FraPlus1.ocx"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WelchButton.ocx"
Begin VB.Form frmOrderReconcile 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Order Cancellation Module (Fully / Partially )"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11370
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   11370
   Begin FramePlusCtl.FramePlus FrmAutoConsumption 
      Height          =   6075
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   10716
      BackgroundPictureAlignment=   5
      BorderStyle     =   10
      BackColorGradient=   12640511
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
      Begin VB.TextBox RMRK 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   1680
         MaxLength       =   200
         TabIndex        =   7
         Top             =   4440
         Width           =   8895
      End
      Begin VB.TextBox TXTQNTY 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   9000
         MaxLength       =   10
         TabIndex        =   6
         Top             =   3600
         Width           =   1575
      End
      Begin VB.TextBox INVNO 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   6600
         MaxLength       =   10
         TabIndex        =   5
         Top             =   3600
         Width           =   1695
      End
      Begin VB.TextBox BALQTY 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   480
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   2
         Top             =   3600
         Width           =   1575
      End
      Begin VB.TextBox DISPQTY 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   2520
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   3
         Top             =   3600
         Width           =   1455
      End
      Begin VB.TextBox txtPCOD 
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
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   600
         Width           =   3495
      End
      Begin VB.TextBox txtBRCD 
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
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   960
         Width           =   3495
      End
      Begin VB.TextBox txttxcd 
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
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   1680
         Width           =   3495
      End
      Begin VB.TextBox txtVBNO 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00C00000&
         Height          =   325
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox txtVBDT 
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
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox txtTTQty 
         Alignment       =   1  'Right Justify
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
         Left            =   8040
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txtICOD 
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
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   1320
         Width           =   3495
      End
      Begin VB.TextBox TXTRATE 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Height          =   285
         Left            =   9840
         TabIndex        =   17
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox TXTOGRD 
         Alignment       =   1  'Right Justify
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
         Left            =   8040
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox TXTRATEFACTOR 
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
         Left            =   8040
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   1680
         Width           =   2895
      End
      Begin VB.TextBox TXTFREIGHT 
         Alignment       =   1  'Right Justify
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
         Left            =   8040
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   960
         Width           =   1095
      End
      Begin ButtonPlusCtl.ButtonPlus btnSelect 
         Height          =   330
         Left            =   1800
         TabIndex        =   0
         Top             =   960
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   582
         BackStyle       =   0
         BorderStyle     =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "..."
      End
      Begin MSComCtl2.DTPicker INVDT 
         Height          =   315
         Left            =   4440
         TabIndex        =   4
         Top             =   3600
         Width           =   1455
         _ExtentX        =   2566
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
         Format          =   50855937
         CurrentDate     =   40289
      End
      Begin WelchButton.lvButtons_H cmdEdit 
         Height          =   495
         Left            =   5640
         TabIndex        =   10
         Top             =   5280
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
         Image           =   "frmOrderReconcile.frx":0000
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdSave 
         Height          =   495
         Left            =   2880
         TabIndex        =   8
         Top             =   5280
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
         Image           =   "frmOrderReconcile.frx":039A
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdCancel 
         Height          =   495
         Left            =   4320
         TabIndex        =   9
         Top             =   5280
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
         Image           =   "frmOrderReconcile.frx":1124
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdExit 
         Height          =   495
         Left            =   6960
         TabIndex        =   11
         Top             =   5280
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
         Image           =   "frmOrderReconcile.frx":1576
         cBack           =   -2147483633
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         X1              =   120
         X2              =   11280
         Y1              =   4200
         Y2              =   4200
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cancellation Quantity"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   9000
         TabIndex        =   41
         Top             =   2880
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Balance Quantity"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   480
         TabIndex        =   40
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Dispatch Quantity"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2520
         TabIndex        =   39
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Cancellation /  Voucher No."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   6600
         TabIndex        =   38
         Top             =   2880
         Width           =   1455
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Cancellation Date"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4440
         TabIndex        =   37
         Top             =   2880
         Width           =   1695
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   36
         Top             =   4440
         Width           =   975
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         Height          =   2175
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   2760
         Width           =   11175
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         X1              =   120
         X2              =   11280
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         X1              =   2280
         X2              =   2280
         Y1              =   2760
         Y2              =   4200
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         X1              =   4200
         X2              =   4200
         Y1              =   2760
         Y2              =   4200
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         X1              =   6240
         X2              =   6240
         Y1              =   2760
         Y2              =   4200
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         X1              =   8640
         X2              =   8640
         Y1              =   2760
         Y2              =   4200
      End
      Begin VB.Label LBLHEAD 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         Caption         =   "ORDER CANCELLATION  ( FULLY / PARTIALLY )"
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
         Left            =   2040
         TabIndex        =   35
         Top             =   120
         Width           =   7095
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         Height          =   2295
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   240
         Width           =   11175
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "Tax Category "
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
         Left            =   1680
         TabIndex        =   34
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Agent "
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
         Left            =   2400
         TabIndex        =   33
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Order Date"
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
         TabIndex        =   32
         Top             =   1320
         Width           =   1125
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Party "
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
         Left            =   2400
         TabIndex        =   31
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Order Number"
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
         Left            =   240
         TabIndex        =   30
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Item"
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
         Left            =   2400
         TabIndex        =   29
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Order Qty"
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
         Left            =   6840
         TabIndex        =   28
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Rate"
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
         TabIndex        =   27
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Grade"
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
         Left            =   6840
         TabIndex        =   26
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Rate Desc."
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
         Left            =   6840
         TabIndex        =   25
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Freight/KG"
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
         Left            =   6840
         TabIndex        =   24
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label9 
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
         TabIndex        =   13
         Tag             =   "S"
         Top             =   -2040
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmOrderReconcile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public ORDBOK As String
Public ORDDBCD As String
Dim SQL As String
'GLOBAL CONSTANT
Dim M_BRCD As String, SCOMP As String, SUNIT As String, SDVCD  As String, SITM  As String, STAX As String, SGRD As String, SUBGRD As String, RATECOD As String, SPARTY As String
Dim SAVEFLAG As Boolean

Private Sub BALQTY_GotFocus()
  BALQTY.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub BALQTY_LostFocus()
  BALQTY.BackColor = vbWhite
End Sub

Private Sub btnSelect_Click()
    frmCancelOrderList.Show 1
    btnSelect.Enabled = False
    Call FindInfo
    INVNO.Text = GenDONO
End Sub

Private Sub cmdCancel_Click()
  ClsData (Me)
  btnSelect.Enabled = True
  btnSelect.SetFocus
End Sub

Private Sub cmdEdit_Click()
If M_USRSECLEVL = "1" Then
    If ReadConfigMaster("000029", 5, "M") = False Then ModuleDeniedMessage: Exit Sub
  End If
  SAVEFLAG = False
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
On Error GoTo LAST

If Not CHKSAVEDATA Then Exit Sub

Call SetGlobal
CN.BeginTrans

Dim INVNO As String
INVNO = GenDONO

SQL = "INSERT INTO ORDTRN (COMP,UNIT, DVCD,VTYP,DBCD,DONO,DODT,PCOD,DCOD,SRCH,BRCD,LTNO,GRAD,SUBGRD," & _
"QNTY,DELQNTY,RATE,ARAT,ORDN,OSRC,ORDQTY,ORDRATE,ORDDATE,BRMK,PRDL,ICOD,TXRT,TXCD,RTCD,FREIGHT_PERKG," & _
"FREIGHT_FACTOR,DFLG,DOSTAT,DOAPRVBY,DOAPRVDATE) VALUES ('" & SCOMP & _
"','" & SUNIT & "','" & SDVCD & "','DOS','" & ORDDBCD & "','" & INVNO & "','" & Format(INVDT, "MM/DD/YYYY") & _
"','" & SPARTY & "','','','" & M_BRCD & "','','" & SGRD & "','','" & Val(TXTQNTY) & _
"','" & Val(TXTQNTY) & "','0','0','" & TXTVBNO & "','1','" & Val(txtTTQty) & "','" & Val(TXTRATE) & _
"','" & Format(Trim(TXTVBDT), "MM/DD/YYYY") & "','" & RMRK & "','','" & SITM & "','','" & STAX & "','" & RATECOD & _
"','0','0','Y','Y','" & Trim(cUName) & "','" & Format(Now, "MM/DD/YYYY HH:MM:SS") & "')"

CN.Execute SQL

SQL = "UPDATE ORDMAN SET CANCELQTY = CANCELQTY + " & Val(TXTQNTY) & " WHERE COMP='" & SCOMP & _
"' AND UNIT='" & SUNIT & "' AND DCOD='" & SDVCD & "' AND DBCD='" & ORDDBCD & _
"' AND ORDN = '" & TXTVBNO & "' AND ICOD = '" & SITM & "' AND TRCD='" & SGRD & "'"

CN.Execute SQL
Call DAILYSTATUS("DOS", SPARTY, ORDDBCD, Val(TXTQNTY), TXTVBNO, 0, cUName, "N", Now, TXTVBDT)
CN.CommitTrans

MsgBox "Your Order Cancellation No. is : " & INVNO

Call cmdCancel_Click

Exit Sub
LAST:
 MsgBox ERR.Description
End Sub

Private Sub DISPQTY_GotFocus()
  DISPQTY.BackColor = RGB(BRED, BGREEN, BBLUE)
  DISPQTY.SelStart = 0
  DISPQTY.SelLength = Len(DISPQTY)
End Sub

Private Sub DISPQTY_LostFocus()
  DISPQTY.BackColor = vbWhite
End Sub

Private Sub Form_Activate()
  Call ColorComponent(Me)
  Me.BackColor = RGB(RED, GREEN, BLUE)
  TXTVBDT.ForeColor = vbWhite: txtPCOD.ForeColor = vbWhite: TXTBRCD.ForeColor = vbWhite: txtTXCD.ForeColor = vbWhite: TXTICOD.ForeColor = vbWhite: TXTOGRD.ForeColor = vbWhite: txtTTQty.ForeColor = vbWhite: TXTRATE.ForeColor = vbWhite: txtFREIGHT.ForeColor = vbWhite: TXTRATEFACTOR.ForeColor = vbWhite
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub Form_Load()
  Call CenterChild(frm_Main, Me)
  Call ColorComponent(Me)
  btnSelect.Enabled = True
  
  INVDT = Date
  
  INVDT.MaxDate = FEDT
  INVDT.MinDate = FSDT
  
  NEW_VISIBLE = False: CANCEL_VISIBLE = False:  M_DESC = Empty:  Key = Empty
  
  '-------SALESMAN MASTER
  ORDBOK = Empty: ORDDBCD = Empty
  ORDBOK = SearchList1("SELECT TOP 20 CODE,NAME FROM SALMANMST", 0, ORDBOK, "SELECT SALESMAN FROM LIST")
  If Key = Empty Then Exit Sub
  ORDDBCD = Key
  
  Me.Caption = Me.Caption + " BOOKED BY SALESMAN : " + ORDBOK

End Sub

Private Sub INVDT_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub INVNO_GotFocus()
  INVNO.BackColor = RGB(BRED, BGREEN, BBLUE)
  INVNO.SelStart = 0
  INVNO.SelLength = Len(INVNO)
End Sub

Private Sub INVNO_LostFocus()
  INVNO.BackColor = vbWhite
End Sub

Private Sub RMRK_GotFocus()
  RMRK.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub RMRK_LostFocus()
 RMRK.BackColor = vbWhite
End Sub

Private Sub TXTQNTY_GotFocus()
  TXTQNTY.BackColor = RGB(BRED, BGREEN, BBLUE)
  TXTQNTY.SelStart = 0
  TXTQNTY.SelLength = Len(TXTQNTY)
End Sub

Private Sub TXTQNTY_KeyPress(KeyAscii As Integer)
 If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 46 Then KeyAscii = 0
End Sub

Private Sub TXTQNTY_LostFocus()
  TXTQNTY.BackColor = vbWhite
End Sub

Private Function CHKSAVEDATA() As Boolean
  CHKSAVEDATA = True
  
  If TXTVBNO = Empty Or txtTTQty = Empty Or BALQTY = Empty Or Val(BALQTY) <= 0 Then
     MsgBox "Invalid Order Details "
     CHKSAVEDATA = False
     Exit Function
  End If
  
  If TXTQNTY = Empty Then
     MsgBox "Enter Valid Cancellation Quantity"
     CHKSAVEDATA = False
     Exit Function
  End If
  
  If Val(TXTQNTY) <= 0 Then
     MsgBox "Enter Valid Cancellation Quantity"
     If TXTQNTY.Enabled Then TXTQNTY.SetFocus
     CHKSAVEDATA = False
     Exit Function
  End If
  
  If INVNO = Empty Then
     MsgBox "Enter Valid Invoice Number"
     CHKSAVEDATA = False
     Exit Function
  End If
  
  If Val(TXTQNTY) > BALQTY Then
    MsgBox "Cancelled Qty Must Be Equal or Less then Balance Qty."
    CHKSAVEDATA = False
    TXTQNTY.Enabled = True
    TXTQNTY.SetFocus
    Exit Function
  End If
  
End Function

Private Sub FindInfo()
Dim INFORS As ADODB.Recordset
Set INFORS = New ADODB.Recordset

Call SetGlobal
Dim SQL As String

SQL = "SELECT  ISNULL(DISPATCHQTY,0) AS DISPATCH,ISNULL(QNTY - DOQTY - DISPATCHQTY - CANCELQTY,0) AS BALQTY FROM ORDMAN WHERE "
SQL = SQL & "COMP='" & SCOMP & "' AND UNIT='" & SUNIT & "' AND DCOD='" & SDVCD & _
"' AND DBCD='" & ORDDBCD & "' AND ORDN = '" & TXTVBNO & "' AND ICOD = '" & SITM & _
"' AND TRCD='" & SGRD & "' "

If INFORS.State = 1 Then INFORS.Close
INFORS.Open SQL, CN, adOpenDynamic, adLockOptimistic
If Not INFORS.EOF Then
    BALQTY = nstr(INFORS!BALQTY, 7, 3)
    DISPQTY = nstr(INFORS!DISPATCH, 7, 3)
End If
INFORS.Close
  
End Sub

Private Sub SetGlobal()
Dim GRRS As ADODB.Recordset
Set GRRS = New ADODB.Recordset

If GRRS.State = 1 Then GRRS.Close
GRRS.Open "SELECT * FROM ORDMAN WHERE COMP= '" & compPth & "' AND DBCD='" & ORDDBCD & "' AND ORDN ='" & TXTVBNO & "'", CN, adOpenDynamic, adLockOptimistic
If Not GRRS.EOF Then
   SCOMP = GRRS!COMP
   SUNIT = GRRS!unit
   SDVCD = GRRS!DCOD
   SITM = FindItemCode
   M_BRCD = GetCode("REFMST", TXTBRCD, "NAME", "CODE")
   STAX = GetCode("TAXMST", txtTXCD, "NAME", "CODE")
   SGRD = GetCode("GRDMST", TXTOGRD, "GRAD", "CODE")
   RATECOD = GetCode("RATEMST", TXTRATEFACTOR, "NAME", "CODE")
End If
GRRS.Close

SPARTY = GetCode("ACCMST", txtPCOD, "NAME", "CODE")

End Sub

Private Function FindItemCode() As String
Dim ITRS As ADODB.Recordset
Set ITRS = New ADODB.Recordset
Dim GRRS As ADODB.Recordset
Set GRRS = New ADODB.Recordset


If GRRS.State = 1 Then GRRS.Close
GRRS.Open "SELECT * FROM ORDMAN WHERE DBCD='" & ORDDBCD & "' AND ORDN ='" & TXTVBNO & "'", CN, adOpenDynamic, adLockOptimistic
If Not GRRS.EOF Then

If ITRS.State = 1 Then ITRS.Close
ITRS.Open "SELECT * FROM FINITMMST WHERE COMP='" & GRRS!COMP & "' AND UNIT='" & GRRS!unit & "' AND DVCD='" & GRRS!DCOD & "' AND NAME ='" & TXTICOD & "'", CN, adOpenDynamic, adLockOptimistic
If Not ITRS.EOF Then
   FindItemCode = ITRS!CODE
Else
   FindItemCode = Empty
End If
ITRS.Close

End If
GRRS.Close

End Function

Private Function GenDONO() As String
Dim DORS As ADODB.Recordset
Set DORS = New ADODB.Recordset
Dim NO As Double

If DORS.State = 1 Then DORS.Close
DORS.Open "SELECT ISNULL(MAX(RIGHT(DONO,4)),0) AS DONUM FROM ORDTRN WHERE ORDN='" & TXTVBNO & "'", CN, adOpenDynamic, adLockOptimistic

NO = Val(DORS!DONUM)
NO = NO + 1
DORS.Close
        
If NO < 10 Then
   GenDONO = "000" + Trim(nstr(NO, 1, 0))
ElseIf NO < 100 Then
   GenDONO = "00" + Trim(nstr(NO, 1, 0))
ElseIf NO < 1000 Then
   GenDONO = "0" + Trim(nstr(NO, 1, 0))
ElseIf NO < 10000 Then
   GenDONO = Trim(nstr(NO, 1, 0))
End If
      
   GenDONO = Mid$(CStr(TXTVBNO), 1, 6) & GenDONO

End Function

