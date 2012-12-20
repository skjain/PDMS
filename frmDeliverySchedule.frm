VERSION 5.00
Object = "{A19332D7-D707-4A30-9F38-796D120AF5B3}#1.2#0"; "BtnPlus1.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{2BF72F7D-D367-4712-9146-5533EF3E691A}#1.2#0"; "FraPlus1.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WELCHB~1.OCX"
Begin VB.Form frmDeliverySchedule 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Delivery Order Scheduling  Module"
   ClientHeight    =   6870
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11325
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   11325
   Begin FramePlusCtl.FramePlus FramePlus1 
      Height          =   6915
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   12197
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
      Begin VB.TextBox TXTCURSTK 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5280
         TabIndex        =   59
         Text            =   "0.000"
         Top             =   3360
         Width           =   1455
      End
      Begin VB.TextBox TXTEXRATE 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Height          =   285
         Left            =   9960
         TabIndex        =   57
         Top             =   1200
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CheckBox ChkReturnable 
         Caption         =   "Returnable Cops"
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
         Left            =   4080
         TabIndex        =   11
         Top             =   3840
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.TextBox M_ARAT 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   9840
         Locked          =   -1  'True
         MaxLength       =   200
         TabIndex        =   9
         Top             =   3360
         Width           =   1335
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
         Left            =   8160
         Locked          =   -1  'True
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   840
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
         Left            =   8160
         Locked          =   -1  'True
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   1560
         Width           =   2895
      End
      Begin VB.TextBox TXTDSPQTY 
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
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   6480
         Width           =   1335
      End
      Begin VB.TextBox TXTBALQTY 
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
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   6480
         Width           =   1335
      End
      Begin VB.ComboBox M_RTTX 
         Height          =   315
         ItemData        =   "frmDeliverySchedule.frx":0000
         Left            =   2160
         List            =   "frmDeliverySchedule.frx":000A
         TabIndex        =   10
         Text            =   "M_RTTX"
         Top             =   3840
         Width           =   1695
      End
      Begin VB.TextBox TXTADDRESS 
         Height          =   525
         Left            =   7560
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   2280
         Width           =   3615
      End
      Begin VB.TextBox txtConsinee 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   2400
         Width           =   3735
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
         Left            =   8160
         Locked          =   -1  'True
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox TXTRATE 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Height          =   285
         Left            =   9960
         TabIndex        =   39
         Top             =   480
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
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   1200
         Width           =   3495
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
         Left            =   8160
         Locked          =   -1  'True
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox BRMK 
         Height          =   285
         Left            =   7320
         MaxLength       =   100
         TabIndex        =   12
         Top             =   3840
         Width           =   2655
      End
      Begin VB.TextBox txtLTNO 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   3360
         Width           =   1455
      End
      Begin VB.TextBox TXTSUBGRD 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   3360
         Width           =   1095
      End
      Begin VB.TextBox M_DORAT 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   8280
         Locked          =   -1  'True
         MaxLength       =   200
         TabIndex        =   8
         Top             =   3360
         Width           =   1335
      End
      Begin VB.TextBox txtQTY 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6960
         TabIndex        =   7
         Top             =   3360
         Width           =   1095
      End
      Begin VB.TextBox txtDCOD 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   2400
         Width           =   3135
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
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox txtVBNO 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00C00000&
         Height          =   325
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   840
         Width           =   1455
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
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   1560
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
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   840
         Width           =   3495
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
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   480
         Width           =   3495
      End
      Begin WelchButton.lvButtons_H cmdSave 
         Height          =   495
         Left            =   7560
         TabIndex        =   15
         Top             =   6360
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
         Image           =   "frmDeliverySchedule.frx":002B
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdCancel 
         Height          =   495
         Left            =   8760
         TabIndex        =   16
         Top             =   6360
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
         Image           =   "frmDeliverySchedule.frx":0DB5
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdExit 
         Height          =   495
         Left            =   9960
         TabIndex        =   17
         Top             =   6360
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
         Image           =   "frmDeliverySchedule.frx":1207
         cBack           =   -2147483633
      End
      Begin ButtonPlusCtl.ButtonPlus btnSelect 
         Height          =   330
         Left            =   1920
         TabIndex        =   0
         Top             =   840
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
      Begin WelchButton.lvButtons_H cmdAdd 
         Height          =   375
         Left            =   10080
         TabIndex        =   13
         Top             =   3765
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
         Image           =   "frmDeliverySchedule.frx":1659
         cBack           =   -2147483633
      End
      Begin MSFlexGridLib.MSFlexGrid ITMFLEX 
         Height          =   1815
         Left            =   240
         TabIndex        =   47
         Top             =   4320
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   3201
         _Version        =   393216
         Cols            =   14
         FixedCols       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComCtl2.DTPicker dtDate 
         Height          =   315
         Left            =   360
         TabIndex        =   4
         Top             =   3360
         Width           =   1605
         _ExtentX        =   2831
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
         Format          =   55836673
         CurrentDate     =   39339
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Current Stock"
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
         Left            =   5280
         TabIndex        =   60
         Tag             =   "S"
         Top             =   3000
         Width           =   1455
      End
      Begin VB.Line Line13 
         BorderColor     =   &H00000080&
         X1              =   6840
         X2              =   6840
         Y1              =   2880
         Y2              =   3720
      End
      Begin VB.Label LBLEXRATE 
         BackStyle       =   0  'Transparent
         Caption         =   "Ex.Rate"
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
         Left            =   10080
         TabIndex        =   58
         Top             =   840
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
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
         Left            =   9840
         TabIndex        =   56
         Tag             =   "S"
         Top             =   3000
         Width           =   1215
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
         Left            =   6960
         TabIndex        =   54
         Top             =   840
         Width           =   1215
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
         Left            =   6960
         TabIndex        =   53
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Dispatch Qty."
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
         Left            =   3840
         TabIndex        =   49
         Top             =   6480
         Width           =   1335
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Balance Qty."
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
         Left            =   720
         TabIndex        =   48
         Top             =   6480
         Width           =   1335
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Retail/Tax Invoice"
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
         Left            =   480
         TabIndex        =   46
         Tag             =   "S"
         Top             =   3840
         Width           =   1695
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00000080&
         X1              =   7440
         X2              =   7440
         Y1              =   1920
         Y2              =   2880
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
         TabIndex        =   45
         Tag             =   "S"
         Top             =   2040
         Width           =   2055
      End
      Begin VB.Line Line12 
         BorderColor     =   &H00000080&
         X1              =   9720
         X2              =   9720
         Y1              =   2880
         Y2              =   3720
      End
      Begin VB.Line Line11 
         BorderColor     =   &H00000080&
         X1              =   8160
         X2              =   8160
         Y1              =   2880
         Y2              =   3720
      End
      Begin VB.Line Line10 
         BorderColor     =   &H00000080&
         X1              =   5160
         X2              =   5160
         Y1              =   2880
         Y2              =   3720
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00000080&
         X1              =   120
         X2              =   11280
         Y1              =   3720
         Y2              =   3720
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00000080&
         X1              =   3840
         X2              =   3840
         Y1              =   2880
         Y2              =   3720
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         Height          =   2295
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   1920
         Width           =   11175
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00000080&
         X1              =   2160
         X2              =   2160
         Y1              =   2880
         Y2              =   3720
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00000080&
         X1              =   120
         X2              =   11280
         Y1              =   3240
         Y2              =   3240
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00000080&
         X1              =   120
         X2              =   11280
         Y1              =   2880
         Y2              =   2880
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
         TabIndex        =   44
         Tag             =   "S"
         Top             =   1995
         Width           =   2175
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
         TabIndex        =   43
         Tag             =   "S"
         Top             =   2000
         Width           =   1455
      End
      Begin VB.Label LBLCFG1 
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
         Left            =   6960
         TabIndex        =   42
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label7 
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
         Left            =   9360
         TabIndex        =   40
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label6 
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
         Left            =   6960
         TabIndex        =   38
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label LBLHEAD 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         Caption         =   "DELIVERY ORDER SCHEDULE"
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
         Left            =   3240
         TabIndex        =   35
         Top             =   0
         Width           =   4455
      End
      Begin VB.Shape BottomHelp 
         BackColor       =   &H00C0C0FF&
         BackStyle       =   1  'Opaque
         Height          =   495
         Left            =   240
         Shape           =   4  'Rounded Rectangle
         Top             =   6360
         Width           =   6855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   6360
         TabIndex        =   34
         Top             =   3840
         Width           =   870
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00000080&
         X1              =   120
         X2              =   7440
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Delivery Date"
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
         Left            =   720
         TabIndex        =   33
         Tag             =   "S"
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000080&
         X1              =   4080
         X2              =   4080
         Y1              =   1920
         Y2              =   2880
      End
      Begin VB.Label Label12 
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
         Left            =   2520
         TabIndex        =   32
         Top             =   1200
         Width           =   495
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
         Left            =   2400
         TabIndex        =   31
         Tag             =   "S"
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Label LBLCFG2 
         Alignment       =   2  'Center
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
         Left            =   3840
         TabIndex        =   30
         Tag             =   "S"
         Top             =   3000
         Width           =   1335
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Rate (In Rs.)"
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
         Left            =   8280
         TabIndex        =   29
         Tag             =   "S"
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "DO Qnty."
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
         Left            =   6960
         TabIndex        =   28
         Tag             =   "S"
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         Height          =   2055
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   4200
         Width           =   11175
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         Height          =   1815
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   11175
      End
      Begin VB.Label Label3 
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
         Left            =   360
         TabIndex        =   22
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label4 
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
         Left            =   2520
         TabIndex        =   21
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label5 
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
         Left            =   480
         TabIndex        =   20
         Top             =   1200
         Width           =   1125
      End
      Begin VB.Label Label8 
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
         Left            =   2520
         TabIndex        =   19
         Top             =   840
         Width           =   735
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
         Left            =   1800
         TabIndex        =   18
         Top             =   1560
         Width           =   1455
      End
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00000080&
      X1              =   0
      X2              =   10920
      Y1              =   4920
      Y2              =   4920
   End
End
Attribute VB_Name = "frmDeliverySchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public ORDBOK As String
Public ORDDBCD As String
Public EXPREQ As Boolean
Dim SQL As String
Dim M_BRCD As String
Dim ROWNO As Long
Dim SWITCH As Boolean
'GLOBAL CONSTANT
Dim SCOMP As String, SUNIT  As String, SDVCD  As String, SITM  As String, STAX As String, SGRD As String, SUBGRD As String, RATECOD As String
'INTERNAL VARIABLE
Dim SPARTY As String, SCONSINEE As String, SADD As String

'ASSESABLE RATE
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
Dim M_BRKPERC As Double
Public ADD_SRNO As Long

Private Sub BRMK_GotFocus()
  BRMK.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub BRMK_LostFocus()
BRMK.BackColor = vbWhite
End Sub

Private Sub ChkReturnable_GotFocus()
  chkReturnable.FontSize = 10
End Sub

Private Sub chkReturnable_LostFocus()
  chkReturnable.FontSize = 8
End Sub

Private Sub cmdAdd_Click()
 Dim INDEX As Long
 
 If Not SWITCH Then
    ROWNO = ITMFLEX.Rows - 1
    If Not IsValidTaxInvoice Then Exit Sub
 End If
 
 If CheckData(ROWNO) Then Exit Sub
    
    If Trim(ITMFLEX.TextMatrix(ROWNO, 0)) = Empty Then
       ITMFLEX.TextMatrix(ROWNO, 0) = "XXXXXXXXXX"
    End If
      
    ITMFLEX.TextMatrix(ROWNO, 1) = Trim(txtCONSINEE)
    ITMFLEX.TextMatrix(ROWNO, 2) = Trim(txtDCOD)
    ITMFLEX.TextMatrix(ROWNO, 3) = Trim(TXTADDRESS)
    ITMFLEX.TextMatrix(ROWNO, 4) = Format(dtDate, "DD/MM/YYYY")
    ITMFLEX.TextMatrix(ROWNO, 5) = Trim(txtLTNo)
    ITMFLEX.TextMatrix(ROWNO, 6) = Trim(TXTSUBGRD)
    ITMFLEX.TextMatrix(ROWNO, 7) = Trim(nstr(Val(txtQty), 12, 3))
    ITMFLEX.TextMatrix(ROWNO, 8) = Trim(M_DORAT)
    ITMFLEX.TextMatrix(ROWNO, 9) = Trim(M_RTTX)
    ITMFLEX.TextMatrix(ROWNO, 10) = Trim(BRMK)
    ITMFLEX.TextMatrix(ROWNO, 12) = Trim(M_ARAT)
    ITMFLEX.TextMatrix(ROWNO, 13) = IIf(chkReturnable.Value = 1, "Y", "N")
             
    If MsgBox("Want to Add More", vbYesNo + vbDefaultButton2) = vbYes Then
          If ITMFLEX.TextMatrix(ITMFLEX.Rows - 1, 1) <> "" Then
             ITMFLEX.Rows = ITMFLEX.Rows + 1
          End If
           If ITMFLEX.Rows > 6 Then ITMFLEX.TopRow = ITMFLEX.TopRow + 2
        txtCONSINEE.SetFocus
    Else
        cmdSave.Enabled = True: cmdSave.SetFocus
    End If
            
    Call CLEARDATA
    cmdAdd.Caption = "&Add"
    SWITCH = False
End Sub

Private Sub cmdCancel_Click()
TXTVBNO = Empty: txtpcod = Empty: TXTBRCD = Empty: txtTXCD = Empty: TXTICOD = Empty
TXTOGRD = Empty: txtTTQty = Empty: txtRate = Empty: txtCONSINEE = Empty: txtDCOD = Empty
TXTADDRESS = Empty: dtDate = Date: txtLTNo = Empty: TXTSUBGRD = Empty: txtQty = Empty
M_DORAT = Empty: BRMK = Empty: TXTBALQTY = Empty: TXTDSPQTY = Empty: TXTVBDT = Empty
TXTFREIGHT = Empty: TXTRATEFACTOR = Empty
TXTCURSTK = "0.000"
dtDate.MinDate = FSDT
ITMFLEX.Rows = 1
ITMFLEX.Rows = 2
btnSelect.Enabled = True: btnSelect.SetFocus
End Sub

Private Sub CMDEXIT_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
On Error GoTo LAST
Dim i As Integer
Dim DONUM As String
Dim DOQTY As Double: DOQTY = 0

If Not IsValidData Then Exit Sub

With ITMFLEX

If .Rows < 2 Then Exit Sub
   
If .TextMatrix(1, 0) = Empty Then
  MsgBox "No Dispatch Entry Found !!", vbInformation, "Transaction Cancelled"
  Exit Sub
End If
   
Call SetGlobal
CN.BeginTrans

For i = 1 To .Rows - 1

If .TextMatrix(i, 0) = "XXXXXXXXXX" Then
   DONUM = GenDONO
Else
   DONUM = ""
End If

Call SetInternal(i)
SQL = "INSERT INTO ORDTRN (COMP,UNIT, DVCD,VTYP,DBCD,DONO,DODT,PCOD,DCOD,SRCH,BRCD,LTNO,GRAD,SUBGRD,QNTY,RATE,ARAT,"
SQL = SQL & "ORDN,OSRC,ORDQTY,ORDRATE,ORDDATE,BRMK,PRDL,ICOD,TXRT,TXCD,RTCD,FREIGHT_PERKG,FREIGHT_FACTOR,ISRETURNABLE) VALUES ('" & SCOMP & _
"','" & SUNIT & "','" & SDVCD & "','DOS','" & ORDDBCD & "','" & DONUM & "','" & Format(Trim(.TextMatrix(i, 4)), "MM/DD/YYYY") & _
"','" & SPARTY & "','" & SCONSINEE & "','" & SADD & "','" & M_BRCD & "','" & .TextMatrix(i, 5) & "','" & SGRD & _
"','" & SUBGRD & "','" & Val(.TextMatrix(i, 7)) & "'," & Val(.TextMatrix(i, 8)) & "," & Val(.TextMatrix(i, 12)) & ",'" & TXTVBNO & _
"','1','" & Val(txtTTQty) & "','" & Val(txtRate) & "','" & Format(Trim(TXTVBDT), "MM/DD/YYYY") & _
"','" & .TextMatrix(i, 10) & "','','" & SITM & "','" & .TextMatrix(i, 9) & "','" & STAX & _
"','" & RATECOD & "','" & Val(TXTFREIGHT) & "','" & nstr(FREIGHT, 8, 5) & "','" & Trim(.TextMatrix(i, 13)) & "')"

DOQTY = DOQTY + Val(.TextMatrix(i, 7))
CN.Execute SQL

Next i

SQL = "UPDATE ORDMAN SET DOQTY = DOQTY + " & DOQTY & " WHERE COMP='" & SCOMP & _
"' AND UNIT='" & SUNIT & "' AND DCOD='" & SDVCD & "' AND DBCD='" & ORDDBCD & _
"' AND ORDN = '" & TXTVBNO & "' AND ICOD = '" & SITM & "' AND TRCD='" & SGRD & "'"

CN.Execute SQL
'----------------------------------------
'DAILYSTATUS ENTRY
 Call DAILYSTATUS("DOS", SPARTY, ORDDBCD, Val(txtTTQty), TXTVBNO, 0, cUName, "N", Now, dtDate)
'-----------------------------------------
CN.CommitTrans

MsgBox "DO Saved Successfully", vbInformation, "Success"

End With

Call cmdCancel_Click

Exit Sub
LAST:
 MsgBox ERR.Description
 CN.RollbackTrans
End Sub

Private Sub dtDate_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Activate()
  If ORDDBCD = Empty Or ORDBOK = Empty Then
       MsgBox "Select Salesman"
       Unload Me
  End If

  Call ColorComponent(Me)
  TXTVBDT.ForeColor = vbWhite: txtpcod.ForeColor = vbWhite: TXTBRCD.ForeColor = vbWhite: txtTXCD.ForeColor = vbWhite: TXTICOD.ForeColor = vbWhite: TXTOGRD.ForeColor = vbWhite: txtTTQty.ForeColor = vbWhite: txtRate.ForeColor = vbWhite: TXTFREIGHT.ForeColor = vbWhite: TXTRATEFACTOR.ForeColor = vbWhite
  BottomHelp.BackColor = &HC0C0FF
  M_RTTX.ListIndex = 0
End Sub

Private Sub Form_Load()
  Call CenterChild(frm_Main, Me)
  Call ColorComponent(Me)
  Call SETFLEX
  btnSelect.Enabled = True
  BottomHelp.BackColor = &HC0C0FF
  
  dtDate = Date
  
  dtDate.MinDate = FSDT: dtDate.MaxDate = FEDT
    
  NEW_VISIBLE = False: CANCEL_VISIBLE = False:  M_DESC = Empty:  Key = Empty
    
  '-------SALESMAN MASTER
  ORDBOK = Empty: ORDDBCD = Empty
  ORDBOK = SearchList1("SELECT TOP 20 CODE,NAME FROM SALMANMST", 0, ORDBOK, "SELECT SALESMAN FROM LIST")
  If Key = Empty Then Exit Sub
  ORDDBCD = Key
  
  EXPREQ = EXP_REQ(ORDDBCD)
  
  If EXPREQ Then
     LBLEXRATE.Visible = True
     txtEXRate.Visible = True
  End If
  
  Me.Caption = Me.Caption + " BOOKED BY SALESMAN : " + ORDBOK
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If UCase(ActiveControl.NAME) = "TXTCONSINEE" And txtCONSINEE = Empty Then Exit Sub
  If UCase(ActiveControl.NAME) = "TXTDCOD" And txtDCOD = Empty Then Exit Sub
  If UCase(ActiveControl.NAME) = "TXTADDRESS" And TXTADDRESS = Empty Then Exit Sub
  If UCase(ActiveControl.NAME) = "TXTLTNO" And txtLTNo = Empty Then Exit Sub
  If UCase(ActiveControl.NAME) = "TXTSUBGRD" And TXTSUBGRD = Empty Then Exit Sub
  If UCase(ActiveControl.NAME) = "TXTQTY" And txtQty = Empty Then Exit Sub
  If UCase(ActiveControl.NAME) = "M_DORAT" And M_DORAT = Empty Then Exit Sub
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub btnSelect_Click()

    frm_SOBList.Show 1
    btnSelect.Enabled = False
    txtCONSINEE.Enabled = True: txtCONSINEE.SetFocus
    Call FindInfo
    If TXTVBDT <> Empty Then
      dtDate.MinDate = Format(TXTVBDT, "DD/MM/YYYY")
    Else
      Call cmdCancel_Click
    End If
    
End Sub

Private Sub ITMFLEX_Click()
   If ITMFLEX.Rows > 1 And ITMFLEX.TextMatrix(ITMFLEX.ROW, 1) <> Empty Then
    cmdAdd.Caption = "Upd&ate"
    ROWNO = ITMFLEX.ROW
       
    txtCONSINEE = ITMFLEX.TextMatrix(ROWNO, 1)
    txtDCOD = ITMFLEX.TextMatrix(ROWNO, 2)
    TXTADDRESS = ITMFLEX.TextMatrix(ROWNO, 3)
    dtDate = Format(ITMFLEX.TextMatrix(ROWNO, 4), "DD/MM/YYYY")
    txtLTNo = ITMFLEX.TextMatrix(ROWNO, 5)
    TXTSUBGRD = ITMFLEX.TextMatrix(ROWNO, 6)
    txtQty = ITMFLEX.TextMatrix(ROWNO, 7)
    M_DORAT = ITMFLEX.TextMatrix(ROWNO, 8)
    M_RTTX = ITMFLEX.TextMatrix(ROWNO, 9)
    BRMK = ITMFLEX.TextMatrix(ROWNO, 10)
    M_ARAT = ITMFLEX.TextMatrix(ROWNO, 12)
    
    SWITCH = True
  End If
   
End Sub

Private Sub ITMFLEX_EnterCell()
   ITMFLEX.CellBackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub ITMFLEX_LostFocus()
ITMFLEX.CellBackColor = vbWhite
End Sub

Private Sub M_ARAT_GotFocus()
M_ARAT.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub M_ARAT_LostFocus()
M_ARAT.BackColor = vbWhite
End Sub

Private Sub M_DORAT_Change()
  If Val(M_DORAT) <= 0 Then Exit Sub
  If TXTRATEFACTOR = Empty Then Exit Sub
  
  BROKERAGE_REQ = False:  M_BRKPERC = 0:  M_CD = 0: VAT = 0: CST = 0: EXCISE = 0
  
  If RS.State = 1 Then RS.Close
  RS.Open "SELECT * FROM REFMST WHERE NAME='" & TXTBRCD.Text & "'", CN, adOpenDynamic, adLockOptimistic
  If Not RS.EOF Then M_BRKPERC = RS!PERC
  
  Dim MID_RATE As Double
  Dim BSC_RATE As Double
  Dim BRK_AMNT As Double
  Dim RATE_FACTOR As Double
  
  If RS.State = 1 Then RS.Close
  RS.Open "SELECT * FROM RATEMST WHERE NAME='" & TXTRATEFACTOR & "'", CN, adOpenDynamic, adLockOptimistic
  If Not RS.EOF Then
    BROKERAGE_REQ = Trim(RS!BROKERAGE & "")
    M_CD = RS!CD
    RATE_FACTOR = RS!RATE_FACTOR
        
    If Val(RS!PERVAT & "") > 0 Then
       VAT = (1 + Round(Val(RS!PERVAT & "") / 100, 5))
    End If
    If Val(RS!PERCST & "") > 0 Then
       CST = (1 + Round(Val(RS!PERCST & "") / 100, 5))
    End If
    If Val(RS!PEREXCISE & "") > 0 Then
       EXCISE = (1 + Round(Val(RS!PEREXCISE & "") / 100, 5))
    End If
  End If
  
  MID_RATE = Val(M_DORAT)
  
  If BROKERAGE_REQ Then
    MID_RATE = Val(M_DORAT) - M_CD 'NET RATE REVISE CD LESS
    Dim FAC1 As Double:   FAC1 = 0
    FAC1 = (1 + Round(M_BRKPERC / 100, 5))
    If FAC1 > 0 Then
       MID_RATE = 1 + (MID_RATE / FAC1)
    Else
       MsgBox "Invalid Brokerage in Broker Master", vbCritical
       Exit Sub
    End If
  End If
  
     If Val(VAT + CST) > 0 Then
       FREIGHT = Val(TXTFREIGHT) / (VAT + CST)
       MID_RATE = MID_RATE / (VAT + CST)
     Else
       FREIGHT = Val(TXTFREIGHT)
       MID_RATE = MID_RATE
     End If
     
     MID_RATE = MID_RATE - FREIGHT
     
     If Val(EXCISE) > 0 Then
       ASSRATE = MID_RATE / EXCISE
     Else
       ASSRATE = MID_RATE
     End If
     
     M_ARAT = nstr(ASSRATE, 12, 4)
End Sub

Private Sub M_DORAT_GotFocus()
M_DORAT.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub M_DORAT_KeyPress(KeyAscii As Integer)
 If CheckNumericKey(KeyAscii, M_DORAT, Me) = 0 Then KeyAscii = 0
End Sub

Private Sub M_DORAT_LostFocus()
M_DORAT.BackColor = vbWhite
End Sub

Private Sub M_RTTX_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub TXTADDRESS_Change()
    txtLTNo = Empty
End Sub

Private Sub TXTADDRESS_GotFocus()
TXTADDRESS.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTADDRESS_KeyDown(KeyCode As Integer, Shift As Integer)
   If txtDCOD = Empty And txtDCOD.Enabled Then txtDCOD.SetFocus: Exit Sub
   TXTADDRESS.FontSize = 8
   If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
    TXTADDRESS = Empty
   ElseIf KeyCode = vbKeyF2 Or (TXTADDRESS = Empty And KeyCode = vbKeyReturn) Then
    TXTADDRESS = SearchAddress("Select SRNO,ADDR From PADDMST WHERE NAME='" & txtDCOD & "'", 0, Empty, "Select Consignee Address")
   End If
   
   Dim TEMPRS As New ADODB.Recordset
   If TEMPRS.State = 1 Then TEMPRS.Close
   TEMPRS.Open "SELECT * FROM PADDMST WHERE NAME='" & txtDCOD & "' AND ADDR='" & TXTADDRESS & "'", CN, adOpenDynamic, adLockOptimistic
   If Not TEMPRS.EOF Then
        ADD_SRNO = TEMPRS!SRNO
   End If
   
   Call SetLastConsigneeLot
   
End Sub

Private Sub TXTADDRESS_LostFocus()
TXTADDRESS.BackColor = vbWhite
End Sub

Private Sub txtDCOD_Change()
  TXTADDRESS = Empty
  txtLTNo = Empty
End Sub

Private Sub TXTDCOD_GotFocus()
txtDCOD.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTDCOD_KeyDown(KeyCode As Integer, Shift As Integer)
 Me.KeyPreview = False
  If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
     txtDCOD = Empty
  ElseIf KeyCode = vbKeyF2 Or txtDCOD = Empty Then
     M_DESC = Empty:   NEW_VISIBLE = False
     txtDCOD = SearchList1("Select DISTINCT CODE,NAME From PADDMST WHERE RECSTAT='A'", 0, Empty, "Select Consinee Name ")
  End If
 Me.KeyPreview = True
End Sub

Private Sub TXTDCOD_LostFocus()
txtDCOD.BackColor = vbWhite
End Sub

Private Sub txtLTNO_Change()
  Call FindCurStock
End Sub

Private Sub txtltno_GotFocus()
txtLTNo.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtltno_KeyDown(KeyCode As Integer, Shift As Integer)
Me.KeyPreview = False

Key = Empty
If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
     txtLTNo = Empty
ElseIf KeyCode = vbKeyF2 Or (KeyCode = 13 And txtLTNo = Empty) Then
   M_DESC = Empty:   NEW_VISIBLE = False
   SQL = GetSql("TXULOT")
   txtLTNo = SearchList(SQL)
End If

Me.KeyPreview = True
End Sub

Private Sub txtltno_LostFocus()
txtLTNo.BackColor = vbWhite
End Sub

Private Sub TXTQTY_GotFocus()
txtQty.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtQTY_KeyPress(KeyAscii As Integer)
 If CheckNumericKey(KeyAscii, txtQty, Me) = 0 Then KeyAscii = 0
End Sub

Private Sub TXTQTY_LostFocus()
txtQty.BackColor = vbWhite
End Sub

Private Sub TXTSUBGRD_Change()
  Call FindCurStock
End Sub

Private Sub TXTSUBGRD_GotFocus()
TXTSUBGRD.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTSUBGRD_KeyDown(KeyCode As Integer, Shift As Integer)
If IsTwistReq = "Y" Then Exit Sub

Me.KeyPreview = False
Key = Empty
If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
     TXTSUBGRD = Empty
ElseIf KeyCode = vbKeyF2 Or (KeyCode = 13 And TXTSUBGRD = Empty) Then
   M_DESC = Empty:   NEW_VISIBLE = False
   SQL = GetSql("SUBGRDMST")
   TXTSUBGRD = SearchList1(SQL, 0, Empty)
   M_DORAT = GetRateIfUnitReq
   If M_DORAT = 0 Then
      M_DORAT = Val(txtRate) - Val(Key)
   End If
   
   If EXPREQ Then
      M_DORAT = M_DORAT * Val(txtEXRate)
   End If
   
   M_DORAT = nstr(M_DORAT, 7, 2)
End If

Me.KeyPreview = True
End Sub

Private Sub TXTSUBGRD_KeyPress(KeyAscii As Integer)
If IsTwistReq = "N" Then Exit Sub
   Select Case KeyAscii
   Case Asc("s"), Asc("S")
        TXTSUBGRD = Empty
        KeyAscii = Asc("S")
   Case Asc("z"), Asc("Z")
        TXTSUBGRD = Empty
        KeyAscii = Asc("Z")
   Case Asc("0")
        TXTSUBGRD = Empty
        KeyAscii = Asc("0")
   Case Else
        KeyAscii = 0
   End Select
End Sub

Private Sub TXTSUBGRD_LostFocus()
TXTSUBGRD.BackColor = vbWhite
End Sub

Private Sub txtConsinee_GotFocus()
  txtCONSINEE.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtConsinee_LostFocus()
  txtCONSINEE.BackColor = vbWhite
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

Private Function GetSql(TABLE As String) As String
Dim GRRS As ADODB.Recordset
Set GRRS = New ADODB.Recordset

If GRRS.State = 1 Then GRRS.Close
GRRS.Open "SELECT * FROM ORDMAN WHERE DBCD='" & ORDDBCD & "' AND ORDN ='" & TXTVBNO & "'", CN, adOpenDynamic, adLockOptimistic
If Not GRRS.EOF Then
   
Select Case (TABLE)
Case "TXULOT"
  GetSql = "SELECT DISTINCT LTNO,LTNO FROM TXULOT WHERE COMP='" & GRRS!COMP & "' AND UNIT='" & GRRS!unit & "' AND DVCD='" & GRRS!DCOD & "' AND FICD='" & FindItemCode & "' AND ACTIVE='Y'"
Case "SUBGRDMST"
  GetSql = "SELECT DISTINCT RDIFF,NAME FROM SUBGRDMST WHERE COMP='" & GRRS!COMP & "' AND UNIT='" & GRRS!unit & "' AND DVCD='" & GRRS!DCOD & "' AND GRAD='" & GetCode("GRDMST", TXTOGRD, "GRAD", "CODE") & "'"
Case Else
  GetSql = Empty
End Select

End If
GRRS.Close
End Function

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

Private Sub SETFLEX()
  ITMFLEX.Clear
  ITMFLEX.ColWidth(0) = 1000
  ITMFLEX.ColWidth(1) = 2000
  ITMFLEX.ColWidth(2) = 2000
  ITMFLEX.ColWidth(3) = 0
  ITMFLEX.ColWidth(4) = 1000
  ITMFLEX.ColWidth(5) = 1000
  ITMFLEX.ColWidth(6) = 1000
  ITMFLEX.ColWidth(7) = 1200
  ITMFLEX.ColWidth(8) = 1000
  ITMFLEX.ColWidth(9) = 1600
  ITMFLEX.ColWidth(10) = 1600
  ITMFLEX.ColWidth(11) = 0
  ITMFLEX.ColWidth(12) = 0
  ITMFLEX.Clear
  ITMFLEX.TextMatrix(0, 0) = "D.O."
  ITMFLEX.TextMatrix(0, 1) = "Account Name"
  ITMFLEX.TextMatrix(0, 2) = "Consinee Name"
  ITMFLEX.TextMatrix(0, 3) = "Consinee Address"
  ITMFLEX.TextMatrix(0, 4) = "Del.Date"
  ITMFLEX.TextMatrix(0, 5) = "LotNo"
  ITMFLEX.TextMatrix(0, 6) = "SubGrade"
  ITMFLEX.TextMatrix(0, 7) = "DO Qty"
  ITMFLEX.TextMatrix(0, 8) = "Rate"
  ITMFLEX.TextMatrix(0, 9) = "Retail/Tax"
  ITMFLEX.TextMatrix(0, 10) = "Remarks"
  ITMFLEX.TextMatrix(0, 11) = "D.O STATUS"
  ITMFLEX.TextMatrix(0, 12) = "Ass.Rate"
  ITMFLEX.TextMatrix(0, 13) = "Returnable"
  'ITMFLEX.ColAlignment(0) = vbLeftJustify
  'ITMFLEX.ColAlignment(1) = vbLeftJustify
  'ITMFLEX.ColAlignment(2) = vbRightJustify
  'ITMFLEX.ColAlignment(3) = vbRightJustify
  'ITMFLEX.ColAlignment(4) = vbRightJustify
End Sub

  
Private Function CheckData(RNO As Long) As Boolean
Dim INDEX As Long
    If Trim(txtCONSINEE) = Empty Then
        txtCONSINEE.SetFocus
        CheckData = True
        Exit Function
    End If
    
    If Trim(txtDCOD) = Empty Then
        txtDCOD.SetFocus
        CheckData = True
        Exit Function
    End If
    
    
    If Trim(TXTADDRESS) = Empty Then
        TXTADDRESS.SetFocus
        CheckData = True
        Exit Function
    End If
    
    'If Trim(dtDate) = Empty Then
    '    dtDate.SetFocus
    '    CheckData = True
    '    Exit Function
    'End If
    
    If Trim(txtLTNo) = Empty Then
        txtLTNo.SetFocus
        CheckData = True
        Exit Function
    End If
    
    If Trim(TXTSUBGRD) = Empty Then
        TXTSUBGRD.SetFocus
        CheckData = True
        Exit Function
    End If
        
    If Trim(M_RTTX) = Empty Then
        M_RTTX.SetFocus
        CheckData = True
        Exit Function
    End If
        
    If Val(txtQty) = 0 Then
        txtQty.SetFocus
        CheckData = True
        Exit Function
    End If
        
    If Val(M_DORAT) = 0 Then
        M_DORAT.SetFocus
        CheckData = True
        Exit Function
    End If
    
    Dim TQTY As Double: TQTY = txtQty
    With ITMFLEX
    For INDEX = 1 To ITMFLEX.Rows - 1
      If .TextMatrix(INDEX, 7) <> Empty And (Not SWITCH Or (SWITCH And INDEX <> RNO)) Then
         TQTY = TQTY + .TextMatrix(INDEX, 7)
      End If
    Next INDEX
    End With
        
    If TQTY > TXTBALQTY Then
        txtQty.SetFocus
        CheckData = True
        Exit Function
    End If
    
End Function

Private Sub CLEARDATA()
        txtCONSINEE = Empty
        txtDCOD = Empty
        TXTADDRESS = Empty
        dtDate = Date
        txtLTNo = Empty
        If LBLCFG1.Caption <> "Shade" Then TXTSUBGRD = Empty
        txtQty = Empty
        TXTCURSTK = "0.00"
        'If LBLCFG1.Caption <> "Shade" Then M_DORAT = Empty
        'If LBLCFG1.Caption <> "Shade" Then M_ARAT = Empty
        BRMK = Empty
End Sub

Private Sub SetGlobal()
Dim GRRS As ADODB.Recordset
Set GRRS = New ADODB.Recordset

If GRRS.State = 1 Then GRRS.Close
GRRS.Open "SELECT * FROM ORDMAN WHERE DBCD='" & ORDDBCD & "' AND ORDN ='" & TXTVBNO & "'", CN, adOpenDynamic, adLockOptimistic
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

End Sub

Private Sub SetInternal(i As Integer)

SPARTY = GetCode("ACCMST", ITMFLEX.TextMatrix(i, 1), "NAME", "CODE")
SCONSINEE = GetCode("PADDMST", ITMFLEX.TextMatrix(i, 2), "NAME", "CODE")

Dim GRRS As ADODB.Recordset
Set GRRS = New ADODB.Recordset

If GRRS.State = 1 Then GRRS.Close
GRRS.Open "SELECT CODE,SRNO FROM PADDMST WHERE NAME='" & ITMFLEX.TextMatrix(i, 2) & "' AND ADDR='" & ITMFLEX.TextMatrix(i, 3) & "'", CN, adOpenDynamic, adLockOptimistic
If Not GRRS.EOF Then
  SADD = GRRS!SRNO
  SCONSINEE = GRRS!CODE
Else
  SADD = Empty
  SCONSINEE = Empty
End If
GRRS.Close

If IsTwistReq = "Y" Then
   SUBGRD = ITMFLEX.TextMatrix(i, 6)
Else
    If GRRS.State = 1 Then GRRS.Close
    GRRS.Open "SELECT * FROM SUBGRDMST WHERE COMP='" & SCOMP & "' AND UNIT ='" & SUNIT & "' AND DVCD = '" & SDVCD & "' AND GRAD = '" & SGRD & "' AND NAME = '" & ITMFLEX.TextMatrix(i, 6) & "'", CN, adOpenDynamic, adLockOptimistic
    If Not GRRS.EOF Then
       SUBGRD = Trim(GRRS!SUBGRD & "")
    End If
    GRRS.Close
End If

End Sub
'Dim SPARTY As String, SCONSINEE As String, SADD As String

Private Function GenDONO() As String
Dim DORS As ADODB.Recordset
Set DORS = New ADODB.Recordset
Dim NO As Double

If DORS.State = 1 Then DORS.Close
DORS.Open "SELECT ISNULL(MAX(RIGHT(DONO,4)),0) AS DONUM FROM ORDTRN WHERE COMP='" & SCOMP & _
          "' AND UNIT ='" & SUNIT & "' AND ORDN='" & TXTVBNO & "'", CN, adOpenDynamic, adLockOptimistic

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


Private Sub FindInfo()
Dim INFORS As ADODB.Recordset
Set INFORS = New ADODB.Recordset

Call SetGlobal

LBLCFG1.Caption = LabelDisplay(SDVCD, SUNIT)

If SetIsShadeReq(SDVCD) = "Y" Then
   LBLCFG2.Caption = "Shade"
   ITMFLEX.TextMatrix(0, 6) = "Shade"
End If


If LBLCFG1.Caption = "Shade" Then
   LBLCFG2.Caption = "Shade"
   TXTSUBGRD = TXTOGRD
   TXTSUBGRD.Enabled = False
   ITMFLEX.TextMatrix(0, 6) = "Shade"
End If

If IsTwistReq = "Y" Then
   LBLCFG2.Caption = "Twist{S/Z/0}"
   TXTSUBGRD.Locked = False
End If

SQL = "SELECT  ISNULL(DISPATCHQTY,0) AS DISPATCH,ISNULL(QNTY - DOQTY - DISPATCHQTY - CANCELQTY,0) AS BALQTY FROM ORDMAN WHERE "
SQL = SQL & "COMP='" & SCOMP & "' AND UNIT='" & SUNIT & "' AND DCOD='" & SDVCD & _
"' AND DBCD='" & ORDDBCD & "' AND ORDN = '" & TXTVBNO & "' AND ICOD = '" & SITM & _
"' AND TRCD='" & SGRD & "' "

If INFORS.State = 1 Then INFORS.Close
INFORS.Open SQL, CN, adOpenDynamic, adLockOptimistic
If Not INFORS.EOF Then
    TXTBALQTY = nstr(INFORS!BALQTY, 7, 3)
    TXTDSPQTY = nstr(INFORS!DISPATCH, 7, 3)
End If
INFORS.Close

txtEXRate = 1

If EXPREQ Then
   SQL = "SELECT EXRAT FROM EXPORD WHERE COMP='" & SCOMP & "' AND UNIT='" & SUNIT & _
         "' AND DBCD='" & ORDDBCD & "' AND ORDN='" & TXTVBNO & "'"
         
   If INFORS.State = 1 Then INFORS.Close
   INFORS.Open SQL, CN, adOpenDynamic, adLockOptimistic
   If Not INFORS.EOF Then
      txtEXRate = nstr(INFORS!EXRAT, 7, 2)
   End If
   INFORS.Close
End If

M_DORAT = GetRateIfUnitReq
If M_DORAT = 0 Then
   M_DORAT = Val(txtRate)
End If
M_DORAT = nstr(M_DORAT, 7, 2)
End Sub

Private Function GetRateIfUnitReq() As Double
On Error GoTo LAST

Dim CFGRS As ADODB.Recordset
Set CFGRS = New ADODB.Recordset
Dim RATEREQ As Boolean

Call SetGlobal
   
'DEAFULT
GetRateIfUnitReq = 0
RATEREQ = False

If CFGRS.State = 1 Then CFGRS.Close
CFGRS.Open "SELECT RATEMST_REQ FROM UNTCFG WHERE COMP='" & SCOMP & "' AND UNIT='" & SUNIT & "'", CN, adOpenDynamic, adLockOptimistic
If Not CFGRS.EOF Then
   If Trim(CFGRS!RATEMST_REQ & "") = "Y" Then
      RATEREQ = True
   Else
      Exit Function
   End If
End If
CFGRS.Close

If CFGRS.State = 1 Then CFGRS.Close
CFGRS.Open "SELECT * FROM SUBGRDMST WHERE COMP='" & SCOMP & "' AND UNIT ='" & SUNIT & _
           "' AND DVCD = '" & SDVCD & "' AND GRAD = '" & SGRD & "' AND NAME = '" & TXTSUBGRD & "'", CN, adOpenDynamic, adLockOptimistic
If Not CFGRS.EOF Then
   SUBGRD = Trim(CFGRS!SUBGRD & "")
End If
CFGRS.Close

If CFGRS.State = 1 Then CFGRS.Close

SQL = "SELECT RATE FROM RATMST WHERE COMP='" & SCOMP & "' AND UNIT='" & SUNIT & "' AND DVCD='" & SDVCD & _
          "' AND ICOD='" & SITM & "' AND GRAD='" & SGRD & "' "
          
If TXTSUBGRD.Enabled And TXTSUBGRD <> Empty Then
   SQL = SQL & "AND SUBGRD='" & SUBGRD & "' "
End If

SQL = SQL & "ORDER BY DATE DESC"
   
CFGRS.Open SQL, CN, adOpenDynamic, adLockOptimistic
If Not CFGRS.EOF Then
   GetRateIfUnitReq = Val(CFGRS!RATE)
End If
CFGRS.Close
Exit Function

LAST:
GetRateIfUnitReq = 0
Exit Function
End Function

Private Function EXP_REQ(SALMANCOD As String) As Boolean
EXP_REQ = False
Dim ISEXPORT As String
Dim EXPRS As ADODB.Recordset
Set EXPRS = New ADODB.Recordset

If EXPRS.State = 1 Then EXPRS.Close
EXPRS.Open "SELECT ISEXPORTORDER FROM SALMANMST WHERE CODE='" & SALMANCOD & "'", CN, adOpenDynamic, adLockOptimistic
If Not EXPRS.EOF Then
   ISEXPORT = Trim(EXPRS!ISEXPORTORDER & "")
End If

If ISEXPORT = "1" Then
  EXP_REQ = True
Else
  EXP_REQ = False
End If
End Function

Private Function IsValidTaxInvoice() As Boolean
On Error GoTo VALIDERR
IsValidTaxInvoice = True

SPARTY = GetCode("ACCMST", txtCONSINEE, "NAME", "CODE")

Dim VALIDRS As ADODB.Recordset
Set VALIDRS = New ADODB.Recordset

If VALIDRS.State = 1 Then VALIDRS.Close
VALIDRS.Open "SELECT * FROM ORDMAN WHERE DBCD='" & ORDDBCD & "' AND ORDN ='" & TXTVBNO & "'", CN, adOpenDynamic, adLockOptimistic
If Not VALIDRS.EOF Then
   SCOMP = VALIDRS!COMP
   SUNIT = VALIDRS!unit
   SITM = FindItemCode
End If
VALIDRS.Close

Dim AYS
If VALIDRS.State = 1 Then VALIDRS.Close
VALIDRS.Open "SELECT TOP 1 TXRT,BRCD,DCOD,ADDRESS FROM SPTRAN WHERE COMP='" & SCOMP & "' AND UNIT='" & SUNIT & _
             "' AND PCOD='" & SPARTY & "' AND ICOD='" & SITM & "' AND RECSTAT<>'D' AND TXRT IS NOT NULL " & _
             " AND VTYP='SAL' ORDER BY DATE DESC", CN, adOpenDynamic, adLockOptimistic
If Not VALIDRS.EOF Then
    If Trim(VALIDRS!TXRT & "") <> M_RTTX Then
       AYS = MsgBox("Your Last Invoice for this Party & Item is " & Trim(VALIDRS!TXRT & "") & ". Do You Want To Continue ", vbYesNo)
       If AYS = VBNO Then
         IsValidTaxInvoice = False
       End If
    End If
End If
VALIDRS.Close

If VALIDRS.State = 1 Then VALIDRS.Close
VALIDRS.Open "SELECT TOP 1 BRCD FROM SPTRAN " & _
             "INNER JOIN PADDMST ON PADDMST.CODE=SPTRAN.DCOD AND PADDMST.SRNO = SPTRAN.ADDRESS " & _
             "AND PADDMST.NAME= '" & txtDCOD & "' AND PADDMST.ADDR = '" & TXTADDRESS & _
             "' WHERE SPTRAN.COMP='" & SCOMP & "' AND SPTRAN.UNIT='" & SUNIT & _
             "' AND SPTRAN.RECSTAT<>'D' AND SPTRAN.VTYP='SAL' ORDER BY SPTRAN.DATE DESC", CN, adOpenDynamic, adLockOptimistic
If Not VALIDRS.EOF Then
    Dim INNERRS As ADODB.Recordset
    Set INNERRS = New ADODB.Recordset
    
    If INNERRS.State = 1 Then INNERRS.Close
    INNERRS.Open "SELECT NAME FROM REFMST WHERE CODE='" & Trim(VALIDRS!BRCD & "") & "' ", CN, adOpenDynamic, adLockOptimistic
    If Not INNERRS.EOF Then
       If Trim(INNERRS!NAME & "") <> Trim(TXTBRCD) Then
          AYS = MsgBox("Your Last Invoice for this Consignee Address having Another Agent. Do You Want To Continue ", vbYesNo)
          If AYS = VBNO Then
             IsValidTaxInvoice = False
          End If
       End If
    End If
    INNERRS.Close
End If
VALIDRS.Close

Exit Function
VALIDERR:
MsgBox ERR.Description
IsValidTaxInvoice = False
End Function

Public Function IsTwistReq() As String
IsTwistReq = "N"
Dim TMPRS As ADODB.Recordset
Set TMPRS = New ADODB.Recordset
If TMPRS.State = 1 Then TMPRS.Close
TMPRS.Open "SELECT TWSTREQ FROM DIVMST WHERE COMP='" & compPth & "' AND UNIT='" & SUNIT & _
           "' AND CODE='" & SDVCD & "' AND RECSTAT='A' AND TWSTREQ='Y'", CN, adOpenDynamic, adLockOptimistic
If Not TMPRS.EOF Then
  IsTwistReq = "Y"
End If
TMPRS.Close
End Function

Private Sub SetLastConsigneeLot()

    Dim LOTRS As ADODB.Recordset
    Set LOTRS = New ADODB.Recordset
    Dim ORDN_DVCD As String
    
    ORDN_DVCD = GetCode("ORDMAN", Trim(TXTVBNO), "ORDN", "DCOD")
            
    If LOTRS.State = 1 Then LOTRS.Close
    LOTRS.Open "SELECT LTNO FROM SPTRAN " & _
               "INNER JOIN FINITMMST ON FINITMMST.COMP = SPTRAN.COMP AND FINITMMST.UNIT = SPTRAN.UNIT " & _
               "AND FINITMMST.DVCD = SPTRAN.DVCD AND FINITMMST.CODE = SPTRAN.ICOD AND FINITMMST.NAME = '" & Trim(TXTICOD) & _
               "' WHERE SPTRAN.COMP='" & compPth & "' AND SPTRAN.UNIT='" & UNCD & _
               "' AND SPTRAN.DVCD='" & ORDN_DVCD & "' AND SPTRAN.DCOD ='" & GetCode("PADDMST", Trim(txtDCOD), "NAME", "CODE") & "' AND SPTRAN.ADDRESS=" & ADD_SRNO & _
               " AND VTYP='DPF' AND RECSTAT<>'D' ORDER BY DATE DESC", CN, adOpenDynamic, adLockOptimistic
    If Not LOTRS.EOF Then
       txtLTNo = Trim(LOTRS!ltno & "")
    End If
    LOTRS.Close

End Sub


Private Function FindCurStock() As Double
FindCurStock = 0

If TXTICOD = Empty Then Exit Function
If TXTOGRD = Empty Then Exit Function
If txtLTNo = Empty Then Exit Function

If LBLCFG2.Caption = "SubGrade" And IsTwistReq = "N" Then
   If TXTSUBGRD = Empty Then Exit Function
End If

Dim ORDN_DVCD As String
ORDN_DVCD = GetCode("ORDMAN", Trim(TXTVBNO), "ORDN", "DCOD")

Dim SQL As String
Dim CURRS As ADODB.Recordset
Set CURRS = New ADODB.Recordset

SQL = "SELECT BOXREGISTER.COMP,BOXREGISTER.UNIT,BOXREGISTER.DVCD,FINITMMST.NAME AS ITEM,GRDMST.GRAD, "

If LBLCFG2.Caption = "SubGrade" And IsTwistReq = "N" Then
   SQL = SQL & "SUBGRDMST.NAME, "
End If

SQL = SQL & "BOXREGISTER.LOTNO,ISNULL(SUM(NTWGT),0) AS FNSTK FROM BOXREGISTER " & _
      "INNER JOIN FINITMMST ON FINITMMST.COMP = BOXREGISTER.COMP AND FINITMMST.UNIT = BOXREGISTER.UNIT AND " & _
      "FINITMMST.DVCD = BOXREGISTER.DVCD And FINITMMST.CODE = BOXREGISTER.ICOD " & _
      "AND NAME='" & TXTICOD & "'  " & _
      "INNER JOIN GRDMST ON GRDMST.CODE=BOXREGISTER.GRAD " & _
      "AND GRDMST.GRAD ='" & TXTOGRD & "'  "
      
If LBLCFG2.Caption = "SubGrade" And IsTwistReq = "N" Then
   SQL = SQL & "INNER JOIN SUBGRDMST ON BOXREGISTER.COMP = SUBGRDMST.COMP AND BOXREGISTER.UNIT = SUBGRDMST.UNIT AND " & _
   "BOXREGISTER.DVCD = SUBGRDMST.DVCD AND BOXREGISTER.GRAD = SUBGRDMST.GRAD AND BOXREGISTER.SUBGRD = SUBGRDMST.SUBGRD " & _
   " AND SUBGRDMST.NAME='" & TXTSUBGRD & "' "
End If
      
SQL = SQL & "WHERE BOXREGISTER.COMP='" & compPth & "' AND BOXREGISTER.UNIT='" & UNCD & _
      "' AND BOXREGISTER.DVCD='" & ORDN_DVCD & "' AND BOXREGISTER.RECSTAT<>'D' AND VTYP IN ('PPF','OPN') " & _
      " AND BOXREGISTER.LOTNO='" & txtLTNo & "' GROUP BY BOXREGISTER.COMP,BOXREGISTER.UNIT,BOXREGISTER.DVCD,FINITMMST.NAME,GRDMST.GRAD,BOXREGISTER.LOTNO "
      
If LBLCFG2.Caption = "SubGrade" And IsTwistReq = "N" Then
   SQL = SQL & ",SUBGRDMST.NAME "
End If
      
If CURRS.State = 1 Then CURRS.Close
CURRS.Open SQL, CN, adOpenDynamic, adLockOptimistic
If Not CURRS.EOF Then
   TXTCURSTK = Trim(nstr(Val(CURRS!FNSTK), 12, 3))
Else
   TXTCURSTK = "0.000"
End If
CURRS.Close

End Function

Private Function IsValidData() As Boolean

IsValidData = True

Dim J As Long

Dim VALIDRS As ADODB.Recordset
Set VALIDRS = New ADODB.Recordset

For J = 1 To ITMFLEX.Rows - 1

If VALIDRS.State = 1 Then VALIDRS.Close
VALIDRS.Open "SELECT * FROM ACCMST WHERE NAME='" & ITMFLEX.TextMatrix(J, 1) & "'", CN, adOpenDynamic, adLockOptimistic
If VALIDRS.EOF Then
   IsValidData = False
   ITMFLEX.ROW = J
   ITMFLEX.COL = 1
   ITMFLEX.SetFocus
   MsgBox "Party Not define Properly", vbCritical
   ITMFLEX.SetFocus
   Exit Function
End If
VALIDRS.Close

If VALIDRS.State = 1 Then VALIDRS.Close
VALIDRS.Open "SELECT * FROM PADDMST WHERE NAME='" & ITMFLEX.TextMatrix(J, 2) & _
             "' AND ADDR='" & ITMFLEX.TextMatrix(J, 3) & "'", CN, adOpenDynamic, adLockOptimistic
If VALIDRS.EOF Then
   IsValidData = False
   ITMFLEX.ROW = J
   ITMFLEX.COL = 2
   ITMFLEX.SetFocus
   MsgBox "Consignee Not define Properly", vbCritical
   ITMFLEX.SetFocus
   Exit Function
End If
VALIDRS.Close

Next J

End Function
