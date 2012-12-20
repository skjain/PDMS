VERSION 5.00
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WelchButton.ocx"
Begin VB.Form frm_Item 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Master"
   ClientHeight    =   4860
   ClientLeft      =   3360
   ClientTop       =   2235
   ClientWidth     =   6930
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   6930
   Begin VB.Frame FRAMCMD 
      Height          =   855
      Left            =   120
      TabIndex        =   66
      Top             =   3840
      Width           =   6735
      Begin WelchButton.lvButtons_H cmdEdit 
         Height          =   495
         Left            =   3360
         TabIndex        =   48
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
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
         Image           =   "frm_Item.frx":0000
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdDelETE 
         Height          =   495
         Left            =   4440
         TabIndex        =   49
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
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
         Image           =   "frm_Item.frx":039A
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdSave 
         Height          =   495
         Left            =   1200
         TabIndex        =   46
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
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
         Image           =   "frm_Item.frx":0734
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdCancel 
         Height          =   495
         Left            =   2280
         TabIndex        =   47
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
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
         Image           =   "frm_Item.frx":14BE
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdExit 
         Height          =   495
         Left            =   5520
         TabIndex        =   50
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
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
         Image           =   "frm_Item.frx":1910
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdAdd 
         Height          =   495
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
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
         Image           =   "frm_Item.frx":1D62
         cBack           =   -2147483633
      End
   End
   Begin VB.Frame frmPlanning 
      Height          =   855
      Left            =   120
      TabIndex        =   58
      Top             =   6240
      Visible         =   0   'False
      Width           =   6705
      Begin VB.TextBox txtGainQty 
         Height          =   285
         Left            =   5160
         TabIndex        =   65
         Top             =   480
         Width           =   1140
      End
      Begin VB.TextBox txtAvgQty 
         Height          =   285
         Left            =   3600
         TabIndex        =   64
         Top             =   480
         Width           =   1140
      End
      Begin VB.TextBox txtitm 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   63
         Top             =   480
         Width           =   2820
      End
      Begin VB.CheckBox chkPlanning 
         Caption         =   "Yarn Dyg."
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
         TabIndex        =   59
         Top             =   0
         Width           =   1695
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Avg.Gain Qty"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5160
         TabIndex        =   62
         Top             =   240
         Width           =   1140
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Avg.Waste Qty"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3480
         TabIndex        =   61
         Top             =   240
         Width           =   1290
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Raw Item"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1200
         TabIndex        =   60
         Top             =   240
         Width           =   810
      End
   End
   Begin VB.Frame FramHead 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   120
      TabIndex        =   53
      Top             =   0
      Width           =   6705
      Begin VB.Label lblHead 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Item Master"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   0
         TabIndex        =   54
         Top             =   120
         Width           =   6705
      End
   End
   Begin VB.Frame FramCont 
      Height          =   3225
      Left            =   120
      TabIndex        =   55
      Top             =   600
      Width           =   6705
      Begin VB.OptionButton OPTOTH 
         Caption         =   "Others"
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
         Left            =   3960
         TabIndex        =   39
         Top             =   3480
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.OptionButton OPTTAKA 
         Caption         =   "Grey Fabrics"
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
         Left            =   1800
         TabIndex        =   38
         Top             =   3480
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.OptionButton OPTBEAM 
         Caption         =   "Beam "
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
         Top             =   3480
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CheckBox CHKRPL 
         Alignment       =   1  'Right Justify
         Caption         =   "Is it Replaceable Item ?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         TabIndex        =   42
         Top             =   3960
         Visible         =   0   'False
         Width           =   2460
      End
      Begin VB.ComboBox cmbLocation 
         Height          =   315
         Left            =   1350
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   2790
         Width           =   4605
      End
      Begin VB.CheckBox chkAutoAprvdINDT 
         Alignment       =   1  'Right Justify
         Caption         =   "Auto Approved Item Indent ?"
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
         Left            =   3120
         TabIndex        =   45
         Top             =   4080
         Visible         =   0   'False
         Width           =   2835
      End
      Begin VB.TextBox txtSALPER 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1350
         TabIndex        =   44
         Top             =   13680
         Width           =   1155
      End
      Begin VB.ComboBox cboUnit 
         Height          =   315
         Left            =   1350
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1290
         Width           =   1275
      End
      Begin VB.CheckBox chkAWIP 
         Alignment       =   1  'Right Justify
         Caption         =   "Item Affect W.I.P"
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   120
         TabIndex        =   51
         Top             =   5475
         Visible         =   0   'False
         Width           =   2565
      End
      Begin VB.CheckBox chkDeleted 
         Alignment       =   1  'Right Justify
         Caption         =   "Deleted Item"
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   2760
         TabIndex        =   52
         Top             =   5355
         Visible         =   0   'False
         Width           =   2610
      End
      Begin VB.TextBox txtROrder 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4920
         TabIndex        =   11
         Top             =   1320
         Width           =   960
      End
      Begin VB.TextBox txtSalRate 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1350
         TabIndex        =   25
         Top             =   2400
         Width           =   1275
      End
      Begin VB.TextBox txtPurRate 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4920
         TabIndex        =   27
         Top             =   2400
         Width           =   960
      End
      Begin VB.TextBox txtGroup 
         Height          =   285
         Left            =   1350
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   555
         Width           =   4500
      End
      Begin VB.TextBox txtOPNP 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4050
         TabIndex        =   23
         Top             =   5760
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.TextBox txtOPNQ 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1350
         TabIndex        =   21
         Top             =   5760
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.CommandButton cmdList 
         Caption         =   "..."
         Height          =   315
         Left            =   6000
         TabIndex        =   56
         Top             =   240
         Width           =   405
      End
      Begin VB.TextBox txtCode 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1350
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   6
         Top             =   930
         Width           =   1815
      End
      Begin VB.CheckBox chkVALID 
         Alignment       =   1  'Right Justify
         Caption         =   "Allow Entry?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   40
         Top             =   3720
         Visible         =   0   'False
         Width           =   2460
      End
      Begin VB.CheckBox chkCHNS 
         Alignment       =   1  'Right Justify
         Caption         =   "Check Negative Stock ?"
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
         Left            =   3120
         TabIndex        =   41
         Top             =   3720
         Visible         =   0   'False
         Width           =   2835
      End
      Begin VB.TextBox txtQty 
         Height          =   285
         Left            =   5160
         MaxLength       =   1
         TabIndex        =   19
         Text            =   "Q"
         Top             =   2040
         Width           =   690
      End
      Begin VB.TextBox txtRate 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1350
         MaxLength       =   6
         TabIndex        =   17
         Top             =   2040
         Width           =   1275
      End
      Begin VB.TextBox txtULIM 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4920
         MaxLength       =   8
         TabIndex        =   15
         Top             =   1680
         Width           =   960
      End
      Begin VB.TextBox txtLLIM 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1350
         MaxLength       =   8
         TabIndex        =   13
         Top             =   1680
         Width           =   1275
      End
      Begin VB.TextBox txtUnit 
         Height          =   285
         Left            =   1350
         MaxLength       =   10
         TabIndex        =   9
         Top             =   1320
         Width           =   1275
      End
      Begin VB.TextBox txtDesc 
         Height          =   285
         Left            =   1350
         MaxLength       =   250
         TabIndex        =   2
         ToolTipText     =   "Enter the Description of Item."
         Top             =   195
         Width           =   4515
      End
      Begin VB.Frame framRate 
         Caption         =   "Excise  :"
         Height          =   810
         Left            =   6000
         TabIndex        =   57
         Top             =   6480
         Width           =   3795
         Begin VB.CheckBox chkITEX 
            Alignment       =   1  'Right Justify
            Caption         =   "Excisable ?"
            Height          =   240
            Left            =   120
            TabIndex        =   36
            Top             =   1320
            Width           =   1590
         End
         Begin VB.TextBox txtSRAT 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1530
            TabIndex        =   35
            Top             =   975
            Width           =   900
         End
         Begin VB.TextBox txtERAT 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1530
            TabIndex        =   34
            Top             =   630
            Width           =   900
         End
         Begin VB.TextBox txtBRAT 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Left            =   1530
            TabIndex        =   31
            Top             =   300
            Width           =   900
         End
         Begin VB.Label lblSRAT 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sales Tax:"
            Height          =   195
            Left            =   150
            TabIndex        =   33
            Top             =   1020
            Width           =   915
         End
         Begin VB.Label lblERAT 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Exicse Rate:"
            Height          =   195
            Left            =   150
            TabIndex        =   32
            Top             =   690
            Width           =   1065
         End
         Begin VB.Label lblBRrat 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Basic Rate:"
            Height          =   195
            Left            =   120
            TabIndex        =   30
            Top             =   315
            Width           =   975
         End
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Location :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   28
         Top             =   2820
         Width           =   870
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sales Tax %"
         Height          =   195
         Left            =   120
         TabIndex        =   43
         Top             =   13710
         Width           =   1155
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Code :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   5
         Top             =   930
         Width           =   570
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Re-Order On:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3600
         TabIndex        =   10
         Top             =   1335
         Width           =   1140
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sal. Rate :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   24
         Top             =   2430
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pur. Rate:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3960
         TabIndex        =   26
         Top             =   2415
         Width           =   885
      End
      Begin VB.Label lblOPNP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Opening Pcs:"
         Height          =   195
         Left            =   2760
         TabIndex        =   22
         Top             =   5760
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Label lblOPNQ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Opn. Qty.:"
         Height          =   195
         Left            =   150
         TabIndex        =   20
         Top             =   5790
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label lblQORP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity/Pieces:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3480
         TabIndex        =   18
         Top             =   2160
         Width           =   1440
      End
      Begin VB.Label lblRate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rate Per :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   16
         Top             =   2055
         Width           =   885
      End
      Begin VB.Label lblULIM 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Upper Limit:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3840
         TabIndex        =   14
         Top             =   1710
         Width           =   1035
      End
      Begin VB.Label lblLLIM 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lower Limit:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   12
         Top             =   1695
         Width           =   1035
      End
      Begin VB.Label lblUnit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "U/M."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   7
         Top             =   1320
         Width           =   450
      End
      Begin VB.Label lblGroup 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Group:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   3
         Top             =   585
         Width           =   585
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   1
         Top             =   225
         Width           =   1035
      End
   End
End
Attribute VB_Name = "frm_Item"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim M_IGCD As String
Dim SAVEFLAG As Boolean, iGrpFlag As Boolean
Dim SQL As String

Private Sub cboUnit_Click()
    If cboUnit.Text = "(New)" Then
        Set LastFrm = New Frm_Ref_FAS
        Ref_Cat = "U"
        LOAD LastFrm
        LastFrm.Tag = Ref_Cat
        LastFrm.Show
        Exit Sub
    End If
    txtUNIT = cboUnit
End Sub

Private Sub cboUnit_GotFocus()
    cboUnit.BackColor = RGB(BRED, BGREEN, BBLUE)
    Msg "Select Unit of Measurement"
    
    If cboUnit.Text = "(New)" Then
        Call FillCmb("Select Name From REFMST Where Cata='U'", cboUnit)
        cboUnit.AddItem "(New)"
    End If
End Sub

Private Sub cboUnit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub cboUnit_LostFocus()
    cboUnit.BackColor = vbWhite
    Msg ""
End Sub

Private Sub chkAutoAprvdINDT_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub chkAWIP_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then chkDeleted.SetFocus
End Sub

Private Sub chkCHNS_GotFocus()
    Msg "Check Whether Stock Is Negative"
End Sub

Private Sub chkCHNS_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
'        If txtSALPER.Visible = True Then
'            txtSALPER.SetFocus
'        Else
'            cmdSave.SetFocus
'        End If
    End If
End Sub

Private Sub chkDeleted_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{taB}"
End Sub

Private Sub chkITEX_GotFocus()
    Msg "Please Tick If Item Excisable"
End Sub

Private Sub chkITEX_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then chkCHNS.SetFocus
End Sub

Private Sub chkPlanning_KeyDown(KeyCode As Integer, Shift As Integer)
If chkPlanning.Value = 1 And KeyCode = vbKeyReturn Then
   TXTITM.SetFocus
ElseIf KeyCode = 32 And chkPlanning.Value = 0 Then
   chkPlanning.Value = 1
Else
   CMDSAVE.SetFocus
End If
End Sub

Private Sub CHKRPL_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub chkVALID_GotFocus()
    Msg "Allow Entry Even Stock Is Negative"
End Sub

Private Sub chkVALID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then chkCHNS.SetFocus
End Sub

Private Sub cmbGroup_GotFocus()
    Msg "Press <F3> to Create new Item Group."
End Sub

Private Sub cmbGroup_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtUNIT.SetFocus
End Sub

Private Sub cmbGroup_LostFocus()
    iGrpFlag = False
End Sub

Private Sub cmbLocation_GotFocus()
    cmbLocation.BackColor = RGB(BRED, BGREEN, BBLUE)
    If SAVEFLAG = True Then
        Call FillCmb("SELECT LOCNAME FROM LOCATION ORDER BY LOCNAME", cmbLocation)
    End If
End Sub

Private Sub cmbLocation_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And frmPlanning.Visible = False And CMDSAVE.Enabled = True Then
       CMDSAVE.SetFocus
    Else
       chkPlanning.Enabled = True
       If chkPlanning.Visible = True Then
         chkPlanning.SetFocus
       End If
    End If
End Sub

Private Sub cmbLocation_LostFocus()
cmbLocation.BackColor = vbWhite
End Sub

Private Sub cmdAdd_Click()
    Call btn_sts(False)
    txtQTY.Text = "Q"
    txtDesc.SetFocus
    SAVEFLAG = True
    cmdCancel.Cancel = True
    'txtCode = "XXXXX" & GENICODE
End Sub

Private Sub cmdAdd_GotFocus()
    Msg cmdAdd.ToolTipText
End Sub

Private Sub cmdCancel_Click()
    Call btn_sts(True)
    Call ClsData(Me)
    cmdAdd.SetFocus
    cmdExit.Cancel = True
End Sub

Private Sub cmdCancel_GotFocus()
    Msg cmdCancel.ToolTipText
End Sub

Private Sub cmdDelete_Click()
    Dim ANS As String
    Dim TEMPRS As New ADODB.Recordset
    
    If M_USRSECLEVL = "1" Then
        If ReadConfigMaster("000010", 6, "M") = False Then ModuleDeniedMessage: Exit Sub
    End If
    
    If txtCode = Empty Then Call cmdEdit_Click
    
    If txtDesc.Text = "" Then
        MsgBox "There is no Record to delete.", vbCritical, App.Title
'        txtDesc.SetFocus
        Exit Sub
    End If
    
    If TEMPRS.State = adStateOpen Then TEMPRS.Close
    TEMPRS.Open "Select * from SPTRAN where ICOD ='" & txtCode.Text & "'", CN, adOpenDynamic, adLockOptimistic
    If TEMPRS.EOF Then
        If TEMPRS.State = adStateOpen Then TEMPRS.Close
        TEMPRS.Open "Select * from PURTRAN where ICOD ='" & txtCode.Text & "'", CN, adOpenDynamic, adLockOptimistic
    End If
    If TEMPRS.EOF Then
        If TEMPRS.State = adStateOpen Then TEMPRS.Close
        TEMPRS.Open "Select * from STORETRAN where ICOD ='" & txtCode.Text & "'", CN, adOpenDynamic, adLockOptimistic
    End If
        
MsgFlash:
        MsgBox "Can Not Delete This Record !!", vbCritical, App.Title
        TEMPRS.Close
        Exit Sub
    
    If TEMPRS.State = 1 Then TEMPRS.Close
''    TEMPRS.Open "Select * From ISS_MST Where V_ICOD='" & txtCode & "' AND TTYP<>'OPN'", CN, adOpenDynamic
''
''    If TEMPRS.EOF = False Then
''        MsgBox "Can Not Delete This Record !!", vbInformation, "Access Denied !!"
''        TEMPRS.Close
''        Exit Sub
''    End If
''
''    ans = MsgBox("Do you want to delete this record?", vbYesNo + vbQuestion, App.Title)
''    'Check MMS Data Structure So Items Not Deleted Accidently
''    If ans = vbYes Then
''        CN.Execute "Delete from STOCK_MST where ICOD ='" & Trim(txtCode.Text) & "' AND COMP='" & compPth & "'"
''        CN.Execute "Delete from ITMMST where CODE ='" & Trim(txtCode.Text) & "'"
''        CN.Execute "INSERT INTO DAILYSTAT (COMP,VTYP,SRNO,PCOD,DBCD,VBNO,QNTY,AMNT,CUSR,DTTM,ACTN) VALUES('" & compPth & "','ITM','XXXXXXXXXXXXX','" & txtDesc & "',NULL,'" & txtCode & "',0,0,'" & cUName & "','" & Format(Now, "MM/dd/yyyy HH:MM:SS AMPM") & "','D')"
''    End If
    ANS = MsgBox("Are You Sure To delete this record ? ", vbYesNo)
    If ANS = vbYes Then
        CN.Execute "Delete from ITMMST where CODE ='" & Trim(txtCode.Text) & "'"
       ' CN.Execute "INSERT INTO DAILYSTAT (COMP,VTYP,SRNO,PCOD,DBCD,VBNO,QNTY,AMNT,CUSR,DTTM,ACTN) VALUES('" & compPth & "','ITM','XXXXXXXXXXXXX','" & txtDesc & "',NULL,'" & txtCode & "',0,0,'" & cUName & "','" & Format(Now, "MM/dd/yyyy HH:MM:SS AMPM") & "','D')"
       '-----------------------------------------------
       'DAILYSTAT
       
       Call DAILYSTATUS("ITM", txtCode.Text, "", 0, "", 0, cUName, "D", Now, Now)
       '-----------------------------------------------
       
    End If
    Call btn_sts(True)
    Call ClsData(Me)
End Sub

Private Sub cmdDelete_GotFocus()
    Msg cmdDelete.ToolTipText
End Sub

Private Sub cmdEdit_Click()
    If M_USRSECLEVL = "1" Then
        If ReadConfigMaster("000010", 5, "M") = False Then ModuleDeniedMessage: Exit Sub
    End If
    
    SAVEFLAG = False
    'If Left(compNm, 4) = "SUBI" Then
    '    txtCode.Text = SearchList("Select Code,[NAME] from ITMMST")
    'Else
    NEW_VISIBLE = False
    txtCode = Empty
    M_DESC = Empty
    Key = Empty
    txtDesc = SearchITEMLIST("Select TOP 20 Code,[NAME] from ITMMST", 0, Empty, "Select Item")
    
    If Not txtDesc = Empty Then
        txtCode = Key
    'End If
        Call btn_sts(False)
        cmdDelete.Enabled = True
        txtCode.Enabled = False
        txtDesc.SetFocus
    Else
        Call cmdCancel_Click
    End If
End Sub

Private Sub cmdEdit_GotFocus()
    Msg cmdEdit.ToolTipText
End Sub

Private Sub cmdExit_Click()
    Msg Empty
    key_PressNew = False
    Unload Me
End Sub

Private Sub cmdExit_GotFocus()
    Msg cmdExit.ToolTipText
End Sub

Private Sub cmdList_Click()
    M_DESC = Empty
    Key = Empty
    txtCode.Text = SearchList("Select Code,[NAME] from ITMMST")
End Sub

Private Sub cmdSave_Click()
On Error GoTo errSaveClick
    Dim M_CHKRPL As String
    Dim strCHNS, strVALID As String, SQL As String, strAutoAprvdINDT As String
    Dim TEMPRS As New ADODB.Recordset
    Dim igcd As String
    Dim M_ITEX As String
    Dim Ctrl As Control
    
    If CHKRPL.Value = 1 Then
      M_CHKRPL = "Yes"
     Else
      M_CHKRPL = "No"
    End If
    For Each Ctrl In Me
        If TypeOf Ctrl Is TextBox Then
            Ctrl = Replace(Ctrl, "'", "", 1)
        End If
    Next
    
    If txtDesc.Text = "" Then
        MsgBox "Description can not be empty.", vbCritical, App.Title
        txtDesc.SetFocus
        Exit Sub
    End If
    
    If txtGroup.Text = "" Then
        MsgBox "Invalid Item Group....!!!", vbCritical, App.Title
        txtGroup.SetFocus
        Exit Sub
    End If
    
    If txtUNIT.Text = "" Then
        MsgBox "Unit Can not be empty.", vbCritical, App.Title
        If txtUNIT.Enabled = True Then
            txtUNIT.SetFocus
        End If
        Exit Sub
    End If
    
       
    If IsNumeric(TXTRATE.Text) = False And TXTRATE.Text <> "" Then
        MsgBox "Please Enter the Value.", vbInformation, App.Title
        TXTRATE.Text = ""
        TXTRATE.SetFocus
        Exit Sub
    End If
    
    If IsNumeric(txtLLIM.Text) = False And txtLLIM.Text <> "" Then
        MsgBox "Please Enter the Value.", vbInformation, App.Title
        txtLLIM.Text = ""
        txtLLIM.SetFocus
        Exit Sub
    End If
    
    If IsNumeric(txtULIM.Text) = False And txtULIM.Text <> "" Then
        MsgBox "Please Enter the Value.", vbInformation, App.Title
        txtULIM.Text = ""
        txtULIM.SetFocus
        Exit Sub
    End If
    
    If txtQTY.Text = "" Or UCase(txtQTY.Text) <> "Q" And UCase(txtQTY.Text) <> "P" And UCase(txtQTY.Text) <> "X" Then
        MsgBox "Please Enter Q / P / X.", vbInformation, App.Title
        txtQTY.SetFocus
        Exit Sub
    End If
    
    If TXTRATE.Text = "" Then
        TXTRATE.Text = 0
    End If
    
    If IsNumeric(txtBRAT.Text) = False And txtBRAT.Text <> "" Then
        MsgBox "Please Enter the Numeric Value..", vbCritical, App.Title
        txtBRAT.Text = ""
        txtBRAT.SetFocus
        Exit Sub
    End If
    
    If IsNumeric(txtERAT.Text) = False And txtERAT.Text <> "" Then
        MsgBox "Please Enter the Numeric Value..", vbCritical, App.Title
        txtERAT.Text = ""
        txtERAT.SetFocus
        Exit Sub
    End If
    
    If IsNumeric(txtSRAT.Text) = False And txtSRAT.Text <> "" Then
        MsgBox "Please Enter the Numeric Value..", vbCritical, App.Title
        txtSRAT.Text = ""
        txtSRAT.SetFocus
        Exit Sub
    End If
    
    If IsNumeric(Val(txtOPNQ.Text)) = False Then
        MsgBox "Invalid Quantity...!!!", vbCritical, App.Title
        txtOPNQ.SetFocus
        Exit Sub
    End If
    
    If IsNumeric(Val(txtOPNP.Text)) = False Then
        MsgBox "Invalid Pieces..!!!", vbCritical, App.Title
        txtOPNP.SetFocus
        Exit Sub
    End If
    
    If chkCHNS.Value = 1 Then strCHNS = "Y" Else strCHNS = "N"
    If chkVALID.Value = 1 Then strVALID = "Y" Else strVALID = "N"
    If chkAutoAprvdINDT.Value = 1 Then strAutoAprvdINDT = "Y" Else strAutoAprvdINDT = "N"
    
    If cmbLocation.ListIndex = -1 Then
        MsgBox "Location Can not be empty.", vbCritical, App.Title
        cmbLocation.SetFocus
        Exit Sub
    End If
    
    TEMPRS.Open "Select * from IGMMST where [NAME] = '" & Trim(txtGroup.Text) & "'", CN, adOpenDynamic, adLockOptimistic
    igcd = TEMPRS!CODE
    TEMPRS.Close
    
    TEMPRS.Open "Select * from LOCATION where LOCNAME = '" & Trim(cmbLocation.Text) & "'", CN, adOpenDynamic, adLockOptimistic
    cmbLocation.Tag = TEMPRS!LOCID
    TEMPRS.Close
    
On Error GoTo LAST
    If chkITEX.Value = 1 Then M_ITEX = "Y" Else M_ITEX = "N"
    
    TEMPRS.Open "Select * from ITMMST where [NAME] = '" & Trim(txtDesc.Text) & "'"
    If TEMPRS.EOF = False Then
        If txtCode = Empty Then
            MsgBox "Can not insert duplicate Record.", vbCritical, App.Title
            TEMPRS.Close
            Exit Sub
        End If
        TEMPRS.Close
    End If
    Dim ITM_DEF As String
    If OPTBEAM.Value = True Then
      ITM_DEF = "B"
    End If
    If OPTTAKA.Value = True Then
      ITM_DEF = "T"
    End If
    If OPTOTH.Value = True Then
      ITM_DEF = "O"
    End If
    If SAVEFLAG = True Then
        'txtCode.Text = GENICODE     'Generating New Item Code
        If TEMPRS.State = 1 Then TEMPRS.Close
        TEMPRS.Open "Select * from ITMMST where CODE = '" & Trim(txtCode.Text) & "'"
        If TEMPRS.EOF = False Then
            MsgBox "Can not insert duplicate Record.", vbCritical, App.Title
            txtCode.SetFocus
            TEMPRS.Close
            Exit Sub
        End If
        
        SQL = "Insert into ITMMST ([COMP],CODE,IGCD,[NAME],UNIT,LLIM,ULIM,RATE,WEIGHTEDRATE,QORP,DENI,GRAD," _
                & "QLTY,CHNS,VALID,BRAT,SRAT,ERAT,ITEX,OPNQ,OPNP,SALR,PURR,RLIM,AWIP,POSN,DELET,LPCOD,LPRAT,V_SCOD,AVGRATE,CLASS,LOCID,EXTRA1,EXTRA2,EXTRA3,EXTRA4) " _
                & " values('" & compPth & "','" & txtCode.Text & "','" & igcd & "','" & Trim(txtDesc.Text) & "','" & txtUNIT.Text & "','" & Val(txtLLIM.Text) & "','" & Val(txtULIM.Text) & "'," & TXTRATE.Text & "," & TXTRATE.Text & ",'" & UCase(txtQTY.Text) & _
                "','','','0','" & strCHNS & "','" & strVALID & "'," & Val(txtBRAT.Text) & "," & Val(txtSRAT.Text) & "," & Val(txtERAT.Text) & ",'" & M_ITEX & "'," & Val(txtOPNQ.Text) & "," & Val(txtOPNP.Text) & "," & Val(txtSalRate) & "," & Val(txtPurRate) & "," & Val(txtROrder) & ",'" & chkAWIP.Value & "','POSN','" & chkDeleted.Value & "','',0,'001'," & Val(TXTRATE) & ",'A','" & cmbLocation.Tag & "','" & TXTITM.Tag & "','" & txtAvgQty & "','" & txtGainQty & "','" & M_CHKRPL & "')"
        
        CN.BeginTrans
            CN.Execute SQL
            CN.Execute "UPDATE ITMMST SET PURQ = 0 WHERE [CODE] = '" & txtCode.Text & "'"
            CN.Execute "UPDATE ITMMST SET SALQ = 0 WHERE [CODE]= '" & txtCode.Text & "'"
            CN.Execute "UPDATE ITMMST SET SRAT='" & Val(txtSALPER) & "' WHERE [CODE]= '" & txtCode.Text & "'"
            CN.Execute "UPDATE ITMMST SET PURP = 0 WHERE [CODE] = '" & txtCode.Text & "'"
            CN.Execute "UPDATE ITMMST SET SALP = 0 WHERE [CODE]= '" & txtCode.Text & "'"
            CN.Execute "Update ITMMST set AUTOAPRVDINDT = '" & strAutoAprvdINDT & "' where CODE ='" & Trim(txtCode.Text) & "'"
            'CN.Execute "INSERT INTO DAILYSTAT (COMP,VTYP,SRNO,PCOD,DBCD,VBNO,QNTY,AMNT,CUSR,DTTM,ACTN) VALUES('" & compPth & "','ITM','XXXXXXXXXXXXX','" & txtDesc & "',NULL,'" & txtCode & "',0,0,'" & cUName & "','" & Format(Now, "MM/dd/yyyy HH:MM:SS AMPM") & "','N')"
            CN.Execute "UPDATE ITMMST SET EXTRA5='" & ITM_DEF & "' WHERE [CODE]= '" & txtCode.Text & "'"
            '------------------------
            'DAILYSTAT
            Call DAILYSTATUS("ITM", txtCode.Text, "", 0, "", 0, cUName, "N", Now, Now)
            '-----------------------
        CN.CommitTrans
    Else
        CN.BeginTrans
            CN.Execute "Update ITMMST set [NAME] = '" & Trim(txtDesc.Text) & "' where CODE ='" & Trim(txtCode.Text) & "'"
            CN.Execute "Update ITMMST set IGCD = '" & igcd & "' where CODE ='" & Trim(txtCode.Text) & "'"
            CN.Execute "Update ITMMST set UNIT = '" & Trim(txtUNIT.Text) & "' where CODE ='" & Trim(txtCode.Text) & "'"
            CN.Execute "Update ITMMST set RATE = " & Trim(TXTRATE.Text) & ",WEIGHTEDRATE = " & Trim(TXTRATE.Text) & " where CODE ='" & Trim(txtCode.Text) & "'"
            CN.Execute "Update ITMMST set RLIM = '" & Trim(txtROrder.Text) & "' where CODE ='" & Trim(txtCode.Text) & "'"
            CN.Execute "Update ITMMST set LLIM = '" & Trim(txtLLIM.Text) & "' where CODE ='" & Trim(txtCode.Text) & "'"
            CN.Execute "Update ITMMST set ULIM = '" & Trim(txtULIM.Text) & "' where CODE ='" & Trim(txtCode.Text) & "'"
            CN.Execute "Update ITMMST set QORP = '" & Trim(txtQTY.Text) & "' where CODE ='" & Trim(txtCode.Text) & "'"
            CN.Execute "Update ITMMST set CHNS = '" & strCHNS & "' where CODE ='" & Trim(txtCode.Text) & "'"
            CN.Execute "Update ITMMST set VALID = '" & strVALID & "' where CODE ='" & Trim(txtCode.Text) & "'"
            CN.Execute "Update ITMMST set BRAT = " & Val(txtBRAT.Text) & " where CODE ='" & Trim(txtCode.Text) & "'"
            CN.Execute "Update ITMMST set SRAT = " & Val(txtSRAT.Text) & " where CODE ='" & Trim(txtCode.Text) & "'"
            CN.Execute "Update ITMMST set ERAT = " & Val(txtERAT.Text) & " where CODE ='" & Trim(txtCode.Text) & "'"
            CN.Execute "Update ITMMST set ITEX = '" & M_ITEX & "' where CODE ='" & Trim(txtCode.Text) & "'"
            CN.Execute "Update ITMMST set SALR= " & Val(txtSalRate) & " where CODE ='" & Trim(txtCode.Text) & "'"
            CN.Execute "Update ITMMST set PURR= " & Val(txtPurRate) & " where CODE ='" & Trim(txtCode.Text) & "'"
            CN.Execute "Update ITMMST set OPNQ = " & Val(txtOPNQ.Text) & " where CODE ='" & Trim(txtCode.Text) & "'"
            CN.Execute "Update ITMMST set OPNP = " & Val(txtOPNP.Text) & " where CODE ='" & Trim(txtCode.Text) & "'"
            CN.Execute "Update ITMMST set AWIP = " & chkAWIP.Value & " where CODE ='" & Trim(txtCode.Text) & "'"
            CN.Execute "Update ITMMST set DELET = " & chkDeleted.Value & " where CODE ='" & Trim(txtCode.Text) & "'"
            CN.Execute "UPDATE ITMMST SET SRAT='" & Val(txtSALPER) & "' WHERE [CODE]= '" & txtCode.Text & "'"
            CN.Execute "Update ITMMST set AUTOAPRVDINDT = '" & strAutoAprvdINDT & "' where CODE ='" & Trim(txtCode.Text) & "'"
            CN.Execute "Update ITMMST set LOCID = '" & cmbLocation.Tag & "' where CODE ='" & Trim(txtCode.Text) & "'"
            CN.Execute "Update ITMMST set EXTRA1 = '" & TXTITM.Tag & "',EXTRA2 = '" & txtAvgQty & "',EXTRA3 = '" & txtGainQty & "' where CODE ='" & Trim(txtCode.Text) & "'"
            CN.Execute "Update ITMMST set EXTRA4 = '" & M_CHKRPL & "' where CODE ='" & Trim(txtCode.Text) & "'"
            CN.Execute "UPDATE ITMMST SET EXTRA5='" & ITM_DEF & "' WHERE [CODE]= '" & txtCode.Text & "'"
            'CN.Execute "INSERT INTO DAILYSTAT (COMP,VTYP,SRNO,PCOD,DBCD,VBNO,QNTY,AMNT,CUSR,DTTM,ACTN) VALUES('" & compPth & "','ITM','XXXXXXXXXXXXX','" & txtDesc & "',NULL,'" & txtCode & "',0,0,'" & cUName & "','" & Format(Now, "MM/dd/yyyy HH:MM:SS AMPM") & "','M')"
            'CN.Execute "IF EXISTS (SELECT * FROM STOCK_MST WHERE ICOD='" & txtCode & "') UPDATE STOCK_MST SET BAL_QTY=BAL_QTY - " & Val(txtOPNQ.Tag) & " + " & Val(txtOPNQ) & " WHERE ICOD='" & txtCode & "' AND COMP='" & compPth & "'"
            'CN.Execute "IF EXISTS (SELECT * FROM ISS_MST WHERE V_ICOD='" & txtCode & "' AND TTYP='OPN') UPDATE ISS_MST SET QNTY=" & Val(txtOPNQ) & " WHERE V_ICOD='" & txtCode & "' AND TTYP='OPN' AND COMP='" & compPth & "'"
            'CN.Execute "IF NOT EXISTS (SELECT * FROM STOCK_MST WHERE ICOD='" & txtCode & "') INSERT INTO STOCK_MST (COMP,ICOD,V_SCOD,BINCD,BAL_QTY,BAL_TEMP) VALUES('" & compPth & "','" & txtCode & "','001','XXXX'," & Val(txtOPNQ) & ",0)"
            'CN.Execute "IF NOT EXISTS (SELECT * FROM ISS_MST WHERE V_ICOD='" & txtCode & "' AND TTYP='OPN') INSERT INTO ISS_MST(COMP,CODE,TTYP,V_ICOD,BINCD,RATE,QNTY,V_SCOD,ISSMODE,D_ISSDATE,C_USERCODE,D_DATE,SRCH) VALUES('" & compPth & "','XXXXXXXXXXXX','OPN','" & txtCode & "','XXXX'," & Val(txtRate) & "," & Val(txtOPNQ) & ",'001',9,'" & Format(DateAdd("D", -1, FSDT), "MM/dd/yyyy") & "','" & cUName & "','" & Format(Date, "MM/dd/yyyy") & "',0)"
            '----------------------------------
            'DAILYSTAT
            Call DAILYSTATUS("ITM", txtCode.Text, "", 0, "", 0, cUName, "M", Now, Now)
            '----------------------------------
        CN.CommitTrans
    End If
    
    sTxt = txtDesc.Text
    Call btn_sts(True)
    Call ClsData(Me)
    
    cmdAdd.SetFocus
    iGrpFlag = False
    cmdExit.Cancel = True
    Exit Sub
    
LAST:

errSaveClick:
    If InStr(1, Err.Description, "more transaction", vbTextCompare) > 0 Then
        CN.RollbackTrans
        'Resume
    ElseIf Err.Number = -2147217873 Then
        MsgBox "Item Name Already Exists....", vbInformation, App.Title
        Exit Sub
    Else
        LOAD frm_ErrorHandler
        ErrNumber = Err.Number
        ErrMessage = Err.Description
        frm_ErrorHandler.Show vbModal
        Err.Clear
    End If
End Sub

Private Sub cmdSave_GotFocus()
    Msg CMDSAVE.ToolTipText
End Sub

Private Sub Form_Activate()
Call ColorComponent(Me)
Me.BackColor = RGB(RED, GREEN, BLUE)
If cmbLocation.ListCount > 1 Then
   cmbLocation.ListIndex = 0
End If
End Sub

Private Sub Form_Load()
Call ColorComponent(Me)
On Error GoTo errLoad
        
    If M_COMPBILL = "VFL" Or M_COMPBILL = "VF1" Then
        Label5.Visible = True
        txtSALPER.Visible = True
        Label5.Caption = "Wt.Per Box"
    Else
        Label5.Visible = False
        txtSALPER.Visible = False
    End If
    
    Call FillCmb("Select RTRIM(Name) From REFMST Where Cata='U'", cboUnit)
    Call FillCmb("SELECT LOCNAME FROM LOCATION ORDER BY LOCNAME", cmbLocation)

    cboUnit.AddItem "(New)"
    txtDesc.Enabled = False
    iGrpFlag = False
    key_PressNew = False

    Call btn_sts(True)
        
    Call CenterChild(frm_Main, Me)
    Exit Sub
    
errLoad:
    ErrNumber = Err.Number
    ErrMessage = Err.Description
    frm_ErrorHandler.Show vbModal
End Sub

Private Sub btn_sts(Yes As Boolean)
    txtCode.Enabled = Not Yes
    txtDesc.Enabled = Not Yes
    txtGroup.Enabled = Not Yes
    txtUNIT.Enabled = False
    txtLLIM.Enabled = Not Yes
    txtULIM.Enabled = Not Yes
    txtOPNQ.Enabled = Not Yes
    txtOPNP.Enabled = Not Yes
    txtSALPER.Enabled = Not Yes
    txtQTY.Enabled = Not Yes
    TXTRATE.Enabled = Not Yes
    chkAWIP.Enabled = Not Yes
    chkDeleted.Enabled = Not Yes
    cboUnit.Enabled = Not Yes
    
    
    CMDSAVE.Enabled = Not Yes
    cmdCancel.Enabled = Not Yes
    cmdAdd.Enabled = Yes
    cmdEdit.Enabled = Yes
    cmdDelete.Enabled = Yes
    framRate.Enabled = Not Yes
    chkCHNS.Enabled = Not Yes
    chkVALID.Enabled = Not Yes
    chkAutoAprvdINDT.Enabled = Not Yes
    txtSalRate.Enabled = Not Yes
    txtPurRate.Enabled = Not Yes
    txtROrder.Enabled = Not Yes
    cmbLocation.Enabled = Not Yes
    OPTBEAM.Enabled = Not Yes
    OPTTAKA.Enabled = Not Yes
    OPTOTH.Enabled = Not Yes
End Sub

Private Sub OPTBEAM_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub OPTOTH_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub OPTTAKA_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtCode_GotFocus()
 txtCode.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cboUnit.SetFocus: Call txtCode_Validate(False)
End Sub

Private Sub txtCode_LostFocus()
 txtCode.BackColor = vbWhite
End Sub

Private Sub txtCode_Validate(Cancel As Boolean)
    If SAVEFLAG = False Then Exit Sub
    If RS.State = 1 Then RS.Close
    RS.Open "Select * from ITMMST where CODE = '" & Trim(txtCode.Text) & "'"
    If RS.EOF = False Then
        MsgBox "Can not insert duplicate Record.", vbCritical, App.Title
        Cancel = True
        txtCode.SetFocus
        RS.Close
        Exit Sub
    End If
End Sub

Private Sub txtDesc_LostFocus()
 txtDesc.BackColor = vbWhite
End Sub

Private Sub txtGroup_LostFocus()
 txtGroup.BackColor = vbWhite
End Sub

Private Sub txtGroup_Validate(Cancel As Boolean)
    If SAVEFLAG = True Then txtCode.Text = GENICODE
End Sub

Private Sub txtLLIM_LostFocus()
txtLLIM.BackColor = vbWhite
End Sub

Private Sub txtPurRate_GotFocus()
    txtPurRate.BackColor = RGB(BRED, BGREEN, BBLUE)
    Msg "Enter Purchase Rate"
End Sub

Private Sub txtPurRate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cmbLocation.SetFocus
End If
End Sub

Private Sub txtPurRate_LostFocus()
    txtPurRate.BackColor = vbWhite
End Sub

Private Sub txtQLTY_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Me.framRate.Visible = True Then
            txtBRAT.SetFocus
        Else
            chkVALID.SetFocus
        End If
    End If
End Sub

Private Sub TXTQTY_LostFocus()
txtQTY.BackColor = vbWhite
End Sub

Private Sub TXTRATE_LostFocus()
   TXTRATE.BackColor = vbWhite
End Sub

Private Sub txtROrder_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtLLIM.SetFocus
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub txtROrder_LostFocus()
 txtROrder.BackColor = vbWhite
End Sub

Private Sub txtSALPER_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtSalRate_GotFocus()
  txtSalRate.BackColor = RGB(BRED, BGREEN, BBLUE)
    Msg "Enter Sales Rate"
End Sub

Private Sub txtROrder_GotFocus()
    txtROrder.BackColor = RGB(BRED, BGREEN, BBLUE)
    Msg "Enter Re Order Level Quantity"
End Sub

Private Sub txtBRAT_GotFocus()
    Msg "Enter Item Basic Rate"
End Sub

Private Sub txtBRAT_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtERAT.SetFocus
End Sub

Private Sub txtCode_Change()
On Error GoTo ERRUOM

    Dim TEMPRS As New ADODB.Recordset
    Dim Grp As New ADODB.Recordset
    
    If txtCode.Text = "" Then Exit Sub
    
    If SAVEFLAG = False Then
        TEMPRS.Open "Select * from ITMMST where Code='" & Trim(txtCode.Text) & "'", CN, adOpenDynamic, adLockOptimistic
        txtDesc.Text = Trim(TEMPRS![NAME])
        txtUNIT.Text = TEMPRS!unit & ""
        If Not IsNull(TEMPRS!unit) Then
            If Not TEMPRS!unit = "" Then
                cboUnit.Text = Trim(TEMPRS!unit & "")
            End If
        End If
        txtLLIM.Text = TEMPRS!LLIM
        txtULIM.Text = TEMPRS!ULIM
        If IsNull(TEMPRS!EXTRA5) = True Then
          OPTOTH.Value = True
          OPTBEAM.Value = False
          OPTTAKA.Value = False
         Else
          If TEMPRS!EXTRA5 = "O" Then
            OPTOTH.Value = True
            OPTBEAM.Value = False
            OPTTAKA.Value = False
          End If
          If TEMPRS!EXTRA5 = "B" Then
            OPTOTH.Value = False
            OPTBEAM.Value = True
            OPTTAKA.Value = False
          End If
          If TEMPRS!EXTRA5 = "T" Then
            OPTOTH.Value = False
            OPTBEAM.Value = False
            OPTTAKA.Value = True
          End If
        End If
        
                        
        txtSALPER.Text = TEMPRS!SRAT
        If IsNull(TEMPRS!EXTRA4) = True Or TEMPRS!EXTRA4 = "No" Then
          CHKRPL.Value = 0
         Else
          CHKRPL.Value = 1
        End If
        
        
        TXTRATE.Text = TEMPRS!RATE
        If IsNull(TEMPRS!QORP) = True Then
            txtQTY.Text = "Q"
        Else
            txtQTY.Text = TEMPRS!QORP
        End If
        txtROrder = TEMPRS!rlim
        txtSalRate = TEMPRS!SALR
        txtPurRate = TEMPRS!PURR
        If IsNull(TEMPRS!extra1) = False Then chkPlanning.Value = 1: TXTITM.Text = GetCode("ITMMST", Trim(TEMPRS!extra1), "CODE", "NAME"): txtAvgQty.Text = Trim(TEMPRS!EXTRA2): txtGainQty.Text = Val(Trim(TEMPRS!EXTRA3))
        If IsNull(TEMPRS!BRAT) = True Then txtBRAT.Text = 0 Else txtBRAT.Text = TEMPRS!BRAT
        If IsNull(TEMPRS!SRAT) = True Then txtSRAT.Text = 0 Else txtSRAT.Text = TEMPRS!SRAT
        If IsNull(TEMPRS!ERAT) = True Then txtERAT.Text = 0 Else txtERAT.Text = TEMPRS!ERAT
        If IsNull(TEMPRS![OPNQ]) = True Then txtOPNQ.Text = 0 Else txtOPNQ.Text = TEMPRS![OPNQ]
        If IsNull(TEMPRS![OPNQ]) = True Then txtOPNQ.Tag = 0 Else txtOPNQ.Tag = TEMPRS![OPNQ]
        If IsNull(TEMPRS![OPNP]) = True Then txtOPNP.Text = 0 Else txtOPNP.Text = TEMPRS![OPNP]
        If TEMPRS!ITEX = "Y" Then chkITEX.Value = 1 Else chkITEX.Value = 0
        If Trim(TEMPRS!CHNS) = "Y" Then chkCHNS.Value = 1 Else chkCHNS.Value = 0
        If Trim(TEMPRS!VALID) = "Y" Then chkVALID.Value = 1 Else chkVALID.Value = 0
        If Trim(TEMPRS!AUTOAPRVDINDT) = "Y" Then chkAutoAprvdINDT.Value = 1 Else chkAutoAprvdINDT.Value = 0
        
        If IsNull(TEMPRS!AWIP) Then chkAWIP.Value = 0 Else chkAWIP.Value = TEMPRS!AWIP
        If IsNull(TEMPRS!DELET) Then chkDeleted.Value = 0 Else chkDeleted.Value = TEMPRS!DELET
        
        Grp.Open "Select * from IGMMST where CODE ='" & Trim(TEMPRS!igcd) & "'", CN, adOpenDynamic, adLockOptimistic
        txtGroup.Text = Grp![NAME]
        Grp.Close
        
        Grp.Open "Select * from LOCATION where LOCID ='" & Trim(TEMPRS!LOCID) & "'", CN, adOpenDynamic, adLockOptimistic
        If Grp.EOF = False Then
            cmbLocation.Text = Grp![LOCNAME]
        End If
        Grp.Close
    End If
    
    Exit Sub
    
ERRUOM:
    'MsgBox Err.Description
    Resume Next
End Sub

Private Sub txtDesc_GotFocus()
 txtDesc.BackColor = RGB(BRED, BGREEN, BBLUE)
    txtDesc.SelStart = 0
    txtDesc.SelLength = Len(txtDesc)
    Msg "Enter Item Name"
    If M_COMPBILL = "SHB" Then
        If txtCode.Text = "XXXXX00002" Then
            txtDesc.Locked = False
        Else
            txtDesc.Locked = True
        End If
    End If
End Sub

Private Sub txtDesc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtGroup.SetFocus
    Select Case KeyAscii
        Case 34, 39
            KeyAscii = 0
    End Select
End Sub

Private Sub txtERAT_GotFocus()
    Msg "Enter Excise Rate"
End Sub

Private Sub txtERAT_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtSRAT.SetFocus
End Sub

Private Sub txtGrade_GotFocus()
    Msg "Enter Item Grade"
End Sub

Private Sub txtGroup_GotFocus()
    txtGroup.BackColor = RGB(BRED, BGREEN, BBLUE)
    Msg "Press <F2> To Get List Of Item Group"
    
    If key_PressNew = False And txtGroup.Text = "" And Not cmdAdd.Enabled Then
        txtGroup_KeyDown vbKeyF2, 16
    End If
End Sub

Private Sub txtGroup_KeyDown(KeyCode As Integer, Shift As Integer)
    M_DESC = Empty
    Key = Empty
    sTxt = ""
    If KeyCode = vbKeyF2 Then
        NEW_VISIBLE = True
        txtGroup.Text = SearchList1("select TOP 20 code, name from IGMMST", 0, "", "List Of Item Group")
        M_IGCD = Key
    End If
    If key_PressNew = True Then
        M_DESC = ""
        FRM_IGRP.Show
    End If
End Sub

Private Sub txtGroup_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtGroup.Text = "" Then
            txtGroup_KeyDown vbKeyF2, 16
        Else
            Call txtGroup_Validate(False)
'            If txtCode.Enabled = True Then txtCode.SetFocus
            SendKeys "{TAB}"
        End If
    End If
End Sub

Private Sub txtLLIM_GotFocus()
    txtLLIM.BackColor = RGB(BRED, BGREEN, BBLUE)
    Msg "Enter Item Lower Limit"
End Sub

Private Sub txtLLIM_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtULIM.SetFocus
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub txtOPNP_GotFocus()
    Msg "Enter Opening Pices"
End Sub

Private Sub txtOPNP_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtSalRate.SetFocus
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub txtOPNQ_GotFocus()
    Msg "Enter Opening Quantity"
End Sub

Private Sub txtOPNQ_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtOPNP.SetFocus
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub txtQLTY_GotFocus()
    Msg "Enter Item Quality"
End Sub

Private Sub txtQty_Change()
    If CMDSAVE.Enabled = False Then Exit Sub
    If txtOPNQ.Visible = True Then
        If Len(txtQTY.Text) = 1 Then txtOPNQ.SetFocus
    End If
End Sub

Private Sub TXTQTY_GotFocus()
    txtQTY.BackColor = RGB(BRED, BGREEN, BBLUE)
    Msg "Enter Q => Quantity / P => Pieces"
End Sub

Private Sub txtQTY_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
    If Not (UCase(Chr(KeyAscii)) = "Q" Or UCase(Chr(KeyAscii)) = "P") And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub TXTRATE_GotFocus()
    TXTRATE.BackColor = RGB(BRED, BGREEN, BBLUE)
    Msg "Enter Item Rate"
End Sub

Private Sub txtRate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtQTY.SetFocus
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 46 Then KeyAscii = 0
End Sub

Private Sub txtSalRate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtPurRate.SetFocus
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 46 Then KeyAscii = 0
End Sub

Private Sub txtSalRate_LostFocus()
 txtSalRate.BackColor = vbWhite
End Sub

Private Sub txtSRAT_GotFocus()
    Msg "Enter Sales Tax Rate"
End Sub

Private Sub txtSRAT_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then chkITEX.SetFocus
End Sub

Private Sub txtULIM_GotFocus()
    txtULIM.BackColor = RGB(BRED, BGREEN, BBLUE)
    Msg "Item Lower Limit"
End Sub

Private Sub txtULIM_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then TXTRATE.SetFocus
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub txtULIM_LostFocus()
  txtULIM.BackColor = vbWhite
End Sub

Private Sub txtUNIT_GotFocus()
    Msg "Enter Item Measurement Unit"
End Sub

Private Sub TXTUNIT_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtROrder.SetFocus
End Sub

Private Function GENICODE() As String
    Dim MS_ICOD As String
    Dim Prfx As String
    
    Set RS = New Recordset
    
    If Mid(M_COMPBILL, 1, 3) = "CIL" Then
      Prfx = GetCODEGEN(txtGroup) & GetSUBIHCD(txtGroup)
     Else
      Prfx = GetSUBIHCD(txtGroup) & GetCODEGEN(txtGroup)
    End If
    If Mid(M_COMPBILL, 1, 3) = "CIL" Then
      RS.Open "Select IsNull(Max(CODE),0) AS CODE From ITMMST Where CODE LIKE '" & Prfx & "0%' AND LEN(CODE)=10", CN, adOpenDynamic
     Else
      RS.Open "Select IsNull(Max(CODE),0) AS CODE From ITMMST Where CODE LIKE '" & Prfx & "0%'", CN, adOpenDynamic
    End If
    If Trim(RS!CODE) = "0" Then
        MS_ICOD = Prfx & "00001"
    Else
        MS_ICOD = Val(Mid(RS!CODE, 6)) + 1
        
        If MS_ICOD < 10 Then
            MS_ICOD = Prfx & "0000" & MS_ICOD
        ElseIf MS_ICOD < 100 Then
            MS_ICOD = Prfx & "000" & MS_ICOD
        ElseIf MS_ICOD < 1000 Then
            MS_ICOD = Prfx & "00" & MS_ICOD
        ElseIf MS_ICOD < 10000 Then
            MS_ICOD = Prfx & "0" & MS_ICOD
        Else
            MS_ICOD = Prfx & MS_ICOD
        End If
    End If
    
    RS.Close
    
    GENICODE = MS_ICOD
    'If M_COMPBILL = "KRAN" Then txtName = Mid(MS_ICOD, 4) & " - " & Mid(Trim(txtName), InStr(1, txtName, "-") + 1)
End Function

Function GetSUBIGCD(Grp As String)
    Dim rsList As Recordset
    Set rsList = New Recordset
    rsList.Open "Select IHCD AS IGCD FROM IGMMST WHERE NAME='" & txtGroup & "'", CN
    'IHCD IS REPLACE WITH COLUMN IGCD - IN MMS ITS IGCD
    If rsList.EOF = False Then GetSUBIGCD = Trim(rsList!igcd)
    rsList.Close
End Function

Function GetCODEGEN(Grp As String)
    Dim rsList As Recordset
    Set rsList = New Recordset
    rsList.Open "Select ISNULL(EXTRA1,'0000') AS IGCD FROM IGMMST WHERE NAME='" & txtGroup & "'", CN
    'IHCD IS REPLACE WITH COLUMN IGCD - IN MMS ITS IGCD
    If rsList.EOF = False Then GetCODEGEN = Trim(rsList!igcd)
    rsList.Close
End Function

Function GetSUBIHCD(Grp As String)
    Dim rsList As Recordset
    Set rsList = New Recordset
    rsList.Open "Select ISNULL(EXTRA1,0) AS IHCD FROM SCAT_MST WHERE CODE =(SELECT IHCD FROM IGMMST WHERE NAME='" & txtGroup & "')", CN
    'IHCD IS REPLACE WITH COLUMN IGCD - IN MMS ITS IGCD
    If rsList.EOF = False Then GetSUBIHCD = Trim(rsList!ihcd)
    rsList.Close
End Function

Private Sub TXTITM_KeyDown(KeyCode As Integer, Shift As Integer)
    Me.KeyPreview = False
    If KeyCode = vbKeyF2 Or (Trim(TXTITM) = Empty And KeyCode = vbKeyReturn) Then
        NEW_VISIBLE = False
        Key = Empty
        
        SQL = "select TOP 20 ITMMST.code,ITMMST.NAME from itmmst "
        SQL = SQL & "inner join igmmst on itmmst.igcd=igmmst.code "
        SQL = SQL & "inner join scat_mst on igmmst.ihcd=scat_mst.code "
        SQL = SQL & "where SCAT_MST.extra2='RM'"
        
        TXTITM.Text = SearchITEMLIST(SQL, 0, TXTITM.Text, "SELECT RAW ITEM FROM LIST")
        TXTITM.ToolTipText = TXTITM
        TXTITM.Tag = Key
    Else
        If KeyCode = vbKeyReturn Then txtAvgQty.SetFocus
    End If
           
    Me.KeyPreview = True
End Sub


Private Sub txtAvgQty_KeyPress(KeyAscii As Integer)
        If InStr(1, txtAvgQty.Text, ".") > 0 And KeyAscii = 46 Then
            KeyAscii = 0
            Exit Sub
        End If

        If KeyAscii = 13 Then txtGainQty.SetFocus
        
        If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
            If (KeyAscii <> 46) Then KeyAscii = 0
        End If
End Sub

Private Sub txtGainQty_KeyPress(KeyAscii As Integer)
        If InStr(1, txtGainQty.Text, ".") > 0 And KeyAscii = 46 Then
            KeyAscii = 0
            Exit Sub
        End If

        If KeyAscii = 13 Then CMDSAVE.SetFocus
        
        If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
            If (KeyAscii <> 46) Then KeyAscii = 0
        End If
End Sub

