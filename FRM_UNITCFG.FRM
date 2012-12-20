VERSION 5.00
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WELCHBUTTON.OCX"
Begin VB.Form FRM_UNITCFG 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Unit Configuration"
   ClientHeight    =   7485
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9555
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7485
   ScaleWidth      =   9555
   Begin VB.TextBox TXTIMPORT 
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
      Left            =   6240
      MaxLength       =   250
      TabIndex        =   97
      Top             =   6840
      Width           =   3255
   End
   Begin VB.TextBox TXTEXPORT 
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
      Left            =   6240
      MaxLength       =   250
      TabIndex        =   95
      Top             =   6240
      Width           =   3255
   End
   Begin VB.TextBox TXTWEXCO 
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
      Left            =   6240
      MaxLength       =   50
      TabIndex        =   91
      Top             =   2760
      Width           =   3255
   End
   Begin VB.TextBox TXTWCHAP 
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
      Left            =   6240
      MaxLength       =   15
      TabIndex        =   90
      Top             =   3360
      Width           =   3255
   End
   Begin VB.TextBox PERCINCAP 
      Alignment       =   1  'Right Justify
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
      Left            =   8520
      MaxLength       =   6
      TabIndex        =   88
      Text            =   "0"
      Top             =   2160
      Width           =   855
   End
   Begin VB.Frame Frame7 
      Height          =   1935
      Left            =   6240
      TabIndex        =   79
      Tag             =   "0"
      Top             =   3960
      Width           =   2535
      Begin VB.CheckBox ChkSal2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "2. Excise Sale"
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
         TabIndex        =   86
         Top             =   360
         Width           =   2355
      End
      Begin VB.CheckBox ChkSal1 
         BackColor       =   &H00C0C0FF&
         Caption         =   "1. Commercial Tax"
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
         TabIndex        =   85
         Top             =   120
         Width           =   2355
      End
      Begin VB.CheckBox ChkSal3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "3. Export Sale"
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
         TabIndex        =   84
         Top             =   600
         Width           =   2355
      End
      Begin VB.CheckBox ChkSal4 
         BackColor       =   &H00FFFFC0&
         Caption         =   "4. Direct Sale"
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
         TabIndex        =   83
         Top             =   840
         Width           =   2355
      End
      Begin VB.CheckBox ChkSal5 
         BackColor       =   &H00FFC0C0&
         Caption         =   "5. Commercial Retail"
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
         TabIndex        =   82
         Top             =   1080
         Width           =   2355
      End
      Begin VB.CheckBox ChkSal6 
         BackColor       =   &H00FFC0FF&
         Caption         =   "6. Trading Sale (Tax)"
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
         TabIndex        =   81
         Top             =   1320
         Width           =   2355
      End
      Begin VB.CheckBox ChkSal7 
         BackColor       =   &H00C0C0C0&
         Caption         =   "7. Trading Sale (Retail)"
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
         TabIndex        =   80
         Top             =   1560
         Width           =   2355
      End
   End
   Begin VB.Frame Frame6 
      ForeColor       =   &H000040C0&
      Height          =   735
      Left            =   120
      TabIndex        =   77
      Top             =   2160
      Width           =   6015
      Begin VB.TextBox EX_TR6 
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
         MaxLength       =   6
         TabIndex        =   15
         Text            =   "000000"
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox EX_STAX 
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
         Left            =   3600
         MaxLength       =   6
         TabIndex        =   19
         Text            =   "000000"
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox EX_EXC 
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
         Left            =   2520
         MaxLength       =   6
         TabIndex        =   17
         Text            =   "000000"
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label25 
         Caption         =   "Last Excise No."
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   120
         TabIndex        =   78
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label Label24 
         Caption         =   "Service tax"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   3600
         TabIndex        =   18
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label23 
         Caption         =   "Excise"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   2520
         TabIndex        =   16
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label22 
         Caption         =   "TR-6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   1560
         TabIndex        =   14
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.PictureBox picToolTip 
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      FillColor       =   &H00404040&
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   120
      ScaleHeight     =   255
      ScaleWidth      =   4815
      TabIndex        =   75
      Top             =   7560
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
         TabIndex        =   76
         Top             =   0
         Width           =   120
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Others  "
      ForeColor       =   &H000040C0&
      Height          =   735
      Left            =   120
      TabIndex        =   74
      Top             =   6720
      Width           =   2655
      Begin VB.TextBox OTH_GATEPASS 
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
         Left            =   600
         MaxLength       =   6
         TabIndex        =   64
         Text            =   "000000"
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox OT_RETURNABLE 
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
         MaxLength       =   6
         TabIndex        =   66
         Text            =   "000000"
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label21 
         Caption         =   "GatePass "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   600
         TabIndex        =   63
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label20 
         Caption         =   "Returnable"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   1560
         TabIndex        =   65
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Last Invoice No."
      ForeColor       =   &H000040C0&
      Height          =   855
      Left            =   120
      TabIndex        =   73
      Top             =   5880
      Width           =   6015
      Begin VB.TextBox SAL_TRD_RET 
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
         Left            =   5160
         MaxLength       =   6
         TabIndex        =   62
         Text            =   "000000"
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox SAL_TRD_TAX 
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
         Left            =   4320
         MaxLength       =   6
         TabIndex        =   60
         Text            =   "000000"
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox SAL_EXPORT 
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
         Left            =   960
         MaxLength       =   6
         TabIndex        =   52
         Text            =   "000000"
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox SAL_DIRECT 
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
         MaxLength       =   6
         TabIndex        =   54
         Text            =   "000000"
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox SAL_EXCISE 
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
         Left            =   120
         MaxLength       =   6
         TabIndex        =   50
         Text            =   "000000"
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox SAL_RETAIL 
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
         Left            =   2640
         MaxLength       =   6
         TabIndex        =   56
         Text            =   "000000"
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox SAL_TAX 
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
         MaxLength       =   6
         TabIndex        =   58
         Text            =   "000000"
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label28 
         Caption         =   "Trd.Retail"
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   5160
         TabIndex        =   61
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label27 
         Caption         =   "Trd.Tax"
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   4320
         TabIndex        =   59
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label19 
         Caption         =   "Excise"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label18 
         Caption         =   "Export"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   1080
         TabIndex        =   51
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label17 
         Caption         =   "Direct"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   1800
         TabIndex        =   53
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label16 
         Caption         =   "Retail"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   2640
         TabIndex        =   55
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label11 
         Caption         =   "Tax"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   3480
         TabIndex        =   57
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Last Dispatch No."
      ForeColor       =   &H000040C0&
      Height          =   855
      Left            =   120
      TabIndex        =   72
      Top             =   4920
      Width           =   6015
      Begin VB.TextBox DPF_RETAIL 
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
         Left            =   5040
         MaxLength       =   6
         TabIndex        =   48
         Text            =   "000000"
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox DPF_WASTAGE 
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
         Left            =   4080
         MaxLength       =   6
         TabIndex        =   47
         Text            =   "000000"
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox DPF_EXPORT 
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
         Left            =   3120
         MaxLength       =   6
         TabIndex        =   45
         Text            =   "000000"
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox DPF_MARKET 
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
         Left            =   240
         MaxLength       =   6
         TabIndex        =   39
         Text            =   "000000"
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox DPF_CAPTIVE 
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
         Left            =   2160
         MaxLength       =   6
         TabIndex        =   43
         Text            =   "000000"
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox DPF_JOB 
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
         Left            =   1200
         MaxLength       =   6
         TabIndex        =   41
         Text            =   "000000"
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label29 
         Caption         =   "Retail"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   5040
         TabIndex        =   94
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label10 
         Caption         =   "Wastage"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   4080
         TabIndex        =   46
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label15 
         Caption         =   "Export"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   3120
         TabIndex        =   44
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label14 
         Caption         =   "Captive"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   2160
         TabIndex        =   42
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Caption         =   "Job"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   1200
         TabIndex        =   40
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label12 
         Caption         =   "Market/Tax"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Last Issue No. "
      ForeColor       =   &H000040C0&
      Height          =   855
      Left            =   120
      TabIndex        =   71
      Top             =   3960
      Width           =   6015
      Begin VB.TextBox ISS_RETURN 
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
         Left            =   4560
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   37
         Text            =   "000000"
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox ISS_JOB 
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
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   31
         Text            =   "000000"
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox ISS_RGP 
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
         Left            =   2520
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   33
         Text            =   "000000"
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox ISS_DIVISION 
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
         Left            =   600
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   29
         Text            =   "000000"
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox ISS_NRGP 
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
         Left            =   3600
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   35
         Text            =   "000000"
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Return"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   4560
         TabIndex        =   36
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Division"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   600
         TabIndex        =   28
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "JOB"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   1560
         TabIndex        =   30
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "RGP"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   2520
         TabIndex        =   32
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "NRGP"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   3600
         TabIndex        =   34
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Last GRN No. "
      ForeColor       =   &H000040C0&
      Height          =   855
      Left            =   120
      TabIndex        =   70
      Top             =   3000
      Width           =   6015
      Begin VB.TextBox GRN_SERVICES 
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
         Left            =   3600
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   27
         Text            =   "000000"
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox GRN_RAW 
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
         Left            =   600
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   21
         Text            =   "000000"
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox GRN_CONTRACT 
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
         Left            =   2520
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   25
         Text            =   "000000"
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox GRN_INHOUSE 
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
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   23
         Text            =   "000000"
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Services"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   3600
         TabIndex        =   26
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Contract"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   2520
         TabIndex        =   24
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Inhouse"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   1560
         TabIndex        =   22
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Raw"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   600
         TabIndex        =   20
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame framCfg1 
      Height          =   2175
      Left            =   120
      TabIndex        =   69
      Tag             =   "0"
      Top             =   0
      Width           =   9375
      Begin VB.CheckBox CHK_15 
         BackColor       =   &H00C0FFFF&
         Caption         =   "14.LR Detail Required on Challan?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   5880
         TabIndex        =   13
         Top             =   1560
         Width           =   3375
      End
      Begin VB.CheckBox CHK_13 
         BackColor       =   &H00C0C0FF&
         Caption         =   "12.Online Challan Print Required?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   5880
         TabIndex        =   11
         Top             =   1080
         Width           =   3255
      End
      Begin VB.CheckBox CHK_14 
         BackColor       =   &H00FFC0C0&
         Caption         =   "13.Online Bill Print Required?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   5880
         TabIndex        =   12
         Top             =   1320
         Width           =   3255
      End
      Begin VB.CheckBox CHK_12 
         BackColor       =   &H00E0E0E0&
         Caption         =   "11.Box Wise Inventory Required ??"
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
         Left            =   5880
         TabIndex        =   10
         Top             =   840
         Width           =   3435
      End
      Begin VB.CheckBox CHK_11 
         BackColor       =   &H00004080&
         Caption         =   "10. Division wise Challan Series ??"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   5880
         TabIndex        =   9
         Top             =   600
         Width           =   3435
      End
      Begin VB.CheckBox CHK_10 
         BackColor       =   &H00FFFF00&
         Caption         =   "9.Item Value Round off Required ??"
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
         Left            =   5880
         TabIndex        =   8
         Top             =   360
         Width           =   3435
      End
      Begin VB.CheckBox CHK_9 
         BackColor       =   &H000000FF&
         Caption         =   "8. Is It Exciseable Unit ??"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000004&
         Height          =   255
         Left            =   5880
         TabIndex        =   7
         Top             =   120
         Width           =   3195
      End
      Begin VB.CheckBox CHK_8 
         BackColor       =   &H00FFFF00&
         Caption         =   "7. VAT,CST Post Req.?"
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
         TabIndex        =   6
         Top             =   1560
         Width           =   5595
      End
      Begin VB.CheckBox CHK_7 
         BackColor       =   &H00C0C0C0&
         Caption         =   "6. Export Packing Required ??"
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
         TabIndex        =   5
         Top             =   1320
         Width           =   5595
      End
      Begin VB.CheckBox CHK_6 
         BackColor       =   &H00FFC0FF&
         Caption         =   "5. Is MMS Installed ??"
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
         Top             =   1080
         Width           =   5595
      End
      Begin VB.CheckBox CHK_5 
         BackColor       =   &H00FFC0C0&
         Caption         =   "4. Rate Master required for Sale Order Booking ??"
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
         TabIndex        =   3
         Top             =   840
         Width           =   5595
      End
      Begin VB.CheckBox CHK_4 
         BackColor       =   &H00FFFFC0&
         Caption         =   "3. Is Current Unit Treat As Company [For Doc. Printing] ??"
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
         Top             =   600
         Width           =   5595
      End
      Begin VB.CheckBox CHK_3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "2. Order Booking required for Dispatch [Delivery Challan] ??"
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
         TabIndex        =   1
         Top             =   360
         Width           =   5595
      End
      Begin VB.CheckBox CHK_2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "1. Challan Transfer required from one Unit to another Unit ??"
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
         TabIndex        =   0
         Top             =   120
         Width           =   5595
      End
   End
   Begin WelchButton.lvButtons_H cmdSave 
      Height          =   495
      Left            =   3120
      TabIndex        =   67
      Top             =   6840
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
      Image           =   "FRM_UNITCFG.frx":0000
      cBack           =   -2147483633
   End
   Begin WelchButton.lvButtons_H cmdExit 
      Height          =   495
      Left            =   4440
      TabIndex        =   68
      Top             =   6840
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
      Image           =   "FRM_UNITCFG.frx":059A
      cBack           =   -2147483633
   End
   Begin VB.Label Label31 
      Caption         =   "Import File Path of Server(Local Drive):"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   6240
      TabIndex        =   98
      Top             =   6600
      Width           =   3495
   End
   Begin VB.Label Label30 
      Caption         =   "Export Path:Server(Local):Map(Node)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   6240
      TabIndex        =   96
      Top             =   6000
      Width           =   3495
   End
   Begin VB.Label LBLWCHAP 
      Caption         =   "Chapter No. of Wastage"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   6240
      TabIndex        =   93
      Top             =   3120
      Width           =   2535
   End
   Begin VB.Label LBLWEXCO 
      Caption         =   "C.Ex Commodity of Wastage    "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   6240
      TabIndex        =   92
      Top             =   2520
      Width           =   2535
   End
   Begin VB.Label Label36 
      BackColor       =   &H00800080&
      Caption         =   "%age In Capital Goods"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   6240
      TabIndex        =   89
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Label Label26 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "Select Type of Sale"
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
      Height          =   255
      Left            =   6240
      TabIndex        =   87
      Top             =   3720
      Width           =   2535
   End
End
Attribute VB_Name = "FRM_UNITCFG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ISFORMLOAD As Boolean

Private Sub CHK_11_Click()
If Not ISFORMLOAD Then Exit Sub

If CHK_11.Value = 1 Then
   Dim STR
   STR = MsgBox("ARE YOU SURE : Division wise series option One Time Only. !!", vbYesNo + vbQuestion, "Not Reversal ?")
   If STR = VBNO Then
      CHK_11.Value = 0
      Frame3.Enabled = True
   End If
Else
   Frame3.Enabled = False
End If


End Sub

Private Sub CHK_9_Click()
  If CHK_9.Value = 1 Then
     LBLWEXCO.Enabled = True
     TXTWEXCO.Enabled = True
     LBLWCHAP.Enabled = True
     TXTWCHAP.Enabled = True
    Else
     LBLWEXCO.Enabled = False
     TXTWEXCO.Enabled = False
     LBLWCHAP.Enabled = False
     TXTWCHAP.Enabled = False
  End If
End Sub

Private Sub CMDEXIT_Click()
   Unload Me
End Sub

Private Sub cmdSave_Click()
On Error GoTo CFGERR

For Each Ctrl In Me
  If TypeOf Ctrl Is TextBox And UCase(Ctrl.NAME) <> "TXTEXPORT" And UCase(Ctrl.NAME) <> "TXTIMPORT" And Ctrl.NAME <> "PERCINCAP" And Ctrl.NAME <> "TXTWEXCO" And Ctrl.NAME <> "TXTWCHAP" Then
     If (Ctrl.NAME = "GRN_RAW" Or Ctrl.NAME = "ISS_DIVISION") And CHK_6.Value = 1 Then
       'OK
     Else
        If Len(Ctrl) <> 6 Then MsgBox "Enter 6 Digit Serial.": Ctrl.SetFocus: Exit Sub
     End If
  End If
Next

If TXTIMPORT = Empty Then
   TXTIMPORT = "C:\"
End If

If TXTEXPORT = Empty Then
   TXTEXPORT = "C:\"
End If

CN.BeginTrans
CN.Execute "UPDATE UNTCFG SET EXPPKG_REQ = '" & IIf(CHK_7.Value = 1, "Y", "N") & _
   "',TRANSFER_REQ = '" & IIf(CHK_2.Value = 1, "Y", "N") & "',ORDER_REQ = '" & IIf(CHK_3.Value = 1, "Y", "N") & _
   "',CUNT = '" & IIf(CHK_4.Value = 1, "Y", "N") & "',RATEMST_REQ = '" & IIf(CHK_5.Value = 1, "Y", "N") & _
   "',MMS_INSTALL = '" & IIf(CHK_6.Value = 1, "Y", "N") & "',POSTAC = '" & IIf(CHK_8.Value = 1, "Y", "N") & _
   "',EXCUNIT = '" & IIf(CHK_9.Value = 1, "Y", "N") & "',EXCCPERC = '" & Val(PERCINCAP.Text) & _
   "',WCHAP='" & Trim(TXTWCHAP) & "',WEXCO='" & Trim(TXTWEXCO) & "',ITEMRO = '" & IIf(CHK_10.Value = 1, "Y", "N") & _
   "',DIVSERIES = '" & IIf(CHK_11.Value = 1, "Y", "N") & _
   "', BOXREQ = '" & IIf(CHK_12.Value = 1, "Y", "N") & "',ONLINECHLN = '" & IIf(CHK_13.Value = 1, "Y", "N") & _
   "', ONLINEBILL = '" & IIf(CHK_14.Value = 1, "Y", "N") & "',LRONCHLN = '" & IIf(CHK_15.Value = 1, "Y", "N") & _
   "' WHERE COMP='" & compPth & _
   "' AND UNIT='" & UNCD & "'"

Call SetData

CN.Execute "UPDATE UNTCFG SET SCANIMP='" & TXTIMPORT & "',SCANEXP='" & TXTEXPORT & _
           "' WHERE COMP='" & compPth & "'AND UNIT='" & UNCD & "'"

CN.CommitTrans
        
        UNT_EXPPKG_REQ = IIf(CHK_7.Value = 1, "Y", "N")
        UNT_MMS_INSTALL = IIf(CHK_6.Value = 1, "Y", "N")
        UNT_DIVSERIES_REQ = IIf(CHK_11.Value = 1, "Y", "N")
        UNT_TRANSFER_REQ = IIf(CHK_2.Value = 1, "Y", "N")
                
        IsOnlineChallanPrintReq = IIf(CHK_13.Value = 1, True, False)
        IsOnlineBillPrintReq = IIf(CHK_14.Value = 1, True, False)
        
        'division wise challan series
        If UNT_DIVSERIES_REQ = "Y" Then
           frm_Main.MNUsETUPOP(6).Visible = True
        Else
           frm_Main.MNUsETUPOP(6).Visible = False
        End If
        '----------------------------------------
               
        'CHALLAN TRANSFER
        If UNT_TRANSFER_REQ = "Y" Then
           frm_Main.mnuTransOp(17).Visible = True
        Else
           frm_Main.mnuTransOp(17).Visible = False
        End If
        '----------------------------------------
        
        If IIf(CHK_9.Value = 1, "Y", "N") = "Y" Then
           frm_Main.mnuMasterGrp1(5).Visible = True
           frm_Main.mnuExcise(0).Visible = True
        Else
           frm_Main.mnuMasterGrp1(5).Visible = False
           frm_Main.mnuExcise(0).Visible = False
        End If
        
        
       'CHALLAN TRANSFER
        If UNT_EXPPKG_REQ = "Y" Then
            frm_Main.mnuCarton(2).Visible = True
            frm_Main.mnuCarton(3).Visible = True
            frm_Main.mnuCarton(4).Visible = True
            frm_Main.mnuCarton(5).Visible = False
            'frm_Main.mnuSale(3).Visible = True
            frm_Main.mnuPacking(0).Visible = True
        Else
            frm_Main.mnuCarton(2).Visible = False
            frm_Main.mnuCarton(3).Visible = False
            frm_Main.mnuCarton(4).Visible = False
            frm_Main.mnuCarton(5).Visible = False
           frm_Main.mnuPacking(0).Visible = False
        End If
          
        'MAIN FORM
        Dim i As Long
        Dim VIS_FLAG As Boolean: VIS_FLAG = True
        If UNT_MMS_INSTALL = "Y" Then VIS_FLAG = False
        For i = 7 To 16
          frm_Main.mnuOrderBooking(i).Visible = VIS_FLAG
        Next i
        '======================================================
        
Unload Me

Exit Sub
CFGERR:
MsgBox ERR.Description
Resume
CN.RollbackTrans
End Sub

Private Sub DPF_RETAIL_GotFocus()
  DPF_RETAIL.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub DPF_RETAIL_LostFocus()
  DPF_RETAIL.BackColor = vbWhite
End Sub

Private Sub Form_Activate()
  LOAD frm_CMPCFGSecurity
  frm_CMPCFGSecurity.Show 1
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If TypeOf ActiveControl Is TextBox Then
   If Len(Trim(ActiveControl.Text)) <> 6 Then Exit Sub
End If
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_Load()

On Error Resume Next
Dim CFGRS As New ADODB.Recordset
Set CFGRS = New ADODB.Recordset

Call CenterChild(frm_Main, Me)
ISFORMLOAD = False
Call GetData

If CFGRS.State = 1 Then CFGRS.Close
CFGRS.Open "SELECT * FROM UNTCFG WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "'", CN, adOpenDynamic, adLockOptimistic
If Not CFGRS.EOF Then
    
   TXTIMPORT = Trim(CFGRS!SCANIMP & "")
   TXTEXPORT = Trim(CFGRS!SCANEXP & "")
   
   If Trim(CFGRS!ONLINECHLN) = "Y" Then CHK_13.Value = 1
   If Trim(CFGRS!ONLINEBILL) = "Y" Then CHK_14.Value = 1
   If Trim(CFGRS!LRONCHLN) = "Y" Then CHK_15.Value = 1
    
   If Trim(CFGRS!TRANSFER_REQ) = "Y" Then CHK_2.Value = 1
   If Trim(CFGRS!ORDER_REQ) = "Y" Then CHK_3.Value = 1
   If Trim(CFGRS!CUNT) = "Y" Then CHK_4.Value = 1
   If Trim(CFGRS!RATEMST_REQ) = "Y" Then CHK_5.Value = 1
   If Trim(CFGRS!MMS_INSTALL) = "Y" Then CHK_6.Value = 1: CHK_6.Enabled = False
   If Trim(CFGRS!EXPPKG_REQ) = "Y" Then CHK_7.Value = 1
   If Trim(CFGRS!POSTAC) = "Y" Then CHK_8.Value = 1
   If Trim(CFGRS!EXCUNIT) = "Y" Then CHK_9.Value = 1
   If Trim(CFGRS!EXCUNIT) = "Y" Then CHK_9.Value = 1
   If Trim(CFGRS!ITEMRO) = "Y" Then CHK_10.Value = 1
   If Trim(CFGRS!BOXREQ) = "Y" Then CHK_12.Value = 1
   If Trim(CFGRS!DIVSERIES) = "Y" Then CHK_11.Value = 1: CHK_11.Enabled = False: Frame3.Enabled = False
   PERCINCAP.Text = Val(CFGRS!exccperc & "")
   TXTWEXCO = Trim(CFGRS!WEXCO & "")
   TXTWCHAP = Trim(CFGRS!WCHAP & "")
End If

For Each Ctrl In Me
        If TypeOf Ctrl Is Frame Then
            Ctrl.BackColor = RGB(RED, GREEN, BLUE)
            Ctrl.ForeColor = RGB(LBLRED, LBLGREEN, LBLBLUE)
        ElseIf TypeOf Ctrl Is Label Then
            Ctrl.BackColor = RGB(RED, GREEN, BLUE)
            Ctrl.ForeColor = RGB(LBLRED, LBLGREEN, LBLBLUE)
        End If
        Me.BackColor = RGB(RED, GREEN, BLUE)
Next

ISFORMLOAD = True

End Sub

Private Sub GRN_RAW_GotFocus()
    GRN_RAW.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub GRN_RAW_KeyPress(KeyAscii As Integer)
   Call CheckNumericKey(KeyAscii, GRN_RAW, Me, False)
End Sub

Private Sub GRN_RAW_LostFocus():  GRN_RAW.BackColor = vbWhite: picToolTip.Visible = False: End Sub

Private Sub GRN_CONTRACT_GotFocus()
  GRN_CONTRACT.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub GRN_CONTRACT_LostFocus():  GRN_CONTRACT.BackColor = vbWhite: picToolTip.Visible = False: End Sub

Private Sub GRN_INHOUSE_GotFocus()
   
    GRN_INHOUSE.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub GRN_INHOUSE_LostFocus():  GRN_INHOUSE.BackColor = vbWhite: picToolTip.Visible = False: End Sub

Private Sub GRN_SERVICES_GotFocus()
  
   GRN_SERVICES.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub GRN_SERVICES_LostFocus():  GRN_SERVICES.BackColor = vbWhite: picToolTip.Visible = False: End Sub

Private Sub ISS_DIVISION_GotFocus()
   
    ISS_DIVISION.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub ISS_DIVISION_LostFocus():  ISS_DIVISION.BackColor = vbWhite: End Sub

Private Sub ISS_JOB_GotFocus()
ToolTip Me, "6 Digit Serial Must.", "", 500 + 1150, 3800
ISS_JOB.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub
Private Sub ISS_JOB_LostFocus():  ISS_JOB.BackColor = vbWhite: picToolTip.Visible = False: End Sub

Private Sub ISS_NRGP_GotFocus()
  
   ISS_NRGP.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub ISS_NRGP_LostFocus():  ISS_NRGP.BackColor = vbWhite: End Sub

Private Sub ISS_RGP_GotFocus()
   
    ISS_RGP.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub ISS_RGP_LostFocus():  ISS_RGP.BackColor = vbWhite: picToolTip.Visible = False: End Sub

Private Sub ISS_RETURN_GotFocus()
 
  ISS_RETURN.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub ISS_RETURN_LostFocus():  ISS_RETURN.BackColor = vbWhite: picToolTip.Visible = False: End Sub

Private Sub DPF_CAPTIVE_GotFocus()
    
     DPF_CAPTIVE.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub DPF_CAPTIVE_LostFocus():  DPF_CAPTIVE.BackColor = vbWhite: picToolTip.Visible = False: End Sub

Private Sub DPF_EXPORT_GotFocus()
     
      DPF_EXPORT.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub DPF_EXPORT_LostFocus():  DPF_EXPORT.BackColor = vbWhite: picToolTip.Visible = False: End Sub

Private Sub DPF_JOB_GotFocus()
  
   DPF_JOB.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub DPF_JOB_LostFocus():  DPF_JOB.BackColor = vbWhite: picToolTip.Visible = False: End Sub

Private Sub DPF_MARKET_GotFocus()
     
      DPF_MARKET.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub DPF_MARKET_LostFocus():  DPF_MARKET.BackColor = vbWhite: picToolTip.Visible = False: End Sub

Private Sub DPF_WASTAGE_GotFocus()
  
   DPF_WASTAGE.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub DPF_WASTAGE_LostFocus():  DPF_WASTAGE.BackColor = vbWhite: picToolTip.Visible = False: End Sub



Private Sub percincap_GotFocus()
  PERCINCAP.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub percincap_KeyPress(KeyAscii As Integer)
   If CheckNumericKey(KeyAscii, PERCINCAP, Me, True) = 0 Then KeyAscii = 0
End Sub

Private Sub percincap_LostFocus()
  PERCINCAP.BackColor = vbWhite
End Sub

Private Sub SAL_DIRECT_GotFocus()
  
   SAL_DIRECT.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub SAL_DIRECT_LostFocus():  SAL_DIRECT.BackColor = vbWhite: picToolTip.Visible = False: End Sub

Private Sub SAL_EXCISE_GotFocus()
  
   SAL_EXCISE.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub SAL_EXCISE_LostFocus():  SAL_EXCISE.BackColor = vbWhite: picToolTip.Visible = False: End Sub

Private Sub SAL_EXPORT_GotFocus()
  
   SAL_EXPORT.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub SAL_EXPORT_LostFocus():  SAL_EXPORT.BackColor = vbWhite: picToolTip.Visible = False: End Sub

Private Sub SAL_RETAIL_GotFocus()
      
       SAL_RETAIL.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub SAL_RETAIL_LostFocus():  SAL_RETAIL.BackColor = vbWhite: picToolTip.Visible = False: End Sub

Private Sub SAL_tax_GotFocus()
  
   SAL_TAX.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub SAL_tax_LostFocus():  SAL_TAX.BackColor = vbWhite: picToolTip.Visible = False: End Sub

Private Sub OT_RETURNABLE_GotFocus()
   
    OT_RETURNABLE.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub OT_RETURNABLE_LostFocus():  OT_RETURNABLE.BackColor = vbWhite: picToolTip.Visible = False: End Sub

Private Sub OTH_GATEPASS_GotFocus()
  
  
   OTH_GATEPASS.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub OTH_GATEPASS_LostFocus():  OTH_GATEPASS.BackColor = vbWhite: picToolTip.Visible = False: End Sub

Private Sub EX_TR6_GotFocus()
   EX_TR6.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub EX_TR6_LostFocus():  EX_TR6.BackColor = vbWhite: picToolTip.Visible = False: End Sub

Private Sub EX_EXC_GotFocus()
  
   EX_EXC.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub EX_EXC_LostFocus():  EX_EXC.BackColor = vbWhite: picToolTip.Visible = False: End Sub

Private Sub EX_STAX_GotFocus()
   EX_STAX.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub EX_STAX_LostFocus():  EX_STAX.BackColor = vbWhite: picToolTip.Visible = False: End Sub

Private Sub GetData()
Dim GETRS As ADODB.Recordset
Set GETRS = New ADODB.Recordset

If GETRS.State = 1 Then GETRS.Close
GETRS.Open "SELECT CODE,SRNO AS LVNO,VTYP,ACTIVE FROM SERIALMASTER WHERE COMP='" & compPth & _
"' AND UNIT='" & UNCD & "' AND FYCD='" & FYCD & "' AND (DVCD IS NULL OR DVCD='')", CN, adOpenDynamic, adLockOptimistic
If GETRS.EOF Then Exit Sub

Do While Not GETRS.EOF
Select Case UCase(Trim(GETRS!VTYP))
Case "IVR"
'GRN-----------------------
    Select Case UCase(Trim(GETRS!CODE))
    Case "000001"
         GRN_RAW = Mid(UCase(Trim(GETRS!LVNO)), 1, 6)
    Case "000002"
         GRN_INHOUSE = Mid(UCase(Trim(GETRS!LVNO)), 1, 6)
    Case "000003"
         GRN_CONTRACT = Mid(UCase(Trim(GETRS!LVNO)), 1, 6)
    Case "000004"
         OT_RETURNABLE = Mid(UCase(Trim(GETRS!LVNO)), 1, 6)
    End Select
Case "PSR"
      If Trim(GETRS!CODE) = "000001" Then
        GRN_SERVICES = Mid(UCase(Trim(GETRS!LVNO)), 1, 6)
      End If
'--------------------------
Case "ISS"
    ISS_DIVISION = Mid(UCase(Trim(GETRS!LVNO)), 1, 6)
Case "ANX"
    ISS_JOB = Mid(UCase(Trim(GETRS!LVNO)), 1, 6)
Case "RGP"
    ISS_RGP = Mid(UCase(Trim(GETRS!LVNO)), 1, 6)
Case "NGP"
    ISS_NRGP = Mid(UCase(Trim(GETRS!LVNO)), 1, 6)
Case "RTI"
    ISS_RETURN = Mid(UCase(Trim(GETRS!LVNO)), 1, 6)

Case "DPF"
'DISPATCH-----------------------
    Select Case UCase(Trim(GETRS!CODE))
    Case "000001"
         DPF_MARKET = Mid(UCase(Trim(GETRS!LVNO)), 1, 6)
    Case "000002"
         DPF_EXPORT = Mid(UCase(Trim(GETRS!LVNO)), 1, 6)
    Case "000003"
         DPF_JOB = Mid(UCase(Trim(GETRS!LVNO)), 1, 6)
    Case "000004"
         DPF_CAPTIVE = Mid(UCase(Trim(GETRS!LVNO)), 1, 6)
    Case "000005"
         DPF_WASTAGE = Mid(UCase(Trim(GETRS!LVNO)), 1, 6)
    Case "000006"
         DPF_RETAIL = Mid(UCase(Trim(GETRS!LVNO)), 1, 6)
    End Select
'--------------------------------
Case "SAL"
'SALE-----------------------
    Select Case UCase(Trim(GETRS!CODE))
    Case "000001"
         SAL_TAX = Mid(UCase(Trim(GETRS!LVNO)), 1, 6)
         If Trim(GETRS!ACTIVE) = "Y" Then ChkSal1.Value = 1
    Case "000002"
         SAL_EXCISE = Mid(UCase(Trim(GETRS!LVNO)), 1, 6)
         If Trim(GETRS!ACTIVE) = "Y" Then ChkSal2.Value = 1
    Case "000003"
         SAL_EXPORT = Mid(UCase(Trim(GETRS!LVNO)), 1, 6)
         If Trim(GETRS!ACTIVE) = "Y" Then ChkSal3.Value = 1
    Case "000004"
         SAL_DIRECT = Mid(UCase(Trim(GETRS!LVNO)), 1, 6)
         If Trim(GETRS!ACTIVE) = "Y" Then ChkSal4.Value = 1
    Case "000005"
         SAL_RETAIL = Mid(UCase(Trim(GETRS!LVNO)), 1, 6)
         If Trim(GETRS!ACTIVE) = "Y" Then ChkSal5.Value = 1
    Case "000006"
         SAL_TRD_TAX = Mid(UCase(Trim(GETRS!LVNO)), 1, 6)
         If Trim(GETRS!ACTIVE) = "Y" Then ChkSal6.Value = 1
    Case "000007"
         SAL_TRD_RET = Mid(UCase(Trim(GETRS!LVNO)), 1, 6)
         If Trim(GETRS!ACTIVE) = "Y" Then ChkSal7.Value = 1
    End Select
'--------------------------------
Case "RET"
    OT_RETURNABLE = Mid(UCase(Trim(GETRS!LVNO)), 1, 6)
Case "GPS"
    OTH_GATEPASS = Mid(UCase(Trim(GETRS!LVNO)), 1, 6)
Case "TR6"
    EX_TR6 = Mid(UCase(Trim(GETRS!LVNO)), 1, 6)
Case "EXC"
    EX_EXC = Mid(UCase(Trim(GETRS!LVNO)), 1, 6)
Case "EXD"
    EX_STAX = Mid(UCase(Trim(GETRS!LVNO)), 1, 6)
End Select
GETRS.MoveNext
Loop
GETRS.Close
End Sub

Private Sub SetData()
On Error GoTo SETERR
Dim SETRS As ADODB.Recordset
Set SETRS = New ADODB.Recordset
Dim NO_OF_ROWS_EFFECT As Long

Dim FY As String
FY = FYCD

If CHK_6.Value = 0 Then
'IVR------SIZING
CN.Execute "UPDATE SERIALMASTER SET SRNO = '" & GRN_RAW & "' WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
           "' AND FYCD ='" & FYCD & "' AND VTYP='IVR' AND CODE='000001'", NO_OF_ROWS_EFFECT
           
If NO_OF_ROWS_EFFECT <> 1 Then MsgBox "CHECK DATA CONSISTENCY"
           
CN.Execute "UPDATE SERIALMASTER SET SRNO = '" & GRN_INHOUSE & "' WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
           "' AND FYCD ='" & FYCD & "' AND VTYP='IVR' AND CODE='000002'", NO_OF_ROWS_EFFECT
           
If NO_OF_ROWS_EFFECT <> 1 Then MsgBox "CHECK DATA CONSISTENCY"

CN.Execute "UPDATE SERIALMASTER SET SRNO = '" & GRN_CONTRACT & "' WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
           "' AND FYCD ='" & FYCD & "' AND VTYP='IVR' AND CODE='000003'", NO_OF_ROWS_EFFECT
           
If NO_OF_ROWS_EFFECT <> 1 Then MsgBox "CHECK DATA CONSISTENCY"

CN.Execute "UPDATE SERIALMASTER SET SRNO = '" & OT_RETURNABLE & "' WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
           "' AND FYCD ='" & FYCD & "' AND VTYP='IVR' AND CODE='000004'", NO_OF_ROWS_EFFECT
           
If NO_OF_ROWS_EFFECT <> 1 Then MsgBox "CHECK DATA CONSISTENCY"

'PSR
CN.Execute "UPDATE SERIALMASTER SET SRNO = '" & GRN_SERVICES & FY & "' WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
           "' AND FYCD ='" & FYCD & "' AND VTYP='PSR' AND CODE='000001'", NO_OF_ROWS_EFFECT
           
If NO_OF_ROWS_EFFECT <> 1 Then MsgBox "CHECK DATA CONSISTENCY"
           
'ISS
CN.Execute "UPDATE SERIALMASTER SET SRNO = '" & ISS_DIVISION & "' WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
           "' AND FYCD ='" & FYCD & "' AND VTYP='ISS' AND CODE='000001'", NO_OF_ROWS_EFFECT
           
If NO_OF_ROWS_EFFECT <> 1 Then MsgBox "CHECK DATA CONSISTENCY"

'ANX
CN.Execute "UPDATE SERIALMASTER SET SRNO = '" & ISS_JOB & FY & "' WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
           "' AND FYCD ='" & FYCD & "' AND VTYP='ANX' AND CODE='000001'", NO_OF_ROWS_EFFECT

If NO_OF_ROWS_EFFECT <> 1 Then MsgBox "CHECK DATA CONSISTENCY"

'RGP
CN.Execute "UPDATE SERIALMASTER SET SRNO = '" & ISS_RGP & FY & "' WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
           "' AND FYCD ='" & FYCD & "' AND VTYP='RGP' AND CODE='000001'", NO_OF_ROWS_EFFECT
           
If NO_OF_ROWS_EFFECT <> 1 Then MsgBox "CHECK DATA CONSISTENCY"
           
'NGP
CN.Execute "UPDATE SERIALMASTER SET SRNO = '" & ISS_NRGP & FY & "' WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
           "' AND FYCD ='" & FYCD & "' AND VTYP='NGP' AND CODE='000001'", NO_OF_ROWS_EFFECT
           
If NO_OF_ROWS_EFFECT <> 1 Then MsgBox "CHECK DATA CONSISTENCY"

'RTI
CN.Execute "UPDATE SERIALMASTER SET SRNO = '" & ISS_RETURN & FY & "' WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
           "' AND FYCD ='" & FYCD & "' AND VTYP='RTI' AND CODE='000001'", NO_OF_ROWS_EFFECT
           
If NO_OF_ROWS_EFFECT <> 1 Then MsgBox "CHECK DATA CONSISTENCY"

End If

'DPF
CN.Execute "UPDATE SERIALMASTER SET SRNO = '" & DPF_MARKET & FY & "' WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
           "' AND FYCD ='" & FYCD & "' AND VTYP='DPF' AND CODE='000001' AND (DVCD='' OR DVCD IS NULL)", NO_OF_ROWS_EFFECT
           
If NO_OF_ROWS_EFFECT <> 1 Then MsgBox "CHECK DATA CONSISTENCY"
           
CN.Execute "UPDATE SERIALMASTER SET SRNO = '" & DPF_EXPORT & FY & "' WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
           "' AND FYCD ='" & FYCD & "' AND VTYP='DPF' AND CODE='000002' AND (DVCD='' OR DVCD IS NULL)", NO_OF_ROWS_EFFECT
           
If NO_OF_ROWS_EFFECT <> 1 Then MsgBox "CHECK DATA CONSISTENCY"
           
CN.Execute "UPDATE SERIALMASTER SET SRNO = '" & DPF_JOB & FY & "' WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
           "' AND FYCD ='" & FYCD & "' AND VTYP='DPF' AND CODE='000003' AND (DVCD='' OR DVCD IS NULL)", NO_OF_ROWS_EFFECT
           
If NO_OF_ROWS_EFFECT <> 1 Then MsgBox "CHECK DATA CONSISTENCY"
           
CN.Execute "UPDATE SERIALMASTER SET SRNO = '" & DPF_CAPTIVE & FY & "' WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
           "' AND FYCD ='" & FYCD & "' AND VTYP='DPF' AND CODE='000004' AND (DVCD='' OR DVCD IS NULL)", NO_OF_ROWS_EFFECT
           
If NO_OF_ROWS_EFFECT <> 1 Then MsgBox "CHECK DATA CONSISTENCY"
           
CN.Execute "UPDATE SERIALMASTER SET SRNO = '" & DPF_WASTAGE & FY & "' WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
           "' AND FYCD ='" & FYCD & "' AND VTYP='DPF' AND CODE='000005' AND (DVCD='' OR DVCD IS NULL)", NO_OF_ROWS_EFFECT
           
If NO_OF_ROWS_EFFECT <> 1 Then MsgBox "CHECK DATA CONSISTENCY"

CN.Execute "UPDATE SERIALMASTER SET SRNO = '" & DPF_RETAIL & FY & "' WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
           "' AND FYCD ='" & FYCD & "' AND VTYP='DPF' AND CODE='000006' AND (DVCD='' OR DVCD IS NULL)", NO_OF_ROWS_EFFECT
           
If NO_OF_ROWS_EFFECT <> 1 Then MsgBox "CHECK DATA CONSISTENCY"


'SAL
CN.Execute "UPDATE SERIALMASTER SET SRNO = '" & SAL_TAX & FY & "',ACTIVE = '" & IIf(ChkSal1.Value = 1, "Y", "N") & "'  WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
           "' AND FYCD ='" & FYCD & "' AND VTYP='SAL' AND CODE='000001'", NO_OF_ROWS_EFFECT
           
If NO_OF_ROWS_EFFECT <> 1 Then MsgBox "CHECK DATA CONSISTENCY"
           
CN.Execute "UPDATE SERIALMASTER SET SRNO = '" & SAL_EXCISE & FY & "',ACTIVE = '" & IIf(ChkSal2.Value = 1, "Y", "N") & "'  WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
           "' AND FYCD ='" & FYCD & "' AND VTYP='SAL' AND CODE='000002'", NO_OF_ROWS_EFFECT
           
If NO_OF_ROWS_EFFECT <> 1 Then MsgBox "CHECK DATA CONSISTENCY"
           
CN.Execute "UPDATE SERIALMASTER SET SRNO = '" & SAL_EXPORT & FY & "',ACTIVE = '" & IIf(ChkSal3.Value = 1, "Y", "N") & "'  WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
           "' AND FYCD ='" & FYCD & "' AND VTYP='SAL' AND CODE='000003'", NO_OF_ROWS_EFFECT
           
If NO_OF_ROWS_EFFECT <> 1 Then MsgBox "CHECK DATA CONSISTENCY"
           
CN.Execute "UPDATE SERIALMASTER SET SRNO = '" & SAL_DIRECT & FY & "',ACTIVE = '" & IIf(ChkSal4.Value = 1, "Y", "N") & "'  WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
           "' AND FYCD ='" & FYCD & "' AND VTYP='SAL' AND CODE='000004'", NO_OF_ROWS_EFFECT
           
If NO_OF_ROWS_EFFECT <> 1 Then MsgBox "CHECK DATA CONSISTENCY"
           
CN.Execute "UPDATE SERIALMASTER SET SRNO = '" & SAL_RETAIL & FY & "',ACTIVE = '" & IIf(ChkSal5.Value = 1, "Y", "N") & "'  WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
           "' AND FYCD ='" & FYCD & "' AND VTYP='SAL' AND CODE='000005'", NO_OF_ROWS_EFFECT
           
If NO_OF_ROWS_EFFECT <> 1 Then MsgBox "CHECK DATA CONSISTENCY"
                
CN.Execute "UPDATE SERIALMASTER SET SRNO = '" & SAL_TRD_TAX & FY & "',ACTIVE = '" & IIf(ChkSal6.Value = 1, "Y", "N") & "'  WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
           "' AND FYCD ='" & FYCD & "' AND VTYP='SAL' AND CODE='000006'", NO_OF_ROWS_EFFECT
           
If NO_OF_ROWS_EFFECT <> 1 Then MsgBox "CHECK DATA CONSISTENCY"
           
CN.Execute "UPDATE SERIALMASTER SET SRNO = '" & SAL_TRD_RET & FY & "',ACTIVE = '" & IIf(ChkSal7.Value = 1, "Y", "N") & "'  WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
           "' AND FYCD ='" & FYCD & "' AND VTYP='SAL' AND CODE='000007'", NO_OF_ROWS_EFFECT
           
If NO_OF_ROWS_EFFECT <> 1 Then MsgBox "CHECK DATA CONSISTENCY"
           
'RET
CN.Execute "UPDATE SERIALMASTER SET SRNO = '" & OT_RETURNABLE & FY & "' WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
           "' AND FYCD ='" & FYCD & "' AND VTYP='RET' AND CODE='000001'", NO_OF_ROWS_EFFECT
           
If NO_OF_ROWS_EFFECT <> 1 Then MsgBox "CHECK DATA CONSISTENCY"

'GPS
CN.Execute "UPDATE SERIALMASTER SET SRNO = '" & OTH_GATEPASS & FY & "' WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
           "' AND FYCD ='" & FYCD & "' AND VTYP='GPS' AND CODE='000001'", NO_OF_ROWS_EFFECT
           
If NO_OF_ROWS_EFFECT <> 1 Then MsgBox "CHECK DATA CONSISTENCY"

'TR6
CN.Execute "UPDATE SERIALMASTER SET SRNO = '" & EX_TR6 & FY & "' WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
           "' AND FYCD ='" & FYCD & "' AND VTYP='TR6' AND CODE='000001'", NO_OF_ROWS_EFFECT
           
If NO_OF_ROWS_EFFECT <> 1 Then MsgBox "CHECK DATA CONSISTENCY"
           
'EXC
CN.Execute "UPDATE SERIALMASTER SET SRNO = '" & EX_EXC & FY & "' WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
           "' AND FYCD ='" & FYCD & "' AND VTYP='EXC' AND CODE='000001'", NO_OF_ROWS_EFFECT
           
If NO_OF_ROWS_EFFECT <> 1 Then MsgBox "CHECK DATA CONSISTENCY"

'EXD
CN.Execute "UPDATE SERIALMASTER SET SRNO = '" & EX_STAX & FY & "' WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
           "' AND FYCD ='" & FYCD & "' AND VTYP='EXD' AND CODE='000001'", NO_OF_ROWS_EFFECT
           
           
If NO_OF_ROWS_EFFECT <> 1 Then MsgBox "CHECK DATA CONSISTENCY"
                    
Exit Sub
SETERR:
MsgBox ERR.Description
CN.RollbackTrans
End Sub

Private Sub TXTWCHAP_GotFocus()
  TXTWCHAP.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTWCHAP_LostFocus()
  TXTWCHAP.BackColor = vbWhite
End Sub

Private Sub TXTWEXCO_GotFocus()
  TXTWEXCO.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTWEXCO_LostFocus()
  TXTWEXCO.BackColor = vbWhite
End Sub
