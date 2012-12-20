VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{2BF72F7D-D367-4712-9146-5533EF3E691A}#1.2#0"; "FraPlus1.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WELCHBUTTON.OCX"
Begin VB.Form frmGatePass 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gate Pass Entry Module"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11400
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   11400
   Begin FramePlusCtl.FramePlus FrmAutoConsumption 
      Height          =   6795
      Left            =   0
      TabIndex        =   29
      Top             =   0
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   11986
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
      Begin VB.Frame framTransDetail 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5175
         Left            =   2160
         TabIndex        =   41
         Top             =   600
         Visible         =   0   'False
         Width           =   6855
         Begin MSMask.MaskEdBox txtIN 
            Height          =   330
            Left            =   1560
            TabIndex        =   10
            Top             =   1680
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   582
            _Version        =   393216
            MaxLength       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin VB.TextBox TXTRMRK 
            Height          =   285
            Left            =   1560
            MaxLength       =   50
            TabIndex        =   23
            Top             =   3840
            Width           =   4935
         End
         Begin VB.TextBox TXTNOTRIP 
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
            Left            =   1560
            MaxLength       =   3
            TabIndex        =   21
            Top             =   3480
            Width           =   735
         End
         Begin VB.TextBox txtPaid 
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
            TabIndex        =   19
            Top             =   3120
            Width           =   975
         End
         Begin VB.TextBox txtLicenceNo 
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
            MaxLength       =   30
            TabIndex        =   17
            Top             =   2760
            Width           =   4935
         End
         Begin VB.TextBox txtDriver 
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
            MaxLength       =   30
            TabIndex        =   15
            Top             =   2400
            Width           =   4935
         End
         Begin MSComCtl2.DTPicker InDate 
            Height          =   315
            Left            =   2640
            TabIndex        =   11
            Top             =   1680
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
            Format          =   54001665
            CurrentDate     =   40474
         End
         Begin VB.TextBox txtTransport 
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
            TabIndex        =   8
            Top             =   1320
            Width           =   4935
         End
         Begin VB.TextBox txtVHCL 
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
            TabIndex        =   6
            Top             =   960
            Width           =   2295
         End
         Begin WelchButton.lvButtons_H cmdAdd 
            Height          =   495
            Left            =   1560
            TabIndex        =   24
            Top             =   4320
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   873
            Caption         =   "&O.k"
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
            Image           =   "frmGatePass.frx":0000
            cBack           =   -2147483633
         End
         Begin MSMask.MaskEdBox txtOUT 
            Height          =   330
            Left            =   1560
            TabIndex        =   13
            Top             =   2040
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   582
            _Version        =   393216
            MaxLength       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Transportation Details"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   1920
            TabIndex        =   42
            Top             =   240
            Width           =   3015
         End
         Begin VB.Shape Shape3 
            BorderColor     =   &H8000000D&
            BorderWidth     =   2
            Height          =   4335
            Left            =   120
            Shape           =   4  'Rounded Rectangle
            Top             =   720
            Width           =   6615
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "Remar&ks "
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
            Left            =   360
            TabIndex        =   22
            Tag             =   "S"
            Top             =   3840
            Width           =   975
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "No. of Tr&ips "
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
            Left            =   360
            TabIndex        =   20
            Tag             =   "S"
            Top             =   3480
            Width           =   1335
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "&Advance Paid"
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
            Left            =   360
            TabIndex        =   18
            Top             =   3135
            Width           =   1215
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "&Licence No."
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
            Left            =   360
            TabIndex        =   16
            Top             =   2760
            Width           =   1215
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "D&river Name"
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
            Left            =   360
            TabIndex        =   14
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "&Out Time"
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
            Left            =   360
            TabIndex        =   12
            Top             =   2055
            Width           =   1215
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "&In Time/Date"
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
            Left            =   360
            TabIndex        =   9
            Top             =   1695
            Width           =   1215
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "&Transport"
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
            Left            =   360
            TabIndex        =   7
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "&Vehicle No."
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
            Left            =   360
            TabIndex        =   5
            Top             =   960
            Width           =   1215
         End
      End
      Begin VB.TextBox TTLQTY 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   39
         Text            =   ".00"
         Top             =   6120
         Width           =   1335
      End
      Begin VB.TextBox TTLPCS 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   405
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   6120
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker TXTVBDT 
         Height          =   330
         Left            =   9360
         TabIndex        =   30
         Top             =   120
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   54001665
         CurrentDate     =   39347
      End
      Begin MSComctlLib.ListView lstChallan 
         Height          =   4335
         Left            =   120
         TabIndex        =   3
         Top             =   1320
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   7646
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   10485760
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   14
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ChallanNo."
            Object.Width           =   2294
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Date"
            Object.Width           =   1853
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Party Name"
            Object.Width           =   2648
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Consignee"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Address"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Item Name"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Grd"
            Object.Width           =   971
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "LotNo"
            Object.Width           =   2029
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Pcs"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Qnty"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "DBCD"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "RTYP"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "SDBC"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "SVBN"
            Object.Width           =   0
         EndProperty
      End
      Begin WelchButton.lvButtons_H cmdPrint 
         Height          =   495
         Left            =   8520
         TabIndex        =   26
         Top             =   6000
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   873
         Caption         =   "&Print"
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
         Image           =   "frmGatePass.frx":039A
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdDelete 
         Height          =   495
         Left            =   9600
         TabIndex        =   27
         Top             =   6000
         Visible         =   0   'False
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
         Image           =   "frmGatePass.frx":0734
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdSave 
         Height          =   495
         Left            =   5280
         TabIndex        =   4
         Top             =   6000
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
         Image           =   "frmGatePass.frx":0ACE
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdCancel 
         Height          =   495
         Left            =   6360
         TabIndex        =   25
         Top             =   6000
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
         Image           =   "frmGatePass.frx":1858
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdExit 
         Height          =   495
         Left            =   7440
         TabIndex        =   28
         Top             =   6000
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
         Image           =   "frmGatePass.frx":1CAA
         cBack           =   -2147483633
      End
      Begin MSComCtl2.DTPicker txtFromDt 
         Height          =   330
         Left            =   4440
         TabIndex        =   0
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   54001665
         CurrentDate     =   38837
      End
      Begin MSComCtl2.DTPicker txtToDt 
         Height          =   330
         Left            =   7320
         TabIndex        =   1
         Top             =   705
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   54001665
         CurrentDate     =   38837
      End
      Begin WelchButton.lvButtons_H cmdSearch 
         Height          =   495
         Left            =   9000
         TabIndex        =   2
         Top             =   640
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   873
         Caption         =   "Search"
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
         Image           =   "frmGatePass.frx":20FC
         cBack           =   -2147483633
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Qty."
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
         Left            =   1920
         TabIndex        =   40
         Tag             =   "S"
         Top             =   5880
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Pcs"
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
         Left            =   360
         TabIndex        =   38
         Tag             =   "S"
         Top             =   5880
         Width           =   855
      End
      Begin VB.Label lblfrom 
         BackStyle       =   0  'Transparent
         Caption         =   "Delivery Challan From Date :"
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
         Left            =   1440
         TabIndex        =   36
         Top             =   720
         Width           =   2895
      End
      Begin VB.Label lblTo 
         BackStyle       =   0  'Transparent
         Caption         =   "To Date :"
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
         Left            =   6240
         TabIndex        =   35
         Top             =   720
         Width           =   1095
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         Height          =   615
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   600
         Width           =   11175
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H80000002&
         BorderStyle     =   5  'Dash-Dot-Dot
         DrawMode        =   1  'Blackness
         FillColor       =   &H00000080&
         Height          =   705
         Left            =   4920
         Shape           =   4  'Rounded Rectangle
         Top             =   5880
         Width           =   6375
      End
      Begin VB.Shape BORDER3 
         BorderColor     =   &H80000002&
         Height          =   300
         Left            =   4200
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   3135
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
         TabIndex        =   34
         Tag             =   "S"
         Top             =   -2040
         Width           =   1335
      End
      Begin VB.Label LBLCHDT 
         BackStyle       =   0  'Transparent
         Caption         =   "Date :"
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
         Left            =   8640
         TabIndex        =   33
         Top             =   120
         Width           =   615
      End
      Begin VB.Label LBLCHLN 
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
         Left            =   5760
         TabIndex        =   32
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label LBLGATE 
         BackStyle       =   0  'Transparent
         Caption         =   "Gate Pass No."
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
         Left            =   4320
         TabIndex        =   31
         Tag             =   "0"
         Top             =   120
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmGatePass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SAVEFLAG As Boolean
Dim INDEX As Long
Dim DIVCODE As String
Dim DIVNAME As String
Dim M_DBCD As String
Dim PKG_SCOD As String
Dim BOX_PKG_REQ As String
Dim FICD As String, MCCD As String, LOCCOD As String
Public CHALLAN As String

Private Sub cmdAdd_Click()
  If Trim(txtVHCL) = Empty Then
     txtVHCL.SetFocus
  End If
  Call cmdSave_Click
End Sub

Private Sub cmdCancel_Click()
txtVHCL = Empty: txtTransport = Empty: TXTNOTRIP = Empty: TXTRMRK = Empty
InDate = Now: txtDriver = Empty: txtLicenceNo = Empty
framTransDetail.Visible = False
txtPaid = Empty:
lstChallan.ListItems.Clear
txtFromDt.Enabled = True
txtFromDt.SetFocus
TTLPCS = Empty
TTLQTY = ".00"
LBLCHLN.Caption = GenVNO("GPS", "000001")
End Sub

Private Sub CMDEXIT_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()

If framTransDetail.Visible = False Then
    framTransDetail.Visible = True
    txtVHCL.SetFocus
    Exit Sub
Else
   If txtVHCL = Empty Then txtVHCL.SetFocus: Exit Sub
   framTransDetail.Visible = False
End If

On Error GoTo LAST

If Not CHKSAVEDATA Then Exit Sub

Dim SAVERS As ADODB.Recordset
Set SAVERS = New ADODB.Recordset
Dim M_VHCD As String, M_TRCD As String

'TRANSPORT CODE
If SAVERS.State = 1 Then SAVERS.Close
SAVERS.Open "SELECT * FROM TRANSPORTMST WHERE NAME ='" & txtTransport & "' AND RECSTAT<>'D'", CN, adOpenDynamic, adLockOptimistic
If Not SAVERS.EOF Then
   M_TRCD = Trim(SAVERS!CODE & "")
Else
   M_TRCD = Empty
End If
SAVERS.Close

'VEHICLE CODE
If SAVERS.State = 1 Then SAVERS.Close
SAVERS.Open "SELECT * FROM VHCLMST WHERE NAME ='" & txtVHCL & "' AND RECSTAT<>'D'", CN, adOpenDynamic, adLockOptimistic
If Not SAVERS.EOF Then
   M_VHCD = Trim(SAVERS!CODE & "")
Else
   M_VHCD = Empty
End If
SAVERS.Close


If SAVERS.State = 1 Then SAVERS.Close
   SAVERS.Open "SELECT * FROM GPMST WHERE COMP = '" & compPth & "' AND UNIT = '" & UNCD & "' AND GPNO = '" & Trim(LBLCHLN) & "'", CN, adOpenDynamic, adLockOptimistic
   If Not SAVERS.EOF Then
   MsgBox "Gate Pass No. Already Exist!!! Change in Unit Configuration ", vbOKOnly
   Exit Sub
End If


CN.BeginTrans  'UPDATE CHALLAN DETAILS

CN.Execute "INSERT INTO GPMST(COMP,UNIT,GPNO,GPDT,TRCD,VHCD,INDATE,INTIME,OUTTIME,DRIVER,LICENCE,ADVANCE,BOXES," & _
           "GROSSWGT,RECSET) VALUES('" & compPth & "','" & UNCD & "','" & Trim(LBLCHLN) & "','" & Format(TXTVBDT, "MM/DD/YYYY") & "','" & M_TRCD & "','" & M_VHCD & _
           "','" & Format(InDate, "MM/DD/YYYY") & "','" & txtIN & "','" & txtOUT & "','" & txtDriver & _
           "','" & txtLicenceNo & "','" & Val(txtPaid) & "','" & Val(TTLPCS) & "','" & Val(TTLQTY) & "','A') "

Dim i As Long
Dim dbcd As String
Dim chln As String
Dim SQL As String
 
For i = 1 To lstChallan.ListItems.COUNT
If lstChallan.ListItems(i).Checked = True Then
   dbcd = lstChallan.ListItems(i).SubItems(10)
   chln = lstChallan.ListItems(i).Text
         
SQL = "UPDATE SPTRAN SET GATEPASSNO='" & Trim(LBLCHLN) & "',VEHICALNO='" & Trim(M_VHCD) & "',TRCD='" & M_TRCD & _
"',NOTRIPS='" & Val(TXTNOTRIP) & "',GP_REMARKS='" & Trim(TXTRMRK) & "' WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
"' AND DBCD='" & dbcd & "' AND VTYP='DPF' AND RECSTAT<>'D' AND VBNO='" & chln & "' AND RECSTAT = 'A'"
      
CN.Execute SQL
     
Dim RTYP As String: RTYP = Trim(lstChallan.ListItems(i).SubItems(11))
Dim SDBC As String: SDBC = Trim(lstChallan.ListItems(i).SubItems(12))
Dim SVBN As String: SVBN = Trim(lstChallan.ListItems(i).SubItems(13))


SQL = "UPDATE BILLMAIN SET TRCD='" & M_TRCD & "',VHCL='" & Trim(M_VHCD) & _
"',INTM = '" & Trim(txtIN) & "', RMTM = '" & Trim(txtOUT) & "'  where COMP='" & compPth & _
"' AND UNIT='" & UNCD & "' AND DBCD='" & SDBC & "' AND VBNO='" & SVBN & "' AND VTYP='SAL' AND RECSTAT='A'"
     
CN.Execute SQL

SQL = "UPDATE SERIALMASTER SET [SRNO]='" & LBLCHLN & "' WHERE COMP='" & compPth & "' AND UNIT = '" & UNCD & _
"' AND VTYP='GPS' AND CODE='000001' AND FYCD='" & FYCD & "'"
   
CN.Execute SQL
'----------------------------
'DAILYSTATUS ENTRY
  Call DAILYSTATUS("GPS", M_TRCD, dbcd, Val(TTLQTY), LBLCHLN, 0, cUName, "N", Now, TXTVBDT)
End If
Next i

CN.CommitTrans
MsgBox "YOUR GATE PASS NO. IS : " & LBLCHLN
  
'Call UPDATESTATUS
Call cmdCancel_Click

Exit Sub
LAST:
MsgBox ERR.Description
End Sub

Private Sub cmdSearch_Click()
Dim SQL As String
Dim M_ROW As Integer

lstChallan.ListItems.Clear

Screen.MousePointer = vbHourglass
Dim RECSET As ADODB.Recordset
Set RECSET = New ADODB.Recordset

SQL = "SELECT DISTINCT SPTRAN.*,ACCMST.NAME AS PARTY,PADDMST.NAME AS CONSIGNEE,PADDMST.ADDR AS ADDRESS,FINITMMST.NAME AS ITEM "
SQL = SQL & "  FROM SPTRAN "
SQL = SQL & " LEFT JOIN ACCMST ON ACCMST.CODE=SPTRAN.DRAC "
SQL = SQL & " INNER JOIN PADDMST ON PADDMST.CODE=SPTRAN.DCOD AND PADDMST.SRNO=SPTRAN.ADDRESS "
SQL = SQL & " INNER JOIN FINITMMST ON SPTRAN.COMP=FINITMMST.COMP AND SPTRAN.UNIT=FINITMMST.UNIT AND SPTRAN.DVCD=FINITMMST.DVCD AND SPTRAN.ICOD=FINITMMST.CODE "
SQL = SQL & " WHERE SPTRAN.COMP='" & compPth & "' AND SPTRAN.UNIT='" & UNCD & _
"' AND SPTRAN.VTYP='DPF' AND SPTRAN.RECSTAT<>'D'AND SPTRAN.DATE>='" & Format(txtFromDt, "MM/DD/YYYY") & _
"' AND SPTRAN.DATE<='" & Format(txtToDt, "MM/DD/YYYY") & "' AND SPTRAN.GATEPASSNO IS NULL "
SQL = SQL & " ORDER BY SPTRAN.DBCD,SPTRAN.VBNO"

If RECSET.State = 1 Then RECSET.Close
RECSET.Open SQL, CN, adOpenDynamic, adLockOptimistic
If RECSET.EOF Then
   MsgBox "No Records Found ": Screen.MousePointer = vbNormal: Exit Sub
End If
   
    Do While RECSET.EOF = False
        Set Item = lstChallan.ListItems.ADD
        Item.Text = RECSET!VBNO
        Item.SubItems(1) = Trim(RECSET!Date)
        Item.SubItems(2) = Trim(RECSET!PARTY & "")
        Item.SubItems(3) = Trim(RECSET!CONSIGNEE & "")
        Item.SubItems(4) = Trim(RECSET!ADDRESS & "")
        Item.SubItems(5) = Trim(RECSET!Item & "")
        If Trim(RECSET!grad & "") <> Empty Then
           Item.SubItems(6) = GetCode("GRDMST", Trim(RECSET!grad & ""), "CODE", "GRAD")
        End If
        Item.SubItems(7) = Trim(RECSET!ltno & "")
        Item.SubItems(8) = Trim(RECSET!PCES & "")
        Item.SubItems(9) = nstr(RECSET!GWGT, 12, 3)
        Item.SubItems(10) = Trim(RECSET!dbcd & "")
        Item.SubItems(11) = Trim(RECSET!RTYP & "")
        Item.SubItems(12) = Trim(RECSET!SDBC & "")
        Item.SubItems(13) = Trim(RECSET!SVBN & "")
        RECSET.MoveNext
    Loop
    RECSET.Close
    
lstChallan.SetFocus
Screen.MousePointer = vbNormal
End Sub

Private Sub Form_Activate()
  Call ColorComponent(Me)
  Me.BackColor = RGB(RED, GREEN, BLUE)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  SendKeys "{TAB}"
End If
End Sub

Private Sub Form_Load()
Call CenterChild(frm_Main, Me)
Call ColorComponent(Me)
On Error GoTo errLoad

M_DESC = Empty: Key = Empty: NEW_VISIBLE = False
  
txtFromDt = GetMinDate
txtToDt = GetMaxDate
InDate = Now
TXTVBDT = Now
InDate.MinDate = FSDT
InDate.MaxDate = FEDT
TXTVBDT.MinDate = FSDT
TXTVBDT.MaxDate = FEDT

  
  LBLCHLN.Caption = GenVNO("GPS", "000001")
  cmdExit.Cancel = True
  Me.Show
  Exit Sub
  
errLoad:
    ErrNumber = ERR.Number
    ErrMessage = ERR.Description
    frm_ErrorHandler.Show vbModal
End Sub

Private Sub InDate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
  SendKeys "{TAB}"
End If
End Sub

Private Sub lstChallan_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim i As Integer, J As Integer, ctr As Integer
    If Item.Checked = True Then
        TTLPCS = Val(TTLPCS) + 1
        TTLQTY = nstr(Val(TTLQTY) + Val(Item.SubItems(9)), 10, 3)
    Else
        TTLPCS = Val(TTLPCS) - 1
        TTLQTY = nstr(Val(TTLQTY) - Val(Item.SubItems(9)), 10, 3)
    End If
    If Item.INDEX < lstChallan.ListItems.COUNT Then lstChallan.ListItems.Item(Item.INDEX + 1).Selected = True: lstChallan.ListItems(Item.INDEX + 1).EnsureVisible
End Sub

Private Sub M_TRNM_GotFocus()
M_TRNM.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub M_TRNM_KeyDown(KeyCode As Integer, Shift As Integer)
  If (KeyCode = vbKeyReturn And M_TRNM = Empty) Or KeyCode = vbKeyF2 Then
   M_TRNM = SearchList1("SELECT TOP 20 Code,name from refmst where cata='R'", 0, Empty, "List of Transporter")
  End If
  If KeyCode = vbKeyDelete Then
   M_TRNM = Empty
  End If
End Sub

Private Sub M_TRNM_LostFocus()
M_TRNM.BackColor = vbWhite
End Sub

Private Sub M_VHCL_GotFocus()
M_VHCL.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub M_VHCL_LostFocus()
M_VHCL.BackColor = vbWhite
End Sub

Private Sub txtFromDt_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{tab}"
End Sub



Private Sub TXTNOTRIP_GotFocus()
TXTNOTRIP.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTNOTRIP_KeyPress(KeyAscii As Integer)
If CheckNumericKey(KeyAscii, TXTNOTRIP, Me) = 0 Then KeyAscii = 0
End Sub

Private Sub TXTNOTRIP_LostFocus()
TXTNOTRIP.BackColor = vbWhite
End Sub


Private Sub txtPaid_KeyPress(KeyAscii As Integer)
 If CheckNumericKey(KeyAscii, txtPaid, Me) = 0 Then KeyAscii = 0
End Sub

Private Sub TXTRMRK_GotFocus()
TXTRMRK.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTRMRK_LostFocus()
TXTRMRK.BackColor = vbWhite
End Sub

Private Sub txtToDt_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{tab}"
End Sub

Private Function CHKSAVEDATA() As Boolean
Dim FLAG As Boolean
Dim INDEX As Long
  CHKSAVEDATA = True
  
  If LBLCHLN = Empty Or Len(LBLCHLN) <> 10 Or (LBLCHLN.Caption) = "XXXXXXXXXX" Then
    CHKSAVEDATA = False
    MsgBox "CHALLAN IS NOT PROPER GENERATED"
    txtFromDt.Enabled = True
    txtFromDt.SetFocus
    Exit Function
  End If
  
  If txtVHCL.Text = Empty Then
    CHKSAVEDATA = False
    framTransDetail.Visible = True
    txtVHCL.Enabled = True
    txtVHCL.SetFocus
    Exit Function
  End If
  
  If txtTransport.Text = Empty Then
    CHKSAVEDATA = False
    framTransDetail.Visible = True
    txtVHCL.Enabled = True
    txtVHCL.SetFocus
    Exit Function
  End If
  
FLAG = False
  
For INDEX = 1 To lstChallan.ListItems.COUNT
  If lstChallan.ListItems(INDEX).Checked = True Then: FLAG = True: Exit For
Next

If FLAG = False Then
    CHKSAVEDATA = False
    txtFromDt.Enabled = True
    txtFromDt.SetFocus
    Exit Function
End If

End Function

Private Sub txtVHCL_KeyDown(KeyCode As Integer, Shift As Integer)
 Me.KeyPreview = False
 
  If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
     txtVHCL = Empty
  ElseIf KeyCode = vbKeyF2 Or txtVHCL = Empty Then
     M_DESC = Empty:   NEW_VISIBLE = False
     txtVHCL = SearchList1("Select DISTINCT CODE,NAME From VHCLMST WHERE RECSTAT='A'", 0, Empty, "Select Vehicle From List. ")
     txtVHCL.Tag = Key
     Call FindDetails
     txtTransport.Tag = GetCode("VHCLMST", txtVHCL.Tag, "CODE", "TRCD")
     txtTransport = GetCode("TRANSPORTMST", txtTransport.Tag, "CODE", "NAME")
  End If
  
 Me.KeyPreview = True
End Sub

Private Sub TXTVHCL_GotFocus(): txtVHCL.BackColor = RGB(BRED, BGREEN, BBLUE): End Sub
Private Sub TXTVHCL_LostFocus(): txtVHCL.BackColor = vbWhite: End Sub
Private Sub txtTransport_GotFocus(): txtTransport.BackColor = RGB(BRED, BGREEN, BBLUE): End Sub
Private Sub txtTransport_LostFocus(): txtTransport.BackColor = vbWhite: End Sub
Private Sub txtDriver_GotFocus(): txtDriver.BackColor = RGB(BRED, BGREEN, BBLUE): End Sub
Private Sub txtDriver_LostFocus(): txtDriver.BackColor = vbWhite: End Sub
Private Sub txtLicenceNo_GotFocus(): txtLicenceNo.BackColor = RGB(BRED, BGREEN, BBLUE): End Sub
Private Sub txtLicenceNo_LostFocus(): txtLicenceNo.BackColor = vbWhite: End Sub
Private Sub txtPaid_GotFocus(): txtPaid.BackColor = RGB(BRED, BGREEN, BBLUE): End Sub
Private Sub txtPaid_LostFocus(): txtPaid.BackColor = vbWhite: End Sub
Private Sub txtNOT_GotFocus(): TXTNOT.BackColor = RGB(BRED, BGREEN, BBLUE): End Sub
Private Sub txtNOT_LostFocus(): TXTNOT.BackColor = vbWhite: End Sub

Private Sub FindDetails()
    If Trim(txtVHCL.Tag) = Empty Then Exit Sub
    
    Dim TMPRS As ADODB.Recordset
    Set TMPRS = New ADODB.Recordset
    If TMPRS.State = 1 Then TMPRS.Close
    TMPRS.Open "SELECT * FROM VHCLENTRY WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND VHCD='" & txtVHCL.Tag & "' ORDER BY CODE DESC", CN, adOpenDynamic, adLockOptimistic
    If Not TMPRS.EOF Then
      txtIN = TMPRS!INTIME & ""
      InDate = Format(TMPRS!Date, "DD/MM/YYYY")
      txtDriver = Trim(TMPRS!DRIVER) & ""
      txtLicenceNo = Trim(TMPRS!LICENCE) & ""
      TXTRMRK = Trim(TMPRS!REMARKS) & ""
    Else
        txtIN = Format(Now, "HH:MM")
        txtDriver = ""
        txtLicenceNo = ""
        TXTRMRK = ""
        InDate = Format(Now, "DD/MM/YYYY")
        
    End If
    TMPRS.Close
End Sub
