VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{2BF72F7D-D367-4712-9146-5533EF3E691A}#1.2#0"; "FraPlus1.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WELCHBUTTON.OCX"
Begin VB.Form frmPurchaseServices 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Service GRN "
   ClientHeight    =   5820
   ClientLeft      =   2865
   ClientTop       =   1155
   ClientWidth     =   10875
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   10875
   Begin FramePlusCtl.FramePlus Frm1 
      Height          =   5835
      Left            =   0
      TabIndex        =   28
      Top             =   0
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   10292
      BackgroundPictureAlignment=   5
      BorderStyle     =   10
      BackColorGradient=   -2147483637
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
         ItemData        =   "frmPurchaseServices.frx":0000
         Left            =   4800
         List            =   "frmPurchaseServices.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Tag             =   "0"
         Top             =   3120
         Width           =   1575
      End
      Begin VB.TextBox txtRMRK 
         Appearance      =   0  'Flat
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
         Height          =   615
         Left            =   1920
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   25
         Top             =   3990
         Width           =   4845
      End
      Begin VB.Frame FRMBTRM 
         Height          =   2655
         Left            =   6960
         TabIndex        =   29
         Top             =   2880
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
            TabIndex        =   32
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
            TabIndex        =   31
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
            TabIndex        =   30
            Text            =   "0.00"
            Top             =   1320
            Visible         =   0   'False
            Width           =   1575
         End
         Begin MSFlexGridLib.MSFlexGrid flexBTRM 
            Height          =   1995
            Left            =   120
            TabIndex        =   26
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
            TabIndex        =   33
            Top             =   2280
            Width           =   1305
         End
      End
      Begin VB.TextBox TXTITOT 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   315
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   23
         Top             =   3690
         Width           =   1575
      End
      Begin VB.TextBox TXTMDESC 
         Appearance      =   0  'Flat
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
         Height          =   615
         Left            =   1920
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Top             =   2190
         Width           =   8805
      End
      Begin VB.TextBox TXTNARR 
         Appearance      =   0  'Flat
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
         Height          =   615
         Left            =   1920
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Top             =   1590
         Width           =   8805
      End
      Begin VB.TextBox TXTVBNO 
         Appearance      =   0  'Flat
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
         Left            =   1920
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   7
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox VBNO 
         Appearance      =   0  'Flat
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
         Height          =   315
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   17
         Top             =   2790
         Width           =   1575
      End
      Begin VB.TextBox TXTSAMT 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   315
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   21
         Top             =   3390
         Width           =   1575
      End
      Begin VB.TextBox TXTDBAC 
         Appearance      =   0  'Flat
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
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   1320
         Width           =   5295
      End
      Begin MSComCtl2.DTPicker VBDT 
         Height          =   315
         Left            =   1935
         TabIndex        =   19
         Top             =   3090
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   53346305
         CurrentDate     =   39383
      End
      Begin MSComCtl2.DTPicker TXTVBDT 
         Height          =   315
         Left            =   1920
         TabIndex        =   9
         Top             =   1005
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   53346305
         CurrentDate     =   39383
      End
      Begin WelchButton.lvButtons_H cmdAdd 
         Height          =   495
         Left            =   405
         TabIndex        =   0
         Top             =   4920
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
         Image           =   "frmPurchaseServices.frx":0004
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdEdit 
         Height          =   495
         Left            =   3600
         TabIndex        =   3
         Top             =   4920
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
         Image           =   "frmPurchaseServices.frx":039E
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdDelete 
         Height          =   495
         Left            =   4680
         TabIndex        =   4
         Top             =   4920
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
         Image           =   "frmPurchaseServices.frx":0738
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdSave 
         Height          =   495
         Left            =   1440
         TabIndex        =   1
         Top             =   4920
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
         Image           =   "frmPurchaseServices.frx":0AD2
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdCancel 
         Height          =   495
         Left            =   2520
         TabIndex        =   2
         Top             =   4920
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
         Image           =   "frmPurchaseServices.frx":185C
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdExit 
         Height          =   495
         Left            =   5760
         TabIndex        =   5
         Top             =   4920
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
         Image           =   "frmPurchaseServices.frx":1CAE
         cBack           =   -2147483633
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Tax Type :"
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
         Left            =   3720
         TabIndex        =   38
         Top             =   3120
         Width           =   975
      End
      Begin VB.Label LBLDIV 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "DIVISION : "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Top             =   240
         Width           =   4575
      End
      Begin VB.Label Label17 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "   Remarks                 ( If Any)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   240
         TabIndex        =   24
         Top             =   3990
         Width           =   1755
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2835
         Index           =   13
         Left            =   6840
         TabIndex        =   35
         Top             =   2760
         Width           =   3855
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Index           =   12
         Left            =   240
         TabIndex        =   34
         Top             =   4800
         Width           =   6615
      End
      Begin VB.Label Label17 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "   GRN Date :"
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
         Index           =   11
         Left            =   240
         TabIndex        =   8
         Top             =   1020
         Width           =   1695
      End
      Begin VB.Label Label17 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "   GRN No. :"
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
         Index           =   10
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label17 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "   Material Cost"
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
         Index           =   7
         Left            =   240
         TabIndex        =   22
         Top             =   3690
         Width           =   1695
      End
      Begin VB.Label Label17 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "   Material                 Description"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   6
         Left            =   240
         TabIndex        =   14
         Top             =   2190
         Width           =   1725
      End
      Begin VB.Label Label17 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "   Service Chln Dt. "
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
         Index           =   5
         Left            =   240
         TabIndex        =   18
         Top             =   3090
         Width           =   1695
      End
      Begin VB.Label Label17 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "   Service Chln No. "
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
         Index           =   4
         Left            =   240
         TabIndex        =   16
         Top             =   2790
         Width           =   1695
      End
      Begin VB.Label Label17 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "   Service Amt. "
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
         Index           =   3
         Left            =   240
         TabIndex        =   20
         Top             =   3390
         Width           =   1695
      End
      Begin VB.Label Label17 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "   Service                  Description"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   2
         Left            =   240
         TabIndex        =   12
         Top             =   1590
         Width           =   1725
      End
      Begin VB.Label Label17 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "   Service Provider "
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
         Index           =   1
         Left            =   240
         TabIndex        =   10
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label LBLHEAD 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         Caption         =   "GRN SERVICES"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   7920
         TabIndex        =   27
         Top             =   720
         Width           =   2535
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         Height          =   5655
         Left            =   120
         Top             =   120
         Width           =   10695
      End
   End
End
Attribute VB_Name = "frmPurchaseServices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public M_DBCD_DIRIVR As String
Public DAYNAM As String
Public SAVEFLAG As Boolean
Public M_VBNO As String
Public M_DVCD As String
Dim DefaultName As String
Dim ALLOWEDITDEL As Boolean

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

Private Sub txtBEdit_GotFocus()
txtBEdit.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtBEdit_KeyDown(KeyCode As Integer, Shift As Integer)
  EditKeyCode flexBTRM, txtBEdit, KeyCode, Shift
End Sub

Private Sub txtBEdit_KeyPress(KeyAscii As Integer)
If KeyAscii = Asc(vbCr) Then KeyAscii = 0
  
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

Private Sub txtBEdit_LostFocus()
  calBTRM 0
  Call calADLS
End Sub

Private Sub TXTITOT_Change()
If flexBTRM.Rows > 0 Then
    flexBTRM.COL = 0
    flexBTRM.ROW = 0
  End If
  calBTRM 0
  Call calADLS
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

Private Sub TXTRMRK_GotFocus()
    txtRMRK.BackColor = RGB(BRED, BGREEN, BBLUE)
    SendKeys "{END}"
End Sub

Private Sub TXTRMRK_LostFocus()
    txtRMRK.BackColor = vbWhite
End Sub

Private Sub TXTSAMT_Change()
If flexBTRM.Rows > 0 Then
    flexBTRM.COL = 0
    flexBTRM.ROW = 0
End If
    calBTRM 0
    Call calADLS
End Sub

Private Sub Form_Load()
    Call CenterChild(frm_Main, Me)
    Me.KeyPreview = True
    M_DBCD_DIRIVR = "000001"
    M_DBCD_DIRIVR = "000001"
    M_VBNO = Empty
 
    VBDT = Date
    VBDT.MaxDate = FEDT
    VBDT.MinDate = FSDT
    TXTVBDT = Date
    TXTVBDT.MaxDate = FEDT
    TXTVBDT.MinDate = FSDT
    
    cmbSelection.AddItem "NONE"
    cmbSelection.AddItem "RG23-A"
    cmbSelection.AddItem "RG23-C"
    cmbSelection.ListIndex = 0
    
  M_DESC = Empty: Key = Empty: NEW_VISIBLE = False: DIVCOD = Empty: DIVNAM = Empty
  
  If DIVCOD = Empty Then
    DIVNAM = SearchList1("SELECT  TOP 20 CODE,NAME FROM DIVMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
                         "' AND RECSTAT='A'", 0, "", "SELECT DIVISION MASTER")
  End If
  
  If RS.State = 1 Then RS.Close
  RS.Open "SELECT * FROM DIVMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND NAME='" & DIVNAM & "'", CN, adOpenKeyset, adLockPessimistic
  If Not RS.EOF Then
     DIVCOD = RS!CODE
     DIVNAM = RS!NAME
     DIVNM = RS!NAME
     Me.Tag = DIVCOD
     LBLDIV.Caption = "DIVISION : " + DIVNAM
    Else
     LBLDIV.Caption = "DIVISION : " + "??????"
  End If
    
    Call FIL_Billingterm
    Call btn_sts(True)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
 If UCase(ActiveControl.NAME) = "M_CVBN" And KeyAscii = vbKeyReturn And M_CVBN <> Empty Then cmdSave.Enabled = True: cmdSave.SetFocus
 If UCase(ActiveControl.NAME) = "VBNO" And Trim(VBNO) = Empty Then Exit Sub
 
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub cmdAdd_Click()
    zoomflag = False
    btn_sts (False)
    M_VBNO = Empty
    TXTVBNO = GenVNO("PSR", M_DBCD_DIRIVR)
    SAVEFLAG = True
    If TXTDBAC.Enabled Then TXTDBAC.SetFocus
End Sub

Private Sub TXTSAMT_GotFocus()
  TXTSAMT.BackColor = RGB(BRED, BGREEN, BBLUE)
  SendKeys "{HOME}+{END}"
End Sub

Private Sub TXTSAMT_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 And InStr(1, TXTSAMT, ".", vbTextCompare) > 0 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If Val(TXTSAMT) <> "0" Then
            TXTSAMT.SetFocus
        Else
            KeyAscii = 0
            MsgBox "Please Enter Amount", vbInformation, "Amount Required"
        End If
    End If
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 46 Then KeyAscii = 0
End Sub

Private Sub TXTSAMT_LostFocus()
TXTSAMT.BackColor = vbWhite
End Sub

Private Sub cmdCancel_Click()
    Call ClsData(Me)
    Call btn_sts(True)
    M_VBNO = Empty
    If zoomflag = True Then
        Call CMDEXIT_Click
        Exit Sub
    End If
End Sub

Private Sub cmdDelete_Click()
  If M_USRSECLEVL = "1" Then
      If ReadConfigMaster("000038", 6, "M") = False Then ModuleDeniedMessage: Exit Sub
  End If
        
  ALLOWEDITDEL = True
  SAVEFLAG = False
  M_VBNO = Empty
  btn_sts (False)
  frmPurServicesList.Show 1
  If ALLOWEDITDEL = False Then
    MsgBox "Purchase of this GRN have been made can not edit/delete ", vbInformation
   Else
    'Check for Receipt and Payment Entires
    If Not M_VBNO = Empty Then
      Dim AYS
      AYS = MsgBox("Are you sure to delete this Service GRN ", vbYesNo)
      If AYS = vbYes Then
        CN.BeginTrans
        CN.Execute "UPDATE GRN SET RECSTAT='D' WHERE COMP='" & compPth & "' AND UNIT ='" & UNCD & "' AND VTYP='PSR' " & _
                   "AND LTRIM(RTRIM(VBNO))='" & M_VBNO & "' AND DBCD='" & M_DBCD_DIRIVR & "' AND RECSTAT<>'D'"
                   
        CN.Execute "DELETE FROM EGPMAN WHERE COMP='" & compPth & "' AND UNIT ='" & UNCD & "' AND VTYP='PSR'" & _
                   "AND LTRIM(RTRIM(VBNO))='" & M_VBNO & "' AND DBCD='" & M_DBCD_DIRIVR & "' AND RECSTAT<>'D'"
                   
        CN.Execute "DELETE FROM EGPMAN WHERE  COMP='" & compPth & "' AND UNIT='" & UNCD & _
                   "' AND DBCD='RG23-C' AND VTYP='PSR' AND VBNO='" & TXTVBNO & "' AND RECSTAT<>'D'"
                           
        CN.Execute "UPDATE STORETRAN SET RECSTAT='D' WHERE COMP='" & compPth & "' AND UNIT ='" & UNCD & _
        "' AND VTYP='PSR' AND DBCD='" & M_DBCD_DIRIVR & "' AND LTRIM(RTRIM(VBNO))='" & M_VBNO & "' AND RECSTAT<>'D'"
                
        
        'Call UPDATEDELSTATUS
        'Call DAILYSTATUS("PSR", GetCode("ACCMST", TXTDBAC, "NAME", "CODE"), M_DBCD_DIRIVR, 0, TXTVBNO, Val(TXTSAMT), cUName, "D", Now, TXTVBDT)
        CN.CommitTrans
      End If
    End If
  End If
  Call cmdCancel_Click
End Sub

Private Sub cmdEdit_Click()
    If M_USRSECLEVL = "1" Then
        If ReadConfigMaster("000038", 5, "M") = False Then ModuleDeniedMessage: Exit Sub
    End If
    SAVEFLAG = False
    frmPurServicesList.Show 1
    If M_VBNO <> Empty Then
        btn_sts (False)
        TXTDBAC.SetFocus
    Else
        btn_sts (True)
        cmdAdd.SetFocus
    End If
End Sub

Private Sub CMDEXIT_Click()
    Unload Me
End Sub



Private Sub TXTDBAC_GotFocus()
    TXTDBAC.BackColor = RGB(BRED, BGREEN, BBLUE)
    SendKeys "{HOME}+{END}"
End Sub

Private Sub TXTDBAC_KeyDown(KeyCode As Integer, Shift As Integer)
    Me.KeyPreview = False
    If KeyCode = vbKeyF2 Or (Trim(TXTDBAC.Text) = Empty And KeyCode = 13) Then
        NEW_VISIBLE = True
        M_DESC = Empty
        Key = Empty
        TXTDBAC.Text = SearchList1("SELECT TOP 20 CODE,NAME FROM ACCMST WHERE DRCR='C'", 0, TXTDBAC, "List of Db A/c")
        If key_PressNew = True Then
            blnNewButton = True
            M_DESC = ""
            Key = ""
            TXTDBAC.Text = ""
            frm_Acc.Show
            frm_Acc.ZOrder
            Exit Sub
        Else
            TXTDBAC.Tag = Key
        End If
    ElseIf KeyCode = vbKeyDelete Then
        TXTDBAC = Empty
    End If
    Me.KeyPreview = True
End Sub

Private Sub TXTDBAC_LostFocus()
  TXTDBAC.BackColor = vbWhite
End Sub

Public Sub btn_sts(Yes As Boolean)
    cmdSave.Enabled = Not Yes
    cmdCancel.Enabled = Not Yes
    cmdAdd.Enabled = Yes
    cmdEdit.Enabled = Yes
    cmdDelete.Enabled = Yes
    TXTDBAC.Enabled = Not Yes
    TXTVBDT.Enabled = Not Yes
    VBDT.Enabled = Not Yes
    VBNO.Enabled = Not Yes
    TXTSAMT.Enabled = Not Yes
    TXTNARR.Enabled = Not Yes
    TXTMDESC.Enabled = Not Yes
    TXTITOT.Enabled = Not Yes
    txtRMRK.Enabled = Not Yes
    VBNO.Enabled = Not Yes
End Sub

Private Sub TXTMDESC_GotFocus()
 TXTMDESC.BackColor = RGB(BRED, BGREEN, BBLUE)
 SendKeys "{END}"
End Sub

Private Sub TXTMDESC_LostFocus()
  TXTMDESC.BackColor = vbWhite
End Sub

Private Sub TXTNARR_GotFocus()
TXTNARR.BackColor = RGB(BRED, BGREEN, BBLUE)
SendKeys "{END}"
End Sub

Private Sub TXTNARR_LostFocus()
TXTNARR.BackColor = vbWhite
End Sub

Private Sub VBDT_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub VBDT_LostFocus()
    If DatePart("W", VBDT.Value) = vbSunday Then
        Dim ANS As String
        ANS = MsgBox("The date is fall on Sunday. Do you wish to continue ?", vbYesNo + vbQuestion)
        If ANS = "7" Then
            Call cmdCancel_Click
        End If
    End If
End Sub

Private Sub VBNO_GotFocus()
  VBNO.BackColor = RGB(BRED, BGREEN, BBLUE)
  SendKeys "{HOME}+{END}"
End Sub

Private Sub VBNO_LostFocus()
VBNO.BackColor = vbWhite
End Sub

Private Sub cmdSave_Click()
On Error GoTo LAST

 If CHKSAVEDATA = False Then
    Exit Sub
 End If
 
 If SAVEFLAG = True Then
    TXTVBNO = GenVNO("PSR", M_DBCD_DIRIVR)
    M_VBNO = TXTVBNO
 End If
 
 Call SAVERECIVR
  
  If SAVEFLAG = True Then
    MsgBox "Your GRN Service No. is " + TXTVBNO.Text
  Else
    MsgBox "Service GRN Successfully Edited."
  End If
  
  Call cmdCancel_Click
  Exit Sub
    
LAST:
    MsgBox ERR.Description
    If RS.State = 1 Then
        RS.CancelUpdate
    End If
    CN.RollbackTrans
    Exit Sub
End Sub

Private Sub UPDATESTATUS()
    Dim DLYSTA As New ADODB.Recordset
    If DLYSTA.State = adStateOpen Then DLYSTA.Close
    DLYSTA.Open "SELECT * FROM DAILYSTAT WHERE CUSR='JNMOPAWCGBDSXZAS'", CN, adOpenKeyset, adLockPessimistic
    DLYSTA.AddNew
    DLYSTA!COMP = compPth & ""
    DLYSTA!VTYP = "PSR"
    DLYSTA!PCOD = ""
    DLYSTA!dbcd = ""
    DLYSTA!QNTY = 0
    DLYSTA!VBNO = VBNO & ""
    DLYSTA!AMNT = Val(TXTSAMT)
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

Private Function CHKSAVEDATA() As Boolean
      
  If Trim(TXTNARR) = Empty Then
     MsgBox "Service Description required.", vbCritical
     TXTNARR.Enabled = True
     TXTNARR.SetFocus
     CHKSAVEDATA = False
     Exit Function
  End If
  
  If Val(TXTITOT) > 0 And Trim(TXTMDESC) = Empty Then
     MsgBox "Material Description required.", vbCritical
     TXTMDESC.Enabled = True
     TXTMDESC.SetFocus
     CHKSAVEDATA = False
     Exit Function
  End If
  
  If cmbSelection.Text = "RG23-A" Or cmbSelection = "RG23-C" Or cmbSelection = "NONE" Then
  Else
     MsgBox "Tax Type Required", vbCritical
     cmbSelection.Enabled = True
     cmbSelection.SetFocus
     CHKSAVEDATA = False
     Exit Function
 End If
  
       
  Dim CHKRS As New ADODB.Recordset
  Set CHKRS = New ADODB.Recordset
  
  'Party  A/c Code
  If CHKRS.State = 1 Then CHKRS.Close
  CHKRS.Open "SELECT CODE from ACCMST WHERE NAME='" & TXTDBAC.Text & "'", CN, adOpenKeyset, adLockPessimistic
  If CHKRS.EOF Then
     MsgBox "Debit A/c Name Not Define ", vbCritical
     TXTDBAC.Enabled = True
     TXTDBAC.SetFocus
     CHKSAVEDATA = False
     Exit Function
  Else
     TXTDBAC.Tag = Trim(CHKRS!CODE & "")
  End If
  
  If SAVEFLAG = True Then
     TXTVBNO = GenVNO("PSR", M_DBCD_DIRIVR)
     If CHKRS.State = 1 Then CHKRS.Close
     CHKRS.Open "SELECT * FROM GRN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
      "' AND VTYP='PSR' AND DBCD='" & M_DBCD_DIRIVR & "' AND VBNO='" & TXTVBNO & "'", CN, adOpenDynamic, adLockOptimistic
     If Not CHKRS.EOF Then
        MsgBox "Duplicate GRN No. !!!! ", vbCritical
        CHKSAVEDATA = False
        Exit Function
     End If
  End If
  
  If Val(TXTSAMT) <= 0 Then
     MsgBox "Amount Should be Greater than 0  !!!! ", vbCritical
     If TXTSAMT.Enabled Then TXTSAMT.SetFocus
     CHKSAVEDATA = False
     Exit Function
  End If
  
  If Trim(VBNO) = Empty Then
     MsgBox "Party Challan No. Not Define", vbCritical
     If VBNO.Enabled Then VBNO.SetFocus
     CHKSAVEDATA = False
     Exit Function
  End If
  
  SQL = "SELECT * FROM GRN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND VTYP='PSR' " & _
        "AND DBCD='" & M_DBCD_DIRIVR & "' AND CVBN='" & VBNO & "' AND RECSTAT<>'D' AND " & _
        "VBNO<>'" & TXTVBNO & "' AND VBNO LIKE '%" & FYCD & "' AND PCOD='" & TXTDBAC.Tag & "'"
            
  If CHKRS.State = 1 Then CHKRS.Close
  CHKRS.Open SQL, CN, adOpenDynamic, adLockOptimistic
  If Not CHKRS.EOF Then
     MsgBox "Party Challan no. Already Exist !!!! ", vbCritical
     If VBNO.Enabled Then VBNO.SetFocus
     CHKSAVEDATA = False
     Exit Function
  End If
  
  CHKSAVEDATA = True
End Function

Private Sub SAVERECIVR()
  On Error GoTo LAST
  Dim SQL As String
  Dim i As Long, J As Long
  
  Dim SAVDAT As New ADODB.Recordset
  Dim MSTDAT As New ADODB.Recordset
  Set SAVDAT = New ADODB.Recordset
  Set MSTDAT = New ADODB.Recordset
  
  '==========================================================================
  Dim M_CRAC As String, M_DRAC As String, M_PCOD As String, M_DCOD As String
  Dim M_ARCD As String, M_TRCD As String, M_BRCD As String, M_CPCD As String
  Dim M_CCCD As String, M_DPTC As String, M_CHEAD As String
  
  
  Dim SDESC As String
  Dim MDESC As String
  
  SDESC = Replace(Trim(TXTNARR), vbCrLf, "")
  MDESC = Replace(Trim(TXTMDESC), vbCrLf, "")
  txtRMRK = Replace(Trim(txtRMRK), vbCrLf, "")
  
  Dim excperc As Double
  excperc = 100
  If cmbSelection = "RG23-C" Then
    If RS.State = 1 Then RS.Close
    RS.Open "select exccperc from untcfg where comp='" & compPth & "' and unit='" & UNCD & "'", CN, adOpenDynamic, adLockOptimistic
    If Not RS.EOF Then
      excperc = RS!exccperc
    End If
  End If
    
  'Party A/c
  If SAVDAT.State = 1 Then SAVDAT.Close
  SAVDAT.Open "SELECT * FROM ACCMST WHERE NAME='" & TXTDBAC.Text & "'", CN, adOpenDynamic, adLockOptimistic
  If Not SAVDAT.EOF Then
    M_DRAC = SAVDAT!CODE & ""
    M_PCOD = SAVDAT!CODE & ""
    M_CPCD = SAVDAT!CPCD & ""
    M_ARCD = SAVDAT!ARCD & ""
    M_BRCD = SAVDAT!BRCD & ""
  End If
  SAVDAT.Close
  '==========================================================================
    
  CN.BeginTrans
  Call DELETEIVR
  
  'GRN DETAILS ==========================================================================
  SQL = Empty
  If SAVDAT.State = 1 Then SAVDAT.Close
  SAVDAT.Open "SELECT * FROM GRN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND VTYP='PSR' " & _
              "AND VBNO='" & TXTVBNO & "' AND DBCD='" & M_DBCD_DIRIVR & "' AND RECSTAT<>'D'", CN, adOpenDynamic, adLockOptimistic
  If SAVDAT.EOF Then
      SAVDAT.AddNew
  End If
  
  SAVDAT!COMP = compPth
  SAVDAT!unit = UNCD
  SAVDAT!dbcd = M_DBCD_DIRIVR
  SAVDAT!VTYP = "PSR"
  SAVDAT!VBNO = TXTVBNO
  
  SAVDAT!CVBN = VBNO
  SAVDAT!LRDT = Format(VBDT.Value, "YYYY/MM/DD")
  
  SAVDAT!DVCD = DIVCOD
  SAVDAT!SRNO = TXTVBNO
  SAVDAT!SRCH = 1
  
  SAVDAT!Date = Format(TXTVBDT.Value, "YYYY/MM/DD")
  SAVDAT!VBNO = Trim(TXTVBNO.Text)
  SAVDAT!TTYP = cmbSelection.Text
        
  'SAVDAT!CRAC = M_CRAC
  SAVDAT!DRAC = M_PCOD
  SAVDAT!PCOD = M_PCOD
  SAVDAT!DCOD = M_DCOD
  SAVDAT!BRCD = M_BRCD
  SAVDAT!CPCD = M_CPCD
  SAVDAT!ARCD = M_ARCD
  SAVDAT!TPCS = 0
  SAVDAT!TQTY = 0
  
  SAVDAT!SAMT = Val(TXTSAMT)
  SAVDAT!ITOT = Val(TXTITOT)
  SAVDAT!BADJ = Val(TXTSAMT) + Val(TXTITOT) - Val(TXTBNET)
  SAVDAT!BNET = Val(TXTBNET)
      
  SAVDAT!MDESC = MDESC
  SAVDAT!SDESC = SDESC
  SAVDAT!BRMK = txtRMRK
    
  If SAVEFLAG = True Then
    SAVDAT!SYSR = "N"
   Else
    SAVDAT!SYSR = "U"
  End If
  
  SAVDAT![User] = cUName & ""
  SAVDAT!RECSTAT = "A"
  
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
  
  'EXCISE DETAILS=========================================================================================
    
  Set EXCISE = New ADODB.Recordset
  If EXCISE.State = 1 Then EXCISE.Close
  
  CN.Execute "DELETE FROM EGPMAN WHERE  COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DBCD='" & _
               M_DBCD_DIRIVR & "' AND VTYP='PSR' AND VBNO='" & TXTVBNO & "' AND RECSTAT<>'D'"
  
  CN.Execute "DELETE FROM EGPMAN WHERE  COMP='" & compPth & "' AND UNIT='" & UNCD & _
             "' AND DBCD='RG23-C' AND VTYP='PSR' AND VBNO='" & TXTVBNO & "' AND RECSTAT<>'D'"
  
  EXCISE.Open "SELECT * FROM EGPMAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DBCD='" & _
               M_DBCD_DIRIVR & "' AND VTYP='PSR' AND VBNO='" & TXTVBNO & "' AND RECSTAT<>'D'", CN, adOpenDynamic, adLockOptimistic
               
  If EXCISE.EOF Then
     EXCISE.AddNew
  End If
  
  EXCISE!COMP = compPth
  EXCISE!unit = UNCD
  EXCISE!dbcd = M_DBCD_DIRIVR
  EXCISE!VTYP = "PSR"
  EXCISE!VBNO = M_VBNO
  EXCISE!SRNO = M_VBNO
  EXCISE!SRCH = 1
  EXCISE!Date = Format(TXTVBDT, "YYYY/MM/DD")
  
  'EXCISE!CRAC = M_CRAC & ""
  EXCISE!DRAC = M_PCOD & ""
  EXCISE!VBNO = TXTVBNO
  EXCISE!chln = Trim(VBNO)
  EXCISE!CHDT = Format(VBDT, "YYYY/MM/DD")
  
  EXCISE!PCES = 0
  EXCISE!QNTY = 0
  EXCISE!AMNT = Val(TXTITOT) + Val(TXTSAMT)
  EXCISE!ITOT = Val(TXTITOT)
  EXCISE!BADJ = Val(TXTITOT) + Val(TXTSAMT) - Val(TXTBNET)
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
    EXCISE.Open "SELECT * FROM EGPMAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DBCD='RG23-C' AND VTYP='PSR' AND VBNO='" & TXTVBNO & "' AND RECSTAT<>'D'", CN, adOpenDynamic, adLockOptimistic
               
    If EXCISE.EOF Then
        EXCISE.AddNew
     End If
  
     EXCISE!COMP = compPth
     EXCISE!unit = UNCD
     EXCISE!dbcd = "RG23-C"
     EXCISE!VTYP = "PSR"
     EXCISE!VBNO = M_VBNO
     EXCISE!SRNO = M_VBNO
     EXCISE!SRCH = 1
     EXCISE!Date = Format(FEDT + 1, "YYYY/MM/DD")
       
     EXCISE!DRAC = M_PCOD & ""
     EXCISE!VBNO = TXTVBNO
     EXCISE!chln = Trim(VBNO)
     EXCISE!CHDT = Format(FEDT + 1, "YYYY/MM/DD")
  
     EXCISE!PCES = 0
     EXCISE!QNTY = 0
     EXCISE!AMNT = Val(TXTITOT) + Val(TXTSAMT)
     EXCISE!ITOT = Val(TXTITOT)
     EXCISE!BADJ = Val(TXTITOT) + Val(TXTSAMT) - Val(TXTBNET)
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
  
  'UPDATE VOUCHER TYPE MASTER
  If SAVEFLAG = True Then
   Call SetSRNO(TXTVBNO, "PSR", M_DBCD_DIRIVR)
  End If
  
  'Call UPDATESTATUS
  '-------------------------------------
  'DAILYSTATUS ENTRY
  If SAVEFLAG = True Then
     Call DAILYSTATUS("PSR", M_PCOD, M_DBCD_DIRIVR, 0, TXTVBNO, Val(TXTSAMT), cUName, "N", Now, TXTVBDT)
  Else
     Call DAILYSTATUS("PSR", M_PCOD, M_DBCD_DIRIVR, 0, TXTVBNO, Val(TXTSAMT), cUName, "M", Now, TXTVBDT)
  End If
  '--------------------------------------
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
  Dim SAVDAT As New ADODB.Recordset
  Dim m_rtyp As String
  Dim m_rsrn As String
  Set SAVDAT = New ADODB.Recordset
  If SAVDAT.State = 1 Then SAVDAT.Close
    
  CN.Execute "DELETE FROM EGPMAN WHERE COMP='" & compPth & "' AND UNIT ='" & UNCD & "' AND VTYP='PSR' AND " & _
             "LTRIM(RTRIM(VBNO))='" & M_VBNO & "' AND DBCD='" & M_DBCD_DIRIVR & "'"
  CN.Execute "DELETE FROM EGPMAN WHERE COMP='" & compPth & "' AND UNIT ='" & UNCD & "' AND VTYP='PSR' AND " & _
             "LTRIM(RTRIM(VBNO))='" & M_VBNO & "' AND DBCD='RG23-C'"
End Sub

Private Sub UPDATEDELSTATUS()
  Dim DLYSTA As New ADODB.Recordset
  If DLYSTA.State = adStateOpen Then DLYSTA.Close
  DLYSTA.Open "SELECT * FROM DAILYSTAT WHERE CUSR='JNMOPAWCGBDSXZAS'", CN, adOpenKeyset, adLockPessimistic
  DLYSTA.AddNew
  DLYSTA!COMP = compPth & ""
  DLYSTA!VTYP = "PSR"
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
    'If M_BILRDOF = "Y" Then
    '    TXTBNET.Text = Format(FormatNumber(Val(TXTSAMT.Text) + Val(TXTITOT.Text) + Val(TXTADLS.Text), 0), "##########.00")
    'Else
    '   TXTBNET.Text = Format(Val(TXTITOT.Text) + Val(TXTSAMT.Text) + Val(TXTADLS.Text), "##########.00")
    'End If
    
    TXTBNET.Text = Format(FormatNumber(Val(TXTITOT.Text) + Val(TXTSAMT.Text) + Val(TXTADLS.Text), 0), "##########.00")
    
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
  RS.Open "SELECT * FROM CONFIG WHERE COMP='" & compPth & "' and vtyp='PSR' " & _
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



