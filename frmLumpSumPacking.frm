VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{2BF72F7D-D367-4712-9146-5533EF3E691A}#1.2#0"; "FraPlus1.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WelchButton.ocx"
Begin VB.Form frmLumpSumPacking 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LumpSum Packing"
   ClientHeight    =   6900
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11325
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6883.333
   ScaleMode       =   0  'User
   ScaleWidth      =   11325
   Begin VB.PictureBox picToolTip 
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      FillColor       =   &H00404040&
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   4815
      TabIndex        =   49
      Top             =   6960
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
         ForeColor       =   &H8000000E&
         Height          =   240
         Left            =   120
         TabIndex        =   50
         Top             =   0
         Width           =   120
      End
   End
   Begin FramePlusCtl.FramePlus FramePlus1 
      Height          =   6915
      Left            =   0
      TabIndex        =   26
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
      Begin VB.TextBox TXTPCOD 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   7320
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1200
         Width           =   3615
      End
      Begin VB.TextBox TXTNETWT 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   9720
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   4320
         Width           =   1335
      End
      Begin VB.TextBox TXTTAREWT 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   7680
         TabIndex        =   23
         Top             =   4320
         Width           =   1335
      End
      Begin VB.TextBox TXTGRSWT 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   4920
         TabIndex        =   22
         Top             =   4320
         Width           =   1455
      End
      Begin VB.TextBox TXTCOPS 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   2400
         TabIndex        =   21
         Top             =   4320
         Width           =   1335
      End
      Begin VB.TextBox TXTNOB 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   360
         TabIndex        =   20
         Top             =   4320
         Width           =   1335
      End
      Begin VB.TextBox TXTSUBGRD 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   9480
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   3480
         Width           =   1695
      End
      Begin VB.TextBox TXTGRAD 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   7560
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   3480
         Width           =   1575
      End
      Begin VB.TextBox TXTITEM 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   3480
         Width           =   3135
      End
      Begin VB.TextBox TXTLOTNO 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   3480
         Width           =   1815
      End
      Begin VB.TextBox TXTSLIP 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   3480
         Width           =   1335
      End
      Begin VB.ComboBox cmbPackingType 
         BackColor       =   &H0080C0FF&
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
         Height          =   315
         ItemData        =   "frmLumpSumPacking.frx":0000
         Left            =   2040
         List            =   "frmLumpSumPacking.frx":0002
         TabIndex        =   7
         Tag             =   "0"
         Text            =   "Select Type of Packing"
         Top             =   1200
         Width           =   3255
      End
      Begin VB.TextBox TXTDVNM 
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
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   33
         TabStop         =   0   'False
         Tag             =   "0"
         Top             =   480
         Width           =   3255
      End
      Begin VB.TextBox TXTPKGSTATION 
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
         Left            =   7320
         Locked          =   -1  'True
         TabIndex        =   32
         TabStop         =   0   'False
         Tag             =   "0"
         Top             =   480
         Width           =   3615
      End
      Begin MSComCtl2.DTPicker TXTVBDT 
         Height          =   315
         Left            =   2040
         TabIndex        =   6
         Top             =   840
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
         Format          =   18939905
         CurrentDate     =   39383
      End
      Begin WelchButton.lvButtons_H cmdAdd 
         Height          =   495
         Left            =   1800
         TabIndex        =   0
         Top             =   6120
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
         Image           =   "frmLumpSumPacking.frx":0004
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdEdit 
         Height          =   495
         Left            =   5760
         TabIndex        =   3
         Top             =   6120
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
         Image           =   "frmLumpSumPacking.frx":039E
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdDelete 
         Height          =   495
         Left            =   7080
         TabIndex        =   4
         Top             =   6120
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
         Image           =   "frmLumpSumPacking.frx":0738
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdSave 
         Height          =   495
         Left            =   3120
         TabIndex        =   1
         Top             =   6120
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
         Image           =   "frmLumpSumPacking.frx":0AD2
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdCancel 
         Height          =   495
         Left            =   4440
         TabIndex        =   2
         Top             =   6120
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
         Image           =   "frmLumpSumPacking.frx":185C
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdExit 
         Height          =   495
         Left            =   8400
         TabIndex        =   5
         Top             =   6120
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
         Image           =   "frmLumpSumPacking.frx":1CAE
         cBack           =   -2147483633
      End
      Begin FramePlusCtl.FramePlus FrmAutoConsumption 
         Height          =   975
         Left            =   0
         TabIndex        =   44
         Top             =   1800
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   1720
         BackgroundPictureAlignment=   5
         BorderStyle     =   10
         BackColorGradient=   12640511
         BackColor       =   12640511
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
         Begin VB.ComboBox cmbPackaging 
            BackColor       =   &H0080C0FF&
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
            Height          =   315
            ItemData        =   "frmLumpSumPacking.frx":2100
            Left            =   1680
            List            =   "frmLumpSumPacking.frx":2102
            Sorted          =   -1  'True
            TabIndex        =   12
            Tag             =   "0"
            Text            =   "Select Type of Packaging"
            Top             =   600
            Width           =   2295
         End
         Begin VB.TextBox TXTLOC 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   9000
            Locked          =   -1  'True
            TabIndex        =   14
            Top             =   600
            Width           =   1935
         End
         Begin VB.TextBox TXTMCCD 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   5160
            Locked          =   -1  'True
            TabIndex        =   13
            Top             =   600
            Width           =   2415
         End
         Begin VB.TextBox TXTCOPSNAME 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   9000
            Locked          =   -1  'True
            TabIndex        =   11
            Top             =   120
            Width           =   1935
         End
         Begin VB.TextBox TXTCARTONNAME 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   5160
            Locked          =   -1  'True
            TabIndex        =   10
            Top             =   120
            Width           =   2415
         End
         Begin VB.TextBox TXTIGRP 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   9
            Top             =   120
            Width           =   2295
         End
         Begin VB.Label Label21 
            BackStyle       =   0  'Transparent
            Caption         =   "PackagingType"
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
            TabIndex        =   54
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label Label20 
            BackStyle       =   0  'Transparent
            Caption         =   "StoreLocation"
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
            Left            =   7680
            TabIndex        =   53
            Tag             =   "S"
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label Label19 
            BackStyle       =   0  'Transparent
            Caption         =   "Machine"
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
            Left            =   4200
            TabIndex        =   52
            Tag             =   "S"
            Top             =   600
            Width           =   855
         End
         Begin VB.Label Label18 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Box Type"
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
            Left            =   3960
            TabIndex        =   48
            Tag             =   "S"
            Top             =   120
            Width           =   1215
         End
         Begin VB.Label Label15 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Cops Type"
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
            Left            =   7680
            TabIndex        =   47
            Tag             =   "S"
            Top             =   120
            Width           =   975
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
            TabIndex        =   46
            Tag             =   "S"
            Top             =   -2040
            Width           =   1335
         End
         Begin VB.Label Label13 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Item Group"
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
            TabIndex        =   45
            Tag             =   "S"
            Top             =   120
            Width           =   975
         End
      End
      Begin MSFlexGridLib.MSFlexGrid FLEXPLY 
         Height          =   885
         Left            =   360
         TabIndex        =   25
         Top             =   4800
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   1561
         _Version        =   393216
         Cols            =   50
         BackColor       =   -2147483628
         BackColorBkg    =   15786495
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
      Begin VB.Line Line9 
         BorderColor     =   &H00000080&
         X1              =   120
         X2              =   11280
         Y1              =   5880
         Y2              =   5880
      End
      Begin VB.Label LBLPCOD 
         BackStyle       =   0  'Transparent
         Caption         =   "Job Party Name"
         Enabled         =   0   'False
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
         Left            =   5640
         TabIndex        =   51
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00000080&
         X1              =   120
         X2              =   11280
         Y1              =   4680
         Y2              =   4680
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00000080&
         X1              =   120
         X2              =   11280
         Y1              =   4200
         Y2              =   4200
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Net Weight / BOX"
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
         Left            =   9480
         TabIndex        =   43
         Tag             =   "S"
         Top             =   3960
         Width           =   1695
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Tare Weight / BOX"
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
         Left            =   7560
         TabIndex        =   42
         Tag             =   "S"
         Top             =   3960
         Width           =   1695
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Gross Weight / BOX"
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
         Left            =   4560
         TabIndex        =   41
         Tag             =   "S"
         Top             =   3960
         Width           =   2175
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cops / BOX"
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
         Left            =   2160
         TabIndex        =   40
         Tag             =   "S"
         Top             =   3960
         Width           =   1815
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "No.of BOXES"
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
         TabIndex        =   39
         Tag             =   "S"
         Top             =   3960
         Width           =   1215
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00000080&
         X1              =   120
         X2              =   11280
         Y1              =   3840
         Y2              =   3840
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Sub Grade"
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
         Left            =   9720
         TabIndex        =   38
         Tag             =   "S"
         Top             =   3075
         Width           =   975
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00000080&
         X1              =   9360
         X2              =   9360
         Y1              =   3000
         Y2              =   4680
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
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
         Left            =   5040
         TabIndex        =   37
         Tag             =   "S"
         Top             =   3120
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Lot No."
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
         TabIndex        =   36
         Tag             =   "S"
         Top             =   3120
         Width           =   975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Date of  Packing"
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
         Left            =   240
         TabIndex        =   35
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Type of  Packing"
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
         Left            =   240
         TabIndex        =   34
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label LBLHEAD 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         Caption         =   "LumpSum Packing"
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
         TabIndex        =   31
         Top             =   0
         Width           =   4455
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         Height          =   1575
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   11055
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000080&
         X1              =   4080
         X2              =   4080
         Y1              =   3000
         Y2              =   4680
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00000080&
         X1              =   120
         X2              =   11280
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Packing Station"
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
         Left            =   5640
         TabIndex        =   30
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Division Name"
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
         Left            =   240
         TabIndex        =   29
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label16 
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
         Left            =   7920
         TabIndex        =   28
         Tag             =   "S"
         Top             =   3075
         Width           =   975
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00000080&
         X1              =   2040
         X2              =   2040
         Y1              =   3000
         Y2              =   4680
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         Height          =   3735
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   3000
         Width           =   11175
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Slip No."
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
         Left            =   600
         TabIndex        =   27
         Tag             =   "S"
         Top             =   3120
         Width           =   975
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00000080&
         X1              =   7440
         X2              =   7440
         Y1              =   3000
         Y2              =   4680
      End
   End
End
Attribute VB_Name = "frmLumpSumPacking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public LSDVCD As String
Dim LSPKGCOD As String
Public M_SRNO As String
Dim M_IGCD As String
Public M_DBCD As String
Dim PKGNG_COD As String
Dim SAVEFLAG As Boolean
Dim ALLOWEDITDEL As Boolean
Dim M_PCOD As String

Private Sub cmbPackaging_Click()
   Call SETPLYLIMIT
End Sub

Private Sub cmbPackingType_Click()
  SendKeys "{HOME}"
  If InStr(1, cmbPackingType.Text, "GR ") <> 0 Then    'PARTY REQUIRED in CASE OF JOB
    TXTPCOD.Enabled = True
    LBLPCOD.Enabled = True
  Else
    TXTPCOD = Empty
    TXTPCOD.Enabled = False
    TXTPCOD.BackColor = vbWhite
  End If
End Sub

Private Sub cmbPackingType_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub cmdAdd_Click()
    SAVEFLAG = True
    Dim Ctrl As Control
    For Each Ctrl In Me
        If TypeOf Ctrl Is TextBox Then
            Ctrl = Replace(Ctrl, "'", "", 1)
        End If
    Next
    Call ClsTxt
    Call btn_sts(True)
    cmdCancel.Cancel = True
    
    TXTSLIP = GenPackSlipNo(LSPKGCOD)
    Call SetUnitConfig
End Sub

Private Sub cmdCancel_Click()
    Call CLEARDATA
    TXTDVNM.Tag = TXTDVNM
    TXTPKGSTATION.Tag = TXTPKGSTATION
    cmbPackingType.Tag = cmbPackingType.Text
    cmbPackaging.Tag = cmbPackaging.Text
    ClsData (Me)
    TXTDVNM = TXTDVNM.Tag
    TXTPKGSTATION = TXTPKGSTATION.Tag
    cmbPackingType.Text = cmbPackingType.Tag
    cmbPackaging.Text = cmbPackaging.Tag
    'If cmbPackingType.ListCount > 1 Then cmbPackingType.ListIndex = 2
    
    Call btn_sts(False)
    M_SRNO = Empty
    If zoomflag = True Then
        Call cmdExit_Click
        Exit Sub
    End If
    TXTSLIP = GenPackSlipNo(LSPKGCOD)
    'change
    cmbPackingType.Enabled = True
End Sub

Private Sub CLEARDATA()
 Dim I As Long
 For I = 1 To FLEXPLY.Cols - 1
  FLEXPLY.TextMatrix(1, I) = ""
 Next
End Sub

Private Sub cmddelete_Click()
  cmbPackingType.Enabled = False
  If M_USRSECLEVL = "1" Then
        If ReadConfigMaster("0017", 6, "M") = False Then ModuleDeniedMessage: Exit Sub
  End If
        
  ALLOWEDITDEL = True
  SAVEFLAG = False
  M_SRNO = Empty
  btn_sts (True)
  frmLumpSumList.Show 1
  If ALLOWEDITDEL = False Then
    MsgBox "Purchase of this GRN have been made can not edit/delete ", vbInformation
   Else
    'Check for Receipt and Payment Entires
    If Not M_SRNO = Empty Then
      Dim AYS
      AYS = MsgBox("Are you sure to delete this Packing Slip", vbYesNo)
      If AYS = vbYes Then
        CN.BeginTrans
        CN.Execute "UPDATE PKGMAN SET RECSTAT='D' WHERE COMP='" & compPth & "' AND UNIT ='" & UNCD & "' AND VTYP='PPF' AND LTRIM(RTRIM(SRNO))='" & M_SRNO & "'"
        CN.Execute "UPDATE STORETRAN SET RECSTAT='D' WHERE COMP='" & compPth & "' AND VTYP='PPF' AND SRNO='" & M_SRNO & "'"
        'Call UPDATESTATUS
        Call DAILYSTATUS("PPF", GetCode("ACCMST", TXTPCOD, "NAME", "CODE"), M_DBCD, Val(TXTNOB) * Val(TXTNETWT), TXTSLIP, 0, cUName, "D", Now, TXTVBDT)
        CN.CommitTrans
      End If
    End If
  End If
  Call cmdCancel_Click
End Sub

Private Sub cmdEdit_Click()
  If M_USRSECLEVL = "1" Then
     If ReadConfigMaster("0017", 5, "M") = False Then ModuleDeniedMessage: Exit Sub
  End If
    SAVEFLAG = False
    Call SetGlobal
    frmLumpSumList.Show 1
    If M_SRNO <> Empty Then
        btn_sts (True)
        TXTLOTNO.SetFocus
        cmbPackingType.Enabled = False
    Else
        btn_sts (False)
        Call cmdCancel_Click
        If cmdAdd.Enabled Then cmdAdd.SetFocus
    End If
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
On Error GoTo LAST
 
 If CHKSAVEDATA = True Then
    Exit Sub
 End If
  
 Call SetGlobal
  
 'Generate Sr. No.
 If M_SRNO = Empty Then
    M_SRNO = pubGenSrNoPKGMAN(TXTVBDT, "PPF")
 End If
 
 If SAVEFLAG = True Then
    TXTSLIP = GenPackSlipNo(LSPKGCOD)
 End If
 
 Dim SAVDAT As ADODB.Recordset
 Set SAVDAT = New ADODB.Recordset
 If SAVDAT.State = 1 Then SAVDAT.Close
 SAVDAT.Open "SELECT * FROM PKGMAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
 "' AND DVCD='" & LSDVCD & "' AND SLIPNO='" & TXTSLIP & "' AND DBCD='" & M_DBCD & "' AND VTYP='PPF' AND RECSTAT<>'D' AND PKG_STCOD ='" & LSPKGCOD & "'", CN, adOpenDynamic, adLockOptimistic
 If Not SAVDAT.EOF Then
    If SAVDAT!SRNO = M_SRNO Then
     Else
      MsgBox "Duplicate Packing No."
      cmdSave.SetFocus
      Exit Sub
    End If
 End If
 
 Call SAVERECPPF
 
 If SAVEFLAG = True Then
    MsgBox "Your Packing Slip No. is " + TXTSLIP.Text
 End If
 
 Call cmdCancel_Click
 Exit Sub
 
Exit Sub
LAST:
MsgBox ERR.Description
Exit Sub
End Sub

Private Sub SAVERECPPF()
  On Error GoTo LAST
  Dim SQL As String
  Dim I As Long, J As Long
  
  Dim SAVDAT As New ADODB.Recordset
  Set SAVDAT = New ADODB.Recordset
       
  CN.BeginTrans
  Call DELETEPPF
  SQL = Empty
  If SAVDAT.State = 1 Then SAVDAT.Close
  SAVDAT.Open "SELECT * FROM PKGMAN WHERE COMP='" & compPth & "' AND VTYP='PPF' AND SRNO='" & M_SRNO & "'", CN, adOpenDynamic, adLockOptimistic
  If SAVDAT.EOF Then
      SAVDAT.AddNew
  End If
  
  SAVDAT!COMP = compPth
  SAVDAT!unit = UNCD
  SAVDAT!DVCD = LSDVCD
  SAVDAT!dbcd = M_DBCD
  SAVDAT!VTYP = "PPF"
  SAVDAT!SRNO = M_SRNO
  SAVDAT!SRCH = 1
  SAVDAT!Date = Format(TXTVBDT.Value, "YYYY/MM/DD")
  SAVDAT!SLIPNO = Trim(TXTSLIP.Text)
  SAVDAT!PKG_STCOD = LSPKGCOD
  SAVDAT!PKGNG_COD = PKGNG_COD
  SAVDAT!PCOD = M_PCOD
  
  SAVDAT!BOX_COD = GetCode("ITMMST", TXTCARTONNAME, "NAME", "CODE")
  SAVDAT!COPS_COD = GetCode("ITMMST", TXTCOPSNAME, "NAME", "CODE")
  SAVDAT!LOTNO = TXTLOTNO
  SAVDAT!MCCD = FindMachineCode
  SAVDAT!LOCCOD = GetCode("LOCMST", TXTLOC, "NAME", "CODE")
  SAVDAT!FINITMCOD = FindFinItemCode
  SAVDAT!grad = GetCode("GRDMST", TXTGRAD, "GRAD", "CODE")
  SAVDAT!SUBGRAD = FindSubGradeCode(GetCode("GRDMST", TXTGRAD, "GRAD", "CODE"), TXTSUBGRD)
  SAVDAT!NOB = Val(TXTNOB)
  SAVDAT!CPB = Val(txtCops)
  SAVDAT!GWPB = Val(TXTGRSWT)
  SAVDAT!TWPB = Val(TXTTAREWT)
  SAVDAT!NWPB = Val(TXTNETWT)
  SAVDAT!QNTY = Val(TXTNOB) * Val(TXTNETWT)
  
  If SAVEFLAG = True Then
    SAVDAT!SYSR = "N"
   Else
    SAVDAT!SYSR = "U"
  End If
  
  SAVDAT![User] = cUName & ""
  SAVDAT!OPER = "+"
  SAVDAT!RECSTAT = "A"
  
  
'PLY UPDATION COMMON FOR BOTH SAVE AND EDIT
If FLEXPLY.Enabled Then

SAVDAT![Top] = 1
SAVDAT!Bottom = 1

I = 0
  For I = 1 To FLEXPLY.Cols - 1
    J = 0
    For J = 0 To SAVDAT.Fields.COUNT - 1
      If Trim(SAVDAT.Fields(J).NAME) = Trim(FLEXPLY.TextMatrix(0, I)) Then
         SAVDAT.Fields(J).Value = Val(FLEXPLY.TextMatrix(1, I))
      End If
    Next
  Next
End If
'----------------------------------------------
 
  SAVDAT.Update
   
  Call SetRawMaterial
   
  'UPDATE VOUCHER TYPE MASTER
  If SAVEFLAG = True Then
     CN.Execute "UPDATE PCKMST SET [LBNO]='" & TXTSLIP & "' WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND CODE='" & LSPKGCOD & "'"
  End If
  '------------------------------
  'DAILYSTATUS ENTRY
   If SAVEFLAG = True Then
      Call DAILYSTATUS("PPF", M_PCOD, M_DBCD, Val(TXTNOB) * Val(TXTNETWT), TXTSLIP, 0, cUName, "N", Now, TXTVBDT)
     Else
      Call DAILYSTATUS("PPF", M_PCOD, M_DBCD, Val(TXTNOB) * Val(TXTNETWT), TXTSLIP, 0, cUName, "M", Now, TXTVBDT)
    End If
   '-----------------------------
  
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

Private Sub Form_Activate()
Me.BackColor = RGB(RED, GREEN, BLUE)
If TXTPKGSTATION = Empty Then Unload Me: Exit Sub
    If zoomflag = True Then
        btn_sts (True)
        SAVEFLAG = False
    Else
        btn_sts (False)
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
 If UCase(ActiveControl.NAME) = "TXTIGRP" And KeyAscii = vbKeyReturn And txtIGRP = Empty Then Exit Sub
 If UCase(ActiveControl.NAME) = "FLEXPLY" Then Exit Sub
 If UCase(ActiveControl.NAME) = "TXTNETWT" Then Exit Sub
 If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
 End If
End Sub

Private Sub Form_Load()
  Call CenterChild(frm_Main, Me)
  Call ColorComponent(Me)
  Me.Left = Me.Left - 100
  'Me.Top = Me.Top + 750
  Me.KeyPreview = True

'-------DIVISION NAME
  M_DESC = Empty: Key = Empty:  NEW_VISIBLE = False
  LSDVCD = Empty
  TXTDVNM = Empty
  If LSDVCD = Empty Then
    TXTDVNM = SearchList1("SELECT  TOP 20 CODE,NAME FROM DIVMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND RECSTAT='A'  AND CODE<>'000001'", 0, "", "SELECT DIVISION MASTER")
    If TXTDVNM <> Empty Then LSDVCD = Key Else TXTDVNM = "???????": Unload Me
  End If
  
 If PackingType(Key) = "C" Then MsgBox "Division Not Allowed Lumpsum Packing.Check Configuration": GoTo JUMP

'-------PACKING STATION MASTER
M_DESC = Empty:  Key = Empty:  NEW_VISIBLE = False
LSPKGCOD = Empty
TXTPKGSTATION = SearchList1("SELECT TOP 20 CODE,NAME FROM PCKMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' ", 0, TXTPKGSTATION, "SELECT PACKING STATION FROM MASTER LIST")
If Key = Empty Then Exit Sub
LSPKGCOD = Key
'---------------------------
       
TXTVBDT.Value = Now
TXTVBDT.MinDate = FSDT
TXTVBDT.MaxDate = FEDT

Call SetPackingType
Call setHeading
JUMP:
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
     cmdSave.SetFocus
  End If
  
  End If   'END<<-----------------------
      
  Exit Sub
LAST:
  MsgBox ERR.Description
  Exit Sub
End Sub

Private Sub FLEXPLY_LeaveCell()
FLEXPLY.CellBackColor = vbWhite
End Sub

Private Sub SetPackingType()
Dim PKTYPRS As ADODB.Recordset
Set PKTYPRS = New ADODB.Recordset
If PKTYPRS.State = 1 Then PKTYPRS.Close
PKTYPRS.Open "SELECT NAME FROM SERIALMASTER WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
"' AND VTYP='PPF' AND NAME NOT LIKE '%WASTAGE%' AND FYCD='" & FYCD & "'", CN, adOpenDynamic, adLockOptimistic

Do While Not PKTYPRS.EOF
 cmbPackingType.AddItem Trim(PKTYPRS!NAME)
 PKTYPRS.MoveNext
Loop

If cmbPackingType.ListCount > 1 Then cmbPackingType.ListIndex = 2

If PKTYPRS.State = 1 Then PKTYPRS.Close
PKTYPRS.Open "SELECT NAME FROM PKGNGMST WHERE STATUS='A' AND RECSTAT='A'", CN, adOpenDynamic, adLockOptimistic
Do While Not PKTYPRS.EOF
 cmbPackaging.AddItem Trim(PKTYPRS!NAME)
PKTYPRS.MoveNext
Loop
If cmbPackaging.ListCount > 1 Then cmbPackaging.ListIndex = 0

End Sub

Private Sub cmbPackaging_KeyPress(KeyAscii As Integer): KeyAscii = 0: End Sub

Private Sub cmbPackaging_KeyDown(KeyCode As Integer, Shift As Integer)
   Call SETPLYLIMIT
End Sub

Private Sub Form_Unload(Cancel As Integer)
Msg ""
End Sub

Private Sub TXTCARTONNAME_KeyDown(KeyCode As Integer, Shift As Integer)
Me.KeyPreview = False
If KeyCode = vbKeyF2 Or (Trim(TXTCARTONNAME) = Empty And KeyCode = vbKeyReturn) Then
   NEW_VISIBLE = False:   Key = Empty
   
   If Trim(M_IGCD) <> Empty Then
    TXTCARTONNAME.Text = SearchITEMLIST("SELECT TOP 20 CODE, NAME FROM ITMMST WHERE IGCD='" & Trim(M_IGCD) & "'", 0, TXTCARTONNAME, "SELECT ITEM FROM LIST")
   Else
    TXTCARTONNAME.Text = SearchITEMLIST("SELECT TOP 20 CODE, NAME FROM ITMMST", 0, TXTCARTONNAME, "SELECT ITEM FROM LIST")
   End If
   
   If key_PressNew = True Then
      M_DESC = "": Key = "": TXTCARTONNAME.Text = ""
      frm_Item.Show
   Else
      TXTCARTONNAME.Tag = Key
   End If
End If
Me.KeyPreview = True
End Sub

Private Sub txtcops_KeyPress(KeyAscii As Integer)
  If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub TXTCOPSNAME_KeyDown(KeyCode As Integer, Shift As Integer)
Me.KeyPreview = False
If KeyCode = vbKeyF2 Or (Trim(TXTCOPSNAME) = Empty And KeyCode = vbKeyReturn) Then
   NEW_VISIBLE = False:   Key = Empty
   
   If Trim(M_IGCD) <> Empty Then
    TXTCOPSNAME.Text = SearchITEMLIST("SELECT TOP 20 CODE, NAME FROM ITMMST WHERE IGCD='" & Trim(M_IGCD) & "'", 0, TXTCOPSNAME, "SELECT ITEM FROM LIST")
   Else
    TXTCOPSNAME.Text = SearchITEMLIST("SELECT TOP 20 CODE, NAME FROM ITMMST", 0, TXTCOPSNAME, "SELECT ITEM FROM LIST")
   End If
   
   If key_PressNew = True Then
      M_DESC = "": Key = "": TXTCOPSNAME.Text = ""
      frm_Item.Show
   Else
      TXTCOPSNAME.Tag = Key
   End If
End If
Me.KeyPreview = True
End Sub

Private Sub TXTGRAD_KeyDown(KeyCode As Integer, Shift As Integer)
  If Trim(TXTGRAD.Text) = Empty Or KeyCode = vbKeyF2 Then
    TXTGRAD.Text = SearchList1("select TOP 20 grad as grade,grad from grdmst", 0, TXTGRAD, "SELECT GRAD FROM MASTER")
  End If
End Sub


Private Sub TXTGRSWT_KeyPress(KeyAscii As Integer)
 If CheckNumericKey(KeyAscii, TXTGRSWT, Me) = 0 Then KeyAscii = 0
End Sub

Private Sub TXTIGRP_KeyDown(KeyCode As Integer, Shift As Integer)
    M_DESC = Empty:    Key = Empty: sTxt = "": NEW_VISIBLE = False
    If KeyCode = vbKeyF2 Or (KeyCode = vbKeyReturn And txtIGRP = Empty) Then
        txtIGRP.Text = SearchList1("select TOP 20 code, name from IGMMST", 0, "", "List Of Item Group")
        M_IGCD = Key
        txtIGRP.Tag = Key
    End If
End Sub

Private Sub TXTLOC_GotFocus()
  TXTLOC.BackColor = RGB(BRED, BGREEN, BBLUE)
  SendKeys "{HOME}+{END}"
  ToolTip Me, "Press {F2} / {Enter} For Location Master Help", "", TXTLOC.Left - 620, TXTLOC.Top + TXTLOC.Height + 100
End Sub

Private Sub TXTLOC_KeyDown(KeyCode As Integer, Shift As Integer)
  If (KeyCode = vbKeyReturn And Trim(TXTLOC.Text) = Empty) Or KeyCode = vbKeyF2 Then
    TXTLOC.Text = SearchList1("SELECT  TOP 20 CODE,NAME FROM LOCMST", 0, TXTLOC, "SELECT LOCATION FROM MASTER")
  End If
End Sub

Private Sub TXTLOC_LostFocus(): TXTLOC.BackColor = vbWhite: picToolTip.Visible = False: End Sub

Private Sub TXTLOTNO_KeyDown(KeyCode As Integer, Shift As Integer)
Me.KeyPreview = False
Dim SQL As String

If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
     TXTLOTNO = Empty
ElseIf KeyCode = vbKeyF2 Or (KeyCode = vbKeyReturn And TXTLOTNO = Empty) Then
   M_DESC = Empty:   NEW_VISIBLE = False: Key = Empty
   SQL = "SELECT DISTINCT LTNO,LTNO FROM TXULOT WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & LSDVCD & "' AND ACTIVE='Y' "
   TXTLOTNO = SearchList(SQL)
End If
If TXTLOTNO <> Empty Then FindFinishItem
Me.KeyPreview = True
End Sub

Private Sub TXTMCCD_GotFocus()
  TXTMCCD.BackColor = RGB(BRED, BGREEN, BBLUE)
  SendKeys "{HOME}+{END}"
  'If COUNTER > 0 Then Exit Sub
  ToolTip Me, "Press {F2} / {Enter} For Machine Master Help", "", TXTMCCD.Left - 620, TXTMCCD.Top + TXTMCCD.Height + 100
End Sub

Private Sub TXTMCCD_KeyDown(KeyCode As Integer, Shift As Integer)
Me.KeyPreview = False
   If KeyCode = vbKeyF2 Or Trim(TXTMCCD.Text) = Empty Then
        NEW_VISIBLE = False:  M_DESC = Empty:   Key = Empty
        TXTMCCD.Text = SearchList1("SELECT  TOP 20 CODE,NAME FROM MACMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & LSDVCD & "'", 0, TXTMCCD, "List of Machine Name")
   ElseIf KeyCode = vbKeyDelete Then
        TXTMCCD = Empty
   End If
Me.KeyPreview = True
End Sub

Private Sub TXTMCCD_LostFocus(): TXTMCCD.BackColor = vbWhite: picToolTip.Visible = False: End Sub

Private Sub TXTNETWT_KeyPress(KeyAscii As Integer)
 If CheckNumericKey(KeyAscii, TXTNETWT, Me) = 0 Then KeyAscii = 0
 
 If KeyAscii = 13 And Val(TXTNETWT) <> 0 Then
    If FLEXPLY.Enabled Then
       FLEXPLY.ROW = 1
       FLEXPLY.COL = 1
       FLEXPLY.SetFocus
    Else
       cmdSave.Enabled = True: cmdSave.SetFocus
    End If
 End If
End Sub

Private Sub TXTNOB_KeyPress(KeyAscii As Integer)
  If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub txtPCOD_GotFocus()
    TXTPCOD.BackColor = RGB(BRED, BGREEN, BBLUE)
    SendKeys "{HOME}+{END}"
End Sub

Private Sub TXTPCOD_KeyDown(KeyCode As Integer, Shift As Integer)
 Me.KeyPreview = False
   If KeyCode = vbKeyF2 Or Trim(TXTPCOD.Text) = Empty Then
        
        NEW_VISIBLE = False
        M_DESC = Empty
        Key = Empty
        TXTPCOD.Text = SearchList1("SELECT  TOP 20 CODE,NAME FROM ACCMST", 0, TXTPCOD, "List of Job Party A/c")
   ElseIf KeyCode = vbKeyDelete Then
        TXTPCOD = Empty
    End If
    Me.KeyPreview = True
End Sub

Private Sub TXTSUBGRD_KeyDown(KeyCode As Integer, Shift As Integer)
Me.KeyPreview = False
Dim SQL As String
If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
     TXTSUBGRD = Empty
ElseIf KeyCode = vbKeyF2 Or (KeyCode = vbKeyReturn And TXTSUBGRD = Empty) Then
   M_DESC = Empty:   NEW_VISIBLE = False
   SQL = "SELECT DISTINCT RDIFF,NAME FROM SUBGRDMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & LSDVCD & "' AND GRAD='" & GetCode("GRDMST", TXTGRAD, "GRAD", "CODE") & "'"
   TXTSUBGRD = SearchList1(SQL, 0, Empty)
End If
Me.KeyPreview = True
End Sub

Private Sub TXTTAREWT_KeyPress(KeyAscii As Integer)
 If CheckNumericKey(KeyAscii, TXTTAREWT, Me) = 0 Then KeyAscii = 0
 TXTNETWT = nstr(Val(TXTGRSWT) - Val(TXTTAREWT), 12, 3)
End Sub

Private Sub TXTVBDT_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
   SendKeys "{TAB}"
End If
End Sub

''----------------------------------------------------------------------------------------
' GRAPHICS
'-----------------------------------------------------------------------------------------
Private Sub TXTSLIP_GotFocus()
    TXTSLIP.BackColor = RGB(BRED, BGREEN, BBLUE)
    SendKeys "{HOME}+{END}"
End Sub

Private Sub TXTSLIP_LostFocus()
    TXTSLIP.BackColor = vbWhite
End Sub

Private Sub TXTLOTNO_GotFocus()
    TXTLOTNO.BackColor = RGB(BRED, BGREEN, BBLUE)
    SendKeys "{HOME}+{END}"
    If TXTLOTNO = Empty Then
      ToolTip Me, "Press {F2} / {Enter} For Lot Master Help", "", TXTLOTNO.Left, TXTLOTNO.Top + TXTLOTNO.Height + 100
    Else
      ToolTip Me, "Press {F2} For Lot Master Help", "", TXTLOTNO.Left, TXTLOTNO.Top + TXTLOTNO.Height + 100
    End If
End Sub

Private Sub TXTLOTNO_LostFocus()
    TXTLOTNO.BackColor = vbWhite
    picToolTip.Visible = False
End Sub

Private Sub txtItem_GotFocus()
    TXTITEM.BackColor = RGB(BRED, BGREEN, BBLUE)
    SendKeys "{HOME}+{END}"
End Sub

Private Sub txtItem_LostFocus()
    TXTITEM.BackColor = vbWhite
End Sub

Private Sub TXTGRAD_GotFocus()
    TXTGRAD.BackColor = RGB(BRED, BGREEN, BBLUE)
    SendKeys "{HOME}+{END}"
    If TXTGRAD = Empty Then
      ToolTip Me, "Press {F2} / {Enter} For Grade Help", "", TXTGRAD.Left, TXTGRAD.Top + TXTGRAD.Height + 100
    Else
      ToolTip Me, "Press {F2} For Grade Help", "", TXTGRAD.Left, TXTGRAD.Top + TXTGRAD.Height + 100
    End If
End Sub

Private Sub TXTGRAD_LostFocus()
    TXTGRAD.BackColor = vbWhite
    picToolTip.Visible = False
End Sub

Private Sub TXTSUBGRD_GotFocus()
    TXTSUBGRD.BackColor = RGB(BRED, BGREEN, BBLUE)
    SendKeys "{HOME}+{END}"
    If TXTSUBGRD = Empty Then
      ToolTip Me, "Press {F2} / {Enter} For SubGrade Help", "", TXTSUBGRD.Left - 2500, TXTSUBGRD.Top + TXTSUBGRD.Height + 100
    Else
      ToolTip Me, "Press {F2} For SubGrade Help", "", TXTSUBGRD.Left - 2500, TXTSUBGRD.Top + TXTSUBGRD.Height + 100
    End If
End Sub

Private Sub TXTSUBGRD_LostFocus()
    TXTSUBGRD.BackColor = vbWhite
    picToolTip.Visible = False
End Sub

Private Sub TXTNOB_GotFocus()
    TXTNOB.BackColor = RGB(BRED, BGREEN, BBLUE)
    SendKeys "{HOME}+{END}"
End Sub

Private Sub TXTNOB_LostFocus()
    TXTNOB.BackColor = vbWhite
End Sub

Private Sub txtCops_GotFocus()
    txtCops.BackColor = RGB(BRED, BGREEN, BBLUE)
    SendKeys "{HOME}+{END}"
End Sub

Private Sub txtCops_LostFocus()
    txtCops.BackColor = vbWhite
End Sub

Private Sub TXTGRSWT_GotFocus()
    TXTGRSWT.BackColor = RGB(BRED, BGREEN, BBLUE)
    SendKeys "{HOME}+{END}"
End Sub

Private Sub TXTGRSWT_LostFocus()
    TXTGRSWT.BackColor = vbWhite
End Sub

Private Sub TXTTAREWT_GotFocus()
    TXTTAREWT.BackColor = RGB(BRED, BGREEN, BBLUE)
    SendKeys "{HOME}+{END}"
End Sub

Private Sub TXTTAREWT_LostFocus()
    TXTTAREWT.BackColor = vbWhite
End Sub

Private Sub TxtNetWt_GotFocus()
    TXTNETWT.BackColor = RGB(BRED, BGREEN, BBLUE)
    SendKeys "{HOME}+{END}"
End Sub

Private Sub TxtNetWt_LostFocus()
    TXTNETWT.BackColor = vbWhite
End Sub

Private Sub txtIGRP_GotFocus()
    txtIGRP.BackColor = RGB(BRED, BGREEN, BBLUE)
    SendKeys "{HOME}+{END}"
    
    If txtIGRP = Empty Then
      ToolTip Me, "Press {F2} / {Enter} For Item Group Help", "", txtIGRP.Left, txtIGRP.Top + (2 * txtIGRP.Height) + 1500
    Else
      ToolTip Me, "Press {F2} For Item Group Help", "", txtIGRP.Left, txtIGRP.Top + (2 * txtIGRP.Height) + 1500
    End If
End Sub

Private Sub txtIGRP_LostFocus()
    txtIGRP.BackColor = vbWhite
    picToolTip.Visible = False
End Sub

Private Sub TXTCARTONNAME_GotFocus()
    TXTCARTONNAME.BackColor = RGB(BRED, BGREEN, BBLUE)
    If TXTCARTONNAME = Empty Then
      ToolTip Me, "Press {F2} / {Enter} For Cartoon Master Help", "", TXTCARTONNAME.Left, TXTCARTONNAME.Top + (2 * TXTCARTONNAME.Height) + 1500
    Else
      ToolTip Me, "Press {F2} For Cartoon Master Help", "", TXTCARTONNAME.Left, TXTCARTONNAME.Top + (2 * TXTCARTONNAME.Height) + 1500
    End If
End Sub

Private Sub TXTCARTONNAME_LostFocus()
    TXTCARTONNAME.BackColor = vbWhite
    picToolTip.Visible = False
End Sub

Private Sub TXTCOPSNAME_GotFocus()
    TXTCOPSNAME.BackColor = RGB(BRED, BGREEN, BBLUE)
    SendKeys "{HOME}+{END}"
    If TXTCOPSNAME = Empty Then
      ToolTip Me, "Press {F2} / {Enter} For Cops Master Help", "", TXTCOPSNAME.Left - 2500, TXTCOPSNAME.Top + (2 * TXTCOPSNAME.Height) + 1500
    Else
      ToolTip Me, "Press {F2} For Cops Master Help", "", TXTCOPSNAME.Left - 2500, TXTCOPSNAME.Top + (2 * TXTCOPSNAME.Height) + 1500
    End If
End Sub

Private Sub TXTCOPSNAME_LostFocus()
    TXTCOPSNAME.BackColor = vbWhite
    picToolTip.Visible = False
End Sub

'-------------------------------------------------------------------------------------
'Local Procedure
'-------------------------------------------------------------------------------------

Private Sub ClsTxt()
    TXTSLIP = Empty
    TXTLOTNO = Empty
    TXTITEM = Empty
    TXTGRAD = Empty
    TXTSUBGRD = Empty
    TXTNOB = Empty
    txtCops = Empty
    TXTGRSWT = Empty
    TXTTAREWT = Empty
    TXTNETWT = Empty
End Sub

Public Sub btn_sts(bool As Boolean)
    cmdSave.Enabled = bool
    cmdCancel.Enabled = bool
    cmdAdd.Enabled = Not bool
    cmdEdit.Enabled = Not bool
    cmdDelete.Enabled = Not bool
    
    'cmbPackingType.Enabled = bool
    TXTLOTNO.Enabled = bool
    TXTITEM.Enabled = bool
    TXTGRAD.Enabled = bool
    TXTSUBGRD.Enabled = bool
    TXTNOB.Enabled = bool
    txtCops.Enabled = bool
    TXTGRSWT.Enabled = bool
    TXTTAREWT.Enabled = bool
    TXTNETWT.Enabled = bool
    txtIGRP.Enabled = bool
    TXTCARTONNAME.Enabled = bool
    TXTCOPSNAME.Enabled = bool
    
End Sub

Private Sub SetUnitConfig()
Dim FLAG As Boolean

'if autoconsumption then
' FLAG = FALSE
'Else
'FLAG = True
'END IF

 txtIGRP.Enabled = Not FLAG
 TXTCARTONNAME.Enabled = Not FLAG
 TXTCOPSNAME.Enabled = Not FLAG
 FrmAutoConsumption.Visible = Not FLAG
 
 If txtIGRP.Enabled Then
    txtIGRP.SetFocus
 Else
   TXTLOTNO.SetFocus
 End If
 
End Sub

Private Sub FindFinishItem()
Dim RSITM As ADODB.Recordset
Set RSITM = New ADODB.Recordset
Dim FICD As String

If RSITM.State = 1 Then RSITM.Close
RSITM.Open "SELECT FICD FROM TXULOT WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & LSDVCD & "' AND LTNO='" & TXTLOTNO & "'", CN, adOpenDynamic, adLockOptimistic
If Not RSITM.EOF Then
   FICD = RSITM!FICD
End If

RSITM.Close

If FICD <> Empty Then
  If RSITM.State = 1 Then RSITM.Close
  RSITM.Open "SELECT NAME FROM FINITMMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & LSDVCD & "' AND CODE='" & FICD & "'", CN, adOpenDynamic, adLockOptimistic
  If Not RSITM.EOF Then
     TXTITEM = RSITM!NAME
  End If
  RSITM.Close
End If
End Sub

Private Sub tmrTool_Timer()
    ' After 5 seconds, hide tooltip
    picToolTip.Visible = False
End Sub

Private Function CHKSAVEDATA() As Boolean

  If TXTMCCD = Empty Then
     MsgBox "Machine Not Defined !!!! ", vbCritical
     If TXTMCCD.Enabled Then TXTMCCD.SetFocus
     CHKSAVEDATA = True
     Exit Function
  End If
 
 If TXTLOC = Empty Then
     MsgBox "Machine Not Defined !!!! ", vbCritical
     If TXTLOC.Enabled Then TXTLOC.SetFocus
     CHKSAVEDATA = True
     Exit Function
  End If
 
  If TXTCARTONNAME = Empty Then
     MsgBox "Carton Not Defined !!!! ", vbCritical
     If TXTCARTONNAME.Enabled Then TXTCARTONNAME.SetFocus
     CHKSAVEDATA = True
     Exit Function
  End If
    
  If TXTCOPSNAME = Empty Then
     MsgBox "Cops Name Not Defined !!!! ", vbCritical
     If TXTCOPSNAME.Enabled Then TXTCOPSNAME.SetFocus
     CHKSAVEDATA = True
     Exit Function
  End If
    
  If TXTLOTNO = Empty Then
     MsgBox "LotNo. Not Defined !!!! ", vbCritical
     If TXTLOTNO.Enabled Then TXTLOTNO.SetFocus
     CHKSAVEDATA = True
     Exit Function
  End If
  
  If TXTITEM = Empty Then
     MsgBox "Item Not Defined !!!! ", vbCritical
     If TXTITEM.Enabled Then TXTITEM.SetFocus
     CHKSAVEDATA = True
     Exit Function
  End If
  
  If TXTGRAD = Empty Then
     MsgBox "Grade Not Defined !!!! ", vbCritical
     If TXTGRAD.Enabled Then TXTGRAD.SetFocus
     CHKSAVEDATA = True
     Exit Function
  End If
  
  If TXTSUBGRD = Empty Then
     MsgBox "Sub Grade Not Defined !!!! ", vbCritical
     If TXTSUBGRD.Enabled Then TXTSUBGRD.SetFocus
     CHKSAVEDATA = True
     Exit Function
  End If
    
  If Val(TXTNOB) = 0 Then
     MsgBox "Number of Boxes Not Defined!!!! ", vbCritical
     If TXTNOB.Enabled Then TXTNOB.SetFocus
     CHKSAVEDATA = True
     Exit Function
  End If
    
  If cmbPackingType.Text = Empty Then
    MsgBox "Please Select Packing Type !!", vbCritical
    If cmbPackingType.Enabled Then cmbPackingType.SetFocus
    CHKSAVEDATA = True
    Exit Function
  End If

  M_PCOD = Empty
  If InStr(1, cmbPackingType.Text, "JobWork") <> 0 And TXTPCOD = Empty Then    'PARTY REQUIRED in CASE OF JOB
    MsgBox "Party Required in case of JobWork !!", vbCritical
    If TXTPCOD.Enabled Then TXTPCOD.SetFocus
    CHKSAVEDATA = True
    Exit Function
  Else
    M_PCOD = GetCode("ACCMST", TXTPCOD, "NAME", "CODE")
  End If
  
If FLEXPLY.Enabled Then
Dim I As Long, TOTPLY As Long: TOTPLY = 0
    For I = 1 To FLEXPLY.Cols - 1
        TOTPLY = TOTPLY + Val(FLEXPLY.TextMatrix(1, I))
    Next I
    
    If TOTPLY <> Val(FLEXPLY.Tag) Then
       MsgBox "Please Enter Exact No. of Ply (" & CStr(Val(FLEXPLY.Tag)) & ") that r defined in Packaging Master!!", vbInformation
       FLEXPLY.SetFocus
       If FLEXPLY.Rows > 1 Then FLEXPLY.ROW = 1
       If FLEXPLY.Cols > 1 Then FLEXPLY.COL = 1
       CHKSAVEDATA = True
       Exit Function
    End If
End If
 
  CHKSAVEDATA = False
End Function


Private Sub DELETEPPF()
  CN.Execute "DELETE FROM PKGMAN WHERE COMP='" & compPth & "' AND UNIT ='" & UNCD & "' AND VTYP='PPF' AND LTRIM(RTRIM(SRNO))='" & M_SRNO & "'"
  CN.Execute "DELETE FROM STORETRAN WHERE COMP='" & compPth & "' AND UNIT ='" & UNCD & "' AND VTYP='PPF' AND SRNO='" & M_SRNO & "'"
  CN.Execute "DELETE FROM JOBIN WHERE COMP='" & compPth & "' AND UNIT ='" & UNCD & "' AND VTYP='PPF' AND SRNO='" & M_SRNO & "'"
End Sub

Private Sub UPDATESTATUS()
    Dim DLYSTA As New ADODB.Recordset
    If DLYSTA.State = adStateOpen Then DLYSTA.Close
    DLYSTA.Open "SELECT * FROM DAILYSTAT WHERE CUSR='JNMOPAWCGBDSXZAS'", CN, adOpenKeyset, adLockPessimistic
    DLYSTA.AddNew
    DLYSTA!COMP = compPth & ""
    DLYSTA!VTYP = "PPF"
    DLYSTA!PCOD = ""
    DLYSTA!dbcd = ""
    DLYSTA!QNTY = Val(TXTNETWT)
    DLYSTA!VBNO = TXTSLIP & ""
    DLYSTA!AMNT = 0
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
  DLYSTA!VTYP = "PPF"
  DLYSTA!PCOD = ""
  DLYSTA!dbcd = ""
  DLYSTA!QNTY = Val(TXTNETWT)
  DLYSTA!VBNO = TXTSLIP & ""
  DLYSTA!AMNT = 0
  DLYSTA!CUSR = cUName
  DLYSTA!ACTN = "D"
  DLYSTA!DTTM = Format(Now, "YYYY/MM/DD HH:MM:SS AMPM")
  DLYSTA.Update
  DLYSTA.Close
End Sub

Private Function FindFinItemCode() As String
Dim GRRS As ADODB.Recordset
Set GRRS = New ADODB.Recordset

If GRRS.State = 1 Then GRRS.Close
GRRS.Open "SELECT * FROM FINITMMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & LSDVCD & "' AND NAME ='" & TXTITEM & "'", CN, adOpenDynamic, adLockOptimistic
If Not GRRS.EOF Then
   FindFinItemCode = GRRS!CODE
Else
   FindFinItemCode = Empty
End If
GRRS.Close
End Function

Private Function FindMachineCode() As String
Dim GRRS As ADODB.Recordset
Set GRRS = New ADODB.Recordset

If GRRS.State = 1 Then GRRS.Close
GRRS.Open "SELECT * FROM MACMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & LSDVCD & "' AND NAME ='" & TXTMCCD & "'", CN, adOpenDynamic, adLockOptimistic
If Not GRRS.EOF Then
   FindMachineCode = GRRS!CODE
Else
   FindMachineCode = Empty
End If
GRRS.Close
End Function

Private Function FindSubGradeCode(GRADCOD As String, SUBGRD As String) As String
Dim GRRS As ADODB.Recordset
Set GRRS = New ADODB.Recordset

If GRRS.State = 1 Then GRRS.Close
GRRS.Open "SELECT * FROM SUBGRDMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
"' AND DVCD='" & LSDVCD & "' AND GRAD ='" & GRADCOD & "' AND NAME = '" & SUBGRD & "'", CN, adOpenDynamic, adLockOptimistic

If Not GRRS.EOF Then
   FindSubGradeCode = Trim(GRRS!SUBGRD)
Else
   FindSubGradeCode = Empty
End If

GRRS.Close
End Function

Private Sub SetGlobal()
Dim DBCDRS As ADODB.Recordset
Set DBCDRS = New ADODB.Recordset
If DBCDRS.State = 1 Then DBCDRS.Close
DBCDRS.Open "SELECT CODE FROM SERIALMASTER WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
"' AND VTYP='PPF' AND NAME = '" & cmbPackingType.Text & "' AND FYCD='" & FYCD & "'", CN, adOpenDynamic, adLockOptimistic

If Not DBCDRS.EOF Then
   M_DBCD = Trim(DBCDRS!CODE & "")
Else
   M_DBCD = Empty
End If
DBCDRS.Close

If DBCDRS.State = 1 Then DBCDRS.Close
DBCDRS.Open "SELECT CODE FROM PKGNGMST WHERE NAME='" & cmbPackaging.Text & "' AND STATUS='A' AND RECSTAT='A'", CN, adOpenDynamic, adLockOptimistic
If Not DBCDRS.EOF Then
   PKGNG_COD = Trim(DBCDRS!CODE & "")
Else
   PKGNG_COD = Empty
End If
DBCDRS.Close

End Sub

Private Sub SetRawMaterial()
If InStr(1, cmbPackingType.Text, "GR") <> 0 Then   'Auto Consumption Stop in GR Case
   Exit Sub
End If

On Error GoTo LAST

Dim COUNT As Long: COUNT = 0
Dim ITMCODE As String
Dim TOTALQTY As Double, ITMQTY As Double, ITMRATE As Double, ITMAMT As Double
Dim TEMPRS As ADODB.Recordset
Set TEMPRS = New ADODB.Recordset
Dim STORERS As ADODB.Recordset
Set STORERS = New ADODB.Recordset
Dim TABLE As String, PCODE As String, FIELD As String

If InStr(1, cmbPackingType.Text, "JobWork") <> 0 Then    'PARTY REQUIRED in CASE OF JOB
    TABLE = "JOBIN"
    PCODE = GetCode("ACCMST", TXTPCOD, "NAME", "CODE")
    FIELD = "JOBQ"
Else
    TABLE = "STORETRAN"
    PCODE = "000000"
    FIELD = "BALQ"
End If

TOTALQTY = Val(TXTNETWT) * Val(TXTNOB)

If STORERS.State = 1 Then STORERS.Close
STORERS.Open "SELECT * FROM " & TABLE & " WHERE COMP='" & compPth & "' AND VTYP='PPF' AND SRNO='" & M_SRNO & "'", CN, adOpenDynamic, adLockOptimistic

If TEMPRS.State = 1 Then TEMPRS.Close
TEMPRS.Open "SELECT RICD,PERC FROM TXULOT WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & LSDVCD & "' AND LTNO='" & TXTLOTNO & "'", CN, adOpenDynamic, adLockOptimistic
Do While Not TEMPRS.EOF
   COUNT = COUNT + 1
   ITMCODE = Trim(TEMPRS!RICD & "")
   ITMQTY = (Val(TEMPRS!PERC) * TOTALQTY) / 100
   ITMRATE = 0
   ITMAMT = ITMQTY * ITMRATE
     
  Dim AI As String
  Dim BQ As Double
    
  STORERS.AddNew
  STORERS!COMP = compPth
  STORERS!unit = UNCD
  STORERS!DVCD = LSDVCD
  STORERS!VTYP = "PPF"
  STORERS!SRNO = M_SRNO
  STORERS!SRCH = COUNT
  STORERS!VBNO = TXTSLIP
  STORERS!chln = TXTSLIP
  STORERS!PCOD = PCODE
  STORERS!Date = Format(TXTVBDT, "YYYY/MM/DD")
  STORERS!dbcd = M_DBCD
  STORERS!ICOD = ITMCODE: AI = ITMCODE
  STORERS!PCES = 0
  STORERS!QNTY = ITMQTY: BQ = ITMQTY
  STORERS!RATE = ITMRATE
  STORERS!AMNT = ITMAMT
  STORERS!QORP = "Q"
  STORERS![User] = cUName
  If SAVEFLAG = True Then
     STORERS!SYSR = "N"
  Else
     STORERS!SYSR = "U"
  End If
  STORERS!OPER = "-"
  STORERS!RECSTAT = "A"
  STORERS.Update
TEMPRS.MoveNext
Loop
TEMPRS.Close

If InStr(1, cmbPackingType.Text, "GR") = 0 Then   'Auto Consumption Stop in GR Case
     
  STORERS.AddNew
  STORERS!COMP = compPth
  STORERS!unit = UNCD
  STORERS!DVCD = LSDVCD
  STORERS!VTYP = "PPF"
  STORERS!SRNO = M_SRNO
  STORERS!SRCH = COUNT + 1
  STORERS!VBNO = TXTSLIP
  STORERS!chln = TXTSLIP
  STORERS!PCOD = PCODE
  STORERS!Date = Format(TXTVBDT, "YYYY/MM/DD")
  STORERS!dbcd = M_DBCD
  STORERS!ICOD = GetCode("ITMMST", TXTCARTONNAME, "NAME", "CODE"): AI = GetCode("ITMMST", TXTCARTONNAME, "NAME", "CODE")
  ITMRATE = Val(GetCode("ITMMST", TXTCARTONNAME, "Name", "PURR"))
  ITMAMT = Val(TXTNOB) * ITMRATE
  STORERS!PCES = 0
  STORERS!QNTY = TXTNOB: BQ = TXTNOB
  STORERS!RATE = ITMRATE
  STORERS!AMNT = ITMAMT
  STORERS!QORP = "Q"
  STORERS![User] = cUName
  If SAVEFLAG = True Then
     STORERS!SYSR = "N"
  Else
     STORERS!SYSR = "U"
  End If
  STORERS!OPER = "-"
  STORERS!RECSTAT = "A"
  STORERS.Update
        
          
  STORERS.AddNew
  STORERS!COMP = compPth
  STORERS!unit = UNCD
  STORERS!DVCD = LSDVCD
  STORERS!VTYP = "PPF"
  STORERS!SRNO = M_SRNO
  STORERS!SRCH = COUNT + 2
  STORERS!VBNO = TXTSLIP
  STORERS!chln = TXTSLIP
  STORERS!PCOD = PCODE
  STORERS!Date = Format(TXTVBDT, "YYYY/MM/DD")
  STORERS!dbcd = M_DBCD
  STORERS!ICOD = GetCode("ITMMST", TXTCOPSNAME, "NAME", "CODE"): AI = GetCode("ITMMST", TXTCOPSNAME, "NAME", "CODE")
  ITMRATE = Val(GetCode("ITMMST", TXTCOPSNAME, "NAME", "PURR"))
  ITMAMT = Val(TXTNOB) * Val(txtCops) * ITMRATE
  STORERS!PCES = 0
  STORERS!QNTY = Val(TXTNOB) * Val(txtCops): BQ = Val(TXTNOB) * Val(txtCops)
  STORERS!RATE = ITMRATE
  STORERS!AMNT = ITMAMT
  STORERS!QORP = "Q"
  STORERS![User] = cUName
  If SAVEFLAG = True Then
     STORERS!SYSR = "N"
  Else
     STORERS!SYSR = "U"
  End If
  STORERS!OPER = "-"
  STORERS!RECSTAT = "A"
  STORERS.Update
 End If  ''Auto Consumption Stop in GR Case

Exit Sub
LAST:
MsgBox ERR.Description
Resume
End Sub

Private Sub SETPLYLIMIT()
Dim LIMITRS As ADODB.Recordset
Set LIMITRS = New ADODB.Recordset

If LIMITRS.State = 1 Then LIMITRS.Close
LIMITRS.Open "SELECT * FROM PKGNGMST WHERE STATUS='A' AND RECSTAT='A' AND NAME='" & cmbPackaging.Text & "' AND PALLET='Y'", CN, adOpenDynamic, adLockOptimistic
If Not LIMITRS.EOF Then
 FLEXPLY.Enabled = True
 FLEXPLY.Tag = Val(Trim(LIMITRS!NOPLY & ""))
Else
 FLEXPLY.Enabled = False
End If

LIMITRS.Close

End Sub

Private Sub setHeading()
Dim I As Long, J As Long
With FLEXPLY
    .TextMatrix(0, 0) = "Ply Name "
    .TextMatrix(1, 0) = "No.of Ply/Pallet "
    .ColWidth(0) = 1650
    .ColWidth(0) = 1650
End With

Dim COUNT As Long
Dim GETRS As ADODB.Recordset
Set GETRS = New ADODB.Recordset

If GETRS.State = 1 Then GETRS.Close
GETRS.Open "SELECT * FROM PLYMST WHERE RECSTAT='A'", CN, adOpenDynamic, adLockOptimistic
Do While Not GETRS.EOF
    COUNT = COUNT + 1
    FLEXPLY.TextMatrix(0, COUNT) = Trim(GETRS!NAME & "")
    FLEXPLY.ColWidth(COUNT) = 155 * Len(Trim(GETRS!NAME & "")) + 150
GETRS.MoveNext
Loop
GETRS.Close

FLEXPLY.Cols = COUNT + 1

End Sub



