VERSION 5.00
Object = "{A19332D7-D707-4A30-9F38-796D120AF5B3}#1.2#0"; "BtnPlus1.ocx"
Object = "{2BF72F7D-D367-4712-9146-5533EF3E691A}#1.2#0"; "FraPlus1.ocx"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WelchButton.ocx"
Begin VB.Form frmPkgStationMst 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Packing Station Master"
   ClientHeight    =   6900
   ClientLeft      =   1890
   ClientTop       =   2985
   ClientWidth     =   10620
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
   Icon            =   "frmPkgStationMst.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6900
   ScaleWidth      =   10620
   Begin FramePlusCtl.FramePlus FramePlus1 
      Height          =   6915
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   10695
      _ExtentX        =   18865
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
      Begin VB.CheckBox CHKONLP 
         BackColor       =   &H0080C0FF&
         Caption         =   "Online Packing Slip"
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
         TabIndex        =   7
         Top             =   2280
         Width           =   3735
      End
      Begin VB.CheckBox chkPallet 
         BackColor       =   &H0080C0FF&
         Caption         =   "Pallet"
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
         Left            =   5640
         TabIndex        =   11
         Top             =   2880
         Width           =   1335
      End
      Begin VB.CheckBox chkCopsWgt 
         BackColor       =   &H0080C0FF&
         Caption         =   "Cops Wgt. "
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
         Left            =   2280
         TabIndex        =   9
         Top             =   2880
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CheckBox chkBoxWgt 
         BackColor       =   &H0080C0FF&
         Caption         =   "Box Wgt. "
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
         Left            =   4080
         TabIndex        =   10
         Top             =   2880
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.CheckBox chkNoCops 
         BackColor       =   &H0080C0FF&
         Caption         =   "No. of Cops"
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
         TabIndex        =   8
         Top             =   2880
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.Frame frmConnection 
         Caption         =   "Connection Preferences"
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
         Height          =   1650
         Left            =   3840
         TabIndex        =   38
         Top             =   4320
         Width           =   3045
         Begin VB.ComboBox cboDataBits 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   41
            Top             =   360
            Width           =   1140
         End
         Begin VB.ComboBox cboParity 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   40
            Top             =   780
            Width           =   1140
         End
         Begin VB.ComboBox cboStopBits 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   39
            Top             =   1200
            Width           =   1140
         End
         Begin VB.Label Label8 
            Caption         =   "Data Bits:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   180
            TabIndex        =   44
            Top             =   375
            Width           =   1545
         End
         Begin VB.Label Label7 
            Caption         =   "Parity:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   180
            TabIndex        =   43
            Top             =   855
            Width           =   615
         End
         Begin VB.Label Label5 
            Caption         =   "Stop Bits:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   180
            TabIndex        =   42
            Top             =   1320
            Width           =   1125
         End
      End
      Begin VB.Frame frmFlow 
         Caption         =   "&Flow Control / Hand Shaking"
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
         Height          =   1650
         Left            =   480
         TabIndex        =   33
         Top             =   4320
         Width           =   3300
         Begin VB.OptionButton optFlow 
            Caption         =   "Xon/RTS"
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
            Index           =   3
            Left            =   1200
            MaskColor       =   &H00000000&
            TabIndex        =   37
            Top             =   720
            Width           =   1275
         End
         Begin VB.OptionButton optFlow 
            Caption         =   "RTS"
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
            Index           =   2
            Left            =   120
            MaskColor       =   &H00000000&
            TabIndex        =   36
            Top             =   720
            Width           =   855
         End
         Begin VB.OptionButton optFlow 
            Caption         =   "Xon/Xoff"
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
            Index           =   1
            Left            =   1200
            MaskColor       =   &H00000000&
            TabIndex        =   35
            Top             =   360
            Width           =   1335
         End
         Begin VB.OptionButton optFlow 
            Caption         =   "None"
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
            Index           =   0
            Left            =   120
            MaskColor       =   &H00000000&
            TabIndex        =   34
            Top             =   345
            Value           =   -1  'True
            Width           =   855
         End
      End
      Begin VB.TextBox txtcomport 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5640
         MaxLength       =   7
         TabIndex        =   15
         Top             =   3720
         Width           =   795
      End
      Begin VB.TextBox TxtBaurdRate 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   3120
         MaxLength       =   7
         TabIndex        =   13
         Top             =   3720
         Width           =   1155
      End
      Begin VB.CheckBox CHKWSCALE 
         BackColor       =   &H0080C0FF&
         Caption         =   "Box Weight through Weighing Scale."
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
         TabIndex        =   12
         Top             =   3360
         Width           =   3975
      End
      Begin VB.TextBox TXTPRFX 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   360
         MaxLength       =   1
         TabIndex        =   2
         Text            =   "M"
         Top             =   1680
         Width           =   435
      End
      Begin VB.OptionButton optPrePrinted 
         BackColor       =   &H0080C0FF&
         Caption         =   "Pre-Printed"
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
         Left            =   5160
         TabIndex        =   6
         Top             =   1680
         Width           =   1695
      End
      Begin VB.OptionButton optPlain 
         BackColor       =   &H0080C0FF&
         Caption         =   "Plain"
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
         Left            =   4200
         TabIndex        =   5
         Top             =   1680
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.TextBox txtCode 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   4560
         MaxLength       =   49
         TabIndex        =   28
         ToolTipText     =   "Enter the Description of Item."
         Top             =   960
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.TextBox txtName 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   360
         MaxLength       =   49
         TabIndex        =   1
         ToolTipText     =   "Enter the Description of Item."
         Top             =   960
         Width           =   5235
      End
      Begin VB.TextBox txtLPNO 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   2520
         MaxLength       =   7
         TabIndex        =   4
         Text            =   "0000000"
         Top             =   1680
         Width           =   1245
      End
      Begin VB.TextBox txtLBNO 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   1080
         MaxLength       =   7
         TabIndex        =   3
         Text            =   "0000000"
         Top             =   1680
         Width           =   1275
      End
      Begin ButtonPlusCtl.ButtonPlus cmdFind 
         Height          =   375
         Left            =   8880
         TabIndex        =   23
         Top             =   5640
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         BackStyle       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "&Find"
      End
      Begin ButtonPlusCtl.ButtonPlus cmdClear 
         Height          =   375
         Left            =   7560
         TabIndex        =   22
         Top             =   5640
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         BackStyle       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "C&lear"
      End
      Begin VB.ListBox lstRef 
         Height          =   4740
         Left            =   7440
         Sorted          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   720
         Width           =   3015
      End
      Begin WelchButton.lvButtons_H cmdAdd 
         Height          =   495
         Left            =   720
         TabIndex        =   0
         Top             =   6240
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
         Image           =   "frmPkgStationMst.frx":058C
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdEdit 
         Height          =   495
         Left            =   4680
         TabIndex        =   18
         Top             =   6240
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
         Image           =   "frmPkgStationMst.frx":0926
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdDelete 
         Height          =   495
         Left            =   6000
         TabIndex        =   19
         Top             =   6240
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
         Image           =   "frmPkgStationMst.frx":0CC0
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdSave 
         Height          =   495
         Left            =   2040
         TabIndex        =   16
         Top             =   6240
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
         Image           =   "frmPkgStationMst.frx":105A
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdCancel 
         Height          =   495
         Left            =   3360
         TabIndex        =   17
         Top             =   6240
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
         Image           =   "frmPkgStationMst.frx":1DE4
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdExit 
         Height          =   495
         Left            =   7320
         TabIndex        =   20
         Top             =   6240
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
         Image           =   "frmPkgStationMst.frx":2236
         cBack           =   -2147483633
      End
      Begin VB.Shape Shape6 
         Height          =   495
         Left            =   240
         Shape           =   4  'Rounded Rectangle
         Top             =   2160
         Width           =   6855
      End
      Begin VB.Shape Shape5 
         Height          =   495
         Left            =   240
         Shape           =   4  'Rounded Rectangle
         Top             =   2760
         Width           =   6855
      End
      Begin VB.Line Line3 
         X1              =   240
         X2              =   7080
         Y1              =   4200
         Y2              =   4200
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Com Port :"
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
         Left            =   4440
         TabIndex        =   14
         Top             =   3720
         Width           =   1080
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Max. Speed / Baud Rate :"
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
         TabIndex        =   32
         Top             =   3720
         Width           =   2520
      End
      Begin VB.Shape Shape4 
         Height          =   2535
         Left            =   240
         Shape           =   4  'Rounded Rectangle
         Top             =   3480
         Width           =   6855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Prefix "
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
         TabIndex        =   31
         Top             =   1320
         Width           =   645
      End
      Begin VB.Shape Shape3 
         Height          =   735
         Left            =   4080
         Shape           =   4  'Rounded Rectangle
         Top             =   1320
         Width           =   3015
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Printing type of Packing Slip"
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
         Left            =   4200
         TabIndex        =   30
         Top             =   1320
         Width           =   2760
      End
      Begin VB.Label LBLHEAD 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Packing Station Master"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   3240
         TabIndex        =   29
         Top             =   195
         Width           =   2415
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H0080C0FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000080&
         Height          =   350
         Left            =   3120
         Shape           =   4  'Rounded Rectangle
         Top             =   195
         Width           =   2655
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Station Name     "
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
         Left            =   405
         TabIndex        =   27
         Top             =   720
         Width           =   1605
      End
      Begin VB.Label lblUnit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Last Pallet No.  "
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
         TabIndex        =   26
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Last Box No.     "
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
         Left            =   1080
         TabIndex        =   25
         Top             =   1320
         Width           =   1365
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         X1              =   7320
         X2              =   7320
         Y1              =   600
         Y2              =   6120
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         X1              =   120
         X2              =   10560
         Y1              =   6120
         Y2              =   6120
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         X1              =   240
         X2              =   10440
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         Height          =   6735
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   10455
      End
   End
End
Attribute VB_Name = "frmPkgStationMst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Settings As String
Dim SAVEFLAG As Boolean
Dim COND As String
Private iFlow                   As Integer
Private NewPort                 As Integer

Private Sub CHKWSCALE_Click()
If CHKWSCALE.Value = 1 Then
   frmFlow.Enabled = True
   frmConnection.Enabled = True
Else
   frmFlow.Enabled = False
   frmConnection.Enabled = False
End If
End Sub

Private Sub cmdAdd_Click()
    Call ClsData(Me)
    Call btn_sts(False)
    
    txtName.SetFocus
    SAVEFLAG = True
    cmdCancel.Cancel = True
    cmdDelete.Enabled = False
End Sub

Private Sub cmdAdd_GotFocus()
    Msg ActiveControl.ToolTipText
End Sub

Private Sub cmdCancel_Click()
    cmdExit.Cancel = True
    Call btn_sts(True)
    Call ClsData(Me)
    Call addWgt
    chkNoCops.Value = 1
    chkBoxWgt.Value = 1
    chkCopsWgt.Value = 1
    chkPallet.Value = 0
    optFlow(0).Value = True
    optPlain.Value = True
    cmdAdd.SetFocus
End Sub

Private Sub cmdCancel_GotFocus()
    Msg ActiveControl.ToolTipText
End Sub

Private Sub cmdCLEAR_Click()
    Call ClsData(Me)
    lstRef.ListIndex = -1
End Sub

Private Sub cmdClear_GotFocus()
    Msg ActiveControl.ToolTipText
End Sub

Private Sub cmdDelete_Click()

  If M_USRSECLEVL = "1" Then
    If ReadConfigMaster("000017", 6, "M") = False Then ModuleDeniedMessage: Exit Sub
  End If
    Dim ANS As String, TEMPRS As New ADODB.Recordset
    
    If isFurtherEntryExist("STATION", txtCode) Then
        MsgBox "Further Entry Exist"
        Call ClsData(Me)
        lstRef.ListIndex = -1
        Call btn_sts(True)
        Exit Sub
    End If
    
    If txtCode.Text = "" Then Exit Sub
    ANS = MsgBox("Do you Want to Delete this record?", vbYesNo + vbQuestion, App.TITLE)
    If ANS = vbYes Then
       CN.Execute "Delete from PCKMST where " & COND & " AND CODE ='" & Trim(txtCode.Text) & "'"
       
       'DAILYSTATUS
       Call DAILYSTATUS("PSM", txtCode, "", 0, "", 0, cUName, "D", Now, Now)
       
       lstRef.RemoveItem lstRef.ListIndex
    End If
                
    Call ClsData(Me)
    lstRef.ListIndex = -1
    Call btn_sts(True)
End Sub

Private Sub cmdDelete_GotFocus()
    Msg ActiveControl.ToolTipText
End Sub

Private Sub cmdEdit_Click()

  If M_USRSECLEVL = "1" Then
    If ReadConfigMaster("000017", 5, "M") = False Then ModuleDeniedMessage: Exit Sub
  End If
    cmdCancel.Cancel = True
    Call btn_sts(False)
    
    If lstRef.ListIndex = -1 Then lstRef.SetFocus Else txtName.SetFocus
    SAVEFLAG = False
    
End Sub

Private Sub cmdEdit_GotFocus()
    Msg ActiveControl.ToolTipText
End Sub

Private Sub CMDEXIT_Click()
    key_PressNew = False
    Unload Me
End Sub

Private Sub cmdExit_GotFocus()
    Msg ActiveControl.ToolTipText
End Sub

Private Sub CMDFIND_Click()
    NEW_VISIBLE = False
    If Me.Tag <> Empty Then Ref_Cat = Me.Tag
    M_DESC = Empty
    Key = Empty
    txtName.Text = SearchList1("Select TOP 20 CODE, NAME FROM PCKMST WHERE " & COND & "", 0, "", "List Of " & Me.Caption)
    txtCode.Text = Key
    
    lstRef.Text = txtName.Text
    'If cmdEdit.Enabled = True Then
    '    cmdEdit.SetFocus
    'End If
    
    If txtName <> Empty Then
       txtName.Enabled = True
       txtName.SetFocus
    End If
End Sub

Private Sub cmdFind_GotFocus()
    Msg ActiveControl.ToolTipText
End Sub

Private Sub cmdSave_Click()
On Error GoTo errPRIMARYKEY
    Dim SQL As String
    Dim PRNTYP As String: PRNTYP = "Y"
           
    If optPrePrinted.Value = True Then
       PRNTYP = "N"
    End If
                
    Dim TEMPRS As New ADODB.Recordset
    Dim Ctrl As Control
    
    txtName.Text = Trim(txtName.Text)
    
    For Each Ctrl In Me
        If TypeOf Ctrl Is TextBox Then
            Ctrl = Replace(Ctrl, "'", "", 1)
        End If
    Next
           
    If Trim(txtName.Text) = "" Then
        MsgBox "Please Enter Packing Station Name.", vbInformation, App.TITLE
        txtName = Trim(txtName)
        txtName.SetFocus
        Exit Sub
    End If
    
    If Len(txtLBNO) <> 7 Then
       MsgBox "Please Enter 7 digit Serial.", vbInformation, App.TITLE
       txtLBNO = Trim(txtLBNO)
       txtLBNO.SetFocus
       Exit Sub
    End If
    
    If Len(txtLPNO) <> 7 Then
       MsgBox "Please Enter 7 digit Serial.", vbInformation, App.TITLE
       txtLPNO = Trim(txtLPNO)
       txtLPNO.SetFocus
       Exit Sub
    End If
    
    If Len(Trim(TXTPRFX)) <> 1 Then
       MsgBox "Please Enter 1 Character for Prefix.", vbInformation, App.TITLE
       TXTPRFX = Trim(TXTPRFX)
       TXTPRFX.SetFocus
       Exit Sub
    End If
              
    If TEMPRS.State = 1 Then TEMPRS.Close
    TEMPRS.Open "Select * from PCKMST where " & COND & " AND (Upper([name])='" & UCase(Trim(txtName.Text)) & "' OR PRFX = '" & UCase(Trim(TXTPRFX.Text)) & "')", CN, adOpenDynamic, adLockOptimistic
    If TEMPRS.EOF = False And SAVEFLAG Then
       MsgBox "This Name Already Is In Use With Same Category Or Other Category !!!", vbInformation, App.TITLE
       TEMPRS.Close
       Exit Sub
    End If
    
    If Not SAVEFLAG Then 'EDIT MODE
    If TEMPRS.State = 1 Then TEMPRS.Close
    TEMPRS.Open "Select * from PCKMST where " & COND & " AND Upper([name])='" & UCase(Trim(txtName.Text)) & "' AND CODE <> '" & txtCode & "' ", CN, adOpenDynamic, adLockOptimistic
    If Not TEMPRS.EOF Then
       MsgBox "This Name Already Is In Use With Same Category Or Other Category !!!", vbInformation, App.TITLE
       txtName.SetFocus
       TEMPRS.Close
       Exit Sub
    End If

    Settings = Trim$(TxtBaurdRate) & "," & Left$(cboParity.Text, 1) & "," & Trim$(cboDataBits.Text) & "," & Trim$(cboStopBits.Text)
    
    If TEMPRS.State = 1 Then TEMPRS.Close
    TEMPRS.Open "Select * from PCKMST where " & COND & " AND Upper([PRFX])='" & UCase(Trim(TXTPRFX.Text)) & "' AND CODE <> '" & txtCode & "' ", CN, adOpenDynamic, adLockOptimistic
    If Not TEMPRS.EOF Then
       MsgBox "This Prefix Already Is In Use With Same Category Or Other Category !!!", vbInformation, App.TITLE
       TXTPRFX.SetFocus
       TEMPRS.Close
       Exit Sub
    End If
    End If
    
    If SAVEFLAG = True Then
        On Error GoTo errPRIMARYKEY
        
        txtCode.Text = GENSIXCOD("Select IsNull(Max(CODE),0) AS CODE From PCKMST WHERE " & COND)
          
        SQL = "insert into PCKMST (COMP,UNIT,CODE,[NAME],PRFX,LBNO,LPNO,PLAIN," & _
              "LSTPCKDT,STATUS,BAURDRATE,WSCALE,COMPORTX,RECSTAT,FLOW,SETTINGS,REQNOCOPS,REQCOPSWGT,REQBOXWGT,REQPALLET,ONLP) " _
              & " values('" & compPth & "','" & UNCD & "','" & Trim(txtCode.Text) & _
              "','" & UCase(Trim(txtName.Text)) & "','" & UCase(Trim(TXTPRFX)) & _
              "','" & Trim(TXTPRFX & txtLBNO.Text & "00") & _
              "','" & Trim(TXTPRFX & txtLPNO.Text & "00") & "','" & PRNTYP & _
              "','" & Format("01/01/2000", "MM/DD/YYYY") & "','A','" & Val(TxtBaurdRate) & _
              "','" & IIf(CHKWSCALE.Value = 1, "Y", "N") & "','" & Val(txtcomport) & _
              "','A','" & iFlow & "','" & Settings & "','" & IIf(chkNoCops.Value = 1, "Y", "N") & _
              "','" & IIf(chkCopsWgt.Value = 1, "Y", "N") & "','" & IIf(chkBoxWgt.Value = 1, "Y", "N") & _
              "','" & IIf(chkPallet.Value = 1, "Y", "N") & "','" & IIf(CHKONLP.Value = 1, "Y", "N") & "') "
                  
        CN.BeginTrans
        CN.Execute SQL
        'DAILYSTATUS
         Call DAILYSTATUS("PSM", txtCode, "", 0, "", 0, cUName, "N", Now, Now)
        CN.CommitTrans
        
        lstRef.AddItem UCase(txtName.Text)
    Else
    CN.BeginTrans
    CN.Execute ("Update PCKMST set PRFX='" & UCase(Trim(TXTPRFX.Text)) & "',NAME = '" & UCase(Trim(txtName.Text)) & _
               "',LBNO = '" & Trim(TXTPRFX & txtLBNO.Text & "00") & _
               "',LPNO='" & Trim(TXTPRFX & txtLPNO.Text & "00") & _
               "',WSCALE ='" & IIf(CHKWSCALE.Value = 1, "Y", "N") & _
               "',BAURDRATE='" & Val(TxtBaurdRate) & "',COMPORTX='" & Val(txtcomport) & _
               "',PLAIN='" & PRNTYP & "',FLOW='" & iFlow & _
               "',SETTINGS='" & Settings & "',REQNOCOPS ='" & IIf(chkNoCops.Value = 1, "Y", "N") & _
               "',REQCOPSWGT ='" & IIf(chkCopsWgt.Value = 1, "Y", "N") & _
               "',REQBOXWGT ='" & IIf(chkBoxWgt.Value = 1, "Y", "N") & _
               "',REQPALLET ='" & IIf(chkPallet.Value = 1, "Y", "N") & _
               "',ONLP ='" & IIf(CHKONLP.Value = 1, "Y", "N") & _
               "' where " & COND & " AND CODE ='" & Trim(txtCode.Text) & "' ")
    
    'DAILYSTATUS
     Call DAILYSTATUS("PSM", txtCode, "", 0, "", 0, cUName, "M", Now, Now)
    
    CN.CommitTrans
    lstRef.Clear
    Call FillList("Select [NAME] from PCKMST where " & COND & "ORDER BY [NAME]", lstRef)
     
    lstRef.ListIndex = -1
    End If
  
    Call btn_sts(True)
    sTxt = txtName.Text
 
    Call ClsData(Me)
    Call cmdCancel_Click
    cmdAdd.SetFocus
    cmdExit.Cancel = True
    Exit Sub

errPRIMARYKEY:
    MsgBox ERR.Description
    CN.RollbackTrans
    If ERR.Number = -2147217873 Or -2147217900 Then
        txtName.SetFocus
        MsgBox "This Name Already Registered With Other Category!!!", vbInformation, "Already Registered"
    Else
        ErrNumber = ERR.Number
        ErrMessage = ERR.Description
        frm_ErrorHandler.Show vbModal
    End If
    ERR.Clear
End Sub

Private Sub cmdSave_GotFocus()
    Msg ActiveControl.ToolTipText
End Sub

Private Sub Form_Activate()
    Call ColorComponent(Me)
    Me.BackColor = RGB(RED, GREEN, BLUE)
    If key_PressNew Then cmdAdd.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If ActiveControl.NAME = "lstRef" Then Exit Sub
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
Call CenterChild(frm_Main, Me)
Call ColorComponent(Me)
On Error GoTo errLoad
  Call btn_sts(True)
  COND = " COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND RECSTAT='A' "
  Call FillList("Select [NAME] from PCKMST where " & COND & "ORDER BY [NAME]", lstRef)
    
  cmdExit.Cancel = True
  Me.Show
  Call addWgt
  
  Exit Sub

errLoad:
    ErrNumber = ERR.Number
    ErrMessage = ERR.Description
    frm_ErrorHandler.Show vbModal
End Sub

Private Sub btn_sts(bool As Boolean)
    cmdSave.Enabled = Not bool
    cmdCancel.Enabled = Not bool
    cmdAdd.Enabled = bool
    cmdEdit.Enabled = bool
    cmdDelete.Enabled = Not bool
    txtName.Enabled = Not bool
    txtLBNO.Enabled = Not bool
    txtLPNO.Enabled = Not bool
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Msg Empty
End Sub

Private Sub lstRef_Click()
    SAVEFLAG = False
    Dim TEMPRS As New ADODB.Recordset
    If lstRef.ListIndex = -1 Then Exit Sub
    If TEMPRS.State = 1 Then TEMPRS.Close
    TEMPRS.Open "Select *,SUBSTRING(LBNO,2,7) AS LASTBOXNO,SUBSTRING(LPNO,2,7) AS LASTPLTNO,BAURDRATE from PCKMST where " & COND & "AND [NAME] = '" & (lstRef.List(lstRef.ListIndex)) & "'", CN, adOpenDynamic, adLockOptimistic
    
    With TEMPRS
        txtCode.Text = !CODE & ""
        txtName.Text = Trim(![NAME] & "")
        TXTPRFX.Text = Trim(![Prfx] & "")
        
        CHKWSCALE.Value = IIf(Trim(![WSCALE] & "") = "Y", 1, 0)
        
        chkNoCops.Value = IIf(Trim(![REQNOCOPS] & "") = "Y", 1, 0)
        chkCopsWgt.Value = IIf(Trim(![REQCOPSWGT] & "") = "Y", 1, 0)
        chkBoxWgt.Value = IIf(Trim(![REQBOXWGT] & "") = "Y", 1, 0)
        chkPallet.Value = IIf(Trim(![REQPALLET] & "") = "Y", 1, 0)
        CHKONLP.Value = IIf(Trim(![ONLP] & "") = "Y", 1, 0)
                
            txtLBNO = Trim(!LASTBOXNO & "")
            txtLPNO = Trim(!LASTPLTNO & "")
            TxtBaurdRate = Val(!BAURDRATE)
            txtcomport = Val(!COMPORTX & "")
                       
        'Call addWgt
                       
        If !PLAIN & "" = "Y" Then
          optPlain.Value = True
        Else
          optPrePrinted.Value = True
        End If
        
        optFlow(!FLOW).Value = True
        
        Settings = Trim(!Settings & "")
        
        ' In all cases the right most part of Settings will be 1 character
        ' except when there are 1.5 stop bits.
          Dim Offset                  As Integer
          If InStr(Settings, ".") > 0 Then
              Offset = 2
          Else
              Offset = 0
          End If
          
          'cboSpeed.Text = Left$(Settings, Len(Settings) - 6 - Offset)
          
          '"UPDATE PCKMST SET SETTINGS='1200,n,8,1'"
     
    If Settings = Empty Then
       Settings = "1200,n,8,1"
    End If
     
    Select Case Mid$(Settings, Len(Settings) - 4 - Offset, 1)

        Case "e", "E"
            cboParity.ListIndex = 0

        Case "m", "M"
            cboParity.ListIndex = 1

        Case "n", "N"
            cboParity.ListIndex = 2

        Case "o", "O"
            cboParity.ListIndex = 3

        Case "s", "S"
            cboParity.ListIndex = 4
    End Select

    cboDataBits.ListIndex = Val(Mid$(Settings, Len(Settings) - 2 - Offset, 1)) - 4
    cboStopBits.Text = Right$(Settings, 1 + Offset)
        
                
    End With
    TEMPRS.Close
End Sub

Private Sub lstRef_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
 txtName.Enabled = True
 txtName.SetFocus
End If
End Sub

Private Sub lstRef_GotFocus()
    lstRef.BackColor = RGB(BRED, BGREEN, BBLUE)
    Msg "Enter Address"
End Sub

Private Sub lstRef_LostFocus()
lstRef.BackColor = vbWhite
End Sub

Private Sub txtLBNO_GotFocus()
    txtLBNO.BackColor = RGB(BRED, BGREEN, BBLUE)
    Msg "Enter Last Box No."
    txtLBNO.ToolTipText = "Enter Last Box No."
    txtLBNO.SelStart = 0
    txtLBNO.SelLength = Len(txtLBNO)
End Sub

Private Sub txtLBNO_KeyPress(KeyAscii As Integer)
   If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub txtLBNO_LostFocus()
txtLBNO.BackColor = vbWhite
End Sub

Private Sub txtLPNO_GotFocus()
  txtLPNO.BackColor = RGB(BRED, BGREEN, BBLUE)
  Msg "Enter Last Pallet Number"
  txtLPNO.SelStart = 0
  txtLPNO.SelLength = Len(txtLPNO)
End Sub

Private Sub txtLPNO_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub txtLPNO_LostFocus()
txtLPNO.BackColor = vbWhite
End Sub

Private Sub txtName_GotFocus()
    txtName.BackColor = RGB(BRED, BGREEN, BBLUE)
    Msg "Enter Packing Station Name"
    txtName.SelStart = 0
    txtName.SelLength = Len(txtName)
End Sub

Private Sub txtName_LostFocus()
txtName.BackColor = vbWhite
End Sub

Public Sub FillList(SQL As String, lst As ListBox)
    Dim TEMPRS As New ADODB.Recordset
    TEMPRS.Open SQL, CN, adOpenDynamic, adLockOptimistic
    If TEMPRS.EOF = True Then Exit Sub
    TEMPRS.MoveFirst
    Do While Not TEMPRS.EOF
        lst.AddItem Trim(TEMPRS.Fields(0).Value)
        TEMPRS.MoveNext
    Loop
    TEMPRS.Close
End Sub

Private Sub TXTPRFX_GotFocus()
  TXTPRFX.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTPRFX_LostFocus()
  TXTPRFX.BackColor = vbWhite
End Sub

Private Sub addWgt()
  ' Load Data Bit Settings
    cboDataBits.Clear
    cboDataBits.AddItem "4"
    cboDataBits.AddItem "5"
    cboDataBits.AddItem "6"
    cboDataBits.AddItem "7"
    cboDataBits.AddItem "8"
    cboDataBits.Text = "8"
  ' Load Parity Settings
    cboParity.Clear
    cboParity.AddItem "Even"
    cboParity.AddItem "Odd"
    cboParity.AddItem "None"
    cboParity.AddItem "Mark"
    cboParity.AddItem "Space"
    cboParity.Text = "None"
  ' Load Stop Bit Settings
    cboStopBits.Clear
    cboStopBits.AddItem "1"
    cboStopBits.AddItem "1.5"
    cboStopBits.AddItem "2"
    cboStopBits.Text = "1"
End Sub

Private Sub optFlow_Click(INDEX As Integer)
    iFlow = INDEX
End Sub

