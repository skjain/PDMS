VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.MDIForm frm_Main 
   BackColor       =   &H80000003&
   Caption         =   "Enterprise Resource Planning (ERP)"
   ClientHeight    =   8190
   ClientLeft      =   270
   ClientTop       =   2115
   ClientWidth     =   11880
   Icon            =   "frm_main.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picInfo 
      Align           =   2  'Align Bottom
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   420
      Left            =   0
      ScaleHeight     =   360
      ScaleWidth      =   11820
      TabIndex        =   1
      Top             =   7440
      Width           =   11880
      Begin VB.Label lblCompanycode 
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
         Height          =   255
         Left            =   10725
         TabIndex        =   7
         Top             =   75
         Width           =   1035
      End
      Begin VB.Label lblNow 
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
         Left            =   6705
         TabIndex        =   6
         Top             =   90
         Width           =   1635
      End
      Begin VB.Label lblUnitName 
         Caption         =   "Unit Name :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1470
         TabIndex        =   5
         Top             =   75
         Width           =   4380
      End
      Begin VB.Label Label3 
         Caption         =   "Company Code :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8970
         TabIndex        =   4
         Top             =   75
         Width           =   1755
      End
      Begin VB.Label Label1 
         Caption         =   "Time :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6015
         TabIndex        =   3
         Top             =   75
         Width           =   945
      End
      Begin VB.Label Label2 
         Caption         =   "Unit Name :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   150
         TabIndex        =   2
         Top             =   75
         Width           =   1335
      End
   End
   Begin VB.Timer Timer1 
      Left            =   5520
      Top             =   2040
   End
   Begin MSComctlLib.StatusBar StsMsg 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   7860
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   8467
            MinWidth        =   8467
            Text            =   "Ready.."
            TextSave        =   "Ready.."
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2117
            MinWidth        =   2117
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2117
            MinWidth        =   2117
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "19:25"
         EndProperty
      EndProperty
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
   Begin Crystal.CrystalReport CRPT 
      Left            =   5520
      Top             =   2640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin MSComctlLib.Toolbar Toolbar2 
      Align           =   3  'Align Left
      Height          =   7440
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   13123
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImgList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "A/c Master"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Item Master"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Agent"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "User Setup"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Design Report"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Unit Setup"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Backup Data"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Data Scanning"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Calculator"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Unit Configuration"
            ImageIndex      =   10
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgList 
      Left            =   6840
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   81
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":08CA
            Key             =   "s_Key1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":0E64
            Key             =   "s_Key2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":13FE
            Key             =   "s_Key3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":1998
            Key             =   "s_Key4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":1F32
            Key             =   "s_Key5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":24CC
            Key             =   "s_Key6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":2A66
            Key             =   "s_Key7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":3000
            Key             =   "s_Key8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":359A
            Key             =   "s_Key9"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":3B34
            Key             =   "s_Key10"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":40CE
            Key             =   "s_Key11"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":4668
            Key             =   "s_Key12"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":4C02
            Key             =   "s_Key13"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":519C
            Key             =   "s_Key14"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":5736
            Key             =   "s_Key15"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":5CD0
            Key             =   "s_Key16"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":626A
            Key             =   "s_Key17"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":6804
            Key             =   "s_Key18"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":6D9E
            Key             =   "s_Key19"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":7338
            Key             =   "s_Key20"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":78D2
            Key             =   "s_Key21"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":7E6C
            Key             =   "s_Key22"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":8406
            Key             =   "s_Key23"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":89A0
            Key             =   "s_Key24"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":8F3A
            Key             =   "s_Key25"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":94D4
            Key             =   "s_Key26"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":9A6E
            Key             =   "s_Key27"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":A008
            Key             =   "s_Key28"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":A5A2
            Key             =   "s_Key29"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":AB3C
            Key             =   "s_Key30"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":B0D6
            Key             =   "s_Key31"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":B670
            Key             =   "s_Key32"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":BC0A
            Key             =   "s_Key33"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":C1A4
            Key             =   "s_Key34"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":C73E
            Key             =   "s_Key35"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":CCD8
            Key             =   "s_Key36"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":D272
            Key             =   "s_Key37"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":D80C
            Key             =   "s_Key38"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":DDA6
            Key             =   "s_Key39"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":E340
            Key             =   "s_Key40"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":E8DA
            Key             =   "s_Key41"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":EE74
            Key             =   "s_Key42"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":F40E
            Key             =   "s_Key43"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":F9A8
            Key             =   "s_Key44"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":FF42
            Key             =   "s_Key45"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":104DC
            Key             =   "s_Key46"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":10A76
            Key             =   "s_Key47"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":11010
            Key             =   "s_Key48"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":115AA
            Key             =   "s_Key49"
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":11B44
            Key             =   "s_Key50"
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":120DE
            Key             =   "s_Key51"
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":12678
            Key             =   "s_Key52"
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":12C12
            Key             =   "s_Key53"
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":131AC
            Key             =   "s_Key54"
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":13746
            Key             =   "s_Key55"
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":13CE0
            Key             =   "s_Key56"
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":1427A
            Key             =   "s_Key57"
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":14814
            Key             =   "s_Key58"
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":14DAE
            Key             =   "s_Key59"
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":15348
            Key             =   "s_Key60"
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":158E2
            Key             =   "s_Key61"
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":15E7C
            Key             =   "s_Key62"
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":16416
            Key             =   "s_Key63"
         EndProperty
         BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":169B0
            Key             =   "s_Key64"
         EndProperty
         BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":16F4A
            Key             =   "s_Key65"
         EndProperty
         BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":174E4
            Key             =   "s_Key66"
         EndProperty
         BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":17A7E
            Key             =   "s_Key67"
         EndProperty
         BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":18018
            Key             =   "s_Key68"
         EndProperty
         BeginProperty ListImage69 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":185B2
            Key             =   "s_Key69"
         EndProperty
         BeginProperty ListImage70 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":18B4C
            Key             =   "s_Key70"
         EndProperty
         BeginProperty ListImage71 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":190E6
            Key             =   "s_Key71"
         EndProperty
         BeginProperty ListImage72 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":19680
            Key             =   "s_Key72"
         EndProperty
         BeginProperty ListImage73 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":19C1A
            Key             =   "s_Key73"
         EndProperty
         BeginProperty ListImage74 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":1A1B4
            Key             =   "s_Key74"
         EndProperty
         BeginProperty ListImage75 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":1A74E
            Key             =   "s_Key75"
         EndProperty
         BeginProperty ListImage76 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":1ACE8
            Key             =   "s_Key76"
         EndProperty
         BeginProperty ListImage77 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":1B282
            Key             =   "s_Key77"
         EndProperty
         BeginProperty ListImage78 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":1B3DC
            Key             =   "s_Key78"
         EndProperty
         BeginProperty ListImage79 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":1B976
            Key             =   "s_Key79"
         EndProperty
         BeginProperty ListImage80 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":1BF10
            Key             =   "s_Key80"
         EndProperty
         BeginProperty ListImage81 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":1C4AA
            Key             =   "s_Key81"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuMaster 
      Caption         =   " Master "
      Begin VB.Menu mnuMasterGrp 
         Caption         =   "&Group Master"
         Index           =   0
         Begin VB.Menu mnuMasterGrpOp 
            Caption         =   "Schedule Master"
            Index           =   0
         End
         Begin VB.Menu mnuMasterGrpOp 
            Caption         =   "A/C. &Group"
            Index           =   1
         End
      End
      Begin VB.Menu mnuMasterGrp 
         Caption         =   "Country / State / City / Account"
         Index           =   1
         Begin VB.Menu mnuMasterGrp1 
            Caption         =   "Country"
            Index           =   0
         End
         Begin VB.Menu mnuMasterGrp1 
            Caption         =   "State"
            Index           =   1
         End
         Begin VB.Menu mnuMasterGrp1 
            Caption         =   "City"
            Index           =   2
         End
         Begin VB.Menu mnuMasterGrp1 
            Caption         =   "Account"
            Index           =   3
         End
         Begin VB.Menu mnuMasterGrp1 
            Caption         =   "Delivery Address"
            Index           =   4
         End
         Begin VB.Menu mnuMasterGrp1 
            Caption         =   "Excise Opening Balance"
            Index           =   5
         End
      End
      Begin VB.Menu mnuMasterGrp 
         Caption         =   "Unit Master"
         Index           =   2
      End
      Begin VB.Menu mnuMasterGrp 
         Caption         =   "Division Master"
         Index           =   3
      End
      Begin VB.Menu mnuMasterGrp 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuMasterGrp 
         Caption         =   "&Finish Item/Product"
         Index           =   6
         Begin VB.Menu mnuFinItmMst 
            Caption         =   "Finish Item Master"
            Index           =   0
         End
         Begin VB.Menu mnuFinItmMst 
            Caption         =   "Grade Master"
            Index           =   1
         End
         Begin VB.Menu mnuFinItmMst 
            Caption         =   "Sub Grade Master"
            Index           =   2
         End
         Begin VB.Menu mnuFinItmMst 
            Caption         =   "Lot/Batch Master"
            Index           =   3
         End
         Begin VB.Menu mnuFinItmMst 
            Caption         =   "Location For Finish Item"
            Index           =   4
         End
      End
      Begin VB.Menu mnuMasterGrp 
         Caption         =   "&Raw Item / Product"
         Index           =   7
         Begin VB.Menu mnuMasterItmOp 
            Caption         =   "Item Category"
            Index           =   0
         End
         Begin VB.Menu mnuMasterItmOp 
            Caption         =   "Item Group"
            Index           =   1
         End
         Begin VB.Menu mnuMasterItmOp 
            Caption         =   "Item Master"
            Index           =   2
         End
         Begin VB.Menu mnuMasterItmOp 
            Caption         =   "Item wise Store &Opening Stock"
            Index           =   5
         End
         Begin VB.Menu mnuMasterItmOp 
            Caption         =   "Division+Machine Wise WIP Opening Stock"
            Index           =   11
         End
         Begin VB.Menu mnuMasterItmOp 
            Caption         =   "Location For Store Items"
            Index           =   13
         End
         Begin VB.Menu mnuMerge 
            Caption         =   "MergeWise Item Opening"
         End
      End
      Begin VB.Menu mnuMasterGrp 
         Caption         =   "References"
         Index           =   8
         Begin VB.Menu mnuMasRefOp 
            Caption         =   "Area"
            Index           =   0
         End
         Begin VB.Menu mnuMasRefOp 
            Caption         =   "Party Group"
            Index           =   1
         End
         Begin VB.Menu mnuMasRefOp 
            Caption         =   "Broker"
            Index           =   2
         End
         Begin VB.Menu mnuMasRefOp 
            Caption         =   "Transport"
            Index           =   3
         End
         Begin VB.Menu mnuMasRefOp 
            Caption         =   "Machine Master"
            Index           =   6
         End
         Begin VB.Menu mnuMasRefOp 
            Caption         =   "Unit Of Measurement (UOM)"
            Index           =   8
         End
         Begin VB.Menu mnuMasRefOp 
            Caption         =   "Bill Entry Charges"
            Index           =   9
         End
         Begin VB.Menu mnuMasRefOp 
            Caption         =   "Production Loss Reason Master"
            Index           =   12
         End
         Begin VB.Menu mnuMasRefOp 
            Caption         =   "Packing Station Master"
            Index           =   16
         End
         Begin VB.Menu mnuMasRefOp 
            Caption         =   "Sales Man Master"
            Index           =   17
         End
         Begin VB.Menu mnuMasRefOp 
            Caption         =   "Packaging Master"
            Index           =   20
         End
         Begin VB.Menu mnuMasRefOp 
            Caption         =   "PLY Master"
            Index           =   21
         End
         Begin VB.Menu mnuMasRefOp 
            Caption         =   "Vehicle Master"
            Index           =   22
         End
         Begin VB.Menu mnuMasRefOp 
            Caption         =   "Bank Master"
            Index           =   23
         End
         Begin VB.Menu mnuMasRefOp 
            Caption         =   "Rate Master "
            Index           =   24
         End
         Begin VB.Menu mnuMasRefOp 
            Caption         =   "Terms And Condition"
            Index           =   25
         End
         Begin VB.Menu mnuMasRefOp 
            Caption         =   "Export Type Master"
            Index           =   26
         End
         Begin VB.Menu mnuMasRefOp 
            Caption         =   "LC Master"
            Index           =   27
         End
         Begin VB.Menu mnuMasRefOp 
            Caption         =   "Power Master"
            Index           =   28
         End
         Begin VB.Menu mnuMasRefOp 
            Caption         =   "Sub-Packaging Master"
            Index           =   29
         End
         Begin VB.Menu mnuMasRefOp 
            Caption         =   "Costing Head"
            Index           =   30
         End
         Begin VB.Menu mnuMasRefOp 
            Caption         =   "Merge Master"
            Index           =   31
         End
      End
      Begin VB.Menu mnuMasterGrp 
         Caption         =   "Tax Refrence"
         Index           =   9
         Begin VB.Menu mnuTaxRef 
            Caption         =   "Tax Group"
            Index           =   0
         End
         Begin VB.Menu mnuTaxRef 
            Caption         =   "Sale Tax"
            Index           =   1
         End
         Begin VB.Menu mnuTaxRef 
            Caption         =   "Rate Factor"
            Index           =   2
         End
      End
      Begin VB.Menu mnuMasterGrp 
         Caption         =   "Bill Entry Setup"
         Index           =   10
         Begin VB.Menu mnuMasBESOp 
            Caption         =   "Sales"
            Index           =   0
         End
         Begin VB.Menu mnuMasBESOp 
            Caption         =   "Stock Inward"
            Index           =   1
         End
         Begin VB.Menu mnuMasBESOp 
            Caption         =   "Service Inward"
            Index           =   2
         End
      End
      Begin VB.Menu mnuMasterGrp 
         Caption         =   "Opening"
         Index           =   11
         Begin VB.Menu mnuOpn 
            Caption         =   "C-Form Receivable"
            Index           =   0
         End
         Begin VB.Menu mnuOpn 
            Caption         =   "C-Form Payable"
            Index           =   1
         End
      End
      Begin VB.Menu mnuMasterGrp 
         Caption         =   "-"
         Index           =   13
      End
      Begin VB.Menu mnuMasterGrp 
         Caption         =   "Change Financial Year"
         Index           =   14
      End
      Begin VB.Menu mnuMasterGrp 
         Caption         =   "Change Unit"
         Index           =   15
      End
      Begin VB.Menu mnuMasterGrp 
         Caption         =   "Log Off"
         Index           =   16
         Shortcut        =   ^{F2}
      End
      Begin VB.Menu mnuMasterGrp 
         Caption         =   "Exit"
         Index           =   17
      End
   End
   Begin VB.Menu mnuTransaction 
      Caption         =   " Transaction  "
      Begin VB.Menu mnuOrderBooking 
         Caption         =   "Order &Booking"
         Index           =   0
      End
      Begin VB.Menu mnuOrderBooking 
         Caption         =   "Order &Approval"
         Index           =   1
      End
      Begin VB.Menu mnuOrderBooking 
         Caption         =   "Delivery Order Schedule"
         Index           =   2
      End
      Begin VB.Menu mnuOrderBooking 
         Caption         =   "Delivery Order Approval"
         Index           =   3
      End
      Begin VB.Menu mnuOrderBooking 
         Caption         =   "Cancellation Order"
         Index           =   4
         Begin VB.Menu mnuordcanc 
            Caption         =   "Cancellation of Order (Partially)"
            Index           =   0
         End
         Begin VB.Menu mnuordcanc 
            Caption         =   "Cancellation of Order (Fully)"
            Index           =   1
         End
      End
      Begin VB.Menu mnuOrderBooking 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuOrderBooking 
         Caption         =   "GRN of Material Without A/c Effect"
         Index           =   6
      End
      Begin VB.Menu mnuOrderBooking 
         Caption         =   "GRN of Raw Material"
         Index           =   7
      End
      Begin VB.Menu mnuOrderBooking 
         Caption         =   "GRN for Inhouse JobWork "
         Index           =   8
      End
      Begin VB.Menu mnuOrderBooking 
         Caption         =   "GRN of Processed job"
         Index           =   9
      End
      Begin VB.Menu mnuOrderBooking 
         Caption         =   "GRN Services"
         Index           =   10
      End
      Begin VB.Menu mnuOrderBooking 
         Caption         =   "-"
         Index           =   11
      End
      Begin VB.Menu mnuOrderBooking 
         Caption         =   "Issue From Store"
         Index           =   12
      End
      Begin VB.Menu mnuOrderBooking 
         Caption         =   "Return To Store"
         Index           =   13
      End
      Begin VB.Menu mnuOrderBooking 
         Caption         =   "Outward for Job/RGP/NRGP"
         Index           =   14
      End
      Begin VB.Menu mnuOrderBooking 
         Caption         =   "Manual Clearance of Job/RGP"
         Index           =   15
      End
      Begin VB.Menu mnuOrderBooking 
         Caption         =   "Manual Clearance of Job GRN"
         Index           =   16
      End
      Begin VB.Menu mnuOrderBooking 
         Caption         =   "Box Wise Inventory"
         Index           =   17
         Begin VB.Menu mnuBox 
            Caption         =   "GRN  (Box Wise)"
            Index           =   0
         End
         Begin VB.Menu mnuBox 
            Caption         =   "Issue From Store (Box Wise)"
            Index           =   1
         End
         Begin VB.Menu mnuBox 
            Caption         =   "Return From Store (Box Wise)"
            Index           =   2
         End
         Begin VB.Menu mnuBox 
            Caption         =   "Direct Sale of Box"
            Index           =   3
         End
      End
      Begin VB.Menu mnuOrderBooking 
         Caption         =   "-"
         Index           =   18
      End
      Begin VB.Menu mnuLumpsum 
         Caption         =   "LumpSum Packing"
         Index           =   0
         Begin VB.Menu mnuLumpsumPacking 
            Caption         =   "Fresh/Export/GR/Job/Captive"
            Index           =   0
         End
         Begin VB.Menu mnuLumpsumPacking 
            Caption         =   "Wastage"
            Index           =   1
         End
         Begin VB.Menu mnuLumpsumPacking 
            Caption         =   "-"
            Index           =   2
         End
         Begin VB.Menu mnuLumpsumPacking 
            Caption         =   "Goods Return [GR] to Wastage"
            Index           =   3
         End
      End
      Begin VB.Menu mnuLumpsum 
         Caption         =   "LumpSum Dispatch"
         Index           =   1
         Begin VB.Menu mnuLumpsumDispatch 
            Caption         =   "Market/Export (With DO)"
            Index           =   0
         End
         Begin VB.Menu mnuLumpsumDispatch 
            Caption         =   "Market/Export/Job (Without DO)"
            Index           =   1
         End
         Begin VB.Menu mnuLumpsumDispatch 
            Caption         =   "Captive"
            Index           =   2
         End
         Begin VB.Menu mnuLumpsumDispatch 
            Caption         =   "Wastage"
            Index           =   3
         End
      End
      Begin VB.Menu mnuLumpsum 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuCarton 
         Caption         =   "Carton Packing"
         Index           =   0
         Begin VB.Menu mnuCartonPacking 
            Caption         =   "Fresh/Export/Job/Captive"
            Index           =   0
         End
         Begin VB.Menu mnuCartonPacking 
            Caption         =   "Update Packing"
            Index           =   1
         End
         Begin VB.Menu mnuCartonPacking 
            Caption         =   "-"
            Index           =   2
         End
         Begin VB.Menu mnuCartonPacking 
            Caption         =   "Goods Return [GR] to Fresh "
            Index           =   3
         End
      End
      Begin VB.Menu mnuCarton 
         Caption         =   "Carton Dispatch"
         Index           =   1
         Begin VB.Menu mnuCartonDispatch 
            Caption         =   "Market/Export (With DO)"
            Index           =   0
         End
         Begin VB.Menu mnuCartonDispatch 
            Caption         =   "Market/Export/Job (Without DO)"
            Index           =   1
         End
         Begin VB.Menu mnuCartonDispatch 
            Caption         =   "Captive"
            Index           =   2
         End
      End
      Begin VB.Menu mnuCarton 
         Caption         =   "Packing Against Order"
         Index           =   2
      End
      Begin VB.Menu mnuCarton 
         Caption         =   "Dispatch for Packed Material Against Order"
         Index           =   3
      End
      Begin VB.Menu mnuCarton 
         Caption         =   "Pallet Deletion"
         Index           =   4
      End
      Begin VB.Menu mnuCarton 
         Caption         =   "Pallet Updation"
         Index           =   5
      End
      Begin VB.Menu mnuCarton 
         Caption         =   "Finish Goods Return (GR) Entry"
         Index           =   6
      End
      Begin VB.Menu mnuCarton 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnuSale 
         Caption         =   "Sale Bill (With DO)"
         Index           =   0
      End
      Begin VB.Menu mnuSale 
         Caption         =   "Sale Bill (Without DO)"
         Index           =   1
      End
      Begin VB.Menu mnuSale 
         Caption         =   "Direct Sale"
         Index           =   2
      End
      Begin VB.Menu mnuSale 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuGatePass 
         Caption         =   "Gate Pass"
         Index           =   0
      End
      Begin VB.Menu mnuGatePass 
         Caption         =   "Vehicle Entry"
         Index           =   1
      End
      Begin VB.Menu mnuReturnable 
         Caption         =   "Returnable Metallic Cops / Pallet"
         Index           =   0
      End
      Begin VB.Menu mnuTransOp 
         Caption         =   "Transfer Challan"
         Index           =   17
      End
      Begin VB.Menu mnuTransOp 
         Caption         =   "Daily Power Consumption Entry"
         Index           =   18
      End
      Begin VB.Menu mnuTransOp 
         Caption         =   "Cops Managment"
         Index           =   19
         Begin VB.Menu copmnu 
            Caption         =   "Position Master"
            Index           =   0
         End
         Begin VB.Menu copmnu 
            Caption         =   "Allotment Of Mearge No."
            Index           =   1
         End
         Begin VB.Menu copmnu 
            Caption         =   "Cops Sticker Printing"
            Index           =   2
         End
      End
      Begin VB.Menu mnuTransOp 
         Caption         =   "Sale tax Form Collection"
         Index           =   20
      End
   End
   Begin VB.Menu mnuExcise 
      Caption         =   "Excise"
      Index           =   0
      Begin VB.Menu mnuExciseTrn 
         Caption         =   "TR-6 Challan"
         Index           =   0
      End
      Begin VB.Menu mnuExciseTrn 
         Caption         =   "Excise Credit Entry"
         Index           =   1
      End
      Begin VB.Menu mnuExciseTrn 
         Caption         =   "Excise Debit Entry"
         Index           =   2
      End
      Begin VB.Menu mnuExciseTrn 
         Caption         =   "Monthly Return (ER-1)"
         Index           =   3
      End
      Begin VB.Menu mnuExciseTrn 
         Caption         =   "Opening As-on Cut-Off Date"
         Index           =   4
      End
      Begin VB.Menu mnuExciseTrn 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu mnuExciseTrn 
         Caption         =   "Modvat Register"
         Index           =   11
      End
      Begin VB.Menu mnuExciseTrn 
         Caption         =   "RG-1 Register"
         Index           =   12
      End
      Begin VB.Menu mnuExciseTrn 
         Caption         =   "RG23-A Part I"
         Index           =   13
      End
      Begin VB.Menu mnuExciseTrn 
         Caption         =   "Excise Direct Credit Register"
         Index           =   14
      End
      Begin VB.Menu mnuExciseTrn 
         Caption         =   "Excise Direct Debit Register"
         Index           =   15
      End
      Begin VB.Menu mnuExciseTrn 
         Caption         =   "Payment Against Service Tax"
         Index           =   16
      End
      Begin VB.Menu mnuExciseTrn 
         Caption         =   "ER-1 Register"
         Index           =   17
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   " Reports  "
      Begin VB.Menu mnuReportOp 
         Caption         =   "Master Reports"
         Index           =   0
         Begin VB.Menu mnuReportMstOp 
            Caption         =   "• Group Master"
            Index           =   0
         End
         Begin VB.Menu mnuReportMstOp 
            Caption         =   "• Account &Master"
            Index           =   1
         End
         Begin VB.Menu mnuReportMstOp 
            Caption         =   "• Item Master"
            Index           =   2
         End
         Begin VB.Menu mnuReportMstOp 
            Caption         =   "• Reference Master"
            Index           =   3
         End
         Begin VB.Menu mnuReportMstOp 
            Caption         =   "• Item Group"
            Index           =   4
         End
         Begin VB.Menu mnuReportMstOp 
            Caption         =   "• Finish Item"
            Index           =   5
         End
         Begin VB.Menu mnuReportMstOp 
            Caption         =   "• Lot Master "
            Index           =   6
         End
      End
      Begin VB.Menu mnuReportOp 
         Caption         =   "Order Reports"
         Index           =   1
         Begin VB.Menu mnuReportOrderOP 
            Caption         =   "• Order Register"
            Index           =   0
         End
         Begin VB.Menu mnuReportOrderOP 
            Caption         =   "• Pending Order"
            Index           =   1
         End
         Begin VB.Menu mnuReportOrderOP 
            Caption         =   "• Booking Summary"
            Index           =   2
         End
         Begin VB.Menu mnuReportOrderOP 
            Caption         =   "• Order History"
            Index           =   3
         End
         Begin VB.Menu mnuReportOrderOP 
            Caption         =   "• Pending D.O"
            Index           =   4
         End
         Begin VB.Menu mnuReportOrderOP 
            Caption         =   "• Pending Order vs Stock"
            Index           =   5
         End
      End
      Begin VB.Menu mnuReportOp 
         Caption         =   "Dispatch Reports"
         Index           =   3
         Begin VB.Menu mnuReportOPDPF 
            Caption         =   "• Dispatch Register"
            Index           =   0
         End
         Begin VB.Menu mnuReportOPDPF 
            Caption         =   "• Dispatch Register With DO"
            Index           =   1
         End
         Begin VB.Menu mnuReportOPDPF 
            Caption         =   "• Order V/s Dispatch "
            Index           =   2
         End
         Begin VB.Menu mnuReportOPDPF 
            Caption         =   "• Production V/s Dispatch "
            Index           =   3
         End
         Begin VB.Menu mnuReportOPDPF 
            Caption         =   "• Lifting Report"
            Index           =   4
         End
         Begin VB.Menu mnuReportOPDPF 
            Caption         =   "• Dispatch Summary"
            Index           =   5
         End
      End
      Begin VB.Menu mnuReportOp 
         Caption         =   "Finish Material Stock Reports"
         Index           =   4
         Begin VB.Menu mnuReportOPStkRpt 
            Caption         =   "• Finish Stock (As On Date)"
            Index           =   0
         End
         Begin VB.Menu mnuReportOPStkRpt 
            Caption         =   "• Ageing Report"
            Index           =   1
         End
         Begin VB.Menu mnuReportOPStkRpt 
            Caption         =   "• Finish Item Ledger"
            Index           =   2
         End
         Begin VB.Menu mnuReportOPStkRpt 
            Caption         =   "• Finish Stock Periodic"
            Index           =   3
         End
         Begin VB.Menu mnuReportOPStkRpt 
            Caption         =   "-"
            Index           =   4
         End
         Begin VB.Menu mnuReportOPStkRpt 
            Caption         =   "• Wastage Report"
            Index           =   5
         End
      End
      Begin VB.Menu mnuReportOp 
         Caption         =   "Goods Return [GR] Register"
         Index           =   5
      End
      Begin VB.Menu mnuReportOp 
         Caption         =   "GRN"
         Index           =   9
         Begin VB.Menu mnuReportGRN 
            Caption         =   "GRN Register"
            Index           =   0
         End
         Begin VB.Menu mnuReportGRN 
            Caption         =   "GRN Register WO A/c Effect"
            Index           =   1
         End
      End
      Begin VB.Menu mnuPack 
         Caption         =   "Packing Report"
         Begin VB.Menu mnuPacking 
            Caption         =   "Pallet Packing Register"
            Index           =   0
         End
         Begin VB.Menu mnuPacking 
            Caption         =   "Packing Register"
            Index           =   1
         End
         Begin VB.Menu mnuPacking 
            Caption         =   "Packing Summary"
            Index           =   2
         End
         Begin VB.Menu mnuPacking 
            Caption         =   "Box History"
            Index           =   3
         End
         Begin VB.Menu mnuPacking 
            Caption         =   "Order V/s Production"
            Index           =   4
         End
         Begin VB.Menu mnuPacking 
            Caption         =   "GR Packing Register"
            Index           =   5
         End
         Begin VB.Menu mnuPacking 
            Caption         =   "Progressive Production Report"
            Index           =   6
         End
         Begin VB.Menu mnuPacking 
            Caption         =   "GR to Fresh Production Register"
            Index           =   7
         End
         Begin VB.Menu mnuPacking 
            Caption         =   "GR to Wastage Production Register"
            Index           =   8
         End
      End
      Begin VB.Menu mnujobwork 
         Caption         =   "Jobwork Reports"
         Begin VB.Menu mnujobtype 
            Caption         =   "Inhouse Jobwork Reports"
            Index           =   0
            Begin VB.Menu mnuInhouse 
               Caption         =   "JobWork Register"
               Index           =   0
            End
         End
         Begin VB.Menu mnujobtype 
            Caption         =   "Contract Jobwork Reports"
            Index           =   1
            Begin VB.Menu mnuContract 
               Caption         =   "Jobwork Register"
               Index           =   0
            End
            Begin VB.Menu mnuContract 
               Caption         =   "Issue Register"
               Index           =   1
            End
            Begin VB.Menu mnuContract 
               Caption         =   "Receive Register"
               Index           =   2
            End
            Begin VB.Menu mnuContract 
               Caption         =   "Joberwise Item Ledger"
               Index           =   3
            End
         End
      End
      Begin VB.Menu mnuRGP 
         Caption         =   "RGP / NRGP Reports"
         Begin VB.Menu mnuRGPNRGP 
            Caption         =   "RGP Register"
            Index           =   0
         End
         Begin VB.Menu mnuRGPNRGP 
            Caption         =   "RGP Issue Register"
            Index           =   1
         End
         Begin VB.Menu mnuRGPNRGP 
            Caption         =   "RGP Receive Register"
            Index           =   2
         End
         Begin VB.Menu mnuRGPNRGP 
            Caption         =   "Partywise Item Ledger"
            Index           =   3
         End
         Begin VB.Menu mnuRGPNRGP 
            Caption         =   "-"
            Index           =   4
         End
         Begin VB.Menu mnuRGPNRGP 
            Caption         =   "NRGP Issue Register"
            Index           =   5
         End
      End
      Begin VB.Menu mnuStore 
         Caption         =   "Store Item"
         Begin VB.Menu mnuStoreReport 
            Caption         =   "Issue Register"
            Index           =   0
         End
         Begin VB.Menu mnuStoreReport 
            Caption         =   "Ledger / Summary"
            Index           =   1
         End
         Begin VB.Menu mnuStoreReport 
            Caption         =   "Dead Stock"
            Index           =   2
         End
         Begin VB.Menu mnuStoreReport 
            Caption         =   "Item Category+Group Wise Stock Valuation Report"
            Index           =   3
         End
         Begin VB.Menu mnuStoreReport 
            Caption         =   "Division+Item Group Wise Consumption"
            Index           =   4
         End
         Begin VB.Menu mnuStoreReport 
            Caption         =   "Age Analysis along with value"
            Index           =   5
         End
         Begin VB.Menu mnuStoreReport 
            Caption         =   "Box Wise Stock"
            Index           =   6
         End
         Begin VB.Menu mnuStoreReport 
            Caption         =   "Specification Wise Report"
            Index           =   7
         End
         Begin VB.Menu mnuStoreReport 
            Caption         =   "Merge No. Wise Ledger"
            Index           =   8
         End
         Begin VB.Menu mnuStoreReport 
            Caption         =   "Store Return Register"
            Index           =   9
         End
      End
      Begin VB.Menu mnuWIP 
         Caption         =   "WIP Reports"
         Begin VB.Menu mnuWIPReport 
            Caption         =   "WIP Stock Status"
            Index           =   0
         End
         Begin VB.Menu mnuWIPReport 
            Caption         =   "WIP Adjustment"
            Index           =   1
         End
      End
      Begin VB.Menu mnuReturnableReg 
         Caption         =   "Returnable Pallet"
         Index           =   0
         Begin VB.Menu mnuPallet 
            Caption         =   "Receivable"
            Index           =   0
            Begin VB.Menu mnuPalletReport 
               Caption         =   "Issue Register"
               Index           =   0
            End
            Begin VB.Menu mnuPalletReport 
               Caption         =   "Received Register"
               Index           =   1
            End
         End
         Begin VB.Menu mnuPallet 
            Caption         =   "Payable"
            Index           =   1
            Begin VB.Menu mnuPalletRpt 
               Caption         =   "Issue Register"
               Index           =   0
            End
            Begin VB.Menu mnuPalletRpt 
               Caption         =   "Received Register"
               Index           =   1
            End
         End
         Begin VB.Menu mnuPallet 
            Caption         =   "Datewise Details"
            Index           =   2
         End
         Begin VB.Menu mnuPallet 
            Caption         =   "Agent+Party Wise Details"
            Index           =   3
         End
         Begin VB.Menu mnuPallet 
            Caption         =   "Party Group + Party Wise O/s Details"
            Index           =   4
         End
         Begin VB.Menu mnuPallet 
            Caption         =   "Partywise O/s Summary"
            Index           =   5
         End
      End
      Begin VB.Menu mnuReturnableReg 
         Caption         =   "Returnable Cops"
         Index           =   1
         Begin VB.Menu mnuCops 
            Caption         =   "Receivable"
            Index           =   0
            Begin VB.Menu mnuCopsReport 
               Caption         =   "Issue Register"
               Index           =   0
            End
            Begin VB.Menu mnuCopsReport 
               Caption         =   "Received Register"
               Index           =   1
            End
         End
         Begin VB.Menu mnuCops 
            Caption         =   "Payable"
            Index           =   1
            Begin VB.Menu mnuCopsRpt 
               Caption         =   "Issue Register"
               Index           =   0
            End
            Begin VB.Menu mnuCopsRpt 
               Caption         =   "Received Register"
               Index           =   1
            End
         End
         Begin VB.Menu mnuCops 
            Caption         =   "Datewise Summary"
            Index           =   2
            Visible         =   0   'False
         End
         Begin VB.Menu mnuCops 
            Caption         =   "Party AddressWise Summary"
            Index           =   3
         End
         Begin VB.Menu mnuCops 
            Caption         =   "Partywise Ledger"
            Index           =   4
         End
         Begin VB.Menu mnuCops 
            Caption         =   "Cops Inventory Report"
            Index           =   5
         End
      End
      Begin VB.Menu mnuTransport 
         Caption         =   "Transport Reports"
         Index           =   0
      End
      Begin VB.Menu mnuTransport 
         Caption         =   "Rate List"
         Index           =   1
      End
      Begin VB.Menu mnuTransport 
         Caption         =   "Excise Report"
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuGatePassReg 
         Caption         =   "Gate Pass Register"
         Index           =   0
      End
      Begin VB.Menu mnuVHCLEntryReg 
         Caption         =   "Vehicle Entry Register"
         Index           =   0
      End
      Begin VB.Menu mnuRateLifting 
         Caption         =   "Rate Realisation Report"
         Index           =   0
      End
      Begin VB.Menu mnuExciseRpt 
         Caption         =   "Excise Modvat Report"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuExciseRpt 
         Caption         =   "Power Consumption Report"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSaleReg 
         Caption         =   "Sale Register"
      End
      Begin VB.Menu mnuTaxForm 
         Caption         =   "Sale Tax Form"
         Index           =   0
         Begin VB.Menu mnuTaxFormTyp 
            Caption         =   "Tax Form Receivable"
            Index           =   0
         End
         Begin VB.Menu mnuTaxFormTyp 
            Caption         =   "Tax Form Payable"
            Index           =   1
         End
         Begin VB.Menu mnuTaxFormTyp 
            Caption         =   "DVAT 31"
            Index           =   2
         End
      End
   End
   Begin VB.Menu mnuDocPrint 
      Caption         =   "  Document Printing   "
      Begin VB.Menu mnuDocPrintOp 
         Caption         =   "Invoice Printing"
         Index           =   2
      End
      Begin VB.Menu mnuDocPrintOp 
         Caption         =   "Gatepass"
         Index           =   3
         Begin VB.Menu mnuDocPrintGP 
            Caption         =   "• Annexture for Jobwork"
            Index           =   0
         End
         Begin VB.Menu mnuDocPrintGP 
            Caption         =   "• Returnable Gatepass"
            Index           =   1
         End
         Begin VB.Menu mnuDocPrintGP 
            Caption         =   "• Non Returnable Gatepass"
            Index           =   2
         End
         Begin VB.Menu mnuDocPrintGP 
            Caption         =   "• Finish Gate Pass"
            Index           =   3
         End
         Begin VB.Menu mnuDocPrintGP 
            Caption         =   "• Returnable Pallets Gate Pass"
            Index           =   4
         End
      End
      Begin VB.Menu mnuDocPrintOp 
         Caption         =   "Loading Statement"
         Index           =   4
      End
      Begin VB.Menu mnuDocPrintOp 
         Caption         =   "Packing List"
         Index           =   6
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDocPrintOp 
         Caption         =   "Delivery Challan"
         Index           =   7
      End
      Begin VB.Menu mnuDocPrintOp 
         Caption         =   "Packing Slip"
         Index           =   8
         Begin VB.Menu mnuSubDocPrintOp 
            Caption         =   "Box"
            Index           =   0
         End
         Begin VB.Menu mnuSubDocPrintOp 
            Caption         =   "Pallet"
            Index           =   1
         End
      End
      Begin VB.Menu mnuDocPrintOp 
         Caption         =   "GRN Print"
         Index           =   9
      End
      Begin VB.Menu mnuDocPrintOp 
         Caption         =   "Form 402"
         Index           =   10
      End
      Begin VB.Menu mnuDocPrintOp 
         Caption         =   "Order Contract"
         Index           =   11
      End
      Begin VB.Menu mnuDocPrintOp 
         Caption         =   "Proforma Invoice"
         Index           =   12
      End
      Begin VB.Menu mnuDocPrintOp 
         Caption         =   "Cops Sticker"
         Index           =   13
      End
      Begin VB.Menu mnuDocPrintOp 
         Caption         =   "Location Wise Box Search"
         Index           =   14
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDocPrintOp 
         Caption         =   "Annexture IV"
         Index           =   15
      End
      Begin VB.Menu mnuDocPrintOp 
         Caption         =   "Annexture V"
         Index           =   16
      End
      Begin VB.Menu mnuDocPrintOp 
         Caption         =   "GR Packing Printing"
         Index           =   17
      End
      Begin VB.Menu mnuDocPrintOp 
         Caption         =   "LR Printing"
         Index           =   18
      End
   End
   Begin VB.Menu MNUsETUP 
      Caption         =   "   Setup   "
      Begin VB.Menu MNUsETUPOP 
         Caption         =   "&Edit Company"
         Index           =   0
      End
      Begin VB.Menu MNUsETUPOP 
         Caption         =   "&Create Company (Auth. Req.)"
         Index           =   1
      End
      Begin VB.Menu MNUsETUPOP 
         Caption         =   "User Setup"
         Index           =   2
      End
      Begin VB.Menu MNUsETUPOP 
         Caption         =   "New Financial Year"
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu MNUsETUPOP 
         Caption         =   "Change Password"
         Index           =   4
      End
      Begin VB.Menu MNUsETUPOP 
         Caption         =   "Unit Configuration"
         Index           =   5
      End
      Begin VB.Menu MNUsETUPOP 
         Caption         =   "Division Configuration"
         Index           =   6
         Visible         =   0   'False
      End
      Begin VB.Menu MNUsETUPOP 
         Caption         =   "Year End Process"
         Index           =   7
      End
   End
   Begin VB.Menu mnuMiscReports 
      Caption         =   "   Misc.   "
      Begin VB.Menu mnuMiscRepoOp1 
         Caption         =   "Daily Entry Status"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMiscRepoOp1 
         Caption         =   "Company Checklist"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMiscRepoOp1 
         Caption         =   "Compress Database"
         Index           =   2
      End
      Begin VB.Menu mnuMiscRepoOp1 
         Caption         =   "Menu Format"
         Index           =   3
         Begin VB.Menu mnuMenuFormat 
            Caption         =   "Windows Standard"
            Index           =   0
         End
         Begin VB.Menu mnuMenuFormat 
            Caption         =   "Others"
            Index           =   1
         End
      End
      Begin VB.Menu mnuMiscRepoOp1 
         Caption         =   "Backup/Restore  Data"
         Index           =   4
      End
      Begin VB.Menu mnuMiscRepoOp1 
         Caption         =   "User Define Reports"
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMiscRepoOp1 
         Caption         =   "User Defined BackGround"
         Index           =   6
      End
      Begin VB.Menu mnuMiscRepoOp1 
         Caption         =   "Clear Approved DO"
         Index           =   7
      End
      Begin VB.Menu mnuMiscRepoOp1 
         Caption         =   "Change Packing Type"
         Index           =   8
      End
      Begin VB.Menu mnuMiscRepoOp1 
         Caption         =   "W.I.P. Adjustment"
         Index           =   9
      End
      Begin VB.Menu mnuMiscRepoOp1 
         Caption         =   "Daily Entry Status"
         Index           =   10
      End
      Begin VB.Menu mnuMiscRepoOp1 
         Caption         =   "Invoice Ammendment"
         Index           =   11
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMiscRepoOp1 
         Caption         =   "Production Date Change"
         Index           =   12
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMiscRepoOp1 
         Caption         =   "Costing Report"
         Index           =   13
      End
      Begin VB.Menu mnuMiscRepoOp1 
         Caption         =   "Additional Info."
         Index           =   14
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuUtility 
      Caption         =   "   Utility   "
      Begin VB.Menu mnuUtilOp 
         Caption         =   "Cancelation of Invoice and Challan"
         Index           =   0
      End
      Begin VB.Menu mnuUtilOp 
         Caption         =   "Export CSV File"
         Index           =   1
      End
      Begin VB.Menu mnuUtilOp 
         Caption         =   "Bill Party Change"
         Index           =   2
      End
   End
   Begin VB.Menu mnuCustomReports 
      Caption         =   "Custom Reports"
      Visible         =   0   'False
      Begin VB.Menu CSMREP 
         Caption         =   "Sales Register"
         Index           =   0
      End
      Begin VB.Menu CSMREP 
         Caption         =   "GRN Register"
         Index           =   5
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuUserManager 
      Caption         =   "UserManager"
      Visible         =   0   'False
      Begin VB.Menu opUserManager 
         Caption         =   "Create New User"
         Index           =   0
      End
      Begin VB.Menu opUserManager 
         Caption         =   "Modify User Detail"
         Index           =   1
      End
      Begin VB.Menu opUserManager 
         Caption         =   "Delete User"
         Index           =   2
      End
   End
End
Attribute VB_Name = "frm_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents MenuEvents As CEvents
Attribute MenuEvents.VB_VarHelpID = -1

Private Sub copmnu_Click(INDEX As Integer)
  Select Case INDEX
  Case 0
   LOAD Frm_positionmaster
   Frm_positionmaster.Show
  Case 1
   LOAD frm_positiontran
   frm_positiontran.Show
  Case 2
   LOAD frm_copssticker
   frm_copssticker.Show
  End Select
End Sub

Private Sub CSMREP_Click(INDEX As Integer)
    
    Select Case INDEX
        Case 0
            RPTPARA = "SAL"
        Case 1
            RPTPARA = "PRM"
        Case 2
            RPTPARA = "RSL"
        Case 3
            RPTPARA = "RPR"
        Case 4
            RPTPARA = "PSR"
        Case 5
            RPTPARA = "IVR"
    End Select
    
    LOAD frm_CUSREP
    frm_CUSREP.Show
    
End Sub

Private Sub MDIForm_Activate()
BILLPRINTONLINE = False
On Error Resume Next
    If IsCompanyFound = False Then
        Dim Ctrl As Control
        For Each Ctrl In frm_Main
            If TypeOf Ctrl Is Menu Then
                Ctrl.Enabled = False
            End If
        Next
    End If
    
On Error GoTo 0
    Me.Caption = compNm + " " + "UNIT--> " + UntNm
End Sub

Private Sub MDIForm_Load()
On Error GoTo errLoad
Call SetScreen
   Call setAnimation
    Dim M_MENI As Menu
    
    If CN.State = 0 Then End
    
    Set RS = New Recordset
    
    RS.Open "SELECT * FROM COMPMAST", CN, adOpenDynamic, adLockOptimistic
    
    If RS.EOF = True Then
        'Load frm_New_Comp
        Dim Ctrl As Control
        For Each Ctrl In Me
            If TypeOf Ctrl Is Menu Then
                On Error Resume Next
                'Ctrl.Enabled = False
            End If
        Next
        frm_New_Comp.Show
        Exit Sub
    Else
        compPth = RS!COMP_PATH
        FSDT = RS!COMP_ACID
        FEDT = RS!comp_acfd
        Me.Show
        Dim rsFY As ADODB.Recordset
        Set rsFY = New Recordset
        'rsFY.Open "Select Distinct COMP_ACID,COMP_ACFD From COMPMAST Order By COMP_ACID,COMP_ACFD", CN, adOpenKeyset
        rsFY.Open "Select Distinct STFY AS COMP_ACID, ENFY AS COMP_ACFD From SERIALMASTER Order By STFY,ENFY", CN, adOpenKeyset
        
        
        If rsFY.RecordCount > 0 Then
            frm_FYrSelection.Show 1
        Else
            LOAD Frm_Login
        End If
    End If
    Timer1.Interval = 100
    PRINTCHAR
    
    On Error Resume Next

    'Loading custom report menu
    Call CustomReportMenu
    
    If GetSetting("FAS", "Color", "Format") = "Misc." Then
        gbSubClassMenu = True
        If gbSubClassMenu Then SubClassMenuXP
    End If
    
  Exit Sub

errLoad:
    ErrNumber = ERR.Number
    ErrMessage = ERR.Description
    frm_ErrorHandler.Show vbModal
End Sub

Private Sub MDIForm_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And Shift = 1 Then
        frmConnectServer.Show
    End If
    If Button = 1 And Shift = 1 Then
        frm_duplicatedo.Show
    End If
    If Button = 2 And Shift = 3 Then
        frmTransfer.Show 1
    End If
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    If IsCompanyFound And CompDtlUpd Then
        CompDtlUpd = False
        Frm_Selection.Show 1
        Cancel = True
    Else
        If IsLoggingOf Then Exit Sub
        
        Dim AYS
        AYS = MsgBox("Do You Want to Take Backup Before Exit ?", vbQuestion + vbYesNoCancel + vbDefaultButton2, "Want to Quit ?")
        If AYS = vbYes Then
            Cancel = True
              exit_bck = True
              LOAD Frm_backupdata
              Frm_backupdata.Show 1
              
        ElseIf AYS = VBNO Then
           CN.Execute "UPDATE USERMAST SET EXTRA1=NULL WHERE UID='" & cUName & "'"
           If cUName = "ADMIN" Then
             CN.Execute "UPDATE USERMAST SET EXTRA1=NULL"
           End If
                      
           For Each LastFrm In Forms
               Unload LastFrm
               Set LastFrm = Nothing
           Next
        Else
           Cancel = True
        End If
    End If
    
End Sub

Private Sub mnuBox_Click(INDEX As Integer)
Select Case INDEX
Case 0
     FRMBOXGRN.TABLENAME = "STORETRAN"
         FRMBOXGRN.SUMMARYTABLE = "GRN"
         LOAD FRMBOXGRN
         FRMBOXGRN.Show
Case 1
     LOAD frmBoxIss
     frmBoxIss.Show
     
Case 2
     LOAD frmRetBox
     frmRetBox.Show
Case 3
     LOAD frmBoxSale
     frmBoxSale.Show
End Select
End Sub

Private Sub mnuCarton_Click(INDEX As Integer)
Select Case INDEX
Case 2
    LOAD frmOrderPacking
    frmOrderPacking.Show
Case 3
    LOAD frmOrderDispatch
    frmOrderDispatch.Show
Case 4
    LOAD frmPalletDeletion
    frmPalletDeletion.Show
Case 5
    LOAD frmPalletEditing
    frmPalletEditing.Show
Case 6
    LOAD frmGRPacking
    frmGRPacking.Show
End Select
End Sub

Private Sub mnuCartonDispatch_Click(INDEX As Integer)
 Select Case INDEX
 Case 0
     If M_USRSECLEVL = "1" Then
          If ReadConfigMaster("000063", 4, "M") = False Then ModuleDeniedMessage: Exit Sub
     Else
          Call chk_ldt
     End If
 
     LOAD frmBoxDispatch
     frmBoxDispatch.Show
 Case 1
     LOAD frmJobDispatch
     frmJobDispatch.Show
 Case 2
     LOAD frmCaptiveBoxChallan
     frmCaptiveBoxChallan.Show
 End Select

 End Sub

Private Sub mnuCartonPacking_Click(INDEX As Integer)
Select Case INDEX
Case 0
If M_USRSECLEVL = "1" Then
       If ReadConfigMaster("000058", 4, "M") = False Then ModuleDeniedMessage: Exit Sub
  Else
       Call chk_ldt
 End If

     LOAD frmBoxPacking
     frmBoxPacking.Show
Case 1
If M_USRSECLEVL = "1" Then
       If ReadConfigMaster("000059", 4, "M") = False Then ModuleDeniedMessage: Exit Sub
    Else
       Call chk_ldt
    End If

     LOAD frmOpnBoxPacking
     frmOpnBoxPacking.Show
     
Case 3
If M_USRSECLEVL = "1" Then
       If ReadConfigMaster("000059", 4, "M") = False Then ModuleDeniedMessage: Exit Sub
    Else
       Call chk_ldt
    End If

     LOAD frmGRToFreshPacking
     frmGRToFreshPacking.Show
     
End Select
End Sub

Private Sub mnuContract_Click(INDEX As Integer)
Select Case INDEX
Case 0
   If M_USRSECLEVL = "1" Then
       If ReadConfigMaster("000054", 7, "R") = False Then ModuleDeniedMessage: Exit Sub
    Else
       Call chk_ldt
    End If

    LOAD frmRPT_ContractJobIssueReg
    frmRPT_ContractJobIssueReg.Tag = "XXX"
    frmRPT_ContractJobIssueReg.Show
Case 1
    If M_USRSECLEVL = "1" Then
       If ReadConfigMaster("000055", 7, "R") = False Then ModuleDeniedMessage: Exit Sub
    Else
       Call chk_ldt
    End If
    LOAD frmRPT_ContractJobIssueReg
    frmRPT_ContractJobIssueReg.Tag = "ANX"
    frmRPT_ContractJobIssueReg.Show
Case 2
If M_USRSECLEVL = "1" Then
       If ReadConfigMaster("000056", 7, "R") = False Then ModuleDeniedMessage: Exit Sub
    Else
       Call chk_ldt
    End If
    LOAD frmRPT_ContractJobIssueReg
    frmRPT_ContractJobIssueReg.Tag = "IVR3"
    frmRPT_ContractJobIssueReg.Show
Case 3
If M_USRSECLEVL = "1" Then
       If ReadConfigMaster("000057", 7, "R") = False Then ModuleDeniedMessage: Exit Sub
    Else
       Call chk_ldt
    End If
    LOAD frmRPT_ContractJobIssueReg
    frmRPT_ContractJobIssueReg.Tag = "JLD"
    frmRPT_ContractJobIssueReg.Show
End Select
End Sub

Private Sub mnuCops_Click(INDEX As Integer)

Select Case INDEX
Case 2
     frmRPT_ReturnableDatewiseSummary.PTYP = "COPS"
     frmRPT_ReturnableDatewiseSummary.Show
     frmRPT_ReturnableDatewiseSummary.Caption = "DATEWISE SUMMARY REPORT FOR RETURNABLE COPS"
Case 3
     frmRPT_ReturnableAddresswiseCopsSummary.Caption = "PARTY ADDRESS WISE SUMMARY REPORT FOR RETURNABLE COPS"
     frmRPT_ReturnableAddresswiseCopsSummary.Show
Case 4
     frmRPT_PartywiseReturnableCopsLedger.Caption = "PARTYWISE LEDGER REPORT FOR RETURNABLE COPS"
     frmRPT_PartywiseReturnableCopsLedger.Show
Case 5
     LOAD frmRPT_CopsInventory
     frmRPT_CopsInventory.Show
End Select

End Sub

Private Sub mnuCopsReport_Click(INDEX As Integer)
With frmRPT_ReturnableRegister
Select Case INDEX
Case 0
    .DBCR = "DB": .PTYP = "COPS": .RTYP = "ISS"
    .Show
    .Caption = "RETURNABLE RECEIVABLE COPS ISSUE REGISTER"
Case 1
    .DBCR = "DB": .PTYP = "COPS": .RTYP = "REC"
    .Show
    .Caption = "RETURNABLE RECEIVABLE COPS RECEIVE REGISTER"
End Select
End With
End Sub

Private Sub mnuCopsRpt_Click(INDEX As Integer)
With frmRPT_ReturnableRegister
Select Case INDEX
Case 0
    .DBCR = "CR": .PTYP = "COPS": .RTYP = "ISS"
    .Show
    .Caption = "RETURNABLE PAYABLE COPS ISSUE REGISTER"
Case 1
    .DBCR = "CR": .PTYP = "COPS": .RTYP = "REC"
    .Show
    .Caption = "RETURNABLE PAYABLE COPS RECEIVE REGISTER"
End Select
End With
End Sub

Private Sub mnuDocPrintGRN_Click(INDEX As Integer)
    LOAD frmRPT_GRNPrinting
    frmRPT_GRNPrinting.Show
End Sub





Private Sub mnuDocPrintGP_Click(INDEX As Integer)
Dim frmRPT_rgp As New frmRPT_ReturanableGP
    
    Select Case INDEX
        Case 0
            RPTPARA = "ANX"
        Case 1
            RPTPARA = "RGP"
        Case 2
            RPTPARA = "NGP"
        Case 3
            LOAD frmGP
            frmGP.Show
            Exit Sub
        Case 4
            LOAD frmRGP
            frmRGP.Show
            Exit Sub
        End Select
    
    LOAD frmRPT_rgp
    frmRPT_rgp.Show
    
    
End Sub

Private Sub mnuDocPrintOp_Click(INDEX As Integer)
    Select Case INDEX
        Case 2
            BILLPRINTONLINE = False
            LOAD frmRPT_InvPrinting
            frmRPT_InvPrinting.Show
        Case 4
            If UNT_LRONCHLN = "N" Then
                RPTPARA = "LOD"
                LOAD frmRPT_LoadStatement
                frmRPT_LoadStatement.Show
            Else
                LOAD frmRPT_PreLoadStatement
                frmRPT_PreLoadStatement.Show
            End If
        Case 6
            LOAD frmRPT_PackingList
            frmRPT_PackingList.Show
        Case 7
            LOAD frmRPT_DelChallanPrint
            frmRPT_DelChallanPrint.Show
        'Case 8
            'LOAD frm_PackingSlip
            'frm_PackingSlip.Show
        Case 9
             LOAD frmRPT_GRNPrinting
             frmRPT_GRNPrinting.Show
        Case 10
            LOAD frmRPT_Form402
            frmRPT_Form402.Show
        Case 11
            LOAD frm_orderfrm
            frm_orderfrm.Show
        Case 12
            LOAD FRM_rptperformainv
            FRM_rptperformainv.Show
        Case 13
            LOAD LBLFLE
            LBLFLE.Show
        Case 14
            RPTPARA = "SRC"
            LOAD frmRPT_LoadStatement
            frmRPT_LoadStatement.Show
        Case 15
            frmRPT_Anx.Tag = "4"
            LOAD frmRPT_Anx
            frmRPT_Anx.Show
        Case 16
            frmRPT_Anx.Tag = "5"
            LOAD frmRPT_Anx
            frmRPT_Anx.Show
        Case 17
            LOAD frmRPT_GRPackPrinting
            frmRPT_GRPackPrinting.Show
        Case 18
            LOAD frmRPT_LRPrinting
            frmRPT_LRPrinting.Show
    End Select
End Sub

Private Sub mnuExciseRpt_Click(INDEX As Integer)
Select Case INDEX
Case 0
  LOAD RPT_EXCISEREG
  RPT_EXCISEREG.Show
Case 1
  LOAD frmRptPowerCost
  frmRptPowerCost.Show
End Select
End Sub

Private Sub mnuExciseTrn_Click(INDEX As Integer)

Select Case INDEX
    Case 0
        If M_USRSECLEVL = "1" Then
           If ReadConfigMaster("000073", 4, "M") = False Then ModuleDeniedMessage: Exit Sub
        Else
           Call chk_ldt
        End If
        LOAD frmTR6
        frmTR6.Show
    Case 1
        If M_USRSECLEVL = "1" Then
           If ReadConfigMaster("000074", 4, "M") = False Then ModuleDeniedMessage: Exit Sub
        Else
           Call chk_ldt
        End If
        LOAD frmExcise
        frmExcise.Show
    Case 2
        If M_USRSECLEVL = "1" Then
           If ReadConfigMaster("000075", 4, "M") = False Then ModuleDeniedMessage: Exit Sub
        Else
           Call chk_ldt
        End If
        LOAD frm_servicetaxdb
        frm_servicetaxdb.Show
    Case 3
        If M_USRSECLEVL = "1" Then
           If ReadConfigMaster("000076", 4, "M") = False Then ModuleDeniedMessage: Exit Sub
        Else
           Call chk_ldt
        End If
        LOAD frm_er1
        frm_er1.Show
    Case 4
        If M_USRSECLEVL = "1" Then
           If ReadConfigMaster("000077", 4, "M") = False Then ModuleDeniedMessage: Exit Sub
        Else
           Call chk_ldt
        End If
        LOAD FRM_EXCADJ
        FRM_EXCADJ.Show
End Select

'FOR REPORTS
Select Case INDEX
    Case 11
        If M_USRSECLEVL = "1" Then
           If ReadConfigMaster("000078", 7, "R") = False Then ModuleDeniedMessage: Exit Sub
        Else
           Call chk_ldt
        End If
        LOAD RPT_EXCISEREG
        RPT_EXCISEREG.Show
    Case 12
        If M_USRSECLEVL = "1" Then
           If ReadConfigMaster("000079", 7, "R") = False Then ModuleDeniedMessage: Exit Sub
        Else
           Call chk_ldt
        End If
        LOAD frmRPT_RG1
        frmRPT_RG1.Show
    Case 13
        If M_USRSECLEVL = "1" Then
           If ReadConfigMaster("000080", 7, "R") = False Then ModuleDeniedMessage: Exit Sub
        Else
           Call chk_ldt
        End If
        LOAD frmrpt_rg23i
        frmrpt_rg23i.Show
    Case 14
        If M_USRSECLEVL = "1" Then
           If ReadConfigMaster("000081", 7, "R") = False Then ModuleDeniedMessage: Exit Sub
        Else
           Call chk_ldt
        End If
        LOAD Frm_exciseregister
        Frm_exciseregister.Show
    Case 15
        If M_USRSECLEVL = "1" Then
           If ReadConfigMaster("000082", 7, "R") = False Then ModuleDeniedMessage: Exit Sub
        Else
           Call chk_ldt
        End If
        LOAD frm_excisedbreg
        frm_excisedbreg.Show
    Case 16
        If M_USRSECLEVL = "1" Then
           If ReadConfigMaster("000083", 7, "R") = False Then ModuleDeniedMessage: Exit Sub
        Else
           Call chk_ldt
        End If
        LOAD FRM_PAYAGSRVTAX
        FRM_PAYAGSRVTAX.Show
    Case 17
        If M_USRSECLEVL = "1" Then
           If ReadConfigMaster("000084", 7, "R") = False Then ModuleDeniedMessage: Exit Sub
        Else
           Call chk_ldt
        End If
        LOAD frm_erp1rep
        frm_erp1rep.Show
        
End Select

End Sub

Private Sub mnuFinItmMst_Click(INDEX As Integer)
Select Case INDEX
Case 0
If M_USRSECLEVL = "1" Then
       If ReadConfigMaster("000007", 4, "M") = False Then ModuleDeniedMessage: Exit Sub
    Else
       Call chk_ldt
End If

   LOAD frm_FinItmMst
   frm_FinItmMst.Show
Case 1
If M_USRSECLEVL = "1" Then
       If ReadConfigMaster("000008", 4, "M") = False Then ModuleDeniedMessage: Exit Sub
    Else
       Call chk_ldt
End If
   LOAD FRM_GRDMST
   FRM_GRDMST.Show
Case 2
If M_USRSECLEVL = "1" Then
       If ReadConfigMaster("000008", 4, "M") = False Then ModuleDeniedMessage: Exit Sub
    Else
       Call chk_ldt
End If
   LOAD FRM_WGHTRANG
   FRM_WGHTRANG.Show
Case 3
If M_USRSECLEVL = "1" Then
       If ReadConfigMaster("000009", 4, "M") = False Then ModuleDeniedMessage: Exit Sub
    Else
       Call chk_ldt
End If

   LOAD frmLotMaster
   frmLotMaster.Show
Case 4
    If M_USRSECLEVL = "1" Then
       If ReadConfigMaster("000009", 4, "M") = False Then ModuleDeniedMessage: Exit Sub
    Else
       Call chk_ldt
    End If
   LOAD frm_mstlocation
   frm_mstlocation.Show
   
End Select
End Sub

Private Sub mnuGatePass_Click(INDEX As Integer)
Select Case INDEX
Case 0
     LOAD frmGatePass
     frmGatePass.Show
Case 1
      LOAD frmVehicleEntry
      frmVehicleEntry.Show
End Select
End Sub

Private Sub MNUGP_Click()
LOAD frmGP
frmGP.Show
End Sub

Private Sub mnuGatePassReg_Click(INDEX As Integer)
    LOAD frmRPT_GatePassReg
    frmRPT_GatePassReg.Show
End Sub

Private Sub mnuInhouse_Click(INDEX As Integer)
Select Case INDEX
Case 0
  LOAD frmRPT_InhouseGRNReg
  frmRPT_InhouseGRNReg.Show
End Select
End Sub

Private Sub mnuLumpsumDispatch_Click(INDEX As Integer)
Select Case INDEX
Case 0
     LOAD frmLumpSumDispatch
     frmLumpSumDispatch.Show
Case 1
     LOAD frmJobChallan
     frmJobChallan.Show
Case 2
     LOAD frmCaptiveChallan
     frmCaptiveChallan.Show
Case 3
     LOAD frmWastageDispatch
     frmWastageDispatch.Show
End Select
End Sub

Private Sub mnuLumpsumPacking_Click(INDEX As Integer)
Select Case INDEX
Case 0
    LOAD frmLumpSumPacking
    frmLumpSumPacking.Show
Case 1
     LOAD frmWastagePacking
     frmWastagePacking.Show
Case 3
     LOAD frmGRWastagePacking
     frmGRWastagePacking.Show
End Select
End Sub

Private Sub mnuMasBESOp_Click(INDEX As Integer)
    'DIVCOD = Empty
    'DIVNAM = Empty
    
    If M_USRSECLEVL = 1 Then
        If ReadConfigMaster("0007", 4, "M") = False Then ModuleDeniedMessage: Exit Sub
    End If
    
    Select Case INDEX
        Case 0
            LOAD frm_BillEntrySetupForSale
            frm_BillEntrySetupForSale.Show
            Exit Sub
        Case 1
            FRMPARA = "IVR"
        Case 2
            FRMPARA = "PSR"
    End Select
        
    Set LastFrm = New frm_BillEntrySetup
    
    LOAD LastFrm
    LastFrm.Tag = FRMPARA
    LastFrm.Show
End Sub

Private Sub mnuMasRefOp_Click(INDEX As Integer)
    Dim RefFrm As Form
   
    Select Case INDEX
        Case 0
            If M_USRSECLEVL = 1 Then
                If ReadConfigMaster("000012", 4, "M") = False Then ModuleDeniedMessage: Exit Sub
            End If
            Ref_Cat = "A"
        Case 1
            If M_USRSECLEVL = 1 Then
                If ReadConfigMaster("000013", 4, "M") = False Then ModuleDeniedMessage: Exit Sub
            End If
            Ref_Cat = "C"
        Case 2
            If M_USRSECLEVL = 1 Then
                If ReadConfigMaster("000014", 4, "M") = False Then ModuleDeniedMessage: Exit Sub
            End If
            Ref_Cat = "B"
        Case 3
            If M_USRSECLEVL = 1 Then
                If ReadConfigMaster("000015", 4, "M") = False Then ModuleDeniedMessage: Exit Sub
            End If
           LOAD frmTransportMaster
           frmTransportMaster.Show
           Exit Sub
        Case 6
            If M_USRSECLEVL = 1 Then
               If ReadConfigMaster("000016", 4, "M") = False Then ModuleDeniedMessage: Exit Sub
            End If
            frm_MACMST.Show
            Exit Sub
        Case 8
            If M_USRSECLEVL = 1 Then
                If ReadConfigMaster("000010", 4, "M") = False Then ModuleDeniedMessage: Exit Sub
            End If
            Ref_Cat = "U"
            
        Case 9
            If Not M_USRSECLEVL = 0 Then
                ModuleDeniedMessage
                Exit Sub
            End If
            LOAD frm_ChargesMaster
            frm_ChargesMaster.Show
            Exit Sub
        Case 12
            If Not M_USRSECLEVL = 0 Then
                ModuleDeniedMessage
                Exit Sub
            End If
            Ref_Cat = "M"
       
       Case 16
           If M_USRSECLEVL = "1" Then
             If ReadConfigMaster("000017", 4, "M") = False Then ModuleDeniedMessage: Exit Sub
            Else
             Call chk_ldt
           End If
           LOAD frmPkgStationMst
           frmPkgStationMst.Show
           Exit Sub
       Case 17
           If M_USRSECLEVL = "1" Then
             If ReadConfigMaster("000018", 4, "M") = False Then ModuleDeniedMessage: Exit Sub
            Else
             Call chk_ldt
           End If
           LOAD frmSalesManMst
           frmSalesManMst.Show
           Exit Sub
        
       Case 20
    If M_USRSECLEVL = "1" Then
       If ReadConfigMaster("000021", 4, "M") = False Then ModuleDeniedMessage: Exit Sub
    Else
       Call chk_ldt
    End If
       
           LOAD frmPackagingMaster
           frmPackagingMaster.Show
           Exit Sub
       Case 21
    If M_USRSECLEVL = "1" Then
       If ReadConfigMaster("000022", 4, "M") = False Then ModuleDeniedMessage: Exit Sub
    Else
       Call chk_ldt
    End If
       
           LOAD frmPlyMaster
           frmPlyMaster.Show
           Exit Sub
       Case 22
    If M_USRSECLEVL = "1" Then
       If ReadConfigMaster("000023", 4, "M") = False Then ModuleDeniedMessage: Exit Sub
    Else
       Call chk_ldt
    End If
           LOAD frmVehicleMaster
           frmVehicleMaster.Show
           Exit Sub
    Case 23
           LOAD FRM_BANKMST
           FRM_BANKMST.Show
           Exit Sub
    Case 24
           LOAD FRM_RATMST
           FRM_RATMST.Show
           Exit Sub
    Case 25
           LOAD frmTermsAndCondition
           frmTermsAndCondition.Show
           Exit Sub
    Case 26
           LOAD frmExportTypeMaster
           frmExportTypeMaster.Show
           Exit Sub
    Case 27
           LOAD FRM_LCMST
           FRM_LCMST.Show
           Exit Sub
    Case 28
           LOAD frmPowerMaster
           frmPowerMaster.Show
           Exit Sub
    Case 29
           LOAD frmSubPkgng
           frmSubPkgng.Show
           Exit Sub
    Case 30
           If M_USRSECLEVL = 1 Then
                If ReadConfigMaster("000012", 4, "M") = False Then ModuleDeniedMessage: Exit Sub
            End If
            Ref_Cat = "N"
            
    Case 31
           LOAD frmMergeMst
           frmMergeMst.Show
           Exit Sub
    End Select
    
    If INDEX <> 14 Then
      Set RefFrm = New Frm_Ref_FAS
    
      RefFrm.Tag = Ref_Cat
      RefFrm.Show
    End If
End Sub


Private Sub mnuMasterGrp_Click(INDEX As Integer)
Dim frm As Form
    Select Case INDEX
    Case 0  'Group Option
        
        
    Case 2
    If M_USRSECLEVL = "1" Then
       If ReadConfigMaster("000006", 4, "M") = False Then ModuleDeniedMessage: Exit Sub
    Else
       Call chk_ldt
    End If
            LOAD UNTMST
            UNTMST.Show
    Case 3
    If M_USRSECLEVL = "1" Then
       If ReadConfigMaster("000006", 4, "M") = False Then ModuleDeniedMessage: Exit Sub
    Else
       Call chk_ldt
    End If
            LOAD FRM_DIVMST
            FRM_DIVMST.Show
        
        Case 14
            frm_FYrSelection.Show 1
        Case 15
            For Each LastFrm In Forms
                If LastFrm.NAME <> Me.NAME Then
                    Unload LastFrm
                    Set LastFrm = Nothing
                End If
            Next
            frm_UnitSelction.Show 1
            Frm_Login.SetMenuVisibility
        Case 16
            CN.Execute "UPDATE USERMAST SET EXTRA1=NULL WHERE UID='" & cUName & "'"
            
            IsLoggingOf = True
            frm_Main.Enabled = False
            Frm_Login.Show 1
            Exit Sub
        Case 17
            
            For Each LastFrm In Forms
                If LastFrm.NAME <> Me.NAME Then
                    Unload LastFrm
                    Set LastFrm = Nothing
                End If
            Next
            
            Unload Me
            
            Set frm_Main = Nothing
    End Select
End Sub

Private Sub mnuMasterGrp1_Click(INDEX As Integer)
   Select Case INDEX
   Case 0
        If M_USRSECLEVL = "1" Then
            If ReadConfigMaster("000003", 4, "M") = False Then ModuleDeniedMessage: Exit Sub
           Else
            Call chk_ldt
        End If
        
        LOAD FRM_COUNTRY
        FRM_COUNTRY.Show
        Exit Sub
   Case 1
        If M_USRSECLEVL = "1" Then
            If ReadConfigMaster("000003", 4, "M") = False Then ModuleDeniedMessage: Exit Sub
           Else
            Call chk_ldt
        End If
        
        LOAD frm_statemaster
        frm_statemaster.Show
        Exit Sub
   Case 2
        If M_USRSECLEVL = "1" Then
            If ReadConfigMaster("000003", 4, "M") = False Then ModuleDeniedMessage: Exit Sub
           Else
            Call chk_ldt
        End If
        
        LOAD frm_citymaster
        frm_citymaster.Show
        Exit Sub
   Case 3
        If M_USRSECLEVL = "1" Then
            If ReadConfigMaster("000004", 4, "M") = False Then ModuleDeniedMessage: Exit Sub
           Else
            Call chk_ldt
        End If
        LOAD frm_Acc
        frm_Acc.Show
        Exit Sub
        
    Case 4
        If M_USRSECLEVL = "1" Then
            If ReadConfigMaster("000005", 4, "M") = False Then ModuleDeniedMessage: Exit Sub
           Else
            Call chk_ldt
        End If
        LOAD FrmDeliveryAddress
        FrmDeliveryAddress.Show
        Exit Sub
        
    Case 5
        LOAD frmexciseopening
        frmexciseopening.Show
        Exit Sub
   End Select
   
End Sub

Private Sub mnuMasterGrpOp_Click(INDEX As Integer)
    Select Case INDEX
     Case 0
       LOAD frm_schedulemaster
       frm_schedulemaster.Show
     Case 1
    If M_USRSECLEVL = "1" Then
       If ReadConfigMaster("000002", 4, "M") = False Then ModuleDeniedMessage: Exit Sub
    Else
       Call chk_ldt
    End If
    LOAD Frm_Grp
    Frm_Grp.Show
    End Select
End Sub
Public Sub SubClassMenuXP()

    '/ this code is made by MenuCreator add-in

    '/ prepare the caption for subclassing. Warning! Don't remove this comment!!!


    Set MenuEvents = New CEvents
    Set objMenuEx = New cMenuEx
'
    Call objMenuEx.Install(Me.hWnd, App.PATH & "\" & Me.NAME, ImgList, 3, MenuEvents)

End Sub

Private Sub mnuMasterItmOp_Click(INDEX As Integer)
    Select Case INDEX
        Case 0
            If M_USRSECLEVL = "1" Then
                If ReadConfigMaster("000010", 4, "M") = False Then ModuleDeniedMessage: Exit Sub
               Else
                Call chk_ldt
            End If
            LOAD frmCatMst
            frmCatMst.Show
        Case 1
            If M_USRSECLEVL = 1 Then
                If ReadConfigMaster("000010", 4, "M") = False Then ModuleDeniedMessage: Exit Sub
            End If
            
            LOAD FRM_IGRP
            FRM_IGRP.Show
        Case 2
            If M_USRSECLEVL = "1" Then
                If ReadConfigMaster("000010", 4, "M") = False Then ModuleDeniedMessage: Exit Sub
               Else
                Call chk_ldt
            End If
            
            LOAD frm_Item
            frm_Item.Show
   
        Case 5
            If M_USRSECLEVL = 1 Then
                If ReadConfigMaster("000011", 4, "M") = False Then ModuleDeniedMessage: Exit Sub
            End If
            
            LOAD frm_ItmOpen
            frm_ItmOpen.Show
        Case 11
            If M_USRSECLEVL = 1 Then
                If ReadConfigMaster("000011", 4, "M") = False Then ModuleDeniedMessage: Exit Sub
            End If
            'LOAD frm_mstlocation
            'frm_mstlocation.Show
            LOAD frm_DivMacItmOpen
            frm_DivMacItmOpen.Show
        
        Case 13
            If M_USRSECLEVL = 1 Then
                If ReadConfigMaster("000011", 4, "M") = False Then ModuleDeniedMessage: Exit Sub
            End If
            LOAD frmLocationMaster
            frmLocationMaster.Show
    End Select
End Sub

Private Sub mnuMenuFormat_Click(INDEX As Integer)
    Select Case INDEX
        Case 0
            Call SaveSetting("FAS", "Color", "Format", "Normal")
            gbSubClassMenu = False
    
        Case 1
            gbSubClassMenu = True
            If gbSubClassMenu Then SubClassMenuXP
            DoEvents
            
            If objMenuEx Is Nothing Then
                MsgBox "The MenuDesigner isn't available, because MenuExtended class is not active." & vbCrLf & _
                       "Click on XP toolbar button to activate subclassing with MenuExtended."
                Exit Sub
            End If
            
            '/ ---- IMPORTANT -----------------------------------------------
            '/ only MAIN MENU can use MenuDesigner!!!
            '/ --------------------------------------------------------------
            '/ MenuDesigner is available ONLY if all child window is close.
            '/ If fail MenuDesigner return False, otherwise True.
            '/
            '/ In a 'real' application need to check if end-user
            '/ try to open MenuDesigner while child windows is displayed.
            '/ the next code trap this eventuality
            If Not objMenuEx.MenuDesigner(Me.hWnd) Then
              MsgBox "Please, close all child windows.", vbExclamation, "Menu Designer isn't available"
              Exit Sub
            End If
            Call SaveSetting("FAS", "Color", "Format", "Misc.")
    End Select

End Sub

Private Sub mnuMerge_Click()
LOAD frmItemOpnMerge
frmItemOpnMerge.Show
End Sub

Private Sub mnuMiscRepoOp1_Click(INDEX As Integer)
        
    Select Case INDEX
        Case 0
            LOAD frmRpt_DailyStatus
            frmRpt_DailyStatus.Show
        Case 1
            If M_USRSECLEVL = "1" Then
                If ReadConfigMaster("0031", 4, "M") = False Then ModuleDeniedMessage: Exit Sub
               Else
                Call chk_ldt
            End If
            'frm_CheckList.Show
        Case 2
            If M_USRSECLEVL = "1" Then
                If ReadConfigMaster("0031", 4, "M") = False Then ModuleDeniedMessage: Exit Sub
               Else
                Call chk_ldt
            End If
            LOAD frmShrinkage
            frmShrinkage.Show
        Case 4
            'LOAD frmBackRest
            'frmBackRest.Show
            DBTYP = "MENU"
            LOAD frmDatabase
            frmDatabase.Show 1
            
            If DBTYP = "NEW" Then
               FRM_CREATEDATABASE.Show
               Unload FRM_CREATEDATABASE
            ElseIf DBTYP = "RESTORE" Then
               Frm_restoredata.Show 1
               Unload Frm_restoredata
            ElseIf DBTYP = "BACKUP" Then
               Frm_backupdata.Show 1
               Unload Frm_backupdata
            End If
            
        Case 5
            LOAD FRM_USRREP
            FRM_USRREP.Show
        Case 6
            LOAD FrmColorTest
            FrmColorTest.Show
       Case 7
            LOAD FRM_DOCLEAR
            FRM_DOCLEAR.Show
       Case 8
            LOAD frmPackingTransfer
            frmPackingTransfer.Show
       Case 9
            LOAD FRMWIPADJ
            FRMWIPADJ.Show
       Case 10
            If UCase(cUName) = "ADMIN" Then
               LOAD frmRPT_DailyEntry
               frmRPT_DailyEntry.Show
              Else
               Call ModuleDeniedMessage
               Exit Sub
            End If
       Case 11
          LOAD Frm_invoiceAmmendment
       Case 12
          LOAD frmProductionChange
       Case 13
          LOAD frmRPTCosting
          frmRPTCosting.Show
        Case 14
          LOAD frmAddress
          frmAddress.Show
    End Select
    
End Sub

Private Sub mnuOpn_Click(INDEX As Integer)
If M_USRSECLEVL = "1" Then
   If ReadConfigMaster("000010", 4, "M") = False Then ModuleDeniedMessage: Exit Sub
Else
   Call chk_ldt
End If
      
Select Case INDEX
Case 0
     LOAD frmOpeningCFormRecievable
     frmOpeningCFormRecievable.Show
Case 1
     LOAD frmOpeningCFormPayable
     frmOpeningCFormPayable.Show
End Select

End Sub

Private Sub mnuordcanc_Click(INDEX As Integer)
  Select Case INDEX
    Case 0
      If M_USRSECLEVL = "1" Then
         If ReadConfigMaster("000029", 4, "M") = False Then ModuleDeniedMessage: Exit Sub
       Else
         Call chk_ldt
      End If
      LOAD frmOrderReconcile
      frmOrderReconcile.Show
    Case 1
      If M_USRSECLEVL = "1" Then
         If ReadConfigMaster("000029", 4, "M") = False Then ModuleDeniedMessage: Exit Sub
       Else
         Call chk_ldt
      End If
      LOAD frmOrderClearing
      frmOrderClearing.Show
  End Select
End Sub

Private Sub mnuOrderBooking_Click(INDEX As Integer)
Select Case INDEX
Case 0
    If M_USRSECLEVL = "1" Then
       If ReadConfigMaster("000025", 4, "M") = False Then ModuleDeniedMessage: Exit Sub
    Else
       Call chk_ldt
    End If
    
        LOAD frm_ORDERBOOK
        frm_ORDERBOOK.Show
Case 1
    If M_USRSECLEVL = "1" Then
       If ReadConfigMaster("000026", 4, "M") = False Then ModuleDeniedMessage: Exit Sub
    Else
       Call chk_ldt
    End If
    LOAD frm_finapproval
    frm_finapproval.Show
Case 2
   If M_USRSECLEVL = "1" Then
       If ReadConfigMaster("000027", 4, "M") = False Then ModuleDeniedMessage: Exit Sub
    Else
       Call chk_ldt
    End If

    LOAD frmDeliverySchedule
    frmDeliverySchedule.Show
Case 3
 If M_USRSECLEVL = "1" Then
       If ReadConfigMaster("000028", 4, "M") = False Then ModuleDeniedMessage: Exit Sub
    Else
       Call chk_ldt
    End If
    LOAD FRM_DOAPPROVAL
    FRM_DOAPPROVAL.Show

Case 6
  If M_USRSECLEVL = "1" Then
       If ReadConfigMaster("000035", 4, "M") = False Then ModuleDeniedMessage: Exit Sub
    Else
       Call chk_ldt
  End If
    LOAD frmGRNWOAcEffect
    frmGRNWOAcEffect.Show
Case 7
  If M_USRSECLEVL = "1" Then
       If ReadConfigMaster("000035", 4, "M") = False Then ModuleDeniedMessage: Exit Sub
    Else
       Call chk_ldt
  End If
    FrmGRNEntry.TABLENAME = "STORETRAN"
    FrmGRNEntry.SUMMARYTABLE = "GRN"
    LOAD FrmGRNEntry
    FrmGRNEntry.Show
Case 8
  If M_USRSECLEVL = "1" Then
       If ReadConfigMaster("000036", 4, "M") = False Then ModuleDeniedMessage: Exit Sub
    Else
       Call chk_ldt
  End If
    FrmGRNEntry.TABLENAME = "JOBIN"
    FrmGRNEntry.SUMMARYTABLE = "JOBGRN"
    LOAD FrmGRNEntry
    FrmGRNEntry.Show
Case 9
  If M_USRSECLEVL = "1" Then
       If ReadConfigMaster("000037", 4, "M") = False Then ModuleDeniedMessage: Exit Sub
    Else
       Call chk_ldt
  End If
    LOAD FrmProcessedJob
    FrmProcessedJob.Show
Case 10
    If M_USRSECLEVL = "1" Then
       If ReadConfigMaster("000038", 4, "M") = False Then ModuleDeniedMessage: Exit Sub
    Else
       Call chk_ldt
    End If
    
    LOAD frmPurchaseServices
    frmPurchaseServices.Show
    
Case 12
    If M_USRSECLEVL = "1" Then
       If ReadConfigMaster("000039", 4, "M") = False Then ModuleDeniedMessage: Exit Sub
    Else
       Call chk_ldt
    End If
  
     LOAD frmStoreIssMerge
     frmStoreIssMerge.Show
Case 13
    If M_USRSECLEVL = "1" Then
       If ReadConfigMaster("000040", 4, "M") = False Then ModuleDeniedMessage: Exit Sub
    Else
       Call chk_ldt
    End If
      
    'LOAD frmReturnToStore
    'frmReturnToStore.Show
     LOAD frmRetStoreMerge
     frmRetStoreMerge.Show
Case 14
    If M_USRSECLEVL = "1" Then
       If ReadConfigMaster("000041", 4, "M") = False Then ModuleDeniedMessage: Exit Sub
    Else
       Call chk_ldt
    End If
    LOAD frmOutwardForJob
    frmOutwardForJob.Show
Case 15
    If M_USRSECLEVL = "1" Then
       If ReadConfigMaster("000042", 4, "M") = False Then ModuleDeniedMessage: Exit Sub
    Else
       Call chk_ldt
    End If
    LOAD Frm_ManualClearance
    Frm_ManualClearance.Show
Case 16
    If M_USRSECLEVL = "1" Then
       If ReadConfigMaster("000042", 4, "M") = False Then ModuleDeniedMessage: Exit Sub
    Else
       Call chk_ldt
    End If
    LOAD Frm_ManualClearanceGRN
    Frm_ManualClearanceGRN.Show
End Select
End Sub

Private Sub mnuPacking_Click(INDEX As Integer)
Select Case INDEX
Case 0
       LOAD frmRPT_PalletPackingRegister
       frmRPT_PalletPackingRegister.Show
Case 1
    If M_USRSECLEVL = "1" Then
       If ReadConfigMaster("000060", 7, "R") = False Then ModuleDeniedMessage: Exit Sub
    Else
       Call chk_ldt
    End If

            LOAD frmRPT_PackingRegister
            frmRPT_PackingRegister.Show
Case 2
    If M_USRSECLEVL = "1" Then
       If ReadConfigMaster("000061", 7, "R") = False Then ModuleDeniedMessage: Exit Sub
    Else
       Call chk_ldt
    End If
 
            LOAD FRMRPT_PCKSMRY
            FRMRPT_PCKSMRY.Show
Case 3
       LOAD FRMBOXHISTORY
       FRMBOXHISTORY.Show
Case 4
       LOAD frmRPT_OrderVsProduction
       frmRPT_OrderVsProduction.Show
Case 5
       LOAD frmRPT_GRPackingRegister
       frmRPT_GRPackingRegister.Show
Case 6
       LOAD frmRpt_Progss
       frmRpt_Progss.Show
Case 7
       LOAD frmRPT_GRFreshPackingRegister
       frmRPT_GRFreshPackingRegister.Show
Case 8
       LOAD frmRPT_GRWastePackingRegister
       frmRPT_GRWastePackingRegister.Show
End Select
End Sub

Private Sub mnuPallet_Click(INDEX As Integer)
With frmRPT_ReturnableDatewiseSummary
Select Case INDEX
Case 2
     .PTYP = "PALLET"
     .Show
     .Caption = "DATEWISE SUMMARY REPORT FOR RETURNABLE PALLET"
Case 3
     frmRPT_ReturnablePartyAgentwiseSummary.PTYP = "PALLET"
     frmRPT_ReturnablePartyAgentwiseSummary.Show
     frmRPT_ReturnablePartyAgentwiseSummary.Caption = "AGENT+PARTY WISE O/S DETAIL FOR RETURNABLE PALLET"
Case 4
     frmRPT_ReturnableGroupwiseSummary.RPT_TYP = "GROUP"
     frmRPT_ReturnableGroupwiseSummary.PTYP = "PALLET"
     frmRPT_ReturnableGroupwiseSummary.Show
     frmRPT_ReturnableGroupwiseSummary.Caption = "GROUP + PARTY WISE O/S DETAIL FOR RETURNABLE PALLET"
Case 5
     frmRPT_ReturnableGroupwiseSummary.RPT_TYP = "PARTY"
     frmRPT_ReturnableGroupwiseSummary.PTYP = "PALLET"
     frmRPT_ReturnableGroupwiseSummary.Show
     frmRPT_ReturnableGroupwiseSummary.Caption = "PARTY WISE O/S SUMMARY FOR RETURNABLE PALLET"
End Select
End With
End Sub

Private Sub mnuPalletReport_Click(INDEX As Integer)
With frmRPT_ReturnableRegister
Select Case INDEX
Case 0
    .DBCR = "DB": .PTYP = "PALLET": .RTYP = "ISS"
    .Show
    .Caption = "RETURNABLE RECEIVABLE PALLET ISSUE REGISTER"
Case 1
    .DBCR = "DB": .PTYP = "PALLET": .RTYP = "REC"
    .Show
    .Caption = "RETURNABLE RECEIVABLE PALLET RECEIVE REGISTER"
End Select
End With

End Sub

Private Sub mnuPalletRpt_Click(INDEX As Integer)
With frmRPT_ReturnableRegister
Select Case INDEX
Case 0
    .DBCR = "CR": .PTYP = "PALLET": .RTYP = "ISS"
    .Show
    .Caption = "RETURNABLE PAYABLE PALLET ISSUE REGISTER"
Case 1
    .DBCR = "CR": .PTYP = "PALLET": .RTYP = "REC"
    .Show
    .Caption = "RETURNABLE PAYABLE PALLET RECEIVE REGISTER"
End Select
End With

End Sub

Private Sub mnuRateLifting_Click(INDEX As Integer)
    LOAD frmRPT_RateLifting
    frmRPT_RateLifting.Show
End Sub

Private Sub mnuReportGRN_Click(INDEX As Integer)
Select Case INDEX
Case 0
  If M_USRSECLEVL = "1" Then
       If ReadConfigMaster("000043", 7, "R") = False Then ModuleDeniedMessage: Exit Sub
    Else
       Call chk_ldt
    End If

  LOAD frmRPT_GRNReg
  frmRPT_GRNReg.Show
Case 1
    If M_USRSECLEVL = "1" Then
       If ReadConfigMaster("000043", 7, "R") = False Then ModuleDeniedMessage: Exit Sub
    Else
       Call chk_ldt
    End If

  LOAD frmRPT_GRNWOAEReg
  frmRPT_GRNWOAEReg.Show
    
End Select

End Sub

Private Sub mnuReportMstOp_Click(INDEX As Integer)
 If M_USRSECLEVL = "1" Then
       If ReadConfigMaster("000024", 7, "R") = False Then ModuleDeniedMessage: Exit Sub
    Else
       Call chk_ldt
 End If

    Select Case INDEX
        Case 0
            frmRPT_GRPMaster.Show
        Case 1
            frmRpt_AccGroup.Show
        Case 2
            frmRPT_ItemMaster.Show
        Case 3
            frmRPT_References.Show
        Case 4
            frmRPT_ItemgrMaster.Show
        Case 5
        RPTPARA = "FIN"
            frmRPT_LOTMASTER.Show
        Case 6
        RPTPARA = "LOT"
            frmRPT_LOTMASTER.Show
    End Select
End Sub

Private Sub mnuReportOp_Click(INDEX As Integer)
Select Case INDEX
Case 5
  LOAD frmRPT_GoodsRetReg
  frmRPT_GoodsRetReg.Show
End Select
End Sub

Private Sub mnuReportOPDPF_Click(INDEX As Integer)
    Select Case INDEX
    Case 0
    If M_USRSECLEVL = "1" Then
       If ReadConfigMaster("000067", 7, "R") = False Then ModuleDeniedMessage: Exit Sub
    Else
       Call chk_ldt
    End If
             LOAD frmRPT_DispatchRegWithoutDO
             frmRPT_DispatchRegWithoutDO.Show
    Case 1
    If M_USRSECLEVL = "1" Then
       If ReadConfigMaster("000068", 7, "R") = False Then ModuleDeniedMessage: Exit Sub
    Else
       Call chk_ldt
    End If
             LOAD frmRPT_DispatchReg
             frmRPT_DispatchReg.Show
    Case 2
    If M_USRSECLEVL = "1" Then
       If ReadConfigMaster("000034", 7, "R") = False Then ModuleDeniedMessage: Exit Sub
    Else
       Call chk_ldt
    End If
             LOAD frmRPT_Last4MonthDispatch
             frmRPT_Last4MonthDispatch.Show
    Case 3
    If M_USRSECLEVL = "1" Then
       If ReadConfigMaster("000069", 7, "R") = False Then ModuleDeniedMessage: Exit Sub
    Else
       Call chk_ldt
    End If
            LOAD frmRPT_Last4MonthPPF_vs_DPF
            frmRPT_Last4MonthPPF_vs_DPF.Show
            
    Case 4
    If M_USRSECLEVL = "1" Then
       'If ReadConfigMaster("000069", 7, "R") = False Then ModuleDeniedMessage: Exit Sub
    Else
       'Call chk_ldt
    End If
            LOAD frmRPT_DailyDispatch
            frmRPT_DailyDispatch.Show
    Case 5
    If M_USRSECLEVL = "1" Then
       If ReadConfigMaster("000067", 7, "R") = False Then ModuleDeniedMessage: Exit Sub
    Else
       Call chk_ldt
    End If
             LOAD frm_despatchsummary
             frm_despatchsummary.Show
    End Select
        
End Sub

Private Sub mnuReportOPStkRpt_Click(INDEX As Integer)
If M_USRSECLEVL = "1" Then
  If ReadConfigMaster("000067", 7, "R") = False Then ModuleDeniedMessage: Exit Sub
Else
   Call chk_ldt
End If

    Select Case INDEX
        Case 0
            'As On Date Finished Material Stock Status Report
            LOAD frmRPT_FinishStkDate
            frmRPT_FinishStkDate.Show
        Case 1
            LOAD FRM_FINISTKAGE
            FRM_FINISTKAGE.Show
        Case 2
            LOAD frmRPT_FinishItemLedger
            frmRPT_FinishItemLedger.Show
        Case 3
          LOAD frmperiodicstockreport
          frmperiodicstockreport.Show
        Case 5
          LOAD frmRPT_WastageStockLedger
          frmRPT_WastageStockLedger.Show
    End Select
End Sub

Private Sub mnuReportOrderOP_Click(INDEX As Integer)
    Select Case INDEX
    Case 0
    If M_USRSECLEVL = "1" Then
       If ReadConfigMaster("000030", 7, "R") = False Then ModuleDeniedMessage: Exit Sub
    Else
       Call chk_ldt
    End If
            LOAD frmRPT_OrderReg
            frmRPT_OrderReg.Show
    Case 1
            LOAD frmRPT_PendingOrder
            frmRPT_PendingOrder.Show
    Case 2
    If M_USRSECLEVL = "1" Then
       If ReadConfigMaster("000031", 7, "R") = False Then ModuleDeniedMessage: Exit Sub
    Else
       Call chk_ldt
    End If
            LOAD frmRPT_BookingStatus
            frmRPT_BookingStatus.Show
    Case 3
    If M_USRSECLEVL = "1" Then
       If ReadConfigMaster("000032", 7, "R") = False Then ModuleDeniedMessage: Exit Sub
    Else
       Call chk_ldt
    End If
            LOAD frmRPT_OrderHistory
            frmRPT_OrderHistory.Show
    Case 4
    If M_USRSECLEVL = "1" Then
       If ReadConfigMaster("000033", 7, "R") = False Then ModuleDeniedMessage: Exit Sub
    Else
       Call chk_ldt
    End If
            LOAD frmRPT_DO_Pending
            frmRPT_DO_Pending.Show
    
    Case 5
            LOAD frmRPT_PendingOrderVsProduction
            frmRPT_PendingOrderVsProduction.Show
            
    End Select
End Sub

Private Sub mnuReturnable_Click(INDEX As Integer)
   LOAD frmReturnable
   frmReturnable.Show
End Sub

Private Sub mnuRGPNRGP_Click(INDEX As Integer)
Select Case INDEX
Case 0
   If M_USRSECLEVL = "1" Then
       If ReadConfigMaster("000050", 7, "R") = False Then ModuleDeniedMessage: Exit Sub
    Else
       Call chk_ldt
   End If
    LOAD frmRPT_ContractJobIssueReg
    frmRPT_ContractJobIssueReg.Tag = "YYY"
    frmRPT_ContractJobIssueReg.Show
Case 1
    If M_USRSECLEVL = "1" Then
       If ReadConfigMaster("000051", 7, "R") = False Then ModuleDeniedMessage: Exit Sub
    Else
       Call chk_ldt
   End If
    LOAD frmRPT_ContractJobIssueReg
    frmRPT_ContractJobIssueReg.Tag = "RGP"
    frmRPT_ContractJobIssueReg.Show
Case 2
If M_USRSECLEVL = "1" Then
       If ReadConfigMaster("000052", 7, "R") = False Then ModuleDeniedMessage: Exit Sub
    Else
       Call chk_ldt
   End If
    LOAD frmRPT_ContractJobIssueReg
    frmRPT_ContractJobIssueReg.Tag = "IVR4"
    frmRPT_ContractJobIssueReg.Show
Case 3
If M_USRSECLEVL = "1" Then
       If ReadConfigMaster("000053", 7, "R") = False Then ModuleDeniedMessage: Exit Sub
    Else
       Call chk_ldt
End If
    LOAD frmRPT_ContractJobIssueReg
    frmRPT_ContractJobIssueReg.Tag = "PLD"
    frmRPT_ContractJobIssueReg.Show
Case 5
If M_USRSECLEVL = "1" Then
       If ReadConfigMaster("000049", 7, "R") = False Then ModuleDeniedMessage: Exit Sub
    Else
       Call chk_ldt
    End If

    LOAD frmRPT_ContractJobIssueReg
    frmRPT_ContractJobIssueReg.Tag = "NGP"
    frmRPT_ContractJobIssueReg.Show
End Select
End Sub

Private Sub mnuSale_Click(INDEX As Integer)
Select Case INDEX
Case 0
    LOAD frmSale
    frmSale.Show
Case 1
    LOAD frmJobSale
    frmJobSale.Show
Case 2
    LOAD frm_Directsal
    frm_Directsal.Show
Case 3
    LOAD Frm_invoiceDeletion
    Frm_invoiceDeletion.Show
End Select
End Sub



Private Sub mnuSaleReg_Click()
LOAD frmRpt_saleRegistermfg
     frmRpt_saleRegistermfg.Show
End Sub

Private Sub MNUsETUPOP_Click(INDEX As Integer)
    
    Select Case INDEX
        Case 0
            If M_USRSECLEVL = "1" Then
               If ReadConfigMaster("0012", 4, "M") = False Then ModuleDeniedMessage: Exit Sub
            End If
            
            LOAD frm_New_Comp
            frm_New_Comp.Tag = "UPD"
            frm_New_Comp.LoadData
            frm_New_Comp.dtInstDate.Enabled = False
            frm_New_Comp.Show
            
        Case 1
            If M_USRSECLEVL = "1" Then
               If ReadConfigMaster("0012", 4, "M") = False Then ModuleDeniedMessage: Exit Sub
            End If
            
            LOAD frm_New_Comp
            frm_New_Comp.Tag = "NEW"
            frm_New_Comp.COMPCOD = frm_New_Comp.GenPath()
            frm_New_Comp.dtInstDate.Value = Now
            frm_New_Comp.Show
            
        Case 2
            If Not M_USRSECLEVL = 0 Then
                ModuleDeniedMessage
                Exit Sub
            End If
            frmUsers.Show
        Case 2
        
        Case 3
            If Not M_USRSECLEVL = 0 Then
                ModuleDeniedMessage
                Exit Sub
            End If
            'frm_CreateNewFY.Show
            LOAD frm_Createolddata
            frm_Createolddata.Show
        Case 4
            LOAD frm_ChangePwd
            frm_ChangePwd.Show
        Case 5
            LOAD FRM_UNITCFG
            FRM_UNITCFG.Show
        Case 6
            LOAD FRM_DIVCFG
            FRM_DIVCFG.Show
        Case 7
            LOAD frm_yearendprocess
            frm_yearendprocess.Show
    End Select
    
End Sub

Private Sub MNUsETUPOP1_Click(INDEX As Integer)
  Select Case INDEX
   Case 0
     LOAD frm_REPCFG
     frm_REPCFG.Show
   
  End Select
End Sub

Private Sub mnuStoreReport_Click(INDEX As Integer)
Select Case INDEX
Case 0
  If M_USRSECLEVL = "1" Then
       If ReadConfigMaster("000044", 7, "R") = False Then ModuleDeniedMessage: Exit Sub
    Else
       Call chk_ldt
    End If
  LOAD frmRPT_StoreItem
  frmRPT_StoreItem.Tag = "ISS"
  frmRPT_StoreItem.Show
Case 1
If M_USRSECLEVL = "1" Then
       If ReadConfigMaster("000045", 7, "R") = False Then ModuleDeniedMessage: Exit Sub
    Else
       Call chk_ldt
    End If
  LOAD frmRPT_StoreItem
  frmRPT_StoreItem.Tag = "IVR"
  frmRPT_StoreItem.Show
Case 2
 If M_USRSECLEVL = "1" Then
       If ReadConfigMaster("000046", 7, "R") = False Then ModuleDeniedMessage: Exit Sub
    Else
       Call chk_ldt
    End If

  LOAD frmRPT_AsOnDeadStock
  frmRPT_AsOnDeadStock.Show
  
Case 3
    If M_USRSECLEVL = "1" Then
       If ReadConfigMaster("000047", 7, "R") = False Then ModuleDeniedMessage: Exit Sub
    Else
       Call chk_ldt
    End If
    
    LOAD frmRPT_Last4MonthStockValuation
    frmRPT_Last4MonthStockValuation.Tag = "CAT"
    frmRPT_Last4MonthStockValuation.Show
    
Case 4
    If M_USRSECLEVL = "1" Then
       If ReadConfigMaster("000047", 7, "R") = False Then ModuleDeniedMessage: Exit Sub
    Else
       Call chk_ldt
    End If
    
    LOAD frmRPT_Last4MonthStockValuation
    frmRPT_Last4MonthStockValuation.Tag = "DIV"
    frmRPT_Last4MonthStockValuation.Show
Case 5
    'If M_USRSECLEVL = "1" Then
    '   If ReadConfigMaster("000047", 7, "R") = False Then ModuleDeniedMessage: Exit Sub
    'Else
    '   Call chk_ldt
    'End If
    
    LOAD frmRPT_FIFOAgeAnalysis
    frmRPT_FIFOAgeAnalysis.Show
    
Case 6
    LOAD FRMRPT_BOXSTOCK
    FRMRPT_BOXSTOCK.Show
    
Case 7
    LOAD frmRptSpecification
    frmRptSpecification.Show
Case 8
   LOAD frmRptMerge
   frmRptMerge.Show
Case 9
    LOAD frmRPT_StoreRetItem
    frmRPT_StoreRetItem.Show
End Select

End Sub

Private Sub mnuSubDocPrintOp_Click(INDEX As Integer)
Select Case INDEX
Case 0
     LOAD frm_PackingSlip
     frm_PackingSlip.Show
Case 1
     LOAD frm_PackingSlipForPallet
     frm_PackingSlipForPallet.Show
End Select
End Sub

Private Sub mnusubexcise_Click(INDEX As Integer)
Select Case INDEX
    Case 0
        LOAD frmTR6
        frmTR6.Show
    Case 1
        LOAD frmExcise
        frmExcise.Show
    Case 2
        LOAD frm_servicetaxdb
        frm_servicetaxdb.Show
  End Select
End Sub

Private Sub mnuTaxFormTyp_Click(INDEX As Integer)
   Select Case INDEX
   Case 0
    RPTPARA = "SAL"
    LOAD frmRPT_TaxCollection
    frmRPT_TaxCollection.Show
   Case 1
    RPTPARA = "PRM"
    LOAD frmRPT_TaxCollection
    frmRPT_TaxCollection.Show
   Case 2
    LOAD frmRPT_CST
    frmRPT_CST.Show
   End Select
End Sub

Private Sub mnuTaxRef_Click(INDEX As Integer)
Select Case INDEX
Case 0
   LOAD frmTaxGroupMaster
   frmTaxGroupMaster.Show
Case 1
     If M_USRSECLEVL = "1" Then
        If ReadConfigMaster("000019", 4, "M") = False Then ModuleDeniedMessage: Exit Sub
     Else
        Call chk_ldt
     End If
     LOAD FrmSaleTaxMaster
     FrmSaleTaxMaster.Show
     Exit Sub
Case 2
   If M_USRSECLEVL = "1" Then
      If ReadConfigMaster("000020", 4, "M") = False Then ModuleDeniedMessage: Exit Sub
   Else
      Call chk_ldt
   End If
   LOAD frmRatFactMst
   frmRatFactMst.Show
   Exit Sub
End Select
End Sub

Private Sub mnuTransOp_Click(INDEX As Integer)
Select Case INDEX
Case 17
       LOAD frmChallanTransfer
       frmChallanTransfer.Show
Case 18
       LOAD frmPowerTran
       frmPowerTran.Show
Case 20
     FRM_TAXENTRY.Tag = "SAL"
     FRMPARA = "SAL"
     LOAD FRM_TAXENTRY
     FRM_TAXENTRY.Show
End Select
End Sub

Private Sub mnuTransport_Click(INDEX As Integer)
Select Case INDEX
Case 0
    LOAD frmRPT_TransportReg
    frmRPT_TransportReg.Show
Case 1
    LOAD frmRPT_RateList
    frmRPT_RateList.Show
Case 2
    LOAD frmRPT_RG1
    frmRPT_RG1.Show
End Select
End Sub

Private Sub mnuUtilOp_Click(INDEX As Integer)
    
    Select Case INDEX
        Case 0
          If cUName = "ADMIN" Then
            LOAD Frm_invoiceDeletion
            Frm_invoiceDeletion.Show
          End If
        Case 1
            LOAD FRM_GENCSV
            FRM_GENCSV.Show
        Case 2
            LOAD frmPartyChange
            frmPartyChange.Show
    End Select
    
End Sub

Private Sub mnuVHCLEntryReg_Click(INDEX As Integer)
    LOAD frmRPT_VehicleEntry
    frmRPT_VehicleEntry.Show
End Sub


Private Sub mnuWIPReport_Click(INDEX As Integer)
Select Case INDEX
Case 0
  If M_USRSECLEVL = "1" Then
       If ReadConfigMaster("000048", 7, "R") = False Then ModuleDeniedMessage: Exit Sub
    Else
       Call chk_ldt
    End If

   LOAD frmRPT_WIPStockStatus
   frmRPT_WIPStockStatus.Show
Case 1
   LOAD frmRPT_ADJRegister
   frmRPT_ADJRegister.Show
End Select
End Sub



Private Sub opUtilMerge_Click(INDEX As Integer)
Dim frmMerging As frmUtil_MergeAc
    Select Case INDEX
        Case 0
            FRMPARA = "ACC"
        Case 1
            FRMPARA = "ITM"
    End Select
    Set frmMerging = New frmUtil_MergeAc
    
    frmMerging.Show
        
End Sub


Private Sub opUserManager_Click(INDEX As Integer)
    If (INDEX = 1 Or INDEX = 2) Then
        If UCase(frmUsers.tvUsers.SelectedItem.Text) = "ADMIN" Then
            MsgBox "Admin User Rights Can't Change !!", vbInformation, "Access Denied"
            Exit Sub
        End If
    End If
    Select Case INDEX
        Case 0
            frmUsers.IsCreatingNew = True
            frm_UserCreation.Show 1
            frm_UserCreation.Tag = "Creating"
            frmUsers.FillTree
        Case 1
            frmUsers.IsCreatingNew = False
            frm_UserCreation.Tag = "Modifying"
            frm_UserCreation.Show 1
        Case 2
            If cUName = frmUsers.tvUsers.SelectedItem.Text Then
                MsgBox "User Is Active !! Please Logon as Super User and Try Again", vbInformation, "Access Denied"
                Exit Sub
            End If
            If MsgBox("Are You Sure ? Want To Remove This ID", vbQuestion + vbYesNo, "Remove This ID") = vbYes Then
                CN.Execute "DELETE FROM USERMAST WHERE  COMP='" & compPth & "' AND UID='" & frmUsers.tvUsers.SelectedItem.Text & "'"
                CN.Execute "DELETE FROM USERRIGHTS WHERE COMP='" & compPth & "' AND USERCODE='" & frmUsers.tvUsers.SelectedItem.Text & "'"
                frmUsers.FillTree
            Else
                frmUsers.tvUsers.SetFocus
            End If
    End Select
End Sub

Private Sub Timer1_Timer()
    lblNow.Caption = Format(Now, "HH:MM")
    If Forms.COUNT > 1 Then picInfo.Visible = False Else picInfo.Visible = True
    If Forms.COUNT <= 1 Then Screen.MousePointer = vbNormal
    Dim SYSDATE As Date
    Dim SYSTIME As Date
    SYSDATE = Now - 1
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
   Select Case Button.INDEX
    Case 1
      If M_USRSECLEVL = "1" Then
        If ReadConfigMaster("0002", 4, "M") = False Then ModuleDeniedMessage: Exit Sub
       Else
        Call chk_ldt
      End If
      LOAD frm_Acc
      frm_Acc.Show
    Case 2
      If M_USRSECLEVL = "1" Then
        If ReadConfigMaster("0005", 4, "M") = False Then ModuleDeniedMessage: Exit Sub
       Else
        Call chk_ldt
      End If
      LOAD frm_Item
      frm_Item.Show
    Case 3
      If M_USRSECLEVL = 1 Then
        If ReadConfigMaster("0008", 4, "M") = False Then ModuleDeniedMessage: Exit Sub
      End If
      Ref_Cat = "B"
      LOAD Frm_Ref_FAS
      Frm_Ref_FAS.Show
    Case 4
     If Not M_USRSECLEVL = 0 Then
       ModuleDeniedMessage
       Exit Sub
     End If
     frmUsers.Show
    Case 5
     LOAD frm_REPCFG
     frm_REPCFG.Show
    Case 6
     If M_USRSECLEVL = "1" Then
        If ReadConfigMaster("0003", 4, "M") = False Then ModuleDeniedMessage: Exit Sub
       Else
        Call chk_ldt
     End If
     LOAD UNTMST
     UNTMST.Show
    Case 7
     LOAD frmBackRest
     frmBackRest.Show
    Case 9
     Shell "Calc.exe", vbNormalFocus
    Case 10
     If Not M_USRSECLEVL = 0 Then
       ModuleDeniedMessage
       Exit Sub
     End If
     LOAD frm_CMPCFG
     frm_CMPCFG.Show
   End Select
End Sub

Public Sub SetScreen()

Dim iWidth As Integer, iHeight As Integer
Dim Size As String
Dim ScreenName As String

iWidth = Screen.WIDTH \ Screen.TwipsPerPixelX
iHeight = Screen.Height \ Screen.TwipsPerPixelY
Size = iWidth & " X " & iHeight

If Size = "800 X 600" Then
   ScreenName = App.PATH & "\Graphics\821x513.JPG"
ElseIf Size = "1024 X 768" Then
   ScreenName = App.PATH & "\Graphics\1024x768.JPG"
ElseIf Size = "1152 X 864" Then
   ScreenName = App.PATH & "\Graphics\1152x864.JPG"
ElseIf Size = "1280 X 768" Then
   ScreenName = App.PATH & "\Graphics\1280x768.JPG"
ElseIf Size = "1366 X 768" Then
   ScreenName = App.PATH & "\Graphics\1366x768.JPG"
End If

If Dir(ScreenName, vbNormal) = Empty Then
   'ReportErrorMessage 1001
   'Exit Sub
Else
      frm_Main.Picture = LoadPicture(ScreenName)
End If

End Sub




Public Sub setAnimation()
Dim POS As Integer: POS = 0
Dim CNT As Integer
    
    For CNT = 1 To 14
      ANIMATE(CNT) = GetToken(Animation, CNT, ",")
    Next
    
       RED = ANIMATE(1)
       GREEN = ANIMATE(2)
       BLUE = ANIMATE(3)

       BRED = ANIMATE(4)
       BGREEN = ANIMATE(5)
       BBLUE = ANIMATE(6)

       FRED = ANIMATE(7)
       FGREEN = ANIMATE(8)
       FBLUE = ANIMATE(9)

       LBLRED = ANIMATE(10)
       LBLGREEN = ANIMATE(11)
       LBLBLUE = ANIMATE(12)
       
       FONTSZ = ANIMATE(13)
       isBold = ANIMATE(14)
        
End Sub


Function GetToken(ByVal StrVal As String, intIndex As Integer, _
    strDelimiter As String) As String

         Dim strSubString() As String
         Dim intIndex2 As Integer
         Dim i As Integer
         Dim intDelimitLen As Integer

         intIndex2 = 1
         i = 0
         intDelimitLen = Len(strDelimiter)

         Do While intIndex2 > 0
         
             ReDim Preserve strSubString(i + 1)
             
             intIndex2 = InStr(1, StrVal, strDelimiter)
         
             If intIndex2 > 0 Then
                 strSubString(i) = Mid(StrVal, 1, (intIndex2 - _
                    1))
                 StrVal = Mid(StrVal, (intIndex2 + _
                   intDelimitLen), Len(StrVal))
             Else
                 strSubString(i) = StrVal
             End If
             
             i = i + 1
             
         Loop

         If intIndex > (i + 1) Or intIndex < 1 Then
             GetToken = ""
         Else
             GetToken = strSubString(intIndex - 1)
         End If

End Function




