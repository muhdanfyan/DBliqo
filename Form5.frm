VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Form5 
   BorderStyle     =   0  'None
   Caption         =   "DATA HALAQAH"
   ClientHeight    =   9720
   ClientLeft      =   0
   ClientTop       =   15
   ClientWidth     =   8010
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form5.frx":0000
   ScaleHeight     =   9720
   ScaleWidth      =   8010
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox F2 
      Height          =   4700
      Left            =   120
      ScaleHeight     =   4635
      ScaleWidth      =   7635
      TabIndex        =   33
      Top             =   4800
      Visible         =   0   'False
      Width           =   7695
      Begin VB.Frame Frame3 
         Height          =   615
         Left            =   960
         TabIndex        =   37
         Top             =   0
         Width           =   6615
         Begin VB.TextBox Text1 
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   5640
            TabIndex        =   39
            Top             =   200
            Width           =   855
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   4080
            TabIndex        =   38
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label10 
            Caption         =   "Tahun"
            Height          =   255
            Left            =   5040
            TabIndex        =   41
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label9 
            Caption         =   "Bulan"
            Height          =   255
            Left            =   3600
            TabIndex        =   40
            Top             =   240
            Width           =   975
         End
      End
      Begin DBLiqo.jcbutton jcbutton5 
         Height          =   255
         Left            =   0
         TabIndex        =   34
         Top             =   480
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         ButtonStyle     =   0
         ShowFocusRect   =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "<"
         UseMaskCOlor    =   -1  'True
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "Form5.frx":31714
         Height          =   3900
         Left            =   10
         TabIndex        =   35
         ToolTipText     =   "Double Click Untuk melihat Data Ikhwah Secara Lengkap"
         Top             =   720
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   6879
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   27
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
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
         ColumnCount     =   5
         BeginProperty Column00 
            DataField       =   "NamaL"
            Caption         =   "Nama"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "Hp"
            Caption         =   "No Telp"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "Keaktifan"
            Caption         =   "Keaktifan"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "KKI"
            Caption         =   "KKI"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "wkt"
            Caption         =   "wkt"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
            BeginProperty Column02 
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   0
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   0
            EndProperty
         EndProperty
      End
      Begin DBLiqo.jcbutton jcbutton6 
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Top             =   480
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         ButtonStyle     =   0
         ShowFocusRect   =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ">"
         UseMaskCOlor    =   -1  'True
      End
   End
   Begin VB.PictureBox F1 
      Height          =   4700
      Left            =   120
      ScaleHeight     =   4635
      ScaleWidth      =   7635
      TabIndex        =   24
      Top             =   4800
      Width           =   7695
      Begin DBLiqo.jcbutton XPButton2 
         Height          =   615
         Left            =   1920
         TabIndex        =   25
         Top             =   120
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1085
         ButtonStyle     =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   15199212
         Caption         =   "Hapus Data"
         UseMaskCOlor    =   -1  'True
      End
      Begin DBLiqo.jcbutton XPButton3 
         Height          =   615
         Left            =   960
         TabIndex        =   26
         Top             =   120
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1085
         ButtonStyle     =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   15199212
         Caption         =   "Edit Ikhwah"
         UseMaskCOlor    =   -1  'True
      End
      Begin DBLiqo.jcbutton XPButton1 
         Height          =   615
         Left            =   0
         TabIndex        =   27
         Top             =   120
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1085
         ButtonStyle     =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   15199212
         Caption         =   "Data baru"
         UseMaskCOlor    =   -1  'True
      End
      Begin DBLiqo.jcbutton XPButton5 
         Height          =   255
         Left            =   7200
         TabIndex        =   28
         Top             =   480
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         ButtonStyle     =   0
         ShowFocusRect   =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "<"
         UseMaskCOlor    =   -1  'True
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "Form5.frx":31729
         Height          =   3900
         Left            =   0
         TabIndex        =   29
         ToolTipText     =   "Double Click Untuk melihat Data Ikhwah Secara Lengkap"
         Top             =   720
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   6879
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   27
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
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
         ColumnCount     =   22
         BeginProperty Column00 
            DataField       =   "NamaL"
            Caption         =   "NAMA LENGKAP"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "NamaKunn"
            Caption         =   "KUNNIYAH"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "Tempat"
            Caption         =   "Tempat"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "Tanggal"
            Caption         =   "Tanggal"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "AlamatD"
            Caption         =   "AlamatD"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "AlamatM"
            Caption         =   "ALAMAT LENGKAP"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "Pendidikan"
            Caption         =   "Pendidikan"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "Fak"
            Caption         =   "Fak"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column08 
            DataField       =   "NamaS"
            Caption         =   "NamaS"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column09 
            DataField       =   "Angkt"
            Caption         =   "Angkt"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column10 
            DataField       =   "Bkt"
            Caption         =   "Bkt"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column11 
            DataField       =   "USaud"
            Caption         =   "USaud"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column12 
            DataField       =   "JSaud"
            Caption         =   "JSaud"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column13 
            DataField       =   "RSD"
            Caption         =   "RSD"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column14 
            DataField       =   "RSMP"
            Caption         =   "RSMP"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column15 
            DataField       =   "RSMA"
            Caption         =   "RSMA"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column16 
            DataField       =   "POrg"
            Caption         =   "POrg"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column17 
            DataField       =   "Tingkat"
            Caption         =   "Tingkat"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column18 
            DataField       =   "NmHalaqa"
            Caption         =   "NmHalaqa"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column19 
            DataField       =   "Ayah"
            Caption         =   "Ayah"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column20 
            DataField       =   "Ibu"
            Caption         =   "Ibu"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column21 
            DataField       =   "hp"
            Caption         =   "TELP"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   0
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   0
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   0
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   0
            EndProperty
            BeginProperty Column05 
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   0
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   0
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   0
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   0
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   0
            EndProperty
            BeginProperty Column11 
               ColumnWidth     =   0
            EndProperty
            BeginProperty Column12 
               ColumnWidth     =   0
            EndProperty
            BeginProperty Column13 
               ColumnWidth     =   0
            EndProperty
            BeginProperty Column14 
               ColumnWidth     =   0
            EndProperty
            BeginProperty Column15 
               ColumnWidth     =   0
            EndProperty
            BeginProperty Column16 
               ColumnWidth     =   0
            EndProperty
            BeginProperty Column17 
               ColumnWidth     =   0
            EndProperty
            BeginProperty Column18 
               ColumnWidth     =   0
            EndProperty
            BeginProperty Column19 
               ColumnWidth     =   0
            EndProperty
            BeginProperty Column20 
               ColumnWidth     =   0
            EndProperty
            BeginProperty Column21 
            EndProperty
         EndProperty
      End
      Begin DBLiqo.jcbutton XpButton6 
         Height          =   255
         Left            =   7440
         TabIndex        =   30
         Top             =   480
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         ButtonStyle     =   0
         ShowFocusRect   =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ">"
         UseMaskCOlor    =   -1  'True
      End
      Begin DBLiqo.jcbutton XPButton7 
         Height          =   615
         Left            =   2880
         TabIndex        =   31
         Top             =   120
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1085
         ButtonStyle     =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   15199212
         Caption         =   "Naqib dan bendahara"
         UseMaskCOlor    =   -1  'True
      End
      Begin DBLiqo.jcbutton jcbutton1 
         Height          =   615
         Left            =   4680
         TabIndex        =   32
         Top             =   120
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1085
         ButtonStyle     =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   15199212
         Caption         =   "Keaktifan"
         UseMaskCOlor    =   -1  'True
      End
      Begin DBLiqo.jcbutton bsms 
         Height          =   615
         Left            =   5640
         TabIndex        =   51
         Top             =   120
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1085
         ButtonStyle     =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   15199212
         Caption         =   "SMS Ikhwah"
         UseMaskCOlor    =   -1  'True
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   5055
      Left            =   120
      TabIndex        =   23
      Top             =   4440
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   8916
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Data Ikhwah"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Keaktifan"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin DBLiqo.jcbutton XPButton4 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   7200
      TabIndex        =   19
      Top             =   40
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      ButtonStyle     =   0
      ShowFocusRect   =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "X"
      UseMaskCOlor    =   -1  'True
   End
   Begin DBLiqo.jcbutton Btn2 
      Height          =   375
      Left            =   1080
      TabIndex        =   18
      Top             =   0
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      ButtonStyle     =   0
      ShowFocusRect   =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ">>"
      UseMaskCOlor    =   -1  'True
   End
   Begin DBLiqo.jcbutton Btn1 
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   0
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      ButtonStyle     =   0
      ShowFocusRect   =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "<<"
      UseMaskCOlor    =   -1  'True
   End
   Begin Crystal.CrystalReport CRKKIIkh 
      Left            =   2760
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.PictureBox XPFrame1 
      Height          =   3975
      Left            =   120
      ScaleHeight     =   3915
      ScaleWidth      =   7740
      TabIndex        =   0
      Top             =   480
      Width           =   7800
      Begin VB.Frame Frame1 
         Caption         =   "SMS KKI"
         Height          =   2295
         Left            =   2400
         TabIndex        =   42
         Top             =   480
         Visible         =   0   'False
         Width           =   3375
         Begin VB.CheckBox Check1 
            Caption         =   "SMS NAQIB"
            Enabled         =   0   'False
            Height          =   255
            Left            =   480
            TabIndex        =   48
            Top             =   1200
            Width           =   1700
         End
         Begin VB.CheckBox Check2 
            Caption         =   "SMS BENDAHARA"
            Enabled         =   0   'False
            Height          =   375
            Left            =   480
            TabIndex        =   47
            Top             =   840
            Width           =   1700
         End
         Begin VB.CheckBox Check3 
            Caption         =   "SMS USTADZ"
            Height          =   255
            Left            =   240
            TabIndex        =   46
            Top             =   1800
            Width           =   1575
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Semua Ikhwah"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   44
            Top             =   360
            Width           =   2535
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Naqib dan bendahara"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   43
            Top             =   600
            Width           =   2535
         End
         Begin DBLiqo.jcbutton button2 
            Height          =   375
            Left            =   2250
            TabIndex        =   45
            Top             =   1900
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
            ButtonStyle     =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   12632064
            Caption         =   "TUTUP"
            UseMaskCOlor    =   -1  'True
         End
         Begin DBLiqo.jcbutton button3 
            Height          =   375
            Left            =   2250
            TabIndex        =   49
            Top             =   1550
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
            ButtonStyle     =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   12632064
            Caption         =   "KIRIM SMS"
            UseMaskCOlor    =   -1  'True
         End
      End
      Begin DBLiqo.jcbutton XPButton8 
         Height          =   615
         Left            =   6720
         TabIndex        =   22
         Top             =   3240
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1085
         ButtonStyle     =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   15199212
         Caption         =   "Cetak"
         UseMaskCOlor    =   -1  'True
      End
      Begin DBLiqo.jcbutton buttonSMS 
         Height          =   615
         Left            =   5760
         TabIndex        =   50
         Top             =   3240
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1085
         ButtonStyle     =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   15199212
         Caption         =   "SMS KKI"
         UseMaskCOlor    =   -1  'True
      End
      Begin VB.Label Label12 
         Height          =   495
         Left            =   2640
         TabIndex        =   53
         Top             =   3600
         Width           =   735
      End
      Begin VB.Label Label11 
         Height          =   495
         Left            =   1800
         TabIndex        =   52
         Top             =   3600
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "TAHUN"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         Height          =   255
         Index           =   8
         Left            =   2280
         TabIndex        =   20
         Top             =   1320
         Width           =   5055
      End
      Begin VB.Label Label16 
         Caption         =   "MARHALAH"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         Height          =   255
         Index           =   1
         Left            =   2280
         TabIndex        =   15
         Top             =   600
         Width           =   5055
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         Height          =   255
         Index           =   7
         Left            =   2280
         TabIndex        =   14
         Top             =   3120
         Width           =   5055
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         Height          =   255
         Index           =   6
         Left            =   2280
         TabIndex        =   13
         Top             =   2760
         Width           =   5055
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         Height          =   255
         Index           =   5
         Left            =   2280
         TabIndex        =   12
         Top             =   2400
         Width           =   5055
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         Height          =   255
         Index           =   4
         Left            =   2280
         TabIndex        =   11
         Top             =   2040
         Width           =   5055
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         Height          =   255
         Index           =   3
         Left            =   2280
         TabIndex        =   10
         ToolTipText     =   "Click Untuk Melihat Data Ust Murobbi"
         Top             =   1680
         Width           =   5055
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         Height          =   255
         Index           =   0
         Left            =   2280
         TabIndex        =   9
         Top             =   240
         Width           =   5055
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         Height          =   255
         Index           =   2
         Left            =   2280
         TabIndex        =   8
         Top             =   960
         Width           =   5055
      End
      Begin VB.Label Label7 
         Caption         =   "BENDAHARA"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   2400
         Width           =   1695
      End
      Begin VB.Label Label6 
         Caption         =   "TAHUN"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "WAKTU"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   3120
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "TEMPAT"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   2760
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "NAQIB"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "MUROBBI"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "NAMA KKI"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   1695
      End
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub ValData()
If Not rsKKI.EOF Then
    lbl(0).Caption = rsKKI!namaKKI
    lbl(1).Caption = rsKKI!Jenis
    lbl(2).Caption = rsKKI!Tahun
    lbl(3).Caption = rsKKI!murobbi
    lbl(4).Caption = rsKKI!Naqib
    lbl(5).Caption = rsKKI!Bendahara
    lbl(6).Caption = rsKKI!Tempat & " - KEC : " & rsKKI!Kecamatan
    lbl(7).Caption = rsKKI!Hari & "," & rsKKI!Waktu & " WITA"
    rsIkh.Filter = "NmHalaqa='" & rsKKI!namaKKI & "'"
    rsupIk.Filter = "KKI='" & rsKKI!namaKKI & "' AND wkt='" & Text1.text & "-" & Combo1.ListIndex + 1 & "'"
    aktif.Filter = rsupIk.Filter & "AND keaktifan='Aktif'"
    Taktif.Filter = rsupIk.Filter & "AND keaktifan='Tidak Aktif' or keaktifan='Kurang'"
 Else
    rsIkh.Filter = "NmHalaqa='~~~'"
    rsupIk.Filter = "KKI='~~~~'"
End If
End Sub

Private Sub bsms_Click()
    If Not rsIkh.EOF Then
        With Form14
            .Tab1.Tabs(3).Selected = True
            .List1.AddItem rsIkh!Hp
            .Show
        End With
    End If
End Sub

Private Sub Btn1_Click()
   If Not rsKKI.EOF Then
    If Not rsKKI.AbsolutePosition <= 1 Then rsKKI.MovePrevious
    ValData
   End If
End Sub

Private Sub Btn2_Click()
   If Not rsKKI.EOF Then
    If Not rsKKI.AbsolutePosition >= rsKKI.RecordCount Then rsKKI.MoveNext
    ValData
   End If
End Sub

Private Sub button2_Click()
    Frame1.Visible = False
End Sub

Private Sub button3_Click()
Dim i As Integer
    With Form14
     .List1.Clear
     rsIkh.Filter = ""
     .Tab1.Tabs.Item(3).Selected = True
     If Option1(1).Value = True Then
        If Check1.Value = 1 Then
            rsIkh.Find "namaL='" & rsKKI!Naqib & "'"
            If Not rsIkh.EOF Then .List1.AddItem rsIkh!Hp
        End If
        If Check2.Value = 1 Then
            
            rsIkh.Find "namaL='" & rsKKI!Bendahara & "'"
            If Not rsIkh.EOF Then .List1.AddItem rsIkh!Hp
        End If
     ElseIf Option1(0).Value = True Then
        rsIkh.Filter = "NmHalaqa='" & rsKKI!namaKKI & "'"
        
        If Not rsIkh.EOF Then
            rsIkh.MoveFirst
            For i = 1 To rsIkh.RecordCount
                .List1.AddItem rsIkh!Hp
                rsIkh.MoveNext
            Next i
        End If
     
     End If
        If Check3.Value = 1 Then
            rsUst.Requery
            rsUst.Find "nama='" & rsKKI!murobbi & "'"
            If Not rsUst.EOF Then .List1.AddItem rsUst!Hp
        End If
     
        .Show
    End With

    Frame1.Visible = False
End Sub

Private Sub buttonSMS_Click()
    Frame1.Visible = True
    Option1(0).Value = True
End Sub

Private Sub Combo1_Change()
    ValData
End Sub

Private Sub Combo1_Click()
        ValData
End Sub

Private Sub DataGrid1_DblClick()
    Form6.Show
End Sub

Private Sub Form_Load()
    Set DataGrid1.DataSource = rsIkh
    Set DataGrid2.DataSource = rsupIk
    With Combo1
        .Clear
        .AddItem "Januari"
        .AddItem "Februari"
        .AddItem "Maret"
        .AddItem "April"
        .AddItem "Mei"
        .AddItem "Juni"
        .AddItem "Juli"
        .AddItem "Agustus"
        .AddItem "September"
        .AddItem "Oktober"
        .AddItem "November"
        .AddItem "Desember"
    End With
    Combo1.ListIndex = Format(Date, "m") - 1
    Text1.text = Format(Date, "YYYY")
    ValData
End Sub

Private Sub Form_Unload(Cancel As Integer)
    rsIkh.Filter = ""
    rsupIk.Filter = ""
End Sub

Private Sub jcbutton1_Click()
If rsIkh.EOF = False Then
    With Form11
       .lbl(0).Caption = rsIkh!namaL
        .lbl(1).Caption = rsIkh!Hp
        .lbl(2).Caption = lbl(0).Caption
        .lbl(3).Caption = lbl(3).Caption
        .lbl(4).Caption = lbl(6).Caption
        .Show
    End With
End If
End Sub

Private Sub lbl_Click(Index As Integer)
    If Index = 3 Then
        If lbl(3).Caption <> "" Then
            rsUst.Find "nama='" & lbl(3).Caption & "'"
            Form9.Show
        End If
    End If
End Sub

Private Sub Option1_Click(Index As Integer)
    If Option1(0).Value = True Then
        Check1.Value = 1
        Check2.Value = 1
        Check1.Enabled = False
        Check2.Enabled = False
    ElseIf Option1(1).Value = True Then
        Check1.Enabled = True
        Check2.Enabled = True
    End If
End Sub

Private Sub TabStrip1_Click()
    If TabStrip1.Tabs(1).Selected Then
        F1.Visible = True
        F2.Visible = False
    ElseIf TabStrip1.Tabs(2).Selected Then
        F1.Visible = False
        F2.Visible = True
        ValData
    End If
End Sub

Private Sub Text1_Change()
    ValData
End Sub

Private Sub XPButton1_Click()
Dim a As String

a = MsgBox("Ikhwah yang Liqox Kosong?", vbYesNo + vbInformation, "Konfirmasi")
If a = vbYes Then
        With Form3
        .CbP.text = "NmHalaqa"
        .txtP.text = "-"
        .Show
    End With
    
Else
If Not rsKKI.EOF Then
    With Form4
        Baru = True
        .CbH1.text = lbl(0).Caption
        .CbH1.Enabled = False
        .lblT.Caption = lbl(1).Caption
        .lblT.Enabled = False
        .Show
    End With
End If
End If
End Sub

Private Sub XPButton3_Click()
If Not rsIkh.EOF Then
    Baru = False
    With Form4
        .txt(0).text = rsIkh!namaL
        .txt(1).text = rsIkh!NamaKunn
        .txt(2).text = rsIkh!Tempat
        If rsIkh!tanggal <> Empty Then .DT1.Value = rsIkh!tanggal
        .txt(3).text = rsIkh!AlamatD
        .txt(4).text = rsIkh!AlamatM
        .CbPen.text = rsIkh!Pendidikan
        .txt(17).text = rsIkh!NamaS
        .txt(18).text = rsIkh!jur
        .txt(5).text = rsIkh!Fak
        .txt(6).text = rsIkh!Angkt
        .txt(7).text = rsIkh!Usaud
        .txt(8).text = rsIkh!Jsaud
        .txt(9).text = rsIkh!RSD
        .txt(10).text = rsIkh!RSMP
        .txt(11).text = rsIkh!RSMA
        .txt(12).text = rsIkh!Bkt
        .txt(13).text = rsIkh!Porg
        .txt(14).text = rsIkh!Ayah
        .txt(15).text = rsIkh!Ibu
        .txt(16).text = rsIkh!Hp
        .CbH1.text = rsIkh!NmHalaqa
        .lblT.Caption = rsIkh!Tingkat
        .Show
    End With
End If
End Sub

Private Sub XPButton4_Click()
    Unload Me
End Sub

Private Sub XPButton5_Click()
If Not rsIkh.EOF Then If Not rsIkh.AbsolutePosition <= 1 Then rsIkh.MovePrevious
End Sub

Private Sub XPButton6_Click()
If Not rsIkh.EOF Then If Not rsIkh.AbsolutePosition >= rsIkh.RecordCount Then rsIkh.MoveNext
End Sub

Private Sub XPButton7_Click()
If Not rsKKI.EOF Then
    Baru = False
    With Form2
        .Text1(0).text = rsKKI!namaKKI
        .Text1(1).text = rsKKI!Tahun
        .Cbm.text = rsKKI!murobbi
        .CbNq.text = rsKKI!Naqib
        .CbJen.text = rsKKI!Tipe
        .CbBd.text = rsKKI!Bendahara
        .cb.text = rsKKI!Hari
        .Text1(5).text = rsKKI!Waktu
        .Text1(6).text = rsKKI!Tempat
        .Text1(7).text = rsKKI!Kecamatan
        .Cbt.text = rsKKI!Jenis
        .Text1(0).Enabled = False
        .Text1(1).Enabled = False
        .Cbm.Enabled = False
        .CbJen.Enabled = False
        .CbNq.Enabled = True
        .CbBd.Enabled = True
        .Text1(5).Enabled = False
        .Text1(6).Enabled = False
        .Text1(7).Enabled = False
        
        With .CbNq
            If rsKKI.EOF Then rsIkh.MoveFirst
            If Not rsIkh.EOF Then
                rsIkh.MoveFirst
                Do While Not rsIkh.EOF
                    .AddItem rsIkh!namaL
                    rsIkh.MoveNext
                Loop
                rsIkh.MoveFirst
            End If
        End With
        
        With .CbBd
            If rsKKI.EOF Then rsIkh.MoveFirst
            If Not rsIkh.EOF Then
                rsIkh.MoveFirst
                Do While Not rsIkh.EOF
                    .AddItem rsIkh!namaL
                    rsIkh.MoveNext
                Loop
            End If
        End With
        
        .Show
    End With
End If
End Sub

Private Sub XPButton8_Click()
If TabStrip1.Tabs(1).Selected Then
    With CRKKIIkh
        .ReportFileName = App.Path & "\RPT\RpKKIikh.rpt"
        .DataFiles(0) = App.Path & "\Db.mdb"
        .SelectionFormula = "{DBKKI.NamaKKI}='" & lbl(0).Caption & "' AND {upIkh.wkt}='" & Text1.text & "-" & Combo1.ListIndex + 1 & "'"
        '.ReplaceSelectionFormula "{DBIkhwah.NmHalaqa}='" & lbl(0).Caption & "'"
        .WindowState = crptMaximized
        .Action = 1
    End With
Else
    With CRKKIIkh
        .ReportFileName = App.Path & "\RPT\Rptaktikh.rpt"
        .DataFiles(0) = App.Path & "\Db.mdb"
        .SelectionFormula = "{DBKKI.NamaKKI}='" & lbl(0).Caption & "'"
        '.ReplaceSelectionFormula "{DBIkhwah.NmHalaqa}='" & lbl(0).Caption & "'"
        .WindowState = crptMaximized
        .Action = 1
    End With
End If
End Sub
