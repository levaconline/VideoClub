VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Statistika 
   BackColor       =   &H00C0C0C0&
   Caption         =   " Statistics"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   FillColor       =   &H000080FF&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "Statistika.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ListView ListView1 
      Height          =   4095
      Left            =   480
      TabIndex        =   27
      Top             =   3000
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   7223
      SortKey         =   1
      View            =   3
      Arrange         =   1
      SortOrder       =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Naslov"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Broj"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdodcArhiva 
      Height          =   330
      Left            =   360
      Top             =   8040
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Program Files\WinVideo\Data\VideoClub.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Program Files\WinVideo\Data\VideoClub.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "selest * from Arhiva where ID"
      Caption         =   "Arhiva"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc AdodcFilmovi 
      Height          =   330
      Left            =   2760
      Top             =   7320
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Program Files\WinVideo\Data\VideoClub.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Program Files\WinVideo\Data\VideoClub.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from filmovi order by Naslov"
      Caption         =   "Filmovi"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc AdodcClanovi 
      Height          =   330
      Left            =   360
      Top             =   7320
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Program Files\WinVideo\Data\VideoClub.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Program Files\WinVideo\Data\VideoClub.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from clanovi where ID"
      Caption         =   "Clanstvo"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0C0C0&
      Caption         =   " Brisanje arhive "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1215
      Left            =   4200
      TabIndex        =   17
      Top             =   5760
      Width           =   2775
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   600
         TabIndex        =   19
         ToolTipText     =   "Odaberite datum pre kojeg zelite da obrisete sve podatke iz arhive."
         Top             =   360
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   12632256
         Format          =   24641537
         CurrentDate     =   37020
      End
      Begin VB.CommandButton Command12 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Del"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Do:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0C0C0&
      Caption         =   " Different "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   2175
      Left            =   4200
      TabIndex        =   13
      Top             =   3360
      Width           =   2775
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFF00&
         Caption         =   "Print"
         Height          =   255
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Print"
         Top             =   240
         Width           =   495
      End
      Begin VB.ComboBox Combo4 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   26
         Text            =   "Godine"
         Top             =   1680
         Width           =   2535
      End
      Begin VB.ComboBox Combo3 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         TabIndex        =   16
         Text            =   "Zanrovi"
         Top             =   1200
         Width           =   2535
      End
      Begin VB.ComboBox Combo2 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         TabIndex        =   15
         Text            =   "Actors:"
         Top             =   720
         Width           =   2535
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         TabIndex        =   14
         Text            =   "Directed by:"
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Caption         =   " Profit "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   6495
      Left            =   7200
      TabIndex        =   10
      Top             =   720
      Width           =   3375
      Begin MSComCtl2.DTPicker DTPicker5 
         Height          =   255
         Left            =   120
         TabIndex        =   33
         ToolTipText     =   "Start Date"
         Top             =   1800
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         CalendarBackColor=   14737632
         Format          =   24641537
         CurrentDate     =   37045
      End
      Begin MSComCtl2.DTPicker DTPicker4 
         Height          =   255
         Left            =   1560
         TabIndex        =   32
         ToolTipText     =   "End Date"
         Top             =   1800
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         CalendarBackColor=   14737632
         Format          =   24641537
         CurrentDate     =   37045
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Refresh"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   2160
         Width           =   735
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "Statistika.frx":030A
         Height          =   4215
         Left            =   120
         TabIndex        =   22
         ToolTipText     =   "Double click to view"
         Top             =   2160
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   7435
         _Version        =   393216
         AllowUpdate     =   -1  'True
         BackColor       =   12632256
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Filmovi"
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   "ID"
            Caption         =   "ID"
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
            DataField       =   "Naslov"
            Caption         =   "Naslov"
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
               ColumnWidth     =   360
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2160
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Profit"
         Height          =   255
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1800
         Width           =   735
      End
      Begin VB.Shape Shape2 
         Height          =   615
         Left            =   120
         Top             =   1440
         Width           =   3135
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Between"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   1440
         Width           =   3135
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "and"
         Height          =   255
         Left            =   1080
         TabIndex        =   34
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Po Naslovu:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ukupno:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   24
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   23
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   1440
         TabIndex        =   12
         Top             =   960
         Width           =   1815
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   " List "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   4935
      Left            =   360
      TabIndex        =   7
      Top             =   2280
      Width           =   3495
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   4095
         Left            =   120
         TabIndex        =   21
         Top             =   720
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   7223
         _Version        =   393216
         BackColor       =   12632256
         HeadLines       =   1
         RowHeight       =   18
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
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
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
            DataField       =   ""
            Caption         =   ""
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
               ColumnWidth     =   434,835
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1335,118
            EndProperty
         EndProperty
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "List  Count"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1560
         TabIndex        =   9
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   " Top List 10 "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   2415
      Left            =   4200
      TabIndex        =   6
      Top             =   720
      Width           =   2775
      Begin VB.CommandButton Command8 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Show"
         Height          =   255
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   1920
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   255
         Left            =   360
         TabIndex        =   30
         ToolTipText     =   "End Date"
         Top             =   1200
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   14737632
         Format          =   24641537
         CurrentDate     =   37045
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   255
         Left            =   360
         TabIndex        =   29
         ToolTipText     =   "Start Date"
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   14737632
         Format          =   24641537
         CurrentDate     =   37045
      End
      Begin VB.CommandButton Command10 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Grafic"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "End Date"
         Height          =   255
         Left            =   1800
         TabIndex        =   40
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Start Date"
         Height          =   255
         Left            =   1800
         TabIndex        =   39
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "and"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   37
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Between"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   36
         Top             =   240
         Width           =   2295
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00FF0000&
         X1              =   360
         X2              =   240
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00FF0000&
         X1              =   360
         X2              =   240
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00FF0000&
         X1              =   240
         X2              =   240
         Y1              =   720
         Y2              =   1320
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00808080&
         FillColor       =   &H00C0C0C0&
         Height          =   2055
         Left            =   120
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   " General "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1335
      Left            =   360
      TabIndex        =   1
      Top             =   720
      Width           =   3495
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Movies  Count"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1440
         TabIndex        =   4
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Members  Count"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1440
         TabIndex        =   2
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FF0000&
      X1              =   4200
      X2              =   3840
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FF0000&
      X1              =   4080
      X2              =   4080
      Y1              =   1200
      Y2              =   2760
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FF0000&
      X1              =   3840
      X2              =   4080
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      X1              =   4080
      X2              =   4200
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "S T A T I S T I C S"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   4080
      TabIndex        =   0
      Top             =   120
      Width           =   4815
   End
End
Attribute VB_Name = "Statistika"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Smer As Integer 'Usmerivac za grafikone
Public Per As String
Public Dy1 As String, Dy2 As String 'Od - Do za grafikon
Option Explicit

Private Sub Combo1_Click()
ListView1.Visible = False
DataGrid1.Caption = "Filmovi"
Set DataGrid1.DataSource = AdodcFilmovi
AdodcFilmovi.RecordSource = "select * from filmovi where reziser= '" & Combo1.Text & "'"
AdodcFilmovi.Refresh
Label6.Caption = DataGrid1.ApproxCount
Label7.Caption = "Grid Count Aprox"
End Sub

Private Sub Combo2_Click()
ListView1.Visible = False
DataGrid1.Caption = "Filmovi"

Set DataGrid1.DataSource = AdodcFilmovi
AdodcFilmovi.RecordSource = "select * from filmovi where GlUloga= '" & Combo2.Text & "'"
AdodcFilmovi.Refresh
Label6.Caption = DataGrid1.ApproxCount
Label7.Caption = "Grid Count Aprox"

End Sub

Private Sub Combo3_Click()
ListView1.Visible = False
DataGrid1.Caption = "Filmovi"

Set DataGrid1.DataSource = AdodcFilmovi
AdodcFilmovi.RecordSource = "select * from filmovi where zanr= '" & Combo3.Text & "'"
AdodcFilmovi.Refresh
Label6.Caption = DataGrid1.ApproxCount
Label7.Caption = "Grid Count Aprox"

End Sub

Private Sub Combo4_Click()
ListView1.Visible = False
DataGrid1.Caption = "Filmovi"

Set DataGrid1.DataSource = AdodcFilmovi
AdodcFilmovi.RecordSource = "select * from filmovi where godina= " & Combo4.Text
AdodcFilmovi.Refresh
Label6.Caption = DataGrid1.ApproxCount
Label7.Caption = "Grid Count Aprox"
End Sub



Private Sub Command1_Click()
'Set DataReport1h.DataSource = AdodcFilmovi
DataReport1h.Show
End Sub

Private Sub Command10_Click()
Smer = 4
Load Grafikoni
End Sub

Private Sub Command12_Click()
Dim Dat As String
Dat = DTPicker1.Value
Dat = Format(Dat, "mm-dd-yyyy")
If MsgBox("Zelite da obrisete sve oodatke iz arhive do: " & Dat, vbYesNo + vbDefaultButton2, "Paznja!") = vbNo Then Exit Sub

AdodcArhiva.RecordSource = "select * from arhiva where datumrentir<  # " & Dat & " #"
AdodcArhiva.Refresh
While Not AdodcArhiva.Recordset.EOF
AdodcArhiva.Recordset.Delete
AdodcArhiva.Recordset.MoveNext
Wend
AdodcArhiva.RecordSource = "select * from arhiva "
AdodcArhiva.Refresh
Command12.Enabled = False
End Sub

Private Sub Command13_Click()
DataGrid2_DblClick
End Sub



Private Sub Command7_Click()
'Ostvareni profit u zadnjoj nedelji, prikaz u label8 (obrada podataka iz baze)
ListView1.Visible = False
Dim Bass As Integer, Naz As String, Skup As Integer
AdodcArhiva.RecordSource = "select * from arhiva"
AdodcArhiva.Refresh
Dim Dat As String, Dat1 As String, Lista As ListItem
Dat = Format(DTPicker5.Value, "mm-dd-yyyy")
Dat1 = Format(DTPicker4.Value, "mm-dd-yyyy")
Set DataGrid1.DataSource = AdodcArhiva
AdodcArhiva.RecordSource = "select * from arhiva where DatumRentir between #" & Dat & "# and  #" & Dat1 & "#"
AdodcArhiva.Refresh

' Proracuni
On Error Resume Next
AdodcFilmovi.RecordSource = "select * from filmovi"
Naz = AdodcFilmovi.Recordset.Fields(1)
While Not AdodcArhiva.Recordset.EOF
If Naz = AdodcArhiva.Recordset.Fields(1) Then
Bass = Bass + AdodcArhiva.Recordset.Fields(2)
End If
Skup = Skup + AdodcArhiva.Recordset.Fields(2)
AdodcArhiva.Recordset.MoveNext
Wend
Label9.Caption = (Skup * Glavni.RV) & " " & Chr$(163)
Label8.Caption = (Bass * Glavni.RV) & " " & Chr$(163)
DataGrid1.Caption = "Arhiva"
Label7.Caption = "Grid Count Aprox"
Label6.Caption = DataGrid1.ApproxCount
Command10.Enabled = False
End Sub

Private Sub Command8_Click()
ListView1.Visible = True

AdodcFilmovi.RecordSource = "select * from filmovi order by naslov"
ListView1.Visible = True
ListView1.ListItems.Clear

Dim Dat As String, Dat1 As String, Lista As ListItem
Dat = Format(DTPicker2.Value, "mm-dd-yyyy")
Dy1 = Dat
Dat1 = Format(DTPicker3.Value, "mm-dd-yyyy")
Dy2 = Dat1
AdodcArhiva.RecordSource = "select * from arhiva where DatumRentir between #" & Dat & "# and #" & Dat1 & "#"
Dim Popop As Long
    AdodcFilmovi.Refresh
    AdodcArhiva.Refresh
' Zbir dana za svaki film posebno
While Not AdodcFilmovi.Recordset.EOF
Set Lista = ListView1.ListItems.Add(, , AdodcFilmovi.Recordset.Fields("Naslov"))
        AdodcArhiva.Refresh
    While Not AdodcArhiva.Recordset.EOF
        If AdodcFilmovi.Recordset.Fields("Naslov") = AdodcArhiva.Recordset.Fields("Film") Then
            Popop = Val(Lista.SubItems(1))
            Lista.SubItems(1) = Val(Popop + AdodcArhiva.Recordset.Fields("Period"))
        End If
    AdodcArhiva.Recordset.MoveNext
    Wend
If Val(Lista.SubItems(1)) < 1 Then ListView1.ListItems.Remove (ListView1.ListItems.Count)

Popop = 0
AdodcFilmovi.Recordset.MoveNext
Wend

Dim Top(1 To 10) As Long, Cic As Long, Prip As Long, Zxf As Integer, Inde As Integer
Dim Lap(1 To 10) As String
On Error Resume Next
    For Zxf = 1 To 10
    Lap(Zxf) = ListView1.ListItems(1)
    Top(Zxf) = ListView1.ListItems(1).ListSubItems(1)
    Prip = 1
         For Cic = 1 To ListView1.ListItems.Count
            Set Lista = ListView1.ListItems(Cic)
            If Top(Zxf) < Val(Lista.ListSubItems(1)) Then
            Lap(Zxf) = Lista
            Top(Zxf) = Lista.ListSubItems(1)
            Prip = Cic
            End If
            
        Next Cic
          Set Lista = ListView1.ListItems(Prip)
        
          ListView1.ListItems.Remove (Prip)
        
    Next Zxf
    
ListView1.ListItems.Clear

For Zxf = 1 To 10
Set Lista = ListView1.ListItems.Add(, , Lap(Zxf))
Lista.SubItems(1) = Top(Zxf)
Next
Label6.Caption = 10
Label7.Caption = " Top List"
Command10.Enabled = True
End Sub

Public Sub DataGrid2_DblClick()
AdodcFilmovi.RecordSource = "select * from filmovi order by naslov"
AdodcFilmovi.Refresh
End Sub

Private Sub DTPicker1_Change()
Command12.Enabled = True
End Sub

Private Sub DTPicker1_Click()
Command12.Enabled = True

End Sub

Private Sub Form_Activate()
On Error Resume Next
AdodcClanovi.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\Data\VideoClub.mdb"
AdodcArhiva.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\Data\VideoClub.mdb"
AdodcFilmovi.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\Data\VideoClub.mdb"

' Upis broja clanova
Pocetna.adoClanovi.Recordset.MoveLast
Label3.Caption = Pocetna.adoClanovi.Recordset.RecordCount
' Upis broja filmova
Pocetna.adoFilmovi.Recordset.MoveLast
Label5.Caption = Pocetna.adoFilmovi.Recordset.RecordCount
' Popinjavanje combo1 - reziseri
Combo1.Clear
AdodcFilmovi.RecordSource = "select distinct Reziser from filmovi"
AdodcFilmovi.Refresh
While Not AdodcFilmovi.Recordset.EOF
If AdodcFilmovi.Recordset.Fields("Reziser") <> "-" Then
Combo1.AddItem (AdodcFilmovi.Recordset.Fields("Reziser"))
End If
AdodcFilmovi.Recordset.MoveNext
Wend
Combo1.Text = "Reziseri"

' Punjenje combo2 - glumci
Combo2.Clear
AdodcFilmovi.RecordSource = "select distinct GlUloga from filmovi"
AdodcFilmovi.Refresh
While Not AdodcFilmovi.Recordset.EOF
'If AdodcFilmovi.Recordset.Fields("GlUloga") <> "-" Then
Combo2.AddItem (AdodcFilmovi.Recordset.Fields("GlUloga"))
'End If
AdodcFilmovi.Recordset.MoveNext
Wend
Combo2.Text = "Glumci"

' Popunjavanje combo3 - zanr
Combo3.Clear
AdodcFilmovi.RecordSource = "select distinct Zanr from filmovi"
AdodcFilmovi.Refresh
While Not AdodcFilmovi.Recordset.EOF
If AdodcFilmovi.Recordset.Fields("Zanr") <> "-" Then
Combo3.AddItem (AdodcFilmovi.Recordset.Fields("Zanr"))
End If
AdodcFilmovi.Recordset.MoveNext
Wend
Combo3.Text = "Zanrovi"

' Popunjavanje combo4 - zanr
Combo4.Clear
AdodcFilmovi.RecordSource = "select distinct Godina from filmovi"
AdodcFilmovi.Refresh
While Not AdodcFilmovi.Recordset.EOF
If AdodcFilmovi.Recordset.Fields("Godina") <> "-" Then
Combo4.AddItem (AdodcFilmovi.Recordset.Fields("Godina"))
End If
AdodcFilmovi.Recordset.MoveNext
Wend
Combo4.Text = "Godine"

' Popuna listi
AdodcFilmovi.RecordSource = "select * from filmovi order by naslov"
AdodcFilmovi.Refresh
' Datum
DTPicker1.Value = Now
DTPicker2.Value = Now - 7
DTPicker3.Value = Now
DTPicker4.Value = Now
DTPicker5.Value = Now - 7
End Sub

Private Sub List1_DblClick()
Load Filmovi: Filmovi.SetFocus
End Sub

