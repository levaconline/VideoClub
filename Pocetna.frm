VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Pocetna 
   BackColor       =   &H00000040&
   Caption         =   " Deal"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Pocetna.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc adoRezervacije 
      Height          =   330
      Left            =   3000
      Top             =   7440
      Visible         =   0   'False
      Width           =   2160
      _ExtentX        =   3810
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
      RecordSource    =   ""
      Caption         =   "Rezervacije"
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2280
      Top             =   7800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ListView LVPi 
      Height          =   615
      Left            =   240
      TabIndex        =   31
      ToolTipText     =   "Double click to close"
      Top             =   7800
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1085
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   49152
      BackColor       =   0
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Title"
         Object.Width           =   5468
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "RentDate"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Amount"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "AmountDays"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "RevertDate"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView LVP 
      Height          =   375
      Left            =   2160
      TabIndex        =   30
      Top             =   7440
      Visible         =   0   'False
      Width           =   135
      _ExtentX        =   238
      _ExtentY        =   661
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   0
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Ime"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Uplac"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Datum"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "IDC"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Undo"
      Enabled         =   0   'False
      Height          =   210
      Left            =   10560
      TabIndex        =   29
      Top             =   6360
      Width           =   735
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Undo"
      Enabled         =   0   'False
      Height          =   210
      Left            =   10560
      TabIndex        =   28
      Top             =   3000
      Width           =   735
   End
   Begin VB.ComboBox Combo2 
      Height          =   330
      ItemData        =   "Pocetna.frx":030A
      Left            =   6120
      List            =   "Pocetna.frx":0329
      TabIndex        =   27
      Text            =   "Title"
      ToolTipText     =   "Odaberite kategoriju za pretrazivanje"
      Top             =   240
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      Height          =   330
      ItemData        =   "Pocetna.frx":037A
      Left            =   1320
      List            =   "Pocetna.frx":0399
      TabIndex        =   26
      Text            =   "Name"
      ToolTipText     =   "Odaberite kategoriju za pretrazivanje"
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   10200
      TabIndex        =   23
      Text            =   "1"
      ToolTipText     =   "Uplata za broj dana"
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Sort"
      Height          =   255
      Left            =   360
      TabIndex        =   21
      ToolTipText     =   "Sort by number"
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   10800
      TabIndex        =   20
      ToolTipText     =   "Koliko je uplaceno"
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Sort"
      Height          =   255
      Left            =   5160
      TabIndex        =   19
      ToolTipText     =   "Sort by number"
      Top             =   720
      Width           =   495
   End
   Begin MSAdodcLib.Adodc adoArhiva 
      Height          =   330
      Left            =   9960
      Top             =   7440
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
      RecordSource    =   "select * from arhiva order by idclana"
      Caption         =   "Arhiva"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc adoIzdato 
      Height          =   330
      Left            =   8040
      Top             =   7440
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
      RecordSource    =   "select * from izdato order by idclana"
      Caption         =   "Izdati"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc adoFilmovi 
      Height          =   330
      Left            =   5280
      Top             =   7440
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      RecordSource    =   "select * from Filmovi order by naslov"
      Caption         =   "Filmovi"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc adoClanovi 
      Height          =   330
      Left            =   240
      Top             =   7440
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
      RecordSource    =   "select * from clanovi order by Prezime"
      Caption         =   "Clanovi"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "Pocetna.frx":03E0
      Height          =   6375
      Left            =   5160
      TabIndex        =   16
      Top             =   720
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   11245
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      BackColor       =   14673638
      ForeColor       =   0
      HeadLines       =   1
      RowHeight       =   19
      TabAction       =   1
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "FILMOVI"
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "ID"
         Caption         =   "BR."
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
      BeginProperty Column02 
         DataField       =   "Broj Kopija"
         Caption         =   "BrKopi"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "Broj Izdatih"
         Caption         =   "Izdato"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0"
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
            ColumnWidth     =   494,929
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2294,929
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   705,26
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   705,26
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Pocetna.frx":03F9
      Height          =   6375
      Left            =   360
      TabIndex        =   15
      ToolTipText     =   "Double Click to details"
      Top             =   720
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   11245
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      HeadLines       =   1
      RowHeight       =   19
      TabAction       =   1
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "CLANOVI"
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "ID"
         Caption         =   "BR"
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
         DataField       =   "Prezime"
         Caption         =   "Prezime"
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
         DataField       =   "Srednje Ime"
         Caption         =   "Sred. Ime"
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
         DataField       =   "Ime"
         Caption         =   "Ime"
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
         MarqueeStyle    =   5
         BeginProperty Column00 
            DividerStyle    =   3
            WrapText        =   -1  'True
            ColumnWidth     =   434,835
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1484,787
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   540,284
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1679,811
         EndProperty
      EndProperty
   End
   Begin VB.ListBox List4 
      BackColor       =   &H00000080&
      ForeColor       =   &H00C0C0FF&
      Height          =   900
      Left            =   10200
      TabIndex        =   12
      ToolTipText     =   "List of movies. Click on item to remove"
      Top             =   4080
      Width           =   1455
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00000080&
      Caption         =   "MOVIES"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Load list"
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00000080&
      Caption         =   "MEMBERS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Load list"
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Search"
      Height          =   255
      Left            =   9120
      TabIndex        =   8
      Top             =   240
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H00000040&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   7200
      TabIndex        =   7
      ToolTipText     =   "Type name"
      Top             =   240
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Search"
      Height          =   255
      Left            =   4200
      TabIndex        =   6
      Top             =   240
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00000040&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      ToolTipText     =   "Type surname"
      Top             =   240
      Width           =   1695
   End
   Begin VB.ListBox List3 
      BackColor       =   &H00000080&
      ForeColor       =   &H00C0C0FF&
      Height          =   1110
      Left            =   10200
      TabIndex        =   4
      ToolTipText     =   "List of movies.  Click to remove item,  Double click to clear list."
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Revert"
      Height          =   375
      Left            =   10200
      TabIndex        =   2
      ToolTipText     =   "Razduzenje"
      Top             =   5880
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Rent"
      Height          =   375
      Left            =   10200
      TabIndex        =   1
      ToolTipText     =   "Zaduzenje"
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   10800
      TabIndex        =   25
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Days"
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   10200
      TabIndex        =   24
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   10200
      TabIndex        =   22
      Top             =   5160
      Width           =   1455
   End
   Begin VB.Label Label7 
      BackColor       =   &H00000080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6615
      Left            =   5040
      TabIndex        =   18
      Top             =   600
      Width           =   4815
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6615
      Left            =   240
      TabIndex        =   17
      Top             =   600
      Width           =   4815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Left            =   10890
      TabIndex        =   11
      ToolTipText     =   "Suma za naplatu"
      Top             =   5400
      Width           =   75
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   10200
      MouseIcon       =   "Pocetna.frx":0412
      MousePointer    =   99  'Custom
      TabIndex        =   14
      ToolTipText     =   "Double click to details"
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00000080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000080FF&
      DrawMode        =   15  'Merge Pen Not
      FillColor       =   &H000040C0&
      FillStyle       =   3  'Vertical Line
      Height          =   495
      Left            =   10200
      Top             =   5160
      Width           =   1455
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   10200
      TabIndex        =   3
      ToolTipText     =   "Member.  Click to clear."
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   10080
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      FillColor       =   &H00000080&
      Height          =   7215
      Left            =   120
      Top             =   120
      Width           =   9855
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   10080
      TabIndex        =   13
      Top             =   3600
      Width           =   1695
   End
End
Attribute VB_Name = "Pocetna"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Rev As String  'Revert/Rent
Public Boroj As Integer 'Broj izdatih filmova u jednom cugu
Public Payment As Currency, Amount As Currency, Acount As Currency
Public Lot  As String 'vreme snimanja - precica za undo
Public IzmeF As String 'za provru izmene imena filma
Option Explicit

Private Sub Command1_Click()
'On Error Resume Next
' Da li je istekla clanarina

Dim Ggin As Integer, Mmes As Integer, Roda As Integer ' Roda = mesec + period clanstva
If Glavni.MemPer > 0 Then
Roda = Month(adoClanovi.Recordset.Fields(9)) + Glavni.MemPer

        Ggin = Year(adoClanovi.Recordset.Fields(9)) + Roda \ 12
        Mmes = Roda Mod 12
        If Year(Now) > Ggin Then GoTo 33
        If Year(Now) >= Ggin And Month(Now) > Mmes Then GoTo 33
        If Year(Now) >= Ggin And Month(Now) >= Mmes And Day(Now) > Day(adoClanovi.Recordset.Fields(9)) Then GoTo 33
 GoTo 23
33:
  If MsgBox("Clanu je isteklo clanstvo. Da li hocete da ga produzite?", vbYesNo + vbDefaultButton2, "Upit") = vbYes Then
     adoClanovi.Recordset.Fields(9) = Format(Now, "dd-mm-yy")
     adoClanovi.Recordset.Update
     MsgBox "Clanstvo je produzeno", vbOKOnly, "Obavestenje"
     Else
     Exit Sub
  End If
23:
End If
'///////////////

Dim Lx As Integer
Dim Dfgtty As String, Hgjhf As Integer, Xero As Integer, Imend As String

If Label4.Caption = "" Then MsgBox "Pronadjite ime clana u listi i klknite na nj.", vbOKOnly, "Greska": Exit Sub
If List3.ListCount = 0 Then MsgBox "Pronadjite filmove u listi i kliknite na nj.", vbOKOnly, "Greska": Exit Sub
If Val(Text3.Text) = 0 Then
If MsgBox("Zelite li da uplatite pretplatu?", vbYesNo, "Nejasnoca") = vbNo Then
Text3.Text = 0
Else
MsgBox "Upisite sumu i pokusajte ponovo", vbOKOnly, "Uputstvo": Exit Sub
End If
End If

Rev = "Rent:"
Load Revers
Revers.Label9.Caption = ""
Revers.Label10.Caption = ""
Revers.Label11.Caption = ""
Revers.Label12.Caption = ""
Lot = Now
For Lx = 0 To List3.ListCount - 1
Imend = Trim$(List3.List(Lx)) '
Revers.List1.AddItem (Imend)
    adoIzdato.Recordset.AddNew
    adoIzdato.Recordset.Fields(1) = adoClanovi.Recordset.Fields(0)
    adoIzdato.Recordset.Fields(2) = Trim$(Imend)
    adoIzdato.Recordset.Fields(3) = Lot
    adoIzdato.Recordset.Fields(4) = Val(Text3.Text) / List3.ListCount
    adoIzdato.Recordset.Update
    '////////////////////////// Dodavanje jednog filma u polje Broj Iadatih
    adoFilmovi.RecordSource = "select * from filmovi where naslov='" & Imend & "'"
    adoFilmovi.Refresh
    Xero = adoFilmovi.Recordset.Fields(9)
    adoFilmovi.Recordset.Fields(9) = Xero + 1
    adoFilmovi.Recordset.Update

Next Lx

Revers.Show vbModal
List3.Clear
Label4.Caption = ""
Text3.Text = ""
Text4.Text = ""
Command9.Enabled = True
Command6_Click
End Sub

Private Sub Command10_Click()
Dim Swq As Integer
If MsgBox("Zelite da ponistite zadnje vracanje filmova?", vbYesNo, "Provera") = vbNo Then Exit Sub
For Swq = 0 To LVP.ListItems.Count - 1
    adoIzdato.Refresh
    adoIzdato.Recordset.AddNew
    adoIzdato.Recordset.Fields(1) = LVP.ListItems(Swq + 1)
    adoIzdato.Recordset.Fields(2) = LVP.ListItems(Swq + 1).ListSubItems(1)
    adoIzdato.Recordset.Fields(3) = LVP.ListItems(Swq + 1).ListSubItems(2)
    adoIzdato.Recordset.Fields(4) = LVP.ListItems(Swq + 1).ListSubItems(3)
    adoIzdato.Recordset.Update
    
    adoFilmovi.Refresh
While Not adoFilmovi.Recordset.EOF
If adoFilmovi.Recordset.Fields(1) = LVP.ListItems(Swq + 1).ListSubItems(1) Then adoFilmovi.Recordset.Fields(9) = Val(adoFilmovi.Recordset.Fields(9)) + 1: GoTo 25
adoFilmovi.Recordset.MoveNext
Wend
25:
adoFilmovi.Recordset.Update
Next Swq
Pocetna.Command6_Click

End Sub

Public Sub Command2_Click()
'REVERT
'_________________
Dim Lx As Integer
Dim Dfgtty As String, Hgjhf As Integer, Imend As String, Xero As Integer
Dim Paimont As Currency, Period As Integer
If List4.ListCount = 0 Then MsgBox "Clan nije zaduzen ni jednim filmom.", vbOKOnly, "Greska": Exit Sub
If MsgBox("Potvrdite razduzenje clana.", vbYesNo, "Razduzenje") = vbYes Then
'[[[[[[
Acount = 0
Period = 0
'[[[[[[
On Error Resume Next
Dim Caz As Integer, Jaz As Date
'Pronalazenje dugovanja clana
Dim Balance As Currency, Dol As Integer, D As Integer

'Dim Mun As MonthConstants, Den As DayConstants, Hades As DayConstants
    Dfgtty = List4.List(0)
  Hgjhf = adoClanovi.Recordset.Fields(0)
']]]]]]

    adoIzdato.RecordSource = "select * from Izdato where  IDClana= " & Hgjhf
    adoIzdato.Refresh
']]]
Rev = "Revert:"
Load Revers
Revers.ListView1.ListItems.Clear
LVP.ListItems.Clear
Dim LV1 As ListItem, LV2 As ListItem, Daw As Date, Saw As Date
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'Petljanje
For Lx = 0 To List4.ListCount - 1
    Imend = List4.List(Lx)
Set LV1 = Revers.ListView1.ListItems.Add(Lx + 1, , Imend)
Set LV2 = LVP.ListItems.Add(Lx + 1, , adoIzdato.Recordset.Fields("IDClana"))
    LV2.SubItems(1) = Imend

'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
adoIzdato.Recordset.MoveFirst
While Not adoIzdato.Recordset.EOF
If adoIzdato.Recordset.Fields("Film") = Imend Then GoTo 19 ' usaglasavanje sa listom
adoIzdato.Recordset.MoveNext
Wend

19:
     Jaz = Format(adoIzdato.Recordset.Fields(3), "dd-mm-yy HH:mm:ss") 'Datum
    Daw = Format(Jaz, "dd/mm/yy")
    LV1.SubItems(2) = Format(Jaz, "dd/mm/yy")
   LV1.SubItems(3) = Format(Now, "dd/mm/yy")
       LV2.SubItems(2) = Format(Jaz, "dd/mm/yy")
   LV2.SubItems(3) = Format(Now, "dd/mm/yy")

   
    Jaz = Format(Now, "dd-mm-yy")
    Caz = DateDiff("d", Daw, Jaz)                'Dani(jaz)
'Broj nedelja
Dol = Weekday(Jaz, vbMonday)
D = (Caz + Dol) \ 7

    If Glavni.InSund = False Then Caz = Caz - D 'Korekcija za slucaj da se nedelja ne obracunava
    If Caz <= 0 Then Caz = 1                    'Nista fraj
    LV1.SubItems(4) = Caz
    Period = Period + Caz
    Acount = Acount + adoIzdato.Recordset.Fields(4)   'Koliko je uplaceno
    LV2.SubItems(3) = adoIzdato.Recordset.Fields(4)
    LV1.SubItems(1) = adoIzdato.Recordset.Fields(4) & Chr$(163)
    
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

' Oduzimanje jednog filma iz polja Broj Izdatih
    adoFilmovi.RecordSource = "select * from filmovi where  naslov = '" & Imend & "'"
    adoFilmovi.Refresh
    Xero = adoFilmovi.Recordset.Fields(9)
    adoFilmovi.Recordset.Fields(9) = Xero - 1
    adoFilmovi.Recordset.Update
    
'///////////////////////// Upis u Arhivu
    adoArhiva.Refresh
    adoArhiva.Recordset.AddNew
    adoArhiva.Recordset.Fields(1) = Trim(Imend)                     'Naslov
    adoArhiva.Recordset.Fields(2) = Period                          'Period koliko je film zadrzan
    adoArhiva.Recordset.Fields(3) = Format(Now, "dd/mm/yy")         'Datum vracanja filma
    adoArhiva.Recordset.Fields(4) = adoClanovi.Recordset.Fields(0)  'IDClana
    adoArhiva.Recordset.Update

' Razduzuvanje
    adoIzdato.Recordset.Delete
    adoIzdato.Recordset.MoveNext
    adoIzdato.RecordSource = "select * from izdato"
    adoIzdato.Recordset.Update

Next Lx


'///////////////////////////////////////////////////
Paimont = Glavni.RV * Period
Balance = Paimont - Acount

Label5.Caption = ""

'////////////////////////////
'Upis u revers i labele
Revers.Label10.Caption = " Ukupan Br. Dana: " & Period  ' Broj dana
Revers.Label11.Caption = " X " & Glavni.RV        'Cena damna
Revers.Label12.Caption = " Payment: " & Paimont
Label1.Caption = Chr$(163) & " " & Balance
If Balance > 0 Then
Revers.Label9 = " Balance: " & Chr$(163) & " " & Balance
Label8.Caption = " Balance:"
Else
Revers.Label9 = " Kusur : " & Chr$(163) & " " & Balance
Label8.Caption = " Kusur:"
End If
Revers.Label8 = " Amount: " & Acount
Command10.Enabled = True
'/////////////////////////////
Revers.Show vbModal
List4.Clear
End If
Pocetna.Command6_Click
End Sub

Private Sub Command3_Click()
On Error Resume Next
Dim Pop As String, Rty As String, C As Integer
'Pronaci kategoriju iz Text1 u bazi podataka i rezultate prikazati u List1
Pop = Trim$(Combo1.Text)

If Pop = "Name" Then Pop = "Ime"
If Pop = "MidName" Then Pop = "Srednje Ime"
If Pop = "Surname" Then Pop = "Prezime"
If Pop = "Address" Then Pop = "Adresa"
If Pop = "City" Then Pop = "Grad"
If Pop = "State" Then Pop = "Drzava"
If Pop = "Phono" Then Pop = "Telefon"
If Pop = "Document" Then Pop = "Dokument"

Rty = Trim$(Text1.Text) '& " * ")
'C = Len(Rty)
'adoClanovi.RecordSource = "select * from clanovi where left$([" & Pop & "] ," & C & ") like '" & Rty & "'"        ' like 's*' " ''" & Rty & "'"
If Pop = "ID" Then
adoClanovi.RecordSource = "select * from clanovi where [" & Pop & "]  >=   " & Rty & " order by [" & Pop & "]"          ' like 's*' " ''" & Rty & "'"
Else
adoClanovi.RecordSource = "select * from clanovi where [" & Pop & "]  >=   '" & Rty & "'order by [" & Pop & "]"          ' like 's*' " ''" & Rty & "'"
End If
adoClanovi.Refresh
End Sub

Private Sub Command4_Click()
'Pronaci kategoriju filma iz Text2 u bazi podataka i rezultat predstaviti u List2
Dim Pop As String, Rty As String, C As Integer
Pop = Trim$(Combo2.Text)
If Pop = "Title" Then Pop = "Naslov"
If Pop = "Directed by" Then Pop = "Reziser"
If Pop = "Actors" Then Pop = "GlUloga"
If Pop = "Zanr" Then Pop = "Zanr"
If Pop = "Year" Then Pop = "Godina"
If Pop = "Time" Then Pop = "Trajanje"
If Pop = "Broj Kopija" Then Pop = "Broj Kopija"
If Pop = "Broj Izdatih" Then Pop = "Broj Izdatih"
Rty = Trim$(Text2.Text)
'C = Len(Rty)

'adoFilmovi.RecordSource = "select * from filmovi where left$([" & Pop & "] ," & C & ") like '" & Rty & "'"
If Pop = "ID" Or Pop = "Godina" Or Pop = "Trajanje" Or Pop = "Broj Kopija" Or Pop = "Broj Izdatih" Then
adoFilmovi.RecordSource = "select * from filmovi where [" & Pop & "] >= " & Rty & " order by [" & Pop & "]"
Else
adoFilmovi.RecordSource = "select * from filmovi where [" & Pop & "] >= '" & Rty & "' order by [" & Pop & "]"
End If
adoFilmovi.Refresh
End Sub

Public Sub Command5_Click()
' Popuna List1 sa Prezimenima, Srednjim Imenoima i Imenima, iz baze podataka
adoClanovi.RecordSource = "select * from clanovi order by prezime"
adoClanovi.Refresh
End Sub

Public Sub Command6_Click()
' Popuna List2 sa Nazivima Folmova, iz baze podataka
adoFilmovi.RecordSource = "select * from filmovi order by naslov"
adoFilmovi.Refresh

End Sub

Private Sub Command7_Click()
adoFilmovi.RecordSource = "select * from filmovi order by id"
adoFilmovi.Refresh
End Sub

Private Sub Command8_Click()
adoClanovi.RecordSource = "select * from clanovi order by id"
adoClanovi.Refresh
End Sub

Private Sub Command9_Click()
If MsgBox("Zelite da ponistite zadnje izdavanje, izvrseno  " & Lot & "?", vbYesNo, "Provera") = vbNo Then Exit Sub
Dim Koss As String
adoFilmovi.RecordSource = "select * from filmovi order by naslov"
adoFilmovi.Refresh
Koss = Right$(Lot, 8)
adoIzdato.RecordSource = "select * from izdato where right$(DatumRentir,8)= '" & Koss & "'"
adoIzdato.Refresh

While Not adoIzdato.Recordset.EOF
        adoFilmovi.Recordset.MoveFirst
    While Not adoFilmovi.Recordset.EOF
        If adoFilmovi.Recordset.Fields("Naslov") = adoIzdato.Recordset.Fields("Film") Then
            adoFilmovi.Recordset.Fields("Broj Izdatih") = adoFilmovi.Recordset.Fields("Broj Izdatih") - 1
            adoFilmovi.Recordset.Update
            GoTo 33
        End If
    adoFilmovi.Recordset.MoveNext
    Wend
33:
adoIzdato.Recordset.Delete
adoIzdato.Recordset.Update
adoIzdato.Recordset.MoveNext
Wend
Pocetna.Command6_Click
MsgBox "Zadnje iznajmljivanje je ponisteno.", vbOKOnly, "Obavestenje"
End Sub

Private Sub DataGrid1_DblClick()
Glavni.Smer1 = 1
Unload Clanstvo: Load Clanstvo: Clanstvo.SetFocus
End Sub

Private Sub DataGrid1_SelChange(Cancel As Integer)

Dim Grrooe As Integer
Label1.Caption = ""
Label8.Caption = ""
On Error Resume Next
Label4.Caption = adoClanovi.Recordset.Fields(1)
Grrooe = adoClanovi.Recordset.Fields(0)
If Grrooe = 0 Then Exit Sub
Label5 = Label4
List4.Clear
adoIzdato.RecordSource = "Select * from izdato where idclana=" & Grrooe
adoIzdato.Refresh

While Not adoIzdato.Recordset.EOF
List4.AddItem Trim$(adoIzdato.Recordset.Fields(2))
adoIzdato.Recordset.MoveNext
Wend
Text4.Text = 1

End Sub

Private Sub DataGrid2_AfterColUpdate(ByVal ColIndex As Integer)
If IzmeF <> adoFilmovi.Recordset.Fields("Naslov") Then
    If MsgBox("Promenjeno je ime filma! Zelite li da ime bude zamenjeno novim u tabelama u arhivi. (u koliko je film vec rentiran pritisnite Yes)", vbYesNo, "Upozorenje!") = vbYes Then
    Izmena.IzmF adoFilmovi.Recordset.Fields("Naslov"), IzmeF
    End If
End If
End Sub

Private Sub DataGrid2_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
IzmeF = adoFilmovi.Recordset.Fields("Naslov")
End Sub

Private Sub DataGrid2_DblClick()
Unload Filmovi: Set Filmovi = Nothing: Load Filmovi: Filmovi.SetFocus
End Sub

Private Sub DataGrid2_SelChange(Cancel As Integer)
On Error Resume Next
If adoFilmovi.Recordset.Fields(9) >= adoFilmovi.Recordset.Fields(8) Then
If MsgBox("Nema vise raspolozivih kopija, zelite li da rezervisete film.", vbYesNo, "Upit") = vbYes Then


Exit Sub
End If
End If
List3.AddItem Trim$(adoFilmovi.Recordset.Fields(1))
Text3.Text = Val(Text4.Text) * Glavni.RV * List3.ListCount

End Sub

Private Sub Form_Activate()

adoClanovi.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\Data\VideoClub.mdb"
adoIzdato.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\Data\VideoClub.mdb"
adoArhiva.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\Data\VideoClub.mdb"
adoFilmovi.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\Data\VideoClub.mdb"
'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Programs\WinVideo\Data\VideoClub.mdb;Persist Security Info=False
End Sub

Public Sub Label2_DblClick()
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
LVPi.ListItems.Clear
With LVPi
.Left = 1
.Top = 1
.Width = 11800
.Height = 7450
.Visible = True
End With
Dim Jop As ListItem, Bid As Integer
Bid = adoClanovi.Recordset.Fields("ID")
adoIzdato.Refresh
adoIzdato.RecordSource = "select * from izdato where idclana=" & Bid
While Not adoIzdato.Recordset.EOF
Set Jop = LVPi.ListItems.Add(, , adoIzdato.Recordset.Fields("Film"))
Jop.SubItems(1) = adoIzdato.Recordset.Fields("DatumRentir")
Jop.SubItems(2) = adoIzdato.Recordset.Fields("Uplaceno")
Jop.SubItems(3) = CInt(adoIzdato.Recordset.Fields("Uplaceno") / Glavni.RV)
Jop.SubItems(4) = Format(DateAdd("d", Jop.SubItems(3), Jop.SubItems(1)), "dd-mm-yy")
adoIzdato.Recordset.MoveNext
Wend
adoIzdato.RecordSource = "select * from izdato"

End Sub

Private Sub Label4_Click()
Label4.Caption = ""
End Sub

Private Sub List1_DblClick()
Glavni.Smer1 = 1
Load Clanstvo: Clanstvo.SetFocus

End Sub

Private Sub List2_DblClick()
Load Filmovi: Filmovi.SetFocus

End Sub

Private Sub Label5_DblClick()
Label2_DblClick
End Sub

Private Sub List3_Click()
List3.RemoveItem (List3.ListIndex)
Text3.Text = Val(Text4.Text) * Glavni.RV * List3.ListCount

End Sub

Private Sub List3_DblClick()
List3.Clear
Text3.Text = Val(Text4.Text) * Glavni.RV * List3.ListCount

End Sub

Private Sub List4_Click()
List4.RemoveItem (List4.ListIndex)

End Sub

Private Sub LVPi_DblClick()
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
With LVPi
    .Visible = False
    .Left = 240
    .Top = 7800
    .Width = 1935
    .Height = 615
    
End With

End Sub

Private Sub Text4_Change()
Text3.Text = Val(Text4.Text) * Glavni.RV * List3.ListCount
End Sub
