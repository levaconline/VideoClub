VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Direct 
   Caption         =   " Data"
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
   Icon            =   "Opste.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command4 
      Caption         =   "Arhiva"
      Height          =   255
      Left            =   4800
      TabIndex        =   4
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Izdato"
      Height          =   255
      Left            =   3240
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Filmovi"
      Height          =   255
      Left            =   1680
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Clanovi"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   6495
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   11456
      _Version        =   393216
      AllowArrows     =   -1  'True
      HeadLines       =   1
      RowHeight       =   19
      TabAction       =   1
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
            ColumnWidth     =   374,74
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1230,236
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc adoArhiva 
      Height          =   330
      Left            =   8280
      Top             =   7080
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
      RecordSource    =   "select * from arhiva order by id"
      Caption         =   "adoArhiva"
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
      Left            =   5520
      Top             =   7080
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
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
      Caption         =   "adoIzdato"
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
      Left            =   2760
      Top             =   7080
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
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
      RecordSource    =   "select * from filmovi order by naslov"
      Caption         =   "adoFilmovi"
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
      Left            =   120
      Top             =   7080
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
      RecordSource    =   "select *  from clanovi  order by prezime"
      Caption         =   "adoClanovi"
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
End
Attribute VB_Name = "Direct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Frbr As String 'kontrola promene imena filma
Public Bibi As Boolean
Option Explicit

Public Sub Command1_Click()
DataGrid1.Caption = "CLANOVI"
Set DataGrid1.DataSource = adoClanovi
DataGrid1.Columns(0).Width = 222
DataGrid1.Columns(1).Width = 222
adoClanovi.Refresh
Bibi = False
End Sub

Private Sub Command2_Click()
DataGrid1.Caption = "FILMOVI"
Set DataGrid1.DataSource = adoFilmovi
adoFilmovi.Refresh
Bibi = True
End Sub

Private Sub Command3_Click()
DataGrid1.Caption = "IZDATO"
Set DataGrid1.DataSource = adoIzdato
adoIzdato.Refresh
Bibi = False
End Sub

Private Sub Command4_Click()
DataGrid1.Caption = "ARHIVA"
Set DataGrid1.DataSource = adoArhiva
adoArhiva.Refresh
Bibi = False
End Sub

Private Sub DataGrid1_AfterColUpdate(ByVal ColIndex As Integer)
If Bibi = True Then
    If Frbr <> adoFilmovi.Recordset.Fields("Naslov") Then
            If MsgBox("Promenjeno je ime filma! Zelite li da ime bude zamenjeno novim u tabelama u arhivi. (u koliko je film vec rentiran pritisnite Yes)", vbYesNo, "Upozorenje!") = vbYes Then
            Izmena.IzmF adoFilmovi.Recordset.Fields("Naslov"), Frbr
            End If
    End If
End If
End Sub

Private Sub DataGrid1_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
If Bibi = True Then
    Frbr = adoFilmovi.Recordset.Fields("Naslov")
End If
End Sub

Private Sub Form_Activate()
adoClanovi.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\Data\VideoClub.mdb"
adoIzdato.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\Data\VideoClub.mdb"
adoArhiva.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\Data\VideoClub.mdb"
adoFilmovi.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\Data\VideoClub.mdb"

DataGrid1.AllowAddNew = True
DataGrid1.AllowDelete = True
DataGrid1.AllowUpdate = True
Command1_Click
End Sub

