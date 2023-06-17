VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form EliCliente 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Eliminar Cliente"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10095
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   10095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnSalir 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Salir"
      Height          =   375
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1680
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   4080
      Top             =   1560
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   "Provider=SQLOLEDB.1;Password=aptoner;Persist Security Info=True;User ID=emmanuel;Initial Catalog=APTONER;Data Source=NEWSERVER"
      OLEDBString     =   "Provider=SQLOLEDB.1;Password=aptoner;Persist Security Info=True;User ID=emmanuel;Initial Catalog=APTONER;Data Source=NEWSERVER"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT ID_CLIENTE, NOMBRE, FECHA_ALTA,VALORACION FROM CLIENTE WHERE VALORACION='0' ORDER BY ID_CLIENTE ASC"
      Caption         =   "Adodc3"
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
   Begin MSDataGridLib.DataGrid DataGrid3 
      Height          =   1695
      Left            =   120
      TabIndex        =   20
      Top             =   2280
      Visible         =   0   'False
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   2990
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
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
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "ID_CLIENTE"
         Caption         =   "Clave"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "NOMBRE"
         Caption         =   "Nombre"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "FECHA_ALTA"
         Caption         =   "FECHA_ALTA"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "VALORACION"
         Caption         =   "Valoración"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   945,071
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   7545,26
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            Object.Visible         =   0   'False
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1049,953
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Height          =   1695
      Left            =   120
      TabIndex        =   19
      Top             =   2280
      Visible         =   0   'False
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   2990
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
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
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "ID_CLIENTE"
         Caption         =   "Clave"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "NOMBRE"
         Caption         =   "Nombre"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "FECHA_ALTA"
         Caption         =   "FECHA_ALTA"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "VALORACION"
         Caption         =   "Valoración"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   884,976
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   7604,788
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            Object.Visible         =   0   'False
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1049,953
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   4080
      Top             =   1200
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   "Provider=SQLOLEDB.1;Password=aptoner;Persist Security Info=True;User ID=emmanuel;Initial Catalog=APTONER;Data Source=NEWSERVER"
      OLEDBString     =   "Provider=SQLOLEDB.1;Password=aptoner;Persist Security Info=True;User ID=emmanuel;Initial Catalog=APTONER;Data Source=NEWSERVER"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT ID_CLIENTE, NOMBRE, FECHA_ALTA,VALORACION FROM CLIENTE WHERE VALORACION='1' ORDER BY ID_CLIENTE ASC"
      Caption         =   "Adodc1"
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
   Begin VB.TextBox Text4 
      DataField       =   "FECHA_ALTA"
      DataSource      =   "Adodc2"
      Height          =   285
      Left            =   3840
      TabIndex        =   2
      Top             =   480
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   5640
      TabIndex        =   16
      Top             =   240
      Width           =   4335
      Begin VB.OptionButton Option5 
         Caption         =   "No Eliminados"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1560
         Width           =   1455
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Eliminados"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1200
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Clave"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Nombre"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   480
         Width           =   975
      End
      Begin VB.OptionButton Option3 
         Caption         =   "General"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Buscar"
         Height          =   375
         Left            =   3000
         TabIndex        =   11
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox Text5 
         DataSource      =   "Adodc2"
         Height          =   285
         Left            =   1200
         TabIndex        =   5
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.TextBox Text1 
      DataField       =   "ID_CLIENTE"
      DataSource      =   "Adodc2"
      Enabled         =   0   'False
      Height          =   285
      Left            =   840
      TabIndex        =   1
      Top             =   480
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      DataField       =   "NOMBRE"
      DataSource      =   "Adodc2"
      Enabled         =   0   'False
      Height          =   285
      Left            =   960
      TabIndex        =   3
      Top             =   960
      Width           =   2655
   End
   Begin VB.TextBox Text3 
      DataField       =   "VALORACION"
      DataSource      =   "Adodc2"
      Enabled         =   0   'False
      Height          =   285
      Left            =   1200
      TabIndex        =   4
      Top             =   1440
      Width           =   975
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Eliminar"
      Height          =   255
      Left            =   2520
      TabIndex        =   0
      Top             =   1440
      Width           =   975
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   4080
      Top             =   840
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   "Provider=SQLOLEDB.1;Password=aptoner;Persist Security Info=True;User ID=emmanuel;Initial Catalog=APTONER;Data Source=NEWSERVER"
      OLEDBString     =   "Provider=SQLOLEDB.1;Password=aptoner;Persist Security Info=True;User ID=emmanuel;Initial Catalog=APTONER;Data Source=NEWSERVER"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT ID_CLIENTE, NOMBRE, FECHA_ALTA,VALORACION FROM CLIENTE ORDER BY ID_CLIENTE ASC"
      Caption         =   "Adodc2"
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1695
      Left            =   120
      TabIndex        =   17
      Top             =   2280
      Visible         =   0   'False
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   2990
      _Version        =   393216
      AllowUpdate     =   -1  'True
      Enabled         =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
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
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "id_cliente"
         Caption         =   "Clave"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "nombre"
         Caption         =   "Nombre"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "valoracion"
         Caption         =   "Valoración"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            Locked          =   -1  'True
            ColumnWidth     =   870,236
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            ColumnWidth     =   7755,024
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            ColumnWidth     =   884,976
         EndProperty
      EndProperty
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Fecha de Alta :"
      Height          =   195
      Left            =   2640
      TabIndex        =   18
      Top             =   480
      Width           =   1200
   End
   Begin VB.Label Label1 
      Caption         =   "Clave :"
      DataSource      =   "Adodc2"
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Nombre :"
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Valoración :"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   1440
      Width           =   855
   End
End
Attribute VB_Name = "EliCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnSalir_Click()
    Unload Me
End Sub
Private Sub Check1_Click()
    Dim buscado, criterio As Integer
    If Check1.Value = 1 Then
        Text3.Text = "1"
    Else
        Text3.Text = "0"
    End If
    Adodc2.Recordset.Update
    buscado = Text1.Text
    If (buscado = "") Then
        Exit Sub
    Else
        buscado = "ID_CLIENTE like '" & buscado & "'"
        Adodc2.Recordset.MoveNext
        If Not Adodc2.Recordset.EOF Then
            Adodc2.Recordset.Find (buscado)
        End If
        If Adodc2.Recordset.EOF Then
            Adodc2.Recordset.MoveFirst
            Adodc2.Recordset.Find (buscado)
            If Adodc2.Recordset.EOF Then
                Adodc2.Recordset.MoveLast
                MsgBox ("No se encontro el registro")
            End If
        End If
    End If
End Sub
Private Sub Command1_Click()
Dim buscado, criterio As Integer
Dim Valoracion As Integer
    If Option1.Value = True Then
        buscado = Text5.Text
        If (buscado = "") Then
            Exit Sub
        Else
            buscado = "ID_CLIENTE like '" & buscado & "'"
            Adodc2.Recordset.MoveNext
            If Not Adodc2.Recordset.EOF Then
                Adodc2.Recordset.Find (buscado)
            End If
            If Adodc2.Recordset.EOF Then
                Adodc2.Recordset.MoveFirst
                Adodc2.Recordset.Find (buscado)
                If Adodc2.Recordset.EOF Then
                    Adodc2.Recordset.MoveLast
                    MsgBox ("No se encontro el registro")
                End If
            End If
        End If
    End If
If Option2.Value = True Then
    buscado = Text5.Text
    If (buscado = "") Then
        Exit Sub
    Else
        buscado = "NOMBRE like '" & buscado & "'"
        Adodc2.Recordset.MoveNext
        If Not Adodc2.Recordset.EOF Then
            Adodc2.Recordset.Find (buscado)
        End If
        If Adodc2.Recordset.EOF Then
            Adodc2.Recordset.MoveFirst
            Adodc2.Recordset.Find (buscado)
            If Adodc2.Recordset.EOF Then
                Adodc2.Recordset.MoveLast
                MsgBox ("No se encontro el registro")
            End If
        End If
    End If
End If
If Option3.Value = True Then
    DataGrid1.Visible = True
End If
If Text3.Text = "1" Then
    Check1.Value = 1
End If
If Text3.Text = "0" Then
    Check1.Value = 0
End If
End Sub
Private Sub DataGrid1_Click()
    Dim buscado, criterio As Integer
    buscado = DataGrid1.Columns(0).Text
        If (buscado = "") Then
             Exit Sub
        Else
            buscado = "ID_CLIENTE like '" & buscado & "'"
            Adodc2.Recordset.MoveNext
            If Not Adodc2.Recordset.EOF Then
                Adodc2.Recordset.Find (buscado)
            End If
            If Adodc2.Recordset.EOF Then
               Adodc2.Recordset.MoveFirst
               Adodc2.Recordset.Find (buscado)
               If Adodc2.Recordset.EOF Then
                  Adodc2.Recordset.MoveLast
                  MsgBox ("No se encontro el registro")
               End If
            End If
        End If
    If Text3.Text = "1" Then
        Check1.Value = 1
    End If
    If Text3.Text = "0" Then
        Check1.Value = 0
    End If
End Sub
Private Sub DataGrid2_Click()
    Dim buscado, criterio As Integer
    buscado = DataGrid2.Columns(0).Text
        If (buscado = "") Then
             Exit Sub
        Else
            buscado = "ID_CLIENTE like '" & buscado & "'"
            Adodc2.Recordset.MoveNext
            If Not Adodc2.Recordset.EOF Then
                Adodc2.Recordset.Find (buscado)
            End If
            If Adodc2.Recordset.EOF Then
               Adodc2.Recordset.MoveFirst
               Adodc2.Recordset.Find (buscado)
               If Adodc2.Recordset.EOF Then
                  Adodc2.Recordset.MoveLast
                  MsgBox ("No se encontro el registro")
               End If
            End If
        End If
    If Text3.Text = "1" Then
        Check1.Value = 1
    End If
    If Text3.Text = "0" Then
        Check1.Value = 0
    End If
End Sub
Private Sub DataGrid3_Click()
    Dim buscado, criterio As Integer
    buscado = DataGrid3.Columns(0).Text
        If (buscado = "") Then
             Exit Sub
        Else
            buscado = "ID_CLIENTE like '" & buscado & "'"
            Adodc2.Recordset.MoveNext
            If Not Adodc2.Recordset.EOF Then
                Adodc2.Recordset.Find (buscado)
            End If
            If Adodc2.Recordset.EOF Then
               Adodc2.Recordset.MoveFirst
               Adodc2.Recordset.Find (buscado)
               If Adodc2.Recordset.EOF Then
                  Adodc2.Recordset.MoveLast
                  MsgBox ("No se encontro el registro")
               End If
            End If
        End If
    If Text3.Text = "1" Then
        Check1.Value = 1
    End If
    If Text3.Text = "0" Then
        Check1.Value = 0
    End If
End Sub

Private Sub Form_Load()
    Option2.Value = True
    Text5 = ""
    If Text3.Text = "1" Then
        Check1.Value = 1
    End If
    If Text3.Text = "0" Then
        Check1.Value = 0
    End If
End Sub
Private Sub Option1_Click()
    DataGrid1.Visible = False
    DataGrid2.Visible = False
    DataGrid3.Visible = False
    Command1.Enabled = True
End Sub
Private Sub Option2_Click()
    DataGrid1.Visible = False
    DataGrid2.Visible = False
    DataGrid3.Visible = False
    Command1.Enabled = True
End Sub
Private Sub Option3_Click()
    DataGrid1.Visible = True
    Command1.Enabled = False
    DataGrid2.Visible = False
    DataGrid3.Visible = False
End Sub
Private Sub Option4_Click()
    DataGrid2.Visible = True
    Command1.Enabled = False
    DataGrid1.Visible = False
    DataGrid3.Visible = False
End Sub
Private Sub Option5_Click()
    DataGrid3.Visible = True
    Command1.Enabled = False
    DataGrid2.Visible = False
    DataGrid1.Visible = False
End Sub
