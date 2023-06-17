VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form EliProveedor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Eliminar Proveedor"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10230
   ControlBox      =   0   'False
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   10230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid DataGrid3 
      Height          =   1695
      Left            =   120
      TabIndex        =   18
      Top             =   2280
      Visible         =   0   'False
      Width           =   9975
      _ExtentX        =   17595
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
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "ID_PROVEEDOR"
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
         DataField       =   "ELIMINADO"
         Caption         =   "Eliminado"
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
            ColumnWidth     =   1289,764
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   7304,882
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   900,284
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Height          =   1695
      Left            =   120
      TabIndex        =   17
      Top             =   2280
      Visible         =   0   'False
      Width           =   9975
      _ExtentX        =   17595
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
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "ID_PROVEEDOR"
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
         DataField       =   "ELIMINADO"
         Caption         =   "Eliminado"
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
            ColumnWidth     =   1289,764
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   7334,93
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   900,284
         EndProperty
      EndProperty
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Eliminar"
      Height          =   255
      Left            =   3600
      TabIndex        =   12
      Top             =   1440
      Width           =   855
   End
   Begin VB.TextBox Text3 
      DataField       =   "ELIMINADO"
      DataSource      =   "Adodc2"
      Enabled         =   0   'False
      Height          =   285
      Left            =   1080
      TabIndex        =   2
      Top             =   1440
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      DataField       =   "NOMBRE"
      DataSource      =   "Adodc2"
      Enabled         =   0   'False
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   960
      Width           =   4695
   End
   Begin VB.TextBox Text1 
      DataField       =   "ID_PROVEEDOR"
      DataSource      =   "Adodc2"
      Enabled         =   0   'False
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Top             =   480
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      Height          =   2055
      Left            =   5880
      TabIndex        =   10
      Top             =   120
      Width           =   4215
      Begin VB.OptionButton Option5 
         Caption         =   "No Eliminados"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1680
         Width           =   1575
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Eliminados"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox Text5 
         DataSource      =   "Adodc2"
         Height          =   285
         Left            =   1080
         TabIndex        =   3
         Top             =   240
         Width           =   3015
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Buscar"
         Height          =   375
         Left            =   2880
         TabIndex        =   9
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton Option3 
         Caption         =   "General"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   960
         Width           =   1095
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Nombre"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Clave"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.CommandButton btnSalir 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Salir"
      Height          =   375
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1680
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   4080
      Top             =   240
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
      RecordSource    =   "SELECT ID_PROVEEDOR, NOMBRE, ELIMINADO FROM PROVEEDOR ORDER BY ID_PROVEEDOR ASC"
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
      TabIndex        =   13
      Top             =   2280
      Visible         =   0   'False
      Width           =   9975
      _ExtentX        =   17595
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
         DataField       =   "ID_PROVEEDOR"
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
         DataField       =   "NOMBRE"
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
         DataField       =   "ELIMINADO"
         Caption         =   "Eliminado"
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
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   360
      Top             =   8040
      Width           =   2040
      _ExtentX        =   3598
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
      Connect         =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=APTONER;Data Source=AP-703D1061112D"
      OLEDBString     =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=APTONER;Data Source=AP-703D1061112D"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT ID_PROVEEDOR, NOMBRE, ELIMINADO FROM PROVEEDOR WHERE ELIMINADO='0' ORDER BY ID_PROVEEDOR ASC"
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   4080
      Top             =   600
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
      RecordSource    =   "SELECT ID_PROVEEDOR, NOMBRE, ELIMINADO FROM PROVEEDOR WHERE ELIMINADO='1' ORDER BY ID_PROVEEDOR ASC"
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
   Begin VB.Label Label3 
      Caption         =   "Valoración :"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Nombre :"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Clave :"
      DataSource      =   "Adodc2"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   480
      Width           =   615
   End
End
Attribute VB_Name = "EliProveedor"
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
        buscado = "ID_PROVEEDOR like '" & buscado & "'"
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
    If Option1.Value = True Then
       buscado = Text5.Text
       If (buscado = "") Then
        Exit Sub
       Else
         buscado = "ID_PROVEDOOR like '" & buscado & "'"
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
        buscado = "ID_PROVEEDOR like '" & buscado & "'"
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
        buscado = "ID_PROVEEDOR like '" & buscado & "'"
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
        buscado = "ID_PROVEEDOR like '" & buscado & "'"
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
    DataGrid1.Visible = False
    Command1.Enabled = False
    DataGrid2.Visible = True
    DataGrid3.Visible = False
End Sub
Private Sub Option5_Click()
    DataGrid1.Visible = False
    Command1.Enabled = False
    DataGrid2.Visible = False
    DataGrid3.Visible = True
End Sub
