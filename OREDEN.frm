VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Orden 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ORDEN DE COMPRA"
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9885
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   9885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00FFFFFF&
      Height          =   6975
      Left            =   8640
      ScaleHeight     =   6915
      ScaleWidth      =   1155
      TabIndex        =   37
      Top             =   0
      Width           =   1215
      Begin VB.Frame Frame9 
         BackColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   120
         TabIndex        =   38
         Top             =   5520
         Width           =   975
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Cancelar"
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
            Left            =   0
            TabIndex        =   39
            Top             =   960
            Width           =   975
         End
         Begin VB.Image cmdCancelar 
            Height          =   705
            Left            =   120
            MouseIcon       =   "OREDEN.frx":0000
            MousePointer    =   99  'Custom
            Picture         =   "OREDEN.frx":030A
            Top             =   240
            Width           =   705
         End
      End
   End
   Begin VB.PictureBox Picture2 
      Height          =   2055
      Left            =   120
      ScaleHeight     =   1995
      ScaleWidth      =   8235
      TabIndex        =   18
      Top             =   4800
      Width           =   8295
      Begin VB.CommandButton Command4 
         Caption         =   "Calcular"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4560
         Picture         =   "OREDEN.frx":1DBC
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   120
         Width           =   1215
      End
      Begin VB.TextBox Text14 
         Enabled         =   0   'False
         Height          =   285
         Left            =   6840
         TabIndex        =   13
         Top             =   120
         Width           =   1215
      End
      Begin VB.TextBox Text9 
         Enabled         =   0   'False
         Height          =   285
         Left            =   6840
         TabIndex        =   14
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         DataField       =   "ID_ORDEN_COMPRA"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         Height          =   285
         Left            =   2040
         TabIndex        =   7
         Top             =   120
         Width           =   2295
      End
      Begin VB.TextBox Text2 
         DataField       =   "FECHA"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   3600
         TabIndex        =   28
         Top             =   480
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox Text3 
         DataField       =   "ENVIARA"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   2040
         TabIndex        =   9
         Top             =   960
         Width           =   3975
      End
      Begin VB.TextBox Text4 
         DataField       =   "DISCOUNT"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   6840
         TabIndex        =   16
         Text            =   "0"
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox Text5 
         DataField       =   "FREIGHT"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   6840
         TabIndex        =   15
         Text            =   "0"
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox Text6 
         DataField       =   "CONFIRMADA"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         Height          =   195
         Left            =   3360
         TabIndex        =   27
         Text            =   "0"
         Top             =   1440
         Width           =   150
      End
      Begin VB.TextBox Text8 
         DataField       =   "TOTAL"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         Height          =   285
         Left            =   6840
         TabIndex        =   17
         Top             =   1560
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Guardar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4560
         Picture         =   "OREDEN.frx":478E
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Confirmada"
         Height          =   375
         Left            =   2160
         TabIndex        =   10
         Top             =   1320
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   2040
         TabIndex        =   8
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   53411841
         CurrentDate     =   38725
      End
      Begin VB.Label Label12 
         Caption         =   "Subtotal"
         Height          =   255
         Left            =   6120
         TabIndex        =   36
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "IVA"
         Height          =   255
         Left            =   6480
         TabIndex        =   35
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "N° de Orden de Compra"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha"
         Height          =   255
         Left            =   1320
         TabIndex        =   33
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Enviar a:"
         Height          =   255
         Left            =   1200
         TabIndex        =   32
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Discount"
         Height          =   255
         Left            =   6120
         TabIndex        =   31
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Freight"
         Height          =   255
         Left            =   6240
         TabIndex        =   30
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label8 
         Caption         =   "Total"
         Height          =   255
         Left            =   6360
         TabIndex        =   29
         Top             =   1560
         Width           =   375
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   3495
      Left            =   120
      ScaleHeight     =   3435
      ScaleWidth      =   8235
      TabIndex        =   22
      Top             =   1080
      Width           =   8295
      Begin VB.TextBox Text11 
         DataField       =   "PROVEEDOR"
         DataSource      =   "Adodc2"
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   3
         Top             =   120
         Width           =   3735
      End
      Begin VB.TextBox Text12 
         DataField       =   "CLIENTE"
         DataSource      =   "Adodc2"
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   4
         Top             =   480
         Width           =   3735
      End
      Begin VB.TextBox Text13 
         DataField       =   "EJECUTIVO"
         DataSource      =   "Adodc2"
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   5
         Top             =   840
         Width           =   3735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Aceptar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5280
         Picture         =   "OREDEN.frx":7160
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   360
         Width           =   1335
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   2175
         Left            =   120
         TabIndex        =   23
         Top             =   1200
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   3836
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Id_Requisición"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Clave"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Cantidad"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Precio"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Descripción"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "IVA"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label9 
         Caption         =   "Proveedor"
         Height          =   255
         Left            =   480
         TabIndex        =   26
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label10 
         Caption         =   "Cliente"
         Height          =   255
         Left            =   600
         TabIndex        =   25
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label11 
         Caption         =   "Ejecutivo"
         Height          =   255
         Left            =   480
         TabIndex        =   24
         Top             =   840
         Width           =   735
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Extraer Requi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      Picture         =   "OREDEN.frx":9B32
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   480
      Width           =   1455
   End
   Begin VB.TextBox Text10 
      Height          =   285
      Left            =   360
      TabIndex        =   1
      Top             =   480
      Width           =   2775
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "OREDEN.frx":C504
      Height          =   255
      Left            =   7560
      TabIndex        =   20
      Top             =   360
      Visible         =   0   'False
      Width           =   135
      _ExtentX        =   238
      _ExtentY        =   450
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
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
            LCID            =   3082
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
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "OREDEN.frx":C519
      Height          =   255
      Left            =   7800
      TabIndex        =   19
      Top             =   240
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   450
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
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
            LCID            =   3082
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
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   7200
      Top             =   720
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
      CommandType     =   2
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
      Connect         =   "Provider=SQLOLEDB.1;Password=SQLPasSap28;Persist Security Info=True;User ID=aptsys5000;Initial Catalog=APTONER;Data Source=LINUX"
      OLEDBString     =   "Provider=SQLOLEDB.1;Password=SQLPasSap28;Persist Security Info=True;User ID=aptsys5000;Initial Catalog=APTONER;Data Source=LINUX"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "REQUISICION_PRODUCTO"
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   6000
      Top             =   360
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
      CommandType     =   2
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
      Connect         =   "Provider=SQLOLEDB.1;Password=SQLPasSap28;Persist Security Info=True;User ID=aptsys5000;Initial Catalog=APTONER;Data Source=LINUX"
      OLEDBString     =   "Provider=SQLOLEDB.1;Password=SQLPasSap28;Persist Security Info=True;User ID=aptsys5000;Initial Catalog=APTONER;Data Source=LINUX"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "REQUISICION"
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   6000
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
      CommandType     =   2
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
      Connect         =   "Provider=SQLOLEDB.1;Password=SQLPasSap28;Persist Security Info=True;User ID=aptsys5000;Initial Catalog=APTONER;Data Source=LINUX"
      OLEDBString     =   "Provider=SQLOLEDB.1;Password=SQLPasSap28;Persist Security Info=True;User ID=aptsys5000;Initial Catalog=APTONER;Data Source=LINUX"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "ORDEN_COMPRA"
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
   Begin VB.TextBox Text7 
      DataField       =   "ID_REQUISICION"
      DataSource      =   "Adodc1"
      Height          =   195
      Left            =   7320
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Frame Frame1 
      Caption         =   "Numero de Requisicion"
      Height          =   735
      Left            =   240
      TabIndex        =   21
      Top             =   240
      Width           =   5535
   End
End
Attribute VB_Name = "Orden"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
    If Check1.Value = 1 Then
        Text6.Text = "1"
    Else
        Text6.Text = "0"
    End If
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
On Error GoTo ManejaError
    Adodc1.Recordset.AddNew
    Text7.Text = Text10.Text
    Text2.Text = DTPicker1.Value
    Command1.Enabled = False
    Picture2.Left = 120
    Text5.Text = "0"
    Text4.Text = "0"
    Command3.Enabled = False
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub Command2_Click()
On Error GoTo ManejaError
    Dim X As ListItem
    Dim buscado, criterio As Integer
        buscado = Text10.Text
        If (buscado = "") Then
             Exit Sub
        Else
            buscado = "ID_REQUISICION like '" & buscado & "'"
            'Adodc2.Recordset.MoveNext
            If Not Adodc2.Recordset.EOF Then
                Adodc2.Recordset.Find (buscado)
            End If
            If Adodc2.Recordset.EOF And Not Adodc2.Recordset.BOF Then
                Adodc2.Recordset.MoveFirst
                Adodc2.Recordset.Find (buscado)
                If Adodc2.Recordset.EOF Then
                    Adodc2.Recordset.MoveLast
                    MsgBox ("ERROR AL INTENTAR LA BUSQUEDA PUEDE QUE EL REGISTRO NO EXISTA")
                End If
            End If
        End If
    With Adodc3.Recordset
    If Not (.BOF) And Not (.EOF) Then
            .MoveFirst
    End If
        Do While Not (.BOF) And Not (.EOF)
        If (.Fields("ID_REQUISICION")) = Text10.Text Then
             Set X = ListView2.ListItems.Add(, , Text10.Text)
                    X.Tag = Text10.Text
            
                If Not IsNull(.Fields("ID_PRODUCTO")) Then
                     X.SubItems(1) = .Fields("ID_PRODUCTO")
                End If
                If Not IsNull(.Fields("CANTIDAD")) Then
                    X.SubItems(2) = .Fields("CANTIDAD")
                End If
                If Not IsNull(.Fields("precio")) Then
                    X.SubItems(3) = .Fields("precio")
                End If
                If Not IsNull(.Fields("DESCRIPCION")) Then
                    X.SubItems(4) = .Fields("DESCRIPCION")
                End If
                 If Not IsNull(.Fields("iva")) Then
                    X.SubItems(5) = .Fields("iva")
                End If
                  End If
            .MoveNext
        Loop
    End With
    Picture1.Left = 120
    Command2.Enabled = False
    Text10.Enabled = False
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub Command3_Click()
On Error GoTo ManejaError
    If Trim(Text3.Text) = "" Or Trim(Text4.Text) = "" Or Trim(Text5.Text) = "" Or Trim(Text6.Text) = "" Or Trim(Text8.Text) = "" Then
        MsgBox "Faltan datos por ingresar", vbCritical, ""
    Else
        Adodc1.Recordset.Update
        ListView2.ListItems.Clear
        Picture1.Left = -20000
        Picture2.Left = -20000
        Command1.Enabled = True
        Command2.Enabled = True
        Text10.Enabled = True
        Command3.Enabled = False
        Command4.Enabled = True
    End If
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub Command4_Click()
On Error GoTo ManejaError
    Dim X, Y, sum, iva, desc, flete, sum2, iva2, importe, subtotal, CANTIDAD, TOTAL As Double
    If Trim(Text5.Text) = "" Or Trim(Text4.Text) = "" Then
        MsgBox "Las casillas de Freight y Discount´ deben tener algun valor", vbCritical, ""
    Else
        flete = Text5.Text
        desc = Text4.Text
        X = ListView2.ListItems.Count
        sum = 0
        iva = 0
        sum2 = 0
        iva2 = 0
        subtotal = 0
        For Y = 1 To X
            Text3.Text = ListView2.ListItems(Y).SubItems(3)
            Text9.Text = ListView2.ListItems(Y).SubItems(5)
            CANTIDAD = ListView2.ListItems(Y).SubItems(2)
            sum2 = Text3.Text
            iva2 = Text9.Text
            importe = sum2 * CANTIDAD
            subtotal = subtotal + importe
            sum = Text3.Text + sum
            iva = Text9.Text + iva
        Next Y
        Text14.Text = subtotal
        Text9.Text = iva
        TOTAL = sum + flete + desc + iva + subtotal
        Text8.Text = TOTAL
        Command3.Enabled = True
        Command4.Enabled = False
    End If
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
        Err.Clear
End Sub
Private Sub DTPicker1_Change()
    Text2.Text = DTPicker1.Value
End Sub
Private Sub Form_Load()
    DTPicker1.Value = Date
    Picture1.Left = -20000
    Picture2.Left = -20000
End Sub
Private Sub Text1_GotFocus()
    Text1.BackColor = &HFFE1E1
End Sub

Private Sub Text1_LostFocus()
    Text1.BackColor = &H80000005
End Sub

Private Sub Text10_GotFocus()
    Text10.BackColor = &HFFE1E1
End Sub

Private Sub Text10_LostFocus()
    Text10.BackColor = &H80000005
End Sub

Private Sub Text11_GotFocus()
    Text11.BackColor = &HFFE1E1
End Sub

Private Sub Text11_LostFocus()
    Text11.BackColor = &H80000005
End Sub

Private Sub Text12_GotFocus()
    Text12.BackColor = &HFFE1E1
End Sub

Private Sub Text12_LostFocus()
    Text12.BackColor = &H80000005
End Sub

Private Sub Text13_GotFocus()
    Text13.BackColor = &HFFE1E1
End Sub

Private Sub Text13_LostFocus()
    Text13.BackColor = &H80000005
End Sub

Private Sub Text14_GotFocus()
    Text14.BackColor = &HFFE1E1
End Sub

Private Sub Text14_LostFocus()
    Text14.BackColor = &H80000005
End Sub

Private Sub Text3_GotFocus()
    Text3.BackColor = &HFFE1E1
End Sub

Private Sub Text3_LostFocus()
    Text3.BackColor = &H80000005
End Sub

Private Sub Text4_GotFocus()
    Text4.BackColor = &HFFE1E1
End Sub

Private Sub Text4_LostFocus()
    Text4.BackColor = &H80000005
End Sub

Private Sub Text5_GotFocus()
    Text5.BackColor = &HFFE1E1
End Sub

Private Sub Text5_LostFocus()
    Text5.BackColor = &H80000005
End Sub

Private Sub Text7_Change()
    Text7.Text = Text10.Text
End Sub
Private Sub Text8_GotFocus()
    Text8.BackColor = &HFFE1E1
End Sub

Private Sub Text8_LostFocus()
    Text8.BackColor = &H80000005
End Sub

Private Sub Text9_GotFocus()
    Text9.BackColor = &HFFE1E1
End Sub

Private Sub Text9_LostFocus()
    Text9.BackColor = &H80000005
End Sub
