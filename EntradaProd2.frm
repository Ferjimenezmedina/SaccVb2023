VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form EntradaProd2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Entrada de Productos ALMACEN 2"
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9975
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   9975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Ver Entrada"
      Height          =   375
      Left            =   2760
      TabIndex        =   16
      Top             =   6960
      Width           =   1335
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   120
      Top             =   6960
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
      Connect         =   "Provider=SQLOLEDB.1;Password=SaccPass*20;Persist Security Info=True;User ID=AdmSACC2210;Initial Catalog=APTONER;Data Source=LINUX"
      OLEDBString     =   "Provider=SQLOLEDB.1;Password=SaccPass*20;Persist Security Info=True;User ID=AdmSACC2210;Initial Catalog=APTONER;Data Source=LINUX"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "EXISTENCIAS"
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
   Begin VB.TextBox Text6 
      Height          =   285
      Index           =   2
      Left            =   1560
      TabIndex        =   47
      Text            =   "Text6"
      Top             =   7080
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Index           =   1
      Left            =   1440
      TabIndex        =   46
      Text            =   "Text6"
      Top             =   7080
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Index           =   0
      Left            =   1320
      TabIndex        =   45
      Text            =   "Text6"
      Top             =   7080
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   7
      Left            =   1800
      MaxLength       =   50
      TabIndex        =   15
      Top             =   6360
      Width           =   7575
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   4080
      TabIndex        =   14
      Top             =   5880
      Width           =   5295
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   7560
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   9
      Top             =   3120
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1560
      MaxLength       =   20
      TabIndex        =   8
      Top             =   3120
      Width           =   4215
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   0
      Left            =   5880
      TabIndex        =   36
      Top             =   3120
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Frame Frame1 
      Caption         =   "Entrada"
      Height          =   1815
      Left            =   120
      TabIndex        =   25
      Top             =   5040
      Width           =   9615
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   1
         Left            =   1080
         MaxLength       =   8
         TabIndex        =   10
         Top             =   360
         Width           =   2175
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   2
         Left            =   4080
         MaxLength       =   8
         TabIndex        =   11
         Top             =   360
         Width           =   2175
      End
      Begin VB.TextBox Text3 
         Height          =   255
         Index           =   3
         Left            =   9240
         MaxLength       =   10
         TabIndex        =   29
         Top             =   360
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox Text3 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   3
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   3000
         MaxLength       =   10
         TabIndex        =   28
         Top             =   840
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox Text3 
         Height          =   195
         Index           =   5
         Left            =   9360
         MaxLength       =   50
         TabIndex        =   27
         Top             =   960
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "EntradaProd2.frx":0000
         Left            =   7320
         List            =   "EntradaProd2.frx":000A
         TabIndex        =   12
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox Text4 
         Height          =   225
         Left            =   9345
         TabIndex        =   26
         Top             =   690
         Visible         =   0   'False
         Width           =   150
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1080
         TabIndex        =   13
         Top             =   840
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         Format          =   16777217
         CurrentDate     =   38663
      End
      Begin VB.Label Label2 
         Caption         =   "Cantidad"
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Precio"
         Height          =   255
         Left            =   3480
         TabIndex        =   34
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Moneda"
         Height          =   255
         Left            =   6480
         TabIndex        =   33
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Fecha"
         Height          =   255
         Left            =   240
         TabIndex        =   32
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "Sucursal"
         Height          =   255
         Left            =   3240
         TabIndex        =   31
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "Codigo de Barras"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   1320
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Guardar/Nuevo"
      Height          =   375
      Left            =   4320
      TabIndex        =   17
      Top             =   6960
      Width           =   1335
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Salir"
      Height          =   375
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6960
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   6
      Left            =   9720
      TabIndex        =   23
      Top             =   2520
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox AyuTxt 
      Height          =   285
      Left            =   120
      TabIndex        =   22
      Top             =   2520
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox MuesTxt 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   285
      Left            =   7920
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   240
      Width           =   1815
   End
   Begin VB.TextBox BusTxt 
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Top             =   240
      Width           =   6375
   End
   Begin VB.CommandButton Guardar 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   5160
      TabIndex        =   7
      Top             =   2520
      Width           =   1335
   End
   Begin VB.TextBox MaTxt 
      Enabled         =   0   'False
      Height          =   285
      Index           =   5
      Left            =   240
      TabIndex        =   21
      Top             =   2520
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox MaTxt 
      Height          =   285
      Index           =   4
      Left            =   2160
      MaxLength       =   50
      TabIndex        =   2
      Top             =   2040
      Width           =   2055
   End
   Begin VB.TextBox MaTxt 
      Height          =   285
      Index           =   3
      Left            =   5160
      TabIndex        =   3
      Top             =   2040
      Width           =   2055
   End
   Begin VB.TextBox MaTxt 
      Height          =   195
      Index           =   2
      Left            =   9720
      TabIndex        =   20
      Top             =   2160
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox MaTxt 
      Height          =   285
      Index           =   1
      Left            =   9720
      TabIndex        =   19
      Top             =   240
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox MaTxt 
      Enabled         =   0   'False
      Height          =   195
      Index           =   0
      Left            =   4440
      TabIndex        =   5
      Top             =   2520
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   2520
      Width           =   2295
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1455
      Left            =   120
      TabIndex        =   24
      Top             =   3480
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   2566
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   1560
      Top             =   6960
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
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
      Connect         =   "Provider=SQLOLEDB.1;Password=SaccPass*20;Persist Security Info=True;User ID=AdmSACC2210;Initial Catalog=APTONER;Data Source=LINUX"
      OLEDBString     =   "Provider=SQLOLEDB.1;Password=SaccPass*20;Persist Security Info=True;User ID=AdmSACC2210;Initial Catalog=APTONER;Data Source=LINUX"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "ENTRADA_PRODUCTO"
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
      Left            =   7920
      Top             =   6960
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
      Connect         =   "Provider=SQLOLEDB.1;Password=SaccPass*20;Persist Security Info=True;User ID=AdmSACC2210;Initial Catalog=APTONER;Data Source=LINUX"
      OLEDBString     =   "Provider=SQLOLEDB.1;Password=SaccPass*20;Persist Security Info=True;User ID=AdmSACC2210;Initial Catalog=APTONER;Data Source=LINUX"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SUCURSALES"
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
   Begin MSAdodcLib.Adodc AdoEnt 
      Height          =   330
      Left            =   6960
      Top             =   2520
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
      Connect         =   "Provider=SQLOLEDB.1;Password=SaccPass*20;Persist Security Info=True;User ID=AdmSACC2210;Initial Catalog=APTONER;Data Source=LINUX"
      OLEDBString     =   "Provider=SQLOLEDB.1;Password=SaccPass*20;Persist Security Info=True;User ID=AdmSACC2210;Initial Catalog=APTONER;Data Source=LINUX"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "ENTRADAS"
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
   Begin MSAdodcLib.Adodc AdoUsu 
      Height          =   330
      Left            =   8160
      Top             =   2520
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
      Connect         =   "Provider=SQLOLEDB.1;Password=SaccPass*20;Persist Security Info=True;User ID=AdmSACC2210;Initial Catalog=APTONER;Data Source=LINUX"
      OLEDBString     =   "Provider=SQLOLEDB.1;Password=SaccPass*20;Persist Security Info=True;User ID=AdmSACC2210;Initial Catalog=APTONER;Data Source=LINUX"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "USUARIOS"
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
   Begin MSComCtl2.DTPicker DTFecha 
      Height          =   375
      Left            =   8040
      TabIndex        =   4
      Top             =   2040
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   16777217
      CurrentDate     =   38667
   End
   Begin MSComctlLib.ListView ListProv 
      Height          =   1335
      Left            =   120
      TabIndex        =   37
      Top             =   600
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   2355
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label1 
      Caption         =   "Buscar producto :"
      Height          =   255
      Left            =   120
      TabIndex        =   44
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   "Clave del producto"
      Height          =   255
      Left            =   6120
      TabIndex        =   43
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label num 
      Caption         =   "Numero de Factura"
      Height          =   255
      Left            =   600
      TabIndex        =   42
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label tot 
      Caption         =   "Total"
      Height          =   255
      Left            =   4440
      TabIndex        =   41
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label fech 
      Caption         =   "Fecha"
      Height          =   255
      Left            =   7320
      TabIndex        =   40
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label prov 
      Caption         =   "Proveedor"
      Height          =   255
      Left            =   240
      TabIndex        =   39
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label ent 
      Caption         =   "Numero de Entrada"
      Height          =   255
      Left            =   480
      TabIndex        =   38
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   9840
      Y1              =   3000
      Y2              =   3000
   End
End
Attribute VB_Name = "EntradaProd2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Private cnn2 As ADODB.Connection
Private WithEvents rst As ADODB.Recordset
Attribute rst.VB_VarHelpID = -1
Private WithEvents rst2 As ADODB.Recordset
Attribute rst2.VB_VarHelpID = -1
Private Sub Combo1_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
     KeyAscii = 0
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear
End Sub
Private Sub Combo2_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
     KeyAscii = 0
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear
End Sub
Private Sub Combo2_LostFocus()
On Error GoTo ManejaError
    Text3(3).Text = Combo2.Text
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear
End Sub
Private Sub cmdAdd_Click()
On Error GoTo ManejaError
    Text3(5).Text = Combo1.Text
'*******************************Agrega a Existencias en Sucursal*************************
    If Text1(0).Text <> "" And Text3(5).Text <> "" And Text3(1).Text <> "" Then
        Dim PRODU As String
        Dim SUC As String
        Dim tLi As ListItem
        Dim Cant As Double
        Dim TPRODU As String 'T de TRAJO
        Dim TSUC As String
        Dim TCant As Double
        SUC = Text3(5).Text
        PRODU = Text1(0).Text
        Cant = CDbl(Text3(1).Text)
        deAPTONER.Qu2Existencia SUC, PRODU
        If Not (deAPTONER.rsQu2Existencia.EOF) Or Not (deAPTONER.rsQu2Existencia.BOF) Then
            TCant = deAPTONER.rsQu2Existencia!CANTIDAD
            Dim OPERA As Double
            OPERA = TCant + Cant
            deAPTONER.ModyExistencia SUC, PRODU, OPERA
            OPERA = 0
        Else
            Dim sBuscar As String
            Dim tRs2 As Recordset
            Dim num As String
            sBuscar = "INSERT INTO EXISTENCIAS (ID_PRODUCTO, CANTIDAD, SUCURSAL) VALUES ('" & PRODU & "', " & Cant & ", '" & SUC & "' );"
            cnn.Execute (sBuscar)
        End If
        deAPTONER.rsQu2Existencia.Close
    Else
        MsgBox "Falta incormacion necesaria para el traspaso o el inventario no existe"
    End If
'***************************************************************************************
    MaTxt(0).Text = Text1(0).Text
    Text3(4).Text = DTPicker1.Value
    Text3(6).Text = Text5.Text
    Text3(3).Text = Combo2.Text
    Text3(5).Text = Text4.Text
    Text3(0).Text = Text1(0).Text
    If Text3(0).Text <> "" And Text3(1).Text <> "" And Text3(2).Text <> "" And Text3(3).Text <> "" And Text3(4).Text <> "" And Text3(5).Text <> "" And Text3(6).Text <> "" Then
        Adodc3.Recordset.Update
        Adodc3.Recordset.AddNew
    Else
        MsgBox ("Falta informacion necesaria del registro")
    End If
Exit Sub
ManejaError:
    If Err.Number = -2147467259 Then
        If MsgBox("SE PERDIO LA CONEXIÓN CON LOS SERVIDORES, ¿DESEA RESTABLECERLA?", vbYesNo + vbCritical + vbDefaultButton1, "SACC") = vbYes Then Reconexion
        Err.Clear
    ElseIf Err.Number = 3704 Then
        If MsgBox("SE PERDIO LA CONEXIÓN CON LOS SERVIDORES, ¿DESEA RESTABLECERLA?", vbYesNo + vbCritical + vbDefaultButton1, "SACC") = vbYes Then Reconexion
        Err.Clear
    Else
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear
    End If
End Sub
Private Sub Command1_Click()
On Error GoTo ManejaError
    If Text5.Text = "" Then
        MsgBox "Debe generar una entrada para abrir esta opcion"
    Else
        VERENTRADA2.Show vbModal
    End If
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear
End Sub
Private Sub DTPicker1_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    KeyAscii = 0
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear
End Sub
Private Sub DTPicker1_LostFocus()
On Error GoTo ManejaError
    Text3(4).Text = DTPicker1.Value
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear
End Sub
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
On Error GoTo ManejaError
    AyuTxt.Text = Menu.Text1(0).Text
    Me.DTFecha.Value = Format(Date, "dd/mm/yyyy")
    Me.DTFecha.Enabled = False
    Me.DTPicker1.Value = Format(Date, "dd/mm/yyyy")
    Text2.Enabled = False
    Text1(0).Enabled = False
    ListView1.Enabled = False
    Text3(1).Enabled = False
    Text3(2).Enabled = False
    Combo2.Enabled = False
    DTPicker1.Enabled = False
    Combo1.Enabled = False
    Text3(7).Enabled = False
    cmdAdd.Enabled = False
    Text3(4).Text = DTPicker1.Value
    Set cnn = New ADODB.Connection
    Set rst = New ADODB.Recordset
    With cnn
        .ConnectionString = _
            "Provider=SQLOLEDB.1;Password=" & GetSetting("APTONER", "ConfigSACC", "PASSWORD", "LINUX") & ";Persist Security Info=True;User ID=" & GetSetting("APTONER", "ConfigSACC", "USUARIO", "LINUX") & ";Initial Catalog=" & GetSetting("APTONER", "ConfigSACC", "DATABASE", "APTONER") & ";" & _
            "Data Source=" & GetSetting("APTONER", "ConfigSACC", "SERVIDOR", "LINUX") & ";"
        .Open
    End With
    rst.Open "SELECT * FROM ALMACEN2", cnn, adOpenDynamic, adLockOptimistic
    With ListView1
        .View = lvwReport
        .Gridlines = True
        .LabelEdit = lvwManual
        .CheckBoxes = True
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "CLAVE DEL PRODUCTO", 2500
        .ColumnHeaders.Add , , "DESCRIPCION", 5400
    End With
       
    Set Text2.DataSource = Adodc2

    
    Set Text3(0).DataSource = Adodc3
    Set Text3(1).DataSource = Adodc3
    Set Text3(2).DataSource = Adodc3
    Set Text3(3).DataSource = Adodc3
    Set Text3(4).DataSource = Adodc3
    Set Text3(5).DataSource = Adodc3
    Set Text3(6).DataSource = Adodc3
    Set Text3(7).DataSource = Adodc3
    
    Text3(0).DataField = "ID_PRODUCTO"
    Text3(1).DataField = "CANTIDAD"
    Text3(2).DataField = "PRECIO"
    Text3(3).DataField = "MONEDA"
    Text3(4).DataField = "FECHA"
    Text3(5).DataField = "ID_SUCURSAL"
    Text3(6).DataField = "ID_ENTRADA"
    Text3(7).DataField = "CODIGO_BARAS"
    Set cnn2 = New ADODB.Connection
    Set rst2 = New ADODB.Recordset
    With cnn2
        .ConnectionString = _
            "Provider=SQLOLEDB.1;Password=" & GetSetting("APTONER", "ConfigSACC", "PASSWORD", "LINUX") & ";Persist Security Info=True;User ID=" & GetSetting("APTONER", "ConfigSACC", "USUARIO", "LINUX") & ";Initial Catalog=" & GetSetting("APTONER", "ConfigSACC", "DATABASE", "APTONER") & ";" & _
            "Data Source=" & GetSetting("APTONER", "ConfigSACC", "SERVIDOR", "LINUX") & ";"
        .Open
    End With
    rst2.Open "SELECT * FROM PROVEEDOR", cnn2, adOpenDynamic, adLockOptimistic
    With ListProv
        .View = lvwReport
        .Gridlines = True
        .LabelEdit = lvwManual
        .CheckBoxes = True
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Add , , "CLAVE DEL PROVEEDOR", 2400
        .ColumnHeaders.Add , , "NOMBRE", 6100
        .ColumnHeaders.Add , , "CIUDAD", 2300
    End With
    Dim i As Integer
    For i = 1 To 5
        Set MaTxt(i).DataSource = AdoEnt
    Next
    MaTxt(1).DataField = "ID_PROVEEDOR"
    MaTxt(2).DataField = "FECHA"
    MaTxt(3).DataField = "TOTAL"
    MaTxt(4).DataField = "FACTURA"
    MaTxt(5).DataField = "ID_USUARIO"
    AdoEnt.Recordset.AddNew
    MaTxt(5).Text = AyuTxt.Text
Exit Sub
ManejaError:
    If Err.Number = -2147467259 Then
        If MsgBox("SE PERDIO LA CONEXIÓN CON LOS SERVIDORES, ¿DESEA RESTABLECERLA?", vbYesNo + vbCritical + vbDefaultButton1, "SACC") = vbYes Then Reconexion
        Err.Clear
    ElseIf Err.Number = 3704 Then
        If MsgBox("SE PERDIO LA CONEXIÓN CON LOS SERVIDORES, ¿DESEA RESTABLECERLA?", vbYesNo + vbCritical + vbDefaultButton1, "SACC") = vbYes Then Reconexion
        Err.Clear
    Else
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear
    End If
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo ManejaError
    Text1(0).Text = Item
    Text3(0).Text = Item
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear
End Sub
Private Sub btnSalir_Click()
On Error GoTo ManejaError
    Unload Me
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear
End Sub
Private Sub Combo1_DropDown()
On Error GoTo ManejaError
    Combo1.Clear
    Buscarcbo
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear
End Sub
Private Sub Combo1_LostFocus()
On Error GoTo ManejaError
    Dim BQuery As String
    BQuery = "NOMBRE Like '" & Combo1.Text & "'"
    Adodc2.Recordset.MoveFirst
    Adodc2.Recordset.Find BQuery
    
    Set Text4.DataSource = Adodc2
    Text4.DataField = "ID_SUCURSAL"
    Text3(5).Text = Text4.Text
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear
End Sub
Private Sub Buscarcbo(Optional ByVal Siguiente As Boolean = False)
On Error GoTo ManejaError
    Dim nRegcbo As Long
    Dim vBookmarkcbo As Variant
    Dim sADOBuscarcbo As String
    On Error Resume Next
    sADOBuscarcbo = "NOMBRE LIKE '" & "%" & "'"
    vBookmarkcbo = Adodc2.Recordset.Bookmark
    If Siguiente = False Then
        Adodc2.Recordset.MoveFirst
        Adodc2.Recordset.Find sADOBuscarcbo
    Else
        Adodc2.Recordset.Find sADOBuscarcbo, 1
    End If
    Dim rs As Recordset
    Set rs = Adodc2.Recordset
    If Adodc2.Recordset.EOF = False Then
    Do While Adodc2.Recordset.EOF = False
        Combo1.AddItem Adodc2.Recordset.Fields("NOMBRE")
        Adodc2.Recordset.MoveNext
    Loop
    End If
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear
End Sub
Private Sub cmdSalir_Click()
On Error GoTo ManejaError
    Unload Me
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear
End Sub

Private Sub MaTxt_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo ManejaError
    Dim Valido As String
    Valido = "1234567890"
    If Index = 3 Or Index = 4 Then
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii > 26 Then
            If InStr(Valido, Chr(KeyAscii)) = 0 Then
                KeyAscii = 0
            End If
        End If
    End If
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 13 Then
        BusProd
    End If
    Dim Valido As String
    Valido = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZ-1234567890. "
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear
End Sub
Private Sub BusProd()
On Error GoTo ManejaError
    Dim tRs As Recordset
    Dim tLi As ListItem
    Dim bus As String
    Dim sBus As String
    sBus = "SELECT * FROM ALMACEN2 WHERE ID_PRODUCTO LIKE '%" & Text2.Text & "%'"
    Set tRs = cnn.Execute(sBus)
    With tRs
        ListView1.ListItems.Clear
        Do While Not .EOF
            Set tLi = ListView1.ListItems.Add(, , .Fields("ID_PRODUCTO") & "")
            tLi.SubItems(1) = .Fields("DESCRIPCION") & ""
            .MoveNext
        Loop
    End With
Exit Sub
ManejaError:
    If Err.Number = -2147467259 Then
        If MsgBox("SE PERDIO LA CONEXIÓN CON LOS SERVIDORES, ¿DESEA RESTABLECERLA?", vbYesNo + vbCritical + vbDefaultButton1, "SACC") = vbYes Then Reconexion
        Err.Clear
    ElseIf Err.Number = 3704 Then
        If MsgBox("SE PERDIO LA CONEXIÓN CON LOS SERVIDORES, ¿DESEA RESTABLECERLA?", vbYesNo + vbCritical + vbDefaultButton1, "SACC") = vbYes Then Reconexion
        Err.Clear
    Else
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear
    End If
End Sub
Private Sub Text3_Change(Index As Integer)
On Error GoTo ManejaError
    If Index = 7 Then
        Text3(7).Text = Replace(Text3(7).Text, ",", "")
        Text3(7).Text = Replace(Text3(7).Text, "-", "")
        Text3(7).Text = Replace(Text3(7).Text, "_", "")
        Text3(7).Text = Replace(Text3(7).Text, ".", "")
        Text3(7).Text = Replace(Text3(7).Text, "*", "")
        Text3(7).Text = Replace(Text3(7).Text, "%", "")
        Text3(7).Text = Replace(Text3(7).Text, "&", "")
        Text3(7).Text = Replace(Text3(7).Text, "/", "")
        Text3(7).Text = Replace(Text3(7).Text, "'", "")
        Text3(7).Text = Replace(Text3(7).Text, "$", "")
        Text3(7).Text = Replace(Text3(7).Text, "=", "")
        Text3(7).Text = Replace(Text3(7).Text, "@", "")
        Text3(7).Text = Replace(Text3(7).Text, "!", "")
        Text3(7).Text = Replace(Text3(7).Text, "?", "")
        Text3(7).Text = Replace(Text3(7).Text, "^", "")
        Text3(7).Text = Replace(Text3(7).Text, "#", "")
        Text3(7).Text = Replace(Text3(7).Text, " ", "")
        Text3(7).Text = Replace(Text3(7).Text, "+", "")
        Text3(7).Text = Replace(Text3(7).Text, ";", "")
        Text3(7).Text = Replace(Text3(7).Text, ":", "")
    End If
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear
End Sub
Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo ManejaError
    Dim Valido As String
    If Index = 1 Or Index = 2 Then
        Valido = "1234567890."
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii > 26 Then
            If InStr(Valido, Chr(KeyAscii)) = 0 Then
                KeyAscii = 0
            End If
        End If
    Else
        Valido = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZ-1234567890. "
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii > 26 Then
            If InStr(Valido, Chr(KeyAscii)) = 0 Then
                KeyAscii = 0
            End If
        End If
    End If
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear
End Sub
Private Sub Guardar_Click()
On Error GoTo ManejaError
    MaTxt(2).Text = DTFecha.Value
    If MaTxt(1).Text <> "" And MaTxt(2).Text <> "" And MaTxt(3).Text <> "" And MaTxt(4).Text <> "" And MaTxt(5).Text <> "" Then
        AdoEnt.Recordset.Update
        Dim nRegcbo As Long
        Dim vBookma As Variant
        Dim sADOBus As String
        Dim rst As Recordset
        On Error Resume Next
        Adodc3.Recordset.AddNew
        sADOBus = "ORDER BY ID_ENTRADA"
        vBookma = AdoEnt.Recordset.Bookmark
        AdoEnt.Recordset.MoveLast
        AdoEnt.Recordset.Find sADOBus
        Set rst = AdoEnt.Recordset
        Set Text5.DataSource = AdoEnt
        Text5.DataField = "ID_ENTRADA"
        Text2.Enabled = True
        ListView1.Enabled = True
        Text3(1).Enabled = True
        Text3(2).Enabled = True
        Combo2.Enabled = True
        Combo1.Enabled = True
        Text3(7).Enabled = True
        cmdAdd.Enabled = True
        cmdSalir.Enabled = True
        BusTxt.Enabled = False
        ListProv.Enabled = False
        MaTxt(3).Enabled = False
        MaTxt(4).Enabled = False
        Guardar.Enabled = False
    Else
        MsgBox ("Falta Informacion necesaria")
    End If
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear
End Sub
Private Sub DTFecha_CloseUp()
On Error GoTo ManejaError
    MaTxt(2).Text = DTFecha.Value
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear
End Sub
Private Sub ListProv_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo ManejaError
    MaTxt(1).Text = Item
    MuesTxt.Text = Item
    BusTxt.Text = Item.SubItems(1)
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear
End Sub
Private Sub BusTxt_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 13 Then
        Dim sBuscar As String
        Dim tRs As Recordset
        Dim tLi As ListItem
        sBuscar = BusTxt
        sBuscar = Replace(sBuscar, "*", "%")
        sBuscar = Replace(sBuscar, "?", "_")
    
        BusTxt = sBuscar
        sBuscar = "SELECT * FROM PROVEEDOR WHERE NOMBRE LIKE '%" & sBuscar & "%' ORDER BY NOMBRE"
        Set tRs = cnn.Execute(sBuscar)
        With tRs
                ListProv.ListItems.Clear
                Do While Not .EOF
                    Set tLi = ListProv.ListItems.Add(, , .Fields("ID_PROVEEDOR") & "")
                    tLi.SubItems(1) = .Fields("NOMBRE") & ""
                    tLi.SubItems(2) = .Fields("CIUDAD") & ""
                    .MoveNext
                Loop
        End With
    End If
    Dim Valido As String
    Valido = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZ-1234567890. "
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
Exit Sub
ManejaError:
    If Err.Number = -2147467259 Then
        If MsgBox("SE PERDIO LA CONEXIÓN CON LOS SERVIDORES, ¿DESEA RESTABLECERLA?", vbYesNo + vbCritical + vbDefaultButton1, "SACC") = vbYes Then Reconexion
        Err.Clear
    ElseIf Err.Number = 3704 Then
        If MsgBox("SE PERDIO LA CONEXIÓN CON LOS SERVIDORES, ¿DESEA RESTABLECERLA?", vbYesNo + vbCritical + vbDefaultButton1, "SACC") = vbYes Then Reconexion
        Err.Clear
    Else
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear
    End If
End Sub
Private Sub Text5_Change()
On Error GoTo ManejaError
    Text3(6).Text = Text5.Text
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear
End Sub

Sub Reconexion()
On Error GoTo ManejaError
    Set cnn = New ADODB.Connection
    With cnn
        .ConnectionString = _
            "Provider=SQLOLEDB.1;Password=" & GetSetting("APTONER", "ConfigSACC", "PASSWORD", "LINUX") & ";Persist Security Info=True;User ID=" & GetSetting("APTONER", "ConfigSACC", "USUARIO", "LINUX") & ";Initial Catalog=" & GetSetting("APTONER", "ConfigSACC", "DATABASE", "APTONER") & ";" & _
            "Data Source=" & GetSetting("APTONER", "ConfigSACC", "SERVIDOR", "LINUX") & ";"
        .Open
        MsgBox "LA CONEXION SE RESABLECIO CON EXITO. PUEDE CONTINUAR CON SU TRABAJO.", vbInformation, "SACC"
    End With
Exit Sub
ManejaError:
    MsgBox Err.Number & Err.Description
    Err.Clear
    If MsgBox("NO PUDIMOS RESTABLECER LA CONEXIÓN, ¿DESEA REINTENTARLO?", vbYesNo + vbCritical + vbDefaultButton1, "SACC") = vbYes Then Reconexion
End Sub


