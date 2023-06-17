VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form VERENTRADA3 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ver Entrada"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11775
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   11775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame17 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   10680
      TabIndex        =   8
      Top             =   3840
      Width           =   975
      Begin VB.Label Label34 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Salir"
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
         TabIndex        =   9
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "VERENTRADA3.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "VERENTRADA3.frx":030A
         Top             =   120
         Width           =   720
      End
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   5760
      TabIndex        =   7
      Top             =   240
      Width           =   1695
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   10680
      TabIndex        =   5
      Top             =   2520
      Width           =   975
      Begin VB.Image Command2 
         Enabled         =   0   'False
         Height          =   735
         Left            =   120
         MouseIcon       =   "VERENTRADA3.frx":23EC
         MousePointer    =   99  'Custom
         Picture         =   "VERENTRADA3.frx":26F6
         Top             =   240
         Width           =   720
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Imprimir"
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
         TabIndex        =   6
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   2520
      TabIndex        =   3
      Top             =   600
      Width           =   4455
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   4440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   7011
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "NOMBRE DEL PROVEEDOR:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "EN ALMACEN 3"
      Height          =   255
      Left            =   7680
      TabIndex        =   2
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "LISTADO DE PRODUCTOS REGISTRADOS EN LA ENTRADA NUMERO  :"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   5535
   End
End
Attribute VB_Name = "VERENTRADA3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Private WithEvents rst As ADODB.Recordset
Attribute rst.VB_VarHelpID = -1
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
On Error GoTo ManejaError
    Dim Total As Double
    Total = 0
    CommonDialog1.Flags = 64
    CommonDialog1.ShowPrinter
    Printer.Print ""
    Printer.Print ""
    Printer.CurrentX = (Printer.Width - Printer.TextWidth(Menu.Text5(0).Text)) / 2
    Printer.Print Menu.Text5(0).Text
    Printer.CurrentX = (Printer.Width - Printer.TextWidth("R.F.C. " & Menu.Text5(8).Text)) / 2
    Printer.Print "R.F.C. " & Menu.Text5(8).Text
    Printer.CurrentX = (Printer.Width - Printer.TextWidth(Menu.Text5(1).Text & " COL. " & Menu.Text5(4).Text)) / 2
    Printer.Print Menu.Text5(1).Text & " COL. " & Menu.Text5(4).Text
    Printer.CurrentX = (Printer.Width - Printer.TextWidth(Menu.Text5(5).Text & ", " & Menu.Text5(6).Text & " C.P. " & Menu.Text5(9).Text)) / 2
    Printer.Print Menu.Text5(5).Text & ", " & Menu.Text5(6).Text & " C.P. " & Menu.Text5(9).Text
    Printer.Print ""
    Printer.Print ""
    Printer.Print "             FECHA : " & Format(Date, "dd/mm/yyyy")
    Printer.Print "             SUCURSAL : " & Menu.Text4(0).Text
    Printer.Print "             IMPRESO POR : " & Menu.Text1(1).Text & " " & Menu.Text1(2).Text
    Printer.Print "             FOLIO: " & Text2.Text
    Printer.Print "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    Printer.CurrentX = (Printer.Width - Printer.TextWidth("COMPROBANTE DE REGISTRO DE ENTRADA DE PRODUCTO")) / 2
    Printer.Print "COMPROBANTE DE REGISTRO DE ENTRADA DE PRODUCTO"
    Printer.Print "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    Printer.Print "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    Printer.CurrentX = (Printer.Width - Printer.TextWidth("NOMBRE DEL PROVEEDOR:  " & Text1.Text)) / 2
    Printer.Print "NOMBRE DEL PROVEEDOR:  " & Text1.Text
    Printer.Print "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    Dim NRegistros As Integer
    NRegistros = ListView1.ListItems.Count
    Dim CON As Integer
    Dim POSY As Integer
    POSY = 3800
    Printer.CurrentY = POSY
    Printer.CurrentX = 100
    Printer.Print "Clave del Producto"
    Printer.CurrentY = POSY
    Printer.CurrentX = 3500
    Printer.Print "Cantidad Registrada"
    Printer.CurrentY = POSY
    Printer.CurrentX = 6500
    Printer.Print "Precio"
    Printer.CurrentY = POSY
    Printer.CurrentX = 7500
    Printer.Print "Sucursal"
    POSY = POSY + 200
    For CON = 1 To NRegistros
        POSY = POSY + 200
        Printer.CurrentY = POSY
        Printer.CurrentX = 100
        Printer.Print ListView1.ListItems(CON).Text
        Printer.CurrentY = POSY
        Printer.CurrentX = 4000
        Printer.Print ListView1.ListItems(CON).SubItems(1)
        Printer.CurrentY = POSY
        Printer.CurrentX = 6500
        Printer.Print ListView1.ListItems(CON).SubItems(3)
        Total = Total + (Val(Replace(ListView1.ListItems(CON).SubItems(3), ",", ".")) * Val(Replace(ListView1.ListItems(CON).SubItems(1), ",", ".")))
        Printer.CurrentY = POSY
        Printer.CurrentX = 7500
        Printer.Print ListView1.ListItems(CON).SubItems(2)
        If POSY >= 14200 Then
            Printer.NewPage
            Printer.Print ""
            Printer.Print ""
            Printer.CurrentX = (Printer.Width - Printer.TextWidth(Menu.Text5(0).Text)) / 2
            Printer.Print Menu.Text5(0).Text
            Printer.CurrentX = (Printer.Width - Printer.TextWidth("R.F.C. " & Menu.Text5(8).Text)) / 2
            Printer.Print "R.F.C. " & Menu.Text5(8).Text
            Printer.CurrentX = (Printer.Width - Printer.TextWidth(Menu.Text5(1).Text & " COL. " & Menu.Text5(4).Text)) / 2
            Printer.Print Menu.Text5(1).Text & " COL. " & Menu.Text5(4).Text
            Printer.CurrentX = (Printer.Width - Printer.TextWidth(Menu.Text5(5).Text & ", " & Menu.Text5(6).Text & " C.P. " & Menu.Text5(9).Text)) / 2
            Printer.Print Menu.Text5(5).Text & ", " & Menu.Text5(6).Text & " C.P. " & Menu.Text5(9).Text
            Printer.Print ""
            Printer.Print ""
            Printer.Print "             FECHA : " & Format(Date, "dd/mm/yyyy")
            Printer.Print "             SUCURSAL : " & Menu.Text4(0).Text
            Printer.Print "             IMPRESO POR : " & Menu.Text1(1).Text & " " & Menu.Text1(2).Text
            Printer.Print "             FOLIO: " & Text2.Text
            Printer.Print "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
            Printer.CurrentX = (Printer.Width - Printer.TextWidth("COMPROBANTE DE REGISTRO DE ENTRADA DE PRODUCTO")) / 2
            Printer.Print "COMPROBANTE DE REGISTRO DE ENTRADA DE PRODUCTO"
            Printer.Print "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
            Printer.Print "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
            Printer.CurrentX = (Printer.Width - Printer.TextWidth("NOMBRE DEL PROVEEDOR:  " & Text1.Text)) / 2
            Printer.Print "NOMBRE DEL PROVEEDOR:  " & Text1.Text
            Printer.Print "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
            POSY = 3800
            Printer.CurrentY = POSY
            Printer.CurrentX = 100
            Printer.Print "Clave del Producto"
            Printer.CurrentY = POSY
            Printer.CurrentX = 3500
            Printer.Print "Cantidad Registrada"
            Printer.CurrentY = POSY
            Printer.CurrentX = 6500
            Printer.Print "Precio"
            Printer.CurrentY = POSY
            Printer.CurrentX = 7500
            Printer.Print "Sucursal"
        End If
    Next CON
    Printer.Print ""
    Printer.Print "             Total = " & Total
    Printer.Print "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    Printer.Print "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    Printer.EndDoc
Exit Sub
ManejaError:
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear
End Sub
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
On Error GoTo ManejaError
    Text2.Text = EntradaProd3.Text5.Text
    Set cnn = New ADODB.Connection
    With cnn
        .ConnectionString = _
            "Provider=SQLOLEDB.1;Password=" & GetSetting("APTONER", "ConfigSACC", "PASSWORD", "LINUX") & ";Persist Security Info=True;User ID=" & GetSetting("APTONER", "ConfigSACC", "USUARIO", "LINUX") & ";Initial Catalog=" & GetSetting("APTONER", "ConfigSACC", "DATABASE", "APTONER") & ";" & _
            "Data Source=" & GetSetting("APTONER", "ConfigSACC", "SERVIDOR", "LINUX") & ";"
        .Open
    End With
    With ListView1
        .View = lvwReport
        .Gridlines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "CLAVE DEL PRODUCTO", 3700
        .ColumnHeaders.Add , , "CANTIDAD", 2700
        .ColumnHeaders.Add , , "SUCURSAL", 3500
        .ColumnHeaders.Add , , "PRECIO", 2500
    End With
    Dim sBuscar As String
    Dim tRs As Recordset
    Dim tLi As ListItem
    sBuscar = "SELECT * FROM ENTRADA_PRODUCTO AS E JOIN SUCURSALES AS S ON E.ID_SUCURSAL = S.ID_SUCURSAL WHERE ID_ENTRADA =" & CDbl(Text2.Text)
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        ListView1.ListItems.Clear
        Do While Not .EOF
            Set tLi = ListView1.ListItems.Add(, , .Fields("ID_PRODUCTO") & "")
                tLi.SubItems(1) = .Fields("CANTIDAD") & ""
                tLi.SubItems(2) = .Fields("NOMBRE") & ""
                tLi.SubItems(3) = .Fields("PRECIO") & ""
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

Private Sub Image9_Click()
    Unload Me
End Sub

Private Sub Text1_Change()
    Command2.Enabled = False
    If Text1.Text <> "" And Text2.Text <> "" Then
        Command2.Enabled = True
    End If
End Sub

Private Sub Text2_Change()
    Command2.Enabled = False
    If Text1.Text <> "" And Text2.Text <> "" Then
        Command2.Enabled = True
    End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        'Busca Entrada
        Dim sBuscar As String
        Dim tRs As Recordset
        Dim tLi As ListItem
        sBuscar = "SELECT * FROM ENTRADA_PRODUCTO AS E JOIN SUCURSALES AS S ON E.ID_SUCURSAL = S.ID_SUCURSAL WHERE ID_ENTRADA =" & CDbl(Text2.Text)
        Set tRs = cnn.Execute(sBuscar)
        With tRs
            ListView1.ListItems.Clear
            Do While Not .EOF
                Set tLi = ListView1.ListItems.Add(, , .Fields("ID_PRODUCTO") & "")
                    tLi.SubItems(1) = .Fields("CANTIDAD") & ""
                    tLi.SubItems(2) = .Fields("NOMBRE") & ""
                    tLi.SubItems(3) = .Fields("PRECIO") & ""
                .MoveNext
            Loop
        End With
    End If
    Dim Valido As String
    Valido = "1234567890"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
