VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form FrmRevDomi 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Imprimir Reporte de Domicilios (recoleccion)"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9975
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   9975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      Height          =   1575
      Left            =   5520
      ScaleHeight     =   1515
      ScaleWidth      =   1155
      TabIndex        =   15
      Top             =   0
      Width           =   1215
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   120
         TabIndex        =   18
         Top             =   120
         Width           =   975
         Begin VB.Image Image1 
            Height          =   870
            Left            =   120
            MouseIcon       =   "FrmRevDomi.frx":0000
            MousePointer    =   99  'Custom
            Picture         =   "FrmRevDomi.frx":030A
            Top             =   120
            Width           =   720
         End
         Begin VB.Label Label6 
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
            TabIndex        =   19
            Top             =   960
            Width           =   975
         End
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   5895
      Left            =   8760
      ScaleHeight     =   5835
      ScaleWidth      =   1155
      TabIndex        =   14
      Top             =   0
      Width           =   1215
      Begin VB.Frame Frame11 
         BackColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   120
         TabIndex        =   20
         Top             =   3360
         Width           =   975
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "EXCEL"
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
            TabIndex        =   21
            Top             =   960
            Width           =   975
         End
         Begin VB.Image Image10 
            Height          =   720
            Left            =   120
            MouseIcon       =   "FrmRevDomi.frx":23EC
            MousePointer    =   99  'Custom
            Picture         =   "FrmRevDomi.frx":26F6
            Top             =   240
            Width           =   720
         End
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   120
         TabIndex        =   16
         Top             =   4560
         Width           =   975
         Begin VB.Label Label27 
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
            TabIndex        =   17
            Top             =   960
            Width           =   975
         End
         Begin VB.Image Image9 
            Height          =   870
            Left            =   120
            MouseIcon       =   "FrmRevDomi.frx":4238
            MousePointer    =   99  'Custom
            Picture         =   "FrmRevDomi.frx":4542
            Top             =   120
            Width           =   720
         End
      End
   End
   Begin VB.TextBox Text2 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   6840
      MaxLength       =   50
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1200
      MaxLength       =   50
      TabIndex        =   6
      Top             =   5400
      Width           =   4695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Terminados"
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
      Left            =   6000
      Picture         =   "FrmRevDomi.frx":6624
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5400
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   3960
      TabIndex        =   3
      Top             =   240
      Width           =   1455
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3735
      Left            =   120
      TabIndex        =   10
      Top             =   1560
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   6588
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Ver"
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
      Left            =   2280
      Picture         =   "FrmRevDomi.frx":8FF6
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   960
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   720
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      PrinterDefault  =   0   'False
   End
   Begin VB.CommandButton Command2 
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
      Height          =   375
      Left            =   7440
      Picture         =   "FrmRevDomi.frx":B9C8
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5400
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   2110
      TabIndex        =   2
      Top             =   240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   50659329
      CurrentDate     =   38833
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   430
      TabIndex        =   1
      Top             =   240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   50659329
      CurrentDate     =   38833
   End
   Begin VB.Label Label5 
      Caption         =   "Codigo del empleado :"
      Height          =   255
      Left            =   6840
      TabIndex        =   13
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Comentarios :"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   5400
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Zona :"
      Height          =   255
      Left            =   3480
      TabIndex        =   11
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Al :"
      Height          =   255
      Left            =   1780
      TabIndex        =   9
      Top             =   240
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Del :"
      Height          =   255
      Left            =   100
      TabIndex        =   0
      Top             =   240
      Width           =   375
   End
End
Attribute VB_Name = "FrmRevDomi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Dim ZoNa As String
Dim FechaDe As String
Dim FechaAl As String
Dim Id_Repartidor As String
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Sub Combo1_DropDown()
    Combo1.Clear
    Combo1.AddItem "NORTE"
    Combo1.AddItem "SUR"
    Combo1.AddItem "ESTE"
    Combo1.AddItem "OESTE"
    Combo1.AddItem "CENTRO"
    Combo1.AddItem "NORESTE"
    Combo1.AddItem "NOROESTE"
    Combo1.AddItem "SURESTE"
    Combo1.AddItem "SUROESTE"
End Sub
Private Sub Combo1_GotFocus()
    Combo1.BackColor = &HFFE1E1
End Sub
Private Sub Combo1_LostFocus()
    Combo1.BackColor = &H80000005
End Sub
Private Sub Command2_Click()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    CommonDialog1.Flags = 64
    CommonDialog1.CancelError = True
    CommonDialog1.ShowPrinter
    Printer.Print " "
    Printer.Print " "
    Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(0).Text)) / 2
    Printer.Print VarMen.Text5(0).Text
    Printer.CurrentX = (Printer.Width - Printer.TextWidth("R.F.C. " & VarMen.Text5(8).Text)) / 2
    Printer.Print "R.F.C. " & VarMen.Text5(8).Text
    Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text)) / 2
    Printer.Print VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text
    Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text)) / 2
    Printer.Print VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text
    Printer.Print " "
    Printer.Print " "
    Printer.Print "             FECHA : " & Format(Date, "dd/mm/yyyy")
    Printer.Print "             SUCURSAL : " & VarMen.Text4(0).Text
    Printer.Print "             IMPRESO POR : " & VarMen.Text1(1).Text & " " & VarMen.Text1(2).Text
    Printer.Print "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    Printer.Print "                                                                          COMPROBANTE DE REGISTRO DE ENTRADA DE PRODUCTO"
    Printer.Print "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    Dim Con As Integer
    Dim POSY As Integer
    Dim fecha As String
    fecha = DTPicker2.Value + 1
    Dim Conta As Integer
    Conta = 16
    POSY = 3800
    Printer.CurrentY = POSY
    Printer.CurrentX = 100
    Printer.Print "Numero"
    Printer.CurrentY = POSY
    Printer.CurrentX = 2500
    Printer.Print "Cliente"
    Printer.CurrentY = POSY
    Printer.CurrentX = 8500
    Printer.Print "Recibio"
    POSY = POSY + 200
    sBuscar = "SELECT * FROM DOMICILIOS ORDER BY ID_DOMICILIO"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            If tRs.Fields("FECHA") <> fecha Then
                POSY = POSY + 200
                Printer.CurrentY = POSY
                Printer.CurrentX = 100
                Printer.Print "Pasar entre las " & tRs.Fields("DE_HORA") & " y las "; tRs.Fields("A_HORA")
                Printer.CurrentY = POSY
                Printer.CurrentX = 4000
                Printer.Print Mid(tRs.Fields("NOM_CLIENTE"), 1, 25)
                Printer.CurrentY = POSY
                Printer.CurrentX = 6500
                Printer.Print "Tel. " & tRs.Fields("TELEFONO")
                POSY = POSY + 200
                Printer.CurrentY = POSY
                Printer.CurrentX = 100
                Printer.Print tRs.Fields("DOMICILIO") & " " & tRs.Fields("COLONIA")
                Printer.CurrentY = POSY
                Printer.CurrentX = 6500
                Printer.Print "_____________________________"
                Conta = Conta + 2
                If Conta >= 69 Then
                    Conta = 16
                    Printer.NewPage
                    Printer.Print " "
                    Printer.Print " "
                    Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(0).Text)) / 2
                    Printer.Print VarMen.Text5(0).Text
                    Printer.CurrentX = (Printer.Width - Printer.TextWidth("R.F.C. " & VarMen.Text5(8).Text)) / 2
                    Printer.Print "R.F.C. " & VarMen.Text5(8).Text
                    Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text)) / 2
                    Printer.Print VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text
                    Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text)) / 2
                    Printer.Print VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text
                    Printer.Print " "
                    Printer.Print " "
                    Printer.Print "             FECHA : " & Format(Date, "dd/mm/yyyy")
                    Printer.Print "             SUCURSAL : " & VarMen.Text4(0).Text
                    Printer.Print "             IMPRESO POR : " & VarMen.Text1(1).Text & " " & VarMen.Text1(2).Text
                    Printer.Print "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
                    Printer.Print "                                                                          COMPROBANTE DE REGISTRO DE ENTRADA DE PRODUCTO"
                    Printer.Print "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
                End If
                tRs.MoveNext
            End If
        Loop
    End If
    Printer.Print ""
    Printer.Print ""
    Printer.Print "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    Printer.Print "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    Me.Command2.Enabled = False
    Me.Command4.Enabled = False
    Printer.EndDoc
    CommonDialog1.Copies = 1
Exit Sub
ManejaError:
    If Err.Number <> 32755 Then
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    End If
    Err.Clear
End Sub
Private Sub Command3_Click()
    FechaDe = DTPicker1.Value
    FechaAl = DTPicker2.Value
    If Combo1.Text <> "" Then
        ZoNa = Combo1.Text
    Else
        ZoNa = "%"
    End If
    BuscarDomi
End Sub
Private Sub Command4_Click()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim NumeroRegistros As Integer
    Dim Conta As Integer
    NumeroRegistros = ListView1.ListItems.Count
    For Conta = 1 To NumeroRegistros
        If Me.ListView1.ListItems.Item(Conta).Checked = True Then
            sBuscar = "UPDATE DOMICILIOS SET ESTADO = 'T', ID_REPA = " & Id_Repartidor & ", FECHA_FIN = '" & Format(Date, "dd/mm/yyyy") & "', COMENTARIOS = '" & Text1.Text & "' WHERE ID_DOMICILIO = " & Me.ListView1.ListItems.Item(Conta)
            Set tRs = cnn.Execute(sBuscar)
        End If
    Next Conta
    Me.Command2.Enabled = False
    Me.Command4.Enabled = False
    BuscarDomi
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
On Error GoTo ManejaError
    FrmRevDomi.Height = 1980
    FrmRevDomi.Width = 6795
    DTPicker1.Value = Format(Date - 1, "dd/mm/yyyy")
    DTPicker2.Value = Format(Date, "dd/mm/yyyy")
    Me.Command2.Enabled = False
    Me.Command4.Enabled = False
    Set cnn = New ADODB.Connection
    With cnn
        .ConnectionString = _
            "Provider=" & NvoMen.TxtProvider.Text & ";Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & NvoMen.TxtBaseDatos.Text & ";Data Source=" & NvoMen.txtServidor.Text & ";"
        .Open
    End With
    With ListView1
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "No. Pedido", 1000
        .ColumnHeaders.Add , , "Cliente", 4500
        .ColumnHeaders.Add , , "Domicilio", 1500
        .ColumnHeaders.Add , , "Colonia", 1500
        .ColumnHeaders.Add , , "Telefono", 1500
        .ColumnHeaders.Add , , "Fecha", 1500
        .ColumnHeaders.Add , , "Horario", 1500
        .ColumnHeaders.Add , , "Nota", 5500
        .ColumnHeaders.Add , , "# de Articulos", 1500
        .ColumnHeaders.Add , , "Fecha/Hora Alta", 2000
        .ColumnHeaders.Add , , "Capturó", 2000
    End With
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub BuscarDomi()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim ItMx As ListItem
    Me.ListView1.ListItems.Clear
    sBuscar = "SELECT * FROM DOMICILIOS WHERE ESTADO = 'P' AND (ZONA LIKE '" & ZoNa & "' OR ZONA = 'DESC') AND FECHA BETWEEN '" & FechaDe & "' AND '" & FechaAl & "' ORDER BY ID_DOMICILIO"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            If tRs.Fields("FECHA") <> Format(Date, "dd/mm/yyyy") Then
                Set ItMx = Me.ListView1.ListItems.Add(, , tRs.Fields("ID_DOMICILIO"))
                If Not IsNull(tRs.Fields("NOM_CLIENTE")) Then ItMx.SubItems(1) = tRs.Fields("NOM_CLIENTE")
                If Not IsNull(tRs.Fields("DOMICILIO")) Then ItMx.SubItems(2) = tRs.Fields("DOMICILIO")
                If Not IsNull(tRs.Fields("COLONIA")) Then ItMx.SubItems(3) = tRs.Fields("COLONIA")
                If Not IsNull(tRs.Fields("TELEFONO")) Then ItMx.SubItems(4) = tRs.Fields("TELEFONO")
                If Not IsNull(tRs.Fields("FECHA")) Then ItMx.SubItems(5) = tRs.Fields("FECHA")
                If Not IsNull(tRs.Fields("DE_HORA")) And Not IsNull(tRs.Fields("A_HORA")) Then ItMx.SubItems(6) = "DE LAS " & tRs.Fields("DE_HORA") & " A LAS " & tRs.Fields("A_HORA")
                If Not IsNull(tRs.Fields("NOTA")) Then ItMx.SubItems(7) = tRs.Fields("NOTA")
                If Not IsNull(tRs.Fields("NUM_ARTICULOS")) Then ItMx.SubItems(8) = tRs.Fields("NUM_ARTICULOS")
                If Not IsNull(tRs.Fields("FECHA_ALTA")) Then ItMx.SubItems(9) = tRs.Fields("FECHA_ALTA")
                If Not IsNull(tRs.Fields("USUARIO")) Then ItMx.SubItems(10) = tRs.Fields("USUARIO")
            End If
            tRs.MoveNext
        Loop
    End If
    If ListView1.ListItems.Count <> 0 Then
        FrmRevDomi.Height = 6375
        FrmRevDomi.Width = 10035
        Picture2.Visible = False
    Else
        MsgBox "NO EXISTEN DOMICILIOS PARA LA ZONA " & ZoNa & " EN ESE RANGO DE FECHAS", vbInformation, "SACC"
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Form_Resize()
    Me.Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
End Sub
Private Sub Image1_Click()
    Unload Me
End Sub
Private Sub Image10_Click()
    Dim StrCopi As String
    Dim Con As Integer
    Dim Con2 As Integer
    Dim NumColum As Integer
    Dim Ruta As String
    Me.CommonDialog1.FileName = ""
    CommonDialog1.DialogTitle = "Guardar como"
    CommonDialog1.Filter = "Excel (*.xls) |*.xls|"
    Me.CommonDialog1.ShowSave
    Ruta = Me.CommonDialog1.FileName
    If ListView1.ListItems.Count > 0 Then
        If Ruta <> "" Then
            NumColum = ListView1.ColumnHeaders.Count
            For Con = 1 To ListView1.ColumnHeaders.Count
                StrCopi = StrCopi & ListView1.ColumnHeaders(Con).Text & Chr(9)
            Next
            StrCopi = StrCopi & Chr(13)
            For Con = 1 To ListView1.ListItems.Count
                StrCopi = StrCopi & ListView1.ListItems.Item(Con) & Chr(9)
                For Con2 = 1 To NumColum - 1
                    StrCopi = StrCopi & ListView1.ListItems.Item(Con).SubItems(Con2) & Chr(9)
                Next
                StrCopi = StrCopi & Chr(13)
            Next
            'archivo TXT
            Dim foo As Integer
            foo = FreeFile
            Open Ruta For Output As #foo
                Print #foo, StrCopi
            Close #foo
        End If
        ShellExecute Me.hWnd, "open", Ruta, "", "", 4
    End If
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub Text1_GotFocus()
    Text1.BackColor = &HFFE1E1
End Sub
Private Sub Text1_LostFocus()
    Text1.BackColor = &H80000005
End Sub
Private Sub Text2_GotFocus()
    Text2.BackColor = &HFFE1E1
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 13 Then
        Dim sBuscar As String
        Dim tRs As ADODB.Recordset
        sBuscar = "SELECT ID_REPA FROM REPAS WHERE CLAVE = '" & Text2.Text & "'"
        Set tRs = cnn.Execute(sBuscar)
        If Not (tRs.EOF And tRs.BOF) Then
            Me.Command2.Enabled = True
            Me.Command4.Enabled = True
            Text2.Text = ""
            Id_Repartidor = tRs.Fields("ID_REPA")
        Else
            Me.Command2.Enabled = False
            Me.Command4.Enabled = False
        End If
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Text2_LostFocus()
    Text2.BackColor = &H80000005
End Sub
