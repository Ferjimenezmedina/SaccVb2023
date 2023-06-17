VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmExportaProductos 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Esportar productos para factura electronica (Contpaq)"
   ClientHeight    =   7575
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11055
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   11055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9960
      TabIndex        =   3
      Top             =   6240
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmExportaProductos.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "FrmExportaProductos.frx":030A
         Top             =   120
         Width           =   720
      End
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
         TabIndex        =   4
         Top             =   960
         Width           =   975
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7215
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   12726
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Productos"
      TabPicture(0)   =   "FrmExportaProductos.frx":23EC
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "ListView1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "CommonDialog1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Command1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "ProgressBar1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Clientes"
      TabPicture(1)   =   "FrmExportaProductos.frx":2408
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   4800
         TabIndex        =   5
         Top             =   6720
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Exportar"
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
         Left            =   8400
         Picture         =   "FrmExportaProductos.frx":2424
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   6600
         Width           =   1215
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   1680
         Top             =   6600
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   6015
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   10610
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
End
Attribute VB_Name = "FrmExportaProductos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As adodb.Connection
Private cnnFox As adodb.Connection
Dim Tabla As String
Private Sub Command1_Click()
    Exporta
    Dim sBuscar As String
    Dim sDescrip As String
    Dim Con As Integer
    ProgressBar1.Visible = True
    ProgressBar1.Max = ListView1.ListItems.Count
    ProgressBar1.Min = 0
    For Con = 1 To ListView1.ListItems.Count
        sDescrip = Replace(ListView1.ListItems(Con).SubItems(2), "'", "")
        sBuscar = "INSERT INTO " & Tabla & " (Cidprodu01, Ccodigop01, Cnombrep01, Ctipopro01, Cfechaal01, Ccontrol01, Cdescrip01, Cmetodoc01, Cimpuesto1, Cprecio1, cidfotop01, cpesopro01, ccomvent01, ccomcobr01, ccostoes01, cmargenu01, cstatusp01, cidunida01, cidunida02, cfechabaja, cimpuesto2, cimpuesto3, cretenci01, cretenci02, cidpadre01, cidpadre02, cidpadre01,  cidpadre03, cidvalor01, cidvalor02, cidvalor03, cidvalor04, cidvalor05, cidvalor06, csegcont01, csegcont02, csegcont03, ctextoex01, ctextoex02, ctextoex03, cfechaex01, cimporte01, cimporte02, cimporte03, cimporte04, cprecio2, cprecio3, cprecio4, cprecio5, cprecio6, cprecio7, cprecio8, cprecio9, cprecio10, cbanunid01, cbancara01, cbanmeto01, cbanmaxmin, cbanprecio, cbanimpu01, cbancodi01, cbancomp01, ctimestamp, cerrorco01, cfechaer01, cprecioc01, cestadop01, cbanubic01, cesexento, cexisten01, ccostoext1, ccostoext2, ccostoext3, ccostoext4, ccostoext5, cfeccosex1, cfeccosex2, cfeccosex3, cfeccosex4, cfeccosex5, " & _
        "cmoncosex1, cmoncosex2, cmoncosex3, cmoncosex4, cmoncosex5, cbancosex, cescuotai2, cescuotai3, cidunicom, ciduniven, csubtipo, ccodaltern, cnomaltern, cdesccorta, cidmoneda, cusabascu, ctipopaque, cprecselec, cdesglosai, csegcont04, csegcont05, csegcont06, csegcont07, cnomodcomp) " & _
        "VALUES(" & Trim(ListView1.ListItems(Con)) & ", '" & Trim(ListView1.ListItems(Con).SubItems(1)) & "', '" & sDescrip & "', " & ListView1.ListItems(Con).SubItems(3) & ", {" & Format(ListView1.ListItems(Con).SubItems(4), "dd/mm/yyyy") & "}, " & ListView1.ListItems(Con).SubItems(5) & ", '" & sDescrip & "', " & ListView1.ListItems(Con).SubItems(7) & ", " & ListView1.ListItems(Con).SubItems(8) & ", " & ListView1.ListItems(Con).SubItems(9) & ", 0, 0, 0, 0, 0, 0, 1, 0, 0, {01/01/1800}, " & ListView1.ListItems(Con).SubItems(8) & ", " & ListView1.ListItems(Con).SubItems(8) & ", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '', '', '', '', '', '', {01/01/1800}, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, {01/01/1800  12:00:00:000}, 0, {01/01/1800}, " & ListView1.ListItems(Con).SubItems(9) & ", 0, 0, 0, 0, 0, 0, 0, 0, 0, {01/01/1800}, {01/01/1800}, {01/01/1800}, {01/01/1800}, {01/01/1800}, 0, 0, 0, 0, 0" & _
        "0, 0, 0, 0, 0, 0, 0, '', '', '', 0, 0, 0, 0, 0, '', '', '', '', 0);"
        'MsgBox sBuscar
        cnnFox.Execute (sBuscar)
        ProgressBar1.Value = Con
    Next
    ProgressBar1.Visible = False
    MsgBox "La expotación a finalizado exitosamente!", vbInformation, "SACC"
End Sub
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    Set cnn = New adodb.Connection
    With cnn
        .ConnectionString = _
            "Provider=SQLOLEDB.1;Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & NvoMen.TxtBaseDatos.Text & ";Data Source=" & NvoMen.txtServidor.Text & ";"
        .Open
    End With
    With ListView1
        .View = lvwReport
        .Gridlines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Numero", 2000
        .ColumnHeaders.Add , , "Clave del Producto", 2000
        .ColumnHeaders.Add , , "Descripción", 4100
        .ColumnHeaders.Add , , "Tipo", 1000
        .ColumnHeaders.Add , , "Fecha alta", 1000
        .ColumnHeaders.Add , , "Control", 1000
        .ColumnHeaders.Add , , "Descripción", 1000
        .ColumnHeaders.Add , , "Metodo", 1000
        .ColumnHeaders.Add , , "Impuesto", 1000
        .ColumnHeaders.Add , , "Precio", 1000
        .ColumnHeaders.Add , , "Costo", 1000
    End With
    Leer
End Sub
Private Sub Leer()
    Dim sBuscar As String
    Dim tRs As Recordset
    Dim tLi As ListItem
    sBuscar = "SELECT * FROM ALMACEN3"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView1.ListItems.Add(, , tRs.Fields("NUM"))
            tLi.SubItems(1) = tRs.Fields("ID_PRODUCTO")
            tLi.SubItems(2) = tRs.Fields("DESCRIPCION")
            tLi.SubItems(3) = "3"
            tLi.SubItems(4) = Date
            tLi.SubItems(5) = "0"
            tLi.SubItems(6) = tRs.Fields("DESCRIPCION")
            tLi.SubItems(7) = "7"
            tLi.SubItems(8) = CDbl(VarMen.Text4(7).Text) * 10
            tLi.SubItems(9) = tRs.Fields("PRECIO_COSTO") * (1 + tRs.Fields("GANANCIA"))
            tLi.SubItems(10) = tRs.Fields("PRECIO_COSTO")
            tRs.MoveNext
        Loop
    End If
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub Exporta()
On Error GoTo ManejaError
    Dim Ruta As String
    Dim Con As Integer
    Dim PPosi As Integer
    Dim NUMREG As Double
    Dim strArchivo As String
    Dim tLi As ListItem
    CommonDialog1.DialogTitle = "Abrir"
    CommonDialog1.Filter = "Base de Datos (*.dbf) |*.dbf|"
    Me.CommonDialog1.ShowOpen
    Ruta = Me.CommonDialog1.FileName
    If Ruta <> "" Then
        Ruta = Mid(Ruta, 1, Len(Ruta) - 4)
        For Con = 1 To Len(Ruta)
            If Mid(Ruta, Con, 1) = "\" Then
                PPosi = Con
            End If
        Next
        Tabla = Mid(Ruta, PPosi + 1, Len(Ruta))
        Ruta = Mid(Ruta, 1, PPosi - 1)
    End If
    Set cnnFox = New adodb.Connection
    With cnnFox
        .ConnectionString = _
            "Provider=MSDASQL.1; Presist Security Info=FALSE;Extended Properties=Driver={Microsoft Visual FoxPro Driver};UID=;SourceDB=" & Ruta & ";SourceType=DBF;Exclusive=No;BackgroundFetch=Yes;Collate=Machine;Null=Yes;Deleted=Yes;"
        .Open
    End With
    Exit Sub
ManejaError:
    If Err.Number = -2147217887 Then
        Err.Clear
        StatusBar1.Panels(3).Picture = Image1.Picture
        StatusBar1.Panels(3).Text = "ABIERTO CON ERRORES"
        StatusBar1.Panels(4).Text = ""
        StatusBar1.Panels(5).Text = ""
        StatusBar1.Panels(6).Text = ""
    Else
        If Err.Number = -2147217865 Then
            MsgBox "EL NOMBRE DE LA TABLA NO EXISTE O EL ARCHIVO NO ES UN DBF", vbCritical, "Hache's system"
        Else
            If Err.Number <> -2147467259 Then
                If Err.Number = -2147352571 Then
                    MsgBox "OCURRIO UN ERROR AL ABRIR LA TABLA, ES POSIBLE QUE NO SE MUESTRE TODA LA INFORMACIÓN DE ESTA, PUEDE FILTRAR LA INFORMACIÓN PARA SECCIONARLA", vbCritical, "Hache's system"
                Else
                    If Err.Number <> 0 Then
                        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "Hache's system"
                    End If
                End If
            End If
        End If
    End If
    Err.Clear
End Sub
