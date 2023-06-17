VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmLicitacion 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Capturar Licitación"
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11175
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   11175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   10080
      TabIndex        =   37
      Top             =   5640
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
         TabIndex        =   38
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmLicitacion.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "FrmLicitacion.frx":030A
         Top             =   120
         Width           =   720
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6735
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   11880
      _Version        =   393216
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Captura"
      TabPicture(0)   =   "FrmLicitacion.frx":23EC
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame5"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Detalle"
      TabPicture(1)   =   "FrmLicitacion.frx":2408
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Command2"
      Tab(1).Control(1)=   "ListView3"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Capturas Anteriores"
      TabPicture(2)   =   "FrmLicitacion.frx":2424
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "BtnQuitar"
      Tab(2).Control(1)=   "ListView4"
      Tab(2).ControlCount=   2
      Begin VB.CommandButton BtnQuitar 
         Caption         =   "Quitar"
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
         Left            =   -66480
         Picture         =   "FrmLicitacion.frx":2440
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   6120
         Width           =   1215
      End
      Begin VB.Frame Frame5 
         Caption         =   "Productos"
         Height          =   1935
         Left            =   5040
         TabIndex        =   27
         Top             =   4680
         Width           =   4695
         Begin VB.CommandButton Command1 
            Caption         =   "Agregar"
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
            Left            =   3240
            Picture         =   "FrmLicitacion.frx":4E12
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   1200
            Width           =   1215
         End
         Begin VB.TextBox Text4 
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   28
            Top             =   600
            Width           =   3135
         End
         Begin VB.TextBox Text5 
            Height          =   285
            Left            =   3360
            MaxLength       =   12
            TabIndex        =   8
            Top             =   600
            Width           =   1215
         End
         Begin VB.TextBox Text6 
            Height          =   285
            Left            =   120
            MaxLength       =   12
            TabIndex        =   9
            Top             =   1320
            Width           =   1335
         End
         Begin VB.TextBox Text8 
            Height          =   285
            Left            =   1560
            TabIndex        =   10
            Top             =   1320
            Width           =   1335
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "Producto"
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   360
            Width           =   3135
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Caption         =   "Precio"
            Height          =   255
            Left            =   3360
            TabIndex        =   31
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Cantidad Min."
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   1080
            Width           =   1095
         End
         Begin VB.Label Label6 
            Caption         =   "Cantidad Max."
            Height          =   255
            Left            =   1560
            TabIndex        =   29
            Top             =   1080
            Width           =   1215
         End
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Quitar"
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
         Left            =   -66480
         Picture         =   "FrmLicitacion.frx":77E4
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   6120
         Width           =   1215
      End
      Begin VB.Frame Frame1 
         Caption         =   "Seleccionar Productos"
         Height          =   4095
         Left            =   5040
         TabIndex        =   24
         Top             =   480
         Width           =   4695
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   120
            TabIndex        =   6
            Top             =   360
            Width           =   2895
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Por Descripción"
            Height          =   195
            Left            =   3120
            TabIndex        =   26
            Top             =   480
            Width           =   1455
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Por Clave"
            Height          =   195
            Left            =   3120
            TabIndex        =   25
            Top             =   240
            Value           =   -1  'True
            Width           =   1455
         End
         Begin MSComctlLib.ListView ListView2 
            Height          =   3135
            Left            =   120
            TabIndex        =   7
            Top             =   720
            Width           =   4455
            _ExtentX        =   7858
            _ExtentY        =   5530
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
      End
      Begin VB.Frame Frame2 
         Caption         =   "Seleccionar Cliente"
         Height          =   3735
         Left            =   120
         TabIndex        =   21
         Top             =   480
         Width           =   4815
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   120
            TabIndex        =   1
            Top             =   360
            Width           =   3255
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Por Clave"
            Height          =   195
            Left            =   3480
            TabIndex        =   23
            Top             =   480
            Width           =   1095
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Por Nombre"
            Height          =   195
            Left            =   3480
            TabIndex        =   22
            Top             =   240
            Value           =   -1  'True
            Width           =   1215
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   2655
            Left            =   120
            TabIndex        =   2
            Top             =   840
            Width           =   4575
            _ExtentX        =   8070
            _ExtentY        =   4683
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Contratos"
         Height          =   2295
         Left            =   120
         TabIndex        =   16
         Top             =   4320
         Width           =   4815
         Begin VB.CommandButton Command3 
            Caption         =   "Nuevo Contrato"
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
            Left            =   3000
            Picture         =   "FrmLicitacion.frx":A1B6
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   1800
            Width           =   1575
         End
         Begin VB.TextBox Text7 
            Height          =   285
            Left            =   1200
            MaxLength       =   25
            TabIndex        =   4
            Top             =   1920
            Width           =   1215
         End
         Begin VB.TextBox Text3 
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   17
            Top             =   480
            Width           =   4575
         End
         Begin VB.TextBox Text9 
            Height          =   285
            Left            =   1200
            MaxLength       =   25
            TabIndex        =   3
            Top             =   1560
            Width           =   1215
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   285
            Left            =   2520
            TabIndex        =   5
            Top             =   1200
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
            _Version        =   393216
            Format          =   50659329
            CurrentDate     =   38987
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   285
            Left            =   2520
            TabIndex        =   39
            Top             =   840
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
            _Version        =   393216
            Format          =   50659329
            CurrentDate     =   38987
         End
         Begin VB.Label Label10 
            Caption         =   "Fecha de Inicio del Contrato"
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   840
            Width           =   2295
         End
         Begin VB.Label Label8 
            Caption         =   "No. Contrato"
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   1920
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Cliente"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   4575
         End
         Begin VB.Label Label4 
            Caption         =   "Fecha de Termino del Contrato"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   1200
            Width           =   2295
         End
         Begin VB.Label Label9 
            Caption         =   "No. Licitación"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   1560
            Width           =   975
         End
      End
      Begin MSComctlLib.ListView ListView3 
         Height          =   5415
         Left            =   -74880
         TabIndex        =   12
         Top             =   480
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   9551
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
      Begin MSComctlLib.ListView ListView4 
         Height          =   5415
         Left            =   -74880
         TabIndex        =   35
         Top             =   480
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   9551
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
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   10080
      TabIndex        =   0
      Top             =   4440
      Width           =   975
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
         TabIndex        =   14
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image1 
         Height          =   735
         Left            =   120
         MouseIcon       =   "FrmLicitacion.frx":F2B8
         MousePointer    =   99  'Custom
         Picture         =   "FrmLicitacion.frx":F5C2
         Top             =   240
         Width           =   720
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   10440
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      PrinterDefault  =   0   'False
   End
End
Attribute VB_Name = "FrmLicitacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Dim ClvClie As String
Dim ClvClie2 As String
Dim IdLiciEli As String
Dim IdLiciEli2 As String
Private Sub btnQuitar_Click()
On Error GoTo ManejaError
    If VarMen.Text1(47).Text = "S" Then
        Dim sBuscar As String
        Dim tRs As ADODB.Recordset
        Dim tLi As ListItem
        sBuscar = "DELETE FROM LICITACIONES WHERE ID = " & IdLiciEli2
        Set tRs = cnn.Execute(sBuscar)
        sBuscar = "SELECT * FROM VsListLicita WHERE ID_CLIENTE = " & ClvClie2 & " AND FECHA_FIN >= '" & Format(Date, "dd/mm/yyyy") & "'"
        Set tRs = cnn.Execute(sBuscar)
        ListView4.ListItems.Clear
        If Not (tRs.EOF And tRs.BOF) Then
            Do While Not tRs.EOF
                Set tLi = Me.ListView4.ListItems.Add(, , tRs.Fields("ID"))
                tLi.SubItems(1) = tRs.Fields("NOMBRE")
                tLi.SubItems(2) = RTrim(tRs.Fields("ID_PRODUCTO"))
                tLi.SubItems(3) = tRs.Fields("Descripcion")
                tLi.SubItems(4) = tRs.Fields("PRECIO_VENTA")
                tLi.SubItems(5) = tRs.Fields("FECHA_FIN")
                tRs.MoveNext
            Loop
        End If
    Else
        MsgBox "NO CUENTA CON PERMISOS NECESARIOS PARA ESTA OPERACIÓN", vbInformation, "SACC"
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Command1_Click()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    sBuscar = "SELECT ID_CLIENTE FROM LICITACIONES WHERE ID_PRODUCTO = '" & Text4.Text & "' AND ID_CLIENTE = " & ClvClie & " AND FECHA_FIN >= '" & Format(Date, "dd/mm/yyyy") & "'"
    Set tRs = cnn.Execute(sBuscar)
    If (tRs.EOF And tRs.BOF) Then
        If Text7.Text <> "" Then
            sBuscar = "INSERT INTO LICITACIONES (ID_PRODUCTO, ID_CLIENTE, PRECIO_VENTA, FECHA_FIN, ID_USUARIO_CAPTURO, CANT_MIN, NO_CONTRATO, CANT_MAX, NO_LICITACION, FECHA_INICIO) VALUES ('" & Text4.Text & "', " & ClvClie & ", " & Text5.Text & ", '" & DTPicker1.Value & "', " & VarMen.Text1(0).Text & ", " & Text6.Text & ", '" & Text7.Text & "', " & Text8.Text & ", '" & Text9.Text & "', '" & Format(DTPicker2.Value, "dd/mm/yyyy") & "');"
            cnn.Execute (sBuscar)
            sBuscar = "SELECT * FROM VsListLicita WHERE ID_CLIENTE = " & ClvClie & " AND FECHA_FIN >= '" & Format(Date, "dd/mm/yyyy") & "'"
            Set tRs = cnn.Execute(sBuscar)
            ListView3.ListItems.Clear
            If Not (tRs.EOF And tRs.BOF) Then
                Do While Not tRs.EOF
                    Set tLi = Me.ListView3.ListItems.Add(, , tRs.Fields("ID"))
                    tLi.SubItems(1) = tRs.Fields("NOMBRE")
                    tLi.SubItems(2) = RTrim(tRs.Fields("ID_PRODUCTO"))
                    tLi.SubItems(3) = tRs.Fields("Descripcion")
                    tLi.SubItems(4) = tRs.Fields("PRECIO_VENTA")
                    tLi.SubItems(5) = tRs.Fields("FECHA_FIN")
                    tLi.SubItems(6) = tRs.Fields("CANT_MIN")
                    tLi.SubItems(7) = tRs.Fields("CANT_MAX")
                    tRs.MoveNext
                Loop
            End If
            Text4.Text = ""
            Text5.Text = ""
        Else
            MsgBox "NO SE HA CAPTURADO EL NUMERO DE CONTRATO, ES NECESARIO PARA EL REGISTRO!", vbInformation, "SACC"
        End If
    Else
        MsgBox "EXISTE UN REGISTRO DE UN PRECIO DE LICITACIÓN AUN VIGENTE PARA ESE CLIENTE EN ESE PRODUCTO, ELIMINE EL ANTERIOR PARA DAR UNO NUEVO", vbInformation, "SACC"
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Command2_Click()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    sBuscar = "DELETE FROM LICITACIONES WHERE ID = " & IdLiciEli
    Set tRs = cnn.Execute(sBuscar)
    sBuscar = "SELECT * FROM VsListLicita WHERE ID_CLIENTE = " & ClvClie & " AND FECHA_FIN >= '" & Format(Date, "dd/mm/yyyy") & "'"
    Set tRs = cnn.Execute(sBuscar)
    ListView3.ListItems.Clear
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = Me.ListView3.ListItems.Add(, , tRs.Fields("ID"))
            tLi.SubItems(1) = tRs.Fields("NOMBRE")
            tLi.SubItems(2) = RTrim(tRs.Fields("ID_PRODUCTO"))
            tLi.SubItems(3) = tRs.Fields("Descripcion")
            tLi.SubItems(4) = tRs.Fields("PRECIO_VENTA")
            tLi.SubItems(5) = tRs.Fields("FECHA_FIN")
            tRs.MoveNext
        Loop
    End If
    Me.Command2.Enabled = False
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Command3_Click()
    Text7.Enabled = True
    Text9.Enabled = True
    ListView1.Enabled = True
    Text1.Enabled = True
End Sub
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
On Error GoTo ManejaError
    DTPicker1.Value = Format(Date + 365, "dd/mm/yyyy")
    DTPicker2.Value = Format(Date, "dd/mm/yyyy")
    Me.Command1.Enabled = False
    Me.Command2.Enabled = False
    With ListView1
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "# DEL CLIENTE", 0
        .ColumnHeaders.Add , , "NOMBRE", 6100
        .ColumnHeaders.Add , , "DESCUENTO", 1200
        .ColumnHeaders.Add , , "DIAS DE CREDITO", 1200
        .ColumnHeaders.Add , , "LIMITE DE CREDITO", 1200
    End With
    With ListView2
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "CLV DEL PRODUCTO", 1600
        .ColumnHeaders.Add , , "Descripcion", 6800
        .ColumnHeaders.Add , , "PRECIO", 1500
    End With
    With ListView3
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "ID", 0
        .ColumnHeaders.Add , , "CLIENTE", 0
        .ColumnHeaders.Add , , "CLV PRODUCTO", 2700
        .ColumnHeaders.Add , , "Descripcion", 7000
        .ColumnHeaders.Add , , "PRECIO DE VENTA", 2000
        .ColumnHeaders.Add , , "FECHA FIN DEL CONTRATO", 2000
        .ColumnHeaders.Add , , "CANTIDAD MINIMA", 2000
        .ColumnHeaders.Add , , "CANTIDAD MAXIMA", 2000
    End With
    With ListView4
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "ID", 0
        .ColumnHeaders.Add , , "CLIENTE", 0
        .ColumnHeaders.Add , , "CLV PRODUCTO", 2700
        .ColumnHeaders.Add , , "Descripcion", 7000
        .ColumnHeaders.Add , , "PRECIO DE VENTA", 2000
        .ColumnHeaders.Add , , "FECHA FIN DEL CONTRATO", 2000
        .ColumnHeaders.Add , , "CANTIDAD MINIMA", 2000
        .ColumnHeaders.Add , , "CANTIDAD MAXIMA", 2000
    End With
    Set cnn = New ADODB.Connection
    With cnn
        .ConnectionString = _
            "Provider=" & NvoMen.TxtProvider.Text & ";Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & NvoMen.TxtBaseDatos.Text & ";Data Source=" & NvoMen.txtServidor.Text & ";"
        .Open
    End With
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Image1_Click()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim Con As Integer
    Dim POSY As Integer
    Dim tot As String
    tot = "0.00"
    CommonDialog1.Flags = 64
    CommonDialog1.CancelError = True
    CommonDialog1.ShowPrinter
    sBuscar = "SELECT * FROM VsLicitacion WHERE NO_CONTRATO = '" & Text7.Text & "' AND ID_CLIENTE = " & ClvClie
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Printer.Print ""
        Printer.Print ""
        Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(0).Text)) / 2
        Printer.Print VarMen.Text5(0).Text
        Printer.CurrentX = (Printer.Width - Printer.TextWidth("R.F.C. " & VarMen.Text5(8).Text)) / 2
        Printer.Print "R.F.C. " & VarMen.Text5(8).Text
        Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text)) / 2
        Printer.Print VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text
        Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text)) / 2
        Printer.Print VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text
        Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        Printer.Print "     FECHA LIMITE : " & tRs.Fields("FECHA_FIN")
        Printer.Print "     LICITACIÓN DEL CLIENTE : " & tRs.Fields("NOMBRE")
        Printer.Print "     NUMERO DE CONTRATO : " & tRs.Fields("NO_CONTRATO")
        Printer.Print "     ATIENDE : " & tRs.Fields("NOMBRE_USUARIO") & " " & tRs.Fields("APELLIDOS")
        Printer.Print ""
        Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        POSY = 2200
        Printer.CurrentY = POSY
        Printer.CurrentX = 600
        Printer.Print "Producto"
        Printer.CurrentY = POSY
        Printer.CurrentX = 2800
        Printer.Print "Descripcion"
        Printer.CurrentY = POSY
        Printer.CurrentX = 8600
        Printer.Print "P. Venta"
        Printer.CurrentY = POSY
        Printer.CurrentX = 9500
        Printer.Print "C. Minima"
        Printer.CurrentY = POSY
        Printer.CurrentX = 10400
        Printer.Print "Total"
        Do While Not tRs.EOF
            POSY = POSY + 200
            Printer.CurrentY = POSY
            Printer.CurrentX = 600
            Printer.Print tRs.Fields("ID_PRODUCTO")
            Printer.CurrentY = POSY
            Printer.CurrentX = 2800
            Printer.Print tRs.Fields("Descripcion")
            Printer.CurrentY = POSY
            Printer.CurrentX = 8600
            Printer.Print "$" & tRs.Fields("PRECIO_VENTA")
            Printer.CurrentY = POSY
            Printer.CurrentX = 9500
            Printer.Print tRs.Fields("CANT_MIN")
            Printer.CurrentY = POSY
            Printer.CurrentX = 10400
            Printer.Print "$" & Format(CDbl(tRs.Fields("PRECIO_VENTA")) * CDbl(tRs.Fields("CANT_MIN")), "###,###,##0.00")
            tot = Format(CDbl(tot) + (CDbl(tRs.Fields("PRECIO_VENTA")) * CDbl(tRs.Fields("CANT_MIN"))), "###,###,##0.00")
            tRs.MoveNext
            If POSY >= 14200 Then
                Printer.NewPage
                Printer.Print ""
                Printer.Print ""
                Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(0).Text)) / 2
                Printer.Print VarMen.Text5(0).Text
                Printer.CurrentX = (Printer.Width - Printer.TextWidth("R.F.C. " & VarMen.Text5(8).Text)) / 2
                Printer.Print "R.F.C. " & VarMen.Text5(8).Text
                Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text)) / 2
                Printer.Print VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text
                Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text)) / 2
                Printer.Print VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text
                Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
                Printer.Print "     FECHA LIMITE : " & tRs.Fields("FECHA_FIN")
                Printer.Print "     LICITACIÓN DEL CLIENTE : " & tRs.Fields("NOMBRE")
                Printer.Print "     NUMERO DE CONTRATO : " & tRs.Fields("NO_CONTRATO")
                Printer.Print "     ATIENDE : " & tRs.Fields("NOMBRE_USUARIO") & " " & tRs.Fields("APELLIDOS")
                Printer.Print ""
                Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
                POSY = 2200
                Printer.CurrentY = POSY
                Printer.CurrentX = 600
                Printer.Print "Producto"
                Printer.CurrentY = POSY
                Printer.CurrentX = 2800
                Printer.Print "Descripcion"
                Printer.CurrentY = POSY
                Printer.CurrentX = 8600
                Printer.Print "P. Venta"
                Printer.CurrentY = POSY
                Printer.CurrentX = 9500
                Printer.Print "C. Minima"
                Printer.CurrentY = POSY
                Printer.CurrentX = 10400
                Printer.Print "Total"
            End If
        Loop
        Printer.CurrentY = POSY + 200
        Printer.CurrentX = 0
        Printer.Print "-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        Printer.CurrentY = POSY + 400
        Printer.CurrentX = 9000
        Printer.Print "Total :   $ " & tot
        Printer.CurrentY = POSY + 600
        Printer.CurrentX = 500
        Printer.Print "FIN DEL LISTADO"
        Printer.EndDoc
        CommonDialog1.Copies = 1
    End If
Exit Sub
ManejaError:
    If Err.Number <> 32755 Then
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    End If
    Err.Clear
End Sub
Private Sub Image9_Click()
On Error GoTo ManejaError
    Unload Me
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo ManejaError
    ClvClie = Item
    Text3.Text = Item.SubItems(1)
    BuscaAnte
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub ListView1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And ListView1.ListItems.Count <> 0 Then
        Text7.SetFocus
    End If
End Sub
Private Sub ListView1_LostFocus()
    If Text3.Text <> "" Then
        Text1.Enabled = False
        ListView1.Enabled = False
    End If
End Sub
Private Sub ListView2_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo ManejaError
    Text4.Text = Item
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub ListView2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And ListView2.ListItems.Count <> 0 Then
        Text5.SetFocus
    End If
End Sub
Private Sub ListView3_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo ManejaError
    IdLiciEli = Item
    Me.Command2.Enabled = True
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub ListView4_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo ManejaError
    IdLiciEli2 = Item
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 13 And Text1.Text <> "" Then
        Me.ListView1.SetFocus
        Dim tRs As ADODB.Recordset
        Dim tLi As ListItem
        Dim sBus As String
        Dim CadClien As String
        If Option2.Value = True Then
            CadClien = Text1.Text
            CadClien = Replace(CadClien, " ", "%")
            sBus = "SELECT ID_CLIENTE, NOMBRE, DESCUENTO, DIAS_CREDITO, LIMITE_CREDITO FROM CLIENTE WHERE NOMBRE LIKE '%" & CadClien & "%'"
        Else
            sBus = "SELECT ID_CLIENTE, NOMBRE, DESCUENTO, DIAS_CREDITO, LIMITE_CREDITO FROM CLIENTE WHERE ID_CLIENTE = " & Text1.Text
        End If
        Set tRs = cnn.Execute(sBus)
        With tRs
            ListView1.ListItems.Clear
            Do While Not .EOF
                Set tLi = ListView1.ListItems.Add(, , .Fields("ID_CLIENTE") & "")
                If Not IsNull(.Fields("NOMBRE")) Then tLi.SubItems(1) = .Fields("NOMBRE") & ""
                If Not IsNull(.Fields("DESCUENTO")) Then
                    tLi.SubItems(2) = .Fields("DESCUENTO") & ""
                Else
                    tLi.SubItems(2) = "0.00"
                End If
                If Not IsNull(.Fields("DIAS_CREDITO")) Then tLi.SubItems(3) = .Fields("DIAS_CREDITO") & ""
                If Not IsNull(.Fields("LIMITE_CREDITO")) Then tLi.SubItems(4) = .Fields("LIMITE_CREDITO") & ""
                .MoveNext
            Loop
        End With
    End If
    If KeyAscii = 13 Then
        ListView1.SetFocus
    End If
    Dim Valido As String
    If Option1.Value = True Then
        Valido = "1234567890"
    Else
        Valido = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZ-1234567890. "
    End If
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
Private Sub Text1_GotFocus()
    Text1.BackColor = &HFFE1E1
End Sub
Private Sub Text1_LostFocus()
    Text1.BackColor = &H80000005
End Sub
Private Sub Text2_GotFocus()
    Text2.BackColor = &HFFE1E1
End Sub
Private Sub Text2_LostFocus()
    Text2.BackColor = &H80000005
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 13 And Text2.Text <> "" Then
        Dim tRs As ADODB.Recordset
        Dim tLi As ListItem
        Dim sBus As String
        Dim SUC As String
        If Option4.Value = True Then
            sBus = "SELECT ID_PRODUCTO, DESCRIPCION, GANANCIA, PRECIO_COSTO FROM ALMACEN3 WHERE ID_PRODUCTO LIKE '%" & Text2.Text & "%'"
        End If
        If Option3.Value = True Then
            sBus = "SELECT ID_PRODUCTO, DESCRIPCION, GANANCIA, PRECIO_COSTO FROM ALMACEN3 WHERE Descripcion LIKE '%" & Text2.Text & "%' "
        End If
        If sBus <> "" Then
            Set tRs = cnn.Execute(sBus)
            With tRs
                ListView2.ListItems.Clear
                Do While Not .EOF
                    If Not IsNull(.Fields("GANANCIA")) And Not IsNull(.Fields("PRECIO_COSTO")) Then
                        Set tLi = ListView2.ListItems.Add(, , .Fields("ID_PRODUCTO") & "")
                            If Not IsNull(.Fields("Descripcion")) Then tLi.SubItems(1) = .Fields("Descripcion") & ""
                            If Not IsNull(.Fields("GANANCIA")) And Not IsNull(.Fields("PRECIO_COSTO")) Then
                                tLi.SubItems(2) = Format((1 + CDbl(.Fields("GANANCIA"))) * CDbl(.Fields("PRECIO_COSTO")), "###,###,##0.00")
                            End If
                    End If
                    .MoveNext
                Loop
            End With
        End If
        Me.ListView2.SetFocus
    End If
    If KeyAscii = 13 Then
        ListView2.SetFocus
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
Private Sub Text3_Change()
On Error GoTo ManejaError
    If Text3.Text <> "" And Text4.Text <> "" And Text5.Text <> "" Then
        Me.Command1.Enabled = True
    Else
        Me.Command1.Enabled = False
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Text4_Change()
On Error GoTo ManejaError
    If Text3.Text <> "" And Text4.Text <> "" And Text5.Text <> "" Then
        Me.Command1.Enabled = True
    Else
        Me.Command1.Enabled = False
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Text5_GotFocus()
    Text5.BackColor = &HFFE1E1
    Text5.SelStart = 0
    Text5.SelLength = Len(Me.Text5.Text)
End Sub
Private Sub Text5_LostFocus()
    Text5.BackColor = &H80000005
End Sub
Private Sub Text5_Change()
On Error GoTo ManejaError
    If Text3.Text <> "" And Text4.Text <> "" And Text5.Text <> "" And Text6.Text <> "" Then
        Me.Command1.Enabled = True
    Else
        Me.Command1.Enabled = False
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Text5_KeyPress(KeyAscii As Integer)
On Error GoTo ManejaError
    If KeyAscii = 13 Then
        Text6.SetFocus
    End If
    Dim Valido As String
    Valido = "1234567890."
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
Private Sub Text6_Change()
    If Text3.Text <> "" And Text4.Text <> "" And Text5.Text <> "" And Text6.Text <> "" Then
        Me.Command1.Enabled = True
    Else
        Me.Command1.Enabled = False
    End If
End Sub
Private Sub Text6_GotFocus()
    Text6.SelStart = 0
    Text6.SelLength = Len(Me.Text6.Text)
End Sub
Private Sub Text6_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text8.SetFocus
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
Private Sub Text7_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        BuscaAnte
        Text2.SetFocus
    End If
    Dim Valido As String
    Valido = "1234567890.ABCDEFGHIJKLMNÑOPQRSTUVWXYZ _ç-,#~<>?¿!¡$@()/&%@!?*+"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Text7_LostFocus()
    If Text7.Text <> "" Then
        Text7.Enabled = False
        Text7.BackColor = &H80000005
    End If
End Sub
Private Sub Text7_GotFocus()
    Text7.BackColor = &HFFE1E1
End Sub
Private Sub Text8_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Command1.SetFocus
    End If
End Sub
Private Sub Text8_LostFocus()
    Text8.BackColor = &H80000005
End Sub
Private Sub Text8_GotFocus()
    Text8.BackColor = &HFFE1E1
    Text8.SelStart = 0
    Text8.SelLength = Len(Me.Text8.Text)
End Sub
Private Sub Text9_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    Valido = "1234567890.ABCDEFGHIJKLMNÑOPQRSTUVWXYZ _ç-,#~<>?¿!¡$@()/&%@!?*+"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Text9_LostFocus()
    If Text9.Text <> "" Then
        Text9.Enabled = False
        Text9.BackColor = &H80000005
    End If
End Sub
Private Sub Text9_GotFocus()
    Text9.BackColor = &HFFE1E1
End Sub
Private Sub BuscaAnte()
    If ClvClie <> "" And Text3.Text <> "" Then
        Dim tRs As ADODB.Recordset
        Dim sBuscar As String
        Dim tLi As ListItem
        sBuscar = "SELECT * FROM VsListLicita WHERE ID_CLIENTE = " & ClvClie & " AND FECHA_FIN >= '" & Format(Date, "dd/mm/yyyy") & "' AND NO_CONTRATO = '" & Text7.Text & "'"
        Set tRs = cnn.Execute(sBuscar)
        ListView4.ListItems.Clear
        If Not (tRs.EOF And tRs.BOF) Then
            Do While Not tRs.EOF
                Set tLi = Me.ListView4.ListItems.Add(, , tRs.Fields("ID"))
                tLi.SubItems(1) = tRs.Fields("NOMBRE")
                tLi.SubItems(2) = RTrim(tRs.Fields("ID_PRODUCTO"))
                tLi.SubItems(3) = tRs.Fields("Descripcion")
                tLi.SubItems(4) = tRs.Fields("PRECIO_VENTA")
                tLi.SubItems(5) = tRs.Fields("FECHA_FIN")
                tLi.SubItems(6) = tRs.Fields("CANT_MIN")
                tRs.MoveNext
            Loop
        End If
    End If
End Sub
