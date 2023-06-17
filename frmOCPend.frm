VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmOCPend 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ORDENES DE COMPRA CON PRODUCTOS PENDIENTES DE LLEGAR"
   ClientHeight    =   7695
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10215
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   10215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   7455
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   13150
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   " "
      TabPicture(0)   =   "frmOCPend.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lvOC"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lvDet"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      Begin MSComctlLib.ListView lvDet 
         Height          =   3255
         Left            =   120
         TabIndex        =   1
         Top             =   4080
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   5741
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView lvOC 
         Height          =   3255
         Left            =   120
         TabIndex        =   0
         Top             =   480
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   5741
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Ordenes de Compra"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   4815
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Ordenes de Compra"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   3840
         Width           =   4815
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9120
      TabIndex        =   2
      Top             =   6360
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
         TabIndex        =   3
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "frmOCPend.frx":001C
         MousePointer    =   99  'Custom
         Picture         =   "frmOCPend.frx":0326
         Top             =   120
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmOCPend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
On Error GoTo ManejaError
    Set cnn = New ADODB.Connection
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    Dim sBus As String
    With cnn
        .ConnectionString = _
            "Provider=" & NvoMen.TxtProvider.Text & ";Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & NvoMen.TxtBaseDatos.Text & ";Data Source=" & NvoMen.txtServidor.Text & ";"
        .Open
    End With
    With lvOC
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "ID_ORDEN", 0
        .ColumnHeaders.Add , , "TIPO", 1500
        .ColumnHeaders.Add , , "# O.C.", 500
        .ColumnHeaders.Add , , "NOMBRE PROVEEDOR", 3440
        .ColumnHeaders.Add , , "TOTAL A PAGAR", 1440
        .ColumnHeaders.Add , , "FECHA", 1000
    End With
    With lvDet
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "ID_PRODUCTO", 1000
        .ColumnHeaders.Add , , "Descripcion", 3000
        .ColumnHeaders.Add , , "PRECIO", 1500
        .ColumnHeaders.Add , , "CANT. PENDIENTE", 1500
        .ColumnHeaders.Add , , "ID", 0
        .ColumnHeaders.Add , , "CANT. SURTIDA", 1000
    End With
    sBus = "SELECT V.ID_ORDEN_COMPRA, V.TIPO, V.NUM_ORDEN, V.TOTAL, V.FECHA, P.NOMBRE FROM VSORDENESP AS V JOIN PROVEEDOR AS P ON P.ID_PROVEEDOR = V.ID_PROVEEDOR ORDER BY TIPO, NUM_ORDEN"
    Set tRs = cnn.Execute(sBus)
    If Not (tRs.EOF And tRs.BOF) Then
        With tRs
            Do While Not .EOF
                Set tLi = lvOC.ListItems.Add(, , .Fields("ID_ORDEN_COMPRA"))
                If .Fields("TIPO") = "N" Then
                    tLi.SubItems(1) = "NACIONAL"
                ElseIf .Fields("TIPO") = "I" Then
                    tLi.SubItems(1) = "INTERNACIONAL"
                Else
                    tLi.SubItems(1) = "INDIRECTA"
                End If
                tLi.SubItems(2) = .Fields("NUM_ORDEN")
                tLi.SubItems(3) = .Fields("NOMBRE")
                If Not IsNull(.Fields("TOTAL")) Then tLi.SubItems(4) = .Fields("TOTAL")
                tLi.SubItems(5) = .Fields("FECHA")
                .MoveNext
            Loop
        End With
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub lvOC_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo ManejaError
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    Dim sBus As String
    If lvOC.ListItems.Count > 0 Then
        sBus = "SELECT ID, ID_PRODUCTO, DESCRIPCION, PRECIO, CANTIDADP, SURTIDO FROM VSORDENES WHERE ID_ORDEN_COMPRA = " & Item & " AND ( CANTIDADP > 0)"
        Set tRs = cnn.Execute(sBus)
        With tRs
            lvDet.ListItems.Clear
            Do While Not .EOF
                Set tLi = lvDet.ListItems.Add(, , .Fields("ID_PRODUCTO") & "")
                If Not IsNull(.Fields("Descripcion")) Then tLi.SubItems(1) = .Fields("Descripcion") & ""
                If Not IsNull(.Fields("PRECIO")) Then tLi.SubItems(2) = Format(CDbl(.Fields("PRECIO")), "###,###,##0.00")
                If Not IsNull(.Fields("CANTIDADP")) Then tLi.SubItems(3) = .Fields("CANTIDADP") & ""
                tLi.SubItems(4) = .Fields("ID") & ""
                If Not IsNull(.Fields("SURTIDO")) Then
                    tLi.SubItems(5) = .Fields("SURTIDO") & ""
                Else
                    tLi.SubItems(5) = "0"
                End If
                
                .MoveNext
            Loop
        End With
    End If
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
