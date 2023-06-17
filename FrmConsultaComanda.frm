VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmConsultaComanda 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consultar cobro por comanda"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7815
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   7815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TxtNomCliente 
      Height          =   285
      Left            =   7560
      TabIndex        =   17
      Text            =   "Text5"
      Top             =   240
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   7200
      TabIndex        =   16
      Text            =   "Combo1"
      Top             =   240
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.TextBox txtComandas 
      Height          =   285
      Left            =   6960
      TabIndex        =   14
      Text            =   "Text5"
      Top             =   240
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox TxtNomCli 
      Height          =   285
      Left            =   6840
      TabIndex        =   13
      Text            =   "Text5"
      Top             =   240
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtIDCOMANDA 
      Height          =   285
      Left            =   6720
      TabIndex        =   12
      Text            =   "Text5"
      Top             =   240
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   6720
      TabIndex        =   10
      Top             =   1800
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
         TabIndex        =   11
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmConsultaComanda.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "FrmConsultaComanda.frx":030A
         Top             =   120
         Width           =   720
      End
   End
   Begin VB.CommandButton cmdVer 
      Caption         =   "Buscar"
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
      Picture         =   "FrmConsultaComanda.frx":23EC
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4560
      TabIndex        =   7
      Text            =   "0.00"
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2640
      TabIndex        =   5
      Text            =   "0.00"
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   840
      TabIndex        =   3
      Text            =   "0.00"
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   240
      Width           =   1815
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   3201
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label32 
      Caption         =   "Label5"
      Height          =   255
      Left            =   7080
      TabIndex        =   15
      Top             =   240
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Total :"
      Height          =   255
      Left            =   3960
      TabIndex        =   8
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "IVA :"
      Height          =   255
      Left            =   2160
      TabIndex        =   6
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Subtotal :"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Comanda :"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "FrmConsultaComanda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Private Sub cmdVer_Click()
    ExtraeComanda
End Sub
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
On Error GoTo ManejaError
    With ListView1
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "CLV DEL PRODUCTO", 2000
        .ColumnHeaders.Add , , "Descripcion", 6800
        .ColumnHeaders.Add , , "CANTIDAD", 2000
        .ColumnHeaders.Add , , "PRECIO", 2000
        .ColumnHeaders.Add , , "COMANDA O ASISTENCIA", 2000
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
Private Sub ExtraeComanda()
On Error GoTo ManejaError
    Dim tRs As ADODB.Recordset
    Dim tRs1 As ADODB.Recordset
    Dim tRs2 As ADODB.Recordset
    Dim tRs3 As ADODB.Recordset
    Dim tRs4 As ADODB.Recordset
    Dim tLi As ListItem
    Dim sBuscar As String
    Dim cant As String
    Dim PreTot As String
    Dim PreTotDes As String
    Dim ClvProd As String
    Dim CLVCLIEN As Integer
    Dim NomClien As String
    Dim DesClien As String
    sBuscar = "SELECT ID_CLIENTE FROM COMANDAS_2 WHERE ID_COMANDA = " & Text1.Text
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        sBuscar = "SELECT ID_PRODUCTO, CANTIDAD, CANTIDAD_NO_SIRVIO FROM COMANDAS_DETALLES_2 WHERE ID_COMANDA = " & Text1.Text & " AND ESTADO_ACTUAL IN ('L','N')"
        Set tRs1 = cnn.Execute(sBuscar)
        If Not (tRs1.EOF And tRs1.BOF) Then
            sBuscar = "SELECT ID_CLIENTE, NOMBRE, DESCUENTO, DIAS_CREDITO, LIMITE_CREDITO FROM CLIENTE WHERE ID_CLIENTE = " & tRs.Fields("ID_CLIENTE")
            Set tRs = cnn.Execute(sBuscar)
            If tRs.EOF And tRs.BOF Then
                MsgBox "EL CLIENTE HA SIDO ELIMINADO, DEBE REASIGNAR UN CLIENTE!", vbInformation, "SACC"
            Else
                Me.TxtNomCliente.Text = tRs.Fields("NOMBRE")
                If Not IsNull(tRs.Fields("NOMBRE")) Then TxtNomCli.Text = tRs.Fields("NOMBRE")
                If Not IsNull(tRs.Fields("ID_CLIENTE")) Then
                    CLVCLIEN = tRs.Fields("ID_CLIENTE")
                    sBuscar = "SELECT ID_DESCUENTO FROM CLIENTE WHERE ID_CLIENTE = " & tRs.Fields("ID_CLIENTE")
                    Set tRs2 = cnn.Execute(sBuscar)
                    If Not (tRs2.EOF And tRs2.BOF) Then
                        If Not IsNull(tRs2.Fields("ID_DESCUENTO")) Then Label32.Caption = tRs2.Fields("ID_DESCUENTO")
                    End If
                End If
                If Not IsNull(tRs.Fields("NOMBRE")) Then NomClien = tRs.Fields("NOMBRE")
                If Not IsNull(tRs.Fields("DESCUENTO")) Then DesClien = tRs.Fields("DESCUENTO")
                If DesClien = "" Then DesClien = "0"
                TxtNomCli.Enabled = False
                ListView1.Enabled = False
                If Not (tRs1.BOF And tRs1.EOF) Then
                    If txtComandas.Text = "" Then
                        txtComandas.Text = Text1.Text
                    Else
                        txtComandas.Text = txtComandas.Text & ", " & Text1.Text
                    End If
                    Do While Not tRs1.EOF
                        sBuscar = "SELECT ID_PRODUCTO, DESCRIPCION, PRECIO_COSTO, GANANCIA, CLASIFICACION FROM ALMACEN3 WHERE ID_PRODUCTO = '" & tRs1.Fields("ID_PRODUCTO") & "'"
                        Set tRs2 = cnn.Execute(sBuscar)
                        If Not (tRs2.BOF And tRs2.EOF) Then
                             ClvProd = tRs2.Fields("ID_PRODUCTO")
                             If CDbl(tRs1.Fields("CANTIDAD")) - CDbl(tRs1.Fields("CANTIDAD_NO_SIRVIO")) <> 0 Then
                                Set tLi = ListView1.ListItems.Add(, , tRs2.Fields("ID_PRODUCTO"))
                                    If Not IsNull(tRs2.Fields("Descripcion")) Then tLi.SubItems(1) = tRs2.Fields("Descripcion")
                                    If Not IsNull(tRs1.Fields("CANTIDAD")) Then tLi.SubItems(2) = CDbl(tRs1.Fields("CANTIDAD")) - CDbl(tRs1.Fields("CANTIDAD_NO_SIRVIO"))
                                    If Label32.Caption = "" Then
                                    If Combo1.Text = "<NINGUNA>" Or Combo1.Text = "" Then
                                        If Label32.Caption <> "" Then
                                            sBuscar = "SELECT PORCENTAJE FROM DESCUENTOS WHERE ID_DESCUENTO = '" & Label32.Caption & "' AND CLASIFICACION = '" & tRs2.Fields("CLASIFICACION") & "'"
                                            Set tRs3 = cnn.Execute(sBuscar)
                                            If Not (tRs3.EOF And tRs3.BOF) Then
                                                PreTot = Format(CDbl(tRs2.Fields("PRECIO_COSTO")) * (1 + CDbl(tRs2.Fields("GANANCIA"))) - (CDbl(tRs2.Fields("PRECIO_COSTO")) * (1 + CDbl(tRs2.Fields("GANANCIA"))) * (1 - (CDbl(tRs3.Fields("PORCENTAJE") / 100)))), "###,###,##0.00")
                                            Else
                                                If Not IsNull(tRs2.Fields("PRECIO_COSTO")) And Not IsNull(tRs2.Fields("GANANCIA")) Then PreTot = Format(CDbl(tRs2.Fields("PRECIO_COSTO")) * (CDbl(tRs2.Fields("GANANCIA")) + 1), "###,###,##0.00")
                                                PreTot = Val(Replace(PreTot, ",", "")) * Val(Replace(tLi.SubItems(2), ",", ""))
                                                If DesClien <> "" Then
                                                    PreTot = Val(Replace(PreTot, ",", "")) * ((100 - Val(Replace(DesClien, ",", ""))) / 100)
                                                End If
                                            End If
                                        Else
                                            PreTot = Format(CDbl((tRs2.Fields("PRECIO_COSTO")) * (CDbl(tRs2.Fields("GANANCIA")) + 1)) * (CDbl(tRs1.Fields("CANTIDAD")) - CDbl(tRs1.Fields("CANTIDAD_NO_SIRVIO"))), "###,###,##0.00")
                                        End If
                                    Else
                                        If Combo1.Text = "LICITACIÓN" Then
                                            sBuscar = "SELECT PRECIO_VENTA FROM LICITACIONES WHERE ID_PRODUCTO = '" & ClvProd & "' AND FECHA_FIN >= '" & Format(Date, "dd/mm/yyyy") & "' AND ID_CLIENTE = " & CLVCLIEN
                                            Set tRs3 = cnn.Execute(sBuscar)
                                            If Not (tRs3.EOF And tRs3.BOF) Then
                                                PreTot = Format(CDbl(tRs3.Fields("PRECIO_VENTA")) * Val(Replace(tLi.SubItems(2), ",", "")), "###,###,##0.00")
                                            Else
                                                PreTot = Format(CDbl(tRs2.Fields("PRECIO_COSTO")) * (CDbl(tRs2.Fields("GANANCIA")) + 1), "###,###,##0.00")
                                                PreTot = Val(Replace(PreTot, ",", "")) * Val(Replace(tLi.SubItems(2), ",", ""))
                                            End If
                                        Else
                                            If VarMen.Text4(8).Text = "S" Then
                                                sBuscar = "SELECT PRECIO_OFERTA FROM PROMOCION WHERE ID_PRODUCTO = '" & ClvProd & "' AND FECHA_FIN >= '" & Format(Date, "dd/mm/yyyy") & "' AND TIPO = '" & Combo1.Text & "'"
                                                Set tRs3 = cnn.Execute(sBuscar)
                                                If Not (tRs3.EOF And tRs3.BOF) Then
                                                    PreTot = CDbl(tRs3.Fields("PRECIO_OFERTA")) * Val(Replace(tLi.SubItems(2), ",", ""))
                                                    If Not IsNull(tRs2.Fields("PRECIO_COSTO")) And Not IsNull(tRs2.Fields("GANANCIA")) Then PreTotDes = Format(CDbl(tRs2.Fields("PRECIO_COSTO")) * (CDbl(tRs2.Fields("GANANCIA")) + 1), "###,###,##0.00")
                                                    PreTotDes = Val(Replace(PreTot, ",", "")) * Val(Replace(tLi.SubItems(2), ",", ""))
                                                    If DesClien <> "" Then
                                                        PreTotDes = Val(Replace(PreTotDes, ",", "")) * ((100 - Val(Replace(DesClien, ",", ""))) / 100)
                                                    End If
                                                    If Val(Replace(PreTot, ",", "")) > Val(Replace(PreTotDes, ",", "")) Then PreTot = PreTotDes
                                                Else
                                                    sBuscar = "SELECT PORCENTAJE FROM DESCUENTOS WHERE ID_DESCUENTO = '" & Label32.Caption & "' AND CLASIFICACION = '" & tRs2.Fields("CLASIFICACION") & "'"
                                                    Set tRs3 = cnn.Execute(sBuscar)
                                                    If Not (tRs3.EOF And tRs3.BOF) Then
                                                        PreTot = Format(CDbl(tRs2.Fields("PRECIO_COSTO")) * (1 + CDbl(tRs2.Fields("GANANCIA"))) - (CDbl(tRs2.Fields("PRECIO_COSTO")) * (1 + CDbl(tRs2.Fields("GANANCIA"))) * (1 - (CDbl(tRs3.Fields("PORCENTAJE") / 100)))), "###,###,##0.00")
                                                    Else
                                                        If Not IsNull(tRs2.Fields("PRECIO_COSTO")) And Not IsNull(tRs2.Fields("GANANCIA")) Then PreTot = Format(CDbl(tRs2.Fields("PRECIO_COSTO")) * (CDbl(tRs2.Fields("GANANCIA")) + 1), "###,###,##0.00")
                                                        PreTot = Val(Replace(PreTot, ",", "")) * Val(Replace(tLi.SubItems(2), ",", ""))
                                                        If DesClien <> "" Then
                                                            PreTot = Val(Replace(PreTot, ",", "")) * ((100 - Val(Replace(DesClien, ",", ""))) / 100)
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                sBuscar = "SELECT PORCENTAJE FROM DESCUENTOS WHERE ID_DESCUENTO = '" & Label32.Caption & "' AND CLASIFICACION = '" & tRs2.Fields("CLASIFICACION") & "'"
                                                Set tRs3 = cnn.Execute(sBuscar)
                                                If Not (tRs3.EOF And tRs3.BOF) Then
                                                    PreTot = Format(CDbl(tRs2.Fields("PRECIO_COSTO")) * (1 + CDbl(tRs2.Fields("GANANCIA"))) - (CDbl(tRs2.Fields("PRECIO_COSTO")) * (1 + CDbl(tRs2.Fields("GANANCIA"))) * (1 - (CDbl(tRs3.Fields("PORCENTAJE") / 100)))), "###,###,##0.00")
                                                Else
                                                    If Not IsNull(tRs2.Fields("PRECIO_COSTO")) And Not IsNull(tRs2.Fields("GANANCIA")) Then PreTot = Format(CDbl(tRs2.Fields("PRECIO_COSTO")) * (CDbl(tRs2.Fields("GANANCIA")) + 1), "###,###,##0.00")
                                                    PreTot = Val(Replace(PreTot, ",", "")) * Val(Replace(tLi.SubItems(2), ",", ""))
                                                    If DesClien <> "" Then
                                                        PreTot = Val(Replace(PreTot, ",", "")) * ((100 - Val(Replace(DesClien, ",", ""))) / 100)
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                Else
                                    sBuscar = "SELECT PORCENTAJE FROM DESCUENTOS WHERE ID_DESCUENTO = '" & Label32.Caption & "' AND CLASIFICACION = '" & tRs2.Fields("CLASIFICACION") & "'"
                                    Set tRs4 = cnn.Execute(sBuscar)
                                    If Not (tRs4.EOF And tRs4.BOF) Then
                                        PreTot = Val(tRs2.Fields("PRECIO_COSTO") * (tRs2.Fields("GANANCIA") + 1))
                                        PreTot = CDbl(PreTot) * (1 - (tRs4.Fields("PORCENTAJE") / 100)) * (tRs1.Fields("CANTIDAD") - tRs1.Fields("CANTIDAD_NO_SIRVIO"))
                                    Else
                                        If Combo1.Text = "<NINGUNA>" Or Combo1.Text = "" Then
                                            If Not IsNull(tRs2.Fields("PRECIO_COSTO")) And Not IsNull(tRs2.Fields("GANANCIA")) Then PreTot = Format(CDbl(tRs2.Fields("PRECIO_COSTO")) * (CDbl(tRs2.Fields("GANANCIA")) + 1), "###,###,##0.00")
                                            PreTot = Val(Replace(PreTot, ",", "")) * Val(Replace(tLi.SubItems(2), ",", ""))
                                            If DesClien <> "" Then
                                                PreTot = Val(Replace(PreTot, ",", "")) * ((100 - Val(Replace(DesClien, ",", ""))) / 100)
                                            End If
                                        Else
                                            If Combo1.Text = "LICITACIÓN" Then
                                                sBuscar = "SELECT PRECIO_VENTA FROM LICITACIONES WHERE ID_PRODUCTO = '" & ClvProd & "' AND FECHA_FIN >= '" & Format(Date, "dd/mm/yyyy") & "' AND ID_CLIENTE = " & CLVCLIEN
                                                Set tRs3 = cnn.Execute(sBuscar)
                                                If Not (tRs3.EOF And tRs3.BOF) Then
                                                    PreTot = Format(CDbl(tRs3.Fields("PRECIO_VENTA")) * Val(Replace(tLi.SubItems(2), ",", "")), "###,###,##0.00")
                                                Else
                                                    PreTot = Format(CDbl(tRs2.Fields("PRECIO_COSTO")) * (CDbl(tRs2.Fields("GANANCIA")) + 1), "###,###,##0.00")
                                                    PreTot = Val(Replace(PreTot, ",", "")) * Val(Replace(tLi.SubItems(2), ",", ""))
                                                End If
                                            Else
                                                If VarMen.Text4(8).Text = "S" Then
                                                    sBuscar = "SELECT PRECIO_OFERTA FROM PROMOCION WHERE ID_PRODUCTO = '" & ClvProd & "' AND FECHA_FIN >= '" & Format(Date, "dd/mm/yyyy") & "' AND TIPO = '" & Combo1.Text & "'"
                                                    Set tRs3 = cnn.Execute(sBuscar)
                                                    If Not (tRs3.EOF And tRs3.BOF) Then
                                                        PreTot = CDbl(tRs3.Fields("PRECIO_OFERTA")) * Val(Replace(tLi.SubItems(2), ",", ""))
                                                        If Not IsNull(tRs2.Fields("PRECIO_COSTO")) And Not IsNull(tRs2.Fields("GANANCIA")) Then PreTotDes = Format(CDbl(tRs2.Fields("PRECIO_COSTO")) * (CDbl(tRs2.Fields("GANANCIA")) + 1), "###,###,##0.00")
                                                        PreTotDes = Val(Replace(PreTot, ",", "")) * Val(Replace(tLi.SubItems(2), ",", ""))
                                                        If DesClien <> "" Then
                                                            PreTotDes = Val(Replace(PreTotDes, ",", "")) * ((100 - Val(Replace(DesClien, ",", ""))) / 100)
                                                        End If
                                                        If Val(Replace(PreTot, ",", "")) > Val(Replace(PreTotDes, ",", "")) Then PreTot = PreTotDes
                                                    Else
                                                        If Not IsNull(tRs2.Fields("PRECIO_COSTO")) And Not IsNull(tRs2.Fields("GANANCIA")) Then PreTot = Format(CDbl(tRs2.Fields("PRECIO_COSTO")) * (CDbl(tRs2.Fields("GANANCIA")) + 1), "###,###,##0.00")
                                                        PreTot = Val(Replace(PreTot, ",", "")) * Val(Replace(tLi.SubItems(2), ",", ""))
                                                        If DesClien <> "" Then
                                                            PreTot = Val(Replace(PreTot, ",", "")) * ((100 - Val(Replace(DesClien, ",", ""))) / 100)
                                                        End If
                                                    End If
                                                Else
                                                    If Not IsNull(tRs2.Fields("PRECIO_COSTO")) And Not IsNull(tRs2.Fields("GANANCIA")) Then PreTot = Format(CDbl(tRs2.Fields("PRECIO_COSTO")) * (CDbl(tRs2.Fields("GANANCIA")) + 1), "###,###,##0.00")
                                                    PreTot = Val(Replace(PreTot, ",", "")) * Val(Replace(tLi.SubItems(2), ",", ""))
                                                    If DesClien <> "" Then
                                                        PreTot = Val(Replace(PreTot, ",", "")) * ((100 - Val(Replace(DesClien, ",", ""))) / 100)
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                                tLi.SubItems(3) = Replace(PreTot, ",", "")
                                tLi.SubItems(4) = "C" & Text1.Text
                                If Not IsNull(tRs1.Fields("CANTIDAD")) Then Text2.Text = Format(Val(Replace(Text2.Text, ",", "")) + Val(Replace(PreTot, ",", "")), "###,###,##0.00")
                             End If
                        End If
                        tRs1.MoveNext
                    Loop
                    Text3.Text = Format(Val(Replace(Text2.Text, ",", "")) * CDbl(CDbl(VarMen.Text4(7).Text) / 100), "###,###,##0.00")
                    Text4.Text = Format(CDbl(Text2.Text) * CDbl(1 + (CDbl(VarMen.Text4(7).Text) / 100)), "###,###,##0.00")
                Else
                    MsgBox "NO SE PUEDEN COBRAR JUNTAS COMANDAS DE CLIENTES DISTINTOS!", vbInformation, "SACC"
                End If
            End If
        Else
            MsgBox "LA COMANDA NO HA SIDO FINALIZADA O YA FUE ENTREGADA AL CLIENTE!", vbInformation, "SACC"
            txtIDCOMANDA.Text = ""
        End If
    Else
        MsgBox "NO EXISTE LA COMANDA!", vbInformation, "SACC"
        txtIDCOMANDA.Text = ""
    End If
    Text1.Text = ""
    Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ExtraeComanda
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
