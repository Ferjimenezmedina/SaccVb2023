VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmSurtir 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Surtidos"
   ClientHeight    =   7815
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   10935
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   10935
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9840
      TabIndex        =   15
      Top             =   6480
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
         TabIndex        =   16
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "frmSurtir.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "frmSurtir.frx":030A
         Top             =   120
         Width           =   720
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9840
      TabIndex        =   13
      Top             =   5280
      Width           =   975
      Begin VB.Image Image1 
         Height          =   735
         Left            =   120
         MouseIcon       =   "frmSurtir.frx":23EC
         MousePointer    =   99  'Custom
         Picture         =   "frmSurtir.frx":26F6
         Top             =   240
         Width           =   720
      End
      Begin VB.Label Label5 
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
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7575
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   13361
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   " "
      TabPicture(0)   =   "frmSurtir.frx":42C8
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Command4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Command3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtCant"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdSurtir"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame2"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      Begin VB.Frame Frame2 
         Height          =   2895
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   9375
         Begin MSComctlLib.ListView lvwSurtir 
            Height          =   2295
            Left            =   120
            TabIndex        =   0
            Top             =   480
            Width           =   9135
            _ExtentX        =   16113
            _ExtentY        =   4048
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483633
            Appearance      =   0
            NumItems        =   0
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "PEDIDOS PENDIENTES"
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
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   9135
         End
      End
      Begin VB.CommandButton cmdSurtir 
         Caption         =   "Surtir"
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
         Picture         =   "frmSurtir.frx":42E4
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   7080
         Width           =   1095
      End
      Begin VB.Frame Frame1 
         Height          =   2895
         Left            =   120
         TabIndex        =   7
         Top             =   4080
         Width           =   9375
         Begin MSComctlLib.ListView lvwSurtir2 
            Height          =   2295
            Left            =   120
            TabIndex        =   3
            Top             =   480
            Width           =   9135
            _ExtentX        =   16113
            _ExtentY        =   4048
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483633
            Appearance      =   0
            NumItems        =   0
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "SURTIR A SUCURSALES"
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
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   9135
         End
      End
      Begin VB.TextBox txtCant 
         Height          =   285
         Left            =   5640
         TabIndex        =   4
         Top             =   7080
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Surtir"
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
         Picture         =   "frmSurtir.frx":6CB6
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   3600
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         Caption         =   "S. Todo"
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
         Left            =   7200
         Picture         =   "frmSurtir.frx":9688
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   3600
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "--------------------------------------------------------------------"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   7080
         Width           =   4215
      End
      Begin VB.Label Label4 
         Caption         =   "Cantidad"
         Height          =   255
         Left            =   4680
         TabIndex        =   11
         Top             =   7080
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmSurtir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Dim ItMx As ListItem
Dim bLvw As Byte
Sub Llenar_Lista_Surtidos(cEstado As String)
On Error GoTo ManejaError
    Me.lvwSurtir.ListItems.Clear
    Dim tRs As ADODB.Recordset
    Dim sBuscar As String
    sBuscar = "SELECT * FROM vsPedido_PD"
    Set tRs = cnn.Execute(sBuscar)
    While Not tRs.EOF
        Set ItMx = Me.lvwSurtir.ListItems.Add(, , tRs.Fields("ID_PEDIDO"))
        If Not IsNull(tRs.Fields("Sucursal")) Then ItMx.SubItems(1) = Trim(tRs.Fields("Sucursal"))
        If Not IsNull(tRs.Fields("PIDIO")) Then ItMx.SubItems(2) = tRs.Fields("PIDIO")
        If Not IsNull(tRs.Fields("fecha")) Then ItMx.SubItems(3) = tRs.Fields("fecha")
        If Not IsNull(tRs.Fields("ID")) Then ItMx.SubItems(4) = tRs.Fields("ID")
        If Not IsNull(tRs.Fields("ID_PRODUCTO")) Then ItMx.SubItems(5) = tRs.Fields("ID_PRODUCTO")
        If Not IsNull(tRs.Fields("CANTIDAD")) Then ItMx.SubItems(6) = tRs.Fields("CANTIDAD")
        If Not IsNull(tRs.Fields("Descripcion")) Then ItMx.SubItems(7) = tRs.Fields("Descripcion")
        tRs.MoveNext
    Wend
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub cmdSurtir_Click()
On Error GoTo ManejaError
Dim cant As Integer
Dim sBuscar As String
Dim Estado As String
    If lvwSurtir2.ListItems.Count > 0 Then
        If Me.lvwSurtir2.SelectedItem.Selected Then
            If Val(txtCant.Text) < Val(lvwSurtir2.SelectedItem.SubItems(4)) Then
                cant = Val(lvwSurtir2.SelectedItem.SubItems(4)) - Val(txtCant.Text)
                Estado = "A"
            Else
                cant = 0
                Estado = "I"
            End If
            sBuscar = "UPDATE SURTIDOS SET ESTADO_ACTUAL = '" & Me.lvwSurtir2.SelectedItem & "', CANTIDAD = " & cant & " WHERE SURTIDO = '" & Estado & "'"
            cnn.Execute (sBuscar)
        End If
    End If
    Me.Llenar_Lista_Surtidos_2
    Label3.Caption = "--------------------------------------------------------------------"
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Imprimir1()
    Dim NRegistros As Integer
    Dim POSY As Integer
    Dim Con As Integer
    Printer.Print "    " & VarMen.Text5(0).Text
    Printer.Print "         SURTIR SUCURSAL"
    Printer.Print "        LISTA DE PRODUCTOS "
    Printer.Print "FECHA : " & Format(Date, "dd/mm/yyyy")
    NRegistros = lvwSurtir2.ListItems.Count
    POSY = 1400
    Printer.CurrentY = POSY
    Printer.CurrentX = 100
    Printer.Print "Producto"
    Printer.CurrentY = POSY
    Printer.CurrentX = 1900
    Printer.Print "Cant."
    Printer.CurrentY = POSY
    Printer.CurrentX = 2400
    Printer.Print "SUCURSAL"
    For Con = 1 To NRegistros
        If POSY > 16000 Then
            Printer.NewPage
            Printer.Print "    " & VarMen.Text5(0).Text
            Printer.Print "         SURTIR SUCURSAL"
            Printer.Print "        LISTA DE PRODUCTOS "
            Printer.Print "FECHA : " & Format(Date, "dd/mm/yyyy")
            NRegistros = lvwSurtir.ListItems.Count
            POSY = 1400
            Printer.CurrentY = POSY
            Printer.CurrentX = 100
            Printer.Print "Producto"
            Printer.CurrentY = POSY
            Printer.CurrentX = 1900
            Printer.Print "Cant."
            Printer.CurrentY = POSY
            Printer.CurrentX = 2400
            Printer.Print "SUCURSAL"
            POSY = POSY + 200
        Else
            POSY = POSY + 200
        End If
        Printer.CurrentY = POSY
        Printer.CurrentX = 100
        Printer.Print lvwSurtir2.ListItems(Con).SubItems(3)
        Printer.CurrentY = POSY
        Printer.CurrentX = 1900
        Printer.Print lvwSurtir2.ListItems(Con).SubItems(4)
        Printer.CurrentY = POSY
        Printer.CurrentX = 2400
        Printer.Print lvwSurtir2.ListItems(Con).SubItems(1)
        POSY = POSY + 200
        Printer.CurrentY = POSY
        Printer.CurrentX = 2400
        Printer.Print "---------------------------------------------------------"
    Next Con
    Printer.Print ""
    Printer.Print "FIN DEL LISTADO"
    Printer.EndDoc
End Sub
Private Sub Imprimir2()
    Dim Con As Integer
    Dim NRegistros As Integer
    Dim POSY As Integer
    If lvwSurtir.ListItems.Count > 0 Then
        Printer.Print "   " & VarMen.Text5(0).Text
        Printer.Print "         SURTIR SUCURSAL"
        Printer.Print "        LISTA DE PRODUCTOS "
        Printer.Print "FECHA : " & Format(Date, "dd/mm/yyyy")
        NRegistros = lvwSurtir.ListItems.Count
        POSY = 1400
        Printer.CurrentY = POSY
        Printer.CurrentX = 100
        Printer.Print "Producto"
        Printer.CurrentY = POSY
        Printer.CurrentX = 1900
        Printer.Print "Cant."
        Printer.CurrentY = POSY
        Printer.CurrentX = 2400
        Printer.Print "SUCURSAL"
        For Con = 1 To NRegistros
            If lvwSurtir.ListItems(Con).Checked Then
                If POSY > 16000 Then
                    Printer.NewPage
                    Printer.Print "    " & VarMen.Text5(0).Text
                    Printer.Print "         SURTIR SUCURSAL"
                    Printer.Print "        LISTA DE PRODUCTOS "
                    Printer.Print "FECHA : " & Format(Date, "dd/mm/yyyy")
                    NRegistros = lvwSurtir.ListItems.Count
                    POSY = 1400
                    Printer.CurrentY = POSY
                    Printer.CurrentX = 100
                    Printer.Print "Producto"
                    Printer.CurrentY = POSY
                    Printer.CurrentX = 1900
                    Printer.Print "Cant."
                    Printer.CurrentY = POSY
                    Printer.CurrentX = 2400
                    Printer.Print "SUCURSAL"
                    POSY = POSY + 200
                Else
                    POSY = POSY + 200
                End If
                Printer.CurrentY = POSY
                Printer.CurrentX = 100
                Printer.Print lvwSurtir.ListItems(Con).SubItems(5)
                Printer.CurrentY = POSY
                Printer.CurrentX = 1900
                Printer.Print lvwSurtir.ListItems(Con).SubItems(6)
                Printer.CurrentY = POSY
                Printer.CurrentX = 2400
                Printer.Print lvwSurtir.ListItems(Con).SubItems(1)
                POSY = POSY + 200
                Printer.CurrentY = POSY
                Printer.CurrentX = 2400
                Printer.Print "---------------------------------------------------------"
            End If
        Next Con
        Printer.Print ""
        Printer.Print "FIN DEL LISTADO"
        Printer.EndDoc
    End If
End Sub
Private Sub Command3_Click()
    Dim ID As Integer
    Dim ID_PEDIDO As Integer
    Dim Con As Integer
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim cant As Integer
    Dim Pend As Integer
    Dim Edo As String
    If lvwSurtir.ListItems.Count > 0 Then
        For Con = 1 To lvwSurtir.ListItems.Count
            If lvwSurtir.ListItems(Con).Checked Then
                Pend = 0
                Edo = "S"
                ID_PEDIDO = Me.lvwSurtir.ListItems.Item(Con)
                ID = lvwSurtir.ListItems.Item(Con).SubItems(4)
                sBuscar = "SELECT * FROM EXISTENCIAS WHERE SUCURSAL = 'BODEGA' AND ID_PRODUCTO = '" & lvwSurtir.ListItems.Item(Con).SubItems(5) & "'"
                Set tRs = cnn.Execute(sBuscar)
                If tRs.EOF And tRs.BOF Then
                    MsgBox "NO HAY EXISTENCIAS PARA SURTIR EL PEDIDO", vbInformation, "SACC"
                Else
                    Do
                        cant = Val(InputBox("INTRODUSCA LA CANTIDAD DE " & lvwSurtir.ListItems.Item(Con).SubItems(5) & " QUE SE VA A SURTIR", "SACC", Val(lvwSurtir.ListItems.Item(Con).SubItems(6))))
                    Loop Until cant <= Val(lvwSurtir.ListItems.Item(Con).SubItems(6)) And cant > 0
                    
                    If Val(tRs.Fields("CANTIDAD")) < cant Then
                        If MsgBox("NO TIENE SUFICIENTE EXISTENCIA PARA SURTIR LA CANTIDAD PEDIDA" & Chr(13) & _
                               "                             DESEA SURTIR " & tRs.Fields("CANTIDAD"), vbYesNo, "SACC") = vbYes Then
                            cant = Val(tRs.Fields("CANTIDAD"))
                        Else
                            cant = 0
                        End If
                    Else
                        If cant < Val(lvwSurtir.ListItems.Item(Con).SubItems(6)) Then
                            Pend = Val(lvwSurtir.ListItems.Item(Con).SubItems(6)) - cant
                            Edo = "R"
                        End If
                        If cant > 0 Then
                            sBuscar = "SELECT CANTIDAD FROM EXISTENCIAS WHERE SUCURSAL = '" & lvwSurtir.ListItems.Item(Con).SubItems(1) & "' AND ID_PRODUCTO = '" & lvwSurtir.ListItems.Item(Con).SubItems(5) & "'"
                            Set tRs = cnn.Execute(sBuscar)
                            If tRs.EOF And tRs.BOF Then
                                sBuscar = "INSERT INTO EXISTENCIAS (ID_PRODUCTO, CANTIDAD, SUCURSAL) VALUES ('" & lvwSurtir.ListItems.Item(Con).SubItems(5) & "', " & cant & ", '" & lvwSurtir.ListItems.Item(Con).SubItems(1) & "');"
                            Else
                                sBuscar = "UPDATE EXISTENCIAS SET CANTIDAD = " & CDbl(tRs.Fields("CANTIDAD")) + CDbl(cant) & " WHERE SUCURSAL = '" & lvwSurtir.ListItems.Item(Con).SubItems(1) & "' AND ID_PRODUCTO = '" & lvwSurtir.ListItems.Item(Con).SubItems(5) & "'"
                            End If
                            cnn.Execute (sBuscar)
                            sBuscar = "UPDATE DETALLE_PEDIDO SET ENTREGADO = '" & Edo & "', CANTIDAD = " & Me.txtCant.Text & " WHERE ID = " & ID & " AND ID_PEDIDO = " & ID_PEDIDO
                            cnn.Execute (sBuscar)
                            sBuscar = "UPDATE EXISTENCIAS SET CANTIDAD = CANTIDAD - " & cant & " WHERE SUCURSAL = 'BODEGA' AND ID_PRODUCTO = '" & lvwSurtir.ListItems.Item(Con).SubItems(5) & "'"
                            cnn.Execute (sBuscar)
                        End If
                    End If
                End If
            End If
        Next Con
    End If
    Me.Llenar_Lista_Surtidos "R"
End Sub
Private Sub Command4_Click()
    Dim Con As Integer
    If lvwSurtir.ListItems.Count > 0 Then
        For Con = 1 To lvwSurtir.ListItems.Count
            lvwSurtir.ListItems(Con).Checked = True
        Next Con
    End If
End Sub
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    Set cnn = New ADODB.Connection
    With cnn
        .ConnectionString = _
            "Provider=" & NvoMen.TxtProvider.Text & ";Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & NvoMen.TxtBaseDatos.Text & ";Data Source=" & NvoMen.txtServidor.Text & ";"
        .Open
    End With
    With lvwSurtir
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "ID_PEDIDO", 300
        .ColumnHeaders.Add , , "SUCURSAL", 1440
        .ColumnHeaders.Add , , "AGENTE", 1440
        .ColumnHeaders.Add , , "FECHA", 1440
        .ColumnHeaders.Add , , "ID", 0
        .ColumnHeaders.Add , , "CLAVE", 2160
        .ColumnHeaders.Add , , "CANTIDAD", 720
        .ColumnHeaders.Add , , "Descripcion", 2880
    End With
    With lvwSurtir2
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "SURTIDO", 0
        .ColumnHeaders.Add , , "SUCURSAL", 2880
        .ColumnHeaders.Add , , "FECHA", 2880
        .ColumnHeaders.Add , , "CLAVE", 2880
        .ColumnHeaders.Add , , "CANTIDAD", 1440
    End With
    Llenar_Lista_Surtidos "R"
    Llenar_Lista_Surtidos_2
End Sub
Sub Llenar_Lista_Surtidos_2()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Me.lvwSurtir2.ListItems.Clear
    sBuscar = "SELECT SURTIDO, SUCURSAL, FECHA, ID_PRODUCTO, CANTIDAD From SURTIDOS WHERE ESTADO_ACTUAL='A'"
    Set tRs = cnn.Execute(sBuscar)
    Do While Not tRs.EOF
        Set ItMx = Me.lvwSurtir2.ListItems.Add(, , tRs.Fields("Surtido"))
        If Not IsNull(tRs.Fields("Sucursal")) Then ItMx.SubItems(1) = Trim(tRs.Fields("Sucursal"))
        If Not IsNull(tRs.Fields("fecha")) Then ItMx.SubItems(2) = tRs.Fields("fecha")
        If Not IsNull(tRs.Fields("ID_PRODUCTO")) Then ItMx.SubItems(3) = tRs.Fields("ID_PRODUCTO")
        If Not IsNull(tRs.Fields("CANTIDAD")) Then ItMx.SubItems(4) = tRs.Fields("CANTIDAD")
        tRs.MoveNext
    Loop
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Image1_Click()
    If bLvw = 1 Then
        Imprimir2
    ElseIf bLvw = 2 Then
        Imprimir1
    End If
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub lvwSurtir_Click()
    bLvw = 1
End Sub
Private Sub lvwSurtir2_Click()
    bLvw = 2
End Sub
Private Sub lvwSurtir2_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Label3.Caption = Item.SubItems(3)
    txtCant.Text = Item.SubItems(4)
End Sub
Private Sub txtCant_GotFocus()
    txtCant.BackColor = &HFFE1E1
End Sub
Private Sub txtCant_LostFocus()
    txtCant.BackColor = &H80000005
End Sub
