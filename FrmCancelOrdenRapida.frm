VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmCancelOrdenRapida 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cancelar Ordne Rapida"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7935
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   7935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   6840
      TabIndex        =   3
      Top             =   840
      Width           =   975
      Begin VB.Label Label1 
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
         TabIndex        =   4
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image6 
         Height          =   705
         Left            =   120
         MouseIcon       =   "FrmCancelOrdenRapida.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "FrmCancelOrdenRapida.frx":030A
         Top             =   240
         Width           =   705
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   6840
      TabIndex        =   1
      Top             =   2040
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
         TabIndex        =   2
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "FrmCancelOrdenRapida.frx":1DBC
         MousePointer    =   99  'Custom
         Picture         =   "FrmCancelOrdenRapida.frx":20C6
         Top             =   120
         Width           =   720
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   5530
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Cancelar Orden Rapida"
      TabPicture(0)   =   "FrmCancelOrdenRapida.frx":41A8
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Text1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Text2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Text3"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   2160
         TabIndex        =   8
         Top             =   2040
         Width           =   3615
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   5400
         TabIndex        =   7
         Top             =   120
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   3240
         TabIndex        =   6
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "Comentario :"
         Height          =   255
         Left            =   1080
         TabIndex        =   9
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Numero de Orden Rapida :"
         Height          =   255
         Left            =   1200
         TabIndex        =   5
         Top             =   1320
         Width           =   2055
      End
   End
End
Attribute VB_Name = "FrmCancelOrdenRapida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Dim ID As Integer
Private Sub Form_Load()
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    Set cnn = New ADODB.Connection
    With cnn
        .ConnectionString = _
            "Provider=" & NvoMen.TxtProvider.Text & ";Password=" & NvoMen.TxtContrasena.Text & ";Persist Security Info=True;User ID=" & NvoMen.TxtUsuario.Text & ";Initial Catalog=" & NvoMen.TxtBaseDatos.Text & ";Data Source=" & NvoMen.txtServidor.Text & ";"
        .Open
    End With
End Sub
Private Sub Image6_Click()
    Label3.Visible = True
    Text3.Visible = True
    If (Text3.Text <> "") And (Text1.Text <> "") Then
        If MsgBox("ESTA SEGURO QUE DESEA CANCELAR LA ORDEN RAPIDA NO. " & Text1.Text & "?", vbYesNo + vbCritical + vbDefaultButton1, "SACC") = vbYes Then
            Dim sBuscar As String
            Dim tRs As ADODB.Recordset
            sBuscar = "SELECT ID_ORDEN_RAPIDA, ESTADO FROM ORDEN_RAPIDA WHERE ID_ORDEN_RAPIDA = " & Text1.Text
            Set tRs = cnn.Execute(sBuscar)
            If Not (tRs.EOF And tRs.BOF) Then
                If tRs.Fields("ESTADO") = "A" Or tRs.Fields("ESTADO") = "M" Then
                    sBuscar = "INSERT INTO ORDEN_CANCE (ORDEN,TIPO,COMENTARIO,FECHA,ID_USUARIO) VALUES (" & Text1.Text & ", 'R', '" & Text3.Text & "','" & Format(Date, "dd/mm/yyyy") & "','" & VarMen.Text1(0) & "');"
                    cnn.Execute (sBuscar)
                    sBuscar = "UPDATE ORDEN_RAPIDA SET ESTADO = 'C' WHERE ID_ORDEN_RAPIDA = " & Text1.Text
                    cnn.Execute (sBuscar)
                    MsgBox "LA ORDEN DE COMPRA HA SIDO CANCELADA", vbInformation, "SACC"
                    ID = Text1.Text
                    'CANCE
                    Text1.Text = ""
                    Text3.Text = ""
                Else
                    If tRs.Fields("ESTADO") = "F" Then
                        If MsgBox("La orden ya esta pagada, ¿Desea cancelarla?", vbYesNo, "SACC") = vbYes Then
                            sBuscar = "INSERT INTO ORDEN_CANCE (ORDEN,TIPO,COMENTARIO,FECHA,ID_USUARIO) VALUES (" & Text1.Text & ", 'R', '" & Text3.Text & "','" & Format(Date, "dd/mm/yyyy") & "','" & VarMen.Text1(0) & "');"
                            cnn.Execute (sBuscar)
                            sBuscar = "UPDATE ORDEN_RAPIDA SET ESTADO = 'C' WHERE ID_ORDEN_RAPIDA = " & Text1.Text
                            cnn.Execute (sBuscar)
                            MsgBox "LA ORDEN DE COMPRA HA SIDO CANCELADA", vbInformation, "SACC"
                            ID = Text1.Text
                            'CANCE
                            Text1.Text = ""
                            Text3.Text = ""
                        End If
                    Else
                        MsgBox "La Orden ya ha sido cancelada", vbExclamation, "SACC"
                    End If
                End If
            Else
                MsgBox "EL NUMERO DE ORDEN DE COMPRA NO EXISTE O YA FUE CANCELADA", vbInformation, "SACC"
            End If
        End If
    Else
        MsgBox "INGRESE MOTIVO DE CANCELACION,COMENTARIOS", vbInformation, "SACC"
    End If
End Sub
Private Sub CANCE()
    Dim oDoc  As cPDF
    Dim dblX  As Double
    Dim dblY  As Double
    Dim Angle As Double
    Dim Cont As Integer
    Dim Posi As Integer
    Dim loca As Integer
    Dim sBuscar As String
    Dim tRs1 As ADODB.Recordset
    Dim PosVer As Integer
    Dim Posabo As Integer
    Dim tRs  As ADODB.Recordset
    Set oDoc = New cPDF
    Dim sumdeuda As Double
    Dim sumIndi As Double
    Dim sNombre As String
    Dim sumabonos As Double
    Dim sumIndiabonos As Double
    Dim ConPag As Integer
    ConPag = 1
    If Not oDoc.PDFCreate(App.Path & "\reportecuentas.pdf") Then
        Exit Sub
    End If
    oDoc.Fonts.Add "F1", Courier, MacRomanEncoding
    oDoc.Fonts.Add "F2", Helvetica_Bold, MacRomanEncoding
    oDoc.Fonts.Add "F3", Helvetica, MacRomanEncoding
    oDoc.Fonts.Add "F4", Courier, MacRomanEncoding
    ' Encabezado del reporte
    oDoc.WTextBox 40, 200, 20, 250, VarMen.TxtEmp(0).Text, "F2", 7, hCenter
    oDoc.WTextBox 60, 200, 20, 250, VarMen.TxtEmp(1).Text & ", " & VarMen.TxtEmp(4).Text, "F2", 7, hCenter
    oDoc.WTextBox 70, 200, 20, 250, VarMen.TxtEmp(5).Text & " " & VarMen.TxtEmp(6).Text, "F2", 7, hCenter
    oDoc.WTextBox 80, 200, 20, 250, "Tel " & VarMen.TxtEmp(2).Text, "F2", 7, hCenter
    oDoc.WTextBox 90, 200, 20, 250, "Cancelaciones", "F2", 10, hCenter
    oDoc.WTextBox 90, 380, 20, 250, Format(Date, "dd/mm/yyyy"), "F3", 8, hCenter
    ' Encabezado de pagina
    oDoc.WTextBox 100, 20, 30, 40, "Folio", "F2", 10, hCenter
    oDoc.WTextBox 100, 60, 30, 80, "Fecha", "F2", 10, hCenter
    oDoc.WTextBox 100, 100, 50, 160, "Tipo", "F2", 10, hCenter
    oDoc.WTextBox 100, 100, 50, 160, "Comentario", "F2", 10, hCenter
    oDoc.WTextBox 100, 250, 50, 160, "Usuario", "F2", 10, hLeft
    ' Cuerpo del reporte
    sBuscar = "SELECT * FROM vsord_cance WHERE   TIPO ='R'  AND ORDEN= '" & ID & "'"
    'AND FECHA BETWEEN '" & DTPicker1.Value & "' AND '" & DTPicker2.Value & " ' ORDER BY NOMBRE"
    Set tRs = cnn.Execute(sBuscar)
    Posi = 140
    oDoc.SetLineFormat 0.5, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 10, 100
    oDoc.WLineTo 580, 100
    oDoc.LineStroke
    oDoc.MoveTo 10, 125
    oDoc.WLineTo 580, 125
    oDoc.LineStroke
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not (tRs.EOF)
            Posi = Posi + 15
            oDoc.WTextBox 100, 20, 30, 40, tRs.Fields("ORDEN"), "F2", 10, hCenter
            oDoc.WTextBox 100, 60, 30, 80, tRs.Fields("FECHA"), "F2", 10, hCenter
            oDoc.WTextBox 100, 100, 50, 160, tRs.Fields("TIPO"), "F2", 10, hCenter
            oDoc.WTextBox 100, 100, 50, 160, tRs.Fields("COMENTARIO"), "F2", 10, hCenter
            oDoc.WTextBox 100, 250, 50, 160, "NOMBRE", "F2", 10, hLeft
            tRs.MoveNext
            If Posi >= 760 Then
                oDoc.WTextBox 780, 500, 20, 175, ConPag, "F2", 7, hLeft
                ConPag = ConPag + 1
                oDoc.NewPage A4_Vertical
                ' Encabezado del reporte
                Posi = 140
                oDoc.WTextBox 40, 200, 20, 250, VarMen.TxtEmp(0).Text, "F2", 7, hCenter
                oDoc.WTextBox 60, 200, 20, 250, VarMen.TxtEmp(1).Text & ", " & VarMen.TxtEmp(4).Text, "F2", 7, hCenter
                oDoc.WTextBox 70, 200, 20, 250, VarMen.TxtEmp(5).Text & " " & VarMen.TxtEmp(6).Text, "F2", 7, hCenter
                oDoc.WTextBox 80, 200, 20, 250, "Tel " & VarMen.TxtEmp(2).Text, "F2", 7, hCenter
                oDoc.WTextBox 90, 200, 20, 250, "Cancelaciones", "F2", 10, hCenter
                oDoc.WTextBox 60, 380, 20, 250, "Rango del reporte", "F2", 10, hCenter
                oDoc.WTextBox 90, 380, 20, 250, Format(Date, "dd/mm/yyyy"), "F3", 8, hCenter
                ' Encabezado de pagina
                oDoc.WTextBox 100, 20, 30, 40, tRs.Fields("ORDEN"), "F2", 10, hCenter
                oDoc.WTextBox 100, 60, 30, 80, tRs.Fields("FECHA"), "F2", 10, hCenter
                oDoc.WTextBox 100, 100, 50, 160, tRs.Fields("TIPO"), "F2", 10, hCenter
                oDoc.WTextBox 100, 100, 50, 160, tRs.Fields("COMENTARIO"), "F2", 10, hCenter
                oDoc.WTextBox 100, 250, 50, 160, "NOMBRE", "F2", 10, hLeft
            End If
        Loop
        Cont = Cont + 1
    End If
    oDoc.PDFClose
    oDoc.Show
End Sub
Private Sub cancelar()
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    If (sBuscar = "SELECT ID_ORDEN_RAPIDA FROM ORDEN_RAPIDA WHERE ID_ORDEN_RAPIDA = " & Text1.Text) = True Then
        Set tRs = cnn.Execute(sBuscar)
    Else
        MsgBox "EL NUMERO DE ORDEN DE COMPRA NO EXISTE O YA FUE CANCELADA", vbInformation, "SACC"
    End If
End Sub
Private Sub Image9_Click()
    Unload Me
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    Valido = "1234567890"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
