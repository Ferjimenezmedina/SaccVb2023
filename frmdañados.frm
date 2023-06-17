VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmdañados 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Modificar Status en Comandas"
   ClientHeight    =   6255
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9015
   LinkTopic       =   "Form1"
   ScaleHeight     =   6255
   ScaleWidth      =   9015
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTraer 
      Caption         =   "Traer"
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
      Left            =   480
      Picture         =   "frmdañados.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox texbus 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   240
      TabIndex        =   22
      Top             =   480
      Width           =   1935
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   7920
      TabIndex        =   4
      Top             =   4920
      Width           =   975
      Begin VB.Image frmdañados 
         Height          =   870
         Left            =   120
         MouseIcon       =   "frmdañados.frx":29D2
         MousePointer    =   99  'Custom
         Picture         =   "frmdañados.frx":2CDC
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
         TabIndex        =   5
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   975
      Left            =   4560
      Picture         =   "frmdañados.frx":4DBE
      ScaleHeight     =   915
      ScaleWidth      =   3195
      TabIndex        =   3
      Top             =   0
      Width           =   3255
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   5400
      TabIndex        =   0
      Top             =   1080
      Visible         =   0   'False
      Width           =   1695
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   255
      Left            =   2760
      TabIndex        =   1
      Top             =   720
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   450
      _Version        =   393216
      Format          =   56229889
      CurrentDate     =   39658
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   255
      Left            =   2760
      TabIndex        =   2
      Top             =   360
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   450
      _Version        =   393216
      Format          =   56229889
      CurrentDate     =   39675
      MinDate         =   39418
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7200
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4335
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   7646
      _Version        =   393216
      Tabs            =   7
      TabsPerRow      =   7
      TabHeight       =   520
      BackColor       =   -2147483639
      TabCaption(0)   =   "Consulta"
      TabPicture(0)   =   "frmdañados.frx":5FB1
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "ListView1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "ListView2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Movimientos"
      TabPicture(1)   =   "frmdañados.frx":5FCD
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label6"
      Tab(1).Control(1)=   "Label5"
      Tab(1).Control(2)=   "Label4"
      Tab(1).Control(3)=   "Label3"
      Tab(1).Control(4)=   "Label1"
      Tab(1).Control(5)=   "Option1"
      Tab(1).Control(6)=   "Text4"
      Tab(1).Control(7)=   "Text3"
      Tab(1).Control(8)=   "Check3"
      Tab(1).Control(9)=   "Check2"
      Tab(1).Control(10)=   "Check1"
      Tab(1).Control(11)=   "Text2"
      Tab(1).Control(12)=   "Text1"
      Tab(1).Control(13)=   "Command1"
      Tab(1).ControlCount=   14
      TabCaption(2)   =   "Nuevas"
      TabPicture(2)   =   "frmdañados.frx":5FE9
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "ListView3"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Proceso"
      TabPicture(3)   =   "frmdañados.frx":6005
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "ListView4"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Calidad"
      TabPicture(4)   =   "frmdañados.frx":6021
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "ListView5"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Dañados"
      TabPicture(5)   =   "frmdañados.frx":603D
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "ListView6"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "Aut Rema"
      TabPicture(6)   =   "frmdañados.frx":6059
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "ListView7"
      Tab(6).ControlCount=   1
      Begin MSComctlLib.ListView ListView7 
         Height          =   3255
         Left            =   -74880
         TabIndex        =   30
         Top             =   720
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   5741
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListView6 
         Height          =   3255
         Left            =   -74880
         TabIndex        =   29
         Top             =   720
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   5741
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListView5 
         Height          =   3255
         Left            =   -74880
         TabIndex        =   28
         Top             =   720
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   5741
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListView4 
         Height          =   3135
         Left            =   -74880
         TabIndex        =   27
         Top             =   720
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   5530
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListView3 
         Height          =   2895
         Left            =   -74760
         TabIndex        =   26
         Top             =   600
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   5106
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   1335
         Left            =   240
         TabIndex        =   24
         Top             =   2940
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   2355
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Guardar"
         Enabled         =   0   'False
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
         Left            =   -68640
         Picture         =   "frmdañados.frx":6075
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   3900
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   375
         Left            =   -73080
         TabIndex        =   14
         Top             =   1260
         Width           =   2055
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   -73080
         TabIndex        =   13
         Top             =   1740
         Width           =   2055
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Comanda"
         Height          =   375
         Left            =   -69960
         TabIndex        =   12
         Top             =   1380
         Width           =   1455
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Produccion"
         Height          =   375
         Left            =   -69960
         TabIndex        =   11
         Top             =   1980
         Width           =   1335
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Sucursal"
         Height          =   375
         Left            =   -69960
         TabIndex        =   10
         Top             =   2700
         Width           =   1575
      End
      Begin VB.TextBox Text3 
         Height          =   405
         Left            =   -73080
         TabIndex        =   9
         Top             =   2220
         Width           =   2055
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   -73080
         TabIndex        =   8
         Top             =   2700
         Width           =   2055
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Descontar Inventarios"
         Height          =   375
         Left            =   -69720
         TabIndex        =   7
         Top             =   780
         Width           =   2055
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1695
         Left            =   240
         TabIndex        =   16
         Top             =   900
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   2990
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label Label1 
         Caption         =   " ID_PRODUCTO :"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   21
         Top             =   1380
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "CANTIDAD :"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   20
         Top             =   1860
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "SOLICITO :"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   19
         Top             =   2340
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "COMENTARIO:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   18
         Top             =   2820
         Width           =   1815
      End
      Begin VB.Label Label6 
         Caption         =   "ID_PRODUCTO :"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   17
         Top             =   1380
         Width           =   1815
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Comanda"
      Height          =   255
      Left            =   240
      TabIndex        =   25
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmdañados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnn As ADODB.Connection
Dim sqlQuery As String
Dim tLi As ListItem
Dim lvSI As ListSubItem
Dim tRs As Recordset
Dim tRs2 As Recordset
Dim intIndex As Integer
Dim StrRep As String
Dim StrRep1 As String
Dim bBandExis As Boolean



Private Sub Check1_Click()
Command1.Enabled = True
Check2.Value = 0
Check3.Value = 0
End Sub

Private Sub Check2_Click()
Command1.Enabled = True
Check1.Value = 0
Check3.Value = 0
End Sub

Private Sub Check3_Click()
Command1.Enabled = True
Check1.Value = 0
Check2.Value = 0
End Sub

Private Sub Imprimir()
Dim Path As String
Dim busc As String
    Dim SelectionFormula As Date
     Path = App.Path
     If Option2 = True Then
         Set crReport = crApplication.OpenReport(Path & "\REPORTES\repcartvacios.rpt")
        crReport.Database.LogOnServer "p2ssql.dll", GetSetting("APTONER", "ConfigSACC", "SERVIDOR", "1"), GetSetting("APTONER", "ConfigSACC", "DATABASE", "APTONER"), GetSetting("APTONER", "ConfigSACC", "USUARIO", "1"), GetSetting("APTONER", "ConfigSACC", "PASSWORD", "1")
        crReport.SQLQueryString = StrRep1
        crReport.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
        frmRep.Show vbModal, Me
        
        
    
        Else
         Set crReport = crApplication.OpenReport(Path & "\REPORTES\repcartvacios.rpt")
        crReport.Database.LogOnServer "p2ssql.dll", GetSetting("APTONER", "ConfigSACC", "SERVIDOR", "1"), GetSetting("APTONER", "ConfigSACC", "DATABASE", "APTONER"), GetSetting("APTONER", "ConfigSACC", "USUARIO", "1"), GetSetting("APTONER", "ConfigSACC", "PASSWORD", "1")
        crReport.SQLQueryString = StrRep
        crReport.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
        frmRep.Show vbModal, Me
            End If
        
        
        
End Sub

Private Sub cmdImprimir_Click()
Dim sBuscar As String
    Dim tRs As Recordset
    Dim tLi As ListItem
    Dim sumab As Double
    Dim sumde As Double
   Dim parcial As String
    ListView1.ListItems.Clear
    parcial = InputBox("Ingresar  el ID_PRODUCTO: ")
     'temporal_abonos
     sBuscar = "SELECT * FROM CARTVAC WHERE ID_PRODUCTO LIKE '%" & parcial & "%'AND FECHA BETWEEN '" & DTPicker2.Value & "' AND '" & DTPicker3.Value & " '   ORDER BY ID_PRODUCTO "
    ' sBuscar = "SELECT ID_VENTA,FOLIO,FECHA,NOMBRE,NO_CHEQUE,BANCO,SUM(CANT_ABONO)AS CREDITO_DISPO,LIMITE_CREDITO, SUM(DEUDA)AS DEUDA_ACTUAL FROM VsRepAbonos WHERE ID_CLIENTE = " & Text1.Text & "%' AND FECHA BETWEEN'" & DTPicker1.Value & " ' AND '" & DTPicker2.Value & "' ORDER BY FECHA"
     cnn.Execute (sBuscar)
     StrRep1 = sBuscar
     Imprimir
      
End Sub

Private Sub cmdTraer_Click()
 Dim sBuscar As String
    Dim tRs As Recordset
    Dim tLi As ListItem
    ListView1.ListItems.Clear
    sBuscar = "SELECT * FROM VSREVCOMANDA123 WHERE ID_COMANDA like  '%" & texbus.Text & "%' AND FECHA BETWEEN '" & DTPicker2.Value & "' AND '" & DTPicker3.Value & " ' ORDER BY FECHA ASC"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView1.ListItems.Add(, , tRs.Fields("ID_COMANDA"))
            If Not IsNull(tRs.Fields("NOMBRE")) Then tLi.SubItems(1) = tRs.Fields("NOMBRE")
        
        
        
        
        '    If Not IsNull(tRs.Fields("ESTADO_ACTUAL")) Then
         '       If tRs.Fields("ESTADO_ACTUAL") = "A" Then
          '          tLi.SubItems(2) = "Nueva"
           '     Else
            '        If tRs.Fields("ESTADO_ACTUAL") = "R" Or tRs.Fields("ESTADO_ACTUAL") = "S" Then
             '           tLi.SubItems(2) = "En Producción"
              '      Else
                '        If tRs.Fields("ESTADO_ACTUAL") = "P" Then
               '             tLi.SubItems(2) = "Probando en Calidad"
                 '       Else
                   '         If tRs.Fields("ESTADO_ACTUAL") = "N" Or tRs.Fields("ESTADO_ACTUAL") = "M" Then
                  '              tLi.SubItems(2) = "Cartuchos Dañados"
                    '        Else
                      '          If tRs.Fields("ESTADO_ACTUAL") = "L" Then
                     '               tLi.SubItems(2) = "Terminado"
                       '         Else
                        ''            If tRs.Fields("ESTADO_ACTUAL") = "Z" Then
                          '              tLi.SubItems(2) = "Aprovar Rema"
                           '         Else
                             '           If tRs.Fields("ESTADO_ACTUAL") = "C" Or tRs.Fields("ESTADO_ACTUAL") = "0" Then
                            '                tLi.SubItems(2) = "CANCELADA"
                              '          End If
                               '     End If
                    '            End If
                      '      End If
                     '   End If
                    'End If
                'End If
            'End If
            'If Not IsNull(tRs.Fields("CANTIDAD")) Then tLi.SubItems(3) = tRs.Fields("CANTIDAD")
            'If Not IsNull(tRs.Fields("CANT_FUNCIONO")) Then tLi.SubItems(4) = tRs.Fields("CANT_FUNCIONO")
            tRs.MoveNext
        Loop
    End If
    
    
    'Buscar = "SELECT * FROM VSREVCOMANDA123 WHERE ID_COMANDA like  '%" & texbus.Text & "%' AND ESTADO_ACTUAL= AND FECHA BETWEEN '" & DTPicker2.Value & "' AND '" & DTPicker3.Value & " ' ORDER BY FECHA ASC"
    'Set tRs = cnn.Execute(sBuscar)
    'If Not (tRs.EOF And tRs.BOF) Then
     '   Do While Not tRs.EOF
      '      Set tLi = ListView3.ListItems.Add(, , tRs.Fields("ID_COMANDA"))
       '     If Not IsNull(tRs.Fields("NOMBRE")) Then tLi.SubItems(1) = tRs.Fields("NOMBRE")
        '     tLi.SubItems(2) = "Nueva"
   '     Loop
    '    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "MENSAJE DEL SISTEMA"
    Err.Clear
End Sub

Private Sub Traer()
 Dim sBuscar As String
    Dim tRs As Recordset
    Dim tLi As ListItem
    
    fecha = Str(Date)
    DTPicker3.Value = Date
     DTPicker2.Value = Date
    sBuscar = "SELECT * FROM VSREVCOMANDA123 WHERE  ESTADO_ACTUAL='L'  AND FECHA BETWEEN '" & DTPicker2.Value & "' AND '" & DTPicker3.Value & " ' ORDER BY FECHA ASC"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView3.ListItems.Add(, , tRs.Fields("ID_COMANDA"))
            If Not IsNull(tRs.Fields("NOMBRE")) Then tLi.SubItems(1) = tRs.Fields("NOMBRE")
            If Not IsNull(tRs.Fields("ESTADO_ACTUAL")) Then tLi.SubItems(1) = tRs.Fields("ESTADO_ACTUAL")
        
        
        
        
        '    If Not IsNull(tRs.Fields("ESTADO_ACTUAL")) Then
         '       If tRs.Fields("ESTADO_ACTUAL") = "A" Then
          '          tLi.SubItems(2) = "Nueva"
           '     Else
            '        If tRs.Fields("ESTADO_ACTUAL") = "R" Or tRs.Fields("ESTADO_ACTUAL") = "S" Then
             '           tLi.SubItems(2) = "En Producción"
              '      Else
                '        If tRs.Fields("ESTADO_ACTUAL") = "P" Then
               '             tLi.SubItems(2) = "Probando en Calidad"
                 '       Else
                   '         If tRs.Fields("ESTADO_ACTUAL") = "N" Or tRs.Fields("ESTADO_ACTUAL") = "M" Then
                  '              tLi.SubItems(2) = "Cartuchos Dañados"
                    '        Else
                      '          If tRs.Fields("ESTADO_ACTUAL") = "L" Then
                     '               tLi.SubItems(2) = "Terminado"
                       '         Else
                        ''            If tRs.Fields("ESTADO_ACTUAL") = "Z" Then
                          '              tLi.SubItems(2) = "Aprovar Rema"
                           '         Else
                             '           If tRs.Fields("ESTADO_ACTUAL") = "C" Or tRs.Fields("ESTADO_ACTUAL") = "0" Then
                            '                tLi.SubItems(2) = "CANCELADA"
                              '          End If
                               '     End If
                    '            End If
                      '      End If
                     '   End If
                    'End If
                'End If
            'End If
            'If Not IsNull(tRs.Fields("CANTIDAD")) Then tLi.SubItems(3) = tRs.Fields("CANTIDAD")
            'If Not IsNull(tRs.Fields("CANT_FUNCIONO")) Then tLi.SubItems(4) = tRs.Fields("CANT_FUNCIONO")
            tRs.MoveNext
        Loop
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



Sub Reconexion()
On Error GoTo ManejaError
    Set cnn = New ADODB.Connection
    With cnn
        'If BanCnn = False Then
            '.Close
            'BanCnn = True
        'End If
        .ConnectionString = _
            "Provider=SQLOLEDB.1;Password=" & GetSetting("APTONER", "ConfigSACC", "PASSWORD", "LINUX") & ";Persist Security Info=True;User ID=" & GetSetting("APTONER", "ConfigSACC", "USUARIO", "LINUX") & ";Initial Catalog=" & GetSetting("APTONER", "ConfigSACC", "DATABASE", "APTONER") & ";" & _
            "Data Source=" & GetSetting("APTONER", "ConfigSACC", "SERVIDOR", "LINUX") & ";"
        .Open
        'BanCnn = False
        MsgBox "LA CONEXION SE RESABLECIO CON EXITO. PUEDE CONTINUAR CON SU TRABAJO.", vbInformation, "SACC"
    End With
Exit Sub
ManejaError:
    MsgBox Err.Number & Err.Description
    Err.Clear
    If MsgBox("NO PUDIMOS RESTABLECER LA CONEXIÓN, ¿DESEA REINTENTARLO?", vbYesNo + vbCritical + vbDefaultButton1, "SACC") = vbYes Then Reconexion
End Sub
Private Sub descontar()
Dim numcoma As Integer
Dim numcom As Integer
Dim numco As Integer
Dim bus As String
Dim tRs As Recordset



bus = "SELECT CANTIDAD FROM EXISTENCIAS WHERE ID_PRODUCTO = '" & Text1.Text & "'"
Set tRs = cnn.Execute(bus)
        If Not (tRs.EOF And tRs.BOF) Then
            bus = "UPDATE EXISTENCIAS SET CANTIDAD = " & CDbl(tRs.Fields("CANTIDAD")) - CDbl(Text2.Text) & " WHERE ID_PRODUCTO = '" & Text1.Text & "'"
            Set tRs = cnn.Execute(bus)
End If

End Sub
Private Sub Command1_Click()
Dim numcoma As Integer
Dim numcom As Integer
Dim numco As String
Dim bus As String
Dim tRs As Recordsets


End Sub

Private Sub Command2_Click()
listadodecomandas.Show vbModal
End Sub

Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
On Error GoTo ManejaError
   DTPicker2.Value = Date - 30
    DTPicker3.Value = Date
    Set cnn = New ADODB.Connection
    With cnn
        .ConnectionString = _
            "Provider=SQLOLEDB.1;Password=" & GetSetting("APTONER", "ConfigSACC", "PASSWORD", "LINUX") & ";Persist Security Info=True;User ID=" & GetSetting("APTONER", "ConfigSACC", "USUARIO", "LINUX") & ";Initial Catalog=" & GetSetting("APTONER", "ConfigSACC", "DATABASE", "APTONER") & ";" & _
            "Data Source=" & GetSetting("APTONER", "ConfigSACC", "SERVIDOR", "LINUX") & ";"
        .Open
    End With
    Dim fecha As Date

  Traer
    With ListView1
        .View = lvwReport
        .Gridlines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "ID", 500
        .ColumnHeaders.Add , , "ID_PRODUCTO", 1440
        .ColumnHeaders.Add , , "CANTIDAD", 1440
    End With
    
    With ListView2
        .View = lvwReport
        .Gridlines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "ID_PRODUCTO", 1200
        .ColumnHeaders.Add , , "DESCRIPCION", 1440
        .ColumnHeaders.Add , , "ESTADO_ACTUAL", 1440
        .ColumnHeaders.Add , , "CANTIDAD", 1000
        .ColumnHeaders.Add , , "FUNCIONARON", 1440
        .ColumnHeaders.Add , , "ID_COMANDA", 500
      
    End With
    
    
     With ListView3
        .View = lvwReport
        .Gridlines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "ID_PRODUCTO", 1200
        .ColumnHeaders.Add , , "DESCRIPCION", 1440
        .ColumnHeaders.Add , , "ESTADO_ACTUAL", 1440
        .ColumnHeaders.Add , , "CANTIDAD", 1000
        .ColumnHeaders.Add , , "FUNCIONARON", 1440
        .ColumnHeaders.Add , , "ID_COMANDA", 500
      
    End With
    
     With ListView4
        .View = lvwReport
        .Gridlines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "ID_PRODUCTO", 1200
        .ColumnHeaders.Add , , "DESCRIPCION", 1440
        .ColumnHeaders.Add , , "ESTADO_ACTUAL", 1440
        .ColumnHeaders.Add , , "CANTIDAD", 1000
        .ColumnHeaders.Add , , "FUNCIONARON", 1440
        .ColumnHeaders.Add , , "ID_COMANDA", 500
      
    End With
    
     With ListView5
        .View = lvwReport
        .Gridlines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "ID_PRODUCTO", 1200
        .ColumnHeaders.Add , , "DESCRIPCION", 1440
        .ColumnHeaders.Add , , "ESTADO_ACTUAL", 1440
        .ColumnHeaders.Add , , "CANTIDAD", 1000
        .ColumnHeaders.Add , , "FUNCIONARON", 1440
        .ColumnHeaders.Add , , "ID_COMANDA", 500
      
    End With
    
     With ListView6
        .View = lvwReport
        .Gridlines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "ID_PRODUCTO", 1200
        .ColumnHeaders.Add , , "DESCRIPCION", 1440
        .ColumnHeaders.Add , , "ESTADO_ACTUAL", 1440
        .ColumnHeaders.Add , , "CANTIDAD", 1000
        .ColumnHeaders.Add , , "FUNCIONARON", 1440
        .ColumnHeaders.Add , , "ID_COMANDA", 500
      
    End With
    
    With ListView7
        .View = lvwReport
        .Gridlines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "ID_PRODUCTO", 1200
        .ColumnHeaders.Add , , "DESCRIPCION", 1440
        .ColumnHeaders.Add , , "ESTADO_ACTUAL", 1440
        .ColumnHeaders.Add , , "CANTIDAD", 1000
        .ColumnHeaders.Add , , "FUNCIONARON", 1440
        .ColumnHeaders.Add , , "ID_COMANDA", 500
      
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

Private Sub Image14_Click()
    FrmReviAutRema.Show vbModal
End Sub



Private Sub Image9_Click()
    Unload Me
End Sub

Private Sub frmdañados_Click()
  Unload Me
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim sBuscar As String
    Dim tRs As Recordset
    Dim tLi As ListItem
    ListView2.ListItems.Clear
    sBuscar = "SELECT * FROM VsRevComanda WHERE ID_COMANDA = " & Item
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            Set tLi = ListView2.ListItems.Add(, , tRs.Fields("ID_PRODUCTO"))
            If Not IsNull(tRs.Fields("DESCRIPCION")) Then tLi.SubItems(1) = tRs.Fields("DESCRIPCION")
            If Not IsNull(tRs.Fields("ESTADO_ACTUAL")) Then
                If tRs.Fields("ESTADO_ACTUAL") = "A" Then
                    tLi.SubItems(2) = "Nueva"
                Else
                    If tRs.Fields("ESTADO_ACTUAL") = "R" Or tRs.Fields("ESTADO_ACTUAL") = "S" Then
                        tLi.SubItems(2) = "En Producción"
                    Else
                        If tRs.Fields("ESTADO_ACTUAL") = "P" Then
                            tLi.SubItems(2) = "Probando en Calidad"
                        Else
                            If tRs.Fields("ESTADO_ACTUAL") = "N" Or tRs.Fields("ESTADO_ACTUAL") = "M" Then
                                tLi.SubItems(2) = "Cartuchos Dañados"
                            Else
                                If tRs.Fields("ESTADO_ACTUAL") = "L" Then
                                    tLi.SubItems(2) = "Terminado"
                                Else
                                    If tRs.Fields("ESTADO_ACTUAL") = "Z" Then
                                        tLi.SubItems(2) = "Aprovar Rema"
                                    Else
                                        If tRs.Fields("ESTADO_ACTUAL") = "C" Or tRs.Fields("ESTADO_ACTUAL") = "0" Then
                                            tLi.SubItems(2) = "CANCELADA"
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
            If Not IsNull(tRs.Fields("CANTIDAD")) Then tLi.SubItems(3) = tRs.Fields("CANTIDAD")
            If Not IsNull(tRs.Fields("CANT_FUNCIONO")) Then tLi.SubItems(4) = tRs.Fields("CANT_FUNCIONO")
            If Not IsNull(tRs.Fields("ID_COMANDA")) Then tLi.SubItems(5) = tRs.Fields("ID_COMANDA")
            tRs.MoveNext
        Loop
    End If

On Error GoTo ManejaError
    
    
    
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









Private Sub Option2_Click()

DTPicker2.Visible = True
DTPicker3.Visible = True
End Sub

Private Sub texbus_KeyPress(KeyAscii As Integer)

On Error GoTo ManejaError

    If KeyAscii = 13 Then
       Me.cmdTraer.Value = True
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









