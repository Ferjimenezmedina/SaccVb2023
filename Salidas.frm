VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form Salidas 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Salida de Uso Interno de Inventario"
   ClientHeight    =   8415
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10170
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8415
   ScaleWidth      =   10170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame29 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9120
      TabIndex        =   37
      Top             =   5760
      Visible         =   0   'False
      Width           =   975
      Begin VB.Image Image2 
         Height          =   720
         Left            =   120
         MouseIcon       =   "Salidas.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "Salidas.frx":030A
         Top             =   240
         Width           =   720
      End
      Begin VB.Label Label31 
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
         TabIndex        =   38
         Top             =   960
         Width           =   975
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   9480
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   9120
      TabIndex        =   21
      Top             =   7080
      Width           =   975
      Begin VB.Image Image9 
         Height          =   870
         Left            =   120
         MouseIcon       =   "Salidas.frx":1E4C
         MousePointer    =   99  'Custom
         Picture         =   "Salidas.frx":2156
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
         TabIndex        =   22
         Top             =   960
         Width           =   975
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8175
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   14420
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   " Salidas"
      TabPicture(0)   =   "Salidas.frx":4238
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label6"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "ListView1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "ListView2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "DTPicker1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "CommonDialog1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Command6"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Command5"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Command2"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Text4"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Command4"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Option2"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Option1"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Text3"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Text2"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Command1"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Text1"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Combo1"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Command7"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Check1"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).ControlCount=   23
      TabCaption(1)   =   "Verficar Salidas"
      TabPicture(1)   =   "Salidas.frx":4254
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(1)=   "Frame2"
      Tab(1).Control(2)=   "Frame4"
      Tab(1).Control(3)=   "ListView3"
      Tab(1).ControlCount=   4
      Begin VB.CheckBox Check1 
         Caption         =   "Activo fijo"
         Height          =   255
         Left            =   6000
         TabIndex        =   39
         Top             =   4680
         Width           =   1215
      End
      Begin MSComctlLib.ListView ListView3 
         Height          =   4815
         Left            =   -74880
         TabIndex        =   36
         Top             =   2520
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   8493
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Frame Frame4 
         Caption         =   "Nombre"
         Height          =   1455
         Left            =   -72600
         TabIndex        =   32
         Top             =   600
         Width           =   3495
         Begin VB.CommandButton Command8 
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
            Left            =   2040
            Picture         =   "Salidas.frx":4270
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   840
            Width           =   1335
         End
         Begin VB.TextBox Text6 
            Height          =   375
            Left            =   120
            TabIndex        =   33
            Top             =   360
            Width           =   3255
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Filtro"
         Height          =   1455
         Left            =   -74880
         TabIndex        =   29
         Top             =   600
         Width           =   2175
         Begin VB.OptionButton Option5 
            Caption         =   "Sucursal"
            Height          =   375
            Left            =   240
            TabIndex        =   35
            Top             =   960
            Width           =   1215
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Producto"
            Height          =   255
            Left            =   240
            TabIndex        =   31
            Top             =   600
            Width           =   1215
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Fecha"
            Height          =   255
            Left            =   240
            TabIndex        =   30
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Rango Del Reporte"
         Height          =   1455
         Left            =   -68880
         TabIndex        =   24
         Top             =   600
         Width           =   2655
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   375
            Left            =   960
            TabIndex        =   25
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            Format          =   50921473
            CurrentDate     =   39885
         End
         Begin MSComCtl2.DTPicker DTPicker3 
            Height          =   375
            Left            =   960
            TabIndex        =   26
            Top             =   840
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            Format          =   50921473
            CurrentDate     =   39885
         End
         Begin VB.Label Label7 
            Caption         =   "AL :"
            Height          =   255
            Left            =   360
            TabIndex        =   28
            Top             =   960
            Width           =   375
         End
         Begin VB.Label Label3 
            Caption         =   "Del:"
            Height          =   255
            Left            =   360
            TabIndex        =   27
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Quitar Prod."
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
         Left            =   1680
         Picture         =   "Salidas.frx":6C42
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   7620
         Width           =   1335
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   960
         TabIndex        =   0
         Top             =   780
         Width           =   2655
      End
      Begin VB.TextBox Text1 
         Height          =   765
         Left            =   1320
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   4260
         Width           =   4455
      End
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
         Left            =   7440
         Picture         =   "Salidas.frx":9614
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   4620
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   3840
         TabIndex        =   15
         Top             =   780
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   840
         TabIndex        =   2
         Top             =   1380
         Width           =   5175
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Por Nombre"
         Height          =   195
         Left            =   6120
         TabIndex        =   4
         Top             =   1500
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Por Clave"
         Height          =   255
         Left            =   6120
         TabIndex        =   3
         Top             =   1260
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
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
         Left            =   7440
         Picture         =   "Salidas.frx":BFE6
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1260
         Width           =   1335
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   6960
         TabIndex        =   8
         Text            =   "1"
         Top             =   4260
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Imp. Seleccion"
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
         Left            =   5880
         Picture         =   "Salidas.frx":E9B8
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   7620
         Width           =   1455
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Guardar"
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
         Picture         =   "Salidas.frx":1138A
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   7620
         Width           =   1335
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Quitar Sel."
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
         Left            =   120
         Picture         =   "Salidas.frx":13D5C
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   7620
         Width           =   1335
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   8160
         Top             =   360
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         PrinterDefault  =   0   'False
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   6720
         TabIndex        =   1
         Top             =   780
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   50921473
         CurrentDate     =   38799
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   2415
         Left            =   120
         TabIndex        =   10
         Top             =   5100
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   4260
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
      Begin MSComctlLib.ListView ListView1 
         Height          =   2415
         Left            =   120
         TabIndex        =   6
         Top             =   1740
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   4260
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
         Caption         =   "Sucursal :"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   780
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Motivo de uso :"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   4260
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Buscar :"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1380
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Cantidad :"
         Height          =   255
         Left            =   6000
         TabIndex        =   17
         Top             =   4260
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Fecha :"
         Height          =   255
         Left            =   6000
         TabIndex        =   16
         Top             =   780
         Width           =   615
      End
   End
End
Attribute VB_Name = "Salidas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Dim SelItm As String
Dim SelImp1 As String
Dim SelImp2 As String
Dim SelImp3 As String
Dim SelImp4 As String
Dim SelImp5 As String
Dim StrRep As String
Dim sCosto As String
Dim ind As Integer
Private Sub Combo1_GotFocus()
    Combo1.BackColor = &HFFE1E1
End Sub
Private Sub Combo1_LostFocus()
    Combo1.BackColor = &H80000005
End Sub
Private Sub Command1_Click()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tRs2 As ADODB.Recordset
    Dim cant As Double
    Dim tLi As ListItem
    Dim ActFij As String
    If Check1.value = 1 Then
        ActFij = "Si"
    Else
        ActFij = "No"
    End If
    If SelItm <> "" And Text4.Text <> "" And Text1.Text <> "" And Text2.Text <> "" And Combo1.Text <> "" Then
        sBuscar = "SELECT CANTIDAD FROM EXISTENCIAS WHERE ID_PRODUCTO = '" & SelItm & "' AND SUCURSAL = '" & Combo1.Text & "' AND CANTIDAD >= " & Text4.Text
        Set tRs = cnn.Execute(sBuscar)
        If Not (tRs.BOF And tRs.EOF) Then
            cant = CDbl(tRs.Fields("CANTIDAD")) - CDbl(Text4.Text)
            sBuscar = "UPDATE EXISTENCIAS SET CANTIDAD = " & cant & " WHERE ID_PRODUCTO = '" & SelItm & "' AND SUCURSAL = '" & Combo1.Text & "'"
            cnn.Execute (sBuscar)
            sBuscar = "INSERT INTO SALIDAS (ID_PRODUCTO, CANTIDAD, JUSTIFICACION, ID_USUARIO, SUCURSAL, FECHA, PRECIO) VALUES ('" & SelItm & "', " & CDbl(Text4.Text) & ", '" & Text1.Text & "', '" & Text2.Text & "', '" & Combo1.Text & "', '" & Format(Date, "dd/mm/yyyy") & "', " & sCosto & ");"
            cnn.Execute (sBuscar)
            If ActFij = "Si" Then
                sBuscar = "INSERT INTO EXISTENCIA_FIJA (ID_PRODUCTO, CANTIDAD) VALUES ('" & SelItm & "', " & CDbl(Text4.Text) & ");"
                cnn.Execute (sBuscar)
            End If
            sBuscar = "SELECT ID_SALIDA FROM SALIDAS ORDER BY ID_SALIDA DESC"
            Set tRs2 = cnn.Execute(sBuscar)
            sBuscar = "SELECT * FROM SALIDAS WHERE ID_SALIDA ='" & tRs2.Fields("ID_SALIDA") & "'"
            Set tRs = cnn.Execute(sBuscar)
            With tRs
                If Not (.BOF And .EOF) Then
                    .MoveFirst
                    Do While Not .EOF
                        Set tLi = ListView2.ListItems.Add(, , .Fields("ID_PRODUCTO") & "")
                        tLi.SubItems(1) = .Fields("CANTIDAD") & ""
                        tLi.SubItems(2) = .Fields("JUSTIFICACION") & ""
                        tLi.SubItems(3) = .Fields("SUCURSAL") & ""
                        tLi.SubItems(4) = .Fields("FECHA") & ""
                        .MoveNext
                    Loop
                End If
            End With
        Else
            MsgBox "No Cuenta con la Existencia para surtir", vbExclamation, "SACC"
        End If
    Else
        MsgBox "Falta informacion necesaria para el registro", vbExclamation, "SACC"
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Command2_Click()
    Dim sqlQuery As String
    Dim tRs As ADODB.Recordset
    Dim tRs2 As ADODB.Recordset
On Error GoTo ManejaError
    sqlQuery = "SELECT ID_SALIDA FROM SALIDAS ORDER BY ID_SALIDA DESC"
    Set tRs = cnn.Execute(sqlQuery)
    CommonDialog1.Flags = 64
    CommonDialog1.CancelError = True
    CommonDialog1.ShowPrinter
    Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(0).Text)) / 2
    Printer.Print VarMen.Text5(0).Text
    Printer.CurrentX = (Printer.Width - Printer.TextWidth("R.F.C. " & VarMen.Text5(8).Text)) / 2
    Printer.Print "R.F.C. " & VarMen.Text5(8).Text
    Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text)) / 2
    Printer.Print VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text
    Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text)) / 2
    Printer.Print VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text & "                       SALIDA: " & tRs.Fields("ID_SALIDA")
    Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    Printer.Print "Clave del Prod.         Cantidad               Fecha                  Sucursal                                                                  Motivo"
    Printer.CurrentY = 1300
    Printer.CurrentX = 100
    Printer.Print SelImp1
    Printer.CurrentY = 1300
    Printer.CurrentX = 1800
    Printer.Print SelImp2
    Printer.CurrentY = 1300
    Printer.CurrentX = 2800
    Printer.Print SelImp5
    Printer.CurrentY = 1300
    Printer.CurrentX = 4000
    Printer.Print SelImp4
    Printer.CurrentY = 1300
    Printer.CurrentX = 5700
    Printer.Print SelImp3
    Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    Printer.Print ""
    Printer.Print ""
    Printer.Print " RECIBIDO ______________________________                                                 ENTREGO :______________________________"
    Printer.EndDoc
    Command2.Enabled = False
    CommonDialog1.Copies = 1
Exit Sub
ManejaError:
    If Err.Number <> 32755 Then
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    End If
    Err.Clear
End Sub
Private Sub Command4_Click()
    Buscas
End Sub
Private Sub Command5_Click()
On Error GoTo ManejaError
    Dim sqlQuery As String
    Dim tRs As ADODB.Recordset
    Dim fo As Integer
    sqlQuery = "SELECT ID_SALIDA FROM SALIDAS ORDER BY ID_SALIDA DESC"
    Set tRs = cnn.Execute(sqlQuery)
    fo = tRs.Fields("ID_SALIDA")
    CommonDialog1.Flags = 64
    CommonDialog1.CancelError = True
    CommonDialog1.ShowPrinter
    Dim pos As Integer
    Dim NumeroRegistros As Integer
    NumeroRegistros = ListView2.ListItems.COUNT
    pos = 1300
    Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(0).Text)) / 2
    Printer.Print VarMen.Text5(0).Text
    Printer.CurrentX = (Printer.Width - Printer.TextWidth("R.F.C. " & VarMen.Text5(8).Text)) / 2
    Printer.Print "R.F.C. " & VarMen.Text5(8).Text
    Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text)) / 2
    Printer.Print VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text
    Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text)) / 2
    Printer.Print VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text & "                       SALIDA: " & tRs.Fields("ID_SALIDA")
    Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    Printer.Print "Clave del Prod.         Cantidad                Fecha               Sucursal                                                        Motivo"
    Dim Conta As Integer
    For Conta = 1 To NumeroRegistros
        Printer.CurrentY = pos
        Printer.CurrentX = 100
        Printer.Print ListView2.ListItems(Conta)
        Printer.CurrentY = pos
        Printer.CurrentX = 1800
        Printer.Print ListView2.ListItems(Conta).SubItems(1)
        Printer.CurrentY = pos
        Printer.CurrentX = 2805
        Printer.Print ListView2.ListItems(Conta).SubItems(4)
        Printer.CurrentY = pos
        Printer.CurrentX = 4000
        Printer.Print ListView2.ListItems(Conta).SubItems(3)
        Printer.CurrentY = pos
        Printer.CurrentX = 5700
        Printer.Print ListView2.ListItems(Conta).SubItems(2)
         pos = pos + 200
        If pos >= 14200 Then
            Printer.NewPage
            pos = 1300
            Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
            Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(0).Text)) / 2
            Printer.Print VarMen.Text5(0).Text
            Printer.CurrentX = (Printer.Width - Printer.TextWidth("R.F.C. " & VarMen.Text5(8).Text)) / 2
            Printer.Print "R.F.C. " & VarMen.Text5(8).Text
            Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text)) / 2
            Printer.Print VarMen.Text5(1).Text & " COL. " & VarMen.Text5(4).Text
            Printer.CurrentX = (Printer.Width - Printer.TextWidth(VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text)) / 2
            Printer.Print VarMen.Text5(5).Text & ", " & VarMen.Text5(6).Text & " C.P. " & VarMen.Text5(9).Text
            Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
            Printer.Print "Clave del Prod.         Cantidad                 Fecha                  Sucursal                                                                  Motivo"
        End If
    Next Conta
    Printer.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    Printer.Print ""
    Printer.Print ""
    Printer.Print " RECIBIDO ______________________________                                                 ENTREGO :______________________________"
    Printer.EndDoc
    CommonDialog1.Copies = 1
    ListView2.ListItems.Clear
Exit Sub
ManejaError:
    If Err.Number <> 32755 Then
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    End If
    Err.Clear
End Sub
Private Sub Command6_Click()
On Error GoTo ManejaError
    If ind > 0 Then
        Dim sBuscar As String
        Dim tRs As ADODB.Recordset
        Dim cant As Double
        sBuscar = "SELECT CANTIDAD FROM EXISTENCIAS WHERE ID_PRODUCTO = '" & SelImp1 & "' AND SUCURSAL = '" & Combo1.Text & "'"
        Set tRs = cnn.Execute(sBuscar)
        If Not (tRs.BOF And tRs.EOF) Then
            cant = CDbl(tRs.Fields("CANTIDAD")) + CDbl(SelImp2)
            sBuscar = "UPDATE EXISTENCIAS SET CANTIDAD = " & cant & " WHERE ID_PRODUCTO = '" & SelImp1 & "' AND SUCURSAL = '" & Combo1.Text & "'"
            Set tRs = cnn.Execute(sBuscar)
            sBuscar = "DELETE FROM SALIDAS WHERE ID_PRODUCTO = '" & SelImp1 & "' AND FECHA = '" & DTPicker1.value & "'"
            cnn.Execute (sBuscar)
            ListView2.ListItems.Remove (ind)
            ind = 0
        Else
            MsgBox "OCURRIO UN ERROR, COMUNIQUESE CON EL ADMINISTRADOR DEL SISTEMA!", vbInformation, "SACC"
        End If
    Else
        MsgBox "No ha seleccionado ningun articulo", vbExclamation, "SACC"
    End If
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Command7_Click()
    ListView2.ListItems.Remove (ListView2.SelectedItem.Index)
End Sub
Private Sub Command8_Click()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    ListView3.ListItems.Clear
    Frame29.Visible = True
    If Option3.value = True Or Option4.value = True Or Option5.value = True Then
        ListView3.ListItems.Clear
        If Option3.value = True Then
            sBuscar = "SELECT * FROM SALIDAS WHERE ID_PRODUCTO LIKE '%" & Text6.Text & "%' AND FECHA BETWEEN '" & DTPicker2.value & "' AND '" & DTPicker3.value & " ' ORDER BY FECHA DESC"
        End If
        If Option4.value = True Then
            sBuscar = "SELECT * FROM SALIDAS WHERE ID_PRODUCTO LIKE '%" & Text6.Text & "%' AND FECHA BETWEEN '" & DTPicker2.value & "' AND '" & DTPicker3.value & " ' ORDER BY ID_PRODUCTO"
        End If
        If Option5.value = True Then
            sBuscar = "SELECT * FROM SALIDAS WHERE SUCURSAL LIKE '%" & Text6.Text & "%' AND FECHA BETWEEN '" & DTPicker2.value & "' AND '" & DTPicker3.value & " ' ORDER BY SUCURSAL"
        End If
        Set tRs = cnn.Execute(sBuscar)
        StrRep = sBuscar
        If Not (tRs.BOF And tRs.EOF) Then
            Do While Not tRs.EOF
                Set tLi = ListView3.ListItems.Add(, , tRs.Fields("ID_SALIDA") & "")
                If Not IsNull(tRs.Fields("ID_PRODUCTO")) Then tLi.SubItems(1) = tRs.Fields("ID_PRODUCTO") & ""
                If Not IsNull(tRs.Fields("CANTIDAD")) Then tLi.SubItems(2) = tRs.Fields("CANTIDAD") & ""
                If Not IsNull(tRs.Fields("FECHA")) Then tLi.SubItems(3) = tRs.Fields("FECHA") & ""
                If Not IsNull(tRs.Fields("JUSTIFICACION")) Then tLi.SubItems(4) = tRs.Fields("JUSTIFICACION") & ""
                tRs.MoveNext
            Loop
        Else
            MsgBox "NO EXISTE  NINGUNA  SALIDA", vbInformation, "SACC"
        End If
    Else
        MsgBox "SELECCIONE  UAN FORMA DE BUSQUEDA", vbInformation, "SACC"
    End If
    Text6.SetFocus
    Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Form_Load()
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
On Error GoTo ManejaError
    Command2.Enabled = False
    Text2.Text = VarMen.Text1(0).Text
    Option2.value = True
    DTPicker1.value = Format(Date, "dd/mm/yyyy")
    DTPicker1.Enabled = False
    DTPicker2.value = Format(Date - 30, "dd/mm/yyyy")
    DTPicker3.value = Format(Date, "dd/mm/yyyy")
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
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
        .ColumnHeaders.Add , , "Clave del Prodcto", 3000
        .ColumnHeaders.Add , , "Descripcion", 5600
        .ColumnHeaders.Add , , "Costo", 0
    End With
    With ListView3
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "SALIDA", 1500
        .ColumnHeaders.Add , , "PRODUCTO", 2500
        .ColumnHeaders.Add , , "CANTIDAD", 2000
        .ColumnHeaders.Add , , "FECHA", 2000
        .ColumnHeaders.Add , , "JUSTIFICACION", 4000
    End With
    With ListView2
        .View = lvwReport
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HotTracking = False
        .HoverSelection = False
        .ColumnHeaders.Add , , "Clave del Prodcto", 3000
        .ColumnHeaders.Add , , "Cantidad", 1600
        .ColumnHeaders.Add , , "Justificacion", 5600
        .ColumnHeaders.Add , , "Sucursal", 1600
        .ColumnHeaders.Add , , "Fecha", 1200
        .ColumnHeaders.Add , , "Activo fijo", 1200
    End With
    sBuscar = "SELECT NOMBRE FROM SUCURSALES WHERE ELIMINADO = 'N' ORDER BY NOMBRE"
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        If (.BOF And .EOF) Then
            MsgBox "No se han encontrado los datos buscados"
        Else
            ListView1.ListItems.Clear
            .MoveFirst
            Do While Not .EOF
                Combo1.AddItem .Fields("NOMBRE")
                .MoveNext
            Loop
        End If
    End With
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Buscas()
On Error GoTo ManejaError
    Dim sBuscar As String
    Dim tRs As ADODB.Recordset
    Dim tLi As ListItem
    ListView1.ListItems.Clear
    If Option1.value Then
        sBuscar = "SELECT ID_PRODUCTO, DESCRIPCION, PRECIO_COSTO FROM ALMACEN1 WHERE Descripcion LIKE '%" & Text3.Text & "%' ORDER BY Descripcion"
    Else
        sBuscar = "SELECT ID_PRODUCTO, DESCRIPCION, PRECIO_COSTO FROM ALMACEN1 WHERE ID_PRODUCTO LIKE '%" & Text3.Text & "%' ORDER BY ID_PRODUCTO"
    End If
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        If Not (.BOF And .EOF) Then
            .MoveFirst
            Do While Not .EOF
                Set tLi = ListView1.ListItems.Add(, , .Fields("ID_PRODUCTO"))
                tLi.SubItems(1) = .Fields("Descripcion")
                tLi.SubItems(2) = .Fields("PRECIO_COSTO")
                .MoveNext
            Loop
        End If
    End With
    If Option1.value Then
        sBuscar = "SELECT ID_PRODUCTO, DESCRIPCION, PRECIO_COSTO FROM ALMACEN2 WHERE Descripcion LIKE '%" & Text3.Text & "%' ORDER BY Descripcion"
    Else
        sBuscar = "SELECT ID_PRODUCTO, DESCRIPCION, PRECIO_COSTO FROM ALMACEN2 WHERE ID_PRODUCTO LIKE '%" & Text3.Text & "%' ORDER BY ID_PRODUCTO"
    End If
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        If Not (.BOF And .EOF) Then
            .MoveFirst
            Do While Not .EOF
                Set tLi = ListView1.ListItems.Add(, , .Fields("ID_PRODUCTO"))
                tLi.SubItems(1) = .Fields("Descripcion")
                tLi.SubItems(2) = .Fields("PRECIO_COSTO")
                .MoveNext
            Loop
        End If
    End With
    If Option1.value Then
        sBuscar = "SELECT ID_PRODUCTO, DESCRIPCION, PRECIO_COSTO FROM ALMACEN3 WHERE Descripcion LIKE '%" & Text3.Text & "%' ORDER BY Descripcion"
    Else
        sBuscar = "SELECT ID_PRODUCTO, DESCRIPCION, PRECIO_COSTO FROM ALMACEN3 WHERE ID_PRODUCTO LIKE '%" & Text3.Text & "%' ORDER BY ID_PRODUCTO"
    End If
    Set tRs = cnn.Execute(sBuscar)
    With tRs
        If Not (.BOF And .EOF) Then
            .MoveFirst
            Do While Not .EOF
                Set tLi = ListView1.ListItems.Add(, , .Fields("ID_PRODUCTO"))
                tLi.SubItems(1) = .Fields("Descripcion")
                tLi.SubItems(2) = .Fields("PRECIO_COSTO")
                .MoveNext
            Loop
        End If
    End With
Exit Sub
ManejaError:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
    Err.Clear
End Sub
Private Sub Image2_Click()
On Error GoTo ManejaError
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
    If ListView1.ListItems.COUNT > 0 Then
        If Ruta <> "" Then
            NumColum = ListView3.ColumnHeaders.COUNT
            For Con = 1 To ListView3.ColumnHeaders.COUNT
                StrCopi = StrCopi & ListView3.ColumnHeaders(Con).Text & Chr(9)
            Next
            StrCopi = StrCopi & Chr(13)
            For Con = 1 To ListView3.ListItems.COUNT
                StrCopi = StrCopi & ListView3.ListItems.Item(Con) & Chr(9)
                For Con2 = 1 To NumColum - 1
                    StrCopi = StrCopi & ListView3.ListItems.Item(Con).SubItems(Con2) & Chr(9)
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
    Exit Sub
ManejaError:
    If Err.Number <> 32755 Then
        MsgBox "Error: " & Err.Number & " " & Err.Description, vbCritical, "SACC"
        Err.Clear
    End If
End Sub
Private Sub Image9_Click()
    If ListView2.ListItems.COUNT = 0 Then
        Unload Me
    Else
        MsgBox "No puede salir si tiene productos pendeintes de imprimir en la salida!", vbExclamation, "SACC"
    End If
End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    SelItm = Item
    sCosto = Item.SubItems(2)
    If Option1.value Then
        Text3.Text = Item.SubItems(1)
    Else
        Text3.Text = Item
    End If
End Sub
Private Sub ListView2_ItemClick(ByVal Item As MSComctlLib.ListItem)
    SelImp1 = Item
    SelImp2 = Item.SubItems(1)
    SelImp3 = Item.SubItems(2)
    SelImp4 = Item.SubItems(3)
    SelImp5 = Item.SubItems(4)
    ind = Item.Index
    Command2.Enabled = True
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
Private Sub Text3_GotFocus()
    Text3.BackColor = &HFFE1E1
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Buscas
    End If
End Sub
Private Sub Text3_LostFocus()
    Text3.BackColor = &H80000005
End Sub
Private Sub Text4_GotFocus()
    Text4.BackColor = &HFFE1E1
End Sub
Private Sub Text4_KeyPress(KeyAscii As Integer)
    Dim Valido As String
    Valido = "1234567890.,"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii > 26 Then
        If InStr(Valido, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Text6_KeyPress(KeyAscii As Integer)
    If Option3.value = True Or Option4.value = True Or Option5.value = True Then
        If KeyAscii = 13 Then
            Command8.value = True
        End If
     Else
      MsgBox "SELECCIONE  UAN FORMA DE BUSQUEDA", vbInformation, "SACC"
    End If
    Text6.SetFocus
End Sub
Private Sub Text4_LostFocus()
    Text4.BackColor = &H80000005
End Sub
