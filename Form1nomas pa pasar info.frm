VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4875
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8940
   LinkTopic       =   "Form1"
   ScaleHeight     =   4875
   ScaleWidth      =   8940
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   3840
      TabIndex        =   0
      Top             =   2160
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnn As ADODB.Connection
Private cnnPRO As ADODB.Connection
Private Sub Command1_Click()
    Dim sBuscar As String
    Dim tRs As Recordset
    sBuscar = "SELECT * FROM USUARIOS where estado = 'A'"
    Set tRs = cnn.Execute(sBuscar)
    If Not (tRs.EOF And tRs.BOF) Then
        Do While Not tRs.EOF
            sBuscar = "INSERT INTO USUARIOS (ID_USUARIO, NOMBRE, APELLIDOS, PUESTO, PASSWORD, ID_SUCURSAL, ESTADO, Mail, MailPass, Pe1, Pe2, Pe3, Pe4, Pe5, Pe6, Pe7, Pe8, Pe9, Pe10, Pe11, Pe12, Pe13, Pe14, Pe15, Pe16, Pe17, Pe18, Pe19, Pe20, Pe21, Pe22, Pe23, Pe24, Pe25, Pe26, Pe27, Pe28, Pe29, Pe30, Pe31, Pe32, Pe33, Pe34, Pe35, Pe36, Pe37, Pe38, Pe39, Pe40, Pe41, Pe42, Pe43, Pe44, Pe45, Pe46, Pe47, Pe48, Pe49, Pe50, Pe51, Pe52, Pe53, Pe54, Pe55, Pe56, Pe57, Pe58, Pe59, Pe60, Pe61, Pe62, Pe63, Pe64, Pe65, Pe66, Pe67, Pe68, Pe69) "
sBuscar = sBuscar & " VALUES ('" & tRs.Fields("ID_USUARIO") & "', '" & tRs.Fields("NOMBRE") & "', '" & tRs.Fields("APELLIDOS") & "', '" & tRs.Fields("PUESTO") & "', '" & tRs.Fields("PASSWORD") & "', '" & tRs.Fields("ID_SUCURSAL") & "', '" & tRs.Fields("ESTADO") & "', '" & tRs.Fields("Mail") & "', '" & tRs.Fields("MailPass") & "', '" & tRs.Fields("Pe1") & "', '" & tRs.Fields("Pe2") & "', '" & tRs.Fields("Pe3") & "', '" & tRs.Fields("Pe4") & "', '" & tRs.Fields("Pe5") & "', '" & tRs.Fields("Pe6") & "', '" & tRs.Fields("Pe7") & "', '" & tRs.Fields("Pe8") & "', '" & tRs.Fields("Pe9") & "', '" & tRs.Fields("Pe10") & "', '" & tRs.Fields("Pe11") & "', '" & tRs.Fields("Pe12") & "', '" & tRs.Fields("Pe13") & "', '" & tRs.Fields("Pe14") & "', '" & tRs.Fields("Pe15") & "', '" & tRs.Fields("Pe16") & "', '" & tRs.Fields("Pe17") & "', '" & tRs.Fields("Pe18") & "', '" & tRs.Fields("Pe19") & "', '" & tRs.Fields("Pe20") & "', '" & tRs.Fields("Pe21") & "', '" & tRs.Fields("Pe22") & "', '" & tRs.Fields("Pe23") & "'"
sBuscar = sBuscar & ", '" & tRs.Fields("Pe24") & "', '" & tRs.Fields("Pe25") & "', '" & tRs.Fields("Pe26") & "', '" & tRs.Fields("Pe27") & "', '" & tRs.Fields("Pe28") & "', '" & tRs.Fields("Pe29") & "', '" & tRs.Fields("Pe30") & "', '" & tRs.Fields("Pe31") & "', '" & tRs.Fields("Pe32") & "', '" & tRs.Fields("Pe33") & "', '" & tRs.Fields("Pe34") & "', '" & tRs.Fields("Pe35") & "', '" & tRs.Fields("Pe36") & "', '" & tRs.Fields("Pe37") & "', '" & tRs.Fields("Pe38") & "', '" & tRs.Fields("Pe39") & "', '" & tRs.Fields("Pe40") & "', '" & tRs.Fields("Pe41") & "', '" & tRs.Fields("Pe42") & "', '" & tRs.Fields("Pe43") & "', '" & tRs.Fields("Pe44") & "', '" & tRs.Fields("Pe45") & "', '" & tRs.Fields("Pe46") & "', '" & tRs.Fields("Pe47") & "', '" & tRs.Fields("Pe48") & "', '" & tRs.Fields("Pe49") & "', '" & tRs.Fields("Pe50") & "', '" & tRs.Fields("Pe51") & "', '" & tRs.Fields("Pe52") & "', '" & tRs.Fields("Pe53") & "', '" & tRs.Fields("Pe54") & "', '" & tRs.Fields("Pe55") & "', '" & tRs.Fields("Pe56") & "', '"
sBuscar = sBuscar & tRs.Fields("Pe57") & "', '" & tRs.Fields("Pe58") & "', '" & tRs.Fields("Pe59") & "', '" & tRs.Fields("Pe60") & "', '" & tRs.Fields("Pe61") & "', '" & tRs.Fields("Pe62") & "', '" & tRs.Fields("Pe63") & "', '" & tRs.Fields("Pe64") & "', '" & tRs.Fields("Pe65") & "', '" & tRs.Fields("Pe66") & "', '" & tRs.Fields("Pe67") & "', '" & tRs.Fields("Pe68") & "', '" & tRs.Fields("Pe69") & "')"
MsgBox sBuscar
            cnnPRO.Execute (sBuscar)
            tRs.MoveNext
        Loop
    End If
    MsgBox "LISTO!!!"
End Sub
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    Set cnn = New ADODB.Connection
    With cnn
        .ConnectionString = _
            "Provider=SQLOLEDB.1;Password=" & GetSetting("APTONER", "ConfigSACC", "PASSWORD", "LINUX") & ";Persist Security Info=True;User ID=" & GetSetting("APTONER", "ConfigSACC", "USUARIO", "LINUX") & ";Initial Catalog=" & GetSetting("APTONER", "ConfigSACC", "DATABASE", "APTONER") & ";" & _
            "Data Source=" & GetSetting("APTONER", "ConfigSACC", "SERVIDOR", "LINUX") & ";"
        .Open
    End With
    Set cnnPRO = New ADODB.Connection
    With cnnPRO
        .ConnectionString = _
            "Provider=SQLOLEDB.1;Password=" & GetSetting("APTONER", "ConfigSACC", "PASSWORD", "LINUX") & ";Persist Security Info=True;User ID=" & GetSetting("APTONER", "ConfigSACC", "USUARIO", "LINUX") & ";Initial Catalog=PROPACO;" & _
            "Data Source=" & GetSetting("APTONER", "ConfigSACC", "SERVIDOR", "LINUX") & ";"
        .Open
    End With
End Sub
