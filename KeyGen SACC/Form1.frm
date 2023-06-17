VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "KeyGen"
   ClientHeight    =   2115
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4515
   LinkTopic       =   "Form1"
   ScaleHeight     =   2115
   ScaleWidth      =   4515
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Generar"
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   1560
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   720
      TabIndex        =   1
      Top             =   960
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   720
      TabIndex        =   0
      Top             =   360
      Width           =   3135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    On Error GoTo Borralo
    Dim LonNum As Integer
    Dim NSRegistro As String
    Dim NSRegistro2 As String
    Dim StrFinal
    Dim ConInt As String
    Dim Con As Integer
    NSRegistro = Int(Text1.Text / (Len(Text1.Text) * 3.5))
    LonNum = Len(NSRegistro)
    NSRegistro = NSRegistro * LonNum
    LonNum = LonNum + 1
    For Con = 1 To Len(NSRegistro)
        StrFinal = StrFinal & Mid(NSRegistro, LonNum, 1)
        LonNum = LonNum - 1
    Next Con
    LonNum = Len(NSRegistro)
    LonNum = CDbl(LonNum) + 3
    StrFinal = StrFinal & "-" & CDbl(NSRegistro) * CDbl(LonNum)
    StrFinal = Replace(StrFinal, "0", "9")
    StrFinal = Replace(StrFinal, "1", "5")
    StrFinal = Replace(StrFinal, "2", "3")
    StrFinal = Replace(StrFinal, "3", "7")
    StrFinal = Replace(StrFinal, "4", "1")
    StrFinal = Replace(StrFinal, "5", "8")
    StrFinal = Replace(StrFinal, "6", "2")
    StrFinal = Replace(StrFinal, "7", "0")
    StrFinal = Replace(StrFinal, "8", "6")
    StrFinal = Replace(StrFinal, "9", "3")
    Me.Text2.Text = StrFinal
Borralo:
    Err.Clear
End Sub
