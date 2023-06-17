VERSION 5.00
Begin VB.Form FrmComparadorPrecios 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Comparador de precios"
   ClientHeight    =   6615
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10425
   LinkTopic       =   "Form1"
   ScaleHeight     =   6615
   ScaleWidth      =   10425
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   8880
      TabIndex        =   0
      Top             =   5760
      Width           =   735
   End
End
Attribute VB_Name = "FrmComparadorPrecios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If Text1.Text <> "" Then
        Dim objExcel As Excel.Application
        Dim xLibro As Excel.Workbook
        Dim sBuscar As String
        Dim tRs1 As ADODB.Recordset
        Dim Col As Integer, Fila As Integer
        Set objExcel = New Excel.Application
        Set xLibro = objExcel.Workbooks.Open(Text1.Text)
        objExcel.Visible = False
        Fila = 1
        With xLibro
            With .Sheets(1)
                Do While .Cells(Fila, 1) <> ""
                    If Option1.Value Then
                        If Option7.Value Then
                            sBuscar = "UPDATE ALMACEN1 SET PRECIO_COSTO = " & .Cells(Fila, 2) & ", PRECIO_EN = 'PESOS' WHERE ID_PRODUCTO = '" & .Cells(Fila, 1) & "'"
                        Else
                            sBuscar = "UPDATE ALMACEN1 SET PRECIO_COSTO = " & .Cells(Fila, 2) & ", PRECIO_EN = 'DOLARES' WHERE ID_PRODUCTO = '" & .Cells(Fila, 1) & "'"
                        End If
                    End If
                    If Option2.Value Then
                        If Option7.Value Then
                            sBuscar = "UPDATE ALMACEN2 SET PRECIO_COSTO = " & .Cells(Fila, 2) & ", PRECIO_EN = 'PESOS' WHERE ID_PRODUCTO = '" & .Cells(Fila, 1) & "'"
                        Else
                            sBuscar = "UPDATE ALMACEN2 SET PRECIO_COSTO = " & .Cells(Fila, 2) & ", PRECIO_EN = 'DOLARES' WHERE ID_PRODUCTO = '" & .Cells(Fila, 1) & "'"
                        End If
                    End If
                    If Option3.Value Then
                        sBuscar = "SELECT PRECIO_COSTO FROM ALMACEN3  WHERE ID_PRODUCTO = '" & .Cells(Fila, 1) & "'"
                        'MsgBox sBuscar
                        Set tRs1 = cnn.Execute(sBuscar)
                        If Not (tRs1.EOF And tRs1.BOF) Then
                            If .Cells(Fila, 2) > tRs1.Fields("PRECIO_COSTO") Then
                                If Option7.Value Then
                                    sBuscar = "UPDATE ALMACEN3 SET PRECIO_COSTO = " & .Cells(Fila, 2) & ", PRECIO_EN = 'PESOS' WHERE ID_PRODUCTO = '" & .Cells(Fila, 1) & "'"
                                Else
                                    sBuscar = "UPDATE ALMACEN3 SET PRECIO_COSTO = " & .Cells(Fila, 2) & ", PRECIO_EN = 'DOLARES' WHERE ID_PRODUCTO = '" & .Cells(Fila, 1) & "'"
                                End If
                                cnn.Execute (sBuscar)
                            End If
                        End If
                        'sBuscar = "UPDATE ALMACEN3 SET PRECIO_COSTO = " & .Cells(Fila, 2) & " WHERE ID_PRODUCTO = '" & .Cells(Fila, 1) & "'"
                    End If
                    cnn.Execute (sBuscar)
                    Fila = Fila + 1
                Loop
            End With
        End With
        objExcel.Workbooks.Close
        Set objExcel = Nothing
        Set xLibro = Nothing
        MsgBox "La importación ha finalizado exitosamente!", vbExclamation, "SACC"
    Else
        MsgBox "No se ha seleccionado ningun archivo a importar!", vbExclamation, "SACC"
    End If
End Sub

