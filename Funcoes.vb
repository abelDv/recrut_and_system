Public Class Funcoes
    Function retornaDadosSemRepetir(dgv As DataGridView, colunaDataGrid As Integer) As String()

        Dim num As Integer = dgv.Rows.Count + 1
        Dim quant As Integer = 0
        Dim jaExiste As Boolean = False
        Dim dadosString(num) As String

        Try
            If num > 1 Then

                For i = 0 To num - 2
                    If dgv.Item(colunaDataGrid, i).Value <> "" Then
                        For j = 0 To quant
                            If dadosString(j) = dgv.Item(colunaDataGrid, i).Value Then
                                jaExiste = True
                                Exit For
                            End If
                        Next j

                        If jaExiste = False Then
                            dadosString(quant) = dgv.Item(colunaDataGrid, i).Value
                            quant += 1
                        End If
                    End If

                    jaExiste = False

                Next i

            End If

            If quant > 0 Then
                quant -= 1
            End If

            ReDim Preserve dadosString(quant)

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        Return dadosString

    End Function

    Function retornaDadosForaLista(lista As String(), dt As DataTable, coluna As String) As String
        Dim dr As DataRow
        Dim quantCargoDt As Integer = dt.Rows.Count - 1
        Dim quant As Integer = lista.Count - 1
        Dim naoTem As String = vbNullString
        Dim jaExiste As Boolean = False

        Try
            For h = 0 To quant
                For k = 0 To quantCargoDt
                    dr = dt.Rows(k)
                    If dr(coluna) = lista(h) Then
                        jaExiste = True
                        Exit For
                    End If
                Next

                If jaExiste = False Then
                    If naoTem = vbNullString Then
                        naoTem = vbCr + lista(h) + vbCr
                    Else
                        naoTem += lista(h) + vbCr
                    End If
                End If

                jaExiste = False
            Next
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        Return naoTem

    End Function
    Sub preencheCamposDgv(dgv As DataGridView, dt As DataTable, colDgvCompara As Integer, colDgvPreenche As Integer, colDt As String)
        Dim dr As DataRow
        Dim ult As Integer = dgv.Rows.Count - 1
        Dim ultDt As Integer = dt.Rows.Count - 1

        Try
            For i = 0 To ult
                For j = 0 To ultDt
                    dr = dt.Rows(j)
                    If dr(0).ToString = dgv.Item(colDgvCompara, i).Value Then
                        dgv.Item(colDgvPreenche, i).Value = dr(colDt).ToString
                        Exit For
                    End If
                Next
            Next

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    Sub preencheCamposDgvDoisValores(dgv As DataGridView, dt As DataTable, colDgvCompara As Integer, colDgvPreenche As Integer, colDt As String, seTiver As String, seNaoTiver As String)
        Dim dr As DataRow
        Dim ult As Integer = dgv.Rows.Count - 1
        Dim ultDt As Integer = dt.Rows.Count - 1
        Dim existe As Boolean = False

        Try
            For i = 0 To ult
                For j = 0 To ultDt
                    dr = dt.Rows(j)
                    If dr(0).ToString = dgv.Item(colDgvCompara, i).Value Then
                        dgv.Item(colDgvPreenche, i).Value = seTiver
                        existe = True
                        Exit For
                    End If
                Next

                If existe = False Then
                    dgv.Item(colDgvPreenche, i).Value = seNaoTiver
                End If

                existe = False

            Next

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    Sub preencheConcatenaQuatroCamposDgv(dgv As DataGridView, campo1 As Integer, campo2 As Integer, campo3 As Integer, campo4 As Integer, separador As String, colPreencheDgv As Integer)
        Dim ult As Integer = dgv.Rows.Count - 1

        Try
            For i = 0 To ult
                dgv.Item(colPreencheDgv, i).Value = dgv.Item(campo1, i).Value + separador + dgv.Item(campo2, i).Value + separador + dgv.Item(campo3, i).Value + separador + dgv.Item(campo4, i).Value
            Next

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    Sub backcolorMDI(formulario As Form, cor As Color)
        formulario.BackColor = cor
    End Sub

End Class
