Public Class principal

    Dim corDeFundo As Color = Color.LightBlue

    Private objBanco As New Access
    Private objCoresStrip As New CoresStrip
    Private objfuncoes As New Funcoes

    Private Sub CargoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CargoToolStripMenuItem.Click
        Dim formulario As Form

        formulario = formCadCargo

        With formulario
            .MdiParent = Me
            .Show()
            .Top = 0
        End With

        objfuncoes.backcolorMDI(formulario, corDeFundo)
    End Sub

    Private Sub MotivoTornaSemEfeitoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MotivoTornaSemEfeitoToolStripMenuItem.Click
        Dim formulario As Form

        formulario = formMotTorna

        With formulario
            .MdiParent = Me
            .Show()
            .Top = 0
        End With

        objfuncoes.backcolorMDI(formulario, corDeFundo)
    End Sub

    Private Sub SecretariaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SecretariaToolStripMenuItem.Click
        Dim formulario As Form

        formulario = FormSecretarias

        With formulario
            .MdiParent = Me
            .Show()
            .Top = 0
        End With

        objfuncoes.backcolorMDI(formulario, corDeFundo)
    End Sub

    Private Sub DocumentoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DocumentoToolStripMenuItem.Click
        Dim formulario As Form

        formulario = formCadDoc

        With formulario
            .MdiParent = Me
            .Show()
            .Top = 0
        End With

        objfuncoes.backcolorMDI(formulario, corDeFundo)
    End Sub

    Private Sub EditalToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles EditalToolStripMenuItem.Click
        Dim formulario As Form

        formulario = formEdital

        With formulario
            .MdiParent = Me
            .Show()
            .Top = 0
        End With

        objfuncoes.backcolorMDI(formulario, corDeFundo)
    End Sub

    Private Sub LeiautePlanilhaEditalToolStripMenuItem_Click_1(sender As Object, e As EventArgs) Handles LeiautePlanilhaEditalToolStripMenuItem.Click
        Dim formulario As Form

        formulario = formEditalLeiaute

        With formulario
            .MdiParent = Me
            .Show()
            .Top = 0
        End With

        objfuncoes.backcolorMDI(formulario, corDeFundo)
    End Sub

    Private Sub CandidatoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CandidatoToolStripMenuItem.Click
        Dim formulario As Form

        formulario = FormCand

        With formulario
            .MdiParent = Me
            .Show()
            .Top = 0
        End With

        objfuncoes.backcolorMDI(formulario, corDeFundo)
    End Sub

    Private Sub InscriçãoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles InscriçãoToolStripMenuItem.Click
        Dim formulario As Form

        formulario = formInsc

        With formulario
            .MdiParent = Me
            .Show()
            .Top = 0
        End With

        objfuncoes.backcolorMDI(formulario, corDeFundo)
    End Sub

    Private Sub VagaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles VagaToolStripMenuItem.Click
        Dim formulario As Form

        formulario = formVagas

        With formulario
            .MdiParent = Me
            .Show()
            .Top = 0
        End With

        objfuncoes.backcolorMDI(formulario, corDeFundo)
    End Sub

    Private Sub ConvocaçãoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ConvocaçãoToolStripMenuItem.Click
        Dim formulario As Form

        formulario = formConvocacao

        With formulario
            .MdiParent = Me
            .Show()
            .Top = 0
        End With

        objfuncoes.backcolorMDI(formulario, corDeFundo)
    End Sub

    Private Sub AndamentoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AndamentoToolStripMenuItem.Click
        Dim formulario As Form

        formulario = FormAndamento

        With formulario
            .MdiParent = Me
            .Show()
            .Top = 0
        End With

        objfuncoes.backcolorMDI(formulario, corDeFundo)
    End Sub

    Private Sub EnviaEmailToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles EnviaEmailToolStripMenuItem.Click
        Dim formulario As Form

        formulario = formEnviaEmail

        With formulario
            .MdiParent = Me
            .Show()
            .Top = 0
        End With

        objfuncoes.backcolorMDI(formulario, corDeFundo)

    End Sub

    Private Sub UsuárioToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles UsuárioToolStripMenuItem.Click
        Dim formulario As Form

        formulario = FormUser

        With formulario
            .MdiParent = Me
            .Show()
            .Top = 0
        End With

        objfuncoes.backcolorMDI(formulario, corDeFundo)
    End Sub

    Private Sub principal_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Try
            objBanco.connect()
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            objBanco.close()
        End Try

        MenuStrip1.RenderMode = ToolStripRenderMode.Professional
        MenuStrip1.Renderer = New MyToolStripProfRender

        For Each controle As Control In Me.Controls
            If controle.GetType Is GetType(MdiClient) Then
                controle.BackColor = Color.LightBlue
            End If
        Next

    End Sub

    Private Sub PerdaTornaCancelaPosseToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PerdaTornaCancelaPosseToolStripMenuItem.Click
        Dim formulario As Form

        formulario = formTornaSemEfeito

        With formulario
            .MdiParent = Me
            .Show()
            .Top = 0
        End With

        objfuncoes.backcolorMDI(formulario, corDeFundo)
    End Sub

    Private Sub InformaRemoçãoETrocaDocumentoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles InformaRemoçãoETrocaDocumentoToolStripMenuItem.Click
        Dim formulario As Form

        formulario = formRemocaoTrocaDoc

        With formulario
            .MdiParent = Me
            .Show()
            .Top = 0
        End With

        objfuncoes.backcolorMDI(formulario, corDeFundo)
    End Sub

    Private Sub ImportaPlanilhaCertameToolStripMenuItem_Click_1(sender As Object, e As EventArgs) Handles ImportaPlanilhaCertameToolStripMenuItem.Click
        Dim formulario As Form

        formulario = formImportaPlan

        With formulario
            .MdiParent = Me
            .Show()
            .Top = 0
        End With

        objfuncoes.backcolorMDI(formulario, corDeFundo)
    End Sub
End Class

