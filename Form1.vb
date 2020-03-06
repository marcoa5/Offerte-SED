Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If System.Diagnostics.Debugger.IsAttached Then
            Me.Text = "SED Offer Debug Mode"
        Else
            Me.Text = "SED Offer Version " & My.Application.Deployment.CurrentVersion.ToString
        End If

        Main()
    End Sub

    Private Sub Form1_Activated(sender As Object, e As EventArgs) Handles Me.Activated

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Testo()
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        Controlla()
    End Sub

    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged
        Controlla()
        Rock_Drill(Me.ListBox1.SelectedItem)
    End Sub

    Private Sub ListBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox2.SelectedIndexChanged
        Controlla()
        Dati(Me.ListBox1.SelectedItem, Me.ListBox2.SelectedItem)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If MsgBox("Vuoi modificare le condizioni standard?", vbYesNo, "Modifica") = vbYes Then
            Form3.ShowDialog()
        Else
            Crea_File()
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Application.Exit()
    End Sub

    Private Sub ComboBox1_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedValueChanged
        Testo()
    End Sub

    Private Sub ComboBox1_ChangeUICues(sender As Object, e As UICuesEventArgs) Handles ComboBox1.ChangeUICues

    End Sub

    Private Sub ComboBox1_TextUpdate(sender As Object, e As EventArgs) Handles ComboBox1.TextUpdate
        If Me.ComboBox1.Text = "" Then
            Me.TextBox1.Text = ""
            Me.TextBox2.Text = ""
        End If
        Controlla()
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        Controlla()
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        Controlla()
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        JobTitle()
        Controlla()
    End Sub

    Private Sub ComboBox2_TextUpdate(sender As Object, e As EventArgs) Handles ComboBox2.TextUpdate
        Controlla()
    End Sub

    Private Sub ListBox1_Click(sender As Object, e As EventArgs) Handles ListBox1.Click
        Controlla()
    End Sub

    Private Sub ListBox2_Click(sender As Object, e As EventArgs) Handles ListBox2.Click
        Controlla()
    End Sub

    Private Sub ListBox1_MouseUp(sender As Object, e As MouseEventArgs) Handles ListBox1.MouseUp
        Controlla()
    End Sub
End Class
