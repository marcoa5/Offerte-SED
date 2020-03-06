Public Class Form3
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Text = Scelta1(0, 1)
        TextBox2.Text = Scelta1(0, 4)
        TextBox3.Text = Scelta1(0, 5)
        TextBox4.Text = Scelta1(0, 23)

        For I As Integer = 5 To 20
            Me.TabPage2.Controls.Item("TextBox" & I.ToString).Text = Replace(Scelta1(0, I + 19), ".", "")
        Next
        For I As Integer = 21 To 27
            Me.TabPage4.Controls.Item("TextBox" & I.ToString).Text = Replace(Scelta1(0, I + 19), ".", "")
        Next
    End Sub

    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Scelta1(0, 1) = TextBox1.Text
        Scelta1(0, 4) = TextBox2.Text
        Scelta1(0, 5) = TextBox3.Text
        Scelta1(0, 23) = TextBox4.Text
        For I As Integer = 5 To 20
            If Me.TabPage2.Controls.Item("TextBox" & I.ToString).Text <> "" Then
                Scelta1(0, I + 19) = Me.TabPage2.Controls.Item("TextBox" & I.ToString).Text
            Else
                Scelta1(0, I + 19) = ""
            End If
        Next
        For I As Integer = 21 To 27
            If Me.TabPage4.Controls.Item("TextBox" & I.ToString).Text <> "." Then
                Scelta1(0, I + 19) = Me.TabPage4.Controls.Item("TextBox" & I.ToString).Text
            Else
                Scelta1(0, I + 19) = ""
            End If
        Next
        For I As Integer = 28 To 30
            Scelta1(0, 46) = Scelta1(0, 46) & vbCr & Me.TabPage4.Controls.Item("TextBox" & I.ToString).Text
        Next


        Crea_File()
    End Sub

End Class