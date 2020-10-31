Imports System.Data.OleDb
Public Class Gestion_des_médicaments

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        Try
            StartForm.Show()
            Me.Close()
        Catch
        End Try
    End Sub

   

    Private Sub Gestion_des_médicaments_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            Affichage_medic_load()
            DataGridView1.DataSource = Affichage_Medicaments
            datgridview_setting()
        Catch
        End Try
    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click
        Try
            Ajouter_Medicaments.Show()
            Me.Close()
        Catch
        End Try
    End Sub

    Private Sub PictureBox3_Click(sender As Object, e As EventArgs) Handles PictureBox3.Click
        Try
            Dim pos As Integer = DataGridView1.CurrentRow.Index
            If MsgBox("Delete " & DataGridView1.Rows(pos).Cells("Nom").Value & " ?", MsgBoxStyle.OkCancel) = MsgBoxResult.Ok Then
                Dim position As Integer = BindingContext(Affichage_Medicaments).Position
                Affichage_Medicaments.Rows(position).Delete()
                Affichage_medic_Save()
                Affichage_medic_load()
            End If
        Catch
        End Try
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        Try
            Search(TextBox1.Text, ComboBox1.Text)
            If TextBox1.Text = "" Then
                Affichage_medic_load()
            End If
        Catch
        End Try
    End Sub
End Class