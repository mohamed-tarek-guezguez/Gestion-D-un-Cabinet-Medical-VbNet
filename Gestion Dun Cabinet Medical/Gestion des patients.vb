Imports System.Data.OleDb

Public Class Gestion_des_patients
    Public pos_patients As Integer
    Public aide As Integer
    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        Try
            StartForm.Show()
            Me.Close()
        Catch
        End Try
    End Sub

    Private Sub Gestion_des_patients_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            Affichage_pat_load()
            DataGridView1.DataSource = Affichage_Patients
            datgridview_setting_patients()
        Catch
        End Try
    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click
        Try
            Ajouter_Patient.Text = "Ajouter Patient"
            Ajouter_Patient.Show()
            Me.Close()
        Catch
        End Try
    End Sub

    Private Sub PictureBox3_Click(sender As Object, e As EventArgs) Handles PictureBox3.Click
        Try
            Dim pos As Integer = DataGridView1.CurrentRow.Index
            If MsgBox("Delete " & DataGridView1.Rows(pos).Cells("Nom").Value & " ?", MsgBoxStyle.OkCancel) = MsgBoxResult.Ok Then
                Dim position As Integer = BindingContext(Affichage_Patients).Position
                Dim posy As Integer = DataGridView1.CurrentRow.Index
                Dim res = Affichage_Patients.Rows(posy).Item("Numero")

                Rendez_Vous_load()
                Dim verif As Boolean = True
                For xx As Integer = 0 To Rendez_Vous.Rows.Count - 1
                    If res = Rendez_Vous.Rows(xx).Item("Numero") Then
                        Rendez_Vous.Rows(xx).Delete()
                    End If
                Next
                Rendez_Vous_Save()

                Affichage_Patients.Rows(position).Delete()
                Affichage_pat_Save()
                Affichage_pat_load()
            End If
        Catch
        End Try
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        Try
            Search_pat(TextBox1.Text, ComboBox1.Text)
            If TextBox1.Text = "" Then
                Affichage_pat_load()
            End If
        Catch
        End Try
    End Sub

    Private Sub PictureBox4_Click(sender As Object, e As EventArgs) Handles PictureBox4.Click
        Try
            Ajouter_Patient.Text = "Modifier Patient"
            Dim pos As Integer = DataGridView1.CurrentRow.Index

            Patients_load()
            Dim resxx = Affichage_Patients.Rows(pos).Item("Numero")
            For rx As Integer = 0 To Patients.Rows.Count - 1
                If resxx = Patients.Rows(rx).Item("Numero") Then
                    pos = rx
                End If
            Next

            pos_patients = pos
            aide = pos
            Ajouter_Patient.Show()
            Me.Close()
        Catch
        End Try
    End Sub

End Class