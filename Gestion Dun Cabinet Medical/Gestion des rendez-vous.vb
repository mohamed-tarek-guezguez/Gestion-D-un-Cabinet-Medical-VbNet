Imports System.Data.OleDb
Public Class Gestion_des_rendez_vous
    Public pos_patients As Integer
    Public aide As Integer
    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        Try
            StartForm.Show()
            Me.Close()
        Catch
        End Try
    End Sub

    Private Sub Gestion_des_rendez_vous_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            If openform Then
                Rendez_Vous_load()
                Affichage_rendezvous_load()
                DataGridView1.DataSource = Affichage_Rendezvous
                datgridview_setting_rendezvous()
            Else
                RendezVousLibre_load()
                RendezVousLibret_load()
                DataGridView1.DataSource = RendezVousLibre
                datgridview_setting_rendezvous()
            End If
        Catch
        End Try
    End Sub

    Private Sub PictureBox2_Click_1(sender As Object, e As EventArgs) Handles PictureBox2.Click
        Try
            Ajouter_RendezVous.Text = "Ajouter Rendez_Vous"
            Ajouter_RendezVous.Show()
            Me.Close()
        Catch
        End Try
    End Sub

    Private Sub PictureBox4_Click_1(sender As Object, e As EventArgs) Handles PictureBox4.Click
        Try
            Ajouter_RendezVous.Text = "Modifier Rendez_Vous"
            Dim pos As Integer = DataGridView1.CurrentRow.Index

            If openform Then
                Rendez_Vous_load()
                Dim resxx = Affichage_Rendezvous.Rows(pos).Item("Date_R")
                For rx As Integer = 0 To Rendez_Vous.Rows.Count - 1
                    If resxx = Rendez_Vous.Rows(rx).Item("Date_R") Then
                        pos = rx
                    End If
                Next
                pos_patients = pos
                aide = pos
            Else
                RendezVousLibret_load()
                Dim resxx = RendezVousLibre.Rows(pos).Item("Date_R")
                For rx As Integer = 0 To RendezVousLibret.Rows.Count - 1
                    If resxx = RendezVousLibret.Rows(rx).Item("Date_R") Then
                        pos = rx
                    End If
                Next
                pos_patients = pos
                aide = pos
            End If

            Ajouter_RendezVous.Show()
            Me.Close()
        Catch
        End Try
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        Try
            If openform Then
                Search_rend(TextBox1.Text, ComboBox1.Text)
                If TextBox1.Text = "" Then
                    Affichage_rendezvous_load()
                End If
            Else
                Search_rendlib(TextBox1.Text, ComboBox1.Text)
                If TextBox1.Text = "" Then
                    RendezVousLibre_load()
                End If
            End If
        Catch
        End Try
    End Sub

    Private Sub PictureBox3_Click(sender As Object, e As EventArgs) Handles PictureBox3.Click
        Try
            Dim pos As Integer = DataGridView1.CurrentRow.Index

            If MsgBox("Delete " & DataGridView1.Rows(pos).Cells("Nom").Value & " ?", MsgBoxStyle.OkCancel) = MsgBoxResult.Ok Then
                If openform Then
                    Rendez_Vous_load()
                    Dim v = Affichage_Rendezvous.Rows(pos).Item("Date_R").ToString

                    Dim b As Integer
                    For xv As Integer = 0 To Rendez_Vous.Rows.Count - 1
                        If v = Rendez_Vous.Rows(xv).Item("Date_R") Then
                            b = xv
                        End If
                    Next

                    Rendez_Vous.Rows(b).Delete()
                    Rendez_Vous_Save()
                    Rendez_Vous_load()
                    Affichage_rendezvous_load()
                Else
                    RendezVousLibret_load()
                    Dim v = RendezVousLibre.Rows(pos).Item("Date_R").ToString

                    Dim b As Integer
                    For xv As Integer = 0 To RendezVousLibret.Rows.Count - 1
                        If v = RendezVousLibret.Rows(xv).Item("Date_R") Then
                            b = xv
                        End If
                    Next

                    RendezVousLibret.Rows(b).Delete()
                    RendezVousLibret_Save()
                    RendezVousLibret_load()
                    RendezVousLibre_load()
                End If
            End If
        Catch

        End Try
    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
End Class