Imports System.Data.OleDb
Public Class Ajouter_Medicaments

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            Gestion_des_médicaments.Show()
            Me.Close()
        Catch
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Try

            If TextBox3.Text.Trim = "" Or TextBox1.Text.Trim = "" Or TextBox2.Text.Trim = "" Then
                If TextBox3.Text.Trim = "" Then
                    MsgBox(Title:="Error", Prompt:="Code Vide ...")
                ElseIf TextBox1.Text.Trim = "" Then
                    MsgBox(Title:="Error", Prompt:="Nom Vide ...")
                Else
                    MsgBox(Title:="Error", Prompt:="Quantité Vide ...")
                End If
            Else
                If IsNumeric(TextBox2.Text.Trim) Then
                    Dim verif As Boolean = True
                    For xx As Integer = 0 To Affichage_Medicaments.Rows.Count - 1
                        If TextBox3.Text.Trim.ToString = Affichage_Medicaments.Rows(xx).Item("Code") Then
                            verif = False
                        End If
                    Next
                    If verif = False Then
                        MsgBox("Code exist ...")
                    Else
                        Affichage_medic_load()
                        Affichage_Medicaments.Rows.Add(TextBox3.Text, TextBox1.Text, TextBox2.Text)
                        Affichage_medic_Save()
                        Gestion_des_médicaments.Show()
                        Me.Close()
                    End If
                Else
                    MsgBox(Title:="Error", Prompt:="Quantité incorrect ...")
                End If
            End If
        Catch
        End Try
    End Sub

    Private Sub Ajouter_Medicaments_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class