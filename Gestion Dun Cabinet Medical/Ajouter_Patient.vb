Imports System.Data.OleDb
Public Class Ajouter_Patient

    Public offff As Integer

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            Gestion_des_patients.Show()
            Me.Close()
        Catch
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Try
            Dim com1 As Boolean = False
            For i = 0 To ComboBox1.Items.Count - 1
                If ComboBox1.Text = ComboBox1.Items(i) Then
                    com1 = True
                End If
            Next
            Dim com2 As Boolean = False
            For i = 0 To ComboBox2.Items.Count - 1
                If ComboBox2.Text = ComboBox2.Items(i) Then
                    com2 = True
                End If
            Next
            If TextBox8.Text.Length = 4 And IsNumeric(TextBox8.Text) And com1 And com2 Then
                TextBox5.Text = Label9.Text + Label8.Text + TextBox8.Text
                If TextBox3.Text.Trim = "" Or TextBox1.Text.Trim = "" Or TextBox2.Text.Trim = "" Or TextBox5.Text.Trim = "" Or TextBox6.Text.Trim = "" Then
                    If TextBox3.Text.Trim = "" Then
                        MsgBox(Title:="Error", Prompt:="Numéro Vide ...")
                    ElseIf TextBox1.Text.Trim = "" Then
                        MsgBox(Title:="Error", Prompt:="Nom Vide ...")
                    ElseIf TextBox2.Text.Trim = "" Then
                        MsgBox(Title:="Error", Prompt:="Prénom Vide ...")
                    ElseIf TextBox5.Text.Trim = "" Then
                        MsgBox(Title:="Error", Prompt:="Date de naissance Vide ...")
                    ElseIf TextBox6.Text.Trim = "" Then
                        MsgBox(Title:="Error", Prompt:="Lieu de naissance Vide ...")
                    End If
                Else
                    If IsNumeric(TextBox3.Text) Then
                        If Me.Text = "Modifier Patient" Then
                            Dim test As Boolean = True
                            For t As Integer = 0 To Affichage_Patients.Rows.Count - 1
                                If TextBox3.Text <> Label10.Text Then
                                    If (TextBox3.Text = Affichage_Patients.Rows(t).Item("Numero")) Then
                                        test = False
                                    End If
                                End If
                            Next
                            If test = False Then
                                MsgBox("Numéro exist ...")
                            Else
                                Affichage_pat_load()
                                Affichage_Patients.Rows(offff).Item("Numero") = TextBox3.Text
                                Affichage_Patients.Rows(offff).Item("Nom") = TextBox1.Text
                                Affichage_Patients.Rows(offff).Item("Prenom") = TextBox2.Text
                                Affichage_Patients.Rows(offff).Item("Adresse") = TextBox4.Text
                                Affichage_Patients.Rows(offff).Item("Date_N") = TextBox5.Text
                                Affichage_Patients.Rows(offff).Item("Lieu") = TextBox6.Text
                                Affichage_Patients.Rows(offff).Item("Etat") = TextBox7.Text
                                Affichage_pat_Save()
                                Affichage_pat_load()
                                Gestion_des_patients.Show()
                                Me.Close()
                            End If

                            If TextBox3.Text <> Label10.Text Then
                                Rendez_Vous_load()
                                Dim toor As Boolean = True
                                For ok As Integer = 0 To Rendez_Vous.Rows.Count - 1
                                    If Label10.Text = Rendez_Vous.Rows(ok).Item("Numero") Then
                                        Rendez_Vous.Rows(ok).Item("Numero") = TextBox3.Text
                                    End If
                                Next
                                Rendez_Vous_Save()
                            End If

                        Else
                            Dim verif As Boolean = True
                            For xx As Integer = 0 To Affichage_Patients.Rows.Count - 1
                                If TextBox3.Text.Trim.ToString = Affichage_Patients.Rows(xx).Item("Numero") Then
                                    verif = False
                                End If
                            Next
                            If verif = False Then
                                MsgBox("Numéro exist ...")
                            Else
                                Affichage_pat_load()
                                Affichage_Patients.Rows.Add(TextBox3.Text, TextBox1.Text, TextBox2.Text, TextBox4.Text, TextBox5.Text, TextBox6.Text, TextBox7.Text)
                                Affichage_pat_Save()
                                Gestion_des_patients.Show()
                                Me.Close()
                            End If
                        End If
                    Else
                        MsgBox(Title:="Error", Prompt:="Numéro incorrect ...")
                    End If
                End If
            Else
                MsgBox(Title:="Error", Prompt:="Date incorrect ...")
            End If
        Catch
        End Try
    End Sub

    Private Sub Ajouter_Patient_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            offff = Gestion_des_patients.aide
            If Me.Text = "Modifier Patient" Then
                Button2.Text = "Modifier"
                Affichage_pat_load()

                TextBox3.Text = Affichage_Patients.Rows(Gestion_des_patients.pos_patients).Item("Numero")
                Label10.Text = TextBox3.Text
                TextBox1.Text = Affichage_Patients.Rows(Gestion_des_patients.pos_patients).Item("Nom")
                TextBox2.Text = Affichage_Patients.Rows(Gestion_des_patients.pos_patients).Item("Prenom")
                TextBox4.Text = Affichage_Patients.Rows(Gestion_des_patients.pos_patients).Item("Adresse")
                TextBox6.Text = Affichage_Patients.Rows(Gestion_des_patients.pos_patients).Item("Lieu")
                TextBox7.Text = Affichage_Patients.Rows(Gestion_des_patients.pos_patients).Item("Etat")

                TextBox5.Text = Affichage_Patients.Rows(Gestion_des_patients.pos_patients).Item("Date_N")
                TextBox8.Text = TextBox5.Text.ToCharArray(6, 4)
                ComboBox2.Text = TextBox5.Text.ToCharArray(0, 2)
                Select Case TextBox5.Text.ToCharArray(3, 2)
                    Case "01"
                        ComboBox1.Text = "Janvier"
                    Case "02"
                        ComboBox1.Text = "Février"
                    Case "03"
                        ComboBox1.Text = "Mars"
                    Case "04"
                        ComboBox1.Text = "Avril"
                    Case "05"
                        ComboBox1.Text = "Mai"
                    Case "06"
                        ComboBox1.Text = "Juin"
                    Case "07"
                        ComboBox1.Text = "Juillet"
                    Case "08"
                        ComboBox1.Text = "Août"
                    Case "09"
                        ComboBox1.Text = "Septembre"
                    Case "10"
                        ComboBox1.Text = "Octobre"
                    Case "11"
                        ComboBox1.Text = "Novembre"
                    Case "12"
                        ComboBox1.Text = "Décembre"
                End Select
            Else
                Button2.Text = "Ajouter"
                TextBox8.Text = Date.Today.Year.ToString
            End If
        Catch
        End Try
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Try
            If ComboBox1.SelectedIndex + 1 < 10 Then
                Label8.Text = "0" + CStr(ComboBox1.SelectedIndex + 1) + "/"
            Else
                Label8.Text = CStr(ComboBox1.SelectedIndex + 1) + "/"
            End If
            If ComboBox1.Text = "Janvier" Or ComboBox1.Text = "Mars" Or ComboBox1.Text = "Mai" Or ComboBox1.Text = "Juillet" Or ComboBox1.Text = "Août" Or ComboBox1.Text = "Octobre" Or ComboBox1.Text = "Décembre" Then
                ComboBox2.Items.Clear()
                For i As Integer = 1 To 31
                    ComboBox2.Items.Add(i)
                Next
            ElseIf ComboBox1.Text = "Avril" Or ComboBox1.Text = "Juin" Or ComboBox1.Text = "Septembre" Or ComboBox1.Text = "Novembre" Then
                ComboBox2.Items.Clear()
                For i As Integer = 1 To 30
                    ComboBox2.Items.Add(i)
                Next
            Else
                If TextBox8.Text.Length = 4 And ComboBox1.Text = "Février" Then
                    If CInt(Date.Today.Year.ToString) Mod 4 = 0 Then
                        ComboBox2.Items.Clear()
                        For i As Integer = 1 To 29
                            ComboBox2.Items.Add(i)
                        Next
                    Else
                        ComboBox2.Items.Clear()
                        For i As Integer = 1 To 28
                            ComboBox2.Items.Add(i)
                        Next
                    End If
                End If
            End If
        Catch
        End Try
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        Try
            If ComboBox2.SelectedIndex + 1 < 10 Then
                Label9.Text = "0" + CStr(ComboBox2.SelectedIndex + 1) + "/"
            Else
                Label9.Text = CStr(ComboBox2.SelectedIndex + 1) + "/"
            End If
        Catch
        End Try
    End Sub

    Private Sub TextBox8_TextChanged(sender As Object, e As EventArgs) Handles TextBox8.TextChanged
        Try
            If IsNumeric(TextBox8.Text) And TextBox8.Text.Length = 4 Then
                If ComboBox1.Text = "Février" Then
                    If CInt(TextBox8.Text) Mod 4 = 0 Then
                        ComboBox2.Items.Clear()
                        For i As Integer = 1 To 29
                            ComboBox2.Items.Add(i)
                        Next
                    Else
                        ComboBox2.Items.Clear()
                        For i As Integer = 1 To 28
                            ComboBox2.Items.Add(i)
                        Next
                    End If
                End If
            End If
        Catch
        End Try
    End Sub

End Class