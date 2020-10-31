Imports System.Data.OleDb
Public Class Ajouter_RendezVous
    Public offff As Integer
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            Gestion_des_rendez_vous.Show()
            Me.Close()
        Catch
        End Try
    End Sub

    Private Sub Ajouter_RendezVous_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            If openform Then
                ComboBox3.Items.Clear()
                Affichage_pat_load()
                For yy As Integer = 0 To Affichage_Patients.Rows.Count - 1
                    ComboBox3.Items.Add(Affichage_Patients.Rows(yy).Item("Numero"))
                Next
            Else
                numt.Visible = True
                nomt.Visible = True
                pret.Visible = True
                ComboBox3.Visible = False
                Label4.Visible = False
                Label5.Visible = False
            End If
            offff = Gestion_des_rendez_vous.aide
            If Me.Text = "Modifier Rendez_Vous" Then
                Button2.Text = "Modifier"
                Rendez_Vous_load()
                RendezVousLibre_load()
                If openform Then
                    ComboBox3.Text = Rendez_Vous.Rows(Gestion_des_rendez_vous.pos_patients).Item("Numero")
                    Label10.Text = ComboBox3.Text
                    ComboBox15.Text = Rendez_Vous.Rows(Gestion_des_rendez_vous.pos_patients).Item("Etat_R")
                    TextBox5.Text = Rendez_Vous.Rows(Gestion_des_rendez_vous.pos_patients).Item("Date_R")
                Else
                    numt.Text = RendezVousLibre.Rows(Gestion_des_rendez_vous.pos_patients).Item("Numero")
                    Label10.Text = numt.Text
                    ComboBox15.Text = RendezVousLibre.Rows(Gestion_des_rendez_vous.pos_patients).Item("Etat_R")
                    TextBox5.Text = RendezVousLibre.Rows(Gestion_des_rendez_vous.pos_patients).Item("Date_R")
                    nomt.Text = RendezVousLibre.Rows(Gestion_des_rendez_vous.pos_patients).Item("Nom")
                    pret.Text = RendezVousLibre.Rows(Gestion_des_rendez_vous.pos_patients).Item("Prenom")
                End If
                TextBox2.Text = TextBox5.Text.ToCharArray(6, 4)
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
                TextBox3.Text = TextBox5.Text.ToCharArray(11, 2)
                TextBox4.Text = TextBox5.Text.ToCharArray(14, 2)
                Label15.Text = Label9.Text + Label8.Text + TextBox2.Text + " " + TextBox3.Text + ":" + TextBox4.Text
            Else
                Button2.Text = "Ajouter"
                TextBox2.Text = Date.Today.Year.ToString
                ComboBox3.Text = ComboBox3.Items(0)
            End If
        Catch
        End Try
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        Try
            If IsNumeric(TextBox2.Text) And TextBox2.Text.Length = 4 Then
                If ComboBox1.Text = "Février" Then
                    If CInt(TextBox2.Text) Mod 4 = 0 Then
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
                If TextBox2.Text.Length = 4 And ComboBox1.Text = "Février" Then
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

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Try

            Affichage_pat_load()
            RendezVousLibre_load()
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
            If openform Then
                If TextBox2.Text.Length = 4 And com1 And com2 Then
                    TextBox5.Text = Label9.Text + Label8.Text + TextBox2.Text + " " + TextBox3.Text + ":" + TextBox4.Text
                    Dim dat_test As Boolean = True
                    For da As Integer = 0 To Rendez_Vous.Rows.Count - 1
                        If TextBox5.Text = Rendez_Vous.Rows(da).Item("Date_R") Then
                            dat_test = False
                        End If
                    Next
                    For da As Integer = 0 To RendezVousLibre.Rows.Count - 1
                        If TextBox5.Text = RendezVousLibre.Rows(da).Item("Date_R") Then
                            dat_test = False
                        End If
                    Next
                    If (TextBox5.Text = Label15.Text) And (Me.Text = "Modifier Rendez_Vous") Then
                        dat_test = True
                    End If
                    If dat_test = True Then
                        If TextBox3.Text.Trim = "" Or ComboBox15.Text = "" Or TextBox2.Text.Trim = "" Or TextBox4.Text.Trim = "" Then
                            If TextBox3.Text.Trim = "" Then
                                MsgBox(Title:="Error", Prompt:="Time Vide ...")
                            ElseIf ComboBox15.Text = "" Then
                                MsgBox(Title:="Error", Prompt:="Etat Vide ...")
                            ElseIf TextBox2.Text.Trim = "" Then
                                MsgBox(Title:="Error", Prompt:="Date incorrect ...")
                            ElseIf TextBox4.Text.Trim = "" Then
                                MsgBox(Title:="Error", Prompt:="Time Vide ...")
                            End If
                        Else
                            If IsNumeric(ComboBox3.Text) Then
                                If (TextBox3.Text > 0 And TextBox3.Text < 24) And ((TextBox4.Text > -1) And (TextBox4.Text < 60)) And TextBox3.Text.Length = 2 And TextBox4.Text.Length = 2 Then
                                    If Me.Text = "Modifier Rendez_Vous" Then
                                        Dim test As Boolean = True
                                        For t As Integer = 0 To Affichage_Patients.Rows.Count - 1
                                            If ComboBox3.Text <> Label10.Text Then
                                                If (ComboBox3.Text = Affichage_Patients.Rows(t).Item("Numero")) Then
                                                    test = False
                                                End If
                                            End If
                                            If ComboBox3.Text = Label10.Text Then
                                                test = False
                                            End If
                                        Next
                                        If test = True Then
                                            MsgBox("Numéro n'existe pas ...")
                                        Else
                                            Rendez_Vous_load()
                                            Rendez_Vous.Rows(offff).Item("Numero") = ComboBox3.Text
                                            Rendez_Vous.Rows(offff).Item("Date_R") = TextBox5.Text
                                            Rendez_Vous.Rows(offff).Item("Etat_R") = ComboBox15.Text


                                            Rendez_Vous_Save()
                                            Rendez_Vous_load()
                                            Gestion_des_rendez_vous.Show()
                                            Me.Close()
                                        End If
                                    Else
                                        Dim verif As Boolean = True
                                        For xx As Integer = 0 To Affichage_Patients.Rows.Count - 1
                                            If ComboBox3.Text = Affichage_Patients.Rows(xx).Item("Numero") Then
                                                verif = False
                                            End If
                                        Next
                                        If verif = True Then
                                            MsgBox("Numéro n'existe pas ...")
                                        Else
                                            Rendez_Vous_load()
                                            Rendez_Vous.Rows.Add(ComboBox3.Text, TextBox5.Text, ComboBox15.Text)
                                            Rendez_Vous_Save()
                                            Gestion_des_rendez_vous.Show()
                                            Me.Close()
                                        End If
                                    End If
                                Else
                                    MsgBox(Title:="Error", Prompt:="Time incorrect ...")
                                End If
                            Else
                                MsgBox(Title:="Error", Prompt:="Numéro incorrect ...")
                            End If
                        End If
                    Else
                        MsgBox(Title:="Error", Prompt:="Date existe ...")
                    End If
                Else
                    MsgBox(Title:="Error", Prompt:="Date incorrect ...")
                End If
            Else
                '**************************************************************************************************

                If TextBox2.Text.Length = 4 And com1 And com2 Then
                    TextBox5.Text = Label9.Text + Label8.Text + TextBox2.Text + " " + TextBox3.Text + ":" + TextBox4.Text
                    Dim dat_test As Boolean = True
                    For da As Integer = 0 To RendezVousLibre.Rows.Count - 1
                        If TextBox5.Text = RendezVousLibre.Rows(da).Item("Date_R") Then
                            dat_test = False
                        End If
                    Next
                    For da As Integer = 0 To Rendez_Vous.Rows.Count - 1
                        If TextBox5.Text = Rendez_Vous.Rows(da).Item("Date_R") Then
                            dat_test = False
                        End If
                    Next
                    If (TextBox5.Text = Label15.Text) And (Me.Text = "Modifier Rendez_Vous") Then
                        dat_test = True
                    End If
                    If dat_test = True Then
                        If TextBox3.Text.Trim = "" Or ComboBox15.Text = "" Or TextBox2.Text.Trim = "" Or TextBox4.Text.Trim = "" Or numt.Text.Trim = "" Or nomt.Text.Trim = "" Or pret.Text.Trim = "" Then
                            If TextBox3.Text.Trim = "" Then
                                MsgBox(Title:="Error", Prompt:="Time Vide ...")
                            ElseIf ComboBox15.Text = "" Then
                                MsgBox(Title:="Error", Prompt:="Etat Vide ...")
                            ElseIf TextBox2.Text.Trim = "" Then
                                MsgBox(Title:="Error", Prompt:="Date incorrect ...")
                            ElseIf TextBox4.Text.Trim = "" Then
                                MsgBox(Title:="Error", Prompt:="Time Vide ...")
                            ElseIf numt.Text.Trim = "" Then
                                MsgBox(Title:="Error", Prompt:="Numero Vide ...")
                            ElseIf nomt.Text.Trim = "" Then
                                MsgBox(Title:="Error", Prompt:="Nom Vide ...")
                            ElseIf pret.Text.Trim = "" Then
                                MsgBox(Title:="Error", Prompt:="Prenom Vide ...")
                            End If
                        Else
                            If IsNumeric(numt.Text.Trim) Then
                                If (TextBox3.Text > 0 And TextBox3.Text < 24) And ((TextBox4.Text > -1) And (TextBox4.Text < 60)) And TextBox3.Text.Length = 2 And TextBox4.Text.Length = 2 Then
                                    If Me.Text = "Modifier Rendez_Vous" Then
                                        Dim test As Boolean = False
                                        For t As Integer = 0 To RendezVousLibre.Rows.Count - 1
                                            If (numt.Text = RendezVousLibre.Rows(t).Item("Numero")) Then
                                                test = True
                                            End If
                                        Next
                                        If numt.Text = Label10.Text Then
                                            test = False
                                        End If
                                        If test = True Then
                                            MsgBox("Numéro existe ...")
                                        Else
                                            RendezVousLibre_load()
                                            RendezVousLibre.Rows(offff).Item("Numero") = numt.Text
                                            RendezVousLibre.Rows(offff).Item("Nom") = nomt.Text
                                            RendezVousLibre.Rows(offff).Item("Prenom") = pret.Text
                                            RendezVousLibre.Rows(offff).Item("Date_R") = TextBox5.Text
                                            RendezVousLibre.Rows(offff).Item("Etat_R") = ComboBox15.Text


                                            RendezVousLibre_Save()
                                            RendezVousLibre_load()
                                            Gestion_des_rendez_vous.Show()
                                            Me.Close()
                                        End If
                                    Else
                                        Dim verif As Boolean = True
                                        For xx As Integer = 0 To RendezVousLibre.Rows.Count - 1
                                            If numt.Text = RendezVousLibre.Rows(xx).Item("Numero") Then
                                                verif = False
                                            End If
                                        Next
                                        If verif = False Then
                                            MsgBox("Numéro existe ...")
                                        Else
                                            RendezVousLibre_load()
                                            RendezVousLibre.Rows.Add(numt.Text, nomt.Text, pret.Text, TextBox5.Text, ComboBox15.Text)
                                            RendezVousLibre_Save()
                                            Gestion_des_rendez_vous.Show()
                                            Me.Close()
                                        End If
                                    End If
                                Else
                                    MsgBox(Title:="Error", Prompt:="Time incorrect ...")
                                End If
                            Else
                                MsgBox(Title:="Error", Prompt:="Numéro incorrect ...")
                            End If
                        End If
                    Else
                        MsgBox(Title:="Error", Prompt:="Date existe ...")
                    End If
                Else
                    MsgBox(Title:="Error", Prompt:="Date incorrect ...")
                End If
                  
                '***************************************************************************************************
            End If
        Catch
        End Try
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        Try
            Dim pp As Integer
            For ppp As Integer = 0 To Affichage_Patients.Rows.Count - 1
                If ComboBox3.Text = Affichage_Patients.Rows(ppp).Item("Numero") Then
                    pp = ppp
                End If
            Next
            Label5.Text = Affichage_Patients.Rows(pp).Item("Nom")
            Label4.Text = Affichage_Patients.Rows(pp).Item("Prenom")
        Catch
        End Try
    End Sub
End Class