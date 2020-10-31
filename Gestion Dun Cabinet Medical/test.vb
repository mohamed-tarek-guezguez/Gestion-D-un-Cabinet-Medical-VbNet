Public Class test
    Dim t2 As Boolean = False

    Private Sub Timer5_Tick(sender As Object, e As EventArgs) Handles Timer5.Tick
        Try
            Rendez_Vous_load()
            For cp As Integer = 0 To Rendez_Vous.Rows.Count - 1
                Rendez_Vous_Save()
                Dim jj As Integer = CInt(System.DateTime.Today.Day)
                Dim mm As Integer = CInt(System.DateTime.Today.Month)
                Dim aa As Integer = CInt(System.DateTime.Today.Year)
                Dim tim As String = Format(TimeOfDay)
                Dim tests As String
                tests = tim.ToCharArray(0, 2)
                Dim he As Integer = CInt(tests)
                tests = tim.ToCharArray(3, 2)
                Dim mi As Integer = CInt(tests)
                '***************************************************************************************************
                Gestion_des_rendez_vous.Label6.Text = Rendez_Vous.Rows(cp).Item("Date_R").ToString
                Dim test As String = Gestion_des_rendez_vous.Label6.Text.ToCharArray(0, 2)
                Dim jjj As Integer = CInt(test)
                test = Gestion_des_rendez_vous.Label6.Text.ToCharArray(3, 2)
                Dim mmm As Integer = CInt(test)
                test = Gestion_des_rendez_vous.Label6.Text.ToCharArray(6, 4)
                Dim aaa As Integer = CInt(test)
                test = Gestion_des_rendez_vous.Label6.Text.ToCharArray(11, 2)
                Dim hehe As Integer = CInt(test)
                test = Gestion_des_rendez_vous.Label6.Text.ToCharArray(14, 2)
                Dim mimi As Integer = CInt(test)

                Dim pre As String = Rendez_Vous.Rows(cp).Item("Etat_R")
                If (pre.ToString.ToUpper = "OUI") Or (pre.ToString.ToUpper = "NON") Then
                    If (aa > aaa) Or ((aa = aaa) And (mm > mmm)) Or ((aa = aaa) And (mm = mmm) And (jj > jjj)) Or ((aa = aaa) And (mm = mmm) And (jj = jjj) And (he > hehe)) Or ((aa = aaa) And (mm = mmm) And (jj = jjj) And (he = hehe) And (mi > mimi)) Then
                        pre = pre + "_Passer"
                        Rendez_Vous.Rows(cp).Item("Etat_R") = pre

                        Rendez_Vous_Save()
                        Rendez_Vous_load()
                        Affichage_rendezvous_load()
                        Gestion_des_rendez_vous.DataGridView1.DataSource = Affichage_Rendezvous
                        datgridview_setting_rendezvous()
                    End If
                End If
                If jj = jjj And mm = mmm And aa = aaa And he = hehe And mimi >= 31 And mimi - 31 = mi Then
                    Timer6.Start()
                    Timer7.Start()
                End If
                If jj = jjj And mm = mmm And aa = aaa And he = hehe - 1 And mimi < 31 And mimi + 60 - 31 = mi Then
                    Timer6.Start()
                    Timer7.Start()
                End If
            Next
            '***************************************************************************************************
            RendezVousLibret_load()
            For cp As Integer = 0 To RendezVousLibret.Rows.Count - 1
                RendezVousLibret_Save()
                Dim jj As Integer = CInt(System.DateTime.Today.Day)
                Dim mm As Integer = CInt(System.DateTime.Today.Month)
                Dim aa As Integer = CInt(System.DateTime.Today.Year)
                Dim tim As String = Format(TimeOfDay)
                Dim tests As String
                tests = tim.ToCharArray(0, 2)
                Dim he As Integer = CInt(tests)
                tests = tim.ToCharArray(3, 2)
                Dim mi As Integer = CInt(tests)
                '***************************************************************************************************
                Gestion_des_rendez_vous.Label6.Text = RendezVousLibret.Rows(cp).Item("Date_R").ToString
                Dim test As String = Gestion_des_rendez_vous.Label6.Text.ToCharArray(0, 2)
                Dim jjj As Integer = CInt(test)
                test = Gestion_des_rendez_vous.Label6.Text.ToCharArray(3, 2)
                Dim mmm As Integer = CInt(test)
                test = Gestion_des_rendez_vous.Label6.Text.ToCharArray(6, 4)
                Dim aaa As Integer = CInt(test)
                test = Gestion_des_rendez_vous.Label6.Text.ToCharArray(11, 2)
                Dim hehe As Integer = CInt(test)
                test = Gestion_des_rendez_vous.Label6.Text.ToCharArray(14, 2)
                Dim mimi As Integer = CInt(test)

                Dim pre As String = RendezVousLibret.Rows(cp).Item("Etat_R")
                If (pre.ToString.ToUpper = "OUI") Or (pre.ToString.ToUpper = "NON") Then
                    If (aa > aaa) Or ((aa = aaa) And (mm > mmm)) Or ((aa = aaa) And (mm = mmm) And (jj > jjj)) Or ((aa = aaa) And (mm = mmm) And (jj = jjj) And (he > hehe)) Or ((aa = aaa) And (mm = mmm) And (jj = jjj) And (he = hehe) And (mi > mimi)) Then
                        pre = pre + "_Passer"
                        RendezVousLibret.Rows(cp).Item("Etat_R") = pre

                        RendezVousLibret_Save()
                        RendezVousLibret_load()
                        RendezVousLibre_load()
                    End If
                End If
                If jj = jjj And mm = mmm And aa = aaa And he = hehe And mimi >= 31 And mimi - 31 = mi Then
                    Timer3.Start()
                    Timer7.Start()
                End If
                If jj = jjj And mm = mmm And aa = aaa And he = hehe - 1 And mimi < 31 And mimi + 60 - 31 = mi Then
                    Timer3.Start()
                    Timer7.Start()
                End If
            Next
        Catch ex As Exception
        End Try
    End Sub
    Private Sub Timer6_Tick(sender As Object, e As EventArgs) Handles Timer6.Tick
        Try
            Timer6.Stop()
            Affichage_rendezvous_load()

            Dim i As Integer = 0
            Do While (Affichage_Rendezvous.Rows(i).Item("Etat_R").ToString.ToUpper <> "OUI")
                i = i + 1
            Loop

            Notification.Label5.Text = Affichage_Rendezvous.Rows(i).Item("Numero")
            Notification.Label4.Text = Affichage_Rendezvous.Rows(i).Item("Prenom") + " " + Affichage_Rendezvous.Rows(i).Item("Nom")
            Notification.Label6.Text = Affichage_Rendezvous.Rows(i).Item("Date_R")

            Notification.Show()
            t2 = True
            conteur = 30
            StartForm.Label3.Text = conteur

            Timer2.Start()
        Catch
        End Try
    End Sub

    Private Sub Timer7_Tick(sender As Object, e As EventArgs) Handles Timer7.Tick
        Try
            conteur = conteur - 1
        Catch
        End Try
    End Sub

    Private Sub test_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        StartForm.Show()
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Try
            If StartForm.Visible = False And Gestion_des_rendez_vous.Visible = False And Gestion_des_patients.Visible = False And Gestion_des_médicaments.Visible = False And Ajouter_Medicaments.Visible = False And Ajouter_Patient.Visible = False And Ajouter_RendezVous.Visible = False Then
                End
            End If

            If conteur >= 1 And t2 = True Then
                StartForm.Label3.Visible = True
                StartForm.Label4.Visible = True
                StartForm.Label3.Text = conteur
            End If

            If CInt(StartForm.Label3.Text) = 1 Then
                StartForm.Label3.Visible = False
                StartForm.Label4.Visible = False
                t2 = False
                Timer7.Stop()
            End If

        Catch
        End Try
    End Sub

    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        Try
            Timer2.Stop()
            My.Computer.Audio.Play(My.Resources.Rendez_Vous, AudioPlayMode.WaitToComplete)
        Catch
        End Try
    End Sub

    Private Sub Timer3_Tick(sender As Object, e As EventArgs) Handles Timer3.Tick
        Try
            Timer3.Stop()
            RendezVousLibre_load()

            Dim i As Integer = 0
            Do While (RendezVousLibre.Rows(i).Item("Etat_R").ToString.ToUpper <> "OUI")
                i = i + 1
            Loop

            Notification.Label5.Text = RendezVousLibre.Rows(i).Item("Numero")
            Notification.Label4.Text = RendezVousLibre.Rows(i).Item("Prenom") + " " + RendezVousLibre.Rows(i).Item("Nom")
            Notification.Label6.Text = RendezVousLibre.Rows(i).Item("Date_R")

            Notification.Show()
            t2 = True
            conteur = 30
            StartForm.Label3.Text = conteur

            Timer2.Start()
        Catch
        End Try
    End Sub
End Class