Public Class StartForm

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            Gestion_des_patients.Show()
            Me.Close()
        Catch
        End Try
    End Sub

    Private Sub Button2_Click_1(sender As Object, e As EventArgs) Handles Button2.Click
        Try
            openform = True
            Gestion_des_rendez_vous.Label1.Text = "Gestion des rendez-vous"
            Gestion_des_rendez_vous.Show()
            Me.Close()
        Catch
        End Try
    End Sub

    Private Sub Button3_Click_1(sender As Object, e As EventArgs) Handles Button3.Click
        Try
            Gestion_des_médicaments.Show()
            Me.Close()
        Catch
        End Try
    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click
        Try
            MsgBox(Title:="Info", Prompt:="This Program Created By Mohamed Tarek GuezGuez")
        Catch
        End Try
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Try
            Timer1.Stop()
            Label2.ForeColor = Color.Black
            Timer2.Start()
        Catch
        End Try
    End Sub

    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        Try
            Timer2.Stop()
            Label2.ForeColor = Color.White
            Timer1.Start()
        Catch
        End Try
    End Sub

    

    Private Sub StartForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Timer6.Start()
    End Sub

    Private Sub Timer3_Tick(sender As Object, e As EventArgs) Handles Timer3.Tick
        Try
            Timer3.Stop()
            PictureBox2.Visible = True
            Timer4.Start()
        Catch
        End Try
    End Sub

    Private Sub Timer4_Tick(sender As Object, e As EventArgs) Handles Timer4.Tick
        Try
            Timer4.Stop()
            PictureBox2.Visible = False
        Catch
        End Try
    End Sub

    Private Sub PictureBox3_Click(sender As Object, e As EventArgs) Handles PictureBox3.Click
        Try
            Affichage_rendezvous_load()
            RendezVousLibre_load()

            Dim ok1 As Boolean = False
            Dim ok2 As Boolean = False
            Dim i As Integer = -1
            Dim k As Integer = -1
            If Affichage_Rendezvous.Rows.Count = 0 And RendezVousLibre.Rows.Count > 0 Then
                Do
                    k = k + 1
                Loop Until ((RendezVousLibre.Rows(k).Item("Etat_R").ToString.ToUpper = "OUI") Or (RendezVousLibre.Rows.Count = k + 1))
                If RendezVousLibre.Rows(k).Item("Etat_R").ToString.ToUpper = "OUI" Then
                    Notification.Label5.Text = RendezVousLibre.Rows(k).Item("Numero")
                    Notification.Label4.Text = RendezVousLibre.Rows(k).Item("Prenom") + " " + RendezVousLibre.Rows(k).Item("Nom")
                    Notification.Label6.Text = RendezVousLibre.Rows(k).Item("Date_R")
                    Notification.Show()
                End If
            End If
            If RendezVousLibre.Rows.Count = 0 And Affichage_Rendezvous.Rows.Count > 0 Then
                Do
                    i = i + 1
                Loop Until ((Affichage_Rendezvous.Rows(i).Item("Etat_R").ToString.ToUpper = "OUI") Or (Affichage_Rendezvous.Rows.Count = i + 1))
                If Affichage_Rendezvous.Rows(i).Item("Etat_R").ToString.ToUpper = "OUI" Then
                    Notification.Label5.Text = Affichage_Rendezvous.Rows(i).Item("Numero")
                    Notification.Label4.Text = Affichage_Rendezvous.Rows(i).Item("Prenom") + " " + Affichage_Rendezvous.Rows(i).Item("Nom")
                    Notification.Label6.Text = Affichage_Rendezvous.Rows(i).Item("Date_R")
                    Notification.Show()
                End If
            End If

            If RendezVousLibre.Rows.Count > 0 And Affichage_Rendezvous.Rows.Count > 0 Then
                i = -1
                Do
                    i = i + 1
                Loop Until ((Affichage_Rendezvous.Rows(i).Item("Etat_R").ToString.ToUpper = "OUI") Or (Affichage_Rendezvous.Rows.Count = i + 1))
                Dim chtest As String
                Dim test As String
                Dim jjj As Integer
                Dim mmm As Integer
                Dim aaa As Integer
                Dim hehe As Integer
                Dim mimi As Integer
                If Affichage_Rendezvous.Rows(i).Item("Etat_R").ToString.ToUpper = "OUI" Then
                    ok1 = True
                    chtest = Affichage_Rendezvous.Rows(i).Item("Date_R").ToString.Trim
                    test = chtest.ToCharArray(0, 2)
                    jjj = CInt(test)
                    test = chtest.ToCharArray(3, 2)
                    mmm = CInt(test)
                    test = chtest.ToCharArray(6, 4)
                    aaa = CInt(test)
                    test = chtest.ToCharArray(11, 2)
                    hehe = CInt(test)
                    test = chtest.ToCharArray(14, 2)
                    mimi = CInt(test)
                End If

                k = -1
                Do
                    k = k + 1
                Loop Until ((RendezVousLibre.Rows(k).Item("Etat_R").ToString.ToUpper = "OUI") Or (RendezVousLibre.Rows.Count = k + 1))
                Dim jj As Integer
                Dim mm As Integer
                Dim aa As Integer
                Dim he As Integer
                Dim mi As Integer
                If RendezVousLibre.Rows(k).Item("Etat_R").ToString.ToUpper = "OUI" Then
                    ok2 = True
                    chtest = RendezVousLibre.Rows(k).Item("Date_R").ToString.Trim
                    test = chtest.ToCharArray(0, 2)
                    jj = CInt(test)
                    test = chtest.ToCharArray(3, 2)
                    mm = CInt(test)
                    test = chtest.ToCharArray(6, 4)
                    aa = CInt(test)
                    test = chtest.ToCharArray(11, 2)
                    he = CInt(test)
                    test = chtest.ToCharArray(14, 2)
                    mi = CInt(test)
                End If

                If ok1 = True And ok2 = False Then
                    Notification.Label5.Text = Affichage_Rendezvous.Rows(i).Item("Numero")
                    Notification.Label4.Text = Affichage_Rendezvous.Rows(i).Item("Prenom") + " " + Affichage_Rendezvous.Rows(i).Item("Nom")
                    Notification.Label6.Text = Affichage_Rendezvous.Rows(i).Item("Date_R")
                    Notification.Show()
                End If

                If ok1 = False And ok2 = True Then
                    Notification.Label5.Text = RendezVousLibre.Rows(k).Item("Numero")
                    Notification.Label4.Text = RendezVousLibre.Rows(k).Item("Prenom") + " " + RendezVousLibre.Rows(k).Item("Nom")
                    Notification.Label6.Text = RendezVousLibre.Rows(k).Item("Date_R")
                    Notification.Show()
                End If

                If ok1 = True And ok2 = True Then
                    If (aaa > aa) Or (aaa = aa And mmm > mm) Or (aaa = aa And mmm = mm And jjj > jj) Or (aaa = aa And mmm = mm And jjj = jj And hehe > he) Or (aaa = aa And mmm = mm And jjj = jj And hehe = he And mimi > mi) Then
                        Notification.Label5.Text = RendezVousLibre.Rows(k).Item("Numero")
                        Notification.Label4.Text = RendezVousLibre.Rows(k).Item("Prenom") + " " + RendezVousLibre.Rows(k).Item("Nom")
                        Notification.Label6.Text = RendezVousLibre.Rows(k).Item("Date_R")
                        Notification.Show()
                    Else
                        Notification.Label5.Text = Affichage_Rendezvous.Rows(i).Item("Numero")
                        Notification.Label4.Text = Affichage_Rendezvous.Rows(i).Item("Prenom") + " " + Affichage_Rendezvous.Rows(i).Item("Nom")
                        Notification.Label6.Text = Affichage_Rendezvous.Rows(i).Item("Date_R")
                        Notification.Show()
                    End If
                End If
            End If
        Catch
        End Try
    End Sub

    Private Sub Timer6_Tick(sender As Object, e As EventArgs) Handles Timer6.Tick
        Try
            Timer6.Stop()
            If t1 = True Then
                t1 = False
                My.Computer.Audio.Play(My.Resources.Bienvenue, AudioPlayMode.WaitToComplete)
            End If
        Catch
        End Try
    End Sub

    
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Try
            openform = False
            Gestion_des_rendez_vous.Label1.Text = "Gestion des rendez-vous Libres"
            Gestion_des_rendez_vous.Show()
            Me.Close()
        Catch
        End Try
    End Sub

End Class
