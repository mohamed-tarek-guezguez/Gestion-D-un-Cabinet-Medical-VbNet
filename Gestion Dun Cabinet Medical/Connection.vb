Imports System.Data.OleDb

Module Connection

    Public t1 As Boolean = True
    Public conteur As Integer = 30
    Public openform As Boolean

    'Public connAccess2003 As New OleDbConnection("provider=Microsoft.jet.OLEDB.4.0;" & "data source=C:\DataBase\Database.mdb")
    Public connAccess2003 As New OleDbConnection("provider=Microsoft.jet.OLEDB.4.0;" & "data source=" & Application.StartupPath & "\Database.mdb")

    Public daAffichage_Medicaments As New OleDbDataAdapter
    Public Affichage_Medicaments As New DataTable

    Public daAffichage_Patients As New OleDbDataAdapter
    Public Affichage_Patients As New DataTable

    Public daAffichage_Rendezvous As New OleDbDataAdapter
    Public Affichage_Rendezvous As New DataTable


    '****************************************************************************
    Public daRendezVousLibre As New OleDbDataAdapter
    Public RendezVousLibre As New DataTable
    Public Sub RendezVousLibre_load()
        RendezVousLibre.Clear()
        daRendezVousLibre = New OleDbDataAdapter("select * from RendezVousLibre order by Date_R", connAccess2003)
        daRendezVousLibre.Fill(RendezVousLibre)
    End Sub
    Public Sub RendezVousLibre_Save()
        Dim Save_RendezVousLibre As OleDbCommandBuilder
        Save_RendezVousLibre = New OleDbCommandBuilder(daRendezVousLibre)
        daRendezVousLibre.Update(RendezVousLibre)
        RendezVousLibre.AcceptChanges()
    End Sub
    Public Sub Search_rendlib(text_s, text2)
        RendezVousLibre.Clear()
        If text2 = "Numéro" Then
            daRendezVousLibre = New OleDbDataAdapter("select * from RendezVousLibre where Numero like '%" & text_s & "%' order by Nom", connAccess2003)
        ElseIf text2 = "Nom" Then
            daRendezVousLibre = New OleDbDataAdapter("select * from RendezVousLibre where Nom like '%" & text_s & "%' order by Nom", connAccess2003)
        Else
            daRendezVousLibre = New OleDbDataAdapter("select * from RendezVousLibre where Date_R like '%" & text_s & "%' order by Nom", connAccess2003)
        End If
        daRendezVousLibre.Fill(RendezVousLibre)
    End Sub

    Public daRendezVousLibret As New OleDbDataAdapter
    Public RendezVousLibret As New DataTable
    Public Sub RendezVousLibret_load()
        RendezVousLibret.Clear()
        daRendezVousLibret = New OleDbDataAdapter("select * from RendezVousLibret order by Date_R", connAccess2003)
        daRendezVousLibret.Fill(RendezVousLibret)
    End Sub
    Public Sub RendezVousLibret_Save()
        Dim Save_RendezVousLibret As OleDbCommandBuilder
        Save_RendezVousLibret = New OleDbCommandBuilder(daRendezVousLibret)
        daRendezVousLibret.Update(RendezVousLibret)
        RendezVousLibret.AcceptChanges()
    End Sub
    '****************************************************************************
    Public daPatients As New OleDbDataAdapter
    Public Patients As New DataTable
    Public Sub Patients_load()
        Patients.Clear()
        daPatients = New OleDbDataAdapter("select * from Patients order by Nom", connAccess2003)
        daPatients.Fill(Patients)
    End Sub
    Public Sub Patients_Save()
        Dim Save_Patients As OleDbCommandBuilder
        Save_Patients = New OleDbCommandBuilder(daPatients)
        daPatients.Update(Patients)
        Patients.AcceptChanges()
    End Sub
    '****************************************************************************
    Public daRendez_Vous As New OleDbDataAdapter
    Public Rendez_Vous As New DataTable
    Public Sub Rendez_Vous_load()
        Rendez_Vous.Clear()
        daRendez_Vous = New OleDbDataAdapter("select * from Rendez_Vous order by Date_R", connAccess2003)
        daRendez_Vous.Fill(Rendez_Vous)
    End Sub
    Public Sub Rendez_Vous_Save()
        Dim Save_rendez As OleDbCommandBuilder
        Save_rendez = New OleDbCommandBuilder(daRendez_Vous)
        daRendez_Vous.Update(Rendez_Vous)
        Rendez_Vous.AcceptChanges()
    End Sub
    '***************************************************************************

    Public Sub Affichage_rendezvous_load()
        Affichage_Rendezvous.Clear()
        daAffichage_Rendezvous = New OleDbDataAdapter("select * from Affichage_Rendezvous order by Date_R", connAccess2003)
        daAffichage_Rendezvous.Fill(Affichage_Rendezvous)
    End Sub

    Public Sub Affichage_rendezvous_Save()
        Dim Save_rend As OleDbCommandBuilder
        Save_rend = New OleDbCommandBuilder(daAffichage_Rendezvous)
        daAffichage_Rendezvous.Update(Affichage_Rendezvous)
        Affichage_Rendezvous.AcceptChanges()
    End Sub

    Public Sub Search_rend(text_s, text2)
        Affichage_Rendezvous.Clear()
        If text2 = "Numéro" Then
            daAffichage_Rendezvous = New OleDbDataAdapter("select * from Affichage_Rendezvous where Numero like '%" & text_s & "%' order by Nom", connAccess2003)
        ElseIf text2 = "Nom" Then
            daAffichage_Rendezvous = New OleDbDataAdapter("select * from Affichage_Rendezvous where Nom like '%" & text_s & "%' order by Nom", connAccess2003)
        Else
            daAffichage_Rendezvous = New OleDbDataAdapter("select * from Affichage_Rendezvous where Date_R like '%" & text_s & "%' order by Nom", connAccess2003)
        End If
        daAffichage_Rendezvous.Fill(Affichage_Rendezvous)
    End Sub

    Public Sub Affichage_medic_load()
        Affichage_Medicaments.Clear()
        daAffichage_Medicaments = New OleDbDataAdapter("select * from Affichage_Medicaments order by nom", connAccess2003)
        daAffichage_Medicaments.Fill(Affichage_Medicaments)
    End Sub

    Public Sub Affichage_medic_Save()
        Dim Save_medic As OleDbCommandBuilder
        Save_medic = New OleDbCommandBuilder(daAffichage_Medicaments)
        daAffichage_Medicaments.Update(Affichage_Medicaments)
        Affichage_Medicaments.AcceptChanges()
    End Sub

    Public Sub Search(text_s, text2)
        Affichage_Medicaments.Clear()
        If text2 = "Code" Then
            daAffichage_Medicaments = New OleDbDataAdapter("select * from Affichage_Medicaments where Code like '%" & text_s & "%' order by Nom", connAccess2003)
        ElseIf text2 = "Nom" Then
            daAffichage_Medicaments = New OleDbDataAdapter("select * from Affichage_Medicaments where Nom like '%" & text_s & "%' order by Nom", connAccess2003)
        Else
            daAffichage_Medicaments = New OleDbDataAdapter("select * from Affichage_Medicaments where Quantite like '%" & text_s & "%' order by Nom", connAccess2003)
        End If
        daAffichage_Medicaments.Fill(Affichage_Medicaments)
    End Sub

    Public Sub Affichage_pat_load()
        Affichage_Patients.Clear()
        daAffichage_Patients = New OleDbDataAdapter("select * from Affichage_Patients order by Nom", connAccess2003)
        daAffichage_Patients.Fill(Affichage_Patients)
    End Sub

    Public Sub Affichage_pat_Save()
        Dim Save_pat As OleDbCommandBuilder
        Save_pat = New OleDbCommandBuilder(daAffichage_Patients)
        daAffichage_Patients.Update(Affichage_Patients)
        Affichage_Patients.AcceptChanges()
    End Sub

    Public Sub Search_pat(text_s, text2)
        Affichage_Patients.Clear()
        If text2 = "Numéro" Then
            daAffichage_Patients = New OleDbDataAdapter("select * from Affichage_Patients where Numero like '%" & text_s & "%' order by Nom", connAccess2003)
        ElseIf text2 = "Nom" Then
            daAffichage_Patients = New OleDbDataAdapter("select * from Affichage_Patients where Nom like '%" & text_s & "%' order by Nom", connAccess2003)
        ElseIf text2 = "Prenom" Then
            daAffichage_Patients = New OleDbDataAdapter("select * from Affichage_Patients where Prenom like '%" & text_s & "%' order by Nom", connAccess2003)
        Else
            daAffichage_Patients = New OleDbDataAdapter("select * from Affichage_Patients where lieu like '%" & text_s & "%' order by Nom", connAccess2003)
        End If
        daAffichage_Patients.Fill(Affichage_Patients)
    End Sub

End Module
