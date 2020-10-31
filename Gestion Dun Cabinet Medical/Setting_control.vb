Module Setting_control

    Public Sub datgridview_setting()
        Gestion_des_médicaments.DataGridView1.Columns(0).Width = 155
        Gestion_des_médicaments.DataGridView1.Columns(0).HeaderText = "Code"
        Gestion_des_médicaments.DataGridView1.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Gestion_des_médicaments.DataGridView1.Columns(1).Width = 255
        Gestion_des_médicaments.DataGridView1.Columns(1).HeaderText = "Nom"
        Gestion_des_médicaments.DataGridView1.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Gestion_des_médicaments.DataGridView1.Columns(2).Width = 254
        Gestion_des_médicaments.DataGridView1.Columns(2).HeaderText = "Quantité"
        Gestion_des_médicaments.DataGridView1.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Gestion_des_médicaments.DataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
    End Sub

    Public Sub datgridview_setting_patients()
        Gestion_des_patients.DataGridView1.Columns(0).Width = 100
        Gestion_des_patients.DataGridView1.Columns(0).HeaderText = "Numéro"
        Gestion_des_patients.DataGridView1.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Gestion_des_patients.DataGridView1.Columns(1).Width = 150
        Gestion_des_patients.DataGridView1.Columns(1).HeaderText = "Nom"
        Gestion_des_patients.DataGridView1.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Gestion_des_patients.DataGridView1.Columns(2).Width = 150
        Gestion_des_patients.DataGridView1.Columns(2).HeaderText = "Prénom"
        Gestion_des_patients.DataGridView1.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Gestion_des_patients.DataGridView1.Columns(3).Width = 200
        Gestion_des_patients.DataGridView1.Columns(3).HeaderText = "Adresse"
        Gestion_des_patients.DataGridView1.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Gestion_des_patients.DataGridView1.Columns(4).Width = 200
        Gestion_des_patients.DataGridView1.Columns(4).HeaderText = "Date de naissance"
        Gestion_des_patients.DataGridView1.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Gestion_des_patients.DataGridView1.Columns(5).Width = 150
        Gestion_des_patients.DataGridView1.Columns(5).HeaderText = "Lieu de naissance"
        Gestion_des_patients.DataGridView1.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Gestion_des_patients.DataGridView1.Columns(6).Width = 120
        Gestion_des_patients.DataGridView1.Columns(6).HeaderText = "Etat Civile"
        Gestion_des_patients.DataGridView1.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Gestion_des_patients.DataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
    End Sub

    Public Sub datgridview_setting_rendezvous()
        Gestion_des_rendez_vous.DataGridView1.Columns(0).Width = 100
        Gestion_des_rendez_vous.DataGridView1.Columns(0).HeaderText = "Numero"
        Gestion_des_rendez_vous.DataGridView1.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Gestion_des_rendez_vous.DataGridView1.Columns(1).Width = 158
        Gestion_des_rendez_vous.DataGridView1.Columns(1).HeaderText = "Nom"
        Gestion_des_rendez_vous.DataGridView1.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Gestion_des_rendez_vous.DataGridView1.Columns(2).Width = 158
        Gestion_des_rendez_vous.DataGridView1.Columns(2).HeaderText = "Prenom"
        Gestion_des_rendez_vous.DataGridView1.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Gestion_des_rendez_vous.DataGridView1.Columns(3).Width = 150
        Gestion_des_rendez_vous.DataGridView1.Columns(3).HeaderText = "Date_R"
        Gestion_des_rendez_vous.DataGridView1.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Gestion_des_rendez_vous.DataGridView1.Columns(4).Width = 98
        Gestion_des_rendez_vous.DataGridView1.Columns(4).HeaderText = "Etat_R"
        Gestion_des_rendez_vous.DataGridView1.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Gestion_des_rendez_vous.DataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
    End Sub

End Module
