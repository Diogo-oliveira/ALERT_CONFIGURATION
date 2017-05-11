Imports Oracle.DataAccess.Client

Public Class DISCHARGE_ADVANCED

    Dim db_access_general As New General
    Dim db_discharge As New DISCHARGE_API
    Dim db_clin_serv As New CLINICAL_SERVICE_API

    'Variável que guarda o sotware selecionado
    Dim g_selected_soft As Int16 = -1

    'Array que vai guardar as REASONS caregadas do default
    Dim g_a_loaded_reasons_default() As DISCHARGE_API.DEFAULT_REASONS

    'Array que vai guardar as DESTINATIONS caregadas do default
    Dim g_a_loaded_destinations_default() As DISCHARGE_API.DEFAULT_REASONS

    'Array com as flags dos tipos de profissionais
    Dim g_a_prof_types(5) As String

    'Array de profile templates disponíveis 
    Public Structure PROFILE_TEMPLATE
        Public ID_PROFILE_TEMPLATE As Integer
        Public PROFILE_NAME As String
        Public FLG_TYPE As String
    End Structure

    Dim g_a_profile_templates() As PROFILE_TEMPLATE

    Function reset_default_reasons()

        ReDim g_a_loaded_reasons_default(0)
        ComboBox3.Items.Clear()

        Dim dr_new As OracleDataReader

        If Not db_discharge.GET_ALL_DEFAULT_REASONS(TextBox1.Text, dr_new) Then

            MsgBox("ERROR GETING DEFAULT DISCHARGE REASONS.", vbCritical)

        Else

            Dim l_index_reason_default As Integer = 0
            ReDim g_a_loaded_reasons_default(0)

            While dr_new.Read()

                ReDim Preserve g_a_loaded_reasons_default(l_index_reason_default)
                g_a_loaded_reasons_default(l_index_reason_default).id_content = dr_new.Item(0)
                g_a_loaded_reasons_default(l_index_reason_default).desccription = dr_new.Item(1)
                l_index_reason_default = l_index_reason_default + 1

                ComboBox3.Items.Add(dr_new.Item(1) & "  -  [" & dr_new.Item(0) & "]")

            End While

        End If

    End Function

    Function reset_default_destinations()

        ReDim g_a_loaded_destinations_default(0)
        CheckedListBox1.Items.Clear()

        Dim dr_new As OracleDataReader

        If Not db_discharge.GET_ALL_DEFAULT_REASONS(TextBox1.Text, dr_new) Then

            MsgBox("ERROR GETING DEFAULT DISCHARGE DESTINATIONS.", vbCritical)

        Else

            Dim l_index_destinations_default As Integer = 0
            ReDim g_a_loaded_destinations_default(0)

            While dr_new.Read()

                ReDim Preserve g_a_loaded_destinations_default(l_index_destinations_default)
                g_a_loaded_destinations_default(l_index_destinations_default).id_content = dr_new.Item(0)
                g_a_loaded_destinations_default(l_index_destinations_default).desccription = dr_new.Item(1)
                l_index_destinations_default = l_index_destinations_default + 1

                CheckedListBox1.Items.Add(dr_new.Item(1) & "  -  [" & dr_new.Item(0) & "]")

            End While

        End If

    End Function

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

        Cursor = Cursors.Arrow

        If TextBox1.Text <> "" Then

            ComboBox1.Text = db_access_general.GET_INSTITUTION(TextBox1.Text)

            ComboBox2.Items.Clear()
            ComboBox2.Text = ""

            Dim dr As OracleDataReader

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
            If Not db_access_general.GET_SOFT_INST(TextBox1.Text, dr) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

                MsgBox("ERROR GETTING SOFTWARES!", vbCritical)

            Else

                While dr.Read()

                    ComboBox2.Items.Add(dr.Item(1))

                End While

            End If

            dr.Dispose()
            dr.Close()

        Else

            ComboBox1.Text = ""
            ComboBox2.Items.Clear()
            ComboBox2.Text = ""

        End If

        Cursor = Cursors.Arrow
    End Sub

    Private Sub DISCHARGE_ADVANCED_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.Text = "DISCHARGE  ::::  Connected to " & Connection.db

        Me.BackColor = Color.FromArgb(215, 215, 180)

        Me.Location = New Point(Form_location.x_position, Form_location.y_position)

        Dim dr As OracleDataReader

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
        If Not db_access_general.GET_ALL_INSTITUTIONS(dr) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

            MsgBox("ERROR GETTING ALL INSTITUTIONS!")

        Else

            While dr.Read()

                ComboBox1.Items.Add(dr.Item(0))

            End While

        End If

        dr.Dispose()
        dr.Close()

        'Tipos de profissionais
        'All
        'Physician
        'Nurse
        'Administrative
        'Other

        g_a_prof_types(0) = "-1"
        g_a_prof_types(1) = "D"
        g_a_prof_types(2) = "N"
        g_a_prof_types(3) = "A"
        g_a_prof_types(4) = "-2"

        ComboBox4.Items.Add("All")
        ComboBox4.Items.Add("Physician")
        ComboBox4.Items.Add("Nurse")
        ComboBox4.Items.Add("Administrative")
        ComboBox4.Items.Add("Other")

        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized

    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged

        Cursor = Cursors.WaitCursor

        g_selected_soft = db_access_general.GET_SELECTED_SOFT(ComboBox2.SelectedIndex, TextBox1.Text)

        reset_default_reasons()

        ReDim g_a_profile_templates(0)
        ComboBox4.SelectedIndex = -1
        CheckedListBox2.Items.Clear()

        Cursor = Cursors.Arrow

    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged

        reset_default_destinations()

    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged

        Cursor = Cursors.WaitCursor

        If ComboBox4.SelectedIndex > -1 Then


            ReDim g_a_profile_templates(0)
            CheckedListBox2.Items.Clear()

            Dim dr_new As OracleDataReader

            If Not db_access_general.GET_PROFILES(g_selected_soft, g_a_prof_types(ComboBox4.SelectedIndex), dr_new) Then

                MsgBox("ERROR GETING PROFILE TEMPLATES.", vbCritical)

            Else

                Dim l_index_prof_templates As Integer = 0
                ReDim g_a_profile_templates(0)

                While dr_new.Read()

                    ReDim Preserve g_a_profile_templates(l_index_prof_templates)
                    g_a_profile_templates(l_index_prof_templates).ID_PROFILE_TEMPLATE = dr_new.Item(0)
                    g_a_profile_templates(l_index_prof_templates).PROFILE_NAME = dr_new.Item(1)
                    g_a_profile_templates(l_index_prof_templates).FLG_TYPE = dr_new.Item(2)
                    l_index_prof_templates = l_index_prof_templates + 1

                    CheckedListBox2.Items.Add(dr_new.Item(0) & "  -  " & dr_new.Item(1))

                End While

            End If

        End If

        Cursor = Cursors.Arrow

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

        Cursor = Cursors.WaitCursor

        TextBox1.Text = db_access_general.GET_INSTITUTION_ID(ComboBox1.SelectedIndex)

        ComboBox2.Items.Clear()
        ComboBox2.Text = ""
        g_selected_soft = -1

        ReDim g_a_profile_templates(0)
        ComboBox4.SelectedIndex = -1
        CheckedListBox2.Items.Clear()

        Dim dr As OracleDataReader

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
        If Not db_access_general.GET_SOFT_INST(TextBox1.Text, dr) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

            MsgBox("ERROR GETTING SOFTWARES!", vbCritical)

        Else

            While dr.Read()

                ComboBox2.Items.Add(dr.Item(1))

            End While

        End If

        dr.Dispose()
        dr.Close()

        Cursor = Cursors.Arrow

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click



    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

        'Código para ver se rank introduzido está correto
        Dim l_correct_rank As Boolean = True
        If TextBox2.Text <> "" Then

            '48 - 57  = Ascii codes for numbers
            For i As Integer = 0 To TextBox2.Text.Length() - 1

                If Asc(TextBox2.Text.Chars(i)) < 48 Or Asc(TextBox2.Text.Chars(i)) > 57 Then

                    l_correct_rank = False

                End If

            Next

        End If

        If l_correct_rank = False Then

            MsgBox("INCORRECT RANK")

        End If

    End Sub
End Class