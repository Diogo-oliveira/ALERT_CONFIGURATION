﻿Imports Oracle.DataAccess.Client

Public Class Translation_Updates

    Dim translation As New Translation_API
    Dim db_access_general As New General

    '################################################################
    '#   Definição das variáves 
    '#   de área a atualizar
    Dim g_analysis_all As String = "1 - ANALYSIS (All content)"
    Dim g_analysis_cat As String = "    1.1 - ANALYSIS CATEGORY"
    Dim g_analysis_sample_type As String = "    1.2 - ANALYSIS SAMPLE TYPE"
    Dim g_analysis As String = "    1.3 - ANALYSIS"
    Dim g_sample_type As String = "    1.4 - SAMPLE TYPES"
    Dim g_analysis_parameters As String = "    1.5 - ANALYSS PARAMETERS"
    Dim g_analysis_recipient As String = "    1.6 - ANALYSIS SAMPLE RECIPIENTS"

    Dim g_exams_all As String = "2 - IMAGING AND OTHER EXAMS (All content)"
    Dim g_exam_categories As String = "    2.1 - EXAM CATEGORIES"
    Dim g_exams As String = "    2.2 - EXAMS"

    Dim g_procedures_all As String = "3 - INTERVENTIONS (All content)"
    Dim g_procedures_cat As String = "    3.1 - INTERVENTION CATEGORIES"
    Dim g_procedures As String = "    3.2 - INTERVENTIONS"

    Dim g_sr_intervs As String = "4 - SURGICAL INTERVENTIONS"

    Dim g_supplies_all As String = "5 - DUPPLIES (All content)"
    Dim g_supplies_cat As String = "    5.1 - SUPPLIES CATEGORIES"
    Dim g_supplies As String = "    5.2 - SUPPLIES"

    '#
    '##################################################################

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Cursor = Cursors.WaitCursor

        Dim l_column_width As Int64 = DataGridView1.Size.Width - 2 'Tentar evitar o scroll

        If (TextBox1.Text <> "" And ComboBox1.Text <> "") Then

            If (ComboBox2.Text = g_exams) Then

                If translation.CREATE_TEMP_TABLE() Then

                    If Not translation.UPDATE_EXAMS(TextBox1.Text) Then

                        MsgBox("ERROR UPDATING " & g_exams & " TRANSLATION!", vbCritical)

                    Else

                        Dim dr As OracleDataReader

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
                        If Not translation.GET_UPDATED_RECORDS(dr) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

                            MsgBox("ERROR GETTING UPDATED RECORDS!", vbCritical)

                        Else
                            DataGridView1.Columns.Clear()

                            Dim Table As New DataTable

                            Table.Load(dr)
                            DataGridView1.DataSource = Table

                            DataGridView1.Columns(0).Width = l_column_width
                            DataGridView1.Columns(0).SortMode = DataGridViewColumnSortMode.NotSortable

                        End If

                        If Not translation.DELETE_TEMP_TABLE() Then

                            MsgBox("ERROR DELETING TEMPORARY TABLE!", vbCritical)

                        End If

                    End If

                End If

            ElseIf (ComboBox2.Text = g_exam_categories) Then

                If translation.CREATE_TEMP_TABLE() Then

                    If Not translation.UPDATE_EXAM_CAT(TextBox1.Text) Then

                        MsgBox("ERROR UPDATING " & g_exam_categories & " TRANSLATION!", vbCritical)

                    Else

                        Dim dr As OracleDataReader

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
                        If Not translation.GET_UPDATED_RECORDS(dr) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

                            MsgBox("ERROR GETTING UPDATED RECORDS!", vbCritical)

                        Else
                            DataGridView1.Columns.Clear()

                            Dim Table As New DataTable

                            Table.Load(dr)
                            DataGridView1.DataSource = Table

                            DataGridView1.Columns(0).Width = l_column_width
                            DataGridView1.Columns(0).SortMode = DataGridViewColumnSortMode.NotSortable

                        End If

                        If Not translation.DELETE_TEMP_TABLE() Then

                            MsgBox("ERROR DELETING TEMPORARY TABLE!", vbCritical)

                        End If

                    End If

                End If

            ElseIf (ComboBox2.Text = g_procedures) Then

                If translation.CREATE_TEMP_TABLE() Then

                    If Not translation.UPDATE_INTERVENTIONS(TextBox1.Text) Then

                        MsgBox("ERROR UPDATING " & g_procedures & " TRANSLATION!", vbCritical)

                    Else

                        Dim dr As OracleDataReader

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
                        If Not translation.GET_UPDATED_RECORDS(dr) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

                            MsgBox("ERROR GETTING UPDATED RECORDS!", vbCritical)

                        Else
                            DataGridView1.Columns.Clear()

                            Dim Table As New DataTable

                            Table.Load(dr)
                            DataGridView1.DataSource = Table

                            DataGridView1.Columns(0).Width = l_column_width
                            DataGridView1.Columns(0).SortMode = DataGridViewColumnSortMode.NotSortable

                        End If

                        If Not translation.DELETE_TEMP_TABLE() Then

                            MsgBox("ERROR DELETING TEMPORARY TABLE!", vbCritical)

                        End If

                    End If

                End If

            ElseIf (ComboBox2.Text = g_procedures_cat) Then

                If translation.CREATE_TEMP_TABLE() Then

                    If Not translation.UPDATE_INTERV_CAT(TextBox1.Text) Then

                        MsgBox("ERROR UPDATING " & g_procedures_cat & " TRANSLATION!", vbCritical)

                    Else

                        Dim dr As OracleDataReader

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
                        If Not translation.GET_UPDATED_RECORDS(dr) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

                            MsgBox("ERROR GETTING UPDATED RECORDS!", vbCritical)

                        Else
                            DataGridView1.Columns.Clear()

                            Dim Table As New DataTable

                            Table.Load(dr)
                            DataGridView1.DataSource = Table

                            DataGridView1.Columns(0).Width = l_column_width
                            DataGridView1.Columns(0).SortMode = DataGridViewColumnSortMode.NotSortable

                        End If

                        If Not translation.DELETE_TEMP_TABLE() Then

                            MsgBox("ERROR DELETING TEMPORARY TABLE!", vbCritical)

                        End If

                    End If

                End If

            ElseIf (ComboBox2.Text = g_analysis) Then

                If translation.CREATE_TEMP_TABLE() Then

                    If Not translation.UPDATE_ANALYSIS(TextBox1.Text) Then

                        MsgBox("ERROR UPDATING " & g_analysis & " TRANSLATION!", vbCritical)

                    Else

                        Dim dr As OracleDataReader

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
                        If Not translation.GET_UPDATED_RECORDS(dr) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

                            MsgBox("ERROR GETTING UPDATED RECORDS!", vbCritical)

                        Else
                            DataGridView1.Columns.Clear()

                            Dim Table As New DataTable

                            Table.Load(dr)
                            DataGridView1.DataSource = Table

                            DataGridView1.Columns(0).Width = l_column_width
                            DataGridView1.Columns(0).SortMode = DataGridViewColumnSortMode.NotSortable

                        End If

                        If Not translation.DELETE_TEMP_TABLE() Then

                            MsgBox("ERROR DELETING TEMPORARY TABLE!", vbCritical)

                        End If

                    End If

                End If

            ElseIf (ComboBox2.Text = g_sample_type) Then

                If translation.CREATE_TEMP_TABLE() Then

                    If Not translation.UPDATE_SAMPLE_TYPE(TextBox1.Text) Then

                        MsgBox("ERROR UPDATING " & g_sample_type & " TRANSLATION!", vbCritical)

                    Else

                        Dim dr As OracleDataReader

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
                        If Not translation.GET_UPDATED_RECORDS(dr) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

                            MsgBox("ERROR GETTING UPDATED RECORDS!", vbCritical)

                        Else
                            DataGridView1.Columns.Clear()

                            Dim Table As New DataTable

                            Table.Load(dr)
                            DataGridView1.DataSource = Table

                            DataGridView1.Columns(0).Width = l_column_width
                            DataGridView1.Columns(0).SortMode = DataGridViewColumnSortMode.NotSortable

                        End If

                        If Not translation.DELETE_TEMP_TABLE() Then

                            MsgBox("ERROR DELETING TEMPORARY TABLE!", vbCritical)

                        End If

                    End If

                End If

            ElseIf (ComboBox2.Text = g_analysis_sample_type) Then

                If translation.CREATE_TEMP_TABLE() Then

                    If Not translation.UPDATE_ANALYSIS_SAMPLE_TYPE(TextBox1.Text) Then

                        MsgBox("ERROR UPDATING " & g_analysis_sample_type & " TRANSLATION!", vbCritical)

                    Else

                        Dim dr As OracleDataReader

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
                        If Not translation.GET_UPDATED_RECORDS(dr) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

                            MsgBox("ERROR GETTING UPDATED RECORDS!", vbCritical)

                        Else
                            DataGridView1.Columns.Clear()

                            Dim Table As New DataTable

                            Table.Load(dr)
                            DataGridView1.DataSource = Table

                            DataGridView1.Columns(0).Width = l_column_width
                            DataGridView1.Columns(0).SortMode = DataGridViewColumnSortMode.NotSortable

                        End If

                        If Not translation.DELETE_TEMP_TABLE() Then

                            MsgBox("ERROR DELETING TEMPORARY TABLE!", vbCritical)

                        End If

                    End If

                End If

            ElseIf (ComboBox2.Text = g_analysis_parameters) Then

                If translation.CREATE_TEMP_TABLE() Then

                    If Not translation.UPDATE_ANALYSIS_PARAMETERS(TextBox1.Text) Then

                        MsgBox("ERROR UPDATING " & g_analysis_parameters & " TRANSLATION!", vbCritical)

                    Else

                        Dim dr As OracleDataReader

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
                        If Not translation.GET_UPDATED_RECORDS(dr) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

                            MsgBox("ERROR GETTING UPDATED RECORDS!", vbCritical)

                        Else
                            DataGridView1.Columns.Clear()

                            Dim Table As New DataTable

                            Table.Load(dr)
                            DataGridView1.DataSource = Table

                            DataGridView1.Columns(0).Width = l_column_width
                            DataGridView1.Columns(0).SortMode = DataGridViewColumnSortMode.NotSortable

                        End If

                        If Not translation.DELETE_TEMP_TABLE() Then

                            MsgBox("ERROR DELETING TEMPORARY TABLE!", vbCritical)

                        End If

                    End If

                End If

            ElseIf (ComboBox2.Text = g_analysis_recipient) Then

                If translation.CREATE_TEMP_TABLE() Then

                    If Not translation.UPDATE_ANALYSIS_SR(TextBox1.Text) Then

                        MsgBox("ERROR UPDATING " & g_analysis_recipient & " TRANSLATION!", vbCritical)

                    Else

                        Dim dr As OracleDataReader

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
                        If Not translation.GET_UPDATED_RECORDS(dr) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

                            MsgBox("ERROR GETTING UPDATED RECORDS!", vbCritical)

                        Else
                            DataGridView1.Columns.Clear()

                            Dim Table As New DataTable

                            Table.Load(dr)
                            DataGridView1.DataSource = Table

                            DataGridView1.Columns(0).Width = l_column_width
                            DataGridView1.Columns(0).SortMode = DataGridViewColumnSortMode.NotSortable

                        End If

                        If Not translation.DELETE_TEMP_TABLE() Then

                            MsgBox("ERROR DELETING TEMPORARY TABLE!", vbCritical)

                        End If

                    End If

                End If

            ElseIf (ComboBox2.Text = g_analysis_cat) Then

                If translation.CREATE_TEMP_TABLE() Then

                    If Not translation.UPDATE_ANALYSIS_CAT(TextBox1.Text) Then

                        MsgBox("ERROR UPDATING " & g_analysis_cat & " TRANSLATION!", vbCritical)

                    Else

                        Dim dr As OracleDataReader

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
                        If Not translation.GET_UPDATED_RECORDS(dr) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

                            MsgBox("ERROR GETTING UPDATED RECORDS!", vbCritical)

                        Else
                            DataGridView1.Columns.Clear()

                            Dim Table As New DataTable

                            Table.Load(dr)
                            DataGridView1.DataSource = Table

                            DataGridView1.Columns(0).Width = l_column_width
                            DataGridView1.Columns(0).SortMode = DataGridViewColumnSortMode.NotSortable

                        End If

                        If Not translation.DELETE_TEMP_TABLE() Then

                            MsgBox("ERROR DELETING TEMPORARY TABLE!", vbCritical)

                        End If

                    End If

                End If

            ElseIf (ComboBox2.Text = g_sr_intervs) Then

                If translation.CREATE_TEMP_TABLE() Then

                    If Not translation.UPDATE_SR_INTERV(TextBox1.Text) Then

                        MsgBox("ERROR UPDATING " & g_sr_intervs & " TRANSLATION!", vbCritical)

                    Else

                        Dim dr As OracleDataReader

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
                        If Not translation.GET_UPDATED_RECORDS(dr) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

                            MsgBox("ERROR GETTING UPDATED RECORDS!", vbCritical)

                        Else
                            DataGridView1.Columns.Clear()

                            Dim Table As New DataTable

                            Table.Load(dr)
                            DataGridView1.DataSource = Table

                            DataGridView1.Columns(0).Width = l_column_width
                            DataGridView1.Columns(0).SortMode = DataGridViewColumnSortMode.NotSortable

                        End If

                        If Not translation.DELETE_TEMP_TABLE() Then

                            MsgBox("ERROR DELETING TEMPORARY TABLE!", vbCritical)

                        End If

                    End If

                End If

            ElseIf (ComboBox2.Text = g_supplies) Then

                If translation.CREATE_TEMP_TABLE() Then

                    If Not translation.UPDATE_SUPPLIES(TextBox1.Text) Then

                        MsgBox("ERROR UPDATING " & g_supplies & " TRANSLATION!", vbCritical)

                    Else

                        Dim dr As OracleDataReader

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
                        If Not translation.GET_UPDATED_RECORDS(dr) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

                            MsgBox("ERROR GETTING UPDATED RECORDS!", vbCritical)

                        Else
                            DataGridView1.Columns.Clear()

                            Dim Table As New DataTable

                            Table.Load(dr)
                            DataGridView1.DataSource = Table

                            DataGridView1.Columns(0).Width = l_column_width
                            DataGridView1.Columns(0).SortMode = DataGridViewColumnSortMode.NotSortable

                        End If

                        If Not translation.DELETE_TEMP_TABLE() Then

                            MsgBox("ERROR DELETING TEMPORARY TABLE!", vbCritical)

                        End If

                    End If

                End If

            ElseIf (ComboBox2.Text = g_supplies_cat) Then

                If translation.CREATE_TEMP_TABLE() Then

                    If Not translation.UPDATE_SUPPLIES_CAT(TextBox1.Text) Then

                        MsgBox("ERROR UPDATING " & g_supplies_cat & " TRANSLATION!", vbCritical)

                    Else

                        Dim dr As OracleDataReader

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
                        If Not translation.GET_UPDATED_RECORDS(dr) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

                            MsgBox("ERROR GETTING UPDATED RECORDS!", vbCritical)

                        Else
                            DataGridView1.Columns.Clear()

                            Dim Table As New DataTable

                            Table.Load(dr)
                            DataGridView1.DataSource = Table

                            DataGridView1.Columns(0).Width = l_column_width
                            DataGridView1.Columns(0).SortMode = DataGridViewColumnSortMode.NotSortable

                        End If

                        If Not translation.DELETE_TEMP_TABLE() Then

                            MsgBox("ERROR DELETING TEMPORARY TABLE!", vbCritical)

                        End If

                    End If

                End If

                '########################################################### Selecção de ALL ##############################################
            ElseIf (ComboBox2.Text = g_exams_all) Then

                If translation.CREATE_TEMP_TABLE() Then

                    If Not translation.UPDATE_EXAM_CAT(TextBox1.Text) Then

                        MsgBox("ERROR UPDATING " & g_exam_categories & " TRANSLATION!", vbCritical)

                    Else

                        Dim dr As OracleDataReader

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
                        If Not translation.GET_UPDATED_RECORDS(dr) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

                            MsgBox("ERROR GETTING UPDATED RECORDS!", vbCritical)

                        Else
                            DataGridView1.Columns.Clear()

                            Dim Table As New DataTable

                            Table.Load(dr)
                            DataGridView1.DataSource = Table

                            DataGridView1.Columns(0).Width = l_column_width
                            DataGridView1.Columns(0).SortMode = DataGridViewColumnSortMode.NotSortable

                        End If

                    End If

                    If Not translation.UPDATE_EXAMS(TextBox1.Text) Then

                        MsgBox("ERROR UPDATING " & g_exams & " TRANSLATION!", vbCritical)

                    Else

                        Dim dr As OracleDataReader

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
                        If Not translation.GET_UPDATED_RECORDS(dr) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

                            MsgBox("ERROR GETTING UPDATED RECORDS!", vbCritical)

                        Else
                            DataGridView1.Columns.Clear()

                            Dim Table As New DataTable

                            Table.Load(dr)
                            DataGridView1.DataSource = Table

                            DataGridView1.Columns(0).Width = l_column_width
                            DataGridView1.Columns(0).SortMode = DataGridViewColumnSortMode.NotSortable

                        End If

                    End If

                End If

                If Not translation.DELETE_TEMP_TABLE() Then

                    MsgBox("ERROR DELETING TEMPORARY TABLE!", vbCritical)

                End If

            ElseIf (ComboBox2.Text = g_analysis_all) Then

                If translation.CREATE_TEMP_TABLE() Then

                    If Not translation.UPDATE_ANALYSIS(TextBox1.Text) Then

                        MsgBox("ERROR UPDATING " & g_analysis & " TRANSLATION!", vbCritical)

                    Else

                        Dim dr As OracleDataReader

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
                        If Not translation.GET_UPDATED_RECORDS(dr) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

                            MsgBox("ERROR GETTING UPDATED RECORDS!", vbCritical)

                        Else
                            DataGridView1.Columns.Clear()

                            Dim Table As New DataTable

                            Table.Load(dr)
                            DataGridView1.DataSource = Table

                            DataGridView1.Columns(0).Width = l_column_width
                            DataGridView1.Columns(0).SortMode = DataGridViewColumnSortMode.NotSortable

                        End If

                    End If

                    If Not translation.UPDATE_SAMPLE_TYPE(TextBox1.Text) Then

                        MsgBox("ERROR UPDATING " & g_sample_type & " TRANSLATION!", vbCritical)

                    Else

                        Dim dr As OracleDataReader

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
                        If Not translation.GET_UPDATED_RECORDS(dr) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

                            MsgBox("ERROR GETTING UPDATED RECORDS!", vbCritical)

                        Else
                            DataGridView1.Columns.Clear()

                            Dim Table As New DataTable

                            Table.Load(dr)
                            DataGridView1.DataSource = Table

                            DataGridView1.Columns(0).Width = l_column_width
                            DataGridView1.Columns(0).SortMode = DataGridViewColumnSortMode.NotSortable

                        End If

                    End If

                    If Not translation.UPDATE_ANALYSIS_SAMPLE_TYPE(TextBox1.Text) Then

                        MsgBox("ERROR UPDATING " & g_analysis_sample_type & " TRANSLATION!", vbCritical)

                    Else

                        Dim dr As OracleDataReader

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
                        If Not translation.GET_UPDATED_RECORDS(dr) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

                            MsgBox("ERROR GETTING UPDATED RECORDS!", vbCritical)

                        Else
                            DataGridView1.Columns.Clear()

                            Dim Table As New DataTable

                            Table.Load(dr)
                            DataGridView1.DataSource = Table

                            DataGridView1.Columns(0).Width = l_column_width
                            DataGridView1.Columns(0).SortMode = DataGridViewColumnSortMode.NotSortable

                        End If

                    End If

                    If Not translation.UPDATE_ANALYSIS_CAT(TextBox1.Text) Then

                        MsgBox("ERROR UPDATING " & g_analysis_cat & " TRANSLATION!", vbCritical)

                    Else

                        Dim dr As OracleDataReader

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
                        If Not translation.GET_UPDATED_RECORDS(dr) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

                            MsgBox("ERROR GETTING UPDATED RECORDS!", vbCritical)

                        Else
                            DataGridView1.Columns.Clear()

                            Dim Table As New DataTable

                            Table.Load(dr)
                            DataGridView1.DataSource = Table

                            DataGridView1.Columns(0).Width = l_column_width
                            DataGridView1.Columns(0).SortMode = DataGridViewColumnSortMode.NotSortable

                        End If

                    End If

                    If Not translation.UPDATE_ANALYSIS_PARAMETERS(TextBox1.Text) Then

                        MsgBox("ERROR UPDATING " & g_analysis_parameters & " TRANSLATION!", vbCritical)

                    Else

                        Dim dr As OracleDataReader

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
                        If Not translation.GET_UPDATED_RECORDS(dr) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

                            MsgBox("ERROR GETTING UPDATED RECORDS!", vbCritical)

                        Else
                            DataGridView1.Columns.Clear()

                            Dim Table As New DataTable

                            Table.Load(dr)
                            DataGridView1.DataSource = Table

                            DataGridView1.Columns(0).Width = l_column_width
                            DataGridView1.Columns(0).SortMode = DataGridViewColumnSortMode.NotSortable

                        End If

                    End If

                    If Not translation.UPDATE_ANALYSIS_SR(TextBox1.Text) Then

                        MsgBox("ERROR UPDATING " & g_analysis_recipient & " TRANSLATION!", vbCritical)

                    Else

                        Dim dr As OracleDataReader

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
                        If Not translation.GET_UPDATED_RECORDS(dr) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

                            MsgBox("ERROR GETTING UPDATED RECORDS!", vbCritical)

                        Else
                            DataGridView1.Columns.Clear()

                            Dim Table As New DataTable

                            Table.Load(dr)
                            DataGridView1.DataSource = Table

                            DataGridView1.Columns(0).Width = l_column_width
                            DataGridView1.Columns(0).SortMode = DataGridViewColumnSortMode.NotSortable

                        End If

                    End If

                End If

                If Not translation.DELETE_TEMP_TABLE() Then

                    MsgBox("ERROR DELETING TEMPORARY TABLE!", vbCritical)

                End If

            ElseIf (ComboBox2.Text = g_procedures_all) Then

                If translation.CREATE_TEMP_TABLE() Then

                    If Not translation.UPDATE_INTERV_CAT(TextBox1.Text) Then

                        MsgBox("ERROR UPDATING " & g_procedures_cat & " TRANSLATION!", vbCritical)

                    Else

                        Dim dr As OracleDataReader

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
                        If Not translation.GET_UPDATED_RECORDS(dr) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

                            MsgBox("ERROR GETTING UPDATED RECORDS!", vbCritical)

                        Else
                            DataGridView1.Columns.Clear()

                            Dim Table As New DataTable

                            Table.Load(dr)
                            DataGridView1.DataSource = Table

                            DataGridView1.Columns(0).Width = l_column_width
                            DataGridView1.Columns(0).SortMode = DataGridViewColumnSortMode.NotSortable

                        End If

                    End If

                    If Not translation.UPDATE_INTERVENTIONS(TextBox1.Text) Then

                        MsgBox("ERROR UPDATING " & g_procedures & " TRANSLATION!", vbCritical)

                    Else

                        Dim dr As OracleDataReader

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
                        If Not translation.GET_UPDATED_RECORDS(dr) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

                            MsgBox("ERROR GETTING UPDATED RECORDS!", vbCritical)

                        Else
                            DataGridView1.Columns.Clear()

                            Dim Table As New DataTable

                            Table.Load(dr)
                            DataGridView1.DataSource = Table

                            DataGridView1.Columns(0).Width = l_column_width
                            DataGridView1.Columns(0).SortMode = DataGridViewColumnSortMode.NotSortable

                        End If

                    End If

                End If

                If Not translation.DELETE_TEMP_TABLE() Then

                    MsgBox("ERROR DELETING TEMPORARY TABLE!", vbCritical)

                End If

            ElseIf (ComboBox2.Text = g_supplies_all) Then

                If translation.CREATE_TEMP_TABLE() Then

                    If Not translation.UPDATE_SUPPLIES_CAT(TextBox1.Text) Then

                        MsgBox("ERROR UPDATING " & g_supplies_cat & " TRANSLATION!", vbCritical)

                    Else

                        Dim dr As OracleDataReader

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
                        If Not translation.GET_UPDATED_RECORDS(dr) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

                            MsgBox("ERROR GETTING UPDATED RECORDS!", vbCritical)

                        Else
                            DataGridView1.Columns.Clear()

                            Dim Table As New DataTable

                            Table.Load(dr)
                            DataGridView1.DataSource = Table

                            DataGridView1.Columns(0).Width = l_column_width
                            DataGridView1.Columns(0).SortMode = DataGridViewColumnSortMode.NotSortable

                        End If

                    End If

                    If Not translation.UPDATE_SUPPLIES(TextBox1.Text) Then

                        MsgBox("ERROR UPDATING " & g_supplies & " TRANSLATION!", vbCritical)

                    Else

                        Dim dr As OracleDataReader

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
                        If Not translation.GET_UPDATED_RECORDS(dr) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

                            MsgBox("ERROR GETTING UPDATED RECORDS!", vbCritical)

                        Else
                            DataGridView1.Columns.Clear()

                            Dim Table As New DataTable

                            Table.Load(dr)
                            DataGridView1.DataSource = Table

                            DataGridView1.Columns(0).Width = l_column_width
                            DataGridView1.Columns(0).SortMode = DataGridViewColumnSortMode.NotSortable

                        End If

                    End If

                End If

                If Not translation.DELETE_TEMP_TABLE() Then

                    MsgBox("ERROR DELETING TEMPORARY TABLE!", vbCritical)

                End If

            Else

                MsgBox("Please select an area to be updated.", vbInformation)

            End If

        Else

            MsgBox("Please select an institution.", vbInformation)

        End If
        Cursor = Cursors.Arrow

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Cursor = Cursors.WaitCursor

        If TextBox1.Text <> "" Then

            ComboBox1.Text = db_access_general.GET_INSTITUTION(TextBox1.Text)

        End If

        Cursor = Cursors.Arrow

    End Sub

    Private Sub Translation_Updates_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.Text = "TRANSLATION UPDATE  ::::  Connected to " & Connection.db

        Me.BackColor = Color.FromArgb(215, 215, 180)

        Me.Location = New Point(Form_location.x_position, Form_location.y_position)

        Dim dr As OracleDataReader

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
        If Not db_access_general.GET_ALL_INSTITUTIONS(dr) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

            MsgBox("ERROR GETTING ALL INSTITUTIONS!", vbCritical)

        Else

            While dr.Read()

                ComboBox1.Items.Add(dr.Item(0))

            End While

        End If

        dr.Dispose()
        dr.Close()


        '###############################################
        '# Bloco para inserir as categorias na BOX
        '# É aqui que se define a ordem de apresentação
        ComboBox2.Items.Add(g_analysis_all)
        ComboBox2.Items.Add(g_analysis_cat)
        ComboBox2.Items.Add(g_analysis_sample_type)
        ComboBox2.Items.Add(g_analysis)
        ComboBox2.Items.Add(g_sample_type)
        ComboBox2.Items.Add(g_analysis_parameters)
        ComboBox2.Items.Add(g_analysis_recipient)

        ComboBox2.Items.Add("")

        ComboBox2.Items.Add(g_exams_all)
        ComboBox2.Items.Add(g_exam_categories)
        ComboBox2.Items.Add(g_exams)

        ComboBox2.Items.Add("")

        ComboBox2.Items.Add(g_procedures_all)
        ComboBox2.Items.Add(g_procedures_cat)
        ComboBox2.Items.Add(g_procedures)

        ComboBox2.Items.Add("")

        ComboBox2.Items.Add(g_sr_intervs)

        ComboBox2.Items.Add("")

        ComboBox2.Items.Add(g_supplies_all)
        ComboBox2.Items.Add(g_supplies_cat)
        ComboBox2.Items.Add(g_supplies)
        '#
        '###############################################

        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        Form_location.x_position = Me.Location.X
        Form_location.y_position = Me.Location.Y

        Dim form1 As New Form1

        form1.Show()

        Me.Close()

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

        DataGridView1.DataSource = ""

    End Sub
End Class