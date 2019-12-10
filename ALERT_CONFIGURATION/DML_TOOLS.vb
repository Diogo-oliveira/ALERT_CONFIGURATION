Imports Oracle.DataAccess.Client
Public Class DML_TOOLS

    Dim dml As New DML_API

    Dim db_access_general As New General

    Public Structure dml_struct
        Public dml_identifier As String
        Public dml_content As String
    End Structure

    Public Structure audit_struct
        Public audit_identifier As String
        Public pos As Int32
    End Structure

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Cursor = Cursors.WaitCursor

        RichTextBox2.Clear()

        Dim l_String_insert As String
        Dim l_String_values As String
        Dim l_tag As String = ""

        Dim l_processed_string As String = ""

        Dim l_insert As Boolean = False

        Dim l_index As Int32
        Dim l_count As Int32
        Dim l_dml_identifier As String
        Dim l_dml_struct() As dml_struct

        Dim l_ok As Boolean = True

        Dim l_pk_search As Boolean = False
        Dim l_pk_string As String = ""

        Dim audit_tags(5) As String
        audit_tags(0) = "CREATE_USER"
        audit_tags(1) = "CREATE_TIME"
        audit_tags(2) = "CREATE_INSTITUTION"
        audit_tags(3) = "UPDATE_USER"
        audit_tags(4) = "UPDATE_TIME"
        audit_tags(5) = "UPDATE_INSTITUTION"

        RichTextBox2.Text = ""

        Try
            For Each strLine As String In RichTextBox1.Text.Split(vbNewLine.ToCharArray, StringSplitOptions.RemoveEmptyEntries)

                If l_pk_search = True Then

                    l_pk_string = l_pk_string + strLine

                    If strLine.ToLower Like "*);*" Then

                        If l_insert = False Then
                            l_String_insert = l_pk_string

                            l_dml_identifier = dml.GET_DML_IDENTIFIER_PK(l_String_insert)
                            l_index = dml.GET_DML_INDEX(l_dml_identifier, l_dml_struct)

                            Try
                                l_count = l_dml_struct.Count
                            Catch ex As Exception
                                l_count = 0
                            End Try

                            If l_index = -1 Then

                                ReDim Preserve l_dml_struct(l_count)
                                l_dml_struct(l_count).dml_identifier = l_dml_identifier
                                l_dml_struct(l_count).dml_content = "   " + l_String_insert

                                l_dml_identifier = ""
                            Else
                                l_dml_struct(l_index).dml_identifier = l_dml_identifier
                                l_dml_struct(l_index).dml_content = l_dml_struct(l_index).dml_content + vbNewLine + "   " + l_String_insert
                                l_dml_identifier = ""
                            End If
                        Else
                            MsgBox("Error", vbCritical)
                            l_ok = False
                            Exit For
                        End If

                        l_pk_search = False
                    End If
                ElseIf strLine.ToLower Like "*insert_into_sys_message*" Or strLine.ToLower Like "*insert_into_translation*" Or
         strLine.ToLower Like "*insert_into_sys_domain*" Or strLine.ToLower Like "*insert_into_functionality_help*" Then


                    If strLine.ToLower Like "*);*" Then
                        If l_insert = False Then
                            l_String_insert = strLine

                            l_dml_identifier = dml.GET_DML_IDENTIFIER_PK(l_String_insert)
                            l_index = dml.GET_DML_INDEX(l_dml_identifier, l_dml_struct)

                            Try
                                l_count = l_dml_struct.Count
                            Catch ex As Exception
                                l_count = 0
                            End Try

                            If l_index = -1 Then

                                ReDim Preserve l_dml_struct(l_count)
                                l_dml_struct(l_count).dml_identifier = l_dml_identifier
                                l_dml_struct(l_count).dml_content = "   " + l_String_insert

                                l_dml_identifier = ""
                            Else
                                l_dml_struct(l_index).dml_identifier = l_dml_identifier
                                l_dml_struct(l_index).dml_content = l_dml_struct(l_index).dml_content + vbNewLine + "   " + l_String_insert
                                l_dml_identifier = ""
                            End If
                        Else
                            MsgBox("Error", vbCritical)
                            l_ok = False
                            Exit For
                        End If

                    Else
                        l_pk_search = True
                        l_pk_string = strLine
                    End If

                ElseIf strLine.ToUpper Like "*INSERT*" Then

                    If l_insert = False Then
                        l_String_insert = strLine
                    Else
                        MsgBox("Error", vbCritical)
                        l_ok = False
                        Exit For
                    End If
                    l_insert = True
                ElseIf strLine.ToUpper Like "*VALUES*" And l_insert = True Then
                    l_String_values = strLine

                    If Not dml.REMOVE_AUDIT(l_String_insert, l_String_values, audit_tags) Then
                        MsgBox("Error removing audit columns")
                    End If

                    l_processed_string = dml.GET_INSERT_PROCESSED(l_String_insert, l_String_values)

                    l_dml_identifier = dml.GET_DML_IDENTIFIER(l_String_insert)
                    l_index = dml.GET_DML_INDEX(l_dml_identifier, l_dml_struct)

                    Try
                        l_count = l_dml_struct.Count
                    Catch ex As Exception
                        l_count = 0
                    End Try

                    If l_index = -1 Then

                        ReDim Preserve l_dml_struct(l_count)
                        l_dml_struct(l_count).dml_identifier = l_dml_identifier
                        l_dml_struct(l_count).dml_content = l_processed_string

                        l_dml_identifier = ""

                    Else
                        l_dml_struct(l_index).dml_identifier = l_dml_identifier
                        l_dml_struct(l_index).dml_content = l_dml_struct(l_index).dml_content + vbNewLine + l_processed_string
                        l_dml_identifier = ""
                    End If

                    l_insert = False
                End If
            Next

            If l_ok = True Then
                For i As Integer = 0 To l_dml_struct.Count - 1
                    If i < l_dml_struct.Count - 1 Then
                        If l_dml_struct(i).dml_identifier.Contains("translation") Or l_dml_struct(i).dml_identifier.Contains("sys_message") Or
                        l_dml_struct(i).dml_identifier.Contains("sys_domain") Or l_dml_struct(i).dml_identifier.Contains("functionality_help") Then
                            RichTextBox2.Text = RichTextBox2.Text + "-->" + l_dml_struct(i).dml_identifier + vbNewLine + dml.GET_PK_PROCESSED(l_dml_struct(i).dml_content) + vbNewLine + vbNewLine
                        Else
                            Dim dr As OracleDataReader

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
                            If Not dml.GET_TAG(l_dml_struct(i).dml_identifier, dr) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
                                MsgBox("ERROR GETTING DML TAGS!")
                                dr.Dispose()
                                dr.Close()
                                Exit For
                            Else
                                While dr.Read()
                                    l_tag = "|" + dr.Item(0) + "|" + dr.Item(1)
                                End While
                                dr.Dispose()
                                dr.Close()
                            End If
                            RichTextBox2.Text = RichTextBox2.Text + "-->" + l_dml_struct(i).dml_identifier + l_tag + vbNewLine + l_dml_struct(i).dml_content + vbNewLine + vbNewLine
                            l_tag = ""
                        End If
                    Else
                        If l_dml_struct(i).dml_identifier.Contains("translation") Or l_dml_struct(i).dml_identifier.Contains("sys_message") Or
                        l_dml_struct(i).dml_identifier.Contains("sys_domain") Or l_dml_struct(i).dml_identifier.Contains("functionality_help") Then
                            RichTextBox2.Text = RichTextBox2.Text + "-->" + l_dml_struct(i).dml_identifier + vbNewLine + dml.GET_PK_PROCESSED(l_dml_struct(i).dml_content)
                        Else
                            Dim dr As OracleDataReader
#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
                            If Not dml.GET_TAG(l_dml_struct(i).dml_identifier, dr) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
                                MsgBox("ERROR GETTING DML TAGS!")
                                dr.Dispose()
                                dr.Close()
                                Exit For
                            Else
                                While dr.Read()
                                    l_tag = "|" + dr.Item(0) + "|" + dr.Item(1)
                                End While
                                dr.Dispose()
                                dr.Close()
                            End If
                            RichTextBox2.Text = RichTextBox2.Text + "-->" + l_dml_struct(i).dml_identifier + l_tag + vbNewLine + l_dml_struct(i).dml_content
                            l_tag = ""
                        End If
                    End If
                Next
            End If

        Catch ex As Exception

            MsgBox("Error generating DML.", vbCritical)
        End Try

        Cursor = Cursors.Arrow

    End Sub

    Private Sub DML_TOOLS_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = "DML TOOLS  ::::  Connected to " & Connection.db

        Me.BackColor = Color.FromArgb(215, 215, 180)

        Me.Location = New Point(Form_location.x_position, Form_location.y_position)

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs)
        MsgBox(dml.GET_DML_IDENTIFIER(RichTextBox1.Text))
    End Sub

    Private Sub RichTextBox2_TextChanged(sender As Object, e As EventArgs) Handles RichTextBox2.TextChanged

    End Sub

    Private Sub Button2_Click_1(sender As Object, e As EventArgs) Handles Button2.Click
        Form_location.x_position = Me.Location.X
        Form_location.y_position = Me.Location.Y

        Me.Enabled = False
        Me.Dispose()
        Form1.Show()
    End Sub
End Class