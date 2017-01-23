Imports Oracle.DataAccess.Client
Public Class LABS_API
    'GET_DEFAULT_VERSIONS(ByVal i_institution As Int64, ByVal i_software As Integer, ByVal i_oradb As String) As OracleDataReader
    'GET_LAB_CATS_DEFAULT(ByVal i_version As String, ByVal i_institution As Int64, ByVal i_software As Integer, ByVal i_oradb As String) As OracleDataReader
    'GET_LABS_DEFAULT_BY_CAT(ByVal i_institution As Int64, ByVal i_software As Integer, ByVal i_version As String, ByVal i_id_cat As String, ByVal i_oradb As String) As OracleDataReader
    'GET_CODE_EXAM_CAT_ALERT(ByVal i_id_content_exam_cat As String, ByVal i_oradb As String) As String
    'GET_CODE_EXAM_CAT_DEFAULT(ByVal i_id_content_exam_cat As String, ByVal i_oradb As String) As String
    'GET_CODE_SAMPLE_TYPE_ALERT(ByVal i_id_content_st As String, ByVal i_oradb As String) As String
    'GET_CODE_SAMPLE_TYPE_DEFAULT(ByVal i_id_content_st As String, ByVal i_oradb As String) As String
    'GET_ID_CAT_ALERT(ByVal i_id_content_exam_cat As String, ByVal i_oradb As String) As Int64
    'GET_CAT_RANK(ByVal i_id_content_exam_cat As String, ByVal i_oradb As String) As Int64
    'GET_CAT_FLG_INTERFACE(ByVal i_id_content_exam_cat As String, ByVal i_oradb As String) As Char
    'GET_DEFAULT_ST_PARAMETERS(ByVal i_id_content_sample_type As String, ByVal i_oradb As String) As OracleDataReader
    'GET_DEFAULT_ANALYSIS_PARAMETERS(ByVal i_id_content_analysis As String, ByVal i_oradb As String) As OracleDataReader
    'GET_ID_SAMPLE_TYPE_ALERT(ByVal i_id_content_st As String, ByVal i_oradb As String) As Int64
    'GET_CODE_ANALYSIS_ALERT(ByVal i_id_content_a As String, ByVal i_oradb As String) As String
    'GET_CODE_ANALYSIS_DEFAULT(ByVal i_id_content_a As String, ByVal i_oradb As String) As String

    'CHECK_RECORD_EXISTENCE(ByVal i_id_content_record As String, ByVal i_sql As String, ByVal i_oradb As String) As Boolean
    'CHECK_RECORD_TRANSLATION_EXISTENCE(ByVal i_institution As Int64, ByVal id_content_record As String, ByVal i_sql As String, ByVal i_oradb As String) As Boolean

    'SET_EXAM_CAT(ByVal i_institution As Int64, ByVal id_content_category As String, ByVal i_oradb As String) As Boolean
    'SET_SAMPLE_TYPE(ByVal i_institution As Int64, ByVal id_content_sample_type As String, ByVal i_oradb As String) As Boolean
    'SET_ANALYSIS(ByVal i_institution As Int64, ByVal id_content_analysis As String, ByVal i_oradb As String) As Boolean

    Dim db_access_general As New General

    Public Structure analysis_default
        Public id_content_category As String
        Public id_content_analysis As String
        Public id_content_sample_type As String
        Public id_content_analysis_sample_type As String
        Public id_content_sample_recipient As String
        Public desc_analysis_sample_type As String
        Public desc_analysis_sample_recipient As String
    End Structure

    Public Structure analysis_alert
        Public id_content_analysis_sample_type As String
        Public desc_analysis_sample_type As String
        Public desc_analysis_sample_recipient As String
    End Structure

    Public Structure analysis_alert_flg
        Public id_content_analysis_sample_type As String
        Public desc_analysis_sample_type As String
        Public desc_analysis_sample_recipient As String
        Public flg_new As String
    End Structure

    Function GET_DEFAULT_VERSIONS(ByVal i_institution As Int64, ByVal i_software As Integer, ByVal i_conn As OracleConnection, ByRef i_dr As OracleDataReader) As Boolean

        Dim sql As String = "Select distinct dastv.version
                              from alert_default.analysis da
                              join alert_default.analysis_sample_type dast
                                on dast.id_analysis = da.id_analysis
                              join alert_default.sample_type dst
                                on dst.id_sample_type = dast.id_sample_type
                              join alert_default.translation dta --VERIFICAR SE TRADUÇÂO DE ANALISE É MESMO NECESSÁRIO
                                on dta.code_translation = da.code_analysis
                              join alert_default.translation dtst --VERIFICAR SE TRADUÇÂO DE SAMPLETYPE É MESMO NECESSÁRIO
                                on dtst.code_translation = dst.code_sample_type
                              join alert_default.translation dtast
                                on dtast.code_translation = dast.code_analysis_sample_type
                              join alert_default.analysis_mrk_vrs dav
                                on dav.id_analysis = da.id_analysis
                              join alert_default.ast_mkt_vrs dastv
                                on dastv.id_content = dast.id_content
                              join alert_default.analysis_instit_soft dais
                                on dais.id_analysis = da.id_analysis
                               and dais.id_sample_type = dst.id_sample_type
                              join alert_default.analysis_instit_recipient dair
                                on dair.id_analysis_instit_soft = dais.id_analysis_instit_soft
                              join alert_default.sample_recipient dsr
                                on dsr.id_sample_recipient = dair.id_sample_recipient
                              join alert_default.translation dtsr
                                on dtsr.code_translation = dsr.code_sample_recipient
                              join alert_default.analysis_param dap
                                on dap.id_analysis = da.id_analysis
                               and dap.id_sample_type = dst.id_sample_type
                              join alert_default.analysis_parameter dparameter
                                on dparameter.id_analysis_parameter = dap.id_analysis_parameter
                              join alert_default.translation dtparameter
                                on dtparameter.code_translation = dparameter.code_analysis_parameter
                              join institution i
                                on i.id_market = dav.id_market
                              join alert_default.exam_cat dec
                                on dec.id_exam_cat = dais.id_exam_cat
                              join alert_default.translation dtec
                                on dtec.code_translation = dec.code_exam_cat

                             where da.flg_available = 'Y'
                               and dst.flg_available = 'Y'
                               and dast.flg_available = 'Y'
                               and dsr.flg_available = 'Y'
                               and dais.flg_available = 'Y'
                               and dap.flg_available = 'Y'
                               and dparameter.flg_available = 'Y'
                               and dec.flg_available = 'Y'
                               and dais.id_software in (0, " & i_software & ")
                               and dap.id_software in (0, " & i_software & ")
                               and i.id_institution= " & i_institution & "
                               and dastv.id_market=i.id_market
                               and dav.id_market=i.id_market
                               and alert_default.pk_translation_default.get_translation_default(" & db_access_general.GET_ID_LANG(i_institution, i_conn) & ",dast.code_analysis_sample_type) is not null
                             order by 1 asc"

            Dim cmd As New OracleCommand(sql, i_conn)
            cmd.CommandType = CommandType.Text

        Try

            i_dr = cmd.ExecuteReader()
            cmd.Dispose()
            Return True

        Catch ex As Exception

            cmd.Dispose()
            Return False

        End Try

    End Function

    Function GET_LAB_CATS_DEFAULT(ByVal i_version As String, ByVal i_institution As Int64, ByVal i_software As Integer, ByVal i_conn As OracleConnection, ByRef i_dr As OracleDataReader) As Boolean

        Dim sql As String = "Select distinct dec.id_content, 
                                 nvl2(dec.parent_id,
                                 (alert_default.pk_translation_default.get_translation_default(" & db_access_general.GET_ID_LANG(i_institution, i_conn) & ", decp.code_exam_cat) || ' - ' ||
                                 alert_default.pk_translation_default.get_translation_default(" & db_access_general.GET_ID_LANG(i_institution, i_conn) & ", dec.code_exam_cat)),
                                 (alert_default.pk_translation_default.get_translation_default(" & db_access_general.GET_ID_LANG(i_institution, i_conn) & ", dec.code_exam_cat)))              
                              from alert_default.analysis da
                              join alert_default.analysis_sample_type dast
                                on dast.id_analysis = da.id_analysis
                              join alert_default.sample_type dst
                                on dst.id_sample_type = dast.id_sample_type
                              join alert_default.translation dta --VERIFICAR SE TRADUÇÂO DE ANALISE É MESMO NECESSÁRIO
                                on dta.code_translation = da.code_analysis
                              join alert_default.translation dtst --VERIFICAR SE TRADUÇÂO DE SAMPLETYPE É MESMO NECESSÁRIO
                                on dtst.code_translation = dst.code_sample_type
                              join alert_default.translation dtast
                                on dtast.code_translation = dast.code_analysis_sample_type
                              join alert_default.analysis_mrk_vrs dav
                                on dav.id_analysis = da.id_analysis
                              join alert_default.ast_mkt_vrs dastv
                                on dastv.id_content = dast.id_content
                              join alert_default.analysis_instit_soft dais
                                on dais.id_analysis = da.id_analysis
                               and dais.id_sample_type = dst.id_sample_type
                              join alert_default.analysis_instit_recipient dair
                                on dair.id_analysis_instit_soft = dais.id_analysis_instit_soft
                              join alert_default.sample_recipient dsr
                                on dsr.id_sample_recipient = dair.id_sample_recipient
                              join alert_default.translation dtsr
                                on dtsr.code_translation = dsr.code_sample_recipient
                              join alert_default.analysis_param dap
                                on dap.id_analysis = da.id_analysis
                               and dap.id_sample_type = dst.id_sample_type
                              join alert_default.analysis_parameter dparameter
                                on dparameter.id_analysis_parameter = dap.id_analysis_parameter
                              join alert_default.translation dtparameter
                                on dtparameter.code_translation = dparameter.code_analysis_parameter
                              join institution i
                                on i.id_market = dav.id_market
                              join alert_default.exam_cat dec
                                on dec.id_exam_cat = dais.id_exam_cat
                              join alert_default.translation dtec
                                on dtec.code_translation = dec.code_exam_cat
                         LEFT JOIN alert_default.exam_cat decp ON decp.id_exam_cat = dec.parent_id
                         LEFT JOIN alert_default.translation dtecp ON dtecp.code_translation = decp.code_exam_cat

                             where da.flg_available = 'Y'
                               and dst.flg_available = 'Y'
                               and dast.flg_available = 'Y'
                               and dsr.flg_available = 'Y'
                               and dais.flg_available = 'Y'
                               and dap.flg_available = 'Y'
                               and dparameter.flg_available = 'Y'
                               and dec.flg_available = 'Y'
                               and dais.id_software in (0, " & i_software & ")
                               and dap.id_software in (0, " & i_software & ")
                               and i.id_institution = " & i_institution & "
                               and dastv.id_market = i.id_market
                               and dav.id_market = i.id_market
                               and dastv.version = '" & i_version & "'
                               and dav.version= '" & i_version & "'
   
                             order by 2 asc"


        Dim cmd As New OracleCommand(sql, i_conn)

        Try
            cmd.CommandType = CommandType.Text
            i_dr = cmd.ExecuteReader()
            cmd.Dispose()
            Return True

        Catch ex As Exception

            cmd.Dispose()
            Return False

        End Try

    End Function

    Function GET_LAB_CATS_INST_SOFT(ByVal i_institution As Int64, ByVal i_software As Integer, ByVal i_conn As OracleConnection, ByRef i_dr As OracleDataReader) As Boolean

        Dim sql As String = "SELECT distinct ec.id_content,
                                        nvl2(ec.parent_id,
                                             (pk_translation.get_translation(" & db_access_general.GET_ID_LANG(i_institution, i_conn) & ", ecp.code_exam_cat) || ' - ' ||
                                              pk_translation.get_translation(" & db_access_general.GET_ID_LANG(i_institution, i_conn) & ", ec.code_exam_cat)),
                                             (pk_translation.get_translation(" & db_access_general.GET_ID_LANG(i_institution, i_conn) & ", ec.code_exam_cat)))     
                                        FROM alert.analysis_room ar

                                        JOIN alert.analysis_sample_type ast ON ast.id_analysis = ar.id_analysis
                                                                        AND ast.id_sample_type = ast.id_sample_type
                                                                        AND ast.flg_available = 'Y'
                                        JOIN alert.analysis a ON a.id_analysis = ast.id_analysis
                                                          AND a.flg_available = 'Y'
                                        JOIN alert.sample_type st ON st.id_sample_type = ast.id_sample_type
                                                              AND st.flg_available = 'Y'
                                        JOIN alert.analysis_instit_soft ais ON ais.id_analysis = ast.id_analysis
                                                                        AND ast.id_sample_type = ais.id_sample_type
                                                                        AND ais.id_institution = ar.id_institution
                                                                        AND ais.flg_available = 'Y'
                                        JOIN alert.analysis_instit_recipient air ON air.id_analysis_instit_soft = ais.id_analysis_instit_soft
                                        JOIN alert.analysis_param ap ON ap.id_analysis = ast.id_analysis
                                                                 AND ap.id_sample_type = ast.id_sample_type
                                                                 AND ap.id_institution = ar.id_institution
                                                                 AND ap.id_software = ais.id_software
                                                                 AND ap.flg_available = 'Y'
                                        JOIN alert.analysis_parameter parameter ON parameter.id_analysis_parameter = ap.id_analysis_parameter
                                                                            AND parameter.flg_available = 'Y'
                                        JOIN alert.exam_cat ec ON ec.id_exam_cat = ais.id_exam_cat
                                                           AND ec.flg_available = 'Y'
                                        JOIN translation tec ON tec.code_translation = ec.code_exam_cat
                                        LEFT JOIN alert.exam_cat ecp ON ecp.id_exam_cat = ec.parent_id
                                        LEFT JOIN translation tecp ON tecp.code_translation = ecp.code_exam_cat

                                        WHERE ar.flg_available = 'Y'
                                        AND ar.id_institution = " & i_institution & "
                                        AND ais.id_software = " & i_software & "
                                        order by 2 asc"

            Dim cmd As New OracleCommand(sql, i_conn)

        Try

            cmd.CommandType = CommandType.Text
            i_dr = cmd.ExecuteReader()
            cmd.Dispose()
            Return True

        Catch ex As Exception

            cmd.Dispose()
            Return False

        End Try

    End Function

    Function GET_LABS_INST_SOFT(ByVal i_institution As Int64, ByVal i_software As Integer, ByVal i_id_content_exam_cat As String, ByVal i_conn As OracleConnection, ByRef i_dr As OracleDataReader) As Boolean

        Dim sql As String = "SELECT distinct ast.id_content,
                                        pk_translation.get_translation(" & db_access_general.GET_ID_LANG(i_institution, i_conn) & ", ast.code_analysis_sample_type),
                                        pk_translation.get_translation(" & db_access_general.GET_ID_LANG(i_institution, i_conn) & ", sr.code_sample_recipient)     
                                    FROM alert.analysis_room ar

                                    JOIN alert.analysis_sample_type ast ON ast.id_analysis = ar.id_analysis
                                                                    AND ast.id_sample_type = ast.id_sample_type
                                                                    AND ast.flg_available = 'Y'
                                    JOIN alert.analysis a ON a.id_analysis = ast.id_analysis
                                                      AND a.flg_available = 'Y'
                                    JOIN alert.sample_type st ON st.id_sample_type = ast.id_sample_type
                                                          AND st.flg_available = 'Y'
                                    JOIN alert.analysis_instit_soft ais ON ais.id_analysis = ast.id_analysis
                                                                    AND ast.id_sample_type = ais.id_sample_type
                                                                    AND ais.id_institution = ar.id_institution
                                                                    AND ais.flg_available = 'Y'
                                    JOIN alert.analysis_instit_recipient air ON air.id_analysis_instit_soft = ais.id_analysis_instit_soft
                                    JOIN alert.analysis_param ap ON ap.id_analysis = ast.id_analysis
                                                             AND ap.id_sample_type = ast.id_sample_type
                                                             AND ap.id_institution = ar.id_institution
                                                             AND ap.id_software = ais.id_software
                                                             AND ap.flg_available = 'Y'
                                    JOIN alert.analysis_parameter parameter ON parameter.id_analysis_parameter = ap.id_analysis_parameter
                                                                        AND parameter.flg_available = 'Y'
                                    JOIN alert.exam_cat ec ON ec.id_exam_cat = ais.id_exam_cat
                                                       AND ec.flg_available = 'Y'                                    
                                    JOIN alert.sample_recipient sr ON sr.id_sample_recipient = air.id_sample_recipient
                                    AND sr.flg_available = 'Y'

                                    WHERE ar.flg_available = 'Y'
                                    AND ar.id_institution = " & i_institution & "
                                    AND ais.id_software = " & i_software

            If i_id_content_exam_cat <> "0" Then

                sql = sql & " And ec.id_content = '" & i_id_content_exam_cat & "'
                          order by 2 asc"

            Else

                sql = sql & " order by 2 asc"

            End If


            Dim cmd As New OracleCommand(sql, i_conn)

        Try

            cmd.CommandType = CommandType.Text
            i_dr = cmd.ExecuteReader()
            cmd.Dispose()
            Return True

        Catch ex As Exception

            cmd.Dispose()
            Return False

        End Try

    End Function

    Function GET_LABS_DEFAULT_BY_CAT(ByVal i_institution As Int64, ByVal i_software As Integer, ByVal i_version As String, ByVal i_id_cat As String, ByVal i_conn As OracleConnection, ByRef i_dr As OracleDataReader) As Boolean

        Dim sql As String = "Select distinct dast.id_content, 
                                             alert_default.pk_translation_default.get_translation_default(" & db_access_general.GET_ID_LANG(i_institution, i_conn) & ", dast.code_analysis_sample_type), 
                                             dsr.id_content,              
                                             alert_default.pk_translation_default.get_translation_default(" & db_access_general.GET_ID_LANG(i_institution, i_conn) & ", dsr.code_sample_recipient), 
                                             da.id_content, 
                                             dst.id_content,
                                             dec.id_content
              
                                              from alert_default.analysis da
                                              join alert_default.analysis_sample_type dast
                                                on dast.id_analysis = da.id_analysis
                                              join alert_default.sample_type dst
                                                on dst.id_sample_type = dast.id_sample_type
                                              join alert_default.translation dta --VERIFICAR SE TRADUÇÂO DE ANALISE É MESMO NECESSÁRIO
                                                on dta.code_translation = da.code_analysis
                                              join alert_default.translation dtst --VERIFICAR SE TRADUÇÂO DE SAMPLETYPE É MESMO NECESSÁRIO
                                                on dtst.code_translation = dst.code_sample_type
                                              join alert_default.translation dtast
                                                on dtast.code_translation = dast.code_analysis_sample_type
                                              join alert_default.analysis_mrk_vrs dav
                                                on dav.id_analysis = da.id_analysis
                                              join alert_default.ast_mkt_vrs dastv
                                                on dastv.id_content = dast.id_content
                                              join alert_default.analysis_instit_soft dais
                                                on dais.id_analysis = da.id_analysis
                                               and dais.id_sample_type = dst.id_sample_type
                                              join alert_default.analysis_instit_recipient dair
                                                on dair.id_analysis_instit_soft = dais.id_analysis_instit_soft
                                              join alert_default.sample_recipient dsr
                                                on dsr.id_sample_recipient = dair.id_sample_recipient
                                              join alert_default.translation dtsr
                                                on dtsr.code_translation = dsr.code_sample_recipient
                                              join alert_default.analysis_param dap
                                                on dap.id_analysis = da.id_analysis
                                               and dap.id_sample_type = dst.id_sample_type
                                              join alert_default.analysis_parameter dparameter
                                                on dparameter.id_analysis_parameter = dap.id_analysis_parameter
                                              join alert_default.translation dtparameter
                                                on dtparameter.code_translation = dparameter.code_analysis_parameter
                                              join institution i
                                                on i.id_market = dav.id_market
                                              join alert_default.exam_cat dec
                                                on dec.id_exam_cat = dais.id_exam_cat
                                              join alert_default.translation dtec
                                                on dtec.code_translation = dec.code_exam_cat

                                             where da.flg_available = 'Y'
                                               and dst.flg_available = 'Y'
                                               and dast.flg_available = 'Y'
                                               and dsr.flg_available = 'Y'
                                               and dais.flg_available = 'Y'
                                               and dap.flg_available = 'Y'
                                               and dparameter.flg_available = 'Y'
                                               and dec.flg_available = 'Y'
                                               and dais.id_software in (0, " & i_software & ")
                                               and dap.id_software in (0, " & i_software & ")
                                               and i.id_institution = " & i_institution & "
                                               and dastv.id_market = i.id_market
                                               and dav.id_market = i.id_market
                                               and dastv.version = '" & i_version & "'
                                               and dav.version= '" & i_version & "'"


        If i_id_cat = "0" Then

            sql = sql & " order by 2 asc, 4 asc"

        Else

            sql = sql & " and dec.id_content= '" & i_id_cat & "'
                         order by 2 asc, 4 asc"
        End If

        Dim cmd As New OracleCommand(sql, i_conn)

        Try

            cmd.CommandType = CommandType.Text
            i_dr = cmd.ExecuteReader()
            cmd.Dispose()
            Return True

        Catch ex As Exception

            cmd.Dispose()
            Return False

        End Try

    End Function

    Function GET_CODE_EXAM_CAT_ALERT(ByVal i_id_content_exam_cat As String, ByVal i_conn As OracleConnection) As String

        Dim sql As String = "Select ec.code_exam_cat from alert.exam_cat ec
                             where ec.id_content='" & i_id_content_exam_cat & "'
                             and ec.flg_available='Y'"

        Dim l_code As String = ""

        Dim cmd As New OracleCommand(sql, i_conn)
        cmd.CommandType = CommandType.Text

        Dim dr As OracleDataReader = cmd.ExecuteReader()

        While dr.Read()

            l_code = dr.Item(0)

        End While

        cmd.Dispose()
        dr.Dispose()
        dr.Close()

        Return l_code

    End Function

    Function GET_CODE_EXAM_CAT_DEFAULT(ByVal i_id_content_exam_cat As String, ByVal i_conn As OracleConnection) As String

        Dim sql As String = "Select ec.code_exam_cat from alert_default.exam_cat ec
                             where ec.id_content='" & i_id_content_exam_cat & "'
                             and ec.flg_available='Y'"

        Dim l_code As String = ""

        Dim cmd As New OracleCommand(sql, i_conn)
        cmd.CommandType = CommandType.Text

        Dim dr As OracleDataReader = cmd.ExecuteReader()

        While dr.Read()

            l_code = dr.Item(0)

        End While

        dr.Dispose()
        dr.Close()
        cmd.Dispose()

        Return l_code

    End Function

    Function GET_CODE_SAMPLE_TYPE_ALERT(ByVal i_id_content_st As String, ByVal i_conn As OracleConnection) As String

        Dim sql As String = "Select st.code_sample_type from alert.sample_type st
                             where st.id_content='" & i_id_content_st & "'
                             and st.flg_available='Y'"

        Dim l_code As String = ""

        Dim cmd As New OracleCommand(sql, i_conn)
        cmd.CommandType = CommandType.Text

        Dim dr As OracleDataReader = cmd.ExecuteReader()

        While dr.Read()

            l_code = dr.Item(0)

        End While

        dr.Dispose()
        dr.Close()
        cmd.Dispose()

        Return l_code

    End Function

    Function GET_CODE_SAMPLE_TYPE_DEFAULT(ByVal i_id_content_st As String, ByVal i_conn As OracleConnection) As String

        Dim sql As String = "Select st.code_sample_type from alert_default.sample_type st
                             where st.id_content='" & i_id_content_st & "'
                             and st.flg_available='Y'"

        Dim l_code As String = ""

        Dim cmd As New OracleCommand(sql, i_conn)
        cmd.CommandType = CommandType.Text

        Dim dr As OracleDataReader = cmd.ExecuteReader()

        While dr.Read()

            l_code = dr.Item(0)

        End While

        dr.Dispose()
        dr.Close()
        cmd.Dispose()

        Return l_code

    End Function

    Function GET_ID_CAT_ALERT(ByVal i_id_content_exam_cat As String, ByVal i_conn As OracleConnection) As Int64

        Dim sql As String = "Select ec.id_exam_cat from alert.exam_cat ec
                             where ec.id_content='" & i_id_content_exam_cat & "'
                             and ec.flg_available='Y'"

        Dim l_id_alert As Int64 = 0

        Dim cmd As New OracleCommand(sql, i_conn)
        cmd.CommandType = CommandType.Text

        Dim dr As OracleDataReader = cmd.ExecuteReader()

        While dr.Read()

            l_id_alert = dr.Item(0)

        End While

        dr.Dispose()
        dr.Close()
        cmd.Dispose()
        Return l_id_alert

    End Function

    Function GET_CAT_RANK(ByVal i_id_content_exam_cat As String, ByVal i_conn As OracleConnection) As Int64

        Dim sql As String = "Select ec.rank from alert.exam_cat ec
                             where ec.id_content='" & i_id_content_exam_cat & "'
                             and ec.flg_available='Y'"

        Dim l_id_alert As Int64 = 0

        Dim cmd As New OracleCommand(sql, i_conn)
        cmd.CommandType = CommandType.Text

        Dim dr As OracleDataReader = cmd.ExecuteReader()

        While dr.Read()

            l_id_alert = dr.Item(0)

        End While

        dr.Dispose()
        dr.Close()
        cmd.Dispose()
        Return l_id_alert

    End Function

    Function GET_CAT_FLG_INTERFACE(ByVal i_id_content_exam_cat As String, ByVal i_conn As OracleConnection) As Char

        Dim sql As String = "Select ec.flg_interface from alert_DEFAULT.exam_cat ec
                             where ec.id_content='" & i_id_content_exam_cat & "'"

        Dim l_flg_interface As Char = ""

        Dim cmd As New OracleCommand(sql, i_conn)
        cmd.CommandType = CommandType.Text

        Dim dr As OracleDataReader = cmd.ExecuteReader()

        While dr.Read()

            l_flg_interface = dr.Item(0)

        End While

        dr.Dispose()
        dr.Close()
        cmd.Dispose()
        Return l_flg_interface

    End Function

    Function GET_DEFAULT_ST_PARAMETERS(ByVal i_id_content_sample_type As String, ByVal i_conn As OracleConnection, ByRef i_dr As OracleDataReader) As Boolean

        Dim sql As String = "Select dst.gender, dst.age_min, dst.age_max from alert_default.sample_type dst
                             where dst.id_content='" & i_id_content_sample_type & "'
                             and dst.flg_available='Y'"

        Dim cmd As New OracleCommand(sql, i_conn)
        cmd.CommandType = CommandType.Text

        Try

            i_dr = cmd.ExecuteReader()
            cmd.Dispose()
            Return True

        Catch ex As Exception

            cmd.Dispose()
            Return False

        End Try

    End Function

    Function GET_DEFAULT_ANALYSIS_PARAMETERS(ByVal i_id_content_analysis As String, ByVal i_conn As OracleConnection, ByRef i_dr As OracleDataReader) As Boolean

        Dim sql As String = "SELECT a.cpt_code, a.gender, a.age_min, a.age_max, a.mdm_coding, a.ref_form_code, st.id_content, a.barcode
                                FROM alert_default.analysis a
                                LEFT JOIN alert_default.sample_type st ON st.id_sample_type = a.id_sample_type                               
                                WHERE a.id_content = '" & i_id_content_analysis & "'
                                AND a.flg_available = 'Y'"

        Dim cmd As New OracleCommand(sql, i_conn)

        Try

            cmd.CommandType = CommandType.Text
            i_dr = cmd.ExecuteReader()
            cmd.Dispose()
            Return True

        Catch ex As Exception

            cmd.Dispose()
            Return False

        End Try

    End Function

    Function GET_ID_SAMPLE_TYPE_ALERT(ByVal i_id_content_st As String, ByVal i_conn As OracleConnection) As Int64

        Dim sql As String = "Select st.id_sample_type from alert.sample_type st
                            where st.id_content='" & i_id_content_st & "'
                            and st.flg_available='Y'"

        Dim l_id_alert As Int64 = 0

        Dim cmd As New OracleCommand(sql, i_conn)
        cmd.CommandType = CommandType.Text

        Dim dr As OracleDataReader = cmd.ExecuteReader()

        While dr.Read()

            l_id_alert = dr.Item(0)

        End While

        dr.Dispose()
        dr.Close()
        cmd.Dispose()

        Return l_id_alert

    End Function

    Function GET_ID_ANALYSIS_ALERT(ByVal i_id_content_a As String, ByVal i_conn As OracleConnection) As Int64

        Dim sql As String = "Select a.id_analysis from alert.ANALYSIS a
                            where a.id_content='" & i_id_content_a & "'
                            and a.flg_available='Y'"

        Dim l_id_alert As Int64 = 0

        Dim cmd As New OracleCommand(sql, i_conn)
        cmd.CommandType = CommandType.Text

        Dim dr As OracleDataReader = cmd.ExecuteReader()

        While dr.Read()

            l_id_alert = dr.Item(0)

        End While

        dr.Dispose()
        dr.Close()
        cmd.Dispose()

        Return l_id_alert

    End Function

    Function GET_CODE_ANALYSIS_ALERT(ByVal i_id_content_a As String, ByVal i_conn As OracleConnection) As String

        Dim sql As String = "Select a.code_analysis from alert.analysis a
                             where a.id_content='" & i_id_content_a & "'
                             and a.flg_available='Y'"

        Dim l_code As String = ""

        Dim cmd As New OracleCommand(sql, i_conn)
        cmd.CommandType = CommandType.Text

        Dim dr As OracleDataReader = cmd.ExecuteReader()

        While dr.Read()

            l_code = dr.Item(0)

        End While

        dr.Dispose()
        dr.Close()
        cmd.Dispose()

        Return l_code

    End Function

    Function GET_CODE_ANALYSIS_DEFAULT(ByVal i_id_content_a As String, ByVal i_conn As OracleConnection) As String

        Dim sql As String = "Select a.code_analysis from alert_default.analysis a
                             where a.id_content='" & i_id_content_a & "'"

        Dim l_code As String = ""

        Dim cmd As New OracleCommand(sql, i_conn)
        cmd.CommandType = CommandType.Text

        Dim dr As OracleDataReader = cmd.ExecuteReader()

        While dr.Read()

            l_code = dr.Item(0)

        End While

        dr.Dispose()
        dr.Close()
        cmd.Dispose()

        Return l_code

    End Function

    Function GET_CODE_ANALYSIS_ST_ALERT(ByVal i_id_content_ast As String, ByVal i_conn As OracleConnection) As String

        Dim sql As String = "Select ast.code_analysis_sample_type from alert.analysis_sample_type ast
                             where ast.id_content='" & i_id_content_ast & "'
                             and ast.flg_available='Y'"

        Dim l_code As String = ""

        Dim cmd As New OracleCommand(sql, i_conn)
        cmd.CommandType = CommandType.Text

        Dim dr As OracleDataReader = cmd.ExecuteReader()

        While dr.Read()

            l_code = dr.Item(0)

        End While

        dr.Dispose()
        dr.Close()
        cmd.Dispose()

        Return l_code

    End Function

    Function GET_CODE_SAMPLE_RECIPIENT_ALERT(ByVal i_id_content_sr As String, ByVal i_conn As OracleConnection) As String

        Dim sql As String = "Select sr.code_sample_recipient from alert.sample_recipient sr
                             where sr.id_content='" & i_id_content_sr & "'
                             and sr.flg_available='Y'"

        Dim l_code As String = ""

        Dim cmd As New OracleCommand(sql, i_conn)
        cmd.CommandType = CommandType.Text

        Dim dr As OracleDataReader = cmd.ExecuteReader()

        While dr.Read()

            l_code = dr.Item(0)

        End While

        dr.Dispose()
        dr.Close()
        cmd.Dispose()

        Return l_code

    End Function

    Function GET_CODE_PARAMETER_ALERT(ByVal i_id_content_parameter As String, ByVal i_conn As OracleConnection) As String

        Dim sql As String = "Select ap.code_analysis_parameter from alert.analysis_parameter ap
                             where ap.id_content='" & i_id_content_parameter & "'
                             and ap.flg_available='Y'"

        Dim l_code As String = ""

        Dim cmd As New OracleCommand(sql, i_conn)
        cmd.CommandType = CommandType.Text

        Dim dr As OracleDataReader = cmd.ExecuteReader()

        While dr.Read()

            l_code = dr.Item(0)

        End While

        dr.Dispose()
        dr.Close()
        cmd.Dispose()

        Return l_code

    End Function

    Function GET_CODE_ANALYSIS_ST_DEFAULT(ByVal i_id_content_ast As String, ByVal i_conn As OracleConnection) As String

        Dim sql As String = "Select ast.code_analysis_sample_type from alert_default.analysis_sample_type ast
                             where ast.id_content='" & i_id_content_ast & "'
                             and ast.flg_available='Y'"

        Dim l_code As String = ""

        Dim cmd As New OracleCommand(sql, i_conn)
        cmd.CommandType = CommandType.Text

        Dim dr As OracleDataReader = cmd.ExecuteReader()

        While dr.Read()

            l_code = dr.Item(0)

        End While

        dr.Dispose()
        dr.Close()
        cmd.Dispose()

        Return l_code

    End Function

    Function GET_CODE_SAMPLE_RECIPIENT_DEFAULT(ByVal i_id_content_sr As String, ByVal i_conn As OracleConnection) As String

        Dim sql As String = "Select sr.code_sample_recipient from alert_default.sample_recipient sr
                             where sr.id_content='" & i_id_content_sr & "'
                             and sr.flg_available='Y'"

        Dim l_code As String = ""

        Dim cmd As New OracleCommand(sql, i_conn)
        cmd.CommandType = CommandType.Text

        Dim dr As OracleDataReader = cmd.ExecuteReader()

        While dr.Read()

            l_code = dr.Item(0)

        End While

        dr.Dispose()
        dr.Close()
        cmd.Dispose()

        Return l_code

    End Function

    Function GET_CODE_PARAMETER_DEFAULT(ByVal i_id_content_parameter As String, ByVal i_conn As OracleConnection) As String

        Dim sql As String = "Select ap.code_analysis_parameter from alert_default.analysis_parameter ap
                                where ap.id_content='" & i_id_content_parameter & "'
                                and ap.flg_available='Y'"

        Dim l_code As String = ""

        Dim cmd As New OracleCommand(sql, i_conn)
        cmd.CommandType = CommandType.Text

        Dim dr As OracleDataReader = cmd.ExecuteReader()

        While dr.Read()

            l_code = dr.Item(0)

        End While

        dr.Dispose()
        dr.Close()
        cmd.Dispose()

        Return l_code

    End Function

    Function GET_DEFAULT_ANALYSIS_ST_PARAMETERS(ByVal i_id_content_analysis_st As String, ByVal i_conn As OracleConnection, ByRef i_Dr As OracleDataReader) As Boolean

        Dim sql As String = "Select ast.gender, ast.age_min, ast.age_max from alert_default.analysis_sample_type ast
                                where ast.id_content='" & i_id_content_analysis_st & "'
                                and ast.flg_available='Y'"


        Dim cmd As New OracleCommand(sql, i_conn)

        Try

            cmd.CommandType = CommandType.Text
            i_Dr = cmd.ExecuteReader()
            cmd.Dispose()
            Return True

        Catch ex As Exception

            cmd.Dispose()
            Return False

        End Try

    End Function

    Function GET_ANALYSIS_PARAMETERS_ID_CONTENT_DEFAULT(ByVal i_id_software As Int16, ByVal i_id_content_ast As String, ByVal i_conn As OracleConnection, ByRef i_dr As OracleDataReader) As Boolean

        Dim sql As String = "SELECT DISTINCT aparam.id_content
                                    FROM alert_default.analysis_sample_type ast
                                    JOIN alert_default.analysis_param ap ON ap.id_analysis = ast.id_analysis
                                                                     AND ap.id_sample_type = ast.id_sample_type
                                                                     AND ap.id_software = " & i_id_software & "
                                    JOIN alert_default.analysis_parameter aparam ON aparam.id_analysis_parameter = ap.id_analysis_parameter
                                                                             AND aparam.flg_available = 'Y'
                                    WHERE ast.flg_available = 'Y'
                                    AND ap.flg_available = 'Y'
                                    AND ast.id_content = '" & i_id_content_ast & "'
                                    ORDER BY 1 ASC"

            Dim cmd As New OracleCommand(sql, i_conn)

        Try

            cmd.CommandType = CommandType.Text
            i_dr = cmd.ExecuteReader()
            cmd.Dispose()
            Return True

        Catch ex As Exception

            cmd.Dispose()
            Return False

        End Try

    End Function

    Function GET_DEFAULT_PARAMETERS(ByVal i_id_content_analysis_st As String, ByVal i_software As Int16, ByVal i_conn As OracleConnection, ByRef i_Dr As OracleDataReader) As Boolean

        Dim sql As String = "SELECT dp.id_content, dap.color_graph, dap.flg_fill_type, dap.rank
                                FROM alert_default.analysis_param dap
                                JOIN alert_default.analysis_sample_type dast ON dast.id_analysis = dap.id_analysis
                                                                         AND dap.id_sample_type = dast.id_sample_type
                                JOIN alert_default.analysis_parameter dp ON dp.id_analysis_parameter = dap.id_analysis_parameter
                                WHERE dast.id_content = '" & i_id_content_analysis_st & "'
                                AND dap.id_software = " & i_software & "
                                AND dap.flg_available = 'Y'
                                AND dp.flg_available = 'Y'"

        Dim cmd As New OracleCommand(sql, i_conn)

        Try

            cmd.CommandType = CommandType.Text
            i_Dr = cmd.ExecuteReader()
            cmd.Dispose()
            Return True

        Catch ex As Exception

            cmd.Dispose()
            Return False

        End Try

    End Function

    Function CHECK_RECORD_EXISTENCE(ByVal i_id_content_record As String, ByVal i_sql As String, ByVal i_conn As OracleConnection) As Boolean

        Dim l_total_records As Int16 = 0

            Dim sql As String = "Select count(*) from " & i_sql & " r
                                 where r.id_content='" & i_id_content_record & "'
                                 and r.flg_available='Y'"

            Dim cmd As New OracleCommand(sql, i_conn)
        cmd.CommandType = CommandType.Text
        Dim dr As OracleDataReader = cmd.ExecuteReader()

        Try

            While dr.Read()

                l_total_records = dr.Item(0)

            End While

            dr.Dispose()
            dr.Close()
            cmd.Dispose()

            'Se l_total_analysis for maior que 0 significa que a análise já existe no ALERT

            If l_total_records > 0 Then

                Return True

            Else

                Return False

            End If

        Catch ex As Exception

            dr.Dispose()
            dr.Close()
            cmd.Dispose()
            Return False

        End Try

    End Function

    Function CHECK_RECORD_TRANSLATION_EXISTENCE(ByVal i_institution As Int64, ByVal id_content_record As String, ByVal i_sql As String, ByVal i_conn As OracleConnection) As Boolean

        Dim l_translation As String = ""

        Dim sql As String = "Select pk_translation.get_translation(" & db_access_general.GET_ID_LANG(i_institution, i_conn) & "," & i_sql & " r
                             where r.id_content='" & id_content_record & "'
                             And r.flg_available='Y'"

        Dim cmd As New OracleCommand(sql, i_conn)
        cmd.CommandType = CommandType.Text

        Dim dr As OracleDataReader = cmd.ExecuteReader()

        Try

            While dr.Read()

                l_translation = dr.Item(0)

            End While

            dr.Dispose()
            dr.Close()
            cmd.Dispose()

        Catch ex As Exception

            dr.Dispose()
            dr.Close()
            cmd.Dispose()

            Return False

        End Try

        Return True

    End Function

    Function GET_DISTINCT_CATEGORIES(ByVal i_selected_default_analysis() As analysis_default, ByVal i_conn As OracleConnection, ByRef i_Dr As OracleDataReader) As Boolean

        Dim sql As String = "Select distinct ec.id_content from alert.exam_cat ec
                                where ec.flg_available = 'Y'
                                and ec.id_content in ("


        For i As Integer = 0 To i_selected_default_analysis.Count() - 1

            If i < i_selected_default_analysis.Count() - 1 Then

                sql = sql & "'" & i_selected_default_analysis(i).id_content_category & "',"

            Else

                sql = sql & "'" & i_selected_default_analysis(i).id_content_category & "')"

            End If

        Next

        Dim cmd As New OracleCommand(sql, i_conn)

        Try

            cmd.CommandType = CommandType.Text
            i_Dr = cmd.ExecuteReader()
            cmd.Dispose()
            Return True

        Catch ex As Exception

            cmd.Dispose()
            Return False

        End Try

    End Function

    Function SET_EXAM_CAT(ByVal i_institution As Int64, ByVal i_selected_default_analysis() As analysis_default, ByVal i_conn As OracleConnection) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution, i_conn)

        '1 - Remover as categorias repetidas do array de entrada
        Dim l_a_distinct_ec() As String
        Dim dr_distinct_ec As OracleDataReader

        Try

            If Not GET_DISTINCT_CATEGORIES(i_selected_default_analysis, i_conn, dr_distinct_ec) Then

                dr_distinct_ec.Dispose()
                dr_distinct_ec.Close()
                Return False

            Else

                Dim l_index As Int64 = 0

                While dr_distinct_ec.Read()

                    ReDim Preserve l_a_distinct_ec(l_index)
                    l_a_distinct_ec(l_index) = dr_distinct_ec.Item(0)
                    l_index = l_index + 1

                End While

                dr_distinct_ec.Dispose()
                dr_distinct_ec.Close()

            End If

        Catch ex As Exception

            dr_distinct_ec.Dispose()
            dr_distinct_ec.Close()
            Return False

        End Try

        '2 - Processar as categorias de exames que já foram filtradas pelo bloco anterior
        Try

            'Ciclo que vai correr as categorias todas enviadas à função
            For i As Integer = 0 To l_a_distinct_ec.Count() - 1

                '' 1 - Verificar se existe Categoria pai
                Dim l_cat_parent As Int64 = 0
                Dim l_id_alert_cat_parent As Int64 = 0
                Dim l_rank As Int64 = 0
                Dim l_flg_interface As Char = ""

                Dim sql As String = "Select ec.parent_id from alert_default.Exam_Cat ec
                                     where ec.id_content='" & l_a_distinct_ec(i) & "'"

                Dim cmd As New OracleCommand(sql, i_conn)

                cmd.CommandType = CommandType.Text

                Dim dr As OracleDataReader = cmd.ExecuteReader()

                Try

                    While dr.Read()

                        l_cat_parent = dr.Item(0)

                    End While

                    dr.Dispose()
                    dr.Close()
                    cmd.Dispose()

                Catch ex As Exception

                    l_cat_parent = 0

                    dr.Dispose()
                    dr.Close()
                    cmd.Dispose()

                End Try

                If l_cat_parent > 0 Then 'Significa que existe Categoria Pai no default

                    '' 1.1 - Verificar se cat pai e tradução existem no alert. Se não existem, inserir.
                    ''1.1.1 - Verificar se existe no ALERT
                    sql = "Select ecp.id_content
                       from alert_default.exam_cat ec
                       join alert_default.exam_cat ecp
                       on ecp.id_exam_cat = ec.parent_id
                       where ec.id_content = '" & l_a_distinct_ec(i) & "'"

                    Dim l_id_content_cat_parent As String = ""
                    Dim cmd_2 As New OracleCommand(sql, i_conn)
                    cmd_2.CommandType = CommandType.Text
                    Dim dr_2 As OracleDataReader = cmd_2.ExecuteReader()

                    While dr_2.Read()

                        l_id_content_cat_parent = dr_2.Item(0)

                    End While

                    dr_2.Dispose()
                    dr_2.Close()
                    cmd_2.Dispose()

                    If Not CHECK_RECORD_EXISTENCE(l_id_content_cat_parent, "alert.exam_cat", i_conn) Then 'Significa que Categoria Pai não existe no ALERT, é necessário inserir.

                        'INSERT EXAM_CAT_PARENT  -Criar função de inserção de categoria(Recursivo)? e função de inserção de tradução ( de tradução deve ir para o generall)
                        'Estrutura auxiliar para ser chamada na recursividade (apenas terá o  id_content da categoria pai)
                        Dim l_analysis(0) As analysis_default
                        l_analysis(0).id_content_category = l_id_content_cat_parent

                        If Not SET_EXAM_CAT(i_institution, l_analysis, i_conn) Then

                            MsgBox("ERROR INSERTING EXAM_CAT_PARENT - LABS_API >> SET_EXAM_CAT")
                            Return False

                        End If

                        'Uma vez que foi adicionada uma nova categoria pai, sérá necessário atualizar o id alert da categoria pai das categorias filho
                        Dim sql_update_parents As String = "UPDATE alert.exam_cat ec
                                                        SET ec.parent_id = (Select ecp_n.id_exam_cat from alert.exam_cat ecp_n where ecp_n.id_content='" & l_id_content_cat_parent & "' and ecp_n.flg_available='Y')
                                                        WHERE ec.parent_id IN (SELECT ecp.id_exam_cat
                                                                               FROM alert.exam_cat ecp
                                                                               WHERE ecp.id_content = '" & l_id_content_cat_parent & "')"

                        Dim cmd_update_parents As New OracleCommand(sql_update_parents, i_conn)
                        cmd_update_parents.CommandType = CommandType.Text

                        cmd_update_parents.ExecuteNonQuery()

                        cmd_update_parents.Dispose()

                        '1.2 - Existe registo no ALERT, verificar se eciste tradução para a língua da isntituição
                    ElseIf Not CHECK_RECORD_TRANSLATION_EXISTENCE(i_institution, l_id_content_cat_parent, "r.code_exam_cat) from alert.exam_cat", i_conn) Then

                        ''Inserir tradução no ALERT
                        Dim l_code_cat_parent As String = GET_CODE_EXAM_CAT_ALERT(l_id_content_cat_parent, i_conn)
                        Dim l_exam_translation_default As String = db_access_general.GET_DEFAULT_TRANSLATION(l_id_language, GET_CODE_EXAM_CAT_DEFAULT(l_id_content_cat_parent, i_conn), i_conn)

                        If Not db_access_general.SET_TRANSLATION(l_id_language, l_code_cat_parent, l_exam_translation_default, i_conn) Then

                            MsgBox("ERROR INSERTING EXAM CATEGORY TRANSLATION - LABS_API >>  SET_TRANSLATION")
                            Return False

                        End If

                        '1.3 - Uma vez que existe no alert e existe tradução, verificar se tradução do alert é igual à do default
                    ElseIf Not db_access_general.CHECK_TRANSLATIONS(l_id_language, GET_CODE_EXAM_CAT_DEFAULT(l_id_content_cat_parent, i_conn), GET_CODE_EXAM_CAT_ALERT(l_id_content_cat_parent, i_conn), i_conn) Then

                        Dim l_code_cat_parent As String = GET_CODE_EXAM_CAT_ALERT(l_id_content_cat_parent, i_conn)
                        Dim l_exam_translation_default As String = db_access_general.GET_DEFAULT_TRANSLATION(l_id_language, GET_CODE_EXAM_CAT_DEFAULT(l_id_content_cat_parent, i_conn), i_conn)

                        If Not db_access_general.SET_TRANSLATION(l_id_language, l_code_cat_parent, l_exam_translation_default, i_conn) Then

                            MsgBox("ERROR INSERTING EXAM CATEGORY TRANSLATION - LABS_API >>  SET_TRANSLATION")
                            Return False

                        End If

                    End If

                    '' 1.2 - Se existir no alert, determinar id. (Neste ponto já vai sempre existir)
                    Try
                        l_id_alert_cat_parent = GET_ID_CAT_ALERT(l_id_content_cat_parent, i_conn)

                    Catch ex As Exception

                        MsgBox("ERROR GETTING ID_EXAM_CATEGORY FROM ALERT - LABS_API >>  GET_ID_CAT_ALERT")
                        Return False

                    End Try

                End If

                '2 - Verificar se categoria já existe no ALERT
                If Not CHECK_RECORD_EXISTENCE(l_a_distinct_ec(i), "alert.exam_cat", i_conn) Then

                    '2.1 - Não existe, Inserir.
                    '2.1.1 - Determinar RANK da categoria E flg_interface
                    Try

                        l_rank = GET_CAT_RANK(l_a_distinct_ec(i), i_conn)

                    Catch ex As Exception

                        l_rank = 0

                    End Try

                    '2.1.2 - Determinar flg_interface da categoria
                    Try

                        l_flg_interface = GET_CAT_FLG_INTERFACE(l_a_distinct_ec(i), i_conn)

                    Catch ex As Exception

                        l_flg_interface = "N"

                    End Try

                    '2.1.3 - Inserir Categoria
                    Dim sql_insert_cat As String

                    If l_id_alert_cat_parent = 0 Then

                        sql_insert_cat = "begin
                                      insert into alert.exam_cat (ID_EXAM_CAT, CODE_EXAM_CAT, FLG_AVAILABLE, FLG_LAB, ID_CONTENT, FLG_INTERFACE, RANK, PARENT_ID)
                                      values (alert.seq_exam_cat.nextval, 'EXAM_CAT.CODE_EXAM_CAT.' || alert.seq_exam_cat.nextval, 'Y', 'Y', '" & l_a_distinct_ec(i) & "', '" & l_flg_interface & "', " & l_rank & ", null);
                                      end;"
                    Else

                        sql_insert_cat = "begin
                                      insert into alert.exam_cat (ID_EXAM_CAT, CODE_EXAM_CAT, FLG_AVAILABLE, FLG_LAB, ID_CONTENT, FLG_INTERFACE, RANK, PARENT_ID)
                                      values (alert.seq_exam_cat.nextval, 'EXAM_CAT.CODE_EXAM_CAT.' || alert.seq_exam_cat.nextval, 'Y', 'Y', '" & l_a_distinct_ec(i) & "', '" & l_flg_interface & "', " & l_rank & ", " & l_id_alert_cat_parent & ");
                                      end;"

                    End If

                    Dim cmd_insert_cat As New OracleCommand(sql_insert_cat, i_conn)
                    cmd_insert_cat.CommandType = CommandType.Text

                    Try
                        cmd_insert_cat.ExecuteNonQuery()
                    Catch ex As Exception

                        MsgBox("ERROR INSERTING EXAM CATEGORY")
                        cmd_insert_cat.Dispose()
                        Return False

                    End Try

                    cmd_insert_cat.Dispose()

                    Dim l_code_ec_default As String = GET_CODE_EXAM_CAT_DEFAULT(l_a_distinct_ec(i), i_conn)
                    Dim l_code_ec_alert As String = GET_CODE_EXAM_CAT_ALERT(l_a_distinct_ec(i), i_conn)

                    '2.1.4 Inserir translation
                    If Not db_access_general.SET_TRANSLATION((l_id_language), (l_code_ec_alert), (db_access_general.GET_DEFAULT_TRANSLATION(l_id_language, l_code_ec_default, i_conn)), (i_conn)) Then

                        MsgBox("ERROR INSERTING CATEGORY TRANSLATION - LABS_API >> SET_TRANSLATION")
                        Return False

                    End If

                    '2.1.5 - Fazer update a todas as análises que utilizavam o id da categoria antiga com o id da nova categoria (alert.analysis_instit_soft)

                    Dim l_id_alert_category As Int64 = GET_ID_CAT_ALERT(l_a_distinct_ec(i), i_conn)

                    Dim sql_update_analysis_cat As String = "update alert.analysis_instit_soft ais 
                                                         set ais.id_exam_cat=" & l_id_alert_category & "
                                                         where ais.id_exam_cat in (select ec.id_exam_cat  from alert.exam_cat ec where ec.id_content='" & l_a_distinct_ec(i) & "')"

                    Dim cmd_update_analysis_cat As New OracleCommand(sql_update_analysis_cat, i_conn)
                    cmd_update_analysis_cat.CommandType = CommandType.Text

                    cmd_update_analysis_cat.ExecuteNonQuery()

                    cmd_update_analysis_cat.Dispose()

                    '2.2 - Uma vez que existe no ALERT, verificar se exsite tradução para a lingua da instituição
                ElseIf Not CHECK_RECORD_TRANSLATION_EXISTENCE(i_institution, l_a_distinct_ec(i), "r.code_exam_cat) from alert.exam_cat", i_conn) Then

                    Dim l_code_ec_default As String = GET_CODE_EXAM_CAT_DEFAULT(l_a_distinct_ec(i), i_conn)
                    Dim l_code_ec_alert As String = GET_CODE_EXAM_CAT_ALERT(l_a_distinct_ec(i), i_conn)

                    If Not db_access_general.SET_TRANSLATION((l_id_language), (l_code_ec_alert), (db_access_general.GET_DEFAULT_TRANSLATION(l_id_language, l_code_ec_default, i_conn)), (i_conn)) Then

                        MsgBox("ERROR INSERTING EXAM_CAT TRANSLATION - LABS_API >> CHECK_RECORD_TRANSLATION_EXISTENCE >> SET_TRANSLATION " & l_id_language)
                        Return False

                    End If

                    '2.3 - Umvez que existe no alert e existe tradução, verificar se tradução do alert é igual à do default
                ElseIf Not db_access_general.CHECK_TRANSLATIONS(l_id_language, GET_CODE_EXAM_CAT_DEFAULT(l_a_distinct_ec(i), i_conn), GET_CODE_EXAM_CAT_ALERT(l_a_distinct_ec(i), i_conn), i_conn) Then

                    Dim l_code_ec_default As String = GET_CODE_EXAM_CAT_DEFAULT(l_a_distinct_ec(i), i_conn)
                    Dim l_code_ec_alert As String = GET_CODE_EXAM_CAT_ALERT(l_a_distinct_ec(i), i_conn)

                    If Not db_access_general.SET_TRANSLATION((l_id_language), (l_code_ec_alert), (db_access_general.GET_DEFAULT_TRANSLATION(l_id_language, l_code_ec_default, i_conn)), (i_conn)) Then

                        MsgBox("ERROR INSERTING EXAM_CAT TRANSLATION - LABS_API >> CHECK_RECORD_TRANSLATION_EXISTENCE >> SET_TRANSLATION" & l_id_language)
                        Return False

                    End If

                End If
            Next

        Catch ex As Exception

            Return False

        End Try

        Return True

    End Function

    Function SET_SAMPLE_TYPE(ByVal i_institution As Int64, ByVal i_selected_default_analysis() As analysis_default, ByVal i_conn As OracleConnection) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution, i_conn)

        Try

            For i As Integer = 0 To i_selected_default_analysis.Count() - 1

                '1 - VErificar se sample_type já existe no alert. Se não existir, inserir, e inserir tradução.
                If Not CHECK_RECORD_EXISTENCE(i_selected_default_analysis(i).id_content_sample_type, "alert.sample_type", i_conn) Then

                    ''1.1 - Obter Rank, Gender. Age_min e Age_max de Sample_Type no default
                    Dim dr As OracleDataReader

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
                    If Not GET_DEFAULT_ST_PARAMETERS(i_selected_default_analysis(i).id_content_sample_type, i_conn, dr) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

                        MsgBox("ERROR GETTING SAMPLE_TYPE PARAMETERS >> SET_SAMPLE_TYPE", vbCritical)
                        dr.Dispose()
                        dr.Close()

                        Return False
                    End If

                    Dim l_gender As String = ""
                    Dim l_age_min As Int16 = -1
                    Dim l_age_max As Int16 = -1

                    While dr.Read()

                        Try

                            l_gender = dr.Item(0)

                        Catch ex As Exception

                            l_gender = ""

                        End Try

                        Try

                            l_age_min = dr.Item(1)

                        Catch ex As Exception

                            l_age_min = -1

                        End Try

                        Try

                            l_age_max = dr.Item(2)

                        Catch ex As Exception

                            l_age_max = -1

                        End Try

                    End While

                    dr.Dispose()
                    dr.Close()

                    ''1.2 - Inserir SAMPLE_TYPE

                    Dim sql_insert_st As String = "begin
                                                   insert into alert.sample_type (ID_SAMPLE_TYPE, CODE_SAMPLE_TYPE, FLG_AVAILABLE, RANK, GENDER, AGE_MIN, AGE_MAX, ID_CONTENT)
                                                   values (alert.seq_sample_type.nextval, 'SAMPLE_TYPE.CODE_SAMPLE_TYPE.' || alert.seq_sample_type.nextval, 'Y', 0, "

                    If l_gender = "" Then

                        sql_insert_st = sql_insert_st & "null, "

                    Else

                        sql_insert_st = sql_insert_st & "'" & l_gender & "', "

                    End If

                    If l_age_min = -1 Then

                        sql_insert_st = sql_insert_st & "null, "

                    Else

                        sql_insert_st = sql_insert_st & l_age_min & ", "

                    End If

                    If l_age_max = -1 Then

                        sql_insert_st = sql_insert_st & "null, "

                    Else

                        sql_insert_st = sql_insert_st & l_age_max & ", "

                    End If


                    sql_insert_st = sql_insert_st & "'" & i_selected_default_analysis(i).id_content_sample_type & "' );
                                end;"

                    Dim cmd_insert_st As New OracleCommand(sql_insert_st, i_conn)
                    cmd_insert_st.CommandType = CommandType.Text
                    cmd_insert_st.ExecuteNonQuery()

                    cmd_insert_st.Dispose()

                    '1.3 - Inserir tradução do sample_type
                    Dim l_code_st_default As String = GET_CODE_SAMPLE_TYPE_DEFAULT(i_selected_default_analysis(i).id_content_sample_type, i_conn) ''IMPORTANTE: ALTERAR FUNÇÂO PARA RECEBER CONN
                    Dim l_code_st_alert As String = GET_CODE_SAMPLE_TYPE_ALERT(i_selected_default_analysis(i).id_content_sample_type, i_conn) ''IMPORTANTE: ALTERAR FUNÇÂO PARA RECEBER CONN

                    If Not db_access_general.SET_TRANSLATION((l_id_language), (l_code_st_alert), (db_access_general.GET_DEFAULT_TRANSLATION(l_id_language, l_code_st_default, i_conn)), (i_conn)) Then

                        MsgBox("ERROR INSERTING SAMPLE_TYPE TRANSLATION - LABS_API >> CHECK_SAMPLE_TYPE_EXISTENCE >> SET_TRANSLATION")

                        Return False

                    End If

                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    ''Pensar na função de atualziar todas as tabelas relacionadas com o sample type
                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''


                    ''2 - Uma vez que já existe, verificar se tem tradução. Se não tem, inserir.
                ElseIf Not CHECK_RECORD_TRANSLATION_EXISTENCE(i_institution, i_selected_default_analysis(i).id_content_sample_type, "r.code_sample_type) from alert.sample_type", i_conn) Then

                    Dim l_code_st_default As String = GET_CODE_SAMPLE_TYPE_DEFAULT(i_selected_default_analysis(i).id_content_sample_type, i_conn)
                    Dim l_code_st_alert As String = GET_CODE_SAMPLE_TYPE_ALERT(i_selected_default_analysis(i).id_content_sample_type, i_conn)

                    If Not db_access_general.SET_TRANSLATION((l_id_language), (l_code_st_alert), (db_access_general.GET_DEFAULT_TRANSLATION(l_id_language, l_code_st_default, i_conn)), (i_conn)) Then

                        MsgBox("ERROR INSERTING SAMPLE_TYPE TRANSLATION - LABS_API >> CHECK_SAMPLE_TYPE_TRANSLATION_EXISTENCE >> SET_TRANSLATION")

                        Return False

                    End If

                    '3 - Umvez que existe no alert e existe tradução, verificar se tradução do alert é igual à do default
                ElseIf Not db_access_general.CHECK_TRANSLATIONS(l_id_language, GET_CODE_SAMPLE_TYPE_DEFAULT(i_selected_default_analysis(i).id_content_sample_type, i_conn), GET_CODE_SAMPLE_TYPE_ALERT(i_selected_default_analysis(i).id_content_sample_type, i_conn), i_conn) Then

                    Dim l_code_st_default As String = GET_CODE_SAMPLE_TYPE_DEFAULT(i_selected_default_analysis(i).id_content_sample_type, i_conn)
                    Dim l_code_st_alert As String = GET_CODE_SAMPLE_TYPE_ALERT(i_selected_default_analysis(i).id_content_sample_type, i_conn)

                    If Not db_access_general.SET_TRANSLATION((l_id_language), (l_code_st_alert), (db_access_general.GET_DEFAULT_TRANSLATION(l_id_language, l_code_st_default, i_conn)), (i_conn)) Then

                        MsgBox("ERROR INSERTING SAMPPLE_TYPE TRANSLATION - LABS_API >> CHECK_TRANSLATIONS >> SET_TRANSLATION" & l_id_language)
                        Return False

                    End If

                End If

            Next

        Catch ex As Exception

            Return False

        End Try

        Return True

    End Function

    Function SET_ANALYSIS(ByVal i_institution As Int64, ByVal i_selected_default_analysis() As analysis_default, ByVal i_conn As OracleConnection) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution, i_conn)

        Try
            For i As Integer = 0 To i_selected_default_analysis.Count() - 1

                '1- VErificar se sample_type já existe no alert. Se não existir, inserir, e inserir tradução.
                If Not CHECK_RECORD_EXISTENCE(i_selected_default_analysis(i).id_content_analysis, "alert.analysis", i_conn) Then

                    Dim l_cpt_code As String = ""
                    Dim l_gender As String = ""
                    Dim l_age_min As Int16 = -1
                    Dim l_age_max As Int16 = -1
                    Dim l_mdm_coding As Int64 = -1
                    Dim l_ref_form_code As String = ""
                    Dim l_id_content_st As String = ""
                    Dim l_barcode As String = ""

                    Dim dr As OracleDataReader

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
                    If Not GET_DEFAULT_ANALYSIS_PARAMETERS(i_selected_default_analysis(i).id_content_analysis, i_conn, dr) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

                        MsgBox("ERROR GETTING ANALYSIS PARAMETERS >> SET_ANALYSIS", vbCritical)
                        dr.Dispose()
                        dr.Close()
                        Return False

                    End If

                    '1.1.1 - Obter os parâmetros da análise
                    While dr.Read()

                        Try

                            l_cpt_code = dr.Item(0)

                        Catch ex As Exception

                            l_cpt_code = ""

                        End Try

                        Try

                            l_gender = dr.Item(1)

                        Catch ex As Exception

                            l_gender = ""

                        End Try

                        Try

                            l_age_min = dr.Item(2)

                        Catch ex As Exception

                            l_age_min = -1

                        End Try

                        Try

                            l_age_max = dr.Item(3)

                        Catch ex As Exception

                            l_age_max = -1

                        End Try

                        Try

                            l_mdm_coding = dr.Item(4)

                        Catch ex As Exception

                            l_mdm_coding = -1

                        End Try

                        Try

                            l_ref_form_code = dr.Item(5)

                        Catch ex As Exception

                            l_ref_form_code = ""

                        End Try

                        Try

                            l_id_content_st = dr.Item(6)

                        Catch ex As Exception

                            l_id_content_st = ""

                        End Try

                        Try

                            l_barcode = dr.Item(7)

                        Catch ex As Exception

                            l_barcode = ""

                        End Try

                    End While

                    dr.Dispose()
                    dr.Close()

                    ' 1.1.2 - Obter o od_alert do sample_type

                    Dim l_id_st As Int64 = -1
                    If l_id_content_st <> "" Then

                        l_id_st = GET_ID_SAMPLE_TYPE_ALERT(l_id_content_st, i_conn)

                    End If

                    '1.1.3 - Inserir análise
                    Dim sql_insert_a As String = "begin
                                                  insert into alert.analysis (ID_ANALYSIS, CODE_ANALYSIS, FLG_AVAILABLE, RANK, ID_SAMPLE_TYPE, GENDER, AGE_MIN, AGE_MAX, MDM_CODING, CPT_CODE, REF_FORM_CODE, ID_CONTENT, BARCODE)
                                                  values (alert.seq_analysis.nextval, 'ANALYSIS.CODE_ANALYSIS.' || alert.seq_analysis.nextval, 'Y', 0, "

                    If l_id_st = -1 Then

                        sql_insert_a = sql_insert_a & "null, "

                    Else

                        sql_insert_a = sql_insert_a & l_id_st & ", "

                    End If

                    If l_gender = "" Then

                        sql_insert_a = sql_insert_a & "null, "

                    Else

                        sql_insert_a = sql_insert_a & "'" & l_gender & "', "

                    End If

                    If l_age_min = -1 Then

                        sql_insert_a = sql_insert_a & "null, "

                    Else

                        sql_insert_a = sql_insert_a & l_age_min & ", "

                    End If

                    If l_age_max = -1 Then

                        sql_insert_a = sql_insert_a & "null, "

                    Else

                        sql_insert_a = sql_insert_a & l_age_max & ", "

                    End If


                    If l_mdm_coding = -1 Then

                        sql_insert_a = sql_insert_a & "null, "

                    Else

                        sql_insert_a = sql_insert_a & l_mdm_coding & ", "

                    End If


                    If l_cpt_code = "" Then

                        sql_insert_a = sql_insert_a & "null, "

                    Else

                        sql_insert_a = sql_insert_a & "'" & l_cpt_code & "', "

                    End If

                    If l_ref_form_code = "" Then

                        sql_insert_a = sql_insert_a & "null, "

                    Else

                        sql_insert_a = sql_insert_a & "'" & l_ref_form_code & "', "

                    End If

                    sql_insert_a = sql_insert_a & "'" & i_selected_default_analysis(i).id_content_analysis & "', "

                    If l_barcode = "" Then

                        sql_insert_a = sql_insert_a & "null); end; "

                    Else

                        sql_insert_a = sql_insert_a & "'" & l_barcode & "'); end;"

                    End If

                    Dim cmd_insert_st As New OracleCommand(sql_insert_a, i_conn)
                    cmd_insert_st.CommandType = CommandType.Text

                    cmd_insert_st.ExecuteNonQuery()

                    cmd_insert_st.Dispose()

                    ''Inserir tradução
                    Dim l_code_analysis_default As String = GET_CODE_ANALYSIS_DEFAULT(i_selected_default_analysis(i).id_content_analysis, i_conn)
                    Dim l_code_analysis_alert As String = GET_CODE_ANALYSIS_ALERT(i_selected_default_analysis(i).id_content_analysis, i_conn)

                    If Not db_access_general.SET_TRANSLATION((l_id_language), (l_code_analysis_alert), (db_access_general.GET_DEFAULT_TRANSLATION(l_id_language, l_code_analysis_default, i_conn)), (i_conn)) Then

                        MsgBox("ERROR INSERTING ANALYSIS TRANSLATION - LABS_API >> CHECK_ANALYSIS_EXISTENCE >> SET_TRANSLATION")

                        Return False

                    End If

                    '2 - Registo já existe no ALERT. Verifica se tem tradução, se não tiver, insere!
                ElseIf Not CHECK_RECORD_TRANSLATION_EXISTENCE(i_institution, i_selected_default_analysis(i).id_content_analysis, "r.code_analysis) from alert.analysis", i_conn) Then

                    Dim l_code_analysis_default As String = GET_CODE_ANALYSIS_DEFAULT(i_selected_default_analysis(i).id_content_analysis, i_conn)
                    Dim l_code_analysis_alert As String = GET_CODE_ANALYSIS_ALERT(i_selected_default_analysis(i).id_content_analysis, i_conn)

                    If Not db_access_general.SET_TRANSLATION((l_id_language), (l_code_analysis_alert), (db_access_general.GET_DEFAULT_TRANSLATION(l_id_language, l_code_analysis_default, i_conn)), (i_conn)) Then

                        MsgBox("ERROR INSERTING ANALYSIS TRANSLATION - LABS_API >> CHECK_ANALYSIS_TRANSLATION_EXISTENCE >> SET_TRANSLATION")

                        Return False

                    End If

                    '3 - Umvez que existe no alert e existe tradução, verificar se tradução do alert é igual à do default
                ElseIf Not db_access_general.CHECK_TRANSLATIONS(l_id_language, GET_CODE_ANALYSIS_DEFAULT(i_selected_default_analysis(i).id_content_analysis, i_conn), GET_CODE_ANALYSIS_ALERT(i_selected_default_analysis(i).id_content_analysis, i_conn), i_conn) Then

                    Dim l_code_analysis_default As String = GET_CODE_ANALYSIS_DEFAULT(i_selected_default_analysis(i).id_content_analysis, i_conn)
                    Dim l_code_analysis_alert As String = GET_CODE_ANALYSIS_ALERT(i_selected_default_analysis(i).id_content_analysis, i_conn)

                    If Not db_access_general.SET_TRANSLATION((l_id_language), (l_code_analysis_alert), (db_access_general.GET_DEFAULT_TRANSLATION(l_id_language, l_code_analysis_default, i_conn)), (i_conn)) Then

                        MsgBox("ERROR INSERTING SAMPPLE_TYPE TRANSLATION - LABS_API >> CHECK_TRANSLATIONS >> SET_TRANSLATION" & l_id_language)
                        Return False

                    End If

                End If

            Next

        Catch ex As Exception

            Return False

        End Try

        Return True

    End Function

    Function SET_ANALYSIS_SAMPLE_TYPE(ByVal i_institution As Int64, ByVal i_selected_default_analysis() As analysis_default, ByVal i_conn As OracleConnection) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution, i_conn)

        Try
            For i As Integer = 0 To i_selected_default_analysis.Count() - 1

                '1 - Verificar se AST já existe no alert (Nesta etapa já se confirmou a existência da análise e do sample_type)
                If Not CHECK_RECORD_EXISTENCE(i_selected_default_analysis(i).id_content_analysis_sample_type, "alert.analysis_sample_type", i_conn) Then

                    Dim l_gender As String = ""
                    Dim l_age_min As Int16 = -1
                    Dim l_age_max As Int16 = -1

                    Dim dr As OracleDataReader

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
                    If Not GET_DEFAULT_ANALYSIS_ST_PARAMETERS(i_selected_default_analysis(i).id_content_analysis_sample_type, i_conn, dr) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

                        MsgBox("ERROR GETTING ANALYSIS_SAMPLE_TYPE PARAMETERS >> SET_ANALYSIS_SAMPLE_TYPE")
                        dr.Dispose()
                        dr.Close()
                        Return False

                    End If

                    '1.1.1 - Obter os parâmetros da análise_SAMPLE_TYPE
                    While dr.Read()

                        Try

                            l_gender = dr.Item(0)

                        Catch ex As Exception

                            l_gender = ""

                        End Try

                        Try

                            l_age_min = dr.Item(1)

                        Catch ex As Exception

                            l_age_min = -1

                        End Try

                        Try

                            l_age_max = dr.Item(2)

                        Catch ex As Exception

                            l_age_max = -1

                        End Try

                    End While

                    dr.Dispose()
                    dr.Close()

                    ''1.1.2  - Obter o ID ALERT da análise
                    Dim l_id_analysis As Int64 = GET_ID_ANALYSIS_ALERT(i_selected_default_analysis(i).id_content_analysis, i_conn)

                    ''1.1.3 - Obter o ID ALERT do sample_type
                    Dim l_id_sample_type As Int64 = GET_ID_SAMPLE_TYPE_ALERT(i_selected_default_analysis(i).id_content_sample_type, i_conn)

                    ''1.1.4 - Inserir AST
                    Dim sql_insert_ast As String = "begin
                                                    insert into alert.analysis_sample_type (ID_ANALYSIS, ID_SAMPLE_TYPE,ID_CONTENT, ID_CONTENT_ANALYSIS, ID_CONTENT_SAMPLE_TYPE, GENDER, AGE_MIN, AGE_MAX, FLG_AVAILABLE)
                                                    values (" & l_id_analysis & ", " & l_id_sample_type & ", '" & i_selected_default_analysis(i).id_content_analysis_sample_type & "', '" & i_selected_default_analysis(i).id_content_analysis & "', '" & i_selected_default_analysis(i).id_content_sample_type & "', "

                    If l_gender = "" Then

                        sql_insert_ast = sql_insert_ast & "null, "

                    Else

                        sql_insert_ast = sql_insert_ast & "'" & l_gender & "', "

                    End If

                    If l_age_min = -1 Then

                        sql_insert_ast = sql_insert_ast & "null, "

                    Else

                        sql_insert_ast = sql_insert_ast & l_age_min & ", "

                    End If

                    If l_age_max = -1 Then

                        sql_insert_ast = sql_insert_ast & "null, "

                    Else

                        sql_insert_ast = sql_insert_ast & l_age_max & ", "

                    End If

                    sql_insert_ast = sql_insert_ast & "'Y'); end;"

                    Try

                        Dim cmd_insert_ast As New OracleCommand(sql_insert_ast, i_conn)
                        cmd_insert_ast.CommandType = CommandType.Text

                        cmd_insert_ast.ExecuteNonQuery()

                        cmd_insert_ast.Dispose()

                    Catch ex As Exception 'Se não der para introduzir, seginfica que já existe mas esta a Not available. Assim, colocar a 'Y'

                        Dim sql_update_ast = "update alert.analysis_sample_type ast
                                              set ast.flg_available='Y'
                                              where ast.id_content='" & i_selected_default_analysis(i).id_content_analysis_sample_type & "'
                                              and ast.id_content_analysis='" & i_selected_default_analysis(i).id_content_analysis & "'
                                              and ast.id_content_sample_type='" & i_selected_default_analysis(i).id_content_sample_type & "'"

                        Dim cmd_update_ast As New OracleCommand(sql_update_ast, i_conn)
                        cmd_update_ast.CommandType = CommandType.Text

                        cmd_update_ast.ExecuteNonQuery()

                        cmd_update_ast.Dispose()

                    End Try

                    ''1.1.5 - Inserir Tradução da AST
                    'Nota: Se só se tiver feito o update, a tradução pode existir, daí a verificação

                    If Not CHECK_RECORD_TRANSLATION_EXISTENCE(i_institution, i_selected_default_analysis(i).id_content_analysis_sample_type, "r.code_analysis_sample_type) from alert.analysis_sample_type", i_conn) Then

                        Dim l_code_analysis_st_default As String = GET_CODE_ANALYSIS_ST_DEFAULT(i_selected_default_analysis(i).id_content_analysis_sample_type, i_conn)
                        Dim l_code_analysis_st_alert As String = GET_CODE_ANALYSIS_ST_ALERT(i_selected_default_analysis(i).id_content_analysis_sample_type, i_conn)
                        If Not db_access_general.SET_TRANSLATION((l_id_language), (l_code_analysis_st_alert), (db_access_general.GET_DEFAULT_TRANSLATION(l_id_language, l_code_analysis_st_default, i_conn)), (i_conn)) Then

                            MsgBox("ERROR INSERTING ANALYSIS SAMPLE TYPE TRANSLATION - LABS_API >> CHECK_ANALYSIS_SAMPLE_TYPE_EXISTENCE >> SET_TRANSLATION")

                            Return False

                        End If

                    End If

                    '2 Verificar se existe tradução. Se não existir, inserir.
                ElseIf Not CHECK_RECORD_TRANSLATION_EXISTENCE(i_institution, i_selected_default_analysis(i).id_content_analysis_sample_type, "R.code_analysis_sample_type) from alert.analysis_sample_type", i_conn) Then

                    Dim l_code_analysis_st_default As String = GET_CODE_ANALYSIS_ST_DEFAULT(i_selected_default_analysis(i).id_content_analysis_sample_type, i_conn)
                    Dim l_code_analysis_st_alert As String = GET_CODE_ANALYSIS_ST_ALERT(i_selected_default_analysis(i).id_content_analysis_sample_type, i_conn)
                    If Not db_access_general.SET_TRANSLATION((l_id_language), (l_code_analysis_st_alert), (db_access_general.GET_DEFAULT_TRANSLATION(l_id_language, l_code_analysis_st_default, i_conn)), (i_conn)) Then

                        MsgBox("ERROR INSERTING ANALYSIS SAMPLE TYPE TRANSLATION - LABS_API >> CHECK_ANALYSIS_ST_TRANSLATION_EXISTENCE >> SET_TRANSLATION")

                        Return False

                    End If


                    '3 - Uma vez que existe no alert e existe tradução, verificar se tradução do alert é igual à do default ----------------------------------------------
                ElseIf Not db_access_general.CHECK_TRANSLATIONS(l_id_language, GET_CODE_ANALYSIS_ST_DEFAULT(i_selected_default_analysis(i).id_content_analysis_sample_type, i_conn), GET_CODE_ANALYSIS_ST_ALERT(i_selected_default_analysis(i).id_content_analysis_sample_type, i_conn), i_conn) Then

                    Dim l_code_analysis_st_default As String = GET_CODE_ANALYSIS_ST_DEFAULT(i_selected_default_analysis(i).id_content_analysis_sample_type, i_conn)
                    Dim l_code_analysis_st_alert As String = GET_CODE_ANALYSIS_ST_ALERT(i_selected_default_analysis(i).id_content_analysis_sample_type, i_conn)

                    If Not db_access_general.SET_TRANSLATION((l_id_language), (l_code_analysis_st_alert), (db_access_general.GET_DEFAULT_TRANSLATION(l_id_language, l_code_analysis_st_default, i_conn)), (i_conn)) Then

                        MsgBox("ERROR INSERTING ANALYSIS_SAPPLE_TYPE TRANSLATION - LABS_API >> CHECK_TRANSLATIONS >> SET_TRANSLATION" & l_id_language)
                        Return False

                    End If

                End If

            Next

        Catch ex As Exception

            Return False

        End Try

        Return True

    End Function

    Function SET_PARAMETER(ByVal i_institution As Int64, ByVal i_software As Int16, ByVal i_selected_default_analysis() As analysis_default, ByVal i_conn As OracleConnection) As Boolean


        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution, i_conn)

        Try

            For ii As Integer = 0 To i_selected_default_analysis.Count() - 1

                'Array que vai guardar os id_contents dos parâmetros do analysis_sample_type
                Dim l_array_parameters() As String
                Dim l_index As Int16 = 0

                '1.1 - Obter os id_contents dos parâmetros da ast
                Dim dr_parameters As OracleDataReader

#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
                If Not GET_ANALYSIS_PARAMETERS_ID_CONTENT_DEFAULT(i_software, i_selected_default_analysis(ii).id_content_analysis_sample_type, i_conn, dr_parameters) Then
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value

                    MsgBox("ERROR GETTING ANALYSIS PARAMETERS ID_CONTENT >> SET_PARAMETER")
                    dr_parameters.Dispose()
                    dr_parameters.Close()
                    Return False

                End If

                ReDim l_array_parameters(0)

                While dr_parameters.Read()

                    ReDim Preserve l_array_parameters(l_index)
                    l_array_parameters(l_index) = dr_parameters.Item(0)
                    l_index = l_index + 1

                End While

                dr_parameters.Dispose()
                dr_parameters.Close()

                '1.2 - Verificar se parâmetros existem no ALERT.
                For i As Integer = 0 To l_array_parameters.Count() - 1


                    '1.2.1 - Inserir registo
                    'cHECKAR SE EXISTEM NO ALERT -  Se não existir, insere, e insere tradução
                    If Not CHECK_RECORD_EXISTENCE(l_array_parameters(i), "alert.analysis_parameter", i_conn) Then

                        Dim sql_parameter As String = "begin
                                                        insert into alert.analysis_parameter (ID_ANALYSIS_PARAMETER, CODE_ANALYSIS_PARAMETER, RANK, FLG_AVAILABLE, ID_CONTENT)
                                                        values (alert.seq_analysis_parameter.nextval, 'ANALYSIS_PARAMETER.CODE_ANALYSIS_PARAMETER.' || seq_analysis_parameter.nextval, 0, 'Y', '" & l_array_parameters(i) & "');
                                                        end;"

                        Dim cmd_insert_parameter As New OracleCommand(sql_parameter, i_conn)
                        cmd_insert_parameter.CommandType = CommandType.Text

                        cmd_insert_parameter.ExecuteNonQuery()
                        cmd_insert_parameter.Dispose()

                        '1.2.2 inserir tradução
                        Dim l_code_analysis_parameter_default As String = GET_CODE_PARAMETER_DEFAULT(l_array_parameters(i), i_conn)
                        Dim l_code_analysis_parameter_alert As String = GET_CODE_PARAMETER_ALERT(l_array_parameters(i), i_conn)

                        If Not db_access_general.SET_TRANSLATION((l_id_language), (l_code_analysis_parameter_alert), (db_access_general.GET_DEFAULT_TRANSLATION(l_id_language, l_code_analysis_parameter_default, i_conn)), (i_conn)) Then

                            MsgBox("ERROR INSERTING PARAMETER TRANSLATION - LABS_API >> CHECK_ANALYSIS_ST_TRANSLATION_EXISTENCE >> SET_TRANSLATION  (Record Verification)")
                            Return False

                        End If

                        '1.2.3 - Como houve uma inserção de um novo parametro, atualizar a analysis_param
                        'Ou seja, parametero que estava inativo e passa a haver  um registo novo activo com o mesmo id_content
                        Dim sql_analysis_param As String = "UPDATE alert.analysis_param ap
                                                            SET ap.id_analysis_parameter =
                                                                (SELECT pn.id_analysis_parameter
                                                                 FROM alert.analysis_parameter pn
                                                                 WHERE pn.id_content = '" & l_array_parameters(i) & "'
                                                                 AND pn.flg_available = 'Y')
                                                            WHERE ap.id_analysis_parameter IN (SELECT po.id_analysis_parameter
                                                                                               FROM alert.analysis_parameter po
                                                                                               WHERE po.id_content = '" & l_array_parameters(i) & "'
                                                                                               AND po.flg_available = 'N')"

                        Dim cmd_analysis_param As New OracleCommand(sql_analysis_param, i_conn)
                        cmd_analysis_param.CommandType = CommandType.Text

                        cmd_analysis_param.ExecuteNonQuery()
                        cmd_analysis_param.Dispose()


                        '1.3 - Verificar se existe tradução. Se não exsitir, insere.
                    ElseIf Not CHECK_RECORD_TRANSLATION_EXISTENCE(i_institution, l_array_parameters(i), "r.code_analysis_parameter) from alert.analysis_parameter", i_conn) Then

                        Dim l_code_analysis_parameter_default As String = GET_CODE_PARAMETER_DEFAULT(l_array_parameters(i), i_conn)
                        Dim l_code_analysis_parameter_alert As String = GET_CODE_PARAMETER_ALERT(l_array_parameters(i), i_conn)

                        If Not db_access_general.SET_TRANSLATION((l_id_language), (l_code_analysis_parameter_alert), (db_access_general.GET_DEFAULT_TRANSLATION(l_id_language, l_code_analysis_parameter_default, i_conn)), (i_conn)) Then

                            MsgBox("ERROR INSERTING PARAMETER TRANSLATION - LABS_API >> CHECK_ANALYSIS_ST_TRANSLATION_EXISTENCE >> SET_TRANSLATION (Translation Verification)")
                            MsgBox(i_selected_default_analysis(ii).id_content_analysis_sample_type)
                            Return False

                        End If

                    ElseIf Not db_access_general.CHECK_TRANSLATIONS(l_id_language, GET_CODE_PARAMETER_DEFAULT(l_array_parameters(i), i_conn), GET_CODE_PARAMETER_ALERT(l_array_parameters(i), i_conn), i_conn) Then

                        Dim l_code_parameter_default As String = GET_CODE_PARAMETER_DEFAULT(l_array_parameters(i), i_conn)
                        Dim l_code_parameter_alert As String = GET_CODE_PARAMETER_ALERT(l_array_parameters(i), i_conn)

                        If Not db_access_general.SET_TRANSLATION((l_id_language), (l_code_parameter_alert), (db_access_general.GET_DEFAULT_TRANSLATION(l_id_language, l_code_parameter_default, i_conn)), (i_conn)) Then

                            MsgBox("ERROR INSERTING PARAMETERTRANSLATION - LABS_API >> CHECK_TRANSLATIONS >> SET_TRANSLATION (update Verification" & l_id_language)
                            Return False

                        End If

                    End If

                Next

                '1.3 - Ver se é necessário fazer update aos registos do ALERT (Registos que no default estão a N e no alert estão a Y)
                If Not UPDATE_PARAMETER_AVAILABILITY(i_software, i_selected_default_analysis(ii).id_content_analysis_sample_type, i_conn) Then

                    MsgBox("ERROR UPDATING PARAMETERS!", vbCritical)

                End If

            Next

        Catch ex As Exception

            Return False

        End Try

        Return True

    End Function

    Function UPDATE_PARAMETER_AVAILABILITY(ByVal i_id_software As Integer, ByVal i_id_ast_content As String, ByVal i_conn As OracleConnection) As Boolean

        Try

            Dim sql_parameter As String = "begin
                                                UPDATE alert.analysis_parameter app
                                                SET app.flg_available = 'N'
                                                WHERE app.id_content IN (SELECT DISTINCT aparam.id_content
                                                                            FROM alert_default.analysis_sample_type ast
                                                                            JOIN alert_default.analysis_param ap ON ap.id_analysis = ast.id_analysis
                                                                                                     AND ap.id_sample_type = ast.id_sample_type
                                                                                                     AND ap.id_software = " & i_id_software & "
                                                                            JOIN alert_default.analysis_parameter aparam ON aparam.id_analysis_parameter = ap.id_analysis_parameter
                                                                            AND aparam.flg_available = 'N'                            
                                                                            WHERE ast.id_content = '" & i_id_ast_content & "');
                                                 end;"


            Dim cmd_update_parameter As New OracleCommand(sql_parameter, i_conn)
            cmd_update_parameter.CommandType = CommandType.Text
            cmd_update_parameter.ExecuteNonQuery()
            cmd_update_parameter.Dispose()

            cmd_update_parameter.Dispose()

            Return True

        Catch ex As Exception

            Return False

        End Try

        Return True

    End Function

    Function SET_PARAM(ByVal i_institution As Int64, ByVal i_software As Int16, ByVal i_selected_default_analysis() As analysis_default, ByVal i_conn As OracleConnection) As Boolean

        Dim dr As OracleDataReader

        Try

            'Ciclo que vai correr as AST todas enviadas à função
            For i As Integer = 0 To i_selected_default_analysis.Count() - 1

                If Not GET_DEFAULT_PARAMETERS(i_selected_default_analysis(i).id_content_analysis_sample_type, i_software, i_conn, dr) Then

                    MsgBox("ERROR INSERTING ANALYSIS_PARAM", vbCritical)
                    dr.Dispose()
                    dr.Close()
                    Return False

                Else

                    Dim l_id_content_parameter As String = ""
                    Dim l_color_graph As String = ""
                    Dim l_flg_fil_type As String = ""
                    Dim l_rank As Int16 = -1

                    While dr.Read()

                        l_id_content_parameter = dr.Item(0)

                        Try

                            l_color_graph = dr.Item(1)

                        Catch ex As Exception

                            l_color_graph = ""

                        End Try

                        Try

                            l_flg_fil_type = dr.Item(2)

                        Catch ex As Exception

                            l_flg_fil_type = ""

                        End Try

                        Try

                            l_rank = dr.Item(3)

                        Catch ex As Exception

                            l_rank = 0

                        End Try

                        Dim sql_insert_analysis_param As String = "DECLARE

                                                                        l_id_analysis    alert.analysis_sample_type.id_analysis%TYPE;
                                                                        l_id_sample_type alert.analysis_sample_type.id_sample_type%TYPE;

                                                                        l_id_analysis_parameter alert.analysis_param.id_analysis_parameter%TYPE;

                                                                    BEGIN

                                                                        SELECT ast.id_analysis, ast.id_sample_type
                                                                        INTO l_id_analysis, l_id_sample_type
                                                                        FROM alert.analysis_sample_type ast
                                                                        WHERE ast.flg_available = 'Y'
                                                                        AND ast.id_content = '" & i_selected_default_analysis(i).id_content_analysis_sample_type & "';

                                                                        SELECT p.id_analysis_parameter
                                                                        INTO l_id_analysis_parameter
                                                                        FROM alert.analysis_parameter p
                                                                        WHERE p.id_content = '" & l_id_content_parameter & "'
                                                                        AND p.flg_available = 'Y';

                                                                        BEGIN
    
                                                                            INSERT INTO alert.analysis_param
                                                                                (id_analysis_param,
                                                                                 id_analysis,
                                                                                 flg_available,
                                                                                 id_institution,
                                                                                 id_software,
                                                                                 id_analysis_parameter,
                                                                                 rank,
                                                                                 color_graph,
                                                                                 flg_fill_type,
                                                                                 id_sample_type)
                                                                            VALUES
                                                                                (alert.seq_analysis_param.nextval,
                                                                                 l_id_analysis,
                                                                                 'Y',
                                                                                 " & i_institution & ",
                                                                                 " & i_software & ",
                                                                                 l_id_analysis_parameter,
                                                                                 " & l_rank & ","

                        If l_color_graph = "" Then

                            sql_insert_analysis_param = sql_insert_analysis_param & " null ,"

                        Else

                            sql_insert_analysis_param = sql_insert_analysis_param & " '" & l_color_graph & "',"

                        End If

                        sql_insert_analysis_param = sql_insert_analysis_param & "'" & l_flg_fil_type & "',
                                                                                 l_id_sample_type);
    
                                                                        EXCEPTION
                                                                            WHEN dup_val_on_index THEN
        
                                                                                UPDATE alert.analysis_param ap
                                                                                SET ap.flg_available = 'Y'
                                                                                WHERE ap.id_analysis_parameter = (SELECT p.id_analysis_parameter
                                                                                                                  FROM alert.analysis_parameter p
                                                                                                                  WHERE p.flg_available = 'Y'
                                                                                                                  AND p.id_content = '" & l_id_content_parameter & "')
                                                                                AND ap.id_institution = " & i_institution & "
                                                                                AND ap.id_software =    " & i_software & "
                                                                                AND ap.id_analysis =    l_id_analysis
                                                                                AND ap.id_sample_type = l_id_sample_type;
        
                                                                        END;
    
                                                                    END;"


                        Dim cmd_insert_analysis_param As New OracleCommand(sql_insert_analysis_param, i_conn)

                        cmd_insert_analysis_param.CommandType = CommandType.Text

                        cmd_insert_analysis_param.ExecuteNonQuery()

                        cmd_insert_analysis_param.Dispose()

                    End While

                End If

            Next

        Catch ex As Exception

            dr.Dispose()
            dr.Close()

            MsgBox("ERROR INSERTING ANALYSIS_PARAM", vbCritical)
            Return False

        End Try

        dr.Dispose()
        dr.Close()

        Return True

    End Function

    Function SET_SAMPLE_RECIPIENT(ByVal i_institution As Int64, ByVal i_selected_default_analysis() As analysis_default, ByVal i_conn As OracleConnection) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution, i_conn)

        Try
            For i As Integer = 0 To i_selected_default_analysis.Count() - 1

                '1 - Verificar se SR já existe no alert. Se não existe, inserir e inserir tradução.
                If Not CHECK_RECORD_EXISTENCE(i_selected_default_analysis(i).id_content_sample_recipient, "alert.sample_recipient", i_conn) Then
                    Dim sql_insert_sr As String = "BEGIN

                                                            insert into ALERT.sample_recipient (ID_SAMPLE_RECIPIENT, CODE_SAMPLE_RECIPIENT, FLG_AVAILABLE, RANK, ID_CONTENT)
                                                            values (ALERT.SEQ_SAMPLE_RECIPIENT.NEXTVAL, 'SAMPLE_RECIPIENT.CODE_SAMPLE_RECIPIENT.'||  ALERT.SEQ_SAMPLE_RECIPIENT.NEXTVAL , 'Y', 0,'" & i_selected_default_analysis(i).id_content_sample_recipient & "');

                                                    EXCEPTION
                                                      WHEN DUP_VAL_ON_INDEX THEN
                                                        UPDATE ALERT.sample_recipient SR
                                                        SET    SR.FLG_AVAILABLE='Y'
                                                        WHERE  SR.ID_CONTENT='" & i_selected_default_analysis(i).id_content_sample_recipient & "';
                                                    
                                                    END;"

                    Dim cmd_insert_sr As New OracleCommand(sql_insert_sr, i_conn)
                    cmd_insert_sr.CommandType = CommandType.Text

                    cmd_insert_sr.ExecuteNonQuery()

                    cmd_insert_sr.Dispose()

                    Dim l_code_analysis_sr_default As String = GET_CODE_SAMPLE_RECIPIENT_DEFAULT(i_selected_default_analysis(i).id_content_sample_recipient, i_conn)
                    Dim l_code_analysis_sr_alert As String = GET_CODE_SAMPLE_RECIPIENT_ALERT(i_selected_default_analysis(i).id_content_sample_recipient, i_conn)

                    If Not db_access_general.SET_TRANSLATION((l_id_language), (l_code_analysis_sr_alert), (db_access_general.GET_DEFAULT_TRANSLATION(l_id_language, l_code_analysis_sr_default, i_conn)), (i_conn)) Then

                        MsgBox("ERROR INSERTING SAMPLE RECIPIENT TRANSLATION - LABS_API >> CHECK_SAMPLE_RECIPIENT_EXISTENCE >> SET_TRANSLATION")

                        Return False

                    End If

                    ''cOMO SE INSERIUM UMA NOVA SR, VERIFICAR SE É NECESSÁRIO FAZER UPDATE À ANALYSIS_INST_RECIPIENT
                    Dim sql_update_ais As String = "UPDATE alert.analysis_instit_recipient air
                                                    SET air.id_sample_recipient =
                                                        (SELECT srn.id_sample_recipient
                                                         FROM alert.sample_recipient srn
                                                         WHERE srn.id_content = '" & i_selected_default_analysis(i).id_content_sample_recipient & "'
                                                         AND srn.flg_available = 'Y')
                                                    WHERE air.id_sample_recipient IN  (SELECT srn.id_sample_recipient
                                                         FROM alert.sample_recipient srn
                                                         WHERE srn.id_content = '" & i_selected_default_analysis(i).id_content_sample_recipient & "'
                                                         AND srn.flg_available = 'N')"

                    Dim cmd_update_ais As New OracleCommand(sql_update_ais, i_conn)

                    cmd_update_ais.CommandType = CommandType.Text

                    cmd_update_ais.ExecuteNonQuery()

                    cmd_update_ais.Dispose()


                    '2 - Uma vez que já existe no alert, verificar se existe translation
                ElseIf Not CHECK_RECORD_TRANSLATION_EXISTENCE(i_institution, i_selected_default_analysis(i).id_content_analysis_sample_type, "R.code_sample_recipient) from alert.sample_recipient", i_conn) Then

                    Dim l_code_analysis_sr_default As String = GET_CODE_SAMPLE_RECIPIENT_DEFAULT(i_selected_default_analysis(i).id_content_sample_recipient, i_conn)
                    Dim l_code_analysis_sr_alert As String = GET_CODE_SAMPLE_RECIPIENT_ALERT(i_selected_default_analysis(i).id_content_sample_recipient, i_conn)

                    If Not db_access_general.SET_TRANSLATION((l_id_language), (l_code_analysis_sr_alert), (db_access_general.GET_DEFAULT_TRANSLATION(l_id_language, l_code_analysis_sr_default, i_conn)), (i_conn)) Then

                        MsgBox("ERROR INSERTING SAMPLE RECIPIENT TRANSLATION - LABS_API >> CHECK_SAMPLE_RECIPIENT_TRANSLATION_EXISTENCE >> SET_TRANSLATION")

                        Return False

                    End If

                    '3 - Uma vez que existe no alert e existe tradução, verificar se tradução do alert é igual à do default
                ElseIf Not db_access_general.CHECK_TRANSLATIONS(l_id_language, GET_CODE_SAMPLE_RECIPIENT_DEFAULT(i_selected_default_analysis(i).id_content_sample_recipient, i_conn), GET_CODE_SAMPLE_RECIPIENT_ALERT(i_selected_default_analysis(i).id_content_sample_recipient, i_conn), i_conn) Then

                    Dim l_code_analysis_sr_default As String = GET_CODE_SAMPLE_RECIPIENT_DEFAULT(i_selected_default_analysis(i).id_content_sample_recipient, i_conn)
                    Dim l_code_analysis_sr_alert As String = GET_CODE_SAMPLE_RECIPIENT_ALERT(i_selected_default_analysis(i).id_content_sample_recipient, i_conn)

                    If Not db_access_general.SET_TRANSLATION((l_id_language), (l_code_analysis_sr_alert), (db_access_general.GET_DEFAULT_TRANSLATION(l_id_language, l_code_analysis_sr_default, i_conn)), (i_conn)) Then

                        MsgBox("ERROR INSERTING SAMPLE RECIPIENT TRANSLATION - LABS_API >> CHECK_TRANSLATIONS >> SET_TRANSLATION" & l_id_language)
                        Return False

                    End If

                End If

            Next

        Catch ex As Exception


            MsgBox("ERROR INSERTING SAMPLE RECIPIENT")
            Return False

        End Try

        Return True

    End Function

    Function SET_ANALYSIS_INST_SOFT(ByVal i_institution As Int64, ByVal i_software As Integer, ByVal i_selected_default_analysis() As analysis_default, ByVal i_conn As OracleConnection) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution, i_conn)

        Try

            Dim sql_verify_ais As String = ""

            For i As Integer = 0 To i_selected_default_analysis.Count() - 1

                Dim l_count_ais As Integer = 0

                '1 - Verificar se SR já existe no alert. Se não existe, inserir e inserir tradução.
                'Nota: A verificação do registo terá que ser diferente dos casos anteriore, pois não tem ID_CONTENT, e é orientada ao software e insituição

                sql_verify_ais = "SELECT COUNT(*)
                                    FROM alert.analysis_instit_soft ais
                                    WHERE ais.id_analysis = (SELECT ast_a.id_analysis
                                                             FROM alert.analysis_sample_type ast_a
                                                             WHERE ast_a.id_content = '" & i_selected_default_analysis(i).id_content_analysis_sample_type & "'
                                                             AND ast_a.flg_available = 'Y')
                                    AND ais.id_sample_type = (SELECT ast_st.id_sample_type
                                                             FROM alert.analysis_sample_type ast_st
                                                             WHERE ast_st.id_content = '" & i_selected_default_analysis(i).id_content_analysis_sample_type & "'
                                                             AND ast_st.flg_available = 'Y')
                                    AND ais.id_institution =" & i_institution & "
                                    AND ais.id_software =" & i_software & "
                                    AND ais.flg_available = 'Y'"


                Dim cmd As New OracleCommand(sql_verify_ais, i_conn)
                cmd.CommandType = CommandType.Text

                Dim dr As OracleDataReader = cmd.ExecuteReader()

                Try

                    While dr.Read()

                        l_count_ais = dr.Item(0)

                    End While

                    dr.Dispose()
                    dr.Close()
                    cmd.Dispose()

                Catch ex As Exception

                    l_count_ais = 0

                    dr.Dispose()
                    dr.Close()
                    cmd.Dispose()

                End Try

                'Significa que não existe registo na analysis_inst_soft
                If l_count_ais = 0 Then

                    'Função que vai insrir o registo
                    Dim sql_insert_ais As String = "DECLARE

                                                        l_id_analysis_alert    alert.analysis_sample_type.id_analysis%TYPE;
                                                        l_id_sample_type_alert alert.analysis_sample_type.id_sample_type%TYPE;

                                                        l_id_analysis_default    alert.analysis_sample_type.id_analysis%TYPE;
                                                        l_id_sample_type_default alert.analysis_sample_type.id_sample_type%TYPE;

                                                        l_flg_type          alert_default.analysis_instit_soft.flg_type%TYPE;
                                                        l_flg_mov_pat       alert_default.analysis_instit_soft.flg_mov_pat%TYPE;
                                                        l_flg_first_result  alert_default.analysis_instit_soft.flg_first_result%TYPE;
                                                        l_flg_mov_recipient alert_default.analysis_instit_soft.flg_mov_recipient%TYPE;
                                                        l_flg_harvest       alert_default.analysis_instit_soft.flg_harvest%TYPE;
                                                        l_qty_harvest       alert_default.analysis_instit_soft.qty_harvest%TYPE;
                                                        l_rank              alert_default.analysis_instit_soft.rank%TYPE;
                                                        l_flg_execute       alert_default.analysis_instit_soft.flg_execute%TYPE;
                                                        l_flg_justify       alert_default.analysis_instit_soft.flg_justify%TYPE;
                                                        l_flg_interface     alert_default.analysis_instit_soft.flg_interface%TYPE;
                                                        l_flg_chargeable    alert_default.analysis_instit_soft.flg_chargeable%TYPE;
                                                        l_flg_fill_type     alert_default.analysis_instit_soft.flg_fill_type%TYPE;
                                                        l_color_text        alert_default.analysis_instit_soft.color_text%TYPE;
                                                        l_color_graph       alert_default.analysis_instit_soft.color_graph%TYPE;
                                                        l_cost              alert_default.analysis_instit_soft.cost%type;
                                                        l_price             alert_default.analysis_instit_soft.price%type;

                                                        l_id_exam_cat_alert alert.exam_cat.id_exam_cat%TYPE;

                                                    BEGIN

                                                        SELECT ast.id_analysis, ast.id_sample_type
                                                        INTO l_id_analysis_alert, l_id_sample_type_alert
                                                        FROM alert.analysis_sample_type ast
                                                        WHERE ast.id_content = '" & i_selected_default_analysis(i).id_content_analysis_sample_type & "'
                                                        AND ast.flg_available = 'Y';

                                                        SELECT ast.id_analysis, ast.id_sample_type
                                                        INTO l_id_analysis_default, l_id_sample_type_default
                                                        FROM alert_default.analysis_sample_type ast
                                                        WHERE ast.id_content = '" & i_selected_default_analysis(i).id_content_analysis_sample_type & "'
                                                        AND ast.flg_available = 'Y';
    
                                                        SELECT EC.ID_EXAM_CAT
                                                        INTO l_id_exam_cat_alert
                                                        FROM ALERT.EXAM_CAT EC
                                                        WHERE EC.ID_CONTENT='" & i_selected_default_analysis(i).id_content_category & "'
                                                        AND EC.FLG_AVAILABLE='Y';
    
                                                        SELECT ais.flg_type, ais.flg_mov_pat, ais.flg_first_result, ais.flg_mov_recipient, ais.flg_harvest,ais.qty_harvest, ais.rank, ais.flg_execute,
                                                        ais.flg_justify, ais.flg_interface, ais.flg_chargeable, ais.flg_fill_type, ais.color_text, ais.color_graph, ais.cost, ais.price
                                                        INTO l_flg_type, l_flg_mov_pat,l_flg_first_result,l_flg_mov_recipient,l_flg_harvest,l_qty_harvest, l_rank, l_flg_execute,
                                                        l_flg_justify,l_flg_interface,l_flg_chargeable,l_flg_fill_type,l_color_text, l_color_graph, l_cost, l_price
                                                        FROM alert_default.analysis_instit_soft ais
                                                        WHERE ais.id_analysis = l_id_analysis_default
                                                        AND ais.id_sample_type = l_id_sample_type_default
                                                        AND ais.flg_available = 'Y'
                                                        AND ais.id_software = " & i_software & ";

                                                        INSERT INTO alert.analysis_instit_soft
                                                            (id_analysis_instit_soft,
                                                             id_analysis,
                                                             flg_type,
                                                             id_institution,
                                                             id_software,
                                                             flg_mov_pat,
                                                             flg_first_result,
                                                             flg_mov_recipient,
                                                             flg_harvest,
                                                             id_exam_cat,
                                                             rank,
                                                             cost,
                                                             price,
                                                             color_graph,
                                                             color_text,
                                                             flg_fill_type,
                                                             id_analysis_group,
                                                             flg_execute,
                                                             flg_justify,
                                                             flg_interface,
                                                             flg_chargeable,
                                                             flg_available,                       
                                                             qty_harvest,
                                                             id_sample_type)
                                                        VALUES
                                                            (alert.seq_analysis_instit_soft.nextval,
                                                             l_id_analysis_alert,
                                                             l_flg_type,
                                                             " & i_institution & ",
                                                             " & i_software & ",
                                                             l_flg_mov_pat,
                                                             l_flg_first_result,
                                                             l_flg_mov_recipient,
                                                             l_flg_harvest,
                                                             l_id_exam_cat_alert,
                                                             l_rank,
                                                             l_cost, --COST
                                                             l_price, --PRICE
                                                             l_color_graph,
                                                             l_color_text,
                                                             l_flg_fill_type,
                                                             NULL,  --ID_ANALYSIS_GROUP
                                                             l_flg_execute,
                                                             l_flg_justify,
                                                             l_flg_interface,
                                                             l_flg_chargeable,
                                                             'Y',        
                                                             l_qty_harvest,
                                                             l_id_sample_type_alert);

                                                    EXCEPTION
                                                      WHEN DUP_VAL_ON_INDEX THEN
                                                        UPDATE ALERT.ANALYSIS_INSTIT_SOFT AIS
                                                        SET AIS.FLG_TYPE = l_flg_type, AIS.FLG_MOV_PAT=l_flg_mov_pat,
                                                        AIS.FLG_FIRST_RESULT= l_flg_first_result, AIS.FLG_MOV_RECIPIENT=l_flg_mov_recipient,
                                                        AIS.FLG_HARVEST=l_flg_harvest, AIS.ID_EXAM_CAT=l_id_exam_cat_alert, AIS.RANK=l_rank,
                                                        AIS.COST=l_cost, AIS.PRICE=l_price, AIS.COLOR_GRAPH=l_color_graph, AIS.COLOR_TEXT=l_color_text,
                                                        AIS.FLG_FILL_TYPE=l_flg_fill_type, AIS.ID_ANALYSIS_GROUP=NULL, AIS.FLG_EXECUTE=l_flg_execute,
                                                        AIS.FLG_JUSTIFY=l_flg_justify, AIS.FLG_INTERFACE=l_flg_interface, AIS.FLG_CHARGEABLE=l_flg_chargeable,
                                                        AIS.FLG_AVAILABLE='Y', AIS.QTY_HARVEST=l_qty_harvest
                                                        WHERE AIS.ID_ANALYSIS=l_id_analysis_alert AND AIS.ID_SAMPLE_TYPE=l_id_sample_type_alert AND AIS.ID_INSTITUTION=" & i_institution & "
                                                        AND AIS.ID_SOFTWARE= " & i_software & ";   

                                                    END;"

                    Try

                        Dim cmd_insert_ais As New OracleCommand(sql_insert_ais, i_conn)
                        cmd_insert_ais.CommandType = CommandType.Text

                        cmd_insert_ais.ExecuteNonQuery()

                        cmd_insert_ais.Dispose()

                    Catch ex As Exception

                        MsgBox("ERROR INSERTING ANALYSIS INST SOFT")
                        Return False

                    End Try

                End If

            Next

        Catch ex As Exception

            MsgBox("ERROR INSERTING SAMPLE RECIPIENT")
            Return False

        End Try

        Return True

    End Function

    Function SET_ANALYSIS_INST_RECIPIENT(ByVal i_institution As Int64, ByVal i_software As Integer, ByVal i_selected_default_analysis() As analysis_default, ByVal i_conn As OracleConnection) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution, i_conn)

        For i As Integer = 0 To i_selected_default_analysis.Count() - 1

            'Função que vai colocar o tubo na análise
            'Nota: Esta tabela permite ter muitos registos para um único id_analysis_inst_soft
            Dim sql_insert_air As String = "DECLARE

                                                    l_id_analysis_alert    alert.analysis_sample_type.id_analysis%TYPE;
                                                    l_id_sample_type_alert alert.analysis_sample_type.id_sample_type%TYPE;

                                                    l_id_sample_recipient_alert alert.sample_recipient.id_sample_recipient%TYPE;

                                                    l_id_analysis_inst_soft_alert alert.analysis_instit_soft.id_analysis_instit_soft%TYPE;

                                                BEGIN

                                                    SELECT ast.id_analysis, ast.id_sample_type
                                                    INTO l_id_analysis_alert, l_id_sample_type_alert
                                                    FROM alert.analysis_sample_type ast
                                                    WHERE ast.id_content = '" & i_selected_default_analysis(i).id_content_analysis_sample_type & "'
                                                    AND ast.flg_available = 'Y';

                                                    SELECT sr.id_sample_recipient
                                                    INTO l_id_sample_recipient_alert
                                                    FROM alert.sample_recipient sr
                                                    WHERE sr.id_content = '" & i_selected_default_analysis(i).id_content_sample_recipient & "'
                                                    AND sr.flg_available = 'Y';

                                                    SELECT ais.id_analysis_instit_soft
                                                    INTO l_id_analysis_inst_soft_alert
                                                    FROM alert.analysis_instit_soft ais
                                                    WHERE ais.id_analysis = l_id_analysis_alert
                                                    AND ais.id_sample_type = l_id_sample_type_alert
                                                    AND ais.id_institution = " & i_institution & " 
                                                    AND ais.id_software = " & i_software & " 
                                                    AND ais.flg_available = 'Y';

                                                    INSERT INTO alert.analysis_instit_recipient
                                                        (id_analysis_instit_recipient, id_analysis_instit_soft, id_sample_recipient, flg_default)
                                                    VALUES
                                                        (alert.seq_analysis_instit_recipient.nextval, l_id_analysis_inst_soft_alert, l_id_sample_recipient_alert, 'Y');

                                                EXCEPTION
                                                    WHEN dup_val_on_index THEN
                                                        UPDATE alert.analysis_instit_recipient air
                                                        SET air.id_sample_recipient = l_id_sample_recipient_alert
                                                        WHERE air.id_analysis_instit_soft = l_id_analysis_inst_soft_alert;
    
                                                END;"

            Try

                Dim cmd_insert_air As New OracleCommand(sql_insert_air, i_conn)
                cmd_insert_air.CommandType = CommandType.Text

                cmd_insert_air.ExecuteNonQuery()

                cmd_insert_air.Dispose()

            Catch ex As Exception

                MsgBox("ERROR INSERTING ANALYSIS INST RECIPIENT")
                MsgBox(sql_insert_air)
                Return False

            End Try

        Next

        Return True

    End Function

    ''''''''''''''''''''''''''''''''''''
    Function SET_ANALYSIS_ROOM(ByVal i_institution As Int64, ByVal i_software As Integer, ByVal i_id_room As Int64, ByVal i_selected_default_analysis() As analysis_default, ByVal i_conn As OracleConnection) As Boolean

        For i As Integer = 0 To i_selected_default_analysis.Count() - 1

            'Função que vai colocar o tubo na análise
            'Nota: Esta tabela permite ter muitos registos para um único id_analysis_inst_soft
            Dim sql_insert_analysis_room As String = "DECLARE

                                              CURSOR c_new_analysis IS
                                                SELECT Ast.ID_ANALYSIS, ast.id_sample_type
                                                  FROM alert.analysis_sample_type ast
                                                 INNER JOIN alert.analysis_instit_soft ais
                                                    on ais.id_analysis = ast.id_analysis
                                                   and ais.id_sample_type = ast.id_sample_type
                                                   and ais.id_software = " & i_software & "
                                                   and ais.flg_available = 'Y'
                                                 WHERE Ast.ID_CONTENT IN ('" & i_selected_default_analysis(i).id_content_analysis_sample_type & "') 
                                                   AND Ast.FLG_AVAILABLE = 'Y';

                                              l_id_analysis    alert.analysis.id_analysis%type;
                                              l_id_sample_type alert.analysis.id_sample_type%type;

                                              l_id_analysis_room alert.analysis_room.id_analysis_room%type;

                                              FUNCTION record_exists(i_id_analysis    IN alert.analysis.id_analysis%type,
                                                                     i_id_sample_type IN alert.sample_type.id_sample_type%type,
                                                                     i_id_inst        IN alert.analysis_instit_soft.id_institution%type,
                                                                     i_flg_type       IN alert.analysis_instit_soft.flg_type%type)
                                                RETURN BOOLEAN IS
  
                                                l_exists    boolean := FALSE;
                                                l_id_a_room alert.analysis_room.id_analysis_room%type := 0;
  
                                              BEGIN
  
                                                BEGIN
    
                                                  Select ar.id_analysis_room
                                                    into l_id_a_room
                                                    from alert.analysis_room ar
                                                   where ar.id_analysis = i_id_analysis
                                                     and ar.id_sample_type = i_id_sample_type
                                                     and ar.id_institution = i_id_inst
                                                     and ar.flg_type = i_flg_type;
    
                                                  IF l_id_a_room <> 0 THEN
      
                                                    l_exists := TRUE;
      
                                                  ELSE
      
                                                    l_exists := FALSE;
      
                                                  END IF;
    
                                                EXCEPTION
                                                  WHEN no_data_found THEN
                                                    l_exists := FALSE;
      
                                                END;
  
                                                RETURN l_exists;
  
                                              END record_exists;

                                            BEGIN

                                              OPEN c_new_analysis;

                                              LOOP
  
                                                FETCH c_new_analysis
                                                  INTO l_id_analysis, l_id_sample_type;
                                                EXIT WHEN c_new_analysis%NOTFOUND;
  
                                                BEGIN
    
                                                  IF not record_exists(l_id_analysis, l_id_sample_type, " & i_institution & ", 'M') THEN
      
                                                    insert into alert.analysis_room
                                                      (ID_ANALYSIS_ROOM,
                                                       ID_ANALYSIS,
                                                       ID_ROOM,
                                                       RANK,
                                                       ADW_LAST_UPDATE,
                                                       FLG_TYPE,
                                                       FLG_AVAILABLE,
                                                       DESC_EXEC_DESTINATION,
                                                       FLG_DEFAULT,
                                                       ID_INSTITUTION,
                                                       CREATE_USER,
                                                       CREATE_TIME,
                                                       CREATE_INSTITUTION,
                                                       UPDATE_USER,
                                                       UPDATE_TIME,
                                                       UPDATE_INSTITUTION,
                                                       ID_SAMPLE_TYPE)
                                                    values
                                                      (alert.seq_analysis_room.nextval,
                                                       l_id_analysis,
                                                       " & i_id_room & ",
                                                       0,
                                                       null,
                                                       'M',
                                                       'Y',
                                                       null,
                                                       'Y',
                                                        " & i_institution & ",
                                                       null,
                                                       null,
                                                       null,
                                                       null,
                                                       null,
                                                       null,
                                                       l_id_sample_type);            
      
                                            END IF;
    
                                                  IF not record_exists(l_id_analysis, l_id_sample_type,  " & i_institution & ", 'T') THEN
      
                                                    insert into alert.analysis_room
                                                      (ID_ANALYSIS_ROOM,
                                                       ID_ANALYSIS,
                                                       ID_ROOM,
                                                       RANK,
                                                       ADW_LAST_UPDATE,
                                                       FLG_TYPE,
                                                       FLG_AVAILABLE,
                                                       DESC_EXEC_DESTINATION,
                                                       FLG_DEFAULT,
                                                       ID_INSTITUTION,
                                                       CREATE_USER,
                                                       CREATE_TIME,
                                                       CREATE_INSTITUTION,
                                                       UPDATE_USER,
                                                       UPDATE_TIME,
                                                       UPDATE_INSTITUTION,
                                                       ID_SAMPLE_TYPE)
                                                    values
                                                      (alert.seq_analysis_room.nextval,
                                                       l_id_analysis,
                                                       " & i_id_room & ", 
                                                       0,
                                                       null,
                                                       'T',
                                                       'Y',
                                                       null,
                                                       'Y',
                                                        " & i_institution & ",
                                                       null,
                                                       null,
                                                       null,
                                                       null,
                                                       null,
                                                       null,
                                                       l_id_sample_type);

                                                  END IF;
    
                                                END;
  
                                              END LOOP;

                                              CLOSE c_new_analysis;

                                            END;"

            Try

                Dim cmd_insert_analysis_room As New OracleCommand(sql_insert_analysis_room, i_conn)
                cmd_insert_analysis_room.CommandType = CommandType.Text

                cmd_insert_analysis_room.ExecuteNonQuery()

                cmd_insert_analysis_room.Dispose()

            Catch ex As Exception

                MsgBox("ERROR INSERTING ANALYSIS ROOM")
                Return False

            End Try

        Next

        Return True

    End Function

    Function DELETE_ANALYSIS_INST_SOFT(ByVal i_institution As Int64, ByVal i_software As Integer, ByVal i_id_content_ast As String, ByVal i_conn As OracleConnection) As Boolean


        Dim sql_delete_ais = "update alert.analysis_instit_soft ais
                              set ais.flg_available='N'
                              where ais.id_analysis = (Select asta.id_analysis from alert.analysis_sample_type asta where asta.id_content='" & i_id_content_ast & "')
                              and ais.id_sample_type = (Select astst.id_sample_type from alert.analysis_sample_type astst where astst.id_content='" & i_id_content_ast & "')
                              and ais.id_software = " & i_software & "
                              and ais.id_institution = " & i_institution

        Try

            Dim cmd_delete_ais As New OracleCommand(sql_delete_ais, i_conn)
            cmd_delete_ais.CommandType = CommandType.Text

            cmd_delete_ais.ExecuteNonQuery()

            cmd_delete_ais.Dispose()

        Catch ex As Exception

            Return False

        End Try

        Return True

    End Function

    Function SET_ANALYSIS_DEP_CLIN_SERV(ByVal i_software As Integer, ByVal i_dep_clin_serv As Int64, ByVal i_id_content_ast As String, ByVal i_conn As OracleConnection) As Boolean

        Dim sql_insert_adps = "DECLARE

                                    l_id_analysis    alert.analysis_sample_type.id_analysis%TYPE;
                                    l_id_sample_type alert.analysis_sample_type.id_sample_type%TYPE;

                                BEGIN

                                    SELECT ast.id_analysis, ast.id_sample_type
                                    INTO l_id_analysis, l_id_sample_type
                                    FROM alert.analysis_sample_type ast
                                    WHERE ast.id_content = '" & i_id_content_ast & "' and ast.flg_available='Y';

                                    INSERT INTO alert.analysis_dep_clin_serv
                                        (id_analysis_dep_clin_serv, id_analysis, id_dep_clin_serv, rank, id_software, flg_available, id_sample_type)
                                    VALUES
                                        (alert.seq_analysis_dep_clin_serv.nextval, l_id_analysis, " & i_dep_clin_serv & ", 0, " & i_software & ", 'Y', l_id_sample_type);

                                EXCEPTION
                                    WHEN dup_val_on_index THEN
                                        UPDATE alert.analysis_dep_clin_serv ad
                                        SET ad.flg_available = 'Y'
                                        WHERE ad.id_analysis = l_id_analysis
                                        AND ad.id_sample_type = l_id_sample_type
                                        AND ad.id_software = " & i_software & "
                                        And ad.id_dep_clin_serv = " & i_dep_clin_serv & ";
    
                                END;"

        Try

            Dim cmd_insert_adps As New OracleCommand(sql_insert_adps, i_conn)
            cmd_insert_adps.CommandType = CommandType.Text

            cmd_insert_adps.ExecuteNonQuery()

            cmd_insert_adps.Dispose()

        Catch ex As Exception

            Return False

        End Try

        Return True

    End Function

    Function GET_ANALYSIS_DEP_CLIN_SERV(ByVal i_institution As Int64, ByVal i_software As Int16, ByVal i_dep_clin_serv As Int64, ByRef i_dr As OracleDataReader, ByVal i_conn As OracleConnection) As Boolean
        Try

            Dim sql As String = "SELECT ast.id_content,
                                    pk_translation.get_translation(" & db_access_general.GET_ID_LANG(i_institution, i_conn) & ", ast.code_analysis_sample_type),
                                    pk_translation.get_translation(" & db_access_general.GET_ID_LANG(i_institution, i_conn) & ", sr.code_sample_recipient)     
                                FROM alert.analysis_dep_clin_serv ad
                                JOIN alert.analysis_sample_type ast ON ast.id_analysis = ad.id_analysis
                                                                AND ast.id_sample_type = ad.id_sample_type
                                                                AND ast.flg_available = 'Y'
                                JOIN translation t ON t.code_translation = ast.code_analysis_sample_type
                                JOIN alert.analysis_instit_soft ais ON ais.id_analysis = ast.id_analysis
                                                                AND ais.id_sample_type = ast.id_sample_type
                                                                AND ais.id_software = ad.id_software
                                                                AND ais.flg_available = 'Y'
                                                                AND ais.id_institution = " & i_institution & " 
                                JOIN alert.analysis_instit_recipient air ON air.id_analysis_instit_soft = ais.id_analysis_instit_soft
                                JOIN alert.sample_recipient sr ON sr.id_sample_recipient = air.id_sample_recipient
                                                           AND sr.flg_available = 'Y'
                                JOIN translation tsr ON tsr.code_translation = sr.code_sample_recipient
                                WHERE ad.id_software = " & i_software & "
                                AND ad.id_dep_clin_serv = " & i_dep_clin_serv & "
                                AND ad.flg_available = 'Y'"


            Dim cmd As New OracleCommand(sql, i_conn)
            cmd.CommandType = CommandType.Text

            i_dr = cmd.ExecuteReader()

            Return True

        Catch ex As Exception

            Return False

        End Try

    End Function

    Function DELETE_ANALYSIS_DEP_CLIN_SERV(ByVal i_software As Integer, ByVal i_dep_clin_serv As Int64, ByVal i_id_content_ast As String, ByVal i_conn As OracleConnection) As Boolean

        Dim sql_delete_adps = "DECLARE

                                    l_id_analysis    alert.analysis_sample_type.id_analysis%TYPE;
                                    l_id_sample_type alert.analysis_sample_type.id_sample_type%TYPE;

                               BEGIN

                                    SELECT ast.id_analysis, ast.id_sample_type
                                    INTO l_id_analysis, l_id_sample_type
                                    FROM alert.analysis_sample_type ast
                                    WHERE ast.id_content = '" & i_id_content_ast & "' and ast.flg_available='Y';

                                   UPDATE alert.analysis_dep_clin_serv ad
                                   SET ad.flg_available = 'N'
                                   WHERE ad.id_analysis = l_id_analysis
                                   And ad.id_sample_type = l_id_sample_type
                                   And ad.id_software = " & i_software

        If i_dep_clin_serv = 0 Then

            sql_delete_adps = sql_delete_adps & "; END;"

        Else

            sql_delete_adps = sql_delete_adps & "AND ad.id_dep_clin_serv = " & i_dep_clin_serv & ";
                                 
                               END;"

        End If


        Try

            Dim cmd_delete_adps As New OracleCommand(sql_delete_adps, i_conn)
            cmd_delete_adps.CommandType = CommandType.Text

            cmd_delete_adps.ExecuteNonQuery()

            cmd_delete_adps.Dispose()

        Catch ex As Exception

            Return False

        End Try

        Return True

    End Function

End Class
