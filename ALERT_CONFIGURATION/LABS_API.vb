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

    Function GET_DEFAULT_VERSIONS(ByVal i_institution As Int64, ByVal i_software As Integer, ByVal i_conn As OracleConnection, ByRef i_dr As OracleDataReader) As Boolean

        Try
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

            i_dr = cmd.ExecuteReader()

            cmd.Dispose()

            Return True

        Catch ex As Exception

            Return False

        End Try

    End Function

    Function GET_LAB_CATS_DEFAULT(ByVal i_version As String, ByVal i_institution As Int64, ByVal i_software As Integer, ByVal i_conn As OracleConnection, ByRef i_dr As OracleDataReader) As Boolean

        Try

            Dim sql As String = "Select distinct dec.id_content, alert_default.pk_translation_default.get_translation_default(" & db_access_general.GET_ID_LANG(i_institution, i_conn) & ",dec.code_exam_cat)
              
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
                                on dec.id_exam_cat = dais.id_exam_cat --Novo
                              join alert_default.translation dtec
                                on dtec.code_translation = dec.code_exam_cat --Novo

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
            cmd.CommandType = CommandType.Text

            i_dr = cmd.ExecuteReader()

            Return True

        Catch ex As Exception

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

        Try


            Dim cmd As New OracleCommand(sql, i_conn)
            cmd.CommandType = CommandType.Text

            i_dr = cmd.ExecuteReader()

            cmd.Dispose()

            Return True

        Catch ex As Exception

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

        dr.Dispose()
        dr.Close()
        cmd.Dispose()
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

        Try

            Dim cmd As New OracleCommand(sql, i_conn)
            cmd.CommandType = CommandType.Text

            i_dr = cmd.ExecuteReader()

            cmd.Dispose()

            Return True

        Catch ex As Exception

            Return False

        End Try

    End Function

    Function GET_DEFAULT_ANALYSIS_PARAMETERS(ByVal i_id_content_analysis As String, ByVal i_conn As OracleConnection, ByRef i_dr As OracleDataReader) As Boolean

        Dim sql As String = "SELECT a.cpt_code, a.gender, a.age_min, a.age_max, a.mdm_coding, a.ref_form_code, st.id_content, a.barcode
                                FROM alert_default.analysis a
                                LEFT JOIN alert_default.sample_type st ON st.id_sample_type = a.id_sample_type                               
                                WHERE a.id_content = '" & i_id_content_analysis & "'
                                AND a.flg_available = 'Y'"

        Try

            Dim cmd As New OracleCommand(sql, i_conn)
            cmd.CommandType = CommandType.Text

            i_dr = cmd.ExecuteReader()

            cmd.Dispose()

            Return True

        Catch ex As Exception

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


        Try

            Dim cmd As New OracleCommand(sql, i_conn)
            cmd.CommandType = CommandType.Text
            i_Dr = cmd.ExecuteReader()

            cmd.Dispose()

            Return True

        Catch ex As Exception

            Return False

        End Try

    End Function

    Function GET_ANALYSIS_PARAMETERS_ID_CONTENT_DEFAULT(ByVal i_id_software As Int16, ByVal i_id_content_ast As String, ByVal i_conn As OracleConnection, ByRef i_dr As OracleDataReader) As Boolean

        Try

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
            cmd.CommandType = CommandType.Text

            i_dr = cmd.ExecuteReader()

            cmd.Dispose()

            Return True

        Catch ex As Exception

            Return False

        End Try

    End Function

    Function CHECK_RECORD_EXISTENCE(ByVal i_id_content_record As String, ByVal i_sql As String, ByVal i_conn As OracleConnection) As Boolean

        Try

            Dim l_total_records As Int16 = 0

            Dim sql As String = "Select count(*) from " & i_sql & " r
                                 where r.id_content='" & i_id_content_record & "'
                                 and r.flg_available='Y'"

            Dim cmd As New OracleCommand(sql, i_conn)
            cmd.CommandType = CommandType.Text

            Dim dr As OracleDataReader = cmd.ExecuteReader()

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

            Return False

        End Try

    End Function

    Function CHECK_RECORD_TRANSLATION_EXISTENCE(ByVal i_institution As Int64, ByVal id_content_record As String, ByVal i_sql As String, ByVal i_conn As OracleConnection) As Boolean

        Dim l_translation As String = ""

        Dim sql As String = "Select pk_translation.get_translation(" & db_access_general.GET_ID_LANG(i_institution, i_conn) & "," & i_sql & " r
                             where r.id_content='" & id_content_record & "'
                             And r.flg_available='Y'"

        Try

            Dim cmd As New OracleCommand(sql, i_conn)
            cmd.CommandType = CommandType.Text

            Dim dr As OracleDataReader = cmd.ExecuteReader()

            While dr.Read()

                l_translation = dr.Item(0)

            End While

            dr.Dispose()
            dr.Close()
            cmd.Dispose()

        Catch ex As Exception

            Return False

        End Try

        Return True

    End Function

    Function SET_EXAM_CAT(ByVal i_institution As Int64, ByVal i_selected_default_analysis() As analysis_default, ByVal i_conn As OracleConnection) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution, i_conn)

        Try
            'Ciclo que vai correr as categorias todas enviadas à função
            For i As Integer = 0 To i_selected_default_analysis.Count() - 1

                '' 1 - Verificar se existe Categoria pai
                Dim l_cat_parent As Int64 = 0
                Dim l_id_alert_cat_parent As Int64 = 0
                Dim l_rank As Int64 = 0
                Dim l_flg_interface As Char = ""

                Dim sql As String = "Select ec.parent_id from alert_default.Exam_Cat ec
                                     where ec.id_content='" & i_selected_default_analysis(i).id_content_category & "'"

                Try
                    Dim cmd As New OracleCommand(sql, i_conn)
                    cmd.CommandType = CommandType.Text

                    Dim dr As OracleDataReader = cmd.ExecuteReader()

                    While dr.Read()

                        l_cat_parent = dr.Item(0)

                    End While

                    dr.Dispose()
                    dr.Close()
                    cmd.Dispose()

                Catch ex As Exception

                    l_cat_parent = 0

                End Try

                If l_cat_parent > 0 Then 'Significa que existe Categoria Pai no default

                    '' 1.1 - Verificar se cat pai e tradução existem no alert. Se não existem, inserir.
                    ''1.1.1 - Verificar se existe no ALERT
                    sql = "Select ecp.id_content
                       from alert_default.exam_cat ec
                       join alert_default.exam_cat ecp
                       on ecp.id_exam_cat = ec.parent_id
                       where ec.id_content = '" & i_selected_default_analysis(i).id_content_category & "'"

                    Dim l_id_content_cat_parent As String = ""
                    Dim cmd As New OracleCommand(sql, i_conn)
                    cmd.CommandType = CommandType.Text
                    Dim dr As OracleDataReader = cmd.ExecuteReader()

                    While dr.Read()

                        l_id_content_cat_parent = dr.Item(0)

                    End While

                    dr.Dispose()
                    dr.Close()
                    cmd.Dispose()

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

                        '1.3 - Umvez que existe no alert e existe tradução, verificar se tradução do alert é igual à do default
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
                If Not CHECK_RECORD_EXISTENCE(i_selected_default_analysis(i).id_content_category, "alert.exam_cat", i_conn) Then

                    '2.1 - Não existe, Inserir.
                    '2.1.1 - Determinar RANK da categoria E flg_interface
                    Try

                        l_rank = GET_CAT_RANK(i_selected_default_analysis(i).id_content_category, i_conn)

                    Catch ex As Exception

                        l_rank = 0

                    End Try

                    '2.1.2 - Determinar flg_interface da categoria
                    Try

                        l_flg_interface = GET_CAT_FLG_INTERFACE(i_selected_default_analysis(i).id_content_category, i_conn)

                    Catch ex As Exception

                        l_flg_interface = "N"

                    End Try

                    '2.1.3 - Inserir Categoria
                    Dim sql_insert_cat As String

                    If l_id_alert_cat_parent = 0 Then

                        sql_insert_cat = "begin
                                      insert into alert.exam_cat (ID_EXAM_CAT, CODE_EXAM_CAT, FLG_AVAILABLE, FLG_LAB, ID_CONTENT, FLG_INTERFACE, RANK, PARENT_ID)
                                      values (alert.seq_exam_cat.nextval, 'EXAM_CAT.CODE_EXAM_CAT.' || alert.seq_exam_cat.nextval, 'Y', 'Y', '" & i_selected_default_analysis(i).id_content_category & "', '" & l_flg_interface & "', " & l_rank & ", null);
                                      end;"
                    Else

                        sql_insert_cat = "begin
                                      insert into alert.exam_cat (ID_EXAM_CAT, CODE_EXAM_CAT, FLG_AVAILABLE, FLG_LAB, ID_CONTENT, FLG_INTERFACE, RANK, PARENT_ID)
                                      values (alert.seq_exam_cat.nextval, 'EXAM_CAT.CODE_EXAM_CAT.' || alert.seq_exam_cat.nextval, 'Y', 'Y', '" & i_selected_default_analysis(i).id_content_category & "', '" & l_flg_interface & "', " & l_rank & ", " & l_id_alert_cat_parent & ");
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

                    Dim l_code_ec_default As String = GET_CODE_EXAM_CAT_DEFAULT(i_selected_default_analysis(i).id_content_category, i_conn)
                    Dim l_code_ec_alert As String = GET_CODE_EXAM_CAT_ALERT(i_selected_default_analysis(i).id_content_category, i_conn)

                    '2.1.4 Inserir translation
                    If Not db_access_general.SET_TRANSLATION((l_id_language), (l_code_ec_alert), (db_access_general.GET_DEFAULT_TRANSLATION(l_id_language, l_code_ec_default, i_conn)), (i_conn)) Then

                        MsgBox("ERROR INSERTING CATEGORY TRANSLATION - LABS_API >> SET_TRANSLATION")
                        Return False

                    End If

                    '2.1.5 - Fazer update a todas as análises que utilizavam o id da categoria antiga com o id da nova categoria (alert.analysis_instit_soft)

                    Dim l_id_alert_category As Int64 = GET_ID_CAT_ALERT(i_selected_default_analysis(i).id_content_category, i_conn)

                    Dim sql_update_analysis_cat As String = "update alert.analysis_instit_soft ais 
                                                         set ais.id_exam_cat=" & l_id_alert_category & "
                                                         where ais.id_exam_cat in (select ec.id_exam_cat  from alert.exam_cat ec where ec.id_content='" & i_selected_default_analysis(i).id_content_category & "')"

                    Dim cmd_update_analysis_cat As New OracleCommand(sql_update_analysis_cat, i_conn)
                    cmd_update_analysis_cat.CommandType = CommandType.Text

                    cmd_update_analysis_cat.ExecuteNonQuery()

                    cmd_update_analysis_cat.Dispose()

                    '2.2 - Uma vez que existe no ALERT, verificar se exsite tradução para a lingua da instituição
                ElseIf Not CHECK_RECORD_TRANSLATION_EXISTENCE(i_institution, i_selected_default_analysis(i).id_content_category, "r.code_exam_cat) from alert.exam_cat", i_conn) Then

                    Dim l_code_ec_default As String = GET_CODE_EXAM_CAT_DEFAULT(i_selected_default_analysis(i).id_content_category, i_conn)
                    Dim l_code_ec_alert As String = GET_CODE_EXAM_CAT_ALERT(i_selected_default_analysis(i).id_content_category, i_conn)

                    If Not db_access_general.SET_TRANSLATION((l_id_language), (l_code_ec_alert), (db_access_general.GET_DEFAULT_TRANSLATION(l_id_language, l_code_ec_default, i_conn)), (i_conn)) Then

                        MsgBox("ERROR INSERTING EXAM_CAT TRANSLATION - LABS_API >> CHECK_RECORD_TRANSLATION_EXISTENCE >> SET_TRANSLATION " & l_id_language)
                        Return False

                    End If

                    '2.3 - Umvez que existe no alert e existe tradução, verificar se tradução do alert é igual à do default
                ElseIf Not db_access_general.CHECK_TRANSLATIONS(l_id_language, GET_CODE_EXAM_CAT_DEFAULT(i_selected_default_analysis(i).id_content_category, i_conn), GET_CODE_EXAM_CAT_ALERT(i_selected_default_analysis(i).id_content_category, i_conn), i_conn) Then

                    Dim l_code_ec_default As String = GET_CODE_EXAM_CAT_DEFAULT(i_selected_default_analysis(i).id_content_category, i_conn)
                    Dim l_code_ec_alert As String = GET_CODE_EXAM_CAT_ALERT(i_selected_default_analysis(i).id_content_category, i_conn)

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

    Function SET_PARAMETER(ByVal i_institution As Int64, ByVal i_software As Int16, ByVal i_id_content_analysis_sample_type As String, ByVal i_conn As OracleConnection) As Boolean

        'Array que vai guardar os id_content dos parâmetros do analysis_sample_type
        Dim l_array_parameters() As String
        Dim l_index As Int16 = 0

        '1.1 - Obter os id_contents dos parâmetros da ast
        Try
            Dim dr_parameters As OracleDataReader


#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
            If Not GET_ANALYSIS_PARAMETERS_ID_CONTENT_DEFAULT(i_software, i_id_content_analysis_sample_type, i_conn, dr_parameters) Then
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

        Catch ex As Exception

            Return False

        End Try

        '1.2 - Verificar se parâmetros existem no ALERT
        For i As Integer = 0 To l_array_parameters.Count() - 1

            '1.2.1 - Inserir registo
            'cHECKAR SE EXISTEM NO ALERT
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
                Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution, i_conn)
                Dim l_code_analysis_parameter_default As String = GET_CODE_PARAMETER_DEFAULT(l_array_parameters(i), i_conn)
                Dim l_code_analysis_parameter_alert As String = GET_CODE_PARAMETER_ALERT(l_array_parameters(i), i_conn)

                If Not db_access_general.SET_TRANSLATION((l_id_language), (l_code_analysis_parameter_alert), (db_access_general.GET_DEFAULT_TRANSLATION(l_id_language, l_code_analysis_parameter_default, i_conn)), (i_conn)) Then

                    MsgBox("ERROR INSERTING PARAMETER TRANSLATION - LABS_API >> CHECK_ANALYSIS_ST_TRANSLATION_EXISTENCE >> SET_TRANSLATION")

                    Return False

                End If

            ElseIf Not CHECK_RECORD_TRANSLATION_EXISTENCE(i_institution, l_array_parameters(i), "'r.code_analysis_parameter from alert.analysis_parameter", i_conn) Then

                Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution, i_conn)
                Dim l_code_analysis_parameter_default As String = GET_CODE_PARAMETER_DEFAULT(l_array_parameters(i), i_conn)
                Dim l_code_analysis_parameter_alert As String = GET_CODE_PARAMETER_ALERT(l_array_parameters(i), i_conn)

                If Not db_access_general.SET_TRANSLATION((l_id_language), (l_code_analysis_parameter_alert), (db_access_general.GET_DEFAULT_TRANSLATION(l_id_language, l_code_analysis_parameter_default, i_conn)), (i_conn)) Then

                    MsgBox("ERROR INSERTING PARAMETER TRANSLATION - LABS_API >> CHECK_ANALYSIS_ST_TRANSLATION_EXISTENCE >> SET_TRANSLATION")

                    Return False

                End If

            End If

        Next

        '1.3 - Ver se é necessário fazer update aos registos do ALERT (Registos que no default estão a N e no alert estão a Y)
        If Not UPDATE_PARAMETER_AVAILABILITY(i_software, i_id_content_analysis_sample_type, i_conn) Then

            MsgBox("ERROR UPDATING PARAMETERS!", vbCritical)

        End If

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

End Class
