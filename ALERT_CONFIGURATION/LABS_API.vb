Imports Oracle.DataAccess.Client
Public Class LABS_API

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

    Function GET_DEFAULT_VERSIONS(ByVal i_institution As Int64, ByVal i_software As Integer, ByVal i_oradb As String) As OracleDataReader

        Dim oradb As String = i_oradb

        Dim conn As New OracleConnection(oradb)

        conn.Open()

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
   and alert_default.pk_translation_default.get_translation_default(" & db_access_general.GET_ID_LANG(i_institution, i_oradb) & ",dast.code_analysis_sample_type) is not null
 order by 1 asc"

        Dim cmd As New OracleCommand(sql, conn)
        cmd.CommandType = CommandType.Text

        Dim dr As OracleDataReader = cmd.ExecuteReader()

        Return dr

    End Function

    Function GET_LAB_CATS_DEFAULT(ByVal i_version As String, ByVal i_institution As Int64, ByVal i_software As Integer, ByVal i_oradb As String) As OracleDataReader

        Dim oradb As String = i_oradb

        Dim conn As New OracleConnection(oradb)

        conn.Open()

        Dim sql As String = "Select distinct dec.id_content, alert_default.pk_translation_default.get_translation_default(" & db_access_general.GET_ID_LANG(i_institution, i_oradb) & ",dec.code_exam_cat)
              
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


        Dim cmd As New OracleCommand(sql, conn)
        cmd.CommandType = CommandType.Text

        Dim dr As OracleDataReader = cmd.ExecuteReader()

        Return dr

    End Function

    Function GET_LABS_DEFAULT_BY_CAT(ByVal i_institution As Int64, ByVal i_software As Integer, ByVal i_version As String, ByVal i_id_cat As String, ByVal i_oradb As String) As OracleDataReader

        Dim oradb As String = i_oradb

        Dim conn As New OracleConnection(oradb)

        conn.Open()

        Dim sql As String = "Select distinct dast.id_content, 
                                             alert_default.pk_translation_default.get_translation_default(" & db_access_general.GET_ID_LANG(i_institution, i_oradb) & ", dast.code_analysis_sample_type), 
                                             dsr.id_content,              
                                             alert_default.pk_translation_default.get_translation_default(" & db_access_general.GET_ID_LANG(i_institution, i_oradb) & ", dsr.code_sample_recipient), 
                                             da.id_content, 
                                             dst.id_content
              
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

        Dim cmd As New OracleCommand(sql, conn)
        cmd.CommandType = CommandType.Text

        Dim dr As OracleDataReader = cmd.ExecuteReader()

        Return dr

    End Function

End Class
