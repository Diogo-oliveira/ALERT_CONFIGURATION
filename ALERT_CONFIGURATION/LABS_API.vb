Imports Oracle.DataAccess.Client
Public Class LABS_API

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

        Dim sql As String = "Select distinct dec.id_content, decode(dav.id_market,
              1,
              dtec.desc_lang_1,
              2,
              dtec.desc_lang_2,
              3,
              dtec.desc_lang_11,
              4,
              dtec.desc_lang_5,
              5,
              dtec.desc_lang_4,
              6,
              dtec.desc_lang_3,
              7,
              dtec.desc_lang_10,
              8,
              dtec.desc_lang_7,
              9,
              dtec.desc_lang_6,
              10,
              dtec.desc_lang_9,
              12,
              dtec.desc_lang_16,
              16,
              dtec.desc_lang_17,
              17,
              dtec.desc_lang_18,
              19,
              dtec.desc_lang_19)
              
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

        Dim sql As String = "Select distinct dast.id_content, decode(dav.id_market,
              1,
              dtast.desc_lang_1,
              2,
              dtast.desc_lang_2,
              3,
              dtast.desc_lang_11,
              4,
              dtast.desc_lang_5,
              5,
              dtast.desc_lang_4,
              6,
              dtast.desc_lang_3,
              7,
              dtast.desc_lang_10,
              8,
              dtast.desc_lang_7,
              9,
              dtast.desc_lang_6,
              10,
              dtast.desc_lang_9,
              12,
              dtast.desc_lang_16,
              16,
              dtast.desc_lang_17,
              17,
              dtast.desc_lang_18,
              19,
              dtast.desc_lang_19), dsr.id_content,
              
              decode(dav.id_market,
              1,
              dtsr.desc_lang_1,
              2,
              dtsr.desc_lang_2,
              3,
              dtsr.desc_lang_11,
              4,
              dtsr.desc_lang_5,
              5,
              dtsr.desc_lang_4,
              6,
              dtsr.desc_lang_3,
              7,
              dtsr.desc_lang_10,
              8,
              dtsr.desc_lang_7,
              9,
              dtsr.desc_lang_6,
              10,
              dtsr.desc_lang_9,
              12,
              dtsr.desc_lang_16,
              16,
              dtsr.desc_lang_17,
              17,
              dtsr.desc_lang_18,
              19,
              dtsr.desc_lang_19)
              
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
