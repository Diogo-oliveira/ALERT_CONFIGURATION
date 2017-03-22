﻿Imports Oracle.DataAccess.Client
Public Class INTERVENTIONS_API

    Dim db_access_general As New General

    Public Structure interventions_default
        Public id_content_category As String
        Public id_content_intervention As String
        Public desc_intervention As String
    End Structure

    Public Structure interventions_alert_flg
        Public id_content_category As String
        Public id_content_intervention As String
        Public desc_intervention As String
        Public flg_new As String
    End Structure

    Function GET_DEFAULT_VERSIONS(ByVal i_institution As Int64, ByVal i_software As Integer, ByVal i_flg_type As Integer, ByRef i_dr As OracleDataReader) As Boolean

        Dim sql As String = "SELECT DISTINCT dim.version
                                FROM alert_default.intervention di
                                JOIN alert_default.translation dti ON dti.code_translation = di.code_intervention
                                JOIN alert_default.interv_int_cat diic ON diic.id_intervention = di.id_intervention
                                JOIN alert_default.interv_mrk_vrs dim ON dim.id_intervention = di.id_intervention
                                JOIN alert.interv_category aic ON aic.id_interv_category = diic.id_interv_category
                                JOIN alert_default.interv_clin_serv dcs ON dcs.id_intervention = di.id_intervention
                                JOIN translation t ON t.code_translation = aic.code_interv_category
                                JOIN institution i ON i.id_market = dim.id_market
                                WHERE di.flg_status = 'A'
                                AND diic.id_software IN (0, " & i_software & ")
                                AND i.id_institution = " & i_institution & "
                                AND dcs.id_software IN (0, " & i_software & ")"

        If i_flg_type = 0 Then

            sql = sql & "And dcs.flg_type IN ('P', 'M', 'B','A') "

        ElseIf i_flg_type = 1 Then

            sql = sql & "And dcs.flg_type IN ('P', 'M') "

        Else

            sql = sql & "And dcs.flg_type IN ('B','A') "

        End If

        sql = sql & "  ORDER BY 1 ASC"

        Dim cmd As New OracleCommand(sql, Connection.conn)
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

    Function GET_INTERV_CATS_INST_SOFT(ByVal i_institution As Int64, ByVal i_software As Integer, ByVal i_flg_type As Integer, ByRef i_dr As OracleDataReader) As Boolean

        'Esta função vai ver as categorias que têm procedimentos disponíveis para a Instituição e Softwares selecionados
        'Os procedimentos têm que respeitar a fla_add_remove da tabela alert_int_cat

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution)
        Dim sql As String = "with tbl_interv_cats (id_content_interv_cat,id_content_interv, cod_interv_cat)
                                as
                                (SELECT DISTINCT ic.id_content, i.id_content, ic.code_interv_category
                                FROM alert.interv_int_cat iic
                                JOIN alert.interv_category ic ON ic.id_interv_category = iic.id_interv_category
                                JOIN alert.intervention i ON i.id_intervention = iic.id_intervention
                                JOIN alert.interv_dep_clin_serv idcs ON idcs.id_intervention = i.id_intervention
                                WHERE i.flg_status = 'A'
                                AND iic.id_software IN (0, " & i_software & ")
                                AND iic.id_institution IN (0, " & i_institution & ")
                                AND iic.flg_add_remove = 'A'
                                AND ic.flg_available = 'Y'
                                AND idcs.id_institution IN (0, " & i_institution & ")
                                AND idcs.id_software IN (0, " & i_software & ") 
                                And i.id_interv_physiatry_area is null " 'Linha para garantir que só aparecem as categorias com procedimentos paa a área normal
        If i_flg_type = 0 Then

            sql = sql & "And idcs.flg_type IN ('P', 'M', 'A', 'B') "

        ElseIf i_flg_type = 1 Then

            sql = sql & "And idcs.flg_type IN ('P', 'M') "

        Else

            sql = sql & "And idcs.flg_type IN ('A', 'B') "

        End If

        sql = sql & "
                                MINUS
                                
                                --Remover para Soft e instituição definidos
                                SELECT DISTINCT ic.id_content , i.id_content , ic.code_interv_category
                                FROM alert.interv_int_cat iic
                                JOIN alert.interv_category ic ON ic.id_interv_category = iic.id_interv_category
                                JOIN alert.intervention i ON i.id_intervention = iic.id_intervention
                                JOIN alert.interv_dep_clin_serv idcs ON idcs.id_intervention = i.id_intervention
                                WHERE iic.id_software IN (" & i_software & ")
                                AND i.flg_status = 'A'
                                AND iic.id_institution IN (" & i_institution & ")
                                AND iic.flg_add_remove = 'R'
                                AND ic.flg_available = 'Y'
                                AND idcs.id_institution IN (0, " & i_institution & ")
                                AND idcs.id_software IN (0, " & i_software & ") "
        If i_flg_type = 0 Then

            sql = sql & "And idcs.flg_type IN ('P','M','B', 'A') "

        ElseIf i_flg_type = 1 Then

            sql = sql & "And idcs.flg_type IN ('P','M') "

        Else

            sql = sql & "And idcs.flg_type IN ('B','A') "

        End If

        sql = sql & "
                                MINUS
                                
                                --Remover para Instituição a 0 e soft definido
                                SELECT DISTINCT ic.id_content , i.id_content ,ic.code_interv_category
                                FROM alert.interv_int_cat iic
                                JOIN alert.interv_category ic ON ic.id_interv_category = iic.id_interv_category
                                JOIN alert.intervention i ON i.id_intervention = iic.id_intervention
                                JOIN alert.interv_dep_clin_serv idcs ON idcs.id_intervention = i.id_intervention
                                WHERE iic.id_software = " & i_software & "
                                AND i.flg_status = 'A'
                                AND iic.id_institution IN (0)
                                AND iic.flg_add_remove = 'R'
                                AND ic.flg_available = 'Y'
                                AND idcs.id_institution IN (0, " & i_institution & ")
                                AND idcs.id_software IN (0, " & i_software & ") "
        If i_flg_type = 0 Then

            sql = sql & "And idcs.flg_type IN ('P','M','B','A') "

        ElseIf i_flg_type = 1 Then

            sql = sql & "And idcs.flg_type IN ('P','M') "

        Else

            sql = sql & "And idcs.flg_type IN ('B','A') "

        End If

        sql = sql & "
                                MINUS
                                
                                --REMOVER Para Soft 0 e Inst definida
                                SELECT DISTINCT ic.id_content , i.id_content ,ic.code_interv_category
                                FROM alert.interv_int_cat iic
                                JOIN alert.interv_category ic ON ic.id_interv_category = iic.id_interv_category
                                JOIN alert.intervention i ON i.id_intervention = iic.id_intervention
                                JOIN alert.interv_dep_clin_serv idcs ON idcs.id_intervention = i.id_intervention
                                WHERE iic.id_software = 0
                                AND i.flg_status = 'A'
                                AND iic.id_institution IN (" & i_institution & ")
                                AND iic.flg_add_remove = 'R'
                                AND ic.flg_available = 'Y'
                                AND idcs.id_institution IN (0, " & i_institution & ")
                                AND idcs.id_software IN (0, " & i_software & ") "
        If i_flg_type = 0 Then

            sql = sql & "And idcs.flg_type IN ('P','M','B', 'A') "

        ElseIf i_flg_type = 1 Then

            sql = sql & "And idcs.flg_type IN ('P','M') "

        Else

            sql = sql & "And idcs.flg_type IN ('B','A') "

        End If

        sql = sql & ")       
                          select distinct id_content_interv_cat, pk_translation.get_translation(" & l_id_language & ",cod_interv_cat) from tbl_interv_cats
                          WHERE  pk_translation.get_translation(" & l_id_language & ",cod_interv_cat) IS NOT NULL
                         -- and cod_interv_cat not like 'SPECIALITY%' (Existem associações a especialidades, por isso tive que remover isto)
                          ORDER BY 2 ASC"

        Dim cmd As New OracleCommand(sql, Connection.conn)
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

    Function GET_FREQ_INTERVS(ByVal i_institution As Int64, ByVal i_software As Integer, ByVal i_flg_type As Integer, ByVal i_id_dep_clin_serv As Int64, ByRef i_dr As OracleDataReader) As Boolean

        'Esta função vai ver as categorias que têm procedimentos disponíveis para a Instituição e Softwares selecionados
        'Os procedimentos têm que respeitar a fla_add_remove da tabela alert_int_cat

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution)
        Dim sql As String = "with tbl_interv (id_content_interv, code_intervention)
                                as
                                (SELECT DISTINCT i.id_content, i.code_intervention
                                FROM alert.interv_int_cat iic
                                JOIN alert.interv_category ic ON ic.id_interv_category = iic.id_interv_category
                                JOIN alert.intervention i ON i.id_intervention = iic.id_intervention
                                JOIN alert.interv_dep_clin_serv idcs ON idcs.id_intervention = i.id_intervention
                                JOIN translation t ON t.code_translation = ic.code_interv_category
                                WHERE i.flg_status = 'A'
                                AND iic.id_software IN (0, " & i_software & ")
                                AND iic.id_institution IN (0, " & i_institution & ")
                                AND iic.flg_add_remove = 'A'
                                AND ic.flg_available = 'Y'
                                AND idcs.id_institution IN (0, " & i_institution & ")
                                AND idcs.id_software IN (0, " & i_software & ")
                                and idcs.id_dep_clin_serv= " & i_id_dep_clin_serv & " 
                                And i.id_interv_physiatry_area is null " ''Linha para garantir que não são mostrados os que têm entrada para a physiatry area

        If i_flg_type = 0 Then

            sql = sql & "And idcs.flg_type IN ('M', 'A') "

        ElseIf i_flg_type = 1 Then

            sql = sql & "And idcs.flg_type IN ('M') "

        Else

            sql = sql & "And idcs.flg_type IN ('A') "

        End If

        sql = sql & "
                                MINUS
                                
                                --Remover para Soft e instituição definidos
                                SELECT DISTINCT i.id_content, i.code_intervention
                                FROM alert.interv_int_cat iic
                                JOIN alert.interv_category ic ON ic.id_interv_category = iic.id_interv_category
                                JOIN alert.intervention i ON i.id_intervention = iic.id_intervention
                                JOIN alert.interv_dep_clin_serv idcs ON idcs.id_intervention = i.id_intervention
                                JOIN translation t ON t.code_translation = ic.code_interv_category
                                WHERE iic.id_software IN (" & i_software & ")
                                AND i.flg_status = 'A'
                                AND iic.id_institution IN (" & i_institution & ")
                                AND iic.flg_add_remove = 'R'
                                AND ic.flg_available = 'Y'
                                AND idcs.id_institution IN (0, " & i_institution & ")
                                AND idcs.id_software IN (0, " & i_software & ") "
        If i_flg_type = 0 Then

            sql = sql & "And idcs.flg_type IN ('P','M','B','A') "

        ElseIf i_flg_type = 1 Then

            sql = sql & "And idcs.flg_type IN ('P','M') "

        Else

            sql = sql & "And idcs.flg_type IN ('B','A') "

        End If

        sql = sql & "
                                MINUS
                                
                                --Remover para Instituição a 0 e soft definido
                                SELECT DISTINCT i.id_content, i.code_intervention
                                FROM alert.interv_int_cat iic
                                JOIN alert.interv_category ic ON ic.id_interv_category = iic.id_interv_category
                                JOIN alert.intervention i ON i.id_intervention = iic.id_intervention
                                JOIN alert.interv_dep_clin_serv idcs ON idcs.id_intervention = i.id_intervention
                                JOIN translation t ON t.code_translation = ic.code_interv_category
                                WHERE iic.id_software = " & i_software & "
                                AND i.flg_status = 'A'
                                AND iic.id_institution IN (0)
                                AND iic.flg_add_remove = 'R'
                                AND ic.flg_available = 'Y'
                                AND idcs.id_institution IN (0, " & i_institution & ")
                                AND idcs.id_software IN (0, " & i_software & ") "
        If i_flg_type = 0 Then

            sql = sql & "And idcs.flg_type IN ('P','M','B','A') "

        ElseIf i_flg_type = 1 Then

            sql = sql & "And idcs.flg_type IN ('P','M') "

        Else

            sql = sql & "And idcs.flg_type IN ('B','A') "

        End If

        sql = sql & "
                                MINUS
                                
                                --REMOVER Para Soft 0 e Inst definida
                                SELECT DISTINCT i.id_content, i.code_intervention
                                FROM alert.interv_int_cat iic
                                JOIN alert.interv_category ic ON ic.id_interv_category = iic.id_interv_category
                                JOIN alert.intervention i ON i.id_intervention = iic.id_intervention
                                JOIN alert.interv_dep_clin_serv idcs ON idcs.id_intervention = i.id_intervention
                                JOIN translation t ON t.code_translation = ic.code_interv_category
                                WHERE iic.id_software = 0
                                AND i.flg_status = 'A'
                                AND iic.id_institution IN (" & i_institution & ")
                                AND iic.flg_add_remove = 'R'
                                AND ic.flg_available = 'Y'
                                AND idcs.id_institution IN (0, " & i_institution & ")
                                AND idcs.id_software IN (0, " & i_software & ") "
        If i_flg_type = 0 Then

            sql = sql & "And idcs.flg_type IN ('P','M','B','A') "

        ElseIf i_flg_type = 1 Then

            sql = sql & "And idcs.flg_type IN ('P','M') "

        Else

            sql = sql & "And idcs.flg_type IN ('B','A') "

        End If

        sql = sql & ")       
                          select distinct id_content_interv, t.desc_lang_" & l_id_language & " from tbl_interv
                          join translation t on t.code_translation=tbl_interv.code_intervention
                          ORDER BY 2 ASC"

        Dim cmd As New OracleCommand(sql, Connection.conn)
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

    Function GET_INTERVS_INST_SOFT(ByVal i_institution As Int64, ByVal i_software As Integer, ByVal i_id_content_interv_cat As String, ByVal i_flg_type As Integer, ByRef i_dr As OracleDataReader) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution)
        Dim sql As String = "WITH tbl_interventions(id_content_interv_cat,
                            id_content_intervention,
                            code_intervention) AS
                             (SELECT DISTINCT ic.id_content, i.id_content, i.code_intervention
                              FROM alert.intervention i
                              JOIN alert.interv_int_cat iic ON iic.id_intervention = i.id_intervention
                              JOIN alert.interv_category ic ON ic.id_interv_category = iic.id_interv_category
                              JOIN alert.interv_dep_clin_serv idcs ON idcs.id_intervention = i.id_intervention
                              WHERE i.flg_status = 'A'
                              AND iic.id_software IN (0, " & i_software & ")
                              AND iic.id_institution IN (0, " & i_institution & ")
                              AND iic.flg_add_remove = 'A'
                              AND ic.flg_available = 'Y'
                              AND idcs.id_institution IN (0, " & i_institution & ")
                              and idcs.id_software in (0," & i_software & ")
                              And i.id_interv_physiatry_area is null  " 'Linha para garantir que só são listados os procedimentos que surgem na área normal

        If i_flg_type = 0 Then

            sql = sql & "And idcs.flg_type IN ('P','M','A','B') "

        ElseIf i_flg_type = 1 Then

            sql = sql & "And idcs.flg_type IN ('P','M') "

        Else

            sql = sql & "And idcs.flg_type IN ('A','B') "

        End If

        sql = sql & "MINUS
  
                      SELECT DISTINCT ic.id_content, i.id_content, i.code_intervention
                      FROM alert.intervention i
                      JOIN alert.interv_int_cat iic ON iic.id_intervention = i.id_intervention
                      JOIN alert.interv_category ic ON ic.id_interv_category = iic.id_interv_category
                      JOIN alert.interv_dep_clin_serv idcs ON idcs.id_intervention = i.id_intervention
                      WHERE i.flg_status = 'A'
                      AND iic.id_software IN (" & i_software & ")
                      AND iic.id_institution IN (" & i_institution & ")
                      AND iic.flg_add_remove = 'R'
                      AND ic.flg_available = 'Y'
                      AND idcs.id_institution IN (0, " & i_institution & ")
                      and idcs.id_software in (0," & i_software & ")"
        If i_flg_type = 0 Then

            sql = sql & "And idcs.flg_type IN ('P','M','A','B') "

        ElseIf i_flg_type = 1 Then

            sql = sql & "And idcs.flg_type IN ('P','M') "

        Else

            sql = sql & "And idcs.flg_type IN ('A','B') "

        End If

        sql = sql & "MINUS
  
                      SELECT DISTINCT ic.id_content, i.id_content, i.code_intervention
                      FROM alert.intervention i
                      JOIN alert.interv_int_cat iic ON iic.id_intervention = i.id_intervention
                      JOIN alert.interv_category ic ON ic.id_interv_category = iic.id_interv_category
                      JOIN alert.interv_dep_clin_serv idcs ON idcs.id_intervention = i.id_intervention
                      WHERE i.flg_status = 'A'
                      AND iic.id_software IN (" & i_software & ")
                      AND iic.id_institution IN (0)
                      AND iic.flg_add_remove = 'R'
                      AND ic.flg_available = 'Y'
                      AND idcs.id_institution IN (0, " & i_institution & ")
                      and idcs.id_software in (0," & i_software & ")"

        If i_flg_type = 0 Then

            sql = sql & "And idcs.flg_type IN ('P','M','A','B') "

        ElseIf i_flg_type = 1 Then

            sql = sql & "And idcs.flg_type IN ('P','M') "

        Else

            sql = sql & "And idcs.flg_type IN ('A','B') "

        End If

        sql = sql & "MINUS

                      SELECT DISTINCT ic.id_content, i.id_content, i.code_intervention
                      FROM alert.intervention i
                      JOIN alert.interv_int_cat iic ON iic.id_intervention = i.id_intervention
                      JOIN alert.interv_category ic ON ic.id_interv_category = iic.id_interv_category
                      JOIN alert.interv_dep_clin_serv idcs ON idcs.id_intervention = i.id_intervention
                      WHERE i.flg_status = 'A'
                      AND iic.id_software IN (0)
                      AND iic.id_institution IN (" & i_institution & ")
                      AND iic.flg_add_remove = 'R'
                      AND ic.flg_available = 'Y'
                      AND idcs.id_institution IN (0, " & i_institution & ")
                      and idcs.id_software in (0," & i_software & ")"

        If i_flg_type = 0 Then

            sql = sql & "And idcs.flg_type IN ('P','M','A','B') "

        ElseIf i_flg_type = 1 Then

            sql = sql & "And idcs.flg_type IN ('P','M') "

        Else

            sql = sql & "And idcs.flg_type IN ('A','B') "

        End If

        sql = sql & ")

                    SELECT DISTINCT id_content_interv_cat, id_content_intervention, pk_translation.get_translation(" & l_id_language & ", code_intervention)                    
                    FROM tbl_interventions
                    WHERE pk_translation.get_translation(" & l_id_language & ", code_intervention) IS NOT NULL"

        If i_id_content_interv_cat <> "0" Then

            sql = sql & " AND ID_CONTENT_INTERV_CAT= '" & i_id_content_interv_cat & "'
                           ORDER BY 3 ASC"

        Else

            sql = sql & " ORDER BY 3 ASC"

        End If

        Dim cmd As New OracleCommand(sql, Connection.conn)
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

    Function GET_INTERV_CATS_DEFAULT(ByVal i_version As String, ByVal i_institution As Int64, ByVal i_software As Integer, ByVal i_flg_type As Integer, ByRef i_dr As OracleDataReader) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution)
        Dim sql As String = "SELECT DISTINCT ic.id_content, pk_translation.get_translation(" & l_id_language & ", ic.code_interv_category)
                                FROM alert_default.intervention di
                                JOIN alert_default.interv_int_cat diic ON diic.id_intervention = di.id_intervention
                                JOIN alert_default.interv_mrk_vrs dim ON dim.id_intervention = di.id_intervention
                                JOIN alert.interv_category ic ON ic.id_interv_category = diic.id_interv_category
                                JOIN alert_default.interv_clin_serv dcs ON dcs.id_intervention = di.id_intervention
                                WHERE di.flg_status = 'A'
                                AND diic.id_software IN (0, " & i_software & ")
                                AND ic.flg_available = 'Y'
                                AND pk_translation.get_translation(" & l_id_language & ", ic.code_interv_category) IS NOT NULL
                                AND alert_default.pk_translation_default.get_translation_default(" & l_id_language & ", di.code_intervention) IS NOT NULL
                                AND dim.version = '" & i_version & "'
                                AND dcs.id_software IN (0, " & i_software & ") "

        If i_flg_type = 0 Then

            sql = sql & "And dcs.flg_type IN ('P','M','A','B') "

        ElseIf i_flg_type = 1 Then

            sql = sql & "And dcs.flg_type IN ('P','M') "

        Else

            sql = sql & "And dcs.flg_type IN ('A','B') "

        End If

        sql = sql & "           ORDER BY 2 ASC"

        Dim cmd As New OracleCommand(sql, Connection.conn)
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

    Function GET_INTERVS_DEFAULT_BY_CAT(ByVal i_institution As Int64, ByVal i_software As Integer, ByVal i_version As String, ByVal i_id_cat As String, ByVal i_flg_type As Integer, ByRef i_dr As OracleDataReader) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution)
        Dim sql As String = "SELECT DISTINCT ic.id_content, di.id_content, alert_default.pk_translation_default.get_translation_default(" & l_id_language & ", di.code_intervention)
                                FROM alert_default.intervention di
                                JOIN alert_default.interv_int_cat diic ON diic.id_intervention = di.id_intervention
                                JOIN alert_default.interv_mrk_vrs dim ON dim.id_intervention = di.id_intervention
                                JOIN alert.interv_category ic ON ic.id_interv_category = diic.id_interv_category
                                JOIN ALERT_DEFAULT.INTERV_CLIN_SERV DCS ON DCS.ID_INTERVENTION=DI.ID_INTERVENTION AND DCS.ID_SOFTWARE IN (0," & i_software & ")                                
                                WHERE di.flg_status = 'A'
                                AND diic.id_software IN (0, " & i_software & ")
                                AND ic.flg_available = 'Y'
                                AND alert_default.pk_translation_default.get_translation_default(" & l_id_language & ", di.code_intervention) IS NOT NULL
                                AND dim.version = '" & i_version & "'"
        If i_flg_type = 0 Then

            sql = sql & "And dcs.flg_type IN ('P','M','A','B') "

        ElseIf i_flg_type = 1 Then

            sql = sql & "And dcs.flg_type IN ('P','M') "

        Else

            sql = sql & "And dcs.flg_type IN ('A','B') "

        End If

        If i_id_cat = "0" Then
            sql = sql & " order by 3 asc"
        Else
            sql = sql & " and ic.id_content= '" & i_id_cat & "'
                          order by 3 asc"
        End If
        Dim cmd As New OracleCommand(sql, Connection.conn)
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

    Function EXISTS_INTERV_INT_CAT_SOFT(ByVal i_institution As Int64, ByVal i_software As Integer, ByVal i_intervention As interventions_default) As Boolean

        Dim sql As String = "DECLARE

                                l_id_interv_int_cat alert.interv_int_cat.id_intervention%type;

                            BEGIN

                                SELECT *
                                INTO l_id_interv_int_cat
                                FROM (
                                        SELECT distinct c.id_intervention                                       
                                        FROM alert.interv_int_cat c
                                        JOIN alert.intervention i ON i.id_intervention = c.id_intervention
                                        JOIN ALERT.INTERV_CATEGORY IC ON IC.ID_INTERV_CATEGORY=C.ID_INTERV_CATEGORY
                                        WHERE i.id_content = '" & i_intervention.id_content_intervention & "'
                                        AND IC.ID_CONTENT='" & i_intervention.id_content_category & "'                                
                                        AND c.id_software = 0
                                        AND c.id_institution IN (0, " & i_institution & ")
                                        AND i.flg_status = 'A'
                                        AND ic.flg_available = 'Y'
                                        and c.flg_add_remove='A'

                                        UNION
                            
                                        SELECT distinct c.id_intervention                                       
                                        FROM alert.interv_int_cat c
                                        JOIN alert.intervention i ON i.id_intervention = c.id_intervention
                                        JOIN ALERT.INTERV_CATEGORY IC ON IC.ID_INTERV_CATEGORY=C.ID_INTERV_CATEGORY
                                        WHERE i.id_content = '" & i_intervention.id_content_intervention & "'
                                        AND IC.ID_CONTENT='" & i_intervention.id_content_category & "'                                
                                        AND c.id_software = " & i_software & "
                                        AND c.id_institution IN (0, " & i_institution & ")
                                        AND i.flg_status = 'A'
                                        AND ic.flg_available = 'Y'
                                        and c.flg_add_remove='A'
                                    );

                            END;"

        Dim cmd_get_interv As New OracleCommand(sql, Connection.conn)

        Try
            cmd_get_interv.CommandType = CommandType.Text
            cmd_get_interv.ExecuteNonQuery()
        Catch ex As Exception
            cmd_get_interv.Dispose()
            Return False
        End Try

        cmd_get_interv.Dispose()
        Return True

    End Function

    Function EXIST_IN_OTHER_CAT(ByVal i_institution As Int64, ByVal i_software As Integer, ByVal i_id_content_intervention As String) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution)
        Dim sql As String = "WITH tbl_interventions(id_content_interv_cat,
                            id_content_intervention,
                            code_intervention) AS
                             (SELECT DISTINCT ic.id_content, i.id_content, i.code_intervention
                              FROM alert.intervention i
                              JOIN alert.interv_int_cat iic ON iic.id_intervention = i.id_intervention
                              JOIN alert.interv_category ic ON ic.id_interv_category = iic.id_interv_category
                              WHERE i.flg_status = 'A'
                              AND iic.id_software IN (0, " & i_software & ")
                              AND iic.id_institution IN (0, " & i_institution & ")
                              AND iic.flg_add_remove = 'A'
                              AND ic.flg_available = 'Y'
                              AND i.id_content='" & i_id_content_intervention & "'"

        sql = sql & "MINUS
  
                      SELECT DISTINCT ic.id_content, i.id_content, i.code_intervention
                      FROM alert.intervention i
                      JOIN alert.interv_int_cat iic ON iic.id_intervention = i.id_intervention
                      JOIN alert.interv_category ic ON ic.id_interv_category = iic.id_interv_category
                      WHERE i.flg_status = 'A'
                      AND iic.id_software IN (" & i_software & ")
                      AND iic.id_institution IN (" & i_institution & ")
                      AND iic.flg_add_remove = 'R'
                      AND ic.flg_available = 'Y'
                      AND i.id_content='" & i_id_content_intervention & "'"

        sql = sql & "MINUS
  
                      SELECT DISTINCT ic.id_content, i.id_content, i.code_intervention
                      FROM alert.intervention i
                      JOIN alert.interv_int_cat iic ON iic.id_intervention = i.id_intervention
                      JOIN alert.interv_category ic ON ic.id_interv_category = iic.id_interv_category
                      WHERE i.flg_status = 'A'
                      AND iic.id_software IN (" & i_software & ")
                      AND iic.id_institution IN (0)
                      AND iic.flg_add_remove = 'R'
                      AND ic.flg_available = 'Y'
                      AND i.id_content='" & i_id_content_intervention & "'"

        sql = sql & "MINUS

                      SELECT DISTINCT ic.id_content, i.id_content, i.code_intervention
                      FROM alert.intervention i
                      JOIN alert.interv_int_cat iic ON iic.id_intervention = i.id_intervention
                      JOIN alert.interv_category ic ON ic.id_interv_category = iic.id_interv_category
                      JOIN alert.interv_dep_clin_serv idcs ON idcs.id_intervention = i.id_intervention
                      WHERE i.flg_status = 'A'
                      AND iic.id_software IN (0)
                      AND iic.id_institution IN (" & i_institution & ")
                      AND iic.flg_add_remove = 'R'
                      AND ic.flg_available = 'Y'
                      AND idcs.id_institution IN (0, " & i_institution & ")
                      and idcs.id_software in (0," & i_software & ")"

        sql = sql & ")

                    SELECT count(*)               
                    FROM tbl_interventions
                    WHERE pk_translation.get_translation(" & l_id_language & ", code_intervention) IS NOT NULL"

        Dim cmd As New OracleCommand(sql, Connection.conn)
        Dim dr As OracleDataReader
        Dim l_total_records As Integer = 0

        Try
            cmd.CommandType = CommandType.Text
            dr = cmd.ExecuteReader()

            While dr.Read()
                l_total_records = dr.Item(0)
            End While

            cmd.Dispose()
            dr.Dispose()
            dr.Close()

            If l_total_records > 0 Then
                Return True
            Else
                Return False
            End If

        Catch ex As Exception

            cmd.Dispose()
#Disable Warning BC42104 ' Variable is used before it has been assigned a value
            dr.Dispose()
#Enable Warning BC42104 ' Variable is used before it has been assigned a value
            dr.Close()
            Return False

        End Try

    End Function

    Function SET_INTERVENTIONS(ByVal i_institution As Int64, ByVal i_a_interventions() As interventions_default) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution)
        Dim sql As String = "DECLARE

                                l_a_interventions table_varchar := table_varchar("

        For i As Integer = 0 To i_a_interventions.Count() - 1

            If (i < i_a_interventions.Count() - 1) Then

                sql = sql & "'" & i_a_interventions(i).id_content_intervention & "', "

            Else

                sql = sql & "'" & i_a_interventions(i).id_content_intervention & "');"

            End If

        Next

        sql = sql & "   l_intervention    alert.intervention.id_intervention%TYPE;

                        l_interv_physiatry_area alert.intervention.id_interv_physiatry_area%type;
                        l_gender            alert.intervention.gender%TYPE;
                        l_age_min           alert.intervention.age_min%TYPE;
                        l_age_max           alert.intervention.age_max%TYPE;
                        l_cpt_code          alert.intervention.cpt_code%TYPE;
                        l_ref_form_code     alert.intervention.ref_form_code%TYPE;
                        l_flg_type          alert.intervention.flg_type%TYPE;
                        l_barcode           alert.intervention.barcode%TYPE;
                        l_flg_category_type alert.intervention.flg_category_type%TYPE;
                        l_flg_move_patient  alert.intervention.flg_mov_pat%type;
    
                        l_sequence_interv   alert.intervention.id_intervention%type;
    
                        l_interv_desc       alert_default.translation.desc_lang_1%type;

                    BEGIN

                        FOR i IN 1 .. l_a_interventions.count()
                        LOOP
                            BEGIN
        
                                SELECT i.id_intervention
                                INTO l_intervention
                                FROM alert.intervention i
                                WHERE i.id_content = l_a_interventions(i)
                                AND i.flg_status = 'A';
        
                            EXCEPTION
                                WHEN no_data_found THEN
                
                                     l_sequence_interv := ALERT.SEQ_INTERVENTION.NEXTVAL;
                
                                    SELECT di.gender, di.age_min, di.age_max, di.cpt_code, di.ref_form_code, di.flg_type, di.barcode, di.flg_category_type, di.flg_mov_pat, ALERT_DEFAULT.PK_TRANSLATION_DEFAULT.get_translation_default(" & l_id_language & ",DI.CODE_INTERVENTION)
                                    INTO l_gender, l_age_min, l_age_max, l_cpt_code, l_ref_form_code, l_flg_type, l_barcode, l_flg_category_type,l_flg_move_patient, l_interv_desc
                                    FROM alert_default.intervention di
                                    WHERE di.id_content = l_a_interventions(i)
                                    AND di.flg_status = 'A';
                
                                    insert into ALERT.INTERVENTION (ID_INTERVENTION, CODE_INTERVENTION, FLG_STATUS, FLG_MOV_PAT, ID_INTERV_PHYSIATRY_AREA, FLG_TYPE, GENDER, AGE_MIN, AGE_MAX, CPT_CODE, REF_FORM_CODE, ID_CONTENT, BARCODE, FLG_CATEGORY_TYPE,rank)
                                    values (l_sequence_interv, 'INTERVENTION.CODE_INTERVENTION.' || l_sequence_interv, 'A', l_flg_move_patient, l_interv_physiatry_area, l_flg_type, l_gender, l_age_min, l_age_max, l_cpt_code, l_ref_form_code, l_a_interventions(i), l_barcode,  l_flg_category_type,0);
                                
                                    begin
                                               PK_TRANSLATION.insert_into_translation(" & l_id_language & ",'INTERVENTION.CODE_INTERVENTION.'||l_sequence_interv,l_interv_desc);
                                    end;
                
                                    continue;
                            END;
                        END LOOP;

                    END;"

        Dim cmd_insert_interv As New OracleCommand(sql, Connection.conn)

        Try
            cmd_insert_interv.CommandType = CommandType.Text
            cmd_insert_interv.ExecuteNonQuery()
        Catch ex As Exception
            cmd_insert_interv.Dispose()
            Return False
        End Try

        cmd_insert_interv.Dispose()
        Return True

    End Function

    Function SET_INTERVS_TRANSLATION(ByVal i_institution As Int64, ByVal i_a_interventions() As interventions_default) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution)
        Dim sql As String = "DECLARE

                                l_a_interventions table_varchar := table_varchar("

        For i As Integer = 0 To i_a_interventions.Count() - 1

            If (i < i_a_interventions.Count() - 1) Then

                sql = sql & "'" & i_a_interventions(i).id_content_intervention & "', "

            Else

                sql = sql & "'" & i_a_interventions(i).id_content_intervention & "');"

            End If

        Next

        sql = sql & "   l_interv_desc alert_default.translation.desc_lang_1%TYPE;
                        l_interv_code alert.intervention.code_intervention%TYPE;

                    BEGIN

                        FOR i IN 1 .. l_a_interventions.count()
                        LOOP
                            BEGIN
        
                                SELECT i.code_intervention
                                INTO l_interv_code
                                FROM alert.intervention i
                                WHERE i.id_content = l_a_interventions(i)
                                AND i.flg_status = 'A'
                                AND pk_translation.get_translation(" & l_id_language & ", i.code_intervention) IS NULL;
        
                                IF l_interv_code IS NOT NULL
                                THEN
            
                                    SELECT alert_default.pk_translation_default.get_translation_default(" & l_id_language & ", di.code_intervention)
                                    INTO l_interv_desc
                                    FROM alert_default.intervention di
                                    WHERE di.id_content = l_a_interventions(i)
                                    AND di.flg_status = 'A';
            
                                    SELECT i.code_intervention
                                    INTO l_interv_code
                                    FROM alert.intervention i
                                    WHERE i.id_content = l_a_interventions(i)
                                    AND i.flg_status = 'A';
            
                                    pk_translation.insert_into_translation(" & l_id_language & ", l_interv_code, l_interv_desc);
            
                                END IF;
        
                                l_interv_code := '';
        
                            EXCEPTION
                                WHEN OTHERS THEN
                                    continue;
                            END;
                        END LOOP;

                    END;"

        Dim cmd_insert_interv As New OracleCommand(sql, Connection.conn)

        Try
            cmd_insert_interv.CommandType = CommandType.Text
            cmd_insert_interv.ExecuteNonQuery()
        Catch ex As Exception
            cmd_insert_interv.Dispose()
            Return False
        End Try

        cmd_insert_interv.Dispose()
        Return True

    End Function

    Function SET_INTERV_INT_CAT(ByVal i_institution As Int64, ByVal i_software As Integer, ByVal i_a_interventions() As interventions_default) As Boolean

        Dim sql As String = "DECLARE

                                l_a_interventions table_varchar := table_varchar("

        For i As Integer = 0 To i_a_interventions.Count() - 1

            If (i < i_a_interventions.Count() - 1) Then

                sql = sql & "'" & i_a_interventions(i).id_content_intervention & "', "

            Else

                sql = sql & "'" & i_a_interventions(i).id_content_intervention & "');"

            End If

        Next

        sql = sql & "l_a_category      table_varchar := table_varchar("

        For i As Integer = 0 To i_a_interventions.Count() - 1

            If (i < i_a_interventions.Count() - 1) Then

                sql = sql & "'" & i_a_interventions(i).id_content_category & "', "

            Else

                sql = sql & "'" & i_a_interventions(i).id_content_category & "');"

            End If

        Next

        sql = sql & "   l_id_intervention    alert.intervention.id_intervention%TYPE;
                        l_id_interv_category alert.interv_category.id_interv_category%TYPE;

                        l_id_iic alert.interv_int_cat.id_interv_category%TYPE;

                    BEGIN

                        FOR i IN 1 .. l_a_interventions.count()
                        LOOP
    
                            BEGIN
        
                                SELECT i.id_intervention
                                INTO l_id_intervention
                                FROM alert.intervention i
                                WHERE i.id_content = l_a_interventions(i)
                                AND i.flg_status = 'A';
        
                                SELECT ic.id_interv_category
                                INTO l_id_interv_category
                                FROM alert.interv_category ic
                                WHERE ic.id_content = l_a_category(i)
                                AND ic.flg_available = 'Y'
                                AND IC.CODE_INTERV_CATEGORY LIKE 'INTERV%'; --EXISTEM ESPECIALDIADES NESTA TABELA!!!
        
                                INSERT INTO alert.interv_int_cat
                                    (id_interv_category, id_intervention, rank, id_software, id_institution, flg_add_remove)
                                VALUES
                                    (l_id_interv_category, l_id_intervention, 0, " & i_software & ", " & i_institution & ", 'A');
        
                            EXCEPTION
    
                            WHEN dup_val_on_index THEN
                                UPDATE alert.interv_int_cat iic
                                SET iic.flg_add_remove = 'A'
                                WHERE iic.id_interv_category = l_id_interv_category
                                AND iic.id_intervention = l_id_intervention
                                AND iic.id_software = " & i_software & "
                                AND iic.id_institution IN (0, " & i_institution & ");

                                WHEN OTHERS THEN
                                    continue;
            
                            END;
    
                        END LOOP;

                    END;"

        Dim cmd_insert_interv_int_cat As New OracleCommand(sql, Connection.conn)

        Try
            cmd_insert_interv_int_cat.CommandType = CommandType.Text
            cmd_insert_interv_int_cat.ExecuteNonQuery()
        Catch ex As Exception
            cmd_insert_interv_int_cat.Dispose()
            Return False
        End Try

        cmd_insert_interv_int_cat.Dispose()
        Return True

    End Function

    Function SET_DEFAULT_INTERV_DEP_CLIN_SERV(ByVal i_institution As Int64, ByVal i_software As Integer, ByVal i_a_interventions() As interventions_default, ByVal i_flg_type As Integer) As Boolean

        Dim sql As String = "DECLARE

                                l_a_interventions table_varchar := table_varchar("

        For i As Integer = 0 To i_a_interventions.Count() - 1

            If (i < i_a_interventions.Count() - 1) Then

                sql = sql & "'" & i_a_interventions(i).id_content_intervention & "', "

            Else

                sql = sql & "'" & i_a_interventions(i).id_content_intervention & "');"

            End If

        Next

        sql = sql & "  l_id_intervention alert.intervention.id_intervention%TYPE;

                        l_id_iic alert.interv_int_cat.id_interv_category%TYPE;

                        l_a_flg_type table_varchar := table_varchar();

                        l_flg_chargeable alert.interv_dep_clin_serv.flg_chargeable%TYPE;
                        l_flg_bandaid    alert.interv_dep_clin_serv.flg_bandaid%TYPE;

                        l_a_flg_chargeable table_varchar := table_varchar();
                        l_a_flg_bandaid    table_varchar := table_varchar();
                        l_a_dep_clin_serv  table_number := table_number();

                        l_flg_type_indicator  integer := " & i_flg_type & " ;

                        BEGIN
  
                            IF l_flg_type_indicator = 0 THEN

                                FOR i IN 1 .. l_a_interventions.count()
                                LOOP
                                    BEGIN
            
                                        SELECT i.id_intervention
                                        INTO l_id_intervention
                                        FROM alert.intervention i
                                        WHERE i.id_content = l_a_interventions(i)
                                        AND i.flg_status = 'A';
            
                                        SELECT DISTINCT dcs.flg_type BULK COLLECT
                                        INTO l_a_flg_type
                                        FROM alert_default.interv_clin_serv dcs
                                        JOIN alert_default.intervention di ON di.id_intervention = dcs.id_intervention
                                        WHERE di.id_content = l_a_interventions(i);
            
                                        FOR j IN 1 .. l_a_flg_type.count()
                                        LOOP
                
                                            IF (l_a_flg_type(j) <> 'A' AND l_a_flg_type(j) <> 'M')
                                            THEN
                    
                                                BEGIN
                                                    SELECT dcs.flg_chargeable, dcs.flg_bandaid
                                                    INTO l_flg_chargeable, l_flg_bandaid
                                                    FROM alert_default.interv_clin_serv dcs
                                                    JOIN alert_default.intervention di ON di.id_intervention = dcs.id_intervention
                                                    WHERE di.id_content = l_a_interventions(i)
                                                    AND di.flg_status = 'A'
                                                    AND dcs.flg_type = l_a_flg_type(j)
                                                    AND dcs.id_software IN (" & i_software & ");
                        
                                                    INSERT INTO alert.interv_dep_clin_serv
                                                        (id_interv_dep_clin_serv,
                                                         id_intervention,
                                                         id_dep_clin_serv,
                                                         flg_type,
                                                         rank,
                                                         id_institution,
                                                         id_software,
                                                         flg_bandaid,
                                                         flg_chargeable,
                                                         flg_execute,
                                                         flg_timeout)
                                                    VALUES
                                                        (alert.seq_interv_dep_clin_serv.nextval,
                                                         l_id_intervention,
                                                         NULL,
                                                         l_a_flg_type(j),
                                                         0,
                                                         " & i_institution & ",
                                                         " & i_software & ",
                                                         l_flg_bandaid,
                                                         l_flg_chargeable,
                                                         'Y',
                                                         'N');
                                                EXCEPTION
                                                    WHEN OTHERS THEN
                                                        continue;
                                                END;
                    
                                            ELSE
                    
                                                SELECT dcs.flg_chargeable, dcs.flg_bandaid, dps.id_dep_clin_serv BULK COLLECT
                                                INTO l_a_flg_chargeable, l_a_flg_bandaid, l_a_dep_clin_serv
                                                FROM alert_default.interv_clin_serv dcs
                                                JOIN alert_default.intervention di ON di.id_intervention = dcs.id_intervention
                                                JOIN alert_default.clinical_service dc ON dc.id_clinical_service = dcs.id_clinical_service
                                                JOIN alert.clinical_service cs ON cs.id_content = dc.id_content
                                                                           AND cs.flg_available = 'Y'
                                                JOIN alert.dep_clin_serv dps ON dps.id_clinical_service = cs.id_clinical_service
                                                                         AND dps.flg_available = 'Y'
                                                JOIN department d ON d.id_department = dps.id_department
                                                WHERE di.id_content = l_a_interventions(i)
                                                AND di.flg_status = 'A'
                                                AND dcs.flg_type IN (l_a_flg_type(j))
                                                AND dcs.id_software IN (" & i_software & ")
                                                AND d.id_institution = " & i_institution & "
                                                AND d.id_software = " & i_software & ";
                    
                                                FOR k IN 1 .. l_a_dep_clin_serv.count()
                                                LOOP
                        
                                                    BEGIN
                                                        INSERT INTO alert.interv_dep_clin_serv
                                                            (id_interv_dep_clin_serv,
                                                             id_intervention,
                                                             id_dep_clin_serv,
                                                             flg_type,
                                                             rank,
                                                             id_institution,
                                                             id_software,
                                                             flg_bandaid,
                                                             flg_chargeable,
                                                             flg_execute,
                                                             flg_timeout)
                                                        VALUES
                                                            (alert.seq_interv_dep_clin_serv.nextval,
                                                             l_id_intervention,
                                                             l_a_dep_clin_serv(k),
                                                             l_a_flg_type(j),
                                                             0,
                                                             " & i_institution & ",
                                                             " & i_software & ",
                                                             l_a_flg_bandaid(k),
                                                             l_a_flg_chargeable(k),
                                                             'Y',
                                                             'N');
                                                    EXCEPTION
                                                        WHEN OTHERS THEN
                                                            continue;
                                                    END;
                        
                                                END LOOP;
                    
                                            END IF;
                
                                        END LOOP;
            
                                    EXCEPTION
                                        WHEN OTHERS THEN
                                            continue;
                
                                    END;
        
                                END LOOP;

                             ELSIF   l_flg_type_indicator = 1 THEN
       
                                   FOR i IN 1 .. l_a_interventions.count()
                                      LOOP
                                          BEGIN
                  
                                              SELECT i.id_intervention
                                              INTO l_id_intervention
                                              FROM alert.intervention i
                                              WHERE i.id_content = l_a_interventions(i)
                                              AND i.flg_status = 'A';
                  
                                              SELECT DISTINCT dcs.flg_type BULK COLLECT
                                              INTO l_a_flg_type
                                              FROM alert_default.interv_clin_serv dcs
                                              JOIN alert_default.intervention di ON di.id_intervention = dcs.id_intervention
                                              WHERE di.id_content = l_a_interventions(i);
                  
                                              FOR j IN 1 .. l_a_flg_type.count()
                                              LOOP
                      
                                                  IF (l_a_flg_type(j) = 'P')
                                                  THEN
                          
                                                      BEGIN
                                                          SELECT dcs.flg_chargeable, dcs.flg_bandaid
                                                          INTO l_flg_chargeable, l_flg_bandaid
                                                          FROM alert_default.interv_clin_serv dcs
                                                          JOIN alert_default.intervention di ON di.id_intervention = dcs.id_intervention
                                                          WHERE di.id_content = l_a_interventions(i)
                                                          AND di.flg_status = 'A'
                                                          AND dcs.flg_type = l_a_flg_type(j)
                                                          AND dcs.id_software IN (" & i_software & ");
                              
                                                          INSERT INTO alert.interv_dep_clin_serv
                                                              (id_interv_dep_clin_serv,
                                                               id_intervention,
                                                               id_dep_clin_serv,
                                                               flg_type,
                                                               rank,
                                                               id_institution,
                                                               id_software,
                                                               flg_bandaid,
                                                               flg_chargeable,
                                                               flg_execute,
                                                               flg_timeout)
                                                          VALUES
                                                              (alert.seq_interv_dep_clin_serv.nextval,
                                                               l_id_intervention,
                                                               NULL,
                                                               l_a_flg_type(j),
                                                               0,
                                                               " & i_institution & ",
                                                               " & i_software & ",
                                                               l_flg_bandaid,
                                                               l_flg_chargeable,
                                                               'Y',
                                                               'N');
                                                      EXCEPTION
                                                          WHEN OTHERS THEN
                                                              continue;
                                                      END;
                          
                                                  ELSIF (l_a_flg_type(j) = 'M') THEN
                          
                                                      SELECT dcs.flg_chargeable, dcs.flg_bandaid, dps.id_dep_clin_serv BULK COLLECT
                                                      INTO l_a_flg_chargeable, l_a_flg_bandaid, l_a_dep_clin_serv
                                                      FROM alert_default.interv_clin_serv dcs
                                                      JOIN alert_default.intervention di ON di.id_intervention = dcs.id_intervention
                                                      JOIN alert_default.clinical_service dc ON dc.id_clinical_service = dcs.id_clinical_service
                                                      JOIN alert.clinical_service cs ON cs.id_content = dc.id_content
                                                                                 AND cs.flg_available = 'Y'
                                                      JOIN alert.dep_clin_serv dps ON dps.id_clinical_service = cs.id_clinical_service
                                                                               AND dps.flg_available = 'Y'
                                                      JOIN department d ON d.id_department = dps.id_department
                                                      WHERE di.id_content = l_a_interventions(i)
                                                      AND di.flg_status = 'A'
                                                      AND dcs.flg_type IN (l_a_flg_type(j))
                                                      AND dcs.id_software IN (" & i_software & ")
                                                      AND d.id_institution = " & i_institution & "
                                                      AND d.id_software = " & i_software & ";
                          
                                                      FOR k IN 1 .. l_a_dep_clin_serv.count()
                                                      LOOP
                              
                                                          BEGIN
                                                              INSERT INTO alert.interv_dep_clin_serv
                                                                  (id_interv_dep_clin_serv,
                                                                   id_intervention,
                                                                   id_dep_clin_serv,
                                                                   flg_type,
                                                                   rank,
                                                                   id_institution,
                                                                   id_software,
                                                                   flg_bandaid,
                                                                   flg_chargeable,
                                                                   flg_execute,
                                                                   flg_timeout)
                                                              VALUES
                                                                  (alert.seq_interv_dep_clin_serv.nextval,
                                                                   l_id_intervention,
                                                                   l_a_dep_clin_serv(k),
                                                                   l_a_flg_type(j),
                                                                   0,
                                                                   " & i_institution & ",
                                                                   " & i_software & ",
                                                                   l_a_flg_bandaid(k),
                                                                   l_a_flg_chargeable(k),
                                                                   'Y',
                                                                   'N');
                                                          EXCEPTION
                                                              WHEN OTHERS THEN
                                                                  continue;
                                                          END;
                              
                                                      END LOOP;
                          
                                                  END IF;
                      
                                              END LOOP;
                  
                                          EXCEPTION
                                              WHEN OTHERS THEN
                                                  continue;
                      
                                          END;
              
                                      END LOOP;
              
                            ELSE
      
                                FOR i IN 1 .. l_a_interventions.count()
                                          LOOP
                                              BEGIN
                      
                                                  SELECT i.id_intervention
                                                  INTO l_id_intervention
                                                  FROM alert.intervention i
                                                  WHERE i.id_content = l_a_interventions(i)
                                                  AND i.flg_status = 'A';
                      
                                                  SELECT DISTINCT dcs.flg_type BULK COLLECT
                                                  INTO l_a_flg_type
                                                  FROM alert_default.interv_clin_serv dcs
                                                  JOIN alert_default.intervention di ON di.id_intervention = dcs.id_intervention
                                                  WHERE di.id_content = l_a_interventions(i);
                      
                                                  FOR j IN 1 .. l_a_flg_type.count()
                                                  LOOP
                          
                                                      IF (l_a_flg_type(j) = 'B')
                                                      THEN
                              
                                                          BEGIN
                                                              SELECT dcs.flg_chargeable, dcs.flg_bandaid
                                                              INTO l_flg_chargeable, l_flg_bandaid
                                                              FROM alert_default.interv_clin_serv dcs
                                                              JOIN alert_default.intervention di ON di.id_intervention = dcs.id_intervention
                                                              WHERE di.id_content = l_a_interventions(i)
                                                              AND di.flg_status = 'A'
                                                              AND dcs.flg_type = l_a_flg_type(j)
                                                              AND dcs.id_software IN (" & i_software & ");
                                  
                                                              INSERT INTO alert.interv_dep_clin_serv
                                                                  (id_interv_dep_clin_serv,
                                                                   id_intervention,
                                                                   id_dep_clin_serv,
                                                                   flg_type,
                                                                   rank,
                                                                   id_institution,
                                                                   id_software,
                                                                   flg_bandaid,
                                                                   flg_chargeable,
                                                                   flg_execute,
                                                                   flg_timeout)
                                                              VALUES
                                                                  (alert.seq_interv_dep_clin_serv.nextval,
                                                                   l_id_intervention,
                                                                   NULL,
                                                                   l_a_flg_type(j),
                                                                   0,
                                                                   " & i_institution & ",
                                                                   " & i_software & ",
                                                                   l_flg_bandaid,
                                                                   l_flg_chargeable,
                                                                   'Y',
                                                                   'N');
                                                          EXCEPTION
                                                              WHEN OTHERS THEN
                                                                  continue;
                                                          END;
                              
                                                      ELSIF (l_a_flg_type(j) = 'A') THEN
                              
                                                          SELECT dcs.flg_chargeable, dcs.flg_bandaid, dps.id_dep_clin_serv BULK COLLECT
                                                          INTO l_a_flg_chargeable, l_a_flg_bandaid, l_a_dep_clin_serv
                                                          FROM alert_default.interv_clin_serv dcs
                                                          JOIN alert_default.intervention di ON di.id_intervention = dcs.id_intervention
                                                          JOIN alert_default.clinical_service dc ON dc.id_clinical_service = dcs.id_clinical_service
                                                          JOIN alert.clinical_service cs ON cs.id_content = dc.id_content
                                                                                     AND cs.flg_available = 'Y'
                                                          JOIN alert.dep_clin_serv dps ON dps.id_clinical_service = cs.id_clinical_service
                                                                                   AND dps.flg_available = 'Y'
                                                          JOIN department d ON d.id_department = dps.id_department
                                                          WHERE di.id_content = l_a_interventions(i)
                                                          AND di.flg_status = 'A'
                                                          AND dcs.flg_type IN (l_a_flg_type(j))
                                                          AND dcs.id_software IN (" & i_software & ")
                                                          AND d.id_institution = " & i_institution & "
                                                          AND d.id_software = " & i_software & ";
                              
                                                          FOR k IN 1 .. l_a_dep_clin_serv.count()
                                                          LOOP
                                  
                                                              BEGIN
                                                                  INSERT INTO alert.interv_dep_clin_serv
                                                                      (id_interv_dep_clin_serv,
                                                                       id_intervention,
                                                                       id_dep_clin_serv,
                                                                       flg_type,
                                                                       rank,
                                                                       id_institution,
                                                                       id_software,
                                                                       flg_bandaid,
                                                                       flg_chargeable,
                                                                       flg_execute,
                                                                       flg_timeout)
                                                                  VALUES
                                                                      (alert.seq_interv_dep_clin_serv.nextval,
                                                                       l_id_intervention,
                                                                       l_a_dep_clin_serv(k),
                                                                       l_a_flg_type(j),
                                                                       0,
                                                                       " & i_institution & ",
                                                                       " & i_software & ",
                                                                       l_a_flg_bandaid(k),
                                                                       l_a_flg_chargeable(k),
                                                                       'Y',
                                                                       'N');
                                                              EXCEPTION
                                                                  WHEN OTHERS THEN
                                                                      continue;
                                                              END;
                                  
                                                          END LOOP;
                              
                                                      END IF;
                          
                                                  END LOOP;
                      
                                              EXCEPTION
                                                  WHEN OTHERS THEN
                                                      continue;
                          
                                              END;
                  
                                          END LOOP;
                     
                            END IF;

                        END;"

        Dim cmd_insert_interv_int_cat As New OracleCommand(sql, Connection.conn)

        Try
            cmd_insert_interv_int_cat.CommandType = CommandType.Text
            cmd_insert_interv_int_cat.ExecuteNonQuery()
        Catch ex As Exception
            cmd_insert_interv_int_cat.Dispose()
            Return False
        End Try

        cmd_insert_interv_int_cat.Dispose()
        Return True

    End Function

    Function SET_INTERV_DEP_CLIN_SERV_FREQ(ByVal i_institution As Int64, ByVal i_software As Integer, ByVal i_a_interventions As interventions_alert_flg, ByVal i_dep_clin_serv As Int64, ByVal i_flg_type As Integer) As Boolean

        Dim sql As String = "DECLARE

                                l_a_interventions table_varchar := table_varchar(" & "'" & i_a_interventions.id_content_intervention & "'); "

        sql = sql & " l_id_intervention alert.intervention.id_intervention%TYPE;
                      l_id_insert_type  INTEGER :=" & i_flg_type & ";
                      l_id_dep_clin_serv  alert.dep_clin_serv.id_dep_clin_serv%type := " & i_dep_clin_serv & ";

                    BEGIN

                        FOR i IN 1 .. l_a_interventions.count()
                        LOOP
    
                            SELECT i.id_intervention
                            INTO l_id_intervention
                            FROM alert.intervention i
                            WHERE i.id_content = l_a_interventions(i)
                            and i.flg_status='A';
    
                      BEGIN
                            IF l_id_insert_type = 1
                            THEN
        
                                INSERT INTO alert.interv_dep_clin_serv
                                    (id_interv_dep_clin_serv, id_intervention, id_dep_clin_serv, flg_type, rank, id_institution, id_software, flg_execute, flg_timeout)
                                VALUES
                                    (alert.seq_interv_dep_clin_serv.nextval, l_id_intervention, l_id_dep_clin_serv, 'M', 0, " & i_institution & ", " & i_software & ", 'Y', 'N');
        
                            ELSIF l_id_insert_type = 2
                            THEN
        
                                INSERT INTO alert.interv_dep_clin_serv
                                    (id_interv_dep_clin_serv, id_intervention, id_dep_clin_serv, flg_type, rank, id_institution, id_software, flg_execute, flg_timeout)
                                VALUES
                                    (alert.seq_interv_dep_clin_serv.nextval, l_id_intervention, l_id_dep_clin_serv, 'A', 0, " & i_institution & ", " & i_software & ", 'Y', 'N');
        
                            ELSE
        
                                INSERT INTO alert.interv_dep_clin_serv
                                    (id_interv_dep_clin_serv, id_intervention, id_dep_clin_serv, flg_type, rank, id_institution, id_software, flg_execute, flg_timeout)
                                VALUES
                                    (alert.seq_interv_dep_clin_serv.nextval, l_id_intervention, l_id_dep_clin_serv, 'M', 0, " & i_institution & ", " & i_software & ", 'Y', 'N');
        
                                INSERT INTO alert.interv_dep_clin_serv
                                    (id_interv_dep_clin_serv, id_intervention, id_dep_clin_serv, flg_type, rank, id_institution, id_software, flg_execute, flg_timeout)
                                VALUES
                                    (alert.seq_interv_dep_clin_serv.nextval, l_id_intervention, l_id_dep_clin_serv, 'A', 0, " & i_institution & ", " & i_software & ", 'Y', 'N');
        
                            END IF;

                    EXCEPTION
                      WHEN DUP_VAL_ON_INDEX THEN
                        CONTINUE;
            
                        END;
    
                        END LOOP;

                    END;"

        Dim cmd_delete_interv_dep_clin_serv As New OracleCommand(sql, Connection.conn)

        Try
            cmd_delete_interv_dep_clin_serv.CommandType = CommandType.Text
            cmd_delete_interv_dep_clin_serv.ExecuteNonQuery()
        Catch ex As Exception
            cmd_delete_interv_dep_clin_serv.Dispose()
            Return False
        End Try

        cmd_delete_interv_dep_clin_serv.Dispose()
        Return True

    End Function

    Function DELETE_INTERV_INT_CAT(ByVal i_institution As Int64, ByVal i_software As Integer, ByVal i_intervention As interventions_default) As Boolean

        Dim sql As String = "DELETE FROM alert.interv_int_cat iic
                                WHERE iic.id_intervention IN (SELECT i.id_intervention
                                                              FROM alert.intervention i
                                                              WHERE i.id_content = '" & i_intervention.id_content_intervention & "'
                                                              AND i.flg_status = 'A')
                                and iic.id_interv_category in (SELECT ic.id_interv_category
                                                              FROM alert.interv_category ic
                                                              WHERE ic.id_content = '" & i_intervention.id_content_category & "'
                                                              AND ic.flg_available = 'Y')
                                 
                                and iic.id_software=" & i_software & "
                                and iic.id_institution in (0," & i_institution & ")"

        Dim cmd_delete_interv_int_cat As New OracleCommand(sql, Connection.conn)

        Try
            cmd_delete_interv_int_cat.CommandType = CommandType.Text
            cmd_delete_interv_int_cat.ExecuteNonQuery()
        Catch ex As Exception
            cmd_delete_interv_int_cat.Dispose()
            Return False
        End Try

        cmd_delete_interv_int_cat.Dispose()
        Return True

    End Function

    Function DELETE_INTERV_DEP_CLIN_SERV(ByVal i_institution As Int64, ByVal i_software As Integer, ByVal i_intervention As interventions_default, ByVal i_most_freq As Boolean, ByVal i_flg_type As Integer) As Boolean

        Dim sql As String

        If i_most_freq = False Then

            sql = "DELETE FROM alert.interv_dep_clin_serv dps
                                WHERE dps.id_intervention IN (SELECT i.id_intervention
                                                              FROM alert.intervention i
                                                              WHERE i.id_content = '" & i_intervention.id_content_intervention & "'
                                                              AND i.flg_status = 'A')
                                AND dps.id_institution IN (0, " & i_institution & ")
                                AND dps.id_software = " & i_software & ""

        Else

            sql = "DELETE FROM alert.interv_dep_clin_serv dps
                                WHERE dps.id_intervention IN (SELECT i.id_intervention
                                                              FROM alert.intervention i
                                                              WHERE i.id_content = '" & i_intervention.id_content_intervention & "'
                                                              AND i.flg_status = 'A')
                                AND dps.id_institution IN (0, " & i_institution & ")
                                AND dps.id_software = " & i_software & " "

            If i_flg_type = 0 Then

                sql = sql & "And dps.flg_type IN ('M', 'A') "

            ElseIf i_flg_type = 1 Then

                sql = sql & "And dps.flg_type IN ('M') "

            Else

                sql = sql & "And dps.flg_type IN ('A') "

            End If

        End If


        Dim cmd_delete_dep_clin_serv As New OracleCommand(sql, Connection.conn)

        Try
            cmd_delete_dep_clin_serv.CommandType = CommandType.Text
            cmd_delete_dep_clin_serv.ExecuteNonQuery()
        Catch ex As Exception
            cmd_delete_dep_clin_serv.Dispose()
            Return False
        End Try

        cmd_delete_dep_clin_serv.Dispose()
        Return True

    End Function

    Function SET_INTERV_INT_CAT_REMOVE(ByVal i_institution As Int64, ByVal i_software As Integer, ByVal i_intervention As interventions_default) As Boolean

        Dim sql As String

        sql = "DECLARE

                  l_id_intervention alert.intervention.id_intervention%type;
                  l_id_interv_cat   alert.interv_category.id_interv_category%type;

                BEGIN
  
                  SELECT I.ID_INTERVENTION
                  INTO   l_id_intervention
                  FROM ALERT.INTERVENTION I
                  WHERE I.ID_CONTENT = '" & i_intervention.id_content_intervention & "'
                  AND I.FLG_STATUS='A';
  
                  SELECT IC.ID_INTERV_CATEGORY
                  INTO l_id_interv_cat
                  FROM ALERT.INTERV_CATEGORY IC
                  WHERE IC.ID_CONTENT='" & i_intervention.id_content_category & "'
                  AND IC.FLG_AVAILABLE='Y';
    
                  INSERT INTO ALERT.INTERV_INT_CAT(ID_INTERV_CATEGORY, ID_INTERVENTION, RANK,ID_SOFTWARE, ID_INSTITUTION, FLG_ADD_REMOVE)
                  VALUES(l_id_interv_cat,l_id_intervention,0," & i_software & ", " & i_institution & ", 'R');

                EXCEPTION
                  WHEN DUP_VAL_ON_INDEX THEN
                    UPDATE ALERT.INTERV_INT_CAT IIC
                    SET IIC.FLG_ADD_REMOVE='R'
                    WHERE IIC.ID_INTERV_CATEGORY=l_id_interv_cat
                    AND IIC.ID_INTERVENTION = l_id_intervention
                    AND IIC.ID_SOFTWARE=" & i_software & "
                    AND IIC.ID_INSTITUTION IN (0," & i_institution & "); 
  
                END;"

        Dim cmd_set_iic_remove As New OracleCommand(sql, Connection.conn)

        Try
            cmd_set_iic_remove.CommandType = CommandType.Text
            cmd_set_iic_remove.ExecuteNonQuery()
        Catch ex As Exception
            cmd_set_iic_remove.Dispose()
            Return False
        End Try

        cmd_set_iic_remove.Dispose()
        Return True

    End Function

End Class
