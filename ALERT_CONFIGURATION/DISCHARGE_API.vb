Imports Oracle.DataAccess.Client
Public Class DISCHARGE_API

    Dim db_access_general As New General

    Public Structure DEFAULT_DISCAHRGE
        Public id_disch_reas_dest As Int64
        Public id_content As String
        Public desccription As String
        Public id_clinical_service As String
        Public type As String
    End Structure

    Public Structure DEFAULT_REASONS
        Public id_content As String
        Public desccription As String
    End Structure

    Public Structure DEFAULT_DISCH_PROFILE
        Public ID_PROFILE_DISCH_REASON As Int64
        Public ID_PROFILE_TEMPLATE As Int64
        Public PROFILE_NAME As String
    End Structure

    Function GET_DEFAULT_VERSIONS(ByVal i_institution As Int64, ByVal i_software As Integer, ByRef i_dr As OracleDataReader) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution)

        Dim sql As String = "SELECT DISTINCT drd.version
                                FROM alert_default.discharge_reason dr
                                JOIN alert_default.discharge_reason_mrk_vrs drmv ON drmv.id_discharge_reason = dr.id_discharge_reason
                                JOIN alert_default.disch_reas_dest drd ON drd.id_discharge_reason = dr.id_discharge_reason
                                                                   AND drd.id_market = drmv.id_market
                                                                   AND drd.version = drmv.version
                                JOIN alert_default.profile_disch_reason pdr ON pdr.id_discharge_reason = dr.id_discharge_reason
                                JOIN institution i ON i.id_market = drmv.id_market
                                WHERE dr.flg_available = 'Y'
                                AND i.id_institution = " & i_institution & "
                                AND drd.flg_active = 'A'
                                AND drd.id_software_param = " & i_software & "
                                AND pdr.flg_available = 'Y'
                                ORDER BY 1 ASC"

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

    Function GET_DEFAULT_REASONS(ByVal i_institution As Int64, ByVal i_software As Integer, ByVal i_version As String, ByRef i_dr As OracleDataReader) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution)

        'Modificar o output. Passar apenas ID_CONTENT e DESCRITIVO. O Resto será chamado diretamente pela função responsável por incluir Reason e Dest na BD
        Dim sql As String = "SELECT DISTINCT dr.id_content,
                                                alert_default.pk_translation_default.get_translation_default(" & l_id_language & ", dr.code_discharge_reason)
                                               
                                FROM alert_default.discharge_reason dr
                                JOIN alert_default.discharge_reason_mrk_vrs drmv ON drmv.id_discharge_reason = dr.id_discharge_reason
                                JOIN alert_default.disch_reas_dest drd ON drd.id_discharge_reason = dr.id_discharge_reason
                                                                   AND drd.id_market = drmv.id_market
                                                                   AND drd.version = drmv.version
                                JOIN alert_default.profile_disch_reason pdr ON pdr.id_discharge_reason = dr.id_discharge_reason
                                JOIN INSTITUTION I ON I.id_market=DRMV.ID_MARKET
                                WHERE dr.flg_available = 'Y'
                                AND I.id_institution=" & i_institution & "
                                AND drmv.version = '" & i_version & "'
                                AND drd.flg_active = 'A'
                                AND drd.id_software_param = " & i_software & "
                                AND pdr.flg_available = 'Y'
                                order by 2 asc"

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

    Function GET_DEFAULT_DESTINATIONS(ByVal i_institution As Int64, ByVal i_software As Integer, ByVal i_reason As String, ByVal i_version As String, ByRef i_dr As OracleDataReader) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution)

        Dim sql As String = "SELECT DISTINCT drd.id_disch_reas_dest,
                                                nvl(d.id_content, dr.id_content) AS id_content,
                                                nvl2(d.id_content, alert_default.pk_translation_default.get_translation_default(" & l_id_language & ", d.code_discharge_dest),
                                                     alert_default.pk_translation_default.get_translation_default(" & l_id_language & ", dr.code_discharge_reason)) AS description,
                                                nvl(dcs.id_content, -1) AS clinical_service,
                                                nvl2(d.id_content, 'D', 'R') AS TYPE

                                FROM alert_default.discharge_reason dr
                                JOIN alert_default.discharge_reason_mrk_vrs drmv ON drmv.id_discharge_reason = dr.id_discharge_reason
                                JOIN alert_default.disch_reas_dest drd ON drd.id_discharge_reason = dr.id_discharge_reason
                                                                   AND drd.id_market = drmv.id_market
                                                                   AND drd.version = drmv.version
                                JOIN alert_default.profile_disch_reason pdr ON pdr.id_discharge_reason = dr.id_discharge_reason
                                LEFT JOIN alert_default.discharge_dest d ON d.id_discharge_dest = drd.id_discharge_dest
                                                                     AND d.flg_available = 'Y'
                                JOIN institution i ON i.id_market = drmv.id_market
                                LEFT JOIN alert_default.discharge_dest_mrk_vrs dv ON dv.id_discharge_dest = d.id_discharge_dest
                                                                              AND dv.id_market = i.id_market
                                                                              AND dv.version = drd.version
                                LEFT JOIN alert_default.clinical_service dcs ON dcs.id_clinical_service = drd.id_clinical_service

                                WHERE dr.flg_available = 'Y'
                                AND i.id_institution = " & i_institution & "
                                AND drmv.version = '" & i_version & "'
                                AND drd.flg_active = 'A'
                                AND drd.id_software_param = " & i_software & "
                                AND pdr.flg_available = 'Y'
                                AND dr.id_content = '" & i_reason & "'
                                ORDER BY 3 ASC"

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

    Function GET_DEFAULT_PROFILE_DISCH_REASON(ByVal id_disch_reason As String, ByRef o_profile_templates As OracleDataReader) As Boolean

        Dim sql As String = "SELECT PDR.ID_PROFILE_DISCH_REASON, pdr.id_profile_template,PT.INTERN_NAME_TEMPL
                                FROM alert_default.discharge_reason dr
                                JOIN alert_default.profile_disch_reason pdr ON pdr.id_discharge_reason = dr.id_discharge_reason
                                JOIN alert.profile_template pt ON pt.id_profile_template = pdr.id_profile_template
                                WHERE dr.flg_available = 'Y'
                                AND dr.id_content = '" & id_disch_reason & "'
                                AND pdr.flg_available = 'Y'
                                ORDER BY 2 ASC"

        Dim cmd As New OracleCommand(sql, Connection.conn)

        Try
            cmd.CommandType = CommandType.Text
            o_profile_templates = cmd.ExecuteReader()
            cmd.Dispose()
            Return True
        Catch ex As Exception
            cmd.Dispose()
            Return False
        End Try

    End Function

    Function CHECK_REASON(ByVal i_id_reason As String) As Boolean

        Dim sql As String = "SELECT COUNT(*)
                                FROM alert.discharge_reason dr
                                WHERE dr.flg_available = 'Y'
                                AND dr.id_content = '" & i_id_reason & "'"

        Dim cmd As New OracleCommand(sql, Connection.conn)

        Dim dr As OracleDataReader

        cmd.CommandType = CommandType.Text
        dr = cmd.ExecuteReader()
        cmd.Dispose()

        Dim l_total_record As Integer = 0

        While dr.Read()
            l_total_record = dr.Item(0)
        End While

        If l_total_record > 0 Then
            Return True
        Else
            Return False
        End If

    End Function

    Function CHECK_REASON_translation(ByVal i_institution As Int64, ByVal i_id_reason As String) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution)

        Dim sql As String = "SELECT COUNT(*)
                                FROM alert.discharge_reason dr
                                JOIN TRANSLATION T ON T.CODE_TRANSLATION=DR.CODE_DISCHARGE_REASON
                                WHERE dr.flg_available = 'Y'
                                AND dr.id_content = '" & i_id_reason & "'
                                And T.DESC_LANG_" & l_id_language & " Is Not NULL"

        Dim cmd As New OracleCommand(sql, Connection.conn)

        Dim dr As OracleDataReader

        cmd.CommandType = CommandType.Text
        dr = cmd.ExecuteReader()
        cmd.Dispose()

        Dim l_total_record As Integer = 0

        While dr.Read()
            l_total_record = dr.Item(0)
        End While

        If l_total_record > 0 Then
            Return True
        Else
            Return False
        End If

    End Function

    Function SET_REASON(ByVal i_institution As Int64, ByVal i_id_reason As String) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution)

        Dim sql As String = "DECLARE

                                    l_flg_admin alert.discharge_reason.flg_admin_medic%TYPE;

                                    l_file_execute alert.discharge_reason.file_to_execute%TYPE;

                                    l_id_content alert.discharge_reason.id_content%TYPE := '" & i_id_reason & "';

                                    l_id_alert_reason alert.discharge_reason.id_discharge_reason%TYPE;

                                    l_default_desc alert_default.translation.desc_lang_6%TYPE;

                                BEGIN

                                    SELECT dr.flg_admin_medic, dr.file_to_execute
                                    INTO l_flg_admin, l_file_execute
                                    FROM alert_default.discharge_reason dr
                                    WHERE dr.id_content = l_id_content
                                    AND dr.flg_available = 'Y';

                                    l_id_alert_reason := alert.seq_discharge_reason.nextval;

                                    INSERT INTO alert.discharge_reason
                                        (id_discharge_reason, code_discharge_reason, flg_admin_medic, flg_available, rank, file_to_execute, id_content)
                                    VALUES
                                        (l_id_alert_reason, 'DISCHARGE_REASON.CODE_DISCHARGE_REASON.' || l_id_alert_reason, l_flg_admin, 'Y', 1, l_file_execute, l_id_content);

                                    SELECT t.desc_lang_" & l_id_language & "
                                    INTO l_default_desc
                                    FROM alert_default.discharge_reason dr
                                    JOIN alert_default.translation t ON t.code_translation = dr.code_discharge_reason
                                    WHERE dr.id_content = l_id_content
                                    AND dr.flg_available = 'Y';

                                    pk_translation.insert_into_translation(" & l_id_language & ", 'DISCHARGE_REASON.CODE_DISCHARGE_REASON.' || l_id_alert_reason, l_default_desc);

                                EXCEPTION
                                    WHEN dup_val_on_index THEN
                                        dbms_output.put_line('Registo já existente.');
    
                                END;"

        Dim cmd_insert_reason As New OracleCommand(sql, Connection.conn)

        Try
            cmd_insert_reason.CommandType = CommandType.Text
            cmd_insert_reason.ExecuteNonQuery()
        Catch ex As Exception
            cmd_insert_reason.Dispose()
            Return False
        End Try

        cmd_insert_reason.Dispose()

        Return True

    End Function

    Function SET_REASON_TRANSLATION(ByVal i_institution As Int64, ByVal i_id_reason As String) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution)

        Dim sql As String = "DECLARE

                                    l_id_content alert.discharge_reason.id_content%TYPE := '" & i_id_reason & "';

                                    l_id_alert_reason alert.discharge_reason.id_discharge_reason%TYPE;

                                    l_default_desc alert_default.translation.desc_lang_6%TYPE;

                                BEGIN

                                    SELECT t.desc_lang_" & l_id_language & "
                                    INTO l_default_desc
                                    FROM alert_default.discharge_reason dr
                                    JOIN alert_default.translation t ON t.code_translation = dr.code_discharge_reason
                                    WHERE dr.id_content = l_id_content
                                    AND dr.flg_available = 'Y';

                                    SELECT dr.id_discharge_reason
                                    INTO  l_id_alert_reason
                                    FROM ALERT.discharge_reason dr
                                    WHERE dr.id_content = l_id_content
                                    AND dr.flg_available = 'Y';

                                    pk_translation.insert_into_translation(" & l_id_language & ", 'DISCHARGE_REASON.CODE_DISCHARGE_REASON.' || l_id_alert_reason, l_default_desc);

                                EXCEPTION
                                    WHEN dup_val_on_index THEN
                                        dbms_output.put_line('Registo já existente.');
    
                                END;"

        Dim cmd_insert_reason_translation As New OracleCommand(sql, Connection.conn)

        Try
            cmd_insert_reason_translation.CommandType = CommandType.Text
            cmd_insert_reason_translation.ExecuteNonQuery()
        Catch ex As Exception
            cmd_insert_reason_translation.Dispose()
            Return False
        End Try

        cmd_insert_reason_translation.Dispose()

        Return True

    End Function

    Function SET_PROFILE_DISCH_REASON(ByVal i_institution As Int64, ByVal a_prof_disch_reason As DEFAULT_DISCH_PROFILE()) As Boolean

        Dim sql As String = "DECLARE

                                l_profile_disch_reason table_number := table_number("

        For i As Integer = 0 To a_prof_disch_reason.Count() - 1

            If (i < a_prof_disch_reason.Count() - 1) Then

                sql = sql & "'" & a_prof_disch_reason(i).ID_PROFILE_DISCH_REASON & "', "

            Else

                sql = sql & "'" & a_prof_disch_reason(i).ID_PROFILE_DISCH_REASON & "');"

            End If

        Next

        sql = sql & "           l_institution institution.id_institution%type := " & i_institution & "; 
    
                                l_id_profile_template  alert.profile_template.id_profile_template%TYPE;
                                l_id_discharge_files   alert.profile_disch_reason.id_discharge_flash_files%TYPE;
                                l_flg_access           alert.profile_disch_reason.flg_access%TYPE;
                                l_rank                 alert.profile_disch_reason.rank%TYPE;
                                l_flg_default          alert.profile_disch_reason.flg_default%TYPE;

                                l_id_disch_reason      alert.profile_disch_reason.id_discharge_reason%TYPE;

                                --#############################################################################################
                                FUNCTION check_prof_disch_reason
                                (
                                    i_discharge_reason  IN alert.profile_disch_reason.id_discharge_reason%TYPE,
                                    i_id_profile_template IN alert.profile_disch_reason.id_profile_template%TYPE,
                                    i_id_institution      IN alert.profile_disch_reason.id_institution%TYPE
                                ) RETURN BOOLEAN IS
    
                                    l_count INTEGER := 0;
    
                                BEGIN
        
                                    SELECT COUNT(*)
                                    INTO l_count
                                    FROM alert.profile_disch_reason pdr
                                    WHERE pdr.id_discharge_reason = i_discharge_reason
                                    AND pdr.id_profile_template = i_id_profile_template
                                    AND pdr.id_institution = i_id_institution
                                    AND pdr.flg_available = 'Y';

                                    IF l_count > 0
                                    THEN
                                        RETURN TRUE;
                                    ELSE
                                        RETURN FALSE;
                                    END IF;
    
                                END check_prof_disch_reason;
                                --#############################################################################################

                            BEGIN

                                FOR i IN 1 .. l_profile_disch_reason.count()
                                LOOP
                                    --1 - OBTER O ID_ALERT DA REASON, E OS DADOS DO PROFILE_DISCH_REASON
                                    SELECT dr.id_discharge_reason, dpdr.id_profile_template, dpdr.id_discharge_flash_files, dpdr.flg_access, dpdr.rank, dpdr.flg_default
                                    INTO l_id_disch_reason, l_id_profile_template, l_id_discharge_files, l_flg_access, l_rank, l_flg_default
                                    FROM alert_default.profile_disch_reason dpdr
                                    JOIN alert_default.discharge_reason ddr ON ddr.id_discharge_reason = dpdr.id_discharge_reason
                                    JOIN alert.discharge_reason dr ON dr.id_content = ddr.id_content
                                                               AND dr.flg_available = 'Y'
                                    WHERE dpdr.id_profile_disch_reason = l_profile_disch_reason(i);
    
                                    --2 - VERIFICAR SE O REGISTO JÁ EXISTE. SE NÃO EXISTIR, INSERE.
                                    IF NOT check_prof_disch_reason(l_id_disch_reason, l_id_profile_template, l_institution)
                                    THEN
                                        INSERT INTO alert.profile_disch_reason
                                            (id_profile_disch_reason,
                                             id_discharge_reason,
                                             id_profile_template,
                                             id_institution,
                                             flg_available,
                                             id_discharge_flash_files,
                                             flg_access,
                                             rank,
                                             flg_default)
                                        VALUES
                                            (alert.seq_profile_disch_reason.nextval,
                                             l_id_disch_reason,
                                             l_id_profile_template,
                                             l_institution,
                                             'Y',
                                             l_id_discharge_files,
                                             l_flg_access,
                                             l_rank,
                                             l_flg_default);
                                    ELSE
                                        dbms_output.put_line('Registo já existente!');
                                    END IF;
                                END LOOP;
                            END;"

        Dim cmd_insert_profile_disch_reason As New OracleCommand(sql, Connection.conn)

        Try
            cmd_insert_profile_disch_reason.CommandType = CommandType.Text
            cmd_insert_profile_disch_reason.ExecuteNonQuery()
        Catch ex As Exception
            cmd_insert_profile_disch_reason.Dispose()
            Return False
        End Try

        cmd_insert_profile_disch_reason.Dispose()

        Return True

    End Function

    Function CHECK_DESTINATION(ByVal i_id_destination As String) As Boolean

        Dim sql As String = "SELECT COUNT(*)
                                FROM alert.discharge_dest dd
                                WHERE dd.flg_available = 'Y'
                                AND dd.id_content = '" & i_id_destination & "'"

        Dim cmd As New OracleCommand(sql, Connection.conn)

        Dim dr As OracleDataReader

        cmd.CommandType = CommandType.Text
        dr = cmd.ExecuteReader()
        cmd.Dispose()

        Dim l_total_record As Integer = 0

        While dr.Read()
            l_total_record = dr.Item(0)
        End While

        If l_total_record > 0 Then
            Return True
        Else
            Return False
        End If

    End Function

    Function CHECK_DESTINATION_TRANSLATION(ByVal i_institution As Int64, ByVal i_id_destination As String) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution)

        Dim sql As String = "SELECT COUNT(*)
                                FROM alert.discharge_dest dd
                                JOIN TRANSLATION T ON T.CODE_TRANSLATION=dd.code_discharge_dest
                                WHERE dd.flg_available = 'Y'
                                AND dd.id_content = '" & i_id_destination & "'
                                And T.DESC_LANG_" & l_id_language & " Is Not NULL"

        Dim cmd As New OracleCommand(sql, Connection.conn)

        Dim dr As OracleDataReader

        cmd.CommandType = CommandType.Text
        dr = cmd.ExecuteReader()
        cmd.Dispose()

        Dim l_total_record As Integer = 0

        While dr.Read()
            l_total_record = dr.Item(0)
        End While

        If l_total_record > 0 Then
            Return True
        Else
            Return False
        End If

    End Function

    Function SET_DESTINATION(ByVal i_institution As Int64, ByVal i_id_destination As String) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution)

        'Modificar o Código PARA ACEITAR ARRAYS
        Dim sql As String = "DECLARE                                  

                                    l_id_content alert.discharge_reason.id_content%TYPE := '" & i_id_destination & "';

                                    l_flg_type alert.discharge_dest.flg_type%type;

                                    l_default_desc alert_default.translation.desc_lang_6%TYPE;

                                    l_id_alert_destination ALERT.DISCHARGE_DEST.ID_DISCHARGE_DEST%TYPE;

                                BEGIN

                                    SELECT dd.flg_type
                                    INTO l_flg_type
                                    FROM alert_default.discharge_dest dd
                                    WHERE dd.id_content = '" & i_id_destination & "'
                                    AND dd.flg_available = 'Y';

                                    l_id_alert_destination := alert.seq_discharge_dest.nextval;

                                    insert into ALERT.discharge_dest (ID_DISCHARGE_DEST, CODE_DISCHARGE_DEST, FLG_AVAILABLE, RANK, FLG_TYPE, ID_CONTENT)
                                    values (l_id_alert_destination, 'DISCHARGE_DEST.CODE_DISCHARGE_DEST.' || l_id_alert_destination, 'Y', 0, l_flg_type, l_id_content);
                                    
                                    SELECT t.desc_lang_" & l_id_language & "
                                    INTO l_default_desc
                                    FROM alert_default.discharge_dest dd
                                    JOIN alert_default.translation t ON t.code_translation = dd.code_discharge_dest
                                    WHERE dd.id_content = l_id_content
                                    AND dd.flg_available = 'Y';     

                                    pk_translation.insert_into_translation(" & l_id_language & ", 'DISCHARGE_DEST.CODE_DISCHARGE_DEST.' || l_id_alert_destination, l_default_desc);

                                EXCEPTION
                                    WHEN dup_val_on_index THEN
                                        dbms_output.put_line('Registo já existente.');
    
                                END;"

        Dim cmd_insert_reason As New OracleCommand(sql, Connection.conn)

        Try
            cmd_insert_reason.CommandType = CommandType.Text
            cmd_insert_reason.ExecuteNonQuery()
        Catch ex As Exception
            cmd_insert_reason.Dispose()
            Return False
        End Try

        cmd_insert_reason.Dispose()

        Return True

    End Function

End Class
