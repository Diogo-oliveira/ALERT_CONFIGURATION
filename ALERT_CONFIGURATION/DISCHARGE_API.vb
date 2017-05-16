Imports Oracle.DataAccess.Client
Public Class DISCHARGE_API

    Dim db_access_general As New General
    Dim db_clin_serv As New CLINICAL_SERVICE_API

    Public Structure DEFAULT_DISCAHRGE
        Public id_disch_reas_dest As Int64
        Public id_content As String
        Public description As String
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

    Public Structure DEFAULT_INSTR '(Esta estrutura vai ser usada pelo Grupo e pelas instruções)
        Public ID_CONTENT As String
        Public DESCRIPTION As String
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
                                JOIN ALERT.PROFILE_TEMPLATE PT ON PT.ID_PROFILE_TEMPLATE=PDR.ID_PROFILE_TEMPLATE AND PT.ID_SOFTWARE=DRD.ID_SOFTWARE_PARAM AND PT.FLG_AVAILABLE='Y'                                
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

        Dim sql As String = "SELECT DISTINCT dr.id_content,
                                                alert_default.pk_translation_default.get_translation_default(" & l_id_language & ", dr.code_discharge_reason)
                                               
                                FROM alert_default.discharge_reason dr
                                JOIN alert_default.discharge_reason_mrk_vrs drmv ON drmv.id_discharge_reason = dr.id_discharge_reason
                                JOIN alert_default.disch_reas_dest drd ON drd.id_discharge_reason = dr.id_discharge_reason
                                                                   AND drd.id_market = drmv.id_market
                                                                   AND drd.version = drmv.version
                                JOIN alert_default.profile_disch_reason pdr ON pdr.id_discharge_reason = dr.id_discharge_reason
                                JOIN ALERT.PROFILE_TEMPLATE PT ON PT.ID_PROFILE_TEMPLATE=PDR.ID_PROFILE_TEMPLATE AND PT.ID_SOFTWARE=DRD.ID_SOFTWARE_PARAM AND PT.FLG_AVAILABLE='Y'
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

    'Obter todas as reasons do default, mesmo as que não estão disponíveisl
    Function GET_ALL_DEFAULT_REASONS(ByVal i_institution As Int64, ByRef i_dr As OracleDataReader) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution)

        Dim sql As String = "SELECT DISTINCT dr.id_content,
                                             alert_default.pk_translation_default.get_translation_default(" & l_id_language & ", dr.code_discharge_reason)
                                               
                                FROM alert_default.discharge_reason dr
 
                                WHERE alert_default.pk_translation_default.get_translation_default(" & l_id_language & ", dr.code_discharge_reason) IS NOT NULL
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
                                JOIN ALERT.PROFILE_TEMPLATE PT ON PT.ID_PROFILE_TEMPLATE=PDR.ID_PROFILE_TEMPLATE AND PT.ID_SOFTWARE=DRD.ID_SOFTWARE_PARAM AND PT.FLG_AVAILABLE='Y'
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

    'Obter todas as destinations do default, mesmo as que não estão disponíveis
    Function GET_ALL_DEFAULT_DESTINATIONS(ByVal i_institution As Int64, ByRef i_dr As OracleDataReader) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution)

        Dim sql As String = "SELECT DISTINCT dd.id_content,
                                             alert_default.pk_translation_default.get_translation_default(" & l_id_language & ", dd.code_discharge_dest)
                                               
                                FROM alert_default.discharge_dest dd
 
                                WHERE alert_default.pk_translation_default.get_translation_default(" & l_id_language & ", dd.code_discharge_dest) IS NOT NULL
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

    Function GET_DEFAULT_PROFILE_DISCH_REASON(ByVal i_software As Integer, ByVal id_disch_reason As String, ByRef o_profile_templates As OracleDataReader) As Boolean

        Dim sql As String = "SELECT PDR.ID_PROFILE_DISCH_REASON, pdr.id_profile_template,PT.INTERN_NAME_TEMPL
                                FROM alert_default.discharge_reason dr
                                JOIN alert_default.profile_disch_reason pdr ON pdr.id_discharge_reason = dr.id_discharge_reason
                                JOIN alert.profile_template pt ON pt.id_profile_template = pdr.id_profile_template
                                WHERE dr.flg_available = 'Y'
                                AND dr.id_content = '" & id_disch_reason & "'
                                AND pdr.flg_available = 'Y'
                                AND pt.id_software = " & i_software & "
                                AND PT.FLG_AVAILABLE='Y'
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

    Function UPDATE_REASON(ByVal i_reason As String, ByVal i_profile_template As String, ByVal i_rank As Integer, ByVal i_file_to_exeecute As String) As Boolean


        Dim Sql As String = "UPDATE alert.discharge_reason dr
                                SET dr.flg_admin_medic = '" & i_profile_template & "', dr.rank = " & i_rank & ", dr.file_to_execute = '" & i_file_to_exeecute & "'
                                WHERE dr.id_content = '" & i_reason & "'
                                AND dr.flg_available = 'Y'"

        Dim cmd_update_reason As New OracleCommand(Sql, Connection.conn)

        Try
            cmd_update_reason.CommandType = CommandType.Text
            cmd_update_reason.ExecuteNonQuery()
        Catch ex As Exception
            cmd_update_reason.Dispose()
            Return False
        End Try

        cmd_update_reason.Dispose()

        Return True

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

    Function SET_MANUAL_REASON(ByVal i_institution As Int64, ByVal i_id_reason As String, ByVal flg_admin_medic As String, ByVal rank As Integer, ByVal i_file_execute As String) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution)

        Dim sql As String = "DECLARE

                                    l_flg_admin alert.discharge_reason.flg_admin_medic%TYPE := '" & flg_admin_medic & "';

                                    l_file_execute alert.discharge_reason.file_to_execute%TYPE := '" & i_file_execute & "';

                                    l_id_content alert.discharge_reason.id_content%TYPE := '" & i_id_reason & "';

                                    l_id_alert_reason alert.discharge_reason.id_discharge_reason%TYPE;

                                    l_default_desc alert_default.translation.desc_lang_6%TYPE;

                                BEGIN

                                    l_id_alert_reason := alert.seq_discharge_reason.nextval;

                                    INSERT INTO alert.discharge_reason
                                        (id_discharge_reason, code_discharge_reason, flg_admin_medic, flg_available, rank, file_to_execute, id_content)
                                    VALUES
                                        (l_id_alert_reason, 'DISCHARGE_REASON.CODE_DISCHARGE_REASON.' || l_id_alert_reason, l_flg_admin, 'Y', " & rank & ", l_file_execute, l_id_content);

                                    SELECT t.desc_lang_" & l_id_language & "
                                    INTO l_default_desc
                                    FROM alert_default.discharge_reason dr
                                    JOIN alert_default.translation t ON t.code_translation = dr.code_discharge_reason
                                    WHERE dr.id_content = l_id_content;

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
                                    WHERE dr.id_content = l_id_content;
                                    --AND dr.flg_available = 'Y'; Para garantir que mesmo que a reason não esteja available no default a sua tradução seja inserida no alert (config manual)

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

    Function SET_DESTINATION(ByVal i_institution As Int64, ByVal i_id_destination As DEFAULT_DISCAHRGE) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution)

        Dim sql As String = "DECLARE

                                l_id_content table_varchar := table_varchar('" & i_id_destination.id_content & "');"

        sql = sql & "
                                    l_flg_type alert.discharge_dest.flg_type%TYPE;

                                    l_default_desc alert_default.translation.desc_lang_6%TYPE;

                                    l_id_alert_destination alert.discharge_dest.id_discharge_dest%TYPE;

                                BEGIN

                                    FOR i IN 1 .. l_id_content.count()
                                    LOOP
    
                                        BEGIN
             
                                        DBMS_OUTPUT.put_line(i);
        
                                            SELECT dd.flg_type
                                            INTO l_flg_type
                                            FROM alert_default.discharge_dest dd
                                            WHERE dd.id_content = l_id_content(i)
                                            AND dd.flg_available = 'Y';
        
                                            l_id_alert_destination := alert.seq_discharge_dest.nextval;
        
                                            INSERT INTO alert.discharge_dest
                                                (id_discharge_dest, code_discharge_dest, flg_available, rank, flg_type, id_content)
                                            VALUES
                                                (l_id_alert_destination, 'DISCHARGE_DEST.CODE_DISCHARGE_DEST.' || l_id_alert_destination, 'Y', 0, l_flg_type, l_id_content(i));
        
                                            SELECT t.desc_lang_" & l_id_language & "
                                            INTO l_default_desc
                                            FROM alert_default.discharge_dest dd
                                            JOIN alert_default.translation t ON t.code_translation = dd.code_discharge_dest
                                            WHERE dd.id_content = l_id_content(i)
                                            AND dd.flg_available = 'Y';
        
                                            pk_translation.insert_into_translation(" & l_id_language & ", 'DISCHARGE_DEST.CODE_DISCHARGE_DEST.' || l_id_alert_destination,
                                                                                   l_default_desc);
        
                                        EXCEPTION
                                            WHEN dup_val_on_index THEN
                                                continue;
            
                                        END;
    
                                    END LOOP;

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

    Function SET_MANUAL_DESTINATION(ByVal i_institution As Int64, ByVal i_id_destination As String, ByVal i_rank As Integer, ByVal i_flg_type As String) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution)

        Dim sql As String = "DECLARE

                               l_id_content alert.discharge_reason.id_content%TYPE := '" & i_id_destination & "';"

        sql = sql & "
                                    l_flg_type alert.discharge_dest.flg_type%TYPE := '" & i_flg_type & "';

                                    l_default_desc alert_default.translation.desc_lang_6%TYPE;

                                    l_id_alert_destination alert.discharge_dest.id_discharge_dest%TYPE;

                             BEGIN
                            
                                            l_id_alert_destination := alert.seq_discharge_dest.nextval;
        
                                            INSERT INTO alert.discharge_dest
                                                (id_discharge_dest, code_discharge_dest, flg_available, rank, flg_type, id_content)
                                            VALUES
                                                (l_id_alert_destination, 'DISCHARGE_DEST.CODE_DISCHARGE_DEST.' || l_id_alert_destination, 'Y', " & i_rank & ", l_flg_type, l_id_content);
        
                                            SELECT t.desc_lang_" & l_id_language & "
                                            INTO l_default_desc
                                            FROM alert_default.discharge_dest dd
                                            JOIN alert_default.translation t ON t.code_translation = dd.code_discharge_dest
                                            WHERE dd.id_content = l_id_content;
        
                                            pk_translation.insert_into_translation(" & l_id_language & ", 'DISCHARGE_DEST.CODE_DISCHARGE_DEST.' || l_id_alert_destination,
                                                                                   l_default_desc);
        
                                        EXCEPTION
                                            WHEN dup_val_on_index THEN
                                                DBMS_OUTPUT.PUT_LINE('REPEATED RECORD');

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

    Function UPDATE_DESTINATION(ByVal i_id_destination As String, ByVal i_rank As Integer, ByVal i_flg_type As String) As Boolean

        Dim Sql As String = "UPDATE alert.discharge_dest dd
                                SET dd.flg_type='" & i_flg_type & "', dd.rank=" & i_rank & "
                                WHERE dd.id_content = '" & i_id_destination & "'
                                AND dd.flg_available = 'Y'"

        Dim cmd_update_destination As New OracleCommand(Sql, Connection.conn)

        Try
            cmd_update_destination.CommandType = CommandType.Text
            cmd_update_destination.ExecuteNonQuery()
        Catch ex As Exception
            cmd_update_destination.Dispose()
            Return False
        End Try

        cmd_update_destination.Dispose()

        Return True

    End Function

    Function SET_DESTINATION_TRANSLATION(ByVal i_institution As Int64, ByVal i_id_destination As DEFAULT_DISCAHRGE) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution)

        Dim sql As String = "DECLARE

                                l_id_content table_varchar := table_varchar('" & i_id_destination.id_content & "');"

        sql = sql & " 
                                l_id_alert_destination alert.discharge_dest.id_discharge_dest%TYPE;

                                l_default_desc alert_default.translation.desc_lang_6%TYPE;

                            BEGIN

                                FOR i IN 1 .. l_id_content.count()
                                LOOP
    
                                    BEGIN
        
                                        SELECT t.desc_lang_" & l_id_language & "
                                        INTO l_default_desc
                                        FROM alert_default.discharge_dest dd
                                        JOIN alert_default.translation t ON t.code_translation = dd.code_discharge_dest
                                        WHERE dd.id_content = l_id_content(i);
                                        --AND dd.flg_available = 'Y'; Por causa de configuração manual
        
                                        SELECT dd.id_discharge_dest
                                        INTO l_id_alert_destination
                                        FROM alert.discharge_dest dd
                                        WHERE dd.id_content = l_id_content(i)
                                        AND dd.flg_available = 'Y';
        
                                        pk_translation.insert_into_translation(" & l_id_language & ", 'DISCHARGE_DEST.CODE_DISCHARGE_DEST.' || l_id_alert_destination, l_default_desc);
        
                                    EXCEPTION
                                        WHEN dup_val_on_index THEN
                                            continue;
            
                                    END;
    
                                END LOOP;

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

    Function SET_DISCH_REAS_DEST(ByVal i_institution As Int64, ByVal i_software As Integer, ByVal i_reason As String, ByVal i_destinations() As DEFAULT_DISCAHRGE) As Boolean

        'Adaptar para receber todos os parêmetros
        'Vão ser passadas todas as Destination, mesmo que sejam do tipo 'R'. A verificação será feita no SQL. (Isto para garantir 
        'que não se perdem clinical services)

        Dim sql As String = "DECLARE
                                    l_id_software       alert.disch_reas_dest.id_software_param%TYPE := " & i_software & ";
                                    l_id_institution    alert.disch_reas_dest.id_instit_param%TYPE := " & i_institution & ";
                                    l_id_content_reason alert.discharge_reason.id_content%TYPE := '" & i_reason & "';

                                    l_id_def_disch_reas_dest table_number := table_number("

        For i As Integer = 0 To i_destinations.Count() - 1

            If (i < i_destinations.Count() - 1) Then

                sql = sql & i_destinations(i).id_disch_reas_dest & ", "

            Else

                sql = sql & i_destinations(i).id_disch_reas_dest & ");"

            End If

        Next

        sql = sql & "  l_id_content_discharge   table_varchar := table_varchar("


        For i As Integer = 0 To i_destinations.Count() - 1

            If (i < i_destinations.Count() - 1) Then

                sql = sql & " '" & i_destinations(i).id_content & "', "

            Else

                sql = sql & "'" & i_destinations(i).id_content & "');"

            End If

        Next

        'A - Tratar dos clinical services
        Dim l_a_clin_serv(i_destinations.Count() - 1) As String 'Array que recebe os id_content dos clnical services
        Dim l_a_dep_clin_serv(i_destinations.Count() - 1) As Int64 'Array que vai receber os ids dos dep_clin_servs

        For i As Integer = 0 To i_destinations.Count() - 1

            l_a_clin_serv(i) = i_destinations(i).id_clinical_service

        Next

        'Tratamento de clinical service
        For i As Integer = 0 To l_a_clin_serv.Count() - 1

            'A1 - VER SE CLINICAL SERVICE EXISTE NO ALERT E/OU SE TEM TRADUÇÃO
            If l_a_clin_serv(i) <> "-1" Then

                If Not db_clin_serv.CHECK_CLIN_SERV(l_a_clin_serv(i)) Then

                    If Not db_clin_serv.SET_CLIN_SERV(i_institution, l_a_clin_serv(i)) Then

                        Return False

                    End If

                ElseIf db_clin_serv.CHECK_CLIN_SERV_TRANSLATION(i_institution, l_a_clin_serv(i)) Then

                    If Not db_clin_serv.SET_CLIN_SERV_TRANSLATION(i_institution, l_a_clin_serv(i)) Then

                        Return False

                    End If

                End If

                Dim l_a_departments As Int64() 'Array que vai receber os ids dos departaments da instituição para um determinado software

                'A2 - Obter/Configurar DEP_CLIN_SERV
                Dim l_override As Boolean = False 'Determinar se utilizador não quer configurar o clin_serv para o software escolhido

                If Not db_clin_serv.CHECK_DEP_CLIN_SERV(i_institution, i_software, l_a_clin_serv(i)) Then

                    Dim result As Integer = 0
                    result = MsgBox("Clinical Service " & db_clin_serv.GET_CLIN_SERV_TRANSLATION(i_institution, l_a_clin_serv(i)) &
                                    " not configured for the chosen institution and software. Do you wish to configure it?", MessageBoxButtons.YesNo)

                    If result = DialogResult.Yes Then

                        If Not db_clin_serv.GET_DEPARTMENTS(i_institution, i_software, l_a_departments) Then

                            Return False

                        Else
                            'Insere-se o clinical service para o 1º departamento do array
                            If Not db_clin_serv.SET_DEP_CLIN_SERV(l_a_clin_serv(i), l_a_departments(0)) Then

                                Return False

                            End If
                        End If
                    Else
                        l_override = True
                    End If
                End If

                'A3 - Preencher Array com valor de dep_clin_serv
                Dim l_a_aux_dep_clin_serv As Int64

                If Not db_clin_serv.GET_DEPARTMENTS(i_institution, i_software, l_a_departments) Then

                    Return False

                End If

                'Vai-se escolher o 1º departamento do array
                If Not db_clin_serv.GET_DEP_CLIN_SERV(i_institution, i_software, l_a_departments(0), l_a_clin_serv(i), l_a_aux_dep_clin_serv) Then

                    Return False

                End If

                If l_override = False Then

                    l_a_dep_clin_serv(i) = l_a_aux_dep_clin_serv

                Else

                    l_a_dep_clin_serv(i) = -1

                End If

            Else
                'Preencher array com -1
                l_a_dep_clin_serv(i) = -1

            End If

        Next

        sql = sql & "               l_id_clinical_services   table_number := table_number("

        For i As Integer = 0 To l_a_dep_clin_serv.Count() - 1

            If (i < l_a_dep_clin_serv.Count() - 1) Then

                sql = sql & " '" & l_a_dep_clin_serv(i) & "', "

            Else

                sql = sql & "'" & l_a_dep_clin_serv(i) & "');"

            End If

        Next

        sql = sql & "               l_type                   table_varchar := table_varchar("

        For i As Integer = 0 To i_destinations.Count() - 1

            If (i < i_destinations.Count() - 1) Then

                sql = sql & " '" & i_destinations(i).type & "', "

            Else

                sql = sql & "'" & i_destinations(i).type & "');"

            End If

        Next

        sql = sql & "               l_id_alert_reason      alert.disch_reas_dest.id_discharge_reason%TYPE;
                                    l_id_alert_destination alert.disch_reas_dest.id_discharge_dest%TYPE;

                                    --Variáveis a inserir no disch_reas_dest
                                    l_flg_diag               alert_default.disch_reas_dest.flg_diag%TYPE;
                                    l_report_name            alert_default.disch_reas_dest.report_name%TYPE;
                                    l_id_edis_type           alert_default.disch_reas_dest.id_epis_type%TYPE;
                                    l_type_screen            alert_default.disch_reas_dest.type_screen%TYPE;
                                    l_id_reports             alert_default.disch_reas_dest.id_reports%TYPE;
                                    l_flg_mcdt               alert_default.disch_reas_dest.flg_mcdt%TYPE;
                                    l_flg_care_stage         alert_default.disch_reas_dest.flg_care_stage%TYPE;
                                    l_flg_default            alert_default.disch_reas_dest.flg_default%TYPE;
                                    l_rank                   alert_default.disch_reas_dest.rank%TYPE;
                                    l_flg_secify_dest        alert_default.disch_reas_dest.flg_specify_dest%TYPE;
                                    l_flg_rep_notes          alert_default.disch_reas_dest.flg_rep_notes%TYPE;
                                    l_flg_def_disch_status   alert_default.disch_reas_dest.flg_def_disch_status%TYPE;
                                    l_id_def_disch_status    alert_default.disch_reas_dest.id_def_disch_status%TYPE;
                                    l_flg_needs_overall_resp alert_default.disch_reas_dest.flg_needs_overall_resp%TYPE;
                                    l_flg_auto_presc_cancel  alert_default.disch_reas_dest.flg_auto_presc_cancel%TYPE;

                                    l_id_dep_clin_serv alert.disch_reas_dest.id_dep_clin_serv%TYPE;

                                    --#############################################################################################
                                    FUNCTION get_disch_reason(i_id_content_reason IN alert.discharge_reason.id_content%TYPE
                              
                                                              ) RETURN alert.discharge_reason.id_discharge_reason%TYPE IS
    
                                        l_id_alert alert.discharge_reason.id_discharge_reason%TYPE;
    
                                    BEGIN
    
                                        SELECT dr.id_discharge_reason
                                        INTO l_id_alert
                                        FROM alert.discharge_reason dr
                                        WHERE dr.id_content = i_id_content_reason
                                        AND dr.flg_available = 'Y';
    
                                        RETURN l_id_alert;
    
                                    END get_disch_reason;

                                    --#############################################################################################

                                    FUNCTION get_disch_destination(i_id_content_destination IN alert.discharge_reason.id_content%TYPE
                                   
                                                                   ) RETURN alert.discharge_dest.id_discharge_dest%TYPE IS
    
                                        l_id_alert alert.discharge_dest.id_discharge_dest%TYPE;
    
                                    BEGIN
    
                                        SELECT d.id_discharge_dest
                                        INTO l_id_alert
                                        FROM alert.discharge_dest d
                                        WHERE d.id_content = i_id_content_destination
                                        AND d.flg_available = 'Y';
    
                                        RETURN l_id_alert;
    
                                    END get_disch_destination;
                                    --#############################################################################################

                                    FUNCTION check_reas_dest
                                    (
                                        i_id_reason      IN alert.discharge_reason.id_discharge_reason%TYPE,
                                        i_id_destination IN alert.discharge_dest.id_discharge_dest%TYPE,
                                        i_id_software    IN alert.disch_reas_dest.id_software_param%TYPE,
                                        i_id_institution IN alert.disch_reas_dest.id_instit_param%TYPE
        
                                    ) RETURN BOOLEAN IS
    
                                        l_count INTEGER := 0;
    
                                    BEGIN
    
                                        SELECT COUNT(*)
                                        INTO l_count
                                        FROM alert.disch_reas_dest d
                                        WHERE d.id_software_param = i_id_software
                                        AND d.id_instit_param = i_id_institution
                                        AND d.id_discharge_reason = i_id_reason
                                        AND d.id_discharge_dest = i_id_destination
                                        AND d.flg_active = 'A';
    
                                        dbms_output.put_line('COUNT: ' || l_count);
                                        dbms_output.put_line(i_id_reason);
                                        dbms_output.put_line(i_id_destination);
    
                                        IF l_count > 0
                                        THEN
                                            RETURN TRUE;
                                        ELSE
                                            RETURN FALSE;
                                        END IF;
    
                                    END check_reas_dest;
                                    --#############################################################################################

                                BEGIN

                                    l_id_alert_reason := get_disch_reason(l_id_content_reason);

                                    FOR i IN 1 .. l_id_content_discharge.count()
                                    LOOP

                                        IF l_type(i) = 'D'
                                        THEN
        
                                            l_id_alert_destination := get_disch_destination(l_id_content_discharge(i));
        
                                        END IF;
        
                                        IF NOT check_reas_dest(l_id_alert_reason, l_id_alert_destination, l_id_software, l_id_institution)
                                        THEN
                
                                            --Obter os dados do default
                                            SELECT d.flg_diag,
                                                   d.report_name,
                                                   d.id_epis_type,
                                                   d.type_screen,
                                                   d.id_reports,
                                                   d.flg_mcdt,
                                                   d.flg_care_stage,
                                                   d.flg_default,
                                                   d.rank,
                                                   d.flg_specify_dest,
                                                   d.flg_rep_notes,
                                                   d.flg_def_disch_status,
                                                   d.id_def_disch_status,
                                                   d.flg_needs_overall_resp,
                                                   d.flg_auto_presc_cancel
                                            INTO l_flg_diag,
                                                 l_report_name,
                                                 l_id_edis_type,
                                                 l_type_screen,
                                                 l_id_reports,
                                                 l_flg_mcdt,
                                                 l_flg_care_stage,
                                                 l_flg_default,
                                                 l_rank,
                                                 l_flg_secify_dest,
                                                 l_flg_rep_notes,
                                                 l_flg_def_disch_status,
                                                 l_id_def_disch_status,
                                                 l_flg_needs_overall_resp,
                                                 l_flg_auto_presc_cancel
                                            FROM alert_default.disch_reas_dest d
                                            WHERE d.id_disch_reas_dest = l_id_def_disch_reas_dest(i);
        
                                            IF l_type(i) = 'D'
                                            THEN
            
                                                IF l_id_clinical_services(i) = -1
                                                THEN
                                                    l_id_dep_clin_serv := NULL;
                                                ELSE
                                                    l_id_dep_clin_serv := l_id_clinical_services(i);
                                                END IF;
            
                                                INSERT INTO alert.disch_reas_dest
                                                    (id_disch_reas_dest,
                                                     id_discharge_reason,
                                                     id_discharge_dest,
                                                     id_dep_clin_serv,
                                                     flg_active,
                                                     flg_diag,
                                                     id_institution,
                                                     id_instit_param,
                                                     id_software_param,
                                                     report_name,
                                                     id_epis_type,
                                                     type_screen,
                                                     id_department,
                                                     id_reports,
                                                     flg_mcdt,
                                                     rank,
                                                     flg_specify_dest,
                                                     flg_care_stage,
                                                     flg_default,
                                                     flg_rep_notes,
                                                     flg_def_disch_status,
                                                     id_def_disch_status,
                                                     flg_needs_overall_resp,
                                                     flg_auto_presc_cancel)
                                                VALUES
                                                    (alert.seq_disch_reas_dest.nextval,
                                                     l_id_alert_reason,
                                                     l_id_alert_destination,
                                                     l_id_dep_clin_serv,
                                                     'A',
                                                     l_flg_diag,
                                                     NULL,
                                                     l_id_institution,
                                                     l_id_software,
                                                     l_report_name,
                                                     l_id_edis_type,
                                                     l_type_screen,
                                                     NULL,
                                                     l_id_reports,
                                                     l_flg_mcdt,
                                                     l_rank,
                                                     l_flg_secify_dest,
                                                     l_flg_care_stage,
                                                     l_flg_default,
                                                     l_flg_rep_notes,
                                                     l_flg_def_disch_status,
                                                     l_id_def_disch_status,
                                                     l_flg_needs_overall_resp,
                                                     l_flg_auto_presc_cancel);
            
                                            ELSE
            
                                                IF l_id_clinical_services(i) = -1
                                                THEN
                                                    l_id_dep_clin_serv := NULL;
                                                ELSE
                                                    l_id_dep_clin_serv := l_id_clinical_services(i);
                                                END IF;
            
                                                INSERT INTO alert.disch_reas_dest
                                                    (id_disch_reas_dest,
                                                     id_discharge_reason,
                                                     id_discharge_dest,
                                                     id_dep_clin_serv,
                                                     flg_active,
                                                     flg_diag,
                                                     id_institution,
                                                     id_instit_param,
                                                     id_software_param,
                                                     report_name,
                                                     id_epis_type,
                                                     type_screen,
                                                     id_department,
                                                     id_reports,
                                                     flg_mcdt,
                                                     rank,
                                                     flg_specify_dest,
                                                     flg_care_stage,
                                                     flg_default,
                                                     flg_rep_notes,
                                                     flg_def_disch_status,
                                                     id_def_disch_status,
                                                     flg_needs_overall_resp,
                                                     flg_auto_presc_cancel)
                                                VALUES
                                                    (alert.seq_disch_reas_dest.nextval,
                                                     l_id_alert_reason,
                                                     NULL,
                                                     l_id_dep_clin_serv,
                                                     'A',
                                                     l_flg_diag,
                                                     NULL,
                                                     l_id_institution,
                                                     l_id_software,
                                                     l_report_name,
                                                     l_id_edis_type,
                                                     l_type_screen,
                                                     NULL,
                                                     l_id_reports,
                                                     l_flg_mcdt,
                                                     l_rank,
                                                     l_flg_secify_dest,
                                                     l_flg_care_stage,
                                                     l_flg_default,
                                                     l_flg_rep_notes,
                                                     l_flg_def_disch_status,
                                                     l_id_def_disch_status,
                                                     l_flg_needs_overall_resp,
                                                     l_flg_auto_presc_cancel);
            
                                            END IF;
        
                                        END IF;
    
                                    END LOOP;

                                END;"

        Dim cmd_insert_reas_dest As New OracleCommand(sql, Connection.conn)

        Try
            cmd_insert_reas_dest.CommandType = CommandType.Text
            cmd_insert_reas_dest.ExecuteNonQuery()
        Catch ex As Exception

            Try
                ''Versões antigas sem flg auto_presc_cancel
                sql = "DECLARE
                                    l_id_software       alert.disch_reas_dest.id_software_param%TYPE := " & i_software & ";
                                    l_id_institution    alert.disch_reas_dest.id_instit_param%TYPE := " & i_institution & ";
                                    l_id_content_reason alert.discharge_reason.id_content%TYPE := '" & i_reason & "';

                                    l_id_def_disch_reas_dest table_number := table_number("

                For i As Integer = 0 To i_destinations.Count() - 1

                    If (i < i_destinations.Count() - 1) Then

                        sql = sql & i_destinations(i).id_disch_reas_dest & ", "

                    Else

                        sql = sql & i_destinations(i).id_disch_reas_dest & ");"

                    End If

                Next

                sql = sql & "  l_id_content_discharge   table_varchar := table_varchar("


                For i As Integer = 0 To i_destinations.Count() - 1

                    If (i < i_destinations.Count() - 1) Then

                        sql = sql & " '" & i_destinations(i).id_content & "', "

                    Else

                        sql = sql & "'" & i_destinations(i).id_content & "');"

                    End If

                Next

                sql = sql & "               l_id_clinical_services   table_number := table_number("

                For i As Integer = 0 To l_a_dep_clin_serv.Count() - 1

                    If (i < l_a_dep_clin_serv.Count() - 1) Then

                        sql = sql & " '" & l_a_dep_clin_serv(i) & "', "

                    Else

                        sql = sql & "'" & l_a_dep_clin_serv(i) & "');"

                    End If

                Next

                sql = sql & "               l_type                   table_varchar := table_varchar("

                For i As Integer = 0 To i_destinations.Count() - 1

                    If (i < i_destinations.Count() - 1) Then

                        sql = sql & " '" & i_destinations(i).type & "', "

                    Else

                        sql = sql & "'" & i_destinations(i).type & "');"

                    End If

                Next

                sql = sql & "       l_id_alert_reason      alert.disch_reas_dest.id_discharge_reason%TYPE;
                                    l_id_alert_destination alert.disch_reas_dest.id_discharge_dest%TYPE;

                                    --Variáveis a inserir no disch_reas_dest
                                    l_flg_diag               alert_default.disch_reas_dest.flg_diag%TYPE;
                                    l_report_name            alert_default.disch_reas_dest.report_name%TYPE;
                                    l_id_edis_type           alert_default.disch_reas_dest.id_epis_type%TYPE;
                                    l_type_screen            alert_default.disch_reas_dest.type_screen%TYPE;
                                    l_id_reports             alert_default.disch_reas_dest.id_reports%TYPE;
                                    l_flg_mcdt               alert_default.disch_reas_dest.flg_mcdt%TYPE;
                                    l_flg_care_stage         alert_default.disch_reas_dest.flg_care_stage%TYPE;
                                    l_flg_default            alert_default.disch_reas_dest.flg_default%TYPE;
                                    l_rank                   alert_default.disch_reas_dest.rank%TYPE;
                                    l_flg_secify_dest        alert_default.disch_reas_dest.flg_specify_dest%TYPE;
                                    l_flg_rep_notes          alert_default.disch_reas_dest.flg_rep_notes%TYPE;
                                    l_flg_def_disch_status   alert_default.disch_reas_dest.flg_def_disch_status%TYPE;
                                    l_id_def_disch_status    alert_default.disch_reas_dest.id_def_disch_status%TYPE;
                                    l_flg_needs_overall_resp alert_default.disch_reas_dest.flg_needs_overall_resp%TYPE;

                                    l_id_dep_clin_serv alert.disch_reas_dest.id_dep_clin_serv%TYPE;

                                    --#############################################################################################
                                    FUNCTION get_disch_reason(i_id_content_reason IN alert.discharge_reason.id_content%TYPE
                              
                                                              ) RETURN alert.discharge_reason.id_discharge_reason%TYPE IS
    
                                        l_id_alert alert.discharge_reason.id_discharge_reason%TYPE;
    
                                    BEGIN
    
                                        SELECT dr.id_discharge_reason
                                        INTO l_id_alert
                                        FROM alert.discharge_reason dr
                                        WHERE dr.id_content = i_id_content_reason
                                        AND dr.flg_available = 'Y';
    
                                        RETURN l_id_alert;
    
                                    END get_disch_reason;

                                    --#############################################################################################

                                    FUNCTION get_disch_destination(i_id_content_destination IN alert.discharge_reason.id_content%TYPE
                                   
                                                                   ) RETURN alert.discharge_dest.id_discharge_dest%TYPE IS
    
                                        l_id_alert alert.discharge_dest.id_discharge_dest%TYPE;
    
                                    BEGIN
    
                                        SELECT d.id_discharge_dest
                                        INTO l_id_alert
                                        FROM alert.discharge_dest d
                                        WHERE d.id_content = i_id_content_destination
                                        AND d.flg_available = 'Y';
    
                                        RETURN l_id_alert;
    
                                    END get_disch_destination;
                                    --#############################################################################################

                                    FUNCTION check_reas_dest
                                    (
                                        i_id_reason      IN alert.discharge_reason.id_discharge_reason%TYPE,
                                        i_id_destination IN alert.discharge_dest.id_discharge_dest%TYPE,
                                        i_id_software    IN alert.disch_reas_dest.id_software_param%TYPE,
                                        i_id_institution IN alert.disch_reas_dest.id_instit_param%TYPE
        
                                    ) RETURN BOOLEAN IS
    
                                        l_count INTEGER := 0;
    
                                    BEGIN
    
                                        SELECT COUNT(*)
                                        INTO l_count
                                        FROM alert.disch_reas_dest d
                                        WHERE d.id_software_param = i_id_software
                                        AND d.id_instit_param = i_id_institution
                                        AND d.id_discharge_reason = i_id_reason
                                        AND d.id_discharge_dest = i_id_destination
                                        AND d.flg_active = 'A';
    
                                        dbms_output.put_line('COUNT: ' || l_count);
                                        dbms_output.put_line(i_id_reason);
                                        dbms_output.put_line(i_id_destination);
    
                                        IF l_count > 0
                                        THEN
                                            RETURN TRUE;
                                        ELSE
                                            RETURN FALSE;
                                        END IF;
    
                                    END check_reas_dest;
                                    --#############################################################################################

                                BEGIN

                                    l_id_alert_reason := get_disch_reason(l_id_content_reason);

                                    FOR i IN 1 .. l_id_content_discharge.count()
                                    LOOP

                                        IF l_type(i) = 'D'
                                        THEN
        
                                            l_id_alert_destination := get_disch_destination(l_id_content_discharge(i));
        
                                        END IF;
        
                                        IF NOT check_reas_dest(l_id_alert_reason, l_id_alert_destination, l_id_software, l_id_institution)
                                        THEN
                
                                            --Obter os dados do default
                                            SELECT d.flg_diag,
                                                   d.report_name,
                                                   d.id_epis_type,
                                                   d.type_screen,
                                                   d.id_reports,
                                                   d.flg_mcdt,
                                                   d.flg_care_stage,
                                                   d.flg_default,
                                                   d.rank,
                                                   d.flg_specify_dest,
                                                   d.flg_rep_notes,
                                                   d.flg_def_disch_status,
                                                   d.id_def_disch_status,
                                                   d.flg_needs_overall_resp
                                            INTO l_flg_diag,
                                                 l_report_name,
                                                 l_id_edis_type,
                                                 l_type_screen,
                                                 l_id_reports,
                                                 l_flg_mcdt,
                                                 l_flg_care_stage,
                                                 l_flg_default,
                                                 l_rank,
                                                 l_flg_secify_dest,
                                                 l_flg_rep_notes,
                                                 l_flg_def_disch_status,
                                                 l_id_def_disch_status,
                                                 l_flg_needs_overall_resp
                                            FROM alert_default.disch_reas_dest d
                                            WHERE d.id_disch_reas_dest = l_id_def_disch_reas_dest(i);
        
                                            IF l_type(i) = 'D'
                                            THEN
            
                                                IF l_id_clinical_services(i) = -1
                                                THEN
                                                    l_id_dep_clin_serv := NULL;
                                                ELSE
                                                    l_id_dep_clin_serv := l_id_clinical_services(i);
                                                END IF;
            
                                                INSERT INTO alert.disch_reas_dest
                                                    (id_disch_reas_dest,
                                                     id_discharge_reason,
                                                     id_discharge_dest,
                                                     id_dep_clin_serv,
                                                     flg_active,
                                                     flg_diag,
                                                     id_institution,
                                                     id_instit_param,
                                                     id_software_param,
                                                     report_name,
                                                     id_epis_type,
                                                     type_screen,
                                                     id_department,
                                                     id_reports,
                                                     flg_mcdt,
                                                     rank,
                                                     flg_specify_dest,
                                                     flg_care_stage,
                                                     flg_default,
                                                     flg_rep_notes,
                                                     flg_def_disch_status,
                                                     id_def_disch_status,
                                                     flg_needs_overall_resp)
                                                VALUES
                                                    (alert.seq_disch_reas_dest.nextval,
                                                     l_id_alert_reason,
                                                     l_id_alert_destination,
                                                     l_id_dep_clin_serv,
                                                     'A',
                                                     l_flg_diag,
                                                     NULL,
                                                     l_id_institution,
                                                     l_id_software,
                                                     l_report_name,
                                                     l_id_edis_type,
                                                     l_type_screen,
                                                     NULL,
                                                     l_id_reports,
                                                     l_flg_mcdt,
                                                     l_rank,
                                                     l_flg_secify_dest,
                                                     l_flg_care_stage,
                                                     l_flg_default,
                                                     l_flg_rep_notes,
                                                     l_flg_def_disch_status,
                                                     l_id_def_disch_status,
                                                     l_flg_needs_overall_resp);
            
                                            ELSE
            
                                                IF l_id_clinical_services(i) = -1
                                                THEN
                                                    l_id_dep_clin_serv := NULL;
                                                ELSE
                                                    l_id_dep_clin_serv := l_id_clinical_services(i);
                                                END IF;
            
                                                INSERT INTO alert.disch_reas_dest
                                                    (id_disch_reas_dest,
                                                     id_discharge_reason,
                                                     id_discharge_dest,
                                                     id_dep_clin_serv,
                                                     flg_active,
                                                     flg_diag,
                                                     id_institution,
                                                     id_instit_param,
                                                     id_software_param,
                                                     report_name,
                                                     id_epis_type,
                                                     type_screen,
                                                     id_department,
                                                     id_reports,
                                                     flg_mcdt,
                                                     rank,
                                                     flg_specify_dest,
                                                     flg_care_stage,
                                                     flg_default,
                                                     flg_rep_notes,
                                                     flg_def_disch_status,
                                                     id_def_disch_status,
                                                     flg_needs_overall_resp)
                                                VALUES
                                                    (alert.seq_disch_reas_dest.nextval,
                                                     l_id_alert_reason,
                                                     NULL,
                                                     l_id_dep_clin_serv,
                                                     'A',
                                                     l_flg_diag,
                                                     NULL,
                                                     l_id_institution,
                                                     l_id_software,
                                                     l_report_name,
                                                     l_id_edis_type,
                                                     l_type_screen,
                                                     NULL,
                                                     l_id_reports,
                                                     l_flg_mcdt,
                                                     l_rank,
                                                     l_flg_secify_dest,
                                                     l_flg_care_stage,
                                                     l_flg_default,
                                                     l_flg_rep_notes,
                                                     l_flg_def_disch_status,
                                                     l_id_def_disch_status,
                                                     l_flg_needs_overall_resp);
            
                                            END IF;
        
                                        END IF;
    
                                    END LOOP;

                                END;"

                cmd_insert_reas_dest = New OracleCommand(sql, Connection.conn)
                cmd_insert_reas_dest.CommandType = CommandType.Text
                cmd_insert_reas_dest.ExecuteNonQuery()

            Catch ex2 As Exception

                cmd_insert_reas_dest.Dispose()
                Return False

            End Try

        End Try

        cmd_insert_reas_dest.Dispose()

        Return True

    End Function

    Function GET_DISCH_INSTR_VERSIONS(ByVal i_institution As Int64, ByVal i_software As Integer, ByRef i_dr As OracleDataReader) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution)

        Dim sql As String = "SELECT DISTINCT dir.version
                                FROM alert_default.disch_instructions di
                                JOIN alert_default.disch_instr_relation dir ON dir.id_disch_instructions = di.id_disch_instructions
                                JOIN alert_default.disch_instructions_group dig ON dig.id_disch_instructions_group = dir.id_disch_instructions_group
                                                                            AND dig.flg_available = 'Y'
                                JOIN institution i ON i.id_market = dir.id_market
                                WHERE di.flg_available = 'Y'
                                AND dir.id_software = " & i_software & "
                                AND i.id_institution = " & i_institution & "
                                AND alert_default.pk_translation_default.get_translation_default(" & l_id_language & ", di.code_disch_instructions) IS NOT NULL
                                AND alert_default.pk_translation_default.get_translation_default(" & l_id_language & ", di.code_disch_instructions_title) IS NOT NULL
                                AND alert_default.pk_translation_default.get_translation_default(" & l_id_language & ", dig.code_disch_instructions_group) IS NOT NULL
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

    Function GET_DEFAULT_INSTR_GROUP(ByVal i_institution As Int64, ByVal i_software As Integer, ByVal i_version As String, ByRef i_dr As OracleDataReader) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution)

        Dim sql As String = "SELECT DISTINCT dig.id_content, alert_default.pk_translation_default.get_translation_default(" & l_id_language & ", dig.code_disch_instructions_group)
                                FROM alert_default.disch_instructions di
                                JOIN alert_default.disch_instr_relation dir ON dir.id_disch_instructions = di.id_disch_instructions
                                JOIN alert_default.disch_instructions_group dig ON dig.id_disch_instructions_group = dir.id_disch_instructions_group
                                                                            AND dig.flg_available = 'Y'
                                JOIN institution i ON i.id_market = dir.id_market
                                WHERE di.flg_available = 'Y'
                                AND dir.id_software = " & i_software & "
                                AND i.id_institution = " & i_institution & "
                                AND dir.version = '" & i_version & "'
                                AND alert_default.pk_translation_default.get_translation_default(" & l_id_language & ", di.code_disch_instructions) IS NOT NULL
                                AND alert_default.pk_translation_default.get_translation_default(" & l_id_language & ", di.code_disch_instructions_title) IS NOT NULL
                                AND alert_default.pk_translation_default.get_translation_default(" & l_id_language & ", dig.code_disch_instructions_group) IS NOT NULL
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

    Function GET_DEFAULT_INSTR_TITLES(ByVal i_institution As Int64, ByVal i_software As Integer, ByVal i_version As String, ByVal i_group As String, ByRef i_dr As OracleDataReader) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution)

        Dim sql As String = "SELECT DISTINCT di.id_content, alert_default.pk_translation_default.get_translation_default(" & l_id_language & ", di.code_disch_instructions_title)
                                FROM alert_default.disch_instructions di
                                JOIN alert_default.disch_instr_relation dir ON dir.id_disch_instructions = di.id_disch_instructions
                                JOIN alert_default.disch_instructions_group dig ON dig.id_disch_instructions_group = dir.id_disch_instructions_group
                                                                            AND dig.flg_available = 'Y'
                                JOIN institution i ON i.id_market = dir.id_market
                                WHERE di.flg_available = 'Y'
                                AND dir.id_software = " & i_software & "
                                AND i.id_institution = " & i_institution & "
                                AND dir.version = '" & i_version & "'
                                AND dig.id_content = '" & i_group & "'
                                AND alert_default.pk_translation_default.get_translation_default(" & l_id_language & ", di.code_disch_instructions) IS NOT NULL
                                AND alert_default.pk_translation_default.get_translation_default(" & l_id_language & ", di.code_disch_instructions_title) IS NOT NULL
                                AND alert_default.pk_translation_default.get_translation_default(" & l_id_language & ", dig.code_disch_instructions_group) IS NOT NULL
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

    Function GET_DEFAULT_INSTR(ByVal i_institution As Int64, ByVal i_instr As String, ByRef o_desc As String) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution)

        Dim sql As String = "SELECT DISTINCT di.id_content, alert_default.pk_translation_default.get_translation_default(" & l_id_language & ", di.code_disch_instructions)
                                FROM alert_default.disch_instructions di
                                WHERE di.flg_available = 'Y'
                                AND di.id_content = '" & i_instr & "'
                                ORDER BY 2 ASC"

        Dim cmd As New OracleCommand(sql, Connection.conn)
        Dim dr As OracleDataReader
        Try
            cmd.CommandType = CommandType.Text
            dr = cmd.ExecuteReader()

            While dr.Read()

                o_desc = dr.Item(1)

            End While

            cmd.Dispose()
            Return True
        Catch ex As Exception
            cmd.Dispose()
            Return False
        End Try

    End Function

    Function GET_ALERT_INSTR(ByVal i_institution As Int64, ByVal i_instr As String, ByRef o_desc As String) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution)

        Dim sql As String = "SELECT DISTINCT di.id_content, PK_TRANSLATION.get_translation(" & l_id_language & ", di.code_disch_instructions)
                                FROM ALERT.disch_instructions di
                                WHERE di.flg_available = 'Y'
                                AND di.id_content = '" & i_instr & "'
                                ORDER BY 2 ASC"

        Dim cmd As New OracleCommand(sql, Connection.conn)
        Dim dr As OracleDataReader
        Try
            cmd.CommandType = CommandType.Text
            dr = cmd.ExecuteReader()

            While dr.Read()

                o_desc = dr.Item(1)

            End While

            cmd.Dispose()
            Return True
        Catch ex As Exception
            cmd.Dispose()
            Return False
        End Try

    End Function

    Function SET_DISCH_INSTRUCTION(ByVal i_institution As Int64, ByVal i_software As Integer, ByVal i_disch_group As String, ByVal i_instructions() As DEFAULT_INSTR) As Boolean

        'A FUNÇÃO SERÁ RESPONSÁVEL POR VERIFICAR SE GRUPO E INSTRUÇÃO JÁ EXISTEM NO ALERT
        'SÓ INSERE SE REGISTO OU TRADUÇÃO NÃO EXISTIREM

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution)

        Dim sql As String = "DECLARE

                                g_id_language alert.language.id_language%TYPE := " & l_id_language & ";

                                g_id_institution institution.id_institution%type := " & i_institution & ";

                                g_id_disch_group alert.disch_instructions_group.id_content%TYPE := '" & i_disch_group & "'; 

                                g_a_disch_instr table_varchar := table_varchar("

        For i As Integer = 0 To i_instructions.Count() - 1

            If (i < i_instructions.Count() - 1) Then

                sql = sql & " '" & i_instructions(i).ID_CONTENT & "', "

            Else

                sql = sql & "'" & i_instructions(i).ID_CONTENT & "');"

            End If

        Next

        sql = sql & "            g_id_software software.id_software%TYPE := " & i_software & ";

                                g_institution institution.id_institution%TYPE := " & i_institution & "; 

                                --#############################################################################################
                                FUNCTION check_instr_group(i_id_content_group IN alert.disch_instructions_group.id_content%TYPE
                               
                                                           ) RETURN BOOLEAN IS
                                    l_count INTEGER;
    
                                BEGIN
    
                                    SELECT COUNT(1)
                                    INTO l_count
                                    FROM alert.disch_instructions_group dg
                                    WHERE dg.id_content = i_id_content_group
                                    AND dg.flg_available = 'Y';
    
                                    IF l_count > 0
                                    THEN
                                        RETURN TRUE;
                                    ELSE
                                        RETURN FALSE;
                                    END IF;
    
                                END check_instr_group;

                                --#############################################################################################
                                FUNCTION check_instr_group_translation
                                (
                                    i_id_lang          IN alert.language.id_language%TYPE,
                                    i_id_content_group IN alert.disch_instructions_group.id_content%TYPE
        
                                ) RETURN BOOLEAN IS
                                    l_count INTEGER;
    
                                BEGIN
    
                                    SELECT COUNT(1)
                                    INTO l_count
                                    FROM alert.disch_instructions_group dg
                                    WHERE dg.id_content = i_id_content_group
                                    AND dg.flg_available = 'Y'
                                    AND pk_translation.get_translation(i_id_lang, dg.code_disch_instructions_group) IS NOT NULL;
    
                                    IF l_count > 0
                                    THEN
                                        RETURN TRUE;
                                    ELSE
                                        RETURN FALSE;
                                    END IF;
    
                                END check_instr_group_translation;

                                --#############################################################################################   
                                FUNCTION get_instr_group(i_id_content_group IN alert.disch_instructions_group.id_content%TYPE)
                                    RETURN alert.disch_instructions_group.id_disch_instructions_group%TYPE IS
    
                                    l_f_disch_group alert.disch_instructions_group.id_disch_instructions_group%TYPE;
    
                                BEGIN
    
                                    SELECT dg.id_disch_instructions_group
                                    INTO l_f_disch_group
                                    FROM alert.disch_instructions_group dg
                                    WHERE dg.id_content = i_id_content_group
                                    AND dg.flg_available = 'Y';
    
                                    RETURN l_f_disch_group;
    
                                END get_instr_group;

                                --############################################################################################# 
                                PROCEDURE set_disch_instr_group(i_id_content_group IN alert.disch_instructions_group.id_content%TYPE) IS
                                BEGIN
                                    INSERT INTO alert.disch_instructions_group
                                        (id_disch_instructions_group, code_disch_instructions_group, flg_available, id_content)
                                    VALUES
                                        (alert.seq_disch_instructions_group.nextval,
                                         'DISCH_INSTRUCTIONS_GROUP.CODE_DISCH_INSTRUCTIONS_GROUP.' || alert.seq_disch_instructions_group.nextval,
                                         'Y',
                                         i_id_content_group);
                                END set_disch_instr_group;

                                --#############################################################################################     
                                PROCEDURE set_disch_instr_group_trans
                                (
                                    i_id_lang          IN alert.language.id_language%TYPE,
                                    i_id_content_group IN alert.disch_instructions_group.id_content%TYPE
                                ) IS
    
                                    l_f_group_desc       translation.desc_lang_6%TYPE;
                                    l_f_code_translation translation.code_translation%TYPE;
    
                                BEGIN
    
                                    SELECT alert_default.pk_translation_default.get_translation_default(i_id_lang, dg.code_disch_instructions_group)
                                    INTO l_f_group_desc
                                    FROM alert_default.disch_instructions_group dg
                                    WHERE dg.flg_available = 'Y'
                                    AND dg.id_content = i_id_content_group;
    
                                    SELECT dg.code_disch_instructions_group
                                    INTO l_f_code_translation
                                    FROM alert.disch_instructions_group dg
                                    WHERE dg.id_content = i_id_content_group
                                    AND dg.flg_available = 'Y';
    
                                    pk_translation.insert_into_translation(i_id_lang, l_f_code_translation, l_f_group_desc);
    
                                END set_disch_instr_group_trans;

                                --#############################################################################################
                                FUNCTION check_disch_instr(i_id_content_disch IN alert.disch_instructions.id_content%TYPE
                               
                                                           ) RETURN BOOLEAN IS
                                    l_count INTEGER;
    
                                BEGIN
    
                                    SELECT COUNT(1)
                                    INTO l_count
                                    FROM alert.disch_instructions di
                                    WHERE di.id_content = i_id_content_disch
                                    AND di.flg_available = 'Y';
    
                                    IF l_count > 0
                                    THEN
                                        RETURN TRUE;
                                    ELSE
                                        RETURN FALSE;
                                    END IF;
    
                                END check_disch_instr;

                                --#############################################################################################    
                                FUNCTION check_disch_instr_trans
                                (
                                    i_id_lang          IN alert.language.id_language%TYPE,
                                    i_id_content_instr IN alert.disch_instructions.id_content%TYPE
        
                                ) RETURN BOOLEAN IS
                                    l_count INTEGER;
    
                                BEGIN
    
                                    SELECT COUNT(1)
                                    INTO l_count
                                    FROM alert.disch_instructions di
                                    WHERE di.id_content = i_id_content_instr
                                    AND di.flg_available = 'Y'
                                    AND pk_translation.get_translation(i_id_lang, di.code_disch_instructions) IS NOT NULL;
    
                                    IF l_count > 0
                                    THEN
                                        RETURN TRUE;
                                    ELSE
                                        RETURN FALSE;
                                    END IF;
    
                                END check_disch_instr_trans;

                                --#############################################################################################    
                                FUNCTION check_disch_instr_title_trans
                                (
                                    i_id_lang          IN alert.language.id_language%TYPE,
                                    i_id_content_instr IN alert.disch_instructions.id_content%TYPE
        
                                ) RETURN BOOLEAN IS
                                    l_count INTEGER;
    
                                BEGIN
    
                                    SELECT COUNT(1)
                                    INTO l_count
                                    FROM alert.disch_instructions di
                                    WHERE di.id_content = i_id_content_instr
                                    AND di.flg_available = 'Y'
                                    AND pk_translation.get_translation(i_id_lang, di.code_disch_instructions_title) IS NOT NULL;
    
                                    IF l_count > 0
                                    THEN
                                        RETURN TRUE;
                                    ELSE
                                        RETURN FALSE;
                                    END IF;
    
                                END check_disch_instr_title_trans;

                                --############################################################################################# 
                                PROCEDURE set_disch_instr(i_id_content_instr IN alert.disch_instructions.id_content%TYPE) IS
                                BEGIN
    
                                    INSERT INTO alert.disch_instructions
                                        (id_disch_instructions, code_disch_instructions, code_disch_instructions_title, flg_available, id_content)
                                    VALUES
                                        (alert.seq_disch_instructions.nextval,
                                         'DISCH_INSTRUCTIONS.CODE_DISCH_INSTRUCTIONS.' || alert.seq_disch_instructions.nextval,
                                         'DISCH_INSTRUCTIONS.CODE_DISCH_INSTRUCTIONS_TITLE.' || alert.seq_disch_instructions.nextval,
                                         'Y',
                                         i_id_content_instr);
    
                                END set_disch_instr;

                                --#############################################################################################     
                                PROCEDURE set_disch_instr_trans
                                (
                                    i_id_lang          IN alert.language.id_language%TYPE,
                                    i_id_content_instr IN alert.disch_instructions.id_content%TYPE
                                ) IS
    
                                    l_f_instr_desc       translation.desc_lang_6%TYPE;
                                    l_f_code_translation translation.code_translation%TYPE;
    
                                BEGIN
    
                                    SELECT alert_default.pk_translation_default.get_translation_default(i_id_lang, di.code_disch_instructions)
                                    INTO l_f_instr_desc
                                    FROM alert_default.disch_instructions di
                                    WHERE di.flg_available = 'Y'
                                    AND di.id_content = i_id_content_instr;
    
                                    SELECT di.code_disch_instructions
                                    INTO l_f_code_translation
                                    FROM alert.disch_instructions di
                                    WHERE di.id_content = i_id_content_instr
                                    AND di.flg_available = 'Y';
    
                                    pk_translation.insert_into_translation(i_id_lang, l_f_code_translation, l_f_instr_desc);
    
                                END set_disch_instr_trans;

                                --#############################################################################################     
                                PROCEDURE set_disch_instr_title_trans
                                (
                                    i_id_lang          IN alert.language.id_language%TYPE,
                                    i_id_content_instr IN alert.disch_instructions.id_content%TYPE
                                ) IS
    
                                    l_f_instr_desc       translation.desc_lang_6%TYPE;
                                    l_f_code_translation translation.code_translation%TYPE;
    
                                BEGIN
    
                                    SELECT alert_default.pk_translation_default.get_translation_default(i_id_lang, di.code_disch_instructions_title)
                                    INTO l_f_instr_desc
                                    FROM alert_default.disch_instructions di
                                    WHERE di.flg_available = 'Y'
                                    AND di.id_content = i_id_content_instr;
    
                                    SELECT di.code_disch_instructions_title
                                    INTO l_f_code_translation
                                    FROM alert.disch_instructions di
                                    WHERE di.id_content = i_id_content_instr
                                    AND di.flg_available = 'Y';
    
                                    pk_translation.insert_into_translation(i_id_lang, l_f_code_translation, l_f_instr_desc);
    
                                END set_disch_instr_title_trans;

                                --#############################################################################################  
                                FUNCTION check_disch_instr_rel
                                (
                                    i_id_content_group IN alert.disch_instructions_group.id_content%TYPE,
                                    i_id_content_instr IN alert.disch_instructions.id_content%TYPE,
                                    i_id_institution   IN alert.disch_instr_relation.id_institution%TYPE,
                                    i_id_software      IN alert.disch_instr_relation.id_software%TYPE
                                ) RETURN BOOLEAN IS
                                    l_count INTEGER;
    
                                BEGIN
    
                                    SELECT COUNT(1)
                                    INTO l_count
                                    FROM alert.disch_instr_relation dr
                                    JOIN alert.disch_instructions di ON di.id_disch_instructions = dr.id_disch_instructions
                                                                 AND di.flg_available = 'Y'
                                    JOIN alert.disch_instructions_group dg ON dg.id_disch_instructions_group = dr.id_disch_instructions_group
                                                                       AND dg.flg_available = 'Y'
                                    WHERE di.id_content = i_id_content_instr
                                    AND dg.id_content = i_id_content_group
                                    AND DR.ID_INSTITUTION=g_id_institution
                                    AND DR.ID_SOFTWARE=g_id_software;
    
                                    IF l_count > 0
                                    THEN
                                        RETURN TRUE;
                                    ELSE
                                        RETURN FALSE;
                                    END IF;
    
                                END check_disch_instr_rel;
                                --############################################################################################# 
                                PROCEDURE set_disch_instr_rel
                                (
                                    i_id_content_group IN alert.disch_instructions_group.id_content%TYPE,
                                    i_id_content_instr IN alert.disch_instructions.id_content%TYPE,
                                    i_id_institution   IN alert.disch_instr_relation.id_institution%TYPE,
                                    i_id_software      IN alert.disch_instr_relation.id_software%TYPE
                                ) IS
    
                                    l_id_disch_group alert.disch_instructions_group.id_disch_instructions_group%TYPE;
                                    l_id_disch_instr alert.disch_instructions.id_disch_instructions%TYPE;
    
                                BEGIN
    
                                    SELECT dg.id_disch_instructions_group
                                    INTO l_id_disch_group
                                    FROM alert.disch_instructions_group dg
                                    WHERE dg.id_content = i_id_content_group
                                    AND dg.flg_available = 'Y';
    
                                    SELECT di.id_disch_instructions
                                    INTO l_id_disch_instr
                                    FROM alert.disch_instructions di
                                    WHERE di.id_content = i_id_content_instr
                                    AND di.flg_available = 'Y';
    
                                    INSERT INTO alert.disch_instr_relation
                                        (id_disch_instr_relation, id_disch_instructions, id_disch_instructions_group, id_institution, id_software)
                                    VALUES
                                        (alert.seq_disch_instr_relation.nextval, l_id_disch_instr, l_id_disch_group, i_id_institution, i_id_software);
    
                                END set_disch_instr_rel;

                                --#############################################################################################     
                            BEGIN

                                --1 - Verificar/Inserir Grupo
                                IF NOT check_instr_group(g_id_disch_group)
                                THEN
    
                                    set_disch_instr_group(g_id_disch_group);
                                    set_disch_instr_group_trans(g_id_language, g_id_disch_group);
    
                                ELSIF NOT check_instr_group_translation(g_id_language, g_id_disch_group)
                                THEN
    
                                    set_disch_instr_group_trans(g_id_language, g_id_disch_group);
    
                                END IF;

                                --2.Instruções
                                FOR i IN 1 .. g_a_disch_instr.count()
                                LOOP
                                    --2.1 - Verificar/Inserir Instruções
                                    IF NOT check_disch_instr(g_a_disch_instr(i))
                                    THEN
        
                                        set_disch_instr(g_a_disch_instr(i));
                                        set_disch_instr_title_trans(g_id_language, g_a_disch_instr(i));
                                        set_disch_instr_trans(g_id_language, g_a_disch_instr(i));
        
                                    ELSIF NOT check_disch_instr_title_trans(g_id_language, g_a_disch_instr(i)) AND NOT check_disch_instr_trans(g_id_language, g_a_disch_instr(i))
                                    THEN
        
                                        set_disch_instr_title_trans(g_id_language, g_a_disch_instr(i));
                                        set_disch_instr_trans(g_id_language, g_a_disch_instr(i));
        
                                    ELSIF NOT check_disch_instr_title_trans(g_id_language, g_a_disch_instr(i))
                                    THEN
        
                                        set_disch_instr_title_trans(g_id_language, g_a_disch_instr(i));
        
                                    ELSIF NOT check_disch_instr_trans(g_id_language, g_a_disch_instr(i))
                                    THEN
        
                                        set_disch_instr_trans(g_id_language, g_a_disch_instr(i));
        
                                    END IF;
    
                                    --2.2 - Verificar/Inserir Relação
    
                                    IF NOT check_disch_instr_rel(g_id_disch_group, g_a_disch_instr(i), g_institution, g_id_software)
                                    THEN
        
                                        set_disch_instr_rel(g_id_disch_group, g_a_disch_instr(i), g_institution, g_id_software);
        
                                    END IF;
    
                                END LOOP;
                            END;"

        Dim cmd_insert_disch_instr As New OracleCommand(sql, Connection.conn)

        Try
            cmd_insert_disch_instr.CommandType = CommandType.Text
            cmd_insert_disch_instr.ExecuteNonQuery()
        Catch ex As Exception
            cmd_insert_disch_instr.Dispose()
            Return False
        End Try

        cmd_insert_disch_instr.Dispose()

        Return True

    End Function

    Function GET_REASONS(ByVal i_institution As Int64, ByVal i_software As Integer, ByRef i_dr As OracleDataReader) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution)

        Dim sql As String = "SELECT DISTINCT dr.id_content, pk_translation.get_translation(" & l_id_language & ", dr.code_discharge_reason)
                                FROM alert.discharge_reason dr
                                JOIN alert.disch_reas_dest drd ON drd.id_discharge_reason = dr.id_discharge_reason
                                JOIN alert.profile_disch_reason pdr ON pdr.id_discharge_reason = dr.id_discharge_reason and pdr.id_institution=drd.id_instit_param
                                JOIN alert.profile_template pt ON pt.id_profile_template = pdr.id_profile_template and pt.id_software=drd.id_software_param
                                WHERE drd.flg_active = 'A'
                                AND dr.flg_available = 'Y'
                                AND drd.id_instit_param = " & i_institution & "
                                AND drd.id_software_param = " & i_software & "
                                AND pdr.flg_available = 'Y'
                                AND pt.flg_available = 'Y'
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

    Function GET_PROFILE_DISCH_REASON(ByVal i_institution As Int64, ByVal i_software As Integer, ByVal i_id_content_reason As String, ByRef i_dr As OracleDataReader) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution)

        Dim sql As String = "SELECT DISTINCT pdr.id_profile_disch_reason, pdr.id_profile_template, pt.intern_name_templ
                                FROM alert.discharge_reason dr
                                JOIN alert.disch_reas_dest drd ON drd.id_discharge_reason = dr.id_discharge_reason
                                JOIN alert.profile_disch_reason pdr ON pdr.id_discharge_reason = dr.id_discharge_reason and pdr.id_institution=drd.id_instit_param
                                JOIN alert.profile_template pt ON pt.id_profile_template = pdr.id_profile_template
                                                           AND pt.id_software = drd.id_software_param
                                WHERE drd.flg_active = 'A'
                                AND dr.flg_available = 'Y'
                                AND drd.id_instit_param = " & i_institution & "
                                AND drd.id_software_param = " & i_software & "
                                AND pdr.flg_available = 'Y'
                                AND pt.flg_available = 'Y'
                                AND dr.id_content = '" & i_id_content_reason & "'
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

    Function GET_DESTINATIONS(ByVal i_institution As Int64, ByVal i_software As Integer, ByVal i_id_content_reason As String, ByRef i_dr As OracleDataReader) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution)

        Dim sql As String = "SELECT DISTINCT drd.id_disch_reas_dest, 
                                             NVL(DD.id_content,DR.id_content),
                                             NVL2(DD.id_content,
                                                  pk_translation.get_translation(" & l_id_language & ", DD.CODE_DISCHARGE_DEST),
                                                  pk_translation.get_translation(" & l_id_language & ", DR.CODE_DISCHARGE_REASON))

                                FROM alert.discharge_reason dr
                                JOIN alert.disch_reas_dest drd ON drd.id_discharge_reason = dr.id_discharge_reason
                                JOIN alert.profile_disch_reason pdr ON pdr.id_discharge_reason = dr.id_discharge_reason and pdr.id_institution=drd.id_instit_param
                                JOIN alert.profile_template pt ON pt.id_profile_template = pdr.id_profile_template  and pt.id_software=drd.id_software_param
                                LEFT JOIN ALERT.DISCHARGE_DEST DD ON DD.ID_DISCHARGE_DEST=DRD.ID_DISCHARGE_DEST AND DD.FLG_AVAILABLE='Y'
                                WHERE drd.flg_active = 'A'
                                AND dr.flg_available = 'Y'
                                AND drd.id_instit_param = " & i_institution & "
                                AND drd.id_software_param = " & i_software & "
                                AND pdr.flg_available = 'Y'
                                AND pt.flg_available = 'Y'                               
                                AND DR.ID_CONTENT='" & i_id_content_reason & "'
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

    Function DELETE_DISCH_REAS_DEST(ByVal i_id_disch_reas_isnt As Int64) As Boolean

        Dim sql As String = "UPDATE ALERT.DISCH_REAS_DEST DRD
                                SET DRD.FLG_ACTIVE='N'
                                WHERE DRD.ID_DISCH_REAS_DEST=" & i_id_disch_reas_isnt

        Dim cmd_delete_disch_reas_inst As New OracleCommand(sql, Connection.conn)

        Try
            cmd_delete_disch_reas_inst.CommandType = CommandType.Text
            cmd_delete_disch_reas_inst.ExecuteNonQuery()
        Catch ex As Exception
            cmd_delete_disch_reas_inst.Dispose()
            Return False
        End Try

        cmd_delete_disch_reas_inst.Dispose()

        Return True

    End Function

    Function DELETE_PROF_DISCH_REAS(ByVal i_id_prof_disch_reas As Int64) As Boolean

        Dim sql As String = "UPDATE ALERT.PROFILE_DISCH_REASON PDR
                                SET PDR.FLG_AVAILABLE='N'
                                WHERE PDR.ID_PROFILE_DISCH_REASON=" & i_id_prof_disch_reas

        Dim cmd_delete_prof_disch_reas As New OracleCommand(sql, Connection.conn)

        Try
            cmd_delete_prof_disch_reas.CommandType = CommandType.Text
            cmd_delete_prof_disch_reas.ExecuteNonQuery()
        Catch ex As Exception
            cmd_delete_prof_disch_reas.Dispose()
            Return False
        End Try

        cmd_delete_prof_disch_reas.Dispose()

        Return True

    End Function

    Function GET_ALERT_INSTR_GROUP(ByVal i_institution As Int64, ByVal i_software As Integer, ByRef i_dr As OracleDataReader) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution)

        Dim sql As String = "SELECT DISTINCT dg.id_content, pk_translation.get_translation(" & l_id_language & ", dg.code_disch_instructions_group)
                                FROM alert.disch_instructions_group dg
                                JOIN alert.disch_instr_relation dir ON dir.id_disch_instructions_group = dg.id_disch_instructions_group
                                JOIN alert.disch_instructions di ON di.id_disch_instructions = dir.id_disch_instructions
                                WHERE dg.flg_available = 'Y'
                                AND di.flg_available = 'Y'
                                AND dir.id_institution = " & i_institution & "
                                AND dir.id_software = " & i_software & "
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

    Function GET_ALERT_INSTR(ByVal i_institution As Int64, ByVal i_software As Integer, ByVal i_id_group As String, ByRef i_dr As OracleDataReader) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution)

        Dim sql As String = "SELECT DISTINCT di.id_content, pk_translation.get_translation(" & l_id_language & ", di.code_disch_instructions_title)
                                FROM alert.disch_instructions_group dg
                                JOIN alert.disch_instr_relation dir ON dir.id_disch_instructions_group = dg.id_disch_instructions_group
                                JOIN alert.disch_instructions di ON di.id_disch_instructions = dir.id_disch_instructions
                                WHERE dg.flg_available = 'Y'
                                AND di.flg_available = 'Y'
                                AND dir.id_institution = " & i_institution & "
                                AND dir.id_software = " & i_software & "
                                and dg.id_content='" & i_id_group & "'
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

    Function DELETE_DISCH_INSTR_REL(ByVal i_institution As Int64, ByVal i_software As Integer, ByVal i_id_group As String, ByVal i_id_instruction As String) As Boolean

        Dim sql As String = "DELETE FROM alert.disch_instr_relation dir
                                WHERE dir.id_disch_instructions_group IN (SELECT DIG.ID_DISCH_INSTRUCTIONS_GROUP
                                                                          FROM alert.disch_instructions_group dig
                                                                          WHERE dig.id_content = '" & i_id_group & "'
                                                                          AND dig.flg_available = 'Y')
                                AND dir.id_disch_instructions IN (SELECT di.id_disch_instructions
                                                                 FROM alert.disch_instructions di
                                                                 WHERE di.id_content = '" & i_id_instruction & "'
                                                                 AND di.flg_available = 'Y')
                                AND dir.id_institution = " & i_institution & "
                                AND dir.id_software = " & i_software

        Dim cmd_delete_disch_instr_rel As New OracleCommand(sql, Connection.conn)

        Try
            cmd_delete_disch_instr_rel.CommandType = CommandType.Text
            cmd_delete_disch_instr_rel.ExecuteNonQuery()
        Catch ex As Exception
            cmd_delete_disch_instr_rel.Dispose()
            Return False
        End Try

        cmd_delete_disch_instr_rel.Dispose()

        Return True

    End Function

    Function GET_ALL_REASON_SCREENS(ByRef o_dr As OracleDataReader) As Boolean

        Dim sql As String = "SELECT DISTINCT dr.file_to_execute
                                FROM alert_default.discharge_reason dr
                                WHERE dr.file_to_execute IS NOT NULL
                                AND UPPER(dr.file_to_execute) LIKE '%SWF'
                                ORDER BY 1 ASC"

        Dim cmd As New OracleCommand(sql, Connection.conn)
        Try
            cmd.CommandType = CommandType.Text
            o_dr = cmd.ExecuteReader()
            cmd.Dispose()
            Return True
        Catch ex As Exception
            cmd.Dispose()
            Return False
        End Try

    End Function

    Function GET_DEFAULT_SCREEN(ByVal i_reason As String, ByRef o_dr As OracleDataReader) As Boolean

        Dim sql As String = "SELECT DISTINCT dr.file_to_execute
                                FROM alert_default.discharge_reason dr
                                WHERE dr.file_to_execute IS NOT NULL
                                and dr.id_content='" & i_reason & "'
                                ORDER BY 1 ASC"

        Dim cmd As New OracleCommand(sql, Connection.conn)
        Try
            cmd.CommandType = CommandType.Text
            o_dr = cmd.ExecuteReader()
            cmd.Dispose()
            Return True
        Catch ex As Exception
            cmd.Dispose()
            Return False
        End Try

    End Function

    'Function GET_SREEN_NAME(ByVal i_reason As String, ByRef o_dr As OracleDataReader) As Boolean

    '    Dim sql As String = "SELECT DISTINCT dr.file_to_execute
    '                            FROM alert_default.discharge_reason dr
    '                            WHERE dr.file_to_execute IS NOT NULL
    '                            and dr.id_content='" & i_reason & "'
    '                            ORDER BY 1 ASC"

    '    Dim cmd As New OracleCommand(sql, Connection.conn)
    '    Try
    '        cmd.CommandType = CommandType.Text
    '        o_dr = cmd.ExecuteReader()
    '        cmd.Dispose()
    '        Return True
    '    Catch ex As Exception
    '        cmd.Dispose()
    '        Return False
    '    End Try

    'End Function

End Class
