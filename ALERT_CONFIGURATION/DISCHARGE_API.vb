Imports Oracle.DataAccess.Client
Public Class DISCHARGE_API

    Dim db_access_general As New General
    Dim db_clin_serv As New CLINICAL_SERVICE_API

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

    Function GET_DEFAULT_PROFILE_DISCH_REASON(ByVal i_software As Integer, ByVal id_disch_reason As String, ByRef o_profile_templates As OracleDataReader) As Boolean

        Dim sql As String = "SELECT PDR.ID_PROFILE_DISCH_REASON, pdr.id_profile_template,PT.INTERN_NAME_TEMPL
                                FROM alert_default.discharge_reason dr
                                JOIN alert_default.profile_disch_reason pdr ON pdr.id_discharge_reason = dr.id_discharge_reason
                                JOIN alert.profile_template pt ON pt.id_profile_template = pdr.id_profile_template
                                WHERE dr.flg_available = 'Y'
                                AND dr.id_content = '" & id_disch_reason & "'
                                AND pdr.flg_available = 'Y'
                                AND pt.id_software = " & i_software & "
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

    Function SET_DESTINATION(ByVal i_institution As Int64, ByVal i_id_destination() As DEFAULT_DISCAHRGE) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution)

        Dim sql As String = "DECLARE

                                l_id_content table_varchar := table_varchar("

        For i As Integer = 0 To i_id_destination.Count() - 1

            If (i < i_id_destination.Count() - 1) Then

                sql = sql & "'" & i_id_destination(i).id_content & "', "

            Else

                sql = sql & "'" & i_id_destination(i).id_content & "');"

            End If

        Next

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

    Function SET_DESTINATION_TRANSLATION(ByVal i_institution As Int64, ByVal i_id_destination() As DEFAULT_DISCAHRGE) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution)

        Dim sql As String = "DECLARE

                                l_id_content table_varchar := table_varchar("

        For i As Integer = 0 To i_id_destination.Count() - 1

            If (i < i_id_destination.Count() - 1) Then

                sql = sql & "'" & i_id_destination(i).id_content & "', "

            Else

                sql = sql & "'" & i_id_destination(i).id_content & "');"

            End If

        Next

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
                                        WHERE dd.id_content = l_id_content(i)
                                        AND dd.flg_available = 'Y';
        
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

        'A1 - VER SE CLINICAL SERVICE EXISTE NO ALERT E/OU SE TEM TRADUÇÃO
        For i As Integer = 0 To l_a_clin_serv.Count() - 1

            If l_a_clin_serv(i) <> -1 Then

                If Not db_clin_serv.CHECK_CLIN_SERV(l_a_clin_serv(i)) Then

                    If Not db_clin_serv.SET_CLIN_SERV(i_institution, l_a_clin_serv(i)) Then

                        Return False

                    End If

                ElseIf db_clin_serv.CHECK_CLIN_SERV_TRANSLATION(i_institution, l_a_clin_serv(i)) Then

                    If Not db_clin_serv.SET_CLIN_SERV_TRANSLATION(i_institution, l_a_clin_serv(i)) Then

                        Return False

                    End If

                End If

            End If

            ' If db_clin_serv.GET_DEP_CLIN_SERV Then

        Next




        sql = sql & "               l_id_clinical_services   table_varchar := table_varchar(-1, -1, -1);
                                    l_type                   table_varchar := table_varchar('D', 'D', 'D');

                                    l_id_alert_reason      alert.disch_reas_dest.id_discharge_reason%TYPE;
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
    
                                        l_id_alert_destination := get_disch_destination(l_id_content_discharge(i));
    
                                        dbms_output.put_line(l_id_alert_destination);
    
                                        IF NOT check_reas_dest(l_id_alert_reason, l_id_alert_destination, l_id_software, l_id_institution)
                                        THEN
        
                                            dbms_output.put_line('NÃO EXISTE');
        
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


    End Function

End Class
