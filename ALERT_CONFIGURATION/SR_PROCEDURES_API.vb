Imports Oracle.DataAccess.Client

Public Class SR_PROCEDURES_API

    Dim db_access_general As New General

    Public Structure sr_interventions_default

        Public id_content_intervention As String
        Public desc_intervention As String

    End Structure

    Function GET_DEFAULT_VERSIONS(ByVal i_institution As Int64, ByVal i_conn As OracleConnection, ByRef i_dr As OracleDataReader) As Boolean

        Dim sql As String = ""

        sql = "SELECT DISTINCT dv.version
                        FROM alert_default.sr_intervention di
                        JOIN alert_default.sr_interv_codification dc ON dc.flg_coding = di.flg_coding
                        JOIN alert_default.sr_intervention_mrk_vrs dv ON dv.id_sr_intervention = di.id_sr_intervention
                        join alert_core_data.ab_institution i on i.id_ab_market=dv.id_market
                        WHERE di.flg_status = 'A'
                        AND i.id_ab_institution= " & i_institution & "
                        ORDER BY 1 ASC"

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

    Function GET_DEFAULT_CODIFICATION(ByVal i_institution As Int64, ByVal i_version As String, ByVal i_conn As OracleConnection, ByRef i_dr As OracleDataReader) As Boolean

        Dim sql As String = ""

        sql = "Select distinct dc.flg_coding from alert_default.sr_intervention di
                join alert_default.sr_interv_codification dc on dc.flg_coding=di.flg_coding
                join alert_default.sr_intervention_mrk_vrs dv on dv.id_sr_intervention=di.id_sr_intervention
                join alert_core_data.ab_institution i on i.id_ab_market=dv.id_market
                where di.flg_status='A'
                and i.id_ab_institution=" & i_institution & "
                and dv.version='" & i_version & "'
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

    Function GET_DEFAULT_SR_INTERVENTIONS(ByVal i_institution As Int64, ByVal i_version As String, ByVal i_codification As String, ByVal i_conn As OracleConnection, ByRef i_dr As OracleDataReader) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution, i_conn)
        Dim sql As String = ""

        sql = "SELECT DISTINCT di.id_content,
                                alert_default.pk_translation_default.get_translation_default(" & l_id_language & ", di.code_sr_intervention),
                                di.flg_type,
                                di.icd,
                                di.gender,
                                di.age_min,
                                di.age_max,
                                di.id_system_organ,
                                di.id_speciality
                FROM alert_default.sr_intervention di
                JOIN alert_default.sr_interv_codification dc ON dc.flg_coding = di.flg_coding
                JOIN alert_default.sr_intervention_mrk_vrs dv ON dv.id_sr_intervention = di.id_sr_intervention
                JOIN alert_core_data.ab_institution i ON i.id_ab_market = dv.id_market
                WHERE di.flg_status = 'A'
                AND I.ID_AB_INSTITUTION=" & i_institution & "
                AND dv.version = '" & i_version & "'
                AND dc.flg_coding = '" & i_codification & "'
                ORDER BY 2 ASC"

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

    Function SET_SR_INTERVENTIONS(ByVal i_institution As Int64, ByVal i_coding As String, ByVal i_a_interventions() As sr_interventions_default, ByVal i_conn As OracleConnection) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution, i_conn)
        Dim sql As String = "DECLARE

                                l_a_interventions table_varchar := table_varchar("

        For i As Integer = 0 To i_a_interventions.Count() - 1

            If (i < i_a_interventions.Count() - 1) Then

                sql = sql & "'" & i_a_interventions(i).id_content_intervention & "', "

            Else

                sql = sql & "'" & i_a_interventions(i).id_content_intervention & "');"

            End If

        Next

        sql = sql & "   l_codification    alert.sr_intervention.flg_coding%type := '" & i_coding & "';     
                        l_intervention    alert.sr_intervention.id_sr_intervention%TYPE;
                        
                        l_flg_type        alert.sr_intervention.flg_type%type;
                        l_icd             alert.sr_intervention.icd%type;
                        l_gender          alert.sr_intervention.gender%TYPE;
                        l_age_min         alert.sr_intervention.age_min%TYPE;
                        l_age_max         alert.sr_intervention.age_max%TYPE;
                        l_id_sysetm_organ alert.sr_intervention.id_system_organ%type;

                        l_sequence_interv   alert.sr_intervention.id_sr_intervention%type;

                        l_interv_desc       alert_default.translation.desc_lang_1%type;
                        
                    BEGIN

                        FOR i IN 1 .. l_a_interventions.count()
                        LOOP
                            BEGIN
        
                                SELECT i.id_sr_intervention
                                INTO l_intervention
                                FROM alert.sr_intervention i
                                WHERE i.id_content = l_a_interventions(i)
                                AND i.flg_status = 'A';
        
                            EXCEPTION
                                WHEN no_data_found THEN
                
                                     l_sequence_interv := alert.seq_sr_intervention.nextval;
                
                                    SELECT di.flg_type,di.icd, di.gender, di.age_min, di.age_max, di.id_system_organ, ALERT_DEFAULT.PK_TRANSLATION_DEFAULT.get_translation_default(" & l_id_language & ",DI.Code_Sr_Intervention)
                                    INTO l_flg_type,l_icd,l_gender, l_age_min, l_age_max, l_id_sysetm_organ, l_interv_desc
                                    FROM alert_default.sr_intervention di
                                    WHERE di.id_content = l_a_interventions(i)
                                    AND di.flg_status = 'A';
                
                                   insert into alert.sr_intervention (ID_SR_INTERVENTION,  CODE_SR_INTERVENTION, FLG_STATUS, FLG_TYPE, ICD, GENDER, AGE_MIN, AGE_MAX,  ID_SYSTEM_ORGAN, FLG_CODING, ID_CONTENT)
                                   values (l_sequence_interv, 'SR_INTERVENTION.CODE_SR_INTERVENTION.' || l_sequence_interv, 'A', l_flg_type,  l_icd, l_gender, l_age_min, l_age_max, l_id_sysetm_organ, l_codification,l_a_interventions(i));
                                    
                                    begin
                                               PK_TRANSLATION.insert_into_translation(" & l_id_language & ",'SR_INTERVENTION.CODE_SR_INTERVENTION.'||l_sequence_interv,l_interv_desc);
                                    end;
                
                                    continue;
                            END;
                        END LOOP;

                    END;"

        Dim cmd_insert_interv As New OracleCommand(sql, i_conn)

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

    Function SET_SR_INTERVS_TRANSLATION(ByVal i_institution As Int64, ByVal i_a_interventions() As sr_interventions_default, ByVal i_conn As OracleConnection) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution, i_conn)
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
        
                                SELECT i.code_sr_intervention
                                INTO l_interv_code
                                FROM alert.sr_intervention i
                                WHERE i.id_content = l_a_interventions(i)
                                AND i.flg_status = 'A'
                                AND pk_translation.get_translation(" & l_id_language & ", i.code_sr_intervention) IS NULL;
        
                                IF l_interv_code IS NOT NULL
                                THEN
            
                                    SELECT alert_default.pk_translation_default.get_translation_default(" & l_id_language & ", di.code_sr_intervention)
                                    INTO l_interv_desc
                                    FROM alert_default.sr_intervention di
                                    WHERE di.id_content = l_a_interventions(i)
                                    AND di.flg_status = 'A';

                                    SELECT i.code_sr_intervention
                                    INTO l_interv_code
                                    FROM alert.sr_intervention i
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

        Dim cmd_insert_interv As New OracleCommand(sql, i_conn)

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

    Function SET_SR_INTERV_DEP_CLIN_SERV(ByVal i_institution As Int64, ByVal i_a_interventions() As sr_interventions_default, ByVal i_conn As OracleConnection) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution, i_conn)
        Dim sql As String = "DECLARE

                                l_a_id_content table_varchar := table_varchar("

        For i As Integer = 0 To i_a_interventions.Count() - 1

            If (i < i_a_interventions.Count() - 1) Then

                sql = sql & "'" & i_a_interventions(i).id_content_intervention & "', "

            Else

                sql = sql & "'" & i_a_interventions(i).id_content_intervention & "');"

            End If

        Next

        sql = sql & "  l_id_sr_intervention alert.sr_intervention.id_sr_intervention%TYPE;

                        l_a_dep_clin_serv table_number := table_number();

                    BEGIN

                        FOR i IN 1 .. l_a_id_content.count()
                        LOOP
    
                            BEGIN
        
                                SELECT sr.id_sr_intervention
                                INTO l_id_sr_intervention
                                FROM alert.sr_intervention sr
                                WHERE sr.id_content = l_a_id_content(i)
                                AND sr.flg_status = 'A';
        
                                INSERT INTO alert.sr_interv_dep_clin_serv
                                    (id_sr_interv_dep_clin_serv, id_dep_clin_serv, id_sr_intervention, flg_type, id_professional, id_institution, id_software, rank)
                                VALUES
                                    (alert.seq_sr_interv_dep_clin_serv.nextval, NULL, l_id_sr_intervention, 'M', NULL, " & i_institution & ", 2, 0);
            
                                EXCEPTION WHEN dup_val_on_index THEN continue;
        
                            END;
    
                            SELECT dps.id_dep_clin_serv  
                            BULK COLLECT
                            INTO l_a_dep_clin_serv
                            FROM alert.dep_clin_serv dps
                            WHERE dps.id_clinical_service IN (SELECT c.id_clinical_service
                                                              FROM alert_default.sr_interv_clin_serv dcs
                                                              JOIN alert_default.sr_intervention di ON di.id_sr_intervention = dcs.id_sr_intervention
                                                              JOIN alert_default.clinical_service dc ON dc.id_clinical_service = dcs.id_clinical_service
                                                                                                 AND dc.flg_available = 'Y'
                                                              JOIN alert.clinical_service c ON c.id_content = dc.id_content
                                                                                        AND c.flg_available = 'Y'
                                                              WHERE di.id_content = l_a_id_content(i)
                                                              AND di.flg_status = 'A')
                            AND dps.id_department IN (SELECT d.id_department
                                                     FROM alert.department d
                                                     WHERE d.id_institution = " & i_institution & "
                                                     AND d.id_software = 2
                                                     AND d.flg_available = 'Y')
                            AND dps.flg_available = 'Y';
    
                            IF l_a_dep_clin_serv.count() > 0
                            THEN
        
                                FOR j IN 1 .. l_a_dep_clin_serv.count()
                                LOOP
            
                                    BEGIN
                
                                        INSERT INTO alert.sr_interv_dep_clin_serv
                                            (id_sr_interv_dep_clin_serv,
                                             id_dep_clin_serv,
                                             id_sr_intervention,
                                             flg_type,
                                             id_professional,
                                             id_institution,
                                             id_software,
                                             rank)
                                        VALUES
                                            (alert.seq_sr_interv_dep_clin_serv.nextval, l_a_dep_clin_serv(j), l_id_sr_intervention, 'M', NULL, " & i_institution & ", 2, 0);
                    
                                        EXCEPTION WHEN dup_val_on_index THEN continue;
                
                                    END;
            
                                END LOOP;
        
                            END IF;
    
                        END LOOP;

                    END;"

        Dim cmd_insert_interv As New OracleCommand(sql, i_conn)

        ' Try
        cmd_insert_interv.CommandType = CommandType.Text
        cmd_insert_interv.ExecuteNonQuery()
        'Catch ex As Exception
        'cmd_insert_interv.Dispose()
        'Return False
        'End Try

        cmd_insert_interv.Dispose()
        Return True

    End Function

End Class
