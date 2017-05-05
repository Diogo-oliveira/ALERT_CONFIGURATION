Imports Oracle.DataAccess.Client
Public Class CLINICAL_SERVICE_API

    Dim db_access_general As New General

    'Verificar se o clinical Service já existe no ALERT
    Function CHECK_CLIN_SERV(ByVal i_id_content As String) As Boolean

        Dim sql As String = "Select COUNT(1)
                                From alert.clinical_service c                                
                                Where c.id_content = '" & i_id_content & "'
                                And c.flg_available = 'Y'"

        Dim cmd As New OracleCommand(sql, Connection.conn)
        cmd.CommandType = CommandType.Text

        Dim dr As OracleDataReader

        dr = cmd.ExecuteReader()
        cmd.Dispose()

        Dim l_total_records As Int64

        While dr.Read()
            l_total_records = dr.Item(0)
        End While

        If l_total_records > 0 Then
            Return True
        Else
            Return False
        End If

    End Function

    'Verificar se tem tradução
    Function CHECK_CLIN_SERV_TRANSLATION(ByVal i_institution As Int64, ByVal i_id_content As String) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution)

        Dim sql As String = "SELECT t.desc_lang_" & l_id_language & "
                                FROM alert.clinical_service c
                                JOIN TRANSLATION T ON T.CODE_TRANSLATION=C.CODE_CLINICAL_SERVICE
                                WHERE c.id_content = '" & i_id_content & "'
                                AND c.flg_available = 'Y'"

        Dim cmd As New OracleCommand(sql, Connection.conn)
        cmd.CommandType = CommandType.Text

        Dim dr As OracleDataReader

        dr = cmd.ExecuteReader()
        cmd.Dispose()

        Try

            While dr.Read()
                If dr.Item(0) = "" Then
                    Return False
                Else
                    Return True
                End If
            End While

        Catch ex As Exception
            Return False
        End Try

    End Function

    'Verificar se tem parent (Só vai verificar na tabela clinical_service. Decidi ignorar a relation)
    Function CHECK_HAS_PARENT(ByVal i_institution As Int64, ByVal i_id_content As String) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution)

        Dim sql As String = "SELECT COUNT(1)
                                FROM alert_default.clinical_service c                                
                                join alert_default.clinical_service dcp on dcp.id_clinical_service=C.ID_CLINICAL_SERVICE_PARENT
                                join alert_default.translation t on t.code_translation=dcp.code_clinical_service and t.desc_lang_" & l_id_language & " is not null
                                
                                WHERE c.flg_available = 'Y'
                                AND dcp.flg_available = 'Y'
                                AND c.id_content = '" & i_id_content & "'"

        Dim cmd As New OracleCommand(sql, Connection.conn)
        cmd.CommandType = CommandType.Text

        Dim dr As OracleDataReader

        dr = cmd.ExecuteReader()
        cmd.Dispose()

        Dim l_total_records As Int64

        While dr.Read()
            l_total_records = dr.Item(0)
        End While

        If l_total_records > 0 Then
            Return True
        Else
            Return False
        End If

    End Function

    ''Devolve o parent
    Function GET_PARENT(ByVal i_institution As Int64, ByVal i_id_content As String, ByRef i_dr As OracleDataReader) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution)

        Dim sql As String = "SELECT dcsp.id_content
                                    FROM alert_default.clinical_service dcs                                    
                                    JOIN alert_default.clinical_service dcsp ON dcsp.id_clinical_service = dcs.ID_CLINICAL_SERVICE_PARENT
                                    join alert_default.translation t on t.code_translation=dcsp.code_clinical_service and t.desc_lang_" & l_id_language & " is not null                               
                                    WHERE dcs.id_content = '" & i_id_content & "'
                                    AND dcs.flg_available = 'Y'
                                    AND dcsp.flg_available = 'Y'"

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

    ''Devolve o id_alert de um clinical service
    Function GET_ID_ALERT(ByVal i_id_content As String, ByRef o_id_alert As Int64) As Boolean

        Dim sql As String = "SELECT c.id_clinical_service
                                FROM alert.clinical_service c
                                WHERE c.id_content = '" & i_id_content & "'
                                AND c.flg_available = 'Y'"

        Dim cmd As New OracleCommand(sql, Connection.conn)
        cmd.CommandType = CommandType.Text

        Dim dr As OracleDataReader

        Try

            dr = cmd.ExecuteReader()
            cmd.Dispose()

            While dr.Read()
                o_id_alert = dr.Item(0)
            End While

            Return True

        Catch ex As Exception

            cmd.Dispose()
            dr.Dispose()
            Return False

        End Try

    End Function

    Function GET_CLIN_SERV_TRANSLATION(ByVal i_institution As Int64, ByVal i_clin_serv As String) As String

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution)

        Dim sql As String = "Select alert_default.pk_translation_default.get_translation_default(" & l_id_language & ",c.code_clinical_service) 
                                from alert_default.clinical_service c
                                where c.flg_available='Y'
                                and c.id_content='" & i_clin_serv & "'"

        Dim cmd As New OracleCommand(sql, Connection.conn)
        cmd.CommandType = CommandType.Text

        Dim dr As OracleDataReader
        Dim l_translation As String = ""

        dr = cmd.ExecuteReader()
        cmd.Dispose()

        While dr.Read()
            l_translation = dr.Item(0)
        End While

        Return l_translation

    End Function

    ''inserir no ALERT o Clinical Service
    Function SET_CLIN_SERV(ByVal i_id_institution As Int64, ByVal i_id_content As String) As Boolean

        Dim l_has_parent As Boolean = False
        Dim l_id_parent As Int64

        If Not CHECK_CLIN_SERV(i_id_content) Then

            '1 - Verificar se o Clinical Service tem Parent
            If CHECK_HAS_PARENT(i_id_institution, i_id_content) Then

                l_has_parent = True

                Dim dr_parent As OracleDataReader

                If Not GET_PARENT(i_id_institution, i_id_content, dr_parent) Then

                    Return False

                End If

                While dr_parent.Read()

                    If Not CHECK_CLIN_SERV(dr_parent.Item(0)) Then

                        If Not SET_CLIN_SERV(i_id_institution, dr_parent.Item(0)) Then

                            Return False

                        End If

                    End If

                    If Not GET_ID_ALERT(dr_parent.Item(0), l_id_parent) Then

                        Return False

                    End If

                End While

                dr_parent.Dispose()
                dr_parent.Close()

            End If

            Dim sql As String

            If l_has_parent = False Then

                sql = "BEGIN
    
                                  insert into alert.clinical_service (ID_CLINICAL_SERVICE,  CODE_CLINICAL_SERVICE, RANK,  FLG_AVAILABLE, ID_CONTENT)
                                  values (alert.seq_clinical_service.nextval, 'CLINICAL_SERVICE.CODE_CLINICAL_SERVICE.' || alert.seq_clinical_service.nextval, 1, 'Y', '" & i_id_content & "');
                            
                   END;"
            Else

                sql = "BEGIN
    
                                  insert into alert.clinical_service (ID_CLINICAL_SERVICE, ID_CLINICAL_SERVICE_PARENT, CODE_CLINICAL_SERVICE, RANK,  FLG_AVAILABLE, ID_CONTENT)
                                  values (alert.seq_clinical_service.nextval, " & l_id_parent & ", 'CLINICAL_SERVICE.CODE_CLINICAL_SERVICE.' || alert.seq_clinical_service.nextval, 1, 'Y', '" & i_id_content & "');
                            
                   END;"
            End If

            Dim cmd_insert_clin_serv As New OracleCommand(sql, Connection.conn)

            Try
                cmd_insert_clin_serv.CommandType = CommandType.Text
                cmd_insert_clin_serv.ExecuteNonQuery()
            Catch ex As Exception
                cmd_insert_clin_serv.Dispose()
                Return False
            End Try

            cmd_insert_clin_serv.Dispose()

            If Not CHECK_CLIN_SERV_TRANSLATION(i_id_institution, i_id_content) Then

                If Not SET_CLIN_SERV_TRANSLATION(i_id_institution, i_id_content) Then

                    Return False

                End If

            End If

        End If

        Return True

    End Function

    ''Inserir Tradução
    Function SET_CLIN_SERV_TRANSLATION(ByVal i_institution As Int64, ByVal i_id_content As String) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution)
        Dim sql As String = "DECLARE

                                l_clin_serv alert.clinical_service.id_content%TYPE := '" & i_id_content & "';

                                l_def_desc translation.desc_lang_6%type;
    
                                l_code_cs translation.code_translation%type;

                            BEGIN

                                SELECT T.DESC_LANG_" & l_id_language & "
                                INTO l_def_desc
                                FROM ALERT_DEFAULT.CLINICAL_SERVICE C
                                JOIN ALERT_DEFAULT.TRANSLATION T ON T.CODE_TRANSLATION=C.CODE_CLINICAL_SERVICE
                                WHERE C.ID_CONTENT=l_clin_serv
                                AND C.FLG_AVAILABLE='Y';
    
                                SELECT C.CODE_CLINICAL_SERVICE
                                INTO l_code_cs
                                FROM ALERT.CLINICAL_SERVICE C
                                WHERE C.ID_CONTENT=l_clin_serv
                                AND C.FLG_AVAILABLE='Y';
    
                                PK_TRANSLATION.insert_into_translation(" & l_id_language & ",l_code_cs,l_def_desc);

                            END;"

        Dim cmd_insert_clin_serv As New OracleCommand(sql, Connection.conn)

        Try
            cmd_insert_clin_serv.CommandType = CommandType.Text
            cmd_insert_clin_serv.ExecuteNonQuery()
        Catch ex As Exception
            cmd_insert_clin_serv.Dispose()
            Return False
        End Try

        cmd_insert_clin_serv.Dispose()

        Return True

    End Function

    Function SET_PARENT(ByVal i_id_content As String, ByVal i_id_parent As Int64) As Boolean

        Dim sql As String = "BEGIN
    
                                  UPDATE alert.clinical_service C
                                  SET c.id_clinical_service_parent=" & i_id_parent & "
                                  WHERE  C.ID_CONTENT='" & i_id_content & "'
                                  AND C.FLG_AVAILABLE='Y';                                  
                            
                             END;"

        Dim cmd_update_clin_serv As New OracleCommand(sql, Connection.conn)

        Try
            cmd_update_clin_serv.CommandType = CommandType.Text
            cmd_update_clin_serv.ExecuteNonQuery()
        Catch ex As Exception
            cmd_update_clin_serv.Dispose()
            Return False
        End Try

        cmd_update_clin_serv.Dispose()

        Return True

    End Function

    'Função para verificar se existe um depl_clin_serv para a inst/soft/clin_Serv
    Function CHECK_DEP_CLIN_SERV(ByVal i_institution As Int64, ByVal i_software As Int16, ByVal i_clinical_service As String) As Boolean

        Dim sql As String = "SELECT count(*)
                                FROM alert.dep_clin_serv dps
                                JOIN alert.department d ON d.id_department = dps.id_department
                                                    AND d.flg_available = 'Y'
                                JOIN alert.clinical_service c ON c.id_clinical_service = dps.id_clinical_service
                                                          AND c.flg_available = 'Y'
                                WHERE dps.flg_available = 'Y'
                                AND d.id_software = " & i_software & "
                                AND d.id_institution = " & i_institution & "
                                AND c.id_content = '" & i_clinical_service & "'"

        Dim cmd As New OracleCommand(sql, Connection.conn)
        cmd.CommandType = CommandType.Text

        Dim dr As OracleDataReader

        dr = cmd.ExecuteReader()
        cmd.Dispose()

        Dim l_total_records As Int64

        While dr.Read()
            l_total_records = dr.Item(0)
        End While

        If l_total_records > 0 Then
            Return True
        Else
            Return False
        End If

    End Function

    'Pode devolver vários resultados?
    Function GET_DEP_CLIN_SERV(ByVal i_institution As Int64, ByVal i_software As Int16, ByVal i_id_department As Int64, ByVal id_clinical_service As String, ByRef o_id_dep_clin_serv As Int64) As Boolean

        Dim sql As String = "SELECT s.id_dep_clin_serv
                                FROM alert.dep_clin_serv s
                                JOIN alert.department d ON d.id_department = s.id_department
                                JOIN alert.clinical_service c ON c.id_clinical_service = s.id_clinical_service
                                WHERE s.flg_available = 'Y'
                                AND d.flg_available = 'Y'
                                AND d.id_institution = " & i_institution & "
                                AND d.id_software = " & i_software & "
                                AND c.flg_available = 'Y'
                                and s.id_department=" & i_id_department & "
                                AND c.id_content = '" & id_clinical_service & "'"


        Dim cmd As New OracleCommand(sql, Connection.conn)
        cmd.CommandType = CommandType.Text

        Dim dr As OracleDataReader

        Try

            dr = cmd.ExecuteReader()
            cmd.Dispose()

            While dr.Read()
                o_id_dep_clin_serv = dr.Item(0)
            End While

            Return True

        Catch ex As Exception

            cmd.Dispose()
            dr.Dispose()
            Return False

        End Try

    End Function

    ''Funcçaõ para devolver todos os departamentos de uma instituição e software
    Function GET_DEPARTMENTS(ByVal i_institution As Int64, ByVal i_software As Int16, ByRef o_a_departments As Int64()) As Boolean

        Dim sql As String = "SELECT d.id_department
                                FROM alert.department d
                                WHERE d.flg_available = 'Y'
                                AND d.id_institution = " & i_institution & "
                                AND d.id_software = " & i_software

        Dim l_array_dimension = 0
        ReDim o_a_departments(l_array_dimension)

        Dim cmd As New OracleCommand(sql, Connection.conn)
        cmd.CommandType = CommandType.Text

        Dim dr As OracleDataReader

        Try

            dr = cmd.ExecuteReader()
            cmd.Dispose()

            While dr.Read()

                ReDim Preserve o_a_departments(l_array_dimension)
                o_a_departments(l_array_dimension) = dr.Item(0)
                l_array_dimension = l_array_dimension + 1

            End While

            Return True

        Catch ex As Exception

            cmd.Dispose()
            dr.Dispose()
            Return False

        End Try

    End Function

    Function SET_DEP_CLIN_SERV(ByVal i_id_clinical_service As String, ByVal i_id_department As Int64) As Boolean

        Dim sql As String = "DECLARE

                                l_id_content_clin_serv alert.clinical_service.id_content%TYPE := '" & i_id_clinical_service & "';

                                l_id_clin_serv alert.clinical_service.id_clinical_service%TYPE;

                                l_id_department alert.department.id_department%TYPE;

                            BEGIN

                                SELECT c.id_clinical_service
                                INTO l_id_clin_serv
                                FROM alert.clinical_service c
                                WHERE c.id_content = l_id_content_clin_serv
                                AND c.flg_available = 'Y';

                                INSERT INTO alert.dep_clin_serv
                                    (id_dep_clin_serv, id_clinical_service, id_department, rank, flg_default, flg_available, flg_coding, flg_just_post_presc, flg_appointment)
                                VALUES
                                    (alert.seq_dep_clin_serv.nextval, l_id_clin_serv, " & i_id_department & ", 0, 'N', 'Y', 'N', 'Y', 'Y');

                            END;"

        Dim cmd_insert_dep_clin_serv As New OracleCommand(sql, Connection.conn)

        Try
            cmd_insert_dep_clin_serv.CommandType = CommandType.Text
            cmd_insert_dep_clin_serv.ExecuteNonQuery()
        Catch ex As Exception
            cmd_insert_dep_clin_serv.Dispose()
            Return False
        End Try

        cmd_insert_dep_clin_serv.Dispose()

        Return True

    End Function

    Function GET_CLIN_SERV_DESC(ByVal i_institution As Int64, ByVal i_id_content_clin_serv As String) As String

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution)

        Dim sql As String = "SELECT t.desc_lang_" & l_id_language & "
                                FROM alert_default.clinical_service c
                                JOIN alert_default.translation t ON t.code_translation = c.code_clinical_service
                                WHERE c.id_content = '" & i_id_content_clin_serv & "'
                                AND c.flg_available = 'Y'"

        Dim cmd As New OracleCommand(sql, Connection.conn)
        cmd.CommandType = CommandType.Text

        Dim dr As OracleDataReader
        Dim cs_translation As String

        dr = cmd.ExecuteReader()
            cmd.Dispose()

            While dr.Read()

            cs_translation = dr.Item(0)

        End While

        Return cs_translation

    End Function


End Class
