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

    'Verificar se tem parent
    Function CHECK_HAS_PARENT(ByVal i_institution As Int64, ByVal i_id_content As String) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution)

        Dim sql As String = "SELECT COUNT(1)
                                FROM alert_default.clinical_service c
                                JOIN alert_default.clinical_serv_rel cp ON cp.id_clinical_service = c.id_clinical_service
                                join alert_default.clinical_service dcp on dcp.id_clinical_service=cp.id_cs_parent
                                join alert_default.translation t on t.code_translation=dcp.code_clinical_service and t.desc_lang_" & l_id_language & " is not null
                                
                                WHERE c.flg_available = 'Y'
                                AND cp.flg_available = 'Y'
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

    ''Devolve os parents todos de um determinado clinical service(Pode devolver mais que um)
    Function GET_PARENTS(ByVal i_institution As Int64, ByVal i_id_content As String, ByRef i_dr As OracleDataReader) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution)
        Dim sql As String = "    SELECT dcsp.id_content
                                    FROM alert_default.clinical_service dcs
                                    JOIN alert_default.clinical_serv_rel dr ON dr.id_clinical_service = dcs.id_clinical_service
                                    JOIN alert_default.clinical_service dcsp ON dcsp.id_clinical_service = dr.id_cs_parent
                                    join alert_default.translation t on t.code_translation=dcsp.code_clinical_service and t.desc_lang_" & l_id_language & " is not null                               
                                    WHERE dcs.id_content = '" & i_id_content & "'
                                    AND dcs.flg_available = 'Y'
                                    AND dr.flg_available = 'Y'
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

    ''inserir no ALERT o Clinical Service
    Function SET_CLIN_SERV(ByVal i_id_content As String) As Boolean

        Dim sql As String = "BEGIN
    
                                  insert into alert.clinical_service (ID_CLINICAL_SERVICE,  CODE_CLINICAL_SERVICE, RANK,  FLG_AVAILABLE, ID_CONTENT)
                                  values (alert.seq_clinical_service.nextval, 'CLINICAL_SERVICE.CODE_CLINICAL_SERVICE.' || alert.seq_clinical_service.nextval, 1, 'Y', '" & i_id_content & "');
                            
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

End Class
