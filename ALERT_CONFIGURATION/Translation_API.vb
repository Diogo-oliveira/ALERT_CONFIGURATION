Imports Oracle.DataAccess.Client

Public Class Translation_API

    Dim db_access_general As New General

    Function CREATE_TEMP_TABLE() As Boolean

        Dim sql As String = "create table output_records
                                (
                                    record_index  number,
                                    updated_records clob,
                                    record_area     varchar2(20)
                                )"

        Dim cmd_create_temp As New OracleCommand(sql, Connection.conn)

        Try
            cmd_create_temp.CommandType = CommandType.Text
            cmd_create_temp.ExecuteNonQuery()
        Catch ex As Exception  'Se bater significa que tabela já existe
            cmd_create_temp.Dispose()
            Return True
        End Try

        cmd_create_temp.Dispose()
        Return True

    End Function

    Function DELETE_TEMP_TABLE() As Boolean

        Dim sql As String = "drop table output_records"

        Dim cmd_drop_temp As New OracleCommand(sql, Connection.conn)

        Try
            cmd_drop_temp.CommandType = CommandType.Text
            cmd_drop_temp.ExecuteNonQuery()
        Catch ex As Exception  'Se bater significa que tabela já não existe
            cmd_drop_temp.Dispose()
            Return True
        End Try

        cmd_drop_temp.Dispose()
        Return True

    End Function

    Function UPDATE_EXAMS(ByVal i_institution As Int64) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution)

        Dim sql As String = "DECLARE

                                l_a_code_translation translation.code_translation%TYPE;

                                l_a_translation translation.desc_lang_6%TYPE;

                                l_d_translation translation.desc_lang_6%TYPE;

                                l_id_content alert.exam.id_content%TYPE;
    
                                l_output     clob := '';

                                contador number;

                                l_index integer := 1;
    
                                l_record_area varchar2(30) := 'EXAMS';
    
                                CURSOR c_exam IS
                                    SELECT e.id_content, e.code_exam
                                    FROM alert.exam e
                                    JOIN translation t ON t.code_translation = e.code_exam;

                                FUNCTION save_output(i_updated_records IN CLOB) RETURN BOOLEAN IS
    
                                BEGIN
        
                                    insert into output_records
                                    values (l_index,i_updated_records,l_record_area);
                                    l_index:=l_index+1;
        
                                    return true;
      
                                EXCEPTION  
                                    when others then          
                                      return false;              
              
                                END save_output;

                            BEGIN

                                contador := 0;
                                OPEN c_exam;

                                LOOP
    
                                    FETCH c_exam
                                        INTO l_id_content, l_a_code_translation;
                                    EXIT WHEN c_exam%NOTFOUND;
    
                                    SELECT t.desc_lang_" & l_id_language & "
                                    INTO l_a_translation
                                    FROM translation t
                                    WHERE t.code_translation = l_a_code_translation;
    
                                    BEGIN
        
                                        SELECT distinct t.desc_lang_" & l_id_language & "
                                        INTO l_d_translation
                                        FROM alert_default.translation t
                                        JOIN alert_default.exam e ON e.code_exam = t.code_translation
                                        WHERE e.id_content = l_id_content
                                        AND t.desc_lang_" & l_id_language & " IS NOT NULL;
        
                                    EXCEPTION
                                        WHEN no_data_found THEN
                                            continue;
                                    END;
    
                                    IF (l_a_translation <> l_d_translation OR (l_a_translation IS NULL AND l_d_translation IS NOT NULL))
                                    THEN
        
                                        l_output:= 'Record ''' || pk_translation.get_translation(6, l_a_code_translation) || ''' has been updated to ''' ;
                                                                       
                                        pk_translation.insert_into_translation(6, l_a_code_translation, l_d_translation);
            
                                        l_output:= l_output || pk_translation.get_translation(6, l_a_code_translation) || '''.';

                                        if not save_output(l_output) then
      
                                             dbms_output.put_line('ERROR');
    
                                        end if;

                                        contador := contador + 1;
        
                                    END IF;
    
                                END LOOP;

                                CLOSE c_exam;
       
                                l_output:= to_char(contador) || ' record(s) updated!';
    
                                if not save_output(l_output) then
      
                                   dbms_output.put_line('ERROR');
    
                                end if;
    
                            END;"


        Dim cmd_update_exams As New OracleCommand(sql, Connection.conn)

        Try
            cmd_update_exams.CommandType = CommandType.Text
            cmd_update_exams.ExecuteNonQuery()
        Catch ex As Exception
            cmd_update_exams.Dispose()
            Return False
        End Try

        cmd_update_exams.Dispose()
        Return True

    End Function

    Function GET_UPDTAED_RECORDS(ByRef i_dr As OracleDataReader) As Boolean

        Dim sql As String = "SELECT desc_record as ""UPDATE LOG""
                            FROM (SELECT r.record_index ""INDEX_RECORD"", r.updated_records ""DESC_RECORD""
                                  FROM alert_config.output_records r
                                  ORDER BY 1 ASC) updated_records"

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

End Class
