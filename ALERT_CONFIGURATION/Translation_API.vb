Imports Oracle.DataAccess.Client

Public Class Translation_API

    Dim db_access_general As New General

    Function CREATE_TEMP_TABLE() As Boolean

        Dim sql As String = "create table output_records
                                (
                                    record_index  number,
                                    updated_records clob,
                                    record_area     varchar2(50)
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
    
                                l_record_area varchar2(50) := 'EXAMS';
    
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

                                   --COLOCAR NO LOG A ÁREA QUE ESTÁ A SER ATUALIZADA
                                   if not save_output(to_char(l_record_area)) then
        
                                               dbms_output.put_line('ERROR');
      
                                   end if;  

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
        
                                        l_output:= 'Record ''' || pk_translation.get_translation(" & l_id_language & ", l_a_code_translation) || ''' has been updated to ''' ;
                                                                       
                                        pk_translation.insert_into_translation(" & l_id_language & ", l_a_code_translation, l_d_translation);
            
                                        l_output:= l_output || pk_translation.get_translation(" & l_id_language & ", l_a_code_translation) || '''  - ' || l_id_content || '.';

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

    Function UPDATE_EXAM_CAT(ByVal i_institution As Int64) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution)

        Dim sql As String = "DECLARE

                                      l_a_code_translation translation.code_translation%type;
      
                                      l_a_translation      translation.desc_lang_6%type;
      
                                      l_d_translation      translation.desc_lang_6%type;
      
                                      l_id_content         alert.diet.id_content%type;
      
                                      l_output     clob := '';
      
                                      contador             integer;
      
                                      l_index integer := 1;
      
                                      l_record_area varchar2(50) := 'EXAM_CATEGORIES';
      
                                      CURSOR c_EXAM_CAT is
                                      select ec.id_content, ec.code_exam_cat
                                      from alert.exam_cat ec
                                      join translation t on t.code_translation=ec.code_exam_cat;

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
       
                                       contador:=0;
                                       OPEN c_EXAM_CAT;

                                       --COLOCAR NO LOG A ÁREA QUE ESTÁ A SER ATUALIZADA
                                       if not save_output(to_char(l_record_area)) then
        
                                                   dbms_output.put_line('ERROR');
      
                                       end if;  
       
                                       LOOP
         
                                            FETCH c_EXAM_CAT into l_id_content,l_a_code_translation;
                                            EXIT WHEN c_EXAM_CAT%notfound;
            
                                            select t.desc_lang_" & l_id_language & "
                                            into  l_a_translation
                                            from translation t 
                                            where t.code_translation=l_a_code_translation;
            
                                            BEGIN
            
                                                select distinct t.desc_lang_" & l_id_language & "
                                                into  l_d_translation
                                                from alert_default.translation t
                                                join alert_default.exam_cat ec on ec.code_exam_cat=t.code_translation
                                                where ec.id_content=l_id_content
                                                and t.desc_lang_" & l_id_language & " is not null;
           
                                           EXCEPTION
                                                WHEN no_data_found then
                                                 CONTINUE;
                                           END;
            
                                            if (l_a_translation<>l_d_translation or (l_a_translation is null and l_d_translation is not null)) THEN
                                                  
                                                    l_output:= 'Record ''' || pk_translation.get_translation(" & l_id_language & ", l_a_code_translation) || ''' has been updated to ''' ;
                                                                       
                                                    pk_translation.insert_into_translation(" & l_id_language & ", l_a_code_translation, l_d_translation);
            
                                                    l_output:= l_output || pk_translation.get_translation(" & l_id_language & ", l_a_code_translation) || '''  - ' || l_id_content || '.';

                                                    if not save_output(l_output) then
      
                                                         dbms_output.put_line('ERROR');
    
                                                    end if;

                                                    contador := contador + 1;
            
                                            END IF;
       
                                       END LOOP;
       
                                       close c_EXAM_CAT;
       
                                       l_output:= to_char(contador) || ' record(s) updated!';
    
                                       if not save_output(l_output) then
      
                                             dbms_output.put_line('ERROR');
    
                                       end if;
             
                                END;"

        Dim cmd_update_exam_cat As New OracleCommand(sql, Connection.conn)

        Try
            cmd_update_exam_cat.CommandType = CommandType.Text
            cmd_update_exam_cat.ExecuteNonQuery()
        Catch ex As Exception
            cmd_update_exam_cat.Dispose()
            Return False
        End Try

        cmd_update_exam_cat.Dispose()
        Return True

    End Function

    Function UPDATE_INTERVENTIONS(ByVal i_institution As Int64) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution)

        Dim sql As String = "DECLARE

                                      l_a_code_translation translation.code_translation%type;
      
                                      l_a_translation      translation.desc_lang_6%type;
      
                                      l_d_translation      translation.desc_lang_6%type;
      
                                      l_id_content         alert.intervention.id_content%type;
      
                                      l_output     clob := '';
      
                                      contador             integer;
      
                                      l_index integer := 1;
      
                                      l_record_area varchar2(50) := 'INTERVENTIONS';
      
                                      CURSOR c_INTERVENTION is
                                      select i.id_content, i.code_intervention
                                      from alert.intervention i
                                      join translation t on t.code_translation=i.code_intervention;
      
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
       
                                       contador:=0;
                                       OPEN c_INTERVENTION;

                                       --COLOCAR NO LOG A ÁREA QUE ESTÁ A SER ATUALIZADA
                                       if not save_output(to_char(l_record_area)) then
        
                                                   dbms_output.put_line('ERROR');
      
                                       end if;  
       
                                       LOOP
         
                                            FETCH c_INTERVENTION into l_id_content,l_a_code_translation;
                                            EXIT WHEN c_INTERVENTION%notfound;
            
                                            select t.desc_lang_" & l_id_language & "
                                            into  l_a_translation
                                            from translation t 
                                            where t.code_translation=l_a_code_translation;
            
            
                                            BEGIN
            
                                                select distinct t.desc_lang_" & l_id_language & "
                                                into  l_d_translation
                                                from alert_default.translation t
                                                join alert_default.intervention i on i.code_intervention=t.code_translation
                                                where i.id_content=l_id_content
                                                and t.desc_lang_" & l_id_language & " is not null;

                                           EXCEPTION
                                                WHEN no_data_found then
                                                 CONTINUE;
                                           END;                
            
                                            if (l_a_translation<>l_d_translation or (l_a_translation is null and l_d_translation is not null)) THEN
                                                  
                                                    l_output:= 'Record ''' || pk_translation.get_translation(" & l_id_language & ", l_a_code_translation) || ''' has been updated to ''' ;
                                                                       
                                                    pk_translation.insert_into_translation(" & l_id_language & ", l_a_code_translation, l_d_translation);
            
                                                    l_output:= l_output || pk_translation.get_translation(" & l_id_language & ", l_a_code_translation) || '''  - ' || l_id_content || '.';

                                                    if not save_output(l_output) then
      
                                                         dbms_output.put_line('ERROR');
    
                                                    end if;

                                                    contador := contador + 1;
            
                                            END IF;
       
                                       END LOOP;
       
                                       close c_INTERVENTION;
       
                                       l_output:= to_char(contador) || ' record(s) updated!';
    
                                       if not save_output(l_output) then
      
                                             dbms_output.put_line('ERROR');
    
                                       end if;
             
                                END;"

        Dim cmd_update_interventions As New OracleCommand(sql, Connection.conn)

        Try
            cmd_update_interventions.CommandType = CommandType.Text
            cmd_update_interventions.ExecuteNonQuery()
        Catch ex As Exception
            cmd_update_interventions.Dispose()
            Return False
        End Try

        cmd_update_interventions.Dispose()
        Return True

    End Function

    Function UPDATE_ANALYSIS(ByVal i_institution As Int64) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution)

        Dim sql As String = "DECLARE

                                  l_a_code_translation translation.code_translation%type;

                                  l_a_translation translation.desc_lang_6%type;

                                  l_d_translation translation.desc_lang_6%type;

                                  l_id_content alert.intervention.id_content%type;

                                  l_output     clob := '';

                                  contador integer;
      
                                  l_index integer := 1;
      
                                  l_record_area varchar2(50) := 'ANALYSIS';

                                  CURSOR c_ANALYSIS is
                                    select a.id_content, a.code_analysis
                                      from alert.analysis a
                                      join translation t
                                        on t.code_translation = a.code_analysis;
            
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
                                  OPEN c_ANALYSIS;

                                   --COLOCAR NO LOG A ÁREA QUE ESTÁ A SER ATUALIZADA
                                   if not save_output(to_char(l_record_area)) then
        
                                               dbms_output.put_line('ERROR');
      
                                   end if;  

                                  LOOP
      
                                    FETCH c_ANALYSIS
                                      into l_id_content, l_a_code_translation;
                                    EXIT WHEN c_ANALYSIS%notfound;
      
                                    select t.desc_lang_" & l_id_language & "
                                      into l_a_translation
                                      from translation t
                                     where t.code_translation = l_a_code_translation;
      
                                    BEGIN
        
                                      select distinct t.desc_lang_" & l_id_language & "
                                        into l_d_translation
                                        from alert_default.translation t
                                        join alert_default.analysis a
                                          on a.code_analysis = t.code_translation
                                       where a.id_content = l_id_content
                                         and t.desc_lang_" & l_id_language & " is not null;
             
                                    EXCEPTION
                                      WHEN no_data_found then
                                        CONTINUE;
                                    END;
      
                                    if (l_a_translation<>l_d_translation or (l_a_translation is null and l_d_translation is not null)) THEN
        
                                              l_output:= 'Record ''' || pk_translation.get_translation(" & l_id_language & ", l_a_code_translation) || ''' has been updated to ''' ;
                                                                       
                                              pk_translation.insert_into_translation(" & l_id_language & ", l_a_code_translation, l_d_translation);
            
                                              l_output:= l_output || pk_translation.get_translation(" & l_id_language & ", l_a_code_translation) || '''  - ' || l_id_content || '.';

                                              if not save_output(l_output) then
      
                                                   dbms_output.put_line('ERROR');
    
                                              end if;

                                              contador := contador + 1;
        
                                    END IF;
      
                                  END LOOP;

                                  close c_ANALYSIS;

                                   l_output:= to_char(contador) || ' record(s) updated!';
    
                                   if not save_output(l_output) then
      
                                         dbms_output.put_line('ERROR');
    
                                   end if;

                            END;"

        Dim cmd_update_analysis As New OracleCommand(sql, Connection.conn)

        Try
            cmd_update_analysis.CommandType = CommandType.Text
            cmd_update_analysis.ExecuteNonQuery()
        Catch ex As Exception
            cmd_update_analysis.Dispose()
            Return False
        End Try

        cmd_update_analysis.Dispose()
        Return True

    End Function

    Function UPDATE_SAMPLE_TYPE(ByVal i_institution As Int64) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution)

        Dim sql As String = "DECLARE

                                  l_a_code_translation translation.code_translation%type;
      
                                  l_a_translation      translation.desc_lang_6%type;
      
                                  l_d_translation      translation.desc_lang_6%type;
      
                                  l_id_content         alert.intervention.id_content%type;
      
                                  l_output     clob := '';      
      
                                  contador             integer;

                                  l_index integer := 1;
      
                                  l_record_area varchar2(50) := 'SAMPLE_TYPE';      
      
                                  CURSOR c_SAMPLE_TYPE is
                                  select st.id_content, st.code_sample_type
                                  from alert.sample_type st
                                  join translation t on t.code_translation=st.code_sample_type
                                  where st.flg_available='Y';
      
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
       
                                   contador:=0;
                                   OPEN c_SAMPLE_TYPE;

                                   --COLOCAR NO LOG A ÁREA QUE ESTÁ A SER ATUALIZADA
                                   if not save_output(to_char(l_record_area)) then
        
                                               dbms_output.put_line('ERROR');
      
                                   end if;  
       
                                   LOOP
         
                                      FETCH c_SAMPLE_TYPE into l_id_content,l_a_code_translation;
                                      EXIT WHEN c_SAMPLE_TYPE%notfound;
          
                                    select t.desc_lang_" & l_id_language & "
                                      into l_a_translation
                                      from translation t
                                     where t.code_translation = l_a_code_translation;
                       
                                      BEGIN
            
                                        select distinct t.desc_lang_" & l_id_language & "
                                          into l_d_translation
                                          from alert_default.translation t
                                          join alert_default.sample_type st
                                            on st.code_sample_type = t.code_translation
                                         where st.id_content = l_id_content
                                         and t.desc_lang_" & l_id_language & " is not null;
           
                                       EXCEPTION
                                            WHEN no_data_found then
                                             CONTINUE;
                                     END;
            
                                        if (l_a_translation<>l_d_translation or (l_a_translation is null and l_d_translation is not null)) THEN
                                                  
                                                l_output:= 'Record ''' || pk_translation.get_translation(" & l_id_language & ", l_a_code_translation) || ''' has been updated to ''' ;
                                                                       
                                                pk_translation.insert_into_translation(" & l_id_language & ", l_a_code_translation, l_d_translation);
            
                                                l_output:= l_output || pk_translation.get_translation(" & l_id_language & ", l_a_code_translation) || '''  - ' || l_id_content || '.';

                                                if not save_output(l_output) then
      
                                                     dbms_output.put_line('ERROR');
    
                                                end if;

                                                contador := contador + 1;
            
                                        END IF;       

                                   END LOOP;
       
                                   close c_SAMPLE_TYPE;
       
                                   l_output:= to_char(contador) || ' record(s) updated!';
    
                                   if not save_output(l_output) then
      
                                         dbms_output.put_line('ERROR');
    
                                   end if;
                
                            END;"

        Dim cmd_update_sampe_type As New OracleCommand(sql, Connection.conn)

        Try
            cmd_update_sampe_type.CommandType = CommandType.Text
            cmd_update_sampe_type.ExecuteNonQuery()
        Catch ex As Exception
            cmd_update_sampe_type.Dispose()
            Return False
        End Try

        cmd_update_sampe_type.Dispose()
        Return True

    End Function

    Function UPDATE_ANALYSIS_SAMPLE_TYPE(ByVal i_institution As Int64) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution)

        Dim sql As String = "DECLARE

                                  l_a_code_translation translation.code_translation%type;
      
                                  l_a_translation      translation.desc_lang_6%type;
      
                                  l_d_translation      translation.desc_lang_6%type;
      
                                  l_id_content         alert.intervention.id_content%type;
      
                                  l_output     clob := '';    
      
                                  contador             integer;
      
                                  l_index integer := 1;
      
                                  l_record_area varchar2(50) := 'ANALYSIS_SAMPLE_TYPE'; 
      
                                  CURSOR c_ANALYSIS_SAMPLE_TYPE is
                                        select ast.id_content, ast.code_analysis_sample_type
                                        from alert.analysis_sample_type ast
                                        join translation t on t.code_translation=ast.code_analysis_sample_type;

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
       
                                   contador:=0;
                                   OPEN c_ANALYSIS_SAMPLE_TYPE;

                                   --COLOCAR NO LOG A ÁREA QUE ESTÁ A SER ATUALIZADA
                                   if not save_output(to_char(l_record_area)) then
        
                                               dbms_output.put_line('ERROR');
      
                                   end if;  
       
                                   LOOP
         
                                      FETCH c_ANALYSIS_SAMPLE_TYPE into l_id_content,l_a_code_translation;
                                        EXIT WHEN c_ANALYSIS_SAMPLE_TYPE%notfound;
            
                                        select t.desc_lang_" & l_id_language & "
                                          into l_a_translation
                                          from translation t
                                         where t.code_translation = l_a_code_translation;      
          
                                      BEGIN
   
                                        select distinct t.desc_lang_" & l_id_language & "
                                          into l_d_translation
                                          from alert_default.translation t
                                          join alert_default.analysis_sample_type ast
                                            on ast.code_analysis_sample_type = t.code_translation
                                         where ast.id_content = l_id_content
                                         and t.desc_lang_" & l_id_language & " is not null;
                          
                                         EXCEPTION
                                            WHEN no_data_found then
                                             CONTINUE;
                                        END;     
                 
                                        if (l_a_translation<>l_d_translation or (l_a_translation is null and l_d_translation is not null)) THEN
                                                  
                                                l_output:= 'Record ''' || pk_translation.get_translation(" & l_id_language & ", l_a_code_translation) || ''' has been updated to ''' ;
                                                                       
                                                pk_translation.insert_into_translation(" & l_id_language & ", l_a_code_translation, l_d_translation);
            
                                                l_output:= l_output || pk_translation.get_translation(" & l_id_language & ", l_a_code_translation) || '''  - ' || l_id_content || '.';

                                                if not save_output(l_output) then
      
                                                     dbms_output.put_line('ERROR');
    
                                                end if;

                                                contador := contador + 1;
            
                                        END IF;
       
                                   END LOOP;
       
                                   close c_ANALYSIS_SAMPLE_TYPE;
       
                                   l_output:= to_char(contador) || ' record(s) updated!';
    
                                   if not save_output(l_output) then
      
                                         dbms_output.put_line('ERROR');
    
                                   end if;          
                            END;"

        Dim cmd_update_analysis_sampe_type As New OracleCommand(sql, Connection.conn)

        Try
            cmd_update_analysis_sampe_type.CommandType = CommandType.Text
            cmd_update_analysis_sampe_type.ExecuteNonQuery()
        Catch ex As Exception
            cmd_update_analysis_sampe_type.Dispose()
            Return False
        End Try

        cmd_update_analysis_sampe_type.Dispose()
        Return True

    End Function

    Function UPDATE_ANALYSIS_PARAMETERS(ByVal i_institution As Int64) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution)

        Dim sql As String = "DECLARE

                                  l_a_code_translation translation.code_translation%type;
      
                                  l_a_translation      translation.desc_lang_6%type;
      
                                  l_d_translation      translation.desc_lang_6%type;
      
                                  l_id_content         alert.intervention.id_content%type;
      
                                  l_output     clob := '';         
      
                                  contador             integer;
      
                                  l_index integer := 1;
      
                                  l_record_area varchar2(50) := 'ANALYSIS_PARAMETERS';              
      
                                  CURSOR c_ANALYSIS_PARAMETER is
                                  select ap.id_content, ap.code_analysis_parameter
                                  from alert.analysis_parameter ap
                                  join translation t on t.code_translation=ap.code_analysis_parameter;

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
       
                                   contador:=0;
                                   OPEN c_ANALYSIS_PARAMETER;

                                   --COLOCAR NO LOG A ÁREA QUE ESTÁ A SER ATUALIZADA
                                   if not save_output(to_char(l_record_area)) then
        
                                               dbms_output.put_line('ERROR');
      
                                   end if;  
       
                                   LOOP
         
                                        FETCH c_ANALYSIS_PARAMETER into l_id_content,l_a_code_translation;
                                        EXIT WHEN c_ANALYSIS_PARAMETER%notfound;
            
                                        select t.desc_lang_" & l_id_language & "
                                          into l_a_translation
                                          from translation t
                                         where t.code_translation = l_a_code_translation;
       
                                        BEGIN
                    
                                              select distinct t.desc_lang_" & l_id_language & "
                                                into l_d_translation
                                                from alert_default.translation t
                                                join alert_default.analysis_parameter ap
                                                  on ap.code_analysis_parameter = t.code_translation
                                               where ap.id_content = l_id_content
                                               and ap.flg_available='Y' and t.desc_lang_" & l_id_language & " is not null;
           
                                        EXCEPTION
                                            WHEN no_data_found then
                                             CONTINUE;
            
                                        END;
            
                                        if (l_a_translation<>l_d_translation or (l_a_translation is null and l_d_translation is not null)) THEN
                                                  
                                                l_output:= 'Record ''' || pk_translation.get_translation(" & l_id_language & ", l_a_code_translation) || ''' has been updated to ''' ;
                                                                       
                                                pk_translation.insert_into_translation(" & l_id_language & ", l_a_code_translation, l_d_translation);
            
                                                l_output:= l_output || pk_translation.get_translation(" & l_id_language & ", l_a_code_translation) || '''  - ' || l_id_content || '.';

                                                if not save_output(l_output) then
      
                                                     dbms_output.put_line('ERROR');
    
                                                end if;

                                                contador := contador + 1;
            
                                        END IF;

                                   END LOOP;
       
                                   close c_ANALYSIS_PARAMETER;
       
                                   l_output:= to_char(contador) || ' record(s) updated!';
      
                                   if not save_output(l_output) then
        
                                         dbms_output.put_line('ERROR');
      
                                   end if;       
                
                            END;"

        Dim cmd_update_analysis_parameters As New OracleCommand(sql, Connection.conn)

        Try
            cmd_update_analysis_parameters.CommandType = CommandType.Text
            cmd_update_analysis_parameters.ExecuteNonQuery()
        Catch ex As Exception
            cmd_update_analysis_parameters.Dispose()
            Return False
        End Try

        cmd_update_analysis_parameters.Dispose()
        Return True

    End Function

    Function UPDATE_ANALYSIS_SR(ByVal i_institution As Int64) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution)

        Dim sql As String = "DECLARE

                                  l_a_code_translation translation.code_translation%type;
      
                                  l_a_translation      translation.desc_lang_6%type;
      
                                  l_d_translation      translation.desc_lang_6%type;
      
                                  l_id_content         alert.intervention.id_content%type;
      
                                  contador             integer;
      
                                  l_output     clob := '';               
      
                                  l_index integer := 1;
      
                                  l_record_area varchar2(50) := 'ANALYSIS_SAMPLE_RECIPIENT';  
      
                                  CURSOR c_SAMPLE_RECIPIENT is
                                  select sr.id_content, sr.code_sample_recipient
                                  from alert.sample_recipient sr
                                  join translation t on t.code_translation=sr.code_sample_recipient;
      
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
       
                                   contador:=0;
                                   OPEN c_SAMPLE_RECIPIENT;
                                 
                                   --COLOCAR NO LOG A ÁREA QUE ESTÁ A SER ATUALIZADA
                                   if not save_output(to_char(l_record_area)) then
        
                                               dbms_output.put_line('ERROR');
      
                                   end if;  
       
                                   LOOP
         
                                      FETCH c_SAMPLE_RECIPIENT into l_id_content,l_a_code_translation;
                                        EXIT WHEN c_SAMPLE_RECIPIENT%notfound;
            
                                        select t.desc_lang_" & l_id_language & "
                                          into l_a_translation
                                          from translation t
                                         where t.code_translation = l_a_code_translation;            
             
                                      BEGIN
            
                                        select distinct t.desc_lang_" & l_id_language & "
                                          into l_d_translation
                                          from alert_default.translation t
                                          join alert_default.sample_recipient sr
                                            on sr.code_sample_recipient = t.code_translation
                                         where sr.id_content = l_id_content
                                         and t.desc_lang_" & l_id_language & " is not null
                                         and sr.flg_available='Y';

                                      EXCEPTION
                                            WHEN no_data_found then
                                             CONTINUE;
                                      END;                                                        

                                        if (l_a_translation<>l_d_translation or (l_a_translation is null and l_d_translation is not null)) THEN
                                                  
                                                l_output:= l_output || 'Record ''' || pk_translation.get_translation(" & l_id_language & ", l_a_code_translation) || ''' has been updated to ''' ;
                                                                       
                                                pk_translation.insert_into_translation(" & l_id_language & ", l_a_code_translation, l_d_translation);
            
                                                l_output:= l_output || pk_translation.get_translation(" & l_id_language & ", l_a_code_translation) || '''  - ' || l_id_content || '.';

                                                if not save_output(l_output) then
      
                                                     dbms_output.put_line('ERROR');
    
                                                end if;

                                                contador := contador + 1;
            
                                        END IF;
       
                                   END LOOP;
       
                                   close c_SAMPLE_RECIPIENT;
       
                                   l_output:= to_char(contador) || ' record(s) updated!';
      
                                   if not save_output(l_output) then
        
                                         dbms_output.put_line('ERROR');
      
                                   end if;                      
                            END;"

        Dim cmd_update_analysis_sr As New OracleCommand(sql, Connection.conn)

        Try
            cmd_update_analysis_sr.CommandType = CommandType.Text
            cmd_update_analysis_sr.ExecuteNonQuery()
        Catch ex As Exception
            cmd_update_analysis_sr.Dispose()
            Return False
        End Try

        cmd_update_analysis_sr.Dispose()
        Return True

    End Function

    Function GET_UPDATED_RECORDS(ByRef i_dr As OracleDataReader) As Boolean

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
