﻿Imports Oracle.DataAccess.Client

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
    
                                l_record_area varchar2(50) := 'EXAMS';
    
                                CURSOR c_exam IS
                                    SELECT e.id_content, e.code_exam
                                    FROM alert.exam e
                                    JOIN translation t ON t.code_translation = e.code_exam;

                                FUNCTION save_output(i_updated_records IN CLOB) RETURN BOOLEAN IS
                                    
                                   l_index integer;  

                                begin
  
                                    select (nvl(max(r.record_index),0)+1)
                                    into l_index  
                                    from output_records r;
        
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

                                --Garantir linha extra no final do log
                                if not save_output(' ') then
      
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
      
                                      l_record_area varchar2(50) := 'EXAM_CATEGORIES';
      
                                      CURSOR c_EXAM_CAT is
                                      select ec.id_content, ec.code_exam_cat
                                      from alert.exam_cat ec
                                      join translation t on t.code_translation=ec.code_exam_cat
                                      where ec.flg_lab='N';  

                                      FUNCTION save_output(i_updated_records IN CLOB) RETURN BOOLEAN IS
    
                                       l_index integer;  

                                    begin
  
                                        select (nvl(max(r.record_index),0)+1)
                                        into l_index  
                                        from output_records r;
        
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
                                                and ec.flg_lab='N'
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

                                --Garantir linha extra no final do log
                                if not save_output(' ') then
      
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
             
                                      l_record_area varchar2(50) := 'INTERVENTIONS';
      
                                      CURSOR c_INTERVENTION is
                                      select i.id_content, i.code_intervention
                                      from alert.intervention i
                                      join translation t on t.code_translation=i.code_intervention;
      
                                      FUNCTION save_output(i_updated_records IN CLOB) RETURN BOOLEAN IS
    
                                       l_index integer;  

                                    begin
  
                                        select (nvl(max(r.record_index),0)+1)
                                        into l_index  
                                        from output_records r;
        
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

                                --Garantir linha extra no final do log
                                if not save_output(' ') then
      
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
      
                                  l_record_area varchar2(50) := 'ANALYSIS';

                                  CURSOR c_ANALYSIS is
                                    select a.id_content, a.code_analysis
                                      from alert.analysis a
                                      join translation t
                                        on t.code_translation = a.code_analysis;
            
                                    FUNCTION save_output(i_updated_records IN CLOB) RETURN BOOLEAN IS
      
                                       l_index integer;  

                                    begin
  
                                        select (nvl(max(r.record_index),0)+1)
                                        into l_index  
                                        from output_records r;
          
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

                                --Garantir linha extra no final do log
                                if not save_output(' ') then
      
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
    
                                  l_record_area varchar2(50) := 'SAMPLE_TYPE';      
      
                                  CURSOR c_SAMPLE_TYPE is
                                  select st.id_content, st.code_sample_type
                                  from alert.sample_type st
                                  join translation t on t.code_translation=st.code_sample_type
                                  where st.flg_available='Y';
      
                                  FUNCTION save_output(i_updated_records IN CLOB) RETURN BOOLEAN IS
      
                                       l_index integer;  

                                    begin
  
                                        select (nvl(max(r.record_index),0)+1)
                                        into l_index  
                                        from output_records r;
          
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

                                --Garantir linha extra no final do log
                                if not save_output(' ') then
      
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
      
                                  l_record_area varchar2(50) := 'ANALYSIS_SAMPLE_TYPE'; 
      
                                  CURSOR c_ANALYSIS_SAMPLE_TYPE is
                                        select ast.id_content, ast.code_analysis_sample_type
                                        from alert.analysis_sample_type ast
                                        join translation t on t.code_translation=ast.code_analysis_sample_type;

                                  FUNCTION save_output(i_updated_records IN CLOB) RETURN BOOLEAN IS
      
                                       l_index integer;  

                                    begin
  
                                        select (nvl(max(r.record_index),0)+1)
                                        into l_index  
                                        from output_records r;
          
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

                                --Garantir linha extra no final do log
                                if not save_output(' ') then
      
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
      
                                  l_record_area varchar2(50) := 'ANALYSIS_PARAMETERS';              
      
                                  CURSOR c_ANALYSIS_PARAMETER is
                                  select ap.id_content, ap.code_analysis_parameter
                                  from alert.analysis_parameter ap
                                  join translation t on t.code_translation=ap.code_analysis_parameter;

                                  FUNCTION save_output(i_updated_records IN CLOB) RETURN BOOLEAN IS
      
                                       l_index integer;  

                                    begin
  
                                        select (nvl(max(r.record_index),0)+1)
                                        into l_index  
                                        from output_records r;
          
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

                                --Garantir linha extra no final do log
                                if not save_output(' ') then
      
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
      
                                  l_record_area varchar2(50) := 'ANALYSIS_SAMPLE_RECIPIENT';  
      
                                  CURSOR c_SAMPLE_RECIPIENT is
                                  select sr.id_content, sr.code_sample_recipient
                                  from alert.sample_recipient sr
                                  join translation t on t.code_translation=sr.code_sample_recipient;
      
                                  FUNCTION save_output(i_updated_records IN CLOB) RETURN BOOLEAN IS
                                       
                                        l_index integer;  

                                    begin
  
                                        select (nvl(max(r.record_index),0)+1)
                                        into l_index  
                                        from output_records r;
          
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
                                                  
                                                l_output:= 'Record ''' || pk_translation.get_translation(" & l_id_language & ", l_a_code_translation) || ''' has been updated to ''' ;
                                                                       
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

                                --Garantir linha extra no final do log
                                if not save_output(' ') then
      
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

    Function UPDATE_ANALYSIS_CAT(ByVal i_institution As Int64) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution)

        Dim sql As String = "DECLARE

                                  l_a_code_translation translation.code_translation%type;
      
                                  l_a_translation      translation.desc_lang_6%type;
      
                                  l_d_translation      translation.desc_lang_6%type;
      
                                  l_id_content         alert.diet.id_content%type;
      
                                  contador             integer;

                                  l_output     clob := '';               
      
                                  l_record_area varchar2(50) := 'ANALYSIS_CATEGORY';              
      
                                  CURSOR c_EXAM_CAT is
                                  select ec.id_content, ec.code_exam_cat
                                  from alert.exam_cat ec
                                  join translation t on t.code_translation=ec.code_exam_cat
                                  where ec.flg_lab='Y';  

                                  FUNCTION save_output(i_updated_records IN CLOB) RETURN BOOLEAN IS
                                       
                                        l_index integer;  

                                    begin
  
                                        select (nvl(max(r.record_index),0)+1)
                                        into l_index  
                                        from output_records r;
          
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
            
                                            select t.desc_lang_" & l_id_language & "
                                            into  l_d_translation
                                            from alert_default.translation t
                                            join alert_default.exam_cat ec on ec.code_exam_cat=t.code_translation
                                            where ec.id_content=l_id_content
                                            and ec.flg_lab='Y'
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

                                  --Garantir linha extra no final do log
                                  if not save_output(' ') then
      
                                     dbms_output.put_line('ERROR');
    
                                  end if;      
             
                            END;"

        Dim cmd_update_analysis_cat As New OracleCommand(sql, Connection.conn)

        Try
            cmd_update_analysis_cat.CommandType = CommandType.Text
            cmd_update_analysis_cat.ExecuteNonQuery()
        Catch ex As Exception
            cmd_update_analysis_cat.Dispose()
            Return False
        End Try

        cmd_update_analysis_cat.Dispose()
        Return True

    End Function

    Function UPDATE_INTERV_CAT(ByVal i_institution As Int64) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution)

        Dim sql As String = "DECLARE

                                  l_a_code_translation translation.code_translation%type;
      
                                  l_a_translation      translation.desc_lang_6%type;
      
                                  l_d_translation      translation.desc_lang_6%type;
      
                                  l_id_content         alert.diet.id_content%type;

                                  l_sql           VARCHAR2(3000) := '';
      
                                  contador             integer;
      
                                   l_output     clob := '';               
      
                                                              l_record_area varchar2(50) := 'INTERVENTION_CATEGORY';         
      
                                  CURSOR c_INTERV_CAT is
                                  select ic.id_content, ic.code_interv_category
                                  from alert.interv_category ic
                                  join translation t on t.code_translation=ic.code_interv_category; 
      
                                  FUNCTION save_output(i_updated_records IN CLOB) RETURN BOOLEAN IS
                                       
                                        l_index integer;  

                                    begin
  
                                        select (nvl(max(r.record_index),0)+1)
                                        into l_index  
                                        from output_records r;
          
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
                                   OPEN c_INTERV_CAT;
       
                                   --COLOCAR NO LOG A ÁREA QUE ESTÁ A SER ATUALIZADA
                                   if not save_output(to_char(l_record_area)) then
        
                                               dbms_output.put_line('ERROR');
      
                                   end if;     
       
                                   LOOP
         
                                        FETCH c_INTERV_CAT into l_id_content,l_a_code_translation;
                                        EXIT WHEN c_INTERV_CAT%notfound;
            
                                        select t.desc_lang_" & l_id_language & "
                                        into  l_a_translation
                                        from translation t 
                                        where t.code_translation=l_a_code_translation;
            
                                        declare
            
                                           no_such_table EXCEPTION;
                                           PRAGMA EXCEPTION_INIT(no_such_table, -942);   
            
                                        BEGIN
            
                                            l_sql := 'select t.desc_lang_" & l_id_language & "      
                                            from alert_default.interv_category ic
                                            join alert_default.translation t on ic.code_interv_category=t.code_translation
                                            where ic.id_content= ''' || l_id_content || '''
                                            and t.desc_lang_" & l_id_language & " is not null';
                                        
                                            EXECUTE immediate l_sql into  l_d_translation;
           
                                       EXCEPTION
                                            WHEN no_data_found then
                                              CONTINUE;
                                  
                                            WHEN no_such_table THEN
                                               if not save_output('Intervention category does not need to be updated on this version of ALERT(R).') then
        
                                                    dbms_output.put_line('ERROR');
      
                                                end if;                                                     
                                                exit;  
             
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
       
                                   close c_INTERV_CAT;
       
       
                                     l_output:= to_char(contador) || ' record(s) updated!';
      
                                     if not save_output(l_output) then
        
                                           dbms_output.put_line('ERROR');
      
                                     end if;    

                                  --Garantir linha extra no final do log
                                  if not save_output(' ') then
      
                                     dbms_output.put_line('ERROR');
    
                                  end if;     

                            END;"

        Dim cmd_update_interv_cat As New OracleCommand(sql, Connection.conn)

        Try
            cmd_update_interv_cat.CommandType = CommandType.Text
            cmd_update_interv_cat.ExecuteNonQuery()
        Catch ex As Exception
            cmd_update_interv_cat.Dispose()
            Return False
        End Try

        cmd_update_interv_cat.Dispose()
        Return True

    End Function

    Function UPDATE_SR_INTERV(ByVal i_institution As Int64) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution)

        Dim sql As String = "DECLARE

                                  l_a_code_translation translation.code_translation%type;
      
                                  l_a_translation      translation.desc_lang_6%type;
      
                                  l_d_translation      translation.desc_lang_6%type;
      
                                  l_id_content         alert.sr_intervention.id_content%type;
      
                                  contador             integer;

                                  l_output     clob := '';               
      
                                  l_record_area varchar2(50) := 'SR_INTERVENTION';   
      
                                  CURSOR c_INTERVENTION is
                                  select i.id_content, i.code_sr_intervention
                                  from alert.sr_intervention i
                                  join translation t on t.code_translation=i.code_sr_intervention;

                                  FUNCTION save_output(i_updated_records IN CLOB) RETURN BOOLEAN IS
                                       
                                        l_index integer;  

                                    begin
  
                                        select (nvl(max(r.record_index),0)+1)
                                        into l_index  
                                        from output_records r;
          
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
              
                                        select t.desc_lang_" & l_id_language & "
                                        into  l_d_translation
                                        from alert_default.translation t
                                        join alert_default.sr_intervention i on i.code_sr_intervention=t.code_translation
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

                                    --Garantir linha extra no final do log
                                    if not save_output(' ') then
      
                                       dbms_output.put_line('ERROR');
    
                                    end if;
             
                            END;"

        Dim cmd_update_sr_interv As New OracleCommand(sql, Connection.conn)

        Try
            cmd_update_sr_interv.CommandType = CommandType.Text
            cmd_update_sr_interv.ExecuteNonQuery()
        Catch ex As Exception
            cmd_update_sr_interv.Dispose()
            Return False
        End Try

        cmd_update_sr_interv.Dispose()
        Return True

    End Function

    Function UPDATE_SUPPLIES(ByVal i_institution As Int64) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution)

        Dim sql As String = "DECLARE

                                  l_a_code_translation translation.code_translation%type;
      
                                  l_a_translation      translation.desc_lang_6%type;
      
                                  l_d_translation      translation.desc_lang_6%type;
      
                                  l_id_content         alert.sr_intervention.id_content%type;
      
                                  contador             integer;
      
                                  l_output     clob := '';               
      
                                  l_record_area varchar2(50) := 'SUPPLY';      
      
                                  CURSOR c_SUPPLIES is
                                  select s.id_content, s.code_supply
                                  from alert.supply s
                                  join translation t on t.code_translation=s.code_supply;

                                  FUNCTION save_output(i_updated_records IN CLOB) RETURN BOOLEAN IS
                                       
                                        l_index integer;  

                                    begin
  
                                        select (nvl(max(r.record_index),0)+1)
                                        into l_index  
                                        from output_records r;
          
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
                                   OPEN c_SUPPLIES;
       
                                   --COLOCAR NO LOG A ÁREA QUE ESTÁ A SER ATUALIZADA
                                   if not save_output(to_char(l_record_area)) then
        
                                               dbms_output.put_line('ERROR');
      
                                   end if;              
       
                                   LOOP
         
                                       FETCH c_SUPPLIES into l_id_content,l_a_code_translation;
                                       EXIT WHEN c_SUPPLIES%notfound;            

                                       select t.desc_lang_" & l_id_language & "
                                       into  l_a_translation
                                       from translation t 
                                       where t.code_translation=l_a_code_translation;
          
                                   BEGIN  
              
                                        select t.desc_lang_" & l_id_language & "
                                        into  l_d_translation
                                        from alert_default.translation t
                                        join alert_default.supply s on s.code_supply=t.code_translation
                                        where s.id_content=l_id_content
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
       
                                   close c_SUPPLIES;
       
                                         l_output:= to_char(contador) || ' record(s) updated!';
      
                                         if not save_output(l_output) then
        
                                               dbms_output.put_line('ERROR');
      
                                         end if;    

                                      --Garantir linha extra no final do log
                                      if not save_output(' ') then
      
                                         dbms_output.put_line('ERROR');
    
                                      end if;   
          
                            END;"

        Dim cmd_update_supplies As New OracleCommand(sql, Connection.conn)

        Try
            cmd_update_supplies.CommandType = CommandType.Text
            cmd_update_supplies.ExecuteNonQuery()
        Catch ex As Exception
            cmd_update_supplies.Dispose()
            Return False
        End Try

        cmd_update_supplies.Dispose()
        Return True

    End Function

    Function UPDATE_SUPPLIES_CAT(ByVal i_institution As Int64) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution)

        Dim sql As String = "DECLARE

                                  l_a_code_translation translation.code_translation%type;
      
                                  l_a_translation      translation.desc_lang_6%type;
      
                                  l_d_translation      translation.desc_lang_6%type;
      
                                  l_id_content         alert.sr_intervention.id_content%type;
      
                                  contador             integer;
      
                                  l_output     clob := '';               
      
                                  l_record_area varchar2(50) := 'SUPPLY_CATEGORY';          
      
                                  CURSOR c_SUPPLIES_CAT is
                                  select s.id_content, s.code_supply_type
                                  from alert.supply_type  s
                                  join translation t on t.code_translation=s.code_supply_type
                                  and s.id_content is not null; -- Existem registos no default sem id_content 

                                  FUNCTION save_output(i_updated_records IN CLOB) RETURN BOOLEAN IS
                                       
                                        l_index integer;  

                                    begin
  
                                        select (nvl(max(r.record_index),0)+1)
                                        into l_index  
                                        from output_records r;
          
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
                                   OPEN c_SUPPLIES_CAT;
       
                                   --COLOCAR NO LOG A ÁREA QUE ESTÁ A SER ATUALIZADA
                                   if not save_output(to_char(l_record_area)) then
        
                                               dbms_output.put_line('ERROR');
      
                                   end if;     
       
                                   LOOP
         
                                        FETCH c_SUPPLIES_CAT into l_id_content,l_a_code_translation;
                                        EXIT WHEN c_SUPPLIES_CAT%notfound;
            

                                       select t.desc_lang_" & l_id_language & "
                                       into  l_a_translation
                                       from translation t 
                                       where t.code_translation=l_a_code_translation;
          
                                   BEGIN  
              
                                        select t.desc_lang_" & l_id_language & "
                                        into  l_d_translation
                                        from alert_default.translation t
                                        join alert_default.supply_type s on s.code_supply_type=t.code_translation
                                        where s.id_content=l_id_content
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
       
                                   close c_SUPPLIES_CAT;
       
                                   l_output:= to_char(contador) || ' record(s) updated!';
      
                                   if not save_output(l_output) then
        
                                         dbms_output.put_line('ERROR');
      
                                   end if;    

                                    --Garantir linha extra no final do log
                                    if not save_output(' ') then
      
                                       dbms_output.put_line('ERROR');
    
                                    end if;   
             
                            END;"

        Dim cmd_update_supplies_cat As New OracleCommand(sql, Connection.conn)

        Try
            cmd_update_supplies_cat.CommandType = CommandType.Text
            cmd_update_supplies_cat.ExecuteNonQuery()
        Catch ex As Exception
            cmd_update_supplies_cat.Dispose()
            Return False
        End Try

        cmd_update_supplies_cat.Dispose()
        Return True

    End Function

    Function GET_UPDATED_RECORDS(ByRef i_dr As OracleDataReader) As Boolean

        Dim sql As String = "SELECT desc_record as ""UPDATE LOG""
                            FROM (SELECT r.record_index ""INDEX_RECORD"", r.updated_records ""DESC_RECORD""
                                  FROM output_records r
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
