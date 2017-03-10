Imports Oracle.DataAccess.Client
Public Class SUPPLIES_API

    Dim db_access_general As New General

    Public Structure SUP_AREAS
        Public id_supply_area As Integer
        Public desc_supply_area As String
    End Structure

    Public Structure supplies_default
        Public id_content_category As String
        Public desc_category As String
        Public id_content_supply As String
        Public desc_supply As String
    End Structure

    Function GET_SUP_AREAS(ByVal i_conn As OracleConnection, ByRef i_dr As OracleDataReader) As Boolean

        Dim sql As String = "SELECT s.id_supply_area, pk_translation.get_translation(2, s.code_supply_area)
                             FROM alert.supply_area s"

        sql = sql & "  ORDER BY 2 ASC"

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

    Function GET_DEFAULT_VERSIONS(ByVal i_institution As Int64, ByVal i_software As Int16, ByVal i_sup_area As Int16, ByVal i_sup_type As String, ByVal i_conn As OracleConnection, ByRef i_dr As OracleDataReader) As Boolean

        Dim sql As String = ""

        If i_sup_type = "ALL" Then

            sql = "SELECT DISTINCT (dmv.version)
                                FROM alert_default.supply_mrk_vrs dmv
                                JOIN alert_default.supply ds ON ds.id_supply = dmv.id_supply
                                JOIN alert_default.supply_soft_inst dssi ON dssi.id_supply = ds.id_supply
                                JOIN alert_default.supply_sup_area dssa ON dssa.id_supply_soft_inst = dssi.id_supply_soft_inst
                                JOIN alert_default.supply_type dst ON dst.id_supply_type = ds.id_supply_type
                                JOIN alert_core_data.ab_institution i ON i.id_ab_market = dmv.id_market
                                WHERE i.id_ab_institution = " & i_institution & "
                                AND ds.flg_available = 'Y'
                                AND dssi.id_software IN (0, " & i_software & ") --Há para software 0
                                AND dssa.id_supply_area = " & i_sup_area & "
                                AND dssa.flg_available = 'Y'
                                AND dst.flg_available = 'Y'
                                ORDER BY 1 ASC"
        Else

            sql = "SELECT DISTINCT (dmv.version)
                                FROM alert_default.supply_mrk_vrs dmv
                                JOIN alert_default.supply ds ON ds.id_supply = dmv.id_supply
                                JOIN alert_default.supply_soft_inst dssi ON dssi.id_supply = ds.id_supply
                                JOIN alert_default.supply_sup_area dssa ON dssa.id_supply_soft_inst = dssi.id_supply_soft_inst
                                JOIN alert_default.supply_type dst ON dst.id_supply_type = ds.id_supply_type
                                JOIN alert_core_data.ab_institution i ON i.id_ab_market = dmv.id_market
                                WHERE i.id_ab_institution = " & i_institution & "
                                AND ds.flg_available = 'Y'
                                AND ds.flg_type = '" & i_sup_type & "'
                                AND dssi.id_software IN (0, " & i_software & ") --Há para software 0
                                AND dssa.id_supply_area = " & i_sup_area & "
                                AND dssa.flg_available = 'Y'
                                AND dst.flg_available = 'Y'
                                ORDER BY 1 ASC"

        End If

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

    Function GET_SUPP_CATS_DEFAULT(ByVal i_institution As Int64, ByVal i_software As Integer, ByVal i_version As String, ByVal i_sup_area As Int16, ByVal i_flg_type As String, ByVal i_conn As OracleConnection, ByRef i_dr As OracleDataReader) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution, i_conn)

        Dim sql As String = ""

        If i_flg_type = "ALL" Then

            sql = "SELECT DISTINCT dst.id_content, alert_default.pk_translation_default.get_translation_default(" & l_id_language & ", dst.code_supply_type)
                                FROM alert_default.supply_mrk_vrs dmv
                                JOIN alert_default.supply ds ON ds.id_supply = dmv.id_supply
                                JOIN alert_default.supply_soft_inst dssi ON dssi.id_supply = ds.id_supply
                                JOIN alert_default.supply_sup_area dssa ON dssa.id_supply_soft_inst = dssi.id_supply_soft_inst
                                JOIN alert_default.supply_type dst ON dst.id_supply_type = ds.id_supply_type
                                JOIN alert_core_data.ab_institution i ON i.id_ab_market = dmv.id_market
                                WHERE i.id_ab_institution = " & i_institution & "
                                AND dmv.version = '" & i_version & "'
                                AND ds.flg_available = 'Y'
                                AND dssi.id_software IN (0, " & i_software & ")
                                AND dssa.id_supply_area = " & i_sup_area & "
                                AND dssa.flg_available = 'Y'
                                AND dst.flg_available = 'Y'
                                AND alert_default.pk_translation_default.get_translation_default(" & l_id_language & ", dst.code_supply_type) IS NOT NULL
                                ORDER BY 2 ASC"

        Else

            sql = "SELECT DISTINCT dst.id_content, alert_default.pk_translation_default.get_translation_default(" & l_id_language & ", dst.code_supply_type)
                                FROM alert_default.supply_mrk_vrs dmv
                                JOIN alert_default.supply ds ON ds.id_supply = dmv.id_supply
                                JOIN alert_default.supply_soft_inst dssi ON dssi.id_supply = ds.id_supply
                                JOIN alert_default.supply_sup_area dssa ON dssa.id_supply_soft_inst = dssi.id_supply_soft_inst
                                JOIN alert_default.supply_type dst ON dst.id_supply_type = ds.id_supply_type
                                JOIN alert_core_data.ab_institution i ON i.id_ab_market = dmv.id_market
                                WHERE i.id_ab_institution = " & i_institution & "
                                AND dmv.version = '" & i_version & "'
                                AND ds.flg_available = 'Y'
                                AND ds.flg_type = '" & i_flg_type & "'
                                AND dssi.id_software IN (0, " & i_software & ")
                                AND dssa.id_supply_area = " & i_sup_area & "
                                AND dssa.flg_available = 'Y'
                                AND dst.flg_available = 'Y'
                                AND alert_default.pk_translation_default.get_translation_default(" & l_id_language & ", dst.code_supply_type) IS NOT NULL
                                ORDER BY 2 ASC"
        End If

        Dim cmd As New OracleCommand(sql, i_conn)
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

    Function GET_SUPP_CATS_ALERT(ByVal i_institution As Int64, ByVal i_software As Integer, ByVal i_sup_area As Int16, ByVal i_flg_type As String, ByVal i_conn As OracleConnection, ByRef i_dr As OracleDataReader) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution, i_conn)

        Dim sql As String = ""

        If i_flg_type = "ALL" Then

            sql = "SELECT distinct st.id_content, tst.desc_lang_" & l_id_language & "
                    FROM alert.supply s
                    JOIN alert.supply_soft_inst ssi ON ssi.id_supply = s.id_supply
                    JOIN alert.supply_sup_area ssa ON ssa.id_supply_soft_inst = ssi.id_supply_soft_inst
                    JOIN alert.supply_type st ON st.id_supply_type = s.id_supply_type
                    join translation tst on tst.code_translation=st.code_supply_type
                    join translation ts on ts.code_translation=s.code_supply
                    join alert.supply_loc_default sl on sl.id_supply_soft_inst=ssi.id_supply_soft_inst
                    WHERE ssi.id_institution= " & i_institution & "
                    AND s.flg_available = 'Y'
                    AND ssi.id_software IN (0, " & i_software & ")
                    AND ssa.id_supply_area = " & i_sup_area & "
                    AND ssa.flg_available = 'Y'
                    AND st.flg_available = 'Y'
                    AND tst.desc_lang_" & l_id_language & " is not null
                    and ts.desc_lang_" & l_id_language & " is not null
                    ORDER BY 2 ASC"

        Else

            sql = "SELECT distinct st.id_content, tst.desc_lang_" & l_id_language & "
                    FROM alert.supply s
                    JOIN alert.supply_soft_inst ssi ON ssi.id_supply = s.id_supply
                    JOIN alert.supply_sup_area ssa ON ssa.id_supply_soft_inst = ssi.id_supply_soft_inst
                    JOIN alert.supply_type st ON st.id_supply_type = s.id_supply_type
                    join translation tst on tst.code_translation=st.code_supply_type
                    join translation ts on ts.code_translation=s.code_supply
                    join alert.supply_loc_default sl on sl.id_supply_soft_inst=ssi.id_supply_soft_inst
                    WHERE ssi.id_institution= " & i_institution & "
                    AND s.flg_available = 'Y'
                    AND ssi.id_software IN (0, " & i_software & ")
                    AND ssa.id_supply_area = " & i_sup_area & "
                    AND s.flg_type = '" & i_flg_type & "'
                    AND ssa.flg_available = 'Y'
                    AND st.flg_available = 'Y'
                    AND tst.desc_lang_" & l_id_language & " is not null
                    and ts.desc_lang_" & l_id_language & " is not null
                    ORDER BY 2 ASC"

        End If

        Dim cmd As New OracleCommand(sql, i_conn)
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

    Function GET_SUPS_DEFAULT_BY_CAT(ByVal i_institution As Int64, ByVal i_software As Integer, ByVal i_version As String, ByVal i_sup_area As Int16, ByVal i_flg_type As String, ByVal i_id_content_cat As String, ByVal i_conn As OracleConnection, ByRef i_dr As OracleDataReader) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution, i_conn)

        Dim sql As String = ""

        If i_id_content_cat = "0" Then

            sql = "SELECT DISTINCT dst.id_content, alert_default.pk_translation_default.get_translation_default(" & l_id_language & ", dst.code_supply_type), DS.ID_CONTENT, alert_default.pk_translation_default.get_translation_default(" & l_id_language & ", ds.Code_Supply)
                                FROM alert_default.supply_mrk_vrs dmv
                                JOIN alert_default.supply ds ON ds.id_supply = dmv.id_supply
                                JOIN alert_default.supply_soft_inst dssi ON dssi.id_supply = ds.id_supply
                                JOIN alert_default.supply_sup_area dssa ON dssa.id_supply_soft_inst = dssi.id_supply_soft_inst
                                JOIN alert_default.supply_type dst ON dst.id_supply_type = ds.id_supply_type
                                JOIN alert_core_data.ab_institution i ON i.id_ab_market = dmv.id_market
                                WHERE i.id_ab_institution = " & i_institution & "
                                AND dmv.version = '" & i_version & "'
                                AND ds.flg_available = 'Y'"

            If i_flg_type <> "ALL" Then

                sql = sql & "AND ds.flg_type = '" & i_flg_type & "'"

            End If

            sql = sql & "      
                                AND dssi.id_software IN (0, " & i_software & ")
                                AND dssa.id_supply_area = " & i_sup_area & "
                                AND dssa.flg_available = 'Y'
                                AND dst.flg_available = 'Y'
                                AND alert_default.pk_translation_default.get_translation_default(" & l_id_language & ", dst.code_supply_type) IS NOT NULL
                                AND alert_default.pk_translation_default.get_translation_default(" & l_id_language & ", ds.Code_Supply) IS NOT NULL
                                ORDER BY 4 ASC"

        Else

            sql = "SELECT DISTINCT dst.id_content, alert_default.pk_translation_default.get_translation_default(" & l_id_language & ", dst.code_supply_type), DS.ID_CONTENT, alert_default.pk_translation_default.get_translation_default(" & l_id_language & ", ds.Code_Supply)
                                FROM alert_default.supply_mrk_vrs dmv
                                JOIN alert_default.supply ds ON ds.id_supply = dmv.id_supply
                                JOIN alert_default.supply_soft_inst dssi ON dssi.id_supply = ds.id_supply
                                JOIN alert_default.supply_sup_area dssa ON dssa.id_supply_soft_inst = dssi.id_supply_soft_inst
                                JOIN alert_default.supply_type dst ON dst.id_supply_type = ds.id_supply_type
                                JOIN alert_core_data.ab_institution i ON i.id_ab_market = dmv.id_market
                                WHERE i.id_ab_institution = " & i_institution & "
                                AND dmv.version = '" & i_version & "'
                                AND ds.flg_available = 'Y'"


            If i_flg_type <> "ALL" Then

                sql = sql & "AND ds.flg_type = '" & i_flg_type & "'"

            End If

            sql = sql & "       And dssi.id_software In (0, " & i_software & ")
                                AND dssa.id_supply_area = " & i_sup_area & "
                                AND dssa.flg_available = 'Y'
                                AND dst.flg_available = 'Y'
                                AND alert_default.pk_translation_default.get_translation_default(" & l_id_language & ", dst.code_supply_type) IS NOT NULL
                                AND alert_default.pk_translation_default.get_translation_default(" & l_id_language & ", ds.Code_Supply) IS NOT NULL
                                AND DST.ID_CONTENT='" & i_id_content_cat & "'
                                ORDER BY 4 ASC"

        End If

        Dim cmd As New OracleCommand(sql, i_conn)
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


    Function GET_SUPS_ALERT_BY_CAT(ByVal i_institution As Int64, ByVal i_software As Integer, ByVal i_sup_area As Int16, ByVal i_flg_type As String, ByVal i_id_content_cat As String, ByVal i_conn As OracleConnection, ByRef i_dr As OracleDataReader) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution, i_conn)

        Dim sql As String = ""

        If i_id_content_cat = "0" Then

            sql = "SELECT DISTINCT st.id_content, tst.desc_lang_" & l_id_language & ", s.id_content, ts.desc_lang_" & l_id_language & "
                    FROM alert.supply s
                    JOIN alert.supply_soft_inst ssi ON ssi.id_supply = s.id_supply
                    JOIN alert.supply_sup_area ssa ON ssa.id_supply_soft_inst = ssi.id_supply_soft_inst
                    JOIN alert.supply_type st ON st.id_supply_type = s.id_supply_type
                    JOIN translation tst ON tst.code_translation = st.code_supply_type
                    JOIN translation ts ON ts.code_translation = s.code_supply
                    JOIN alert.supply_loc_default sl ON sl.id_supply_soft_inst = ssi.id_supply_soft_inst
                    WHERE ssi.id_institution = " & i_institution & "
                    AND s.flg_available = 'Y'
                    AND ssi.id_software IN (0, " & i_software & ")
                    AND ssa.id_supply_area = " & i_sup_area & "                   
                    AND ssa.flg_available = 'Y'
                    AND st.flg_available = 'Y'
                    AND tst.desc_lang_" & l_id_language & " IS NOT NULL
                    AND ts.desc_lang_" & l_id_language & " IS NOT NULL "

            If i_flg_type <> "ALL" Then

                sql = sql & " And s.flg_type = '" & i_flg_type & "'"

            End If

            sql = sql & " ORDER BY 4 asc , 2 ASC"

        Else

            sql = sql & "SELECT DISTINCT st.id_content, tst.desc_lang_" & l_id_language & ", s.id_content, ts.desc_lang_" & l_id_language & "
                            FROM alert.supply s
                            JOIN alert.supply_soft_inst ssi ON ssi.id_supply = s.id_supply
                            JOIN alert.supply_sup_area ssa ON ssa.id_supply_soft_inst = ssi.id_supply_soft_inst
                            JOIN alert.supply_type st ON st.id_supply_type = s.id_supply_type
                            JOIN translation tst ON tst.code_translation = st.code_supply_type
                            JOIN translation ts ON ts.code_translation = s.code_supply
                            JOIN alert.supply_loc_default sl ON sl.id_supply_soft_inst = ssi.id_supply_soft_inst
                            WHERE ssi.id_institution = " & i_institution & "
                            AND s.flg_available = 'Y'
                            AND ssi.id_software IN (0, " & i_software & ")
                            AND ssa.id_supply_area = " & i_sup_area & "                   
                            AND ssa.flg_available = 'Y'
                            AND st.flg_available = 'Y'
                            AND tst.desc_lang_" & l_id_language & " IS NOT NULL
                            AND ts.desc_lang_" & l_id_language & " IS NOT NULL
                            AND ST.ID_CONTENT='" & i_id_content_cat & "'"

            If i_flg_type <> "ALL" Then

                sql = sql & " And s.flg_type = '" & i_flg_type & "'"

            End If

            sql = sql & " ORDER BY 4 asc , 2 ASC"

        End If

        Dim cmd As New OracleCommand(sql, i_conn)
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

    Function GET_DISTINCT_SUPPLY_TYPE(ByVal i_a_supplies() As supplies_default, ByVal i_conn As OracleConnection, ByRef i_dr As OracleDataReader) As Boolean

        Dim sql As String = "Select dst.id_content from alert_default.supply_type dst
                             where dst.id_content in ("

        For i As Integer = 0 To i_a_supplies.Count() - 1

            If i < i_a_supplies.Count() - 1 Then

                sql = sql & "'" & i_a_supplies(i).id_content_category & "',"

            Else

                sql = sql & "'" & i_a_supplies(i).id_content_category & "')"

            End If

        Next

        Dim cmd As New OracleCommand(sql, i_conn)
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

    Function SET_SUPPLY_TYPE(ByVal i_institution As Int64, ByVal i_a_supplies() As supplies_default, ByVal i_conn As OracleConnection) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution, i_conn)

        Dim dr_filtered_st As OracleDataReader

        If Not GET_DISTINCT_SUPPLY_TYPE(i_a_supplies, i_conn, dr_filtered_st) Then

            Return False

        End If

        Dim l_a_filtered_supplies(0) As supplies_default
        Dim l_dimension_fs As Integer = 0

        Dim sql As String = "DECLARE

                              l_a_id_content_st  table_varchar := table_varchar("

        While dr_filtered_st.Read()

            'Código para depois se inserir as trduções
            ReDim Preserve l_a_filtered_supplies(l_dimension_fs)
            l_a_filtered_supplies(l_dimension_fs).id_content_category = dr_filtered_st.Item(0)
            l_dimension_fs = l_dimension_fs + 1

            'SQL para a inserção dos SUPPLY TYPES
            sql = sql & " '" & dr_filtered_st.Item(0) & "', "

        End While

        While dr_filtered_st.Read()

            sql = sql & " '" & dr_filtered_st.Item(0) & "', "

        End While

        dr_filtered_st.Dispose()
        dr_filtered_st.Close()

        'Garantir que o array é fechado
        sql = sql & "'');"

        sql = sql & "         l_st_aux ALERT.SUPPLY_TYPE.ID_CONTENT%TYPE;
  
                            BEGIN
  
                              FOR i IN 1 .. l_a_id_content_st.COUNT() LOOP
     
                                 BEGIN
        
                                     SELECT st.id_content
                                     INTO l_st_aux
                                     FROM alert.supply_type st
                                     WHERE st.id_content = l_a_id_content_st(i)
                                     AND st.flg_available = 'Y';
     
                                 EXCEPTION
                                     WHEN TOO_MANY_ROWS THEN
                                       CONTINUE;
         
                                     WHEN NO_DATA_FOUND THEN
                                       insert into ALERT.SUPPLY_TYPE (ID_SUPPLY_TYPE, CODE_SUPPLY_TYPE, ID_CONTENT, FLG_AVAILABLE)
                                       values (ALERT.SEQ_SUPPLY_TYPE.NEXTVAL, 'SUPPLY_TYPE.CODE_SUPPLY_TYPE.' || ALERT.SEQ_SUPPLY_TYPE.NEXTVAL, l_a_id_content_st(i), 'Y');
                                       CONTINUE;     
                                 END;     

                              END LOOP;
  
                            END;"

        Dim cmd_insert_st As New OracleCommand(sql, i_conn)

        Try
            cmd_insert_st.CommandType = CommandType.Text
            cmd_insert_st.ExecuteNonQuery()
        Catch ex As Exception
            cmd_insert_st.Dispose()
            Return False
        End Try

        cmd_insert_st.Dispose()

        If Not SET_SUPPLY_TYPE_TRANSLATION(i_institution, l_a_filtered_supplies, i_conn) Then

            Return False

        End If

        Return True

    End Function


    Function SET_SUPPLY_TYPE_TRANSLATION(ByVal i_institution As Int64, ByVal i_a_supplies() As supplies_default, ByVal i_conn As OracleConnection) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution, i_conn)

        Dim sql As String = "DECLARE

                                l_a_supply_types table_varchar := table_varchar("


        For i As Integer = 0 To i_a_supplies.Count() - 1

            If (i < i_a_supplies.Count() - 1) Then

                sql = sql & "'" & i_a_supplies(i).id_content_category & "', "

            Else

                sql = sql & "'" & i_a_supplies(i).id_content_category & "');"

            End If

        Next

        sql = sql & "
                                l_st_desc TRANSLATION.DESC_LANG_1%TYPE;
                                l_st_code table_varchar := table_varchar();

                            BEGIN

                                FOR i IN 1 .. l_a_supply_types.count()
                                LOOP
                                    BEGIN
        
                                        SELECT st.code_supply_type BULK COLLECT
                                        INTO l_st_code
                                        FROM alert.supply_type st
                                        WHERE st.id_content = l_a_supply_types(i)
                                        AND st.flg_available = 'Y'
                                        AND pk_translation.get_translation(" & l_id_language & ", st.code_supply_type) IS NULL;
        
                                        FOR ii IN 1 .. l_st_code.count()
                                        LOOP
            
                                            IF l_st_code(ii) IS NOT NULL
                                            THEN
                
                                                SELECT alert_default.pk_translation_default.get_translation_default(" & l_id_language & ", dst.code_supply_type)
                                                INTO l_st_desc
                                                FROM alert_default.Supply_Type DST
                                                WHERE DST.id_content = l_a_supply_types(i)
                                                AND dST.Flg_Available = 'Y';
                                
                                                pk_translation.insert_into_translation(" & l_id_language & ", l_st_code(II), l_st_desc);
                
                                            END IF;

                                        END LOOP;
        
                                    EXCEPTION
                                        WHEN OTHERS THEN
                                            continue;
                                    END;
        
                                END LOOP;

                            END;"

        Dim cmd_insert_st_translation As New OracleCommand(sql, i_conn)

        Try
            cmd_insert_st_translation.CommandType = CommandType.Text
            cmd_insert_st_translation.ExecuteNonQuery()
        Catch ex As Exception
            cmd_insert_st_translation.Dispose()
            Return False
        End Try

        cmd_insert_st_translation.Dispose()

        Return True

    End Function

    'Esta Função faz Insert de novos supplies, e insert de traduções para supplies que não tinham tradução
    Function SET_SUPPLY(ByVal i_institution As Int64, ByVal i_a_supplies() As supplies_default, ByVal i_conn As OracleConnection) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution, i_conn)

        Dim sql As String = "DECLARE

                                l_a_supply_types table_varchar := table_varchar("

        For i As Integer = 0 To i_a_supplies.Count() - 1

            If (i < i_a_supplies.Count() - 1) Then

                sql = sql & "'" & i_a_supplies(i).id_content_category & "', "

            Else

                sql = sql & "'" & i_a_supplies(i).id_content_category & "');"

            End If

        Next

        sql = sql & "l_a_supplies     table_varchar := table_varchar("

        For i As Integer = 0 To i_a_supplies.Count() - 1

            If (i < i_a_supplies.Count() - 1) Then

                sql = sql & "'" & i_a_supplies(i).id_content_supply & "', "

            Else

                sql = sql & "'" & i_a_supplies(i).id_content_supply & "');"

            End If

        Next

        sql = sql & "       l_a_id_supliy_type table_varchar := table_varchar();

                            l_id_supply   alert.supply.id_supply%TYPE;
                            l_desc_supply translation.desc_lang_1%TYPE;

                            l_flg_type alert.supply.flg_type%TYPE;

                            l_id_supply_verifif table_number := table_number();

                        BEGIN

                            FOR i IN 1 .. l_a_supplies.count()
                            LOOP
    
                                SELECT st.id_supply_type BULK COLLECT
                                INTO l_a_id_supliy_type
                                FROM alert.supply_type st
                                WHERE st.id_content = l_a_supply_types(i)
                                AND st.flg_available = 'Y';
    
                                SELECT ds.flg_type
                                INTO l_flg_type
                                FROM alert_default.supply ds
                                JOIN alert_default.supply_type dst ON dst.id_supply_type = ds.id_supply_type
                                WHERE ds.id_content = l_a_supplies(i)
                                AND dst.id_content = l_a_supply_types(i)
                                AND dst.flg_available = 'Y'
                                AND ds.flg_available = 'Y';
    
                                BEGIN
        
                                    SELECT s.id_supply BULK COLLECT
                                    INTO l_id_supply_verifif
                                    FROM alert.supply s
                                    JOIN ALERT.SUPPLY_TYPE ST ON ST.ID_SUPPLY_TYPE=S.ID_SUPPLY_TYPE
                                    WHERE s.id_content = l_a_supplies(i)
                                    AND s.flg_available = 'Y'            
                                    AND ST.FLG_AVAILABLE='Y';
        
                                    IF l_id_supply_verifif.count() = 0
                                    THEN
            
                                        l_id_supply := alert.seq_supply.nextval;
            
                                        INSERT INTO alert.supply
                                            (id_supply, code_supply, id_supply_type, flg_type, id_content, flg_available)
                                        VALUES
                                            (l_id_supply, 'SUPPLY.CODE_SUPPLY.' || l_id_supply, l_a_id_supliy_type(1), l_flg_type, l_a_supplies(i), 'Y');
            
                                        SELECT alert_default.pk_translation_default.get_translation_default(6, ds.code_supply)
                                        INTO l_desc_supply
                                        FROM alert_default.supply ds
                                        WHERE ds.id_content = l_a_supplies(i)
                                        AND ds.flg_available = 'Y';
            
                                        pk_translation.insert_into_translation(" & l_id_language & ", 'SUPPLY.CODE_SUPPLY.' || l_id_supply, l_desc_supply);
            
                                    ELSE
            
                                        --VERIFICAR SE TEM TRADUÇÃO (ATENÇÃO AOS DUPLICADOS)
                                        SELECT s.id_supply BULK COLLECT
                                        INTO l_id_supply_verifif
                                        FROM alert.supply s
                                        JOIN translation t ON t.code_translation = s.code_supply
                                        WHERE s.id_content = l_a_supplies(i)
                                        AND s.flg_available = 'Y'
                                        AND t.desc_lang_" & l_id_language & " IS NULL;
            
                                        IF l_id_supply_verifif.count() > 0
                                        THEN
                
                                            DECLARE
                    
                                                l_desc_supply translation.desc_lang_1%TYPE;
                    
                                            BEGIN
                    
                                                FOR j IN 1 .. l_id_supply_verifif.count()
                                                LOOP
                        
                                                    SELECT alert_default.pk_translation_default.get_translation_default(" & l_id_language & ", ds.code_supply)
                                                    INTO l_desc_supply
                                                    FROM alert_default.supply ds
                                                    JOIN alert.supply s ON s.id_content = ds.id_content
                                                    WHERE s.id_supply = l_id_supply_verifif(j)
                                                    AND ds.flg_available = 'Y';
                        
                                                    pk_translation.insert_into_translation(" & l_id_language & ", 'SUPPLY.CODE_SUPPLY.' || l_id_supply_verifif(j), l_desc_supply);
                        
                                                END LOOP;
                    
                                            END;
                
                                        END IF;
            
                                    END IF;
                                END;
    
                            END LOOP;

                        END;"

        Dim cmd_insert_st_translation As New OracleCommand(sql, i_conn)

        Try
            cmd_insert_st_translation.CommandType = CommandType.Text
            cmd_insert_st_translation.ExecuteNonQuery()
        Catch ex As Exception
            cmd_insert_st_translation.Dispose()
            Return False
        End Try

        cmd_insert_st_translation.Dispose()

        Return True

    End Function

    Function SET_SUPPLY_SOFT_INST(ByVal i_institution As Int64, ByVal i_software As Int16, ByVal i_a_supplies() As supplies_default, ByVal i_conn As OracleConnection) As Boolean

        Dim sql As String = "DECLARE

                                 l_a_supplies table_varchar := table_varchar("

        For i As Integer = 0 To i_a_supplies.Count() - 1

            If (i < i_a_supplies.Count() - 1) Then

                sql = sql & "'" & i_a_supplies(i).id_content_supply & "', "

            Else

                sql = sql & "'" & i_a_supplies(i).id_content_supply & "');"

            End If

        Next

        sql = sql & "       l_id_supply table_varchar := table_varchar(); --Como existem supplies repetidos é necessário um table_varchar

                            l_id_supply_soft_inst  alert.supply_soft_inst.id_supply_soft_inst%TYPE;
                            l_quantity             alert_default.supply_soft_inst.quantity%TYPE;
                            l_id_unit_measure      alert_default.supply_soft_inst.id_unit_measure%TYPE;
                            l_flg_cons_type        alert_default.supply_soft_inst.flg_cons_type%TYPE;
                            l_flg_reusable         alert_default.supply_soft_inst.flg_reusable%TYPE;
                            l_flg_editable         alert_default.supply_soft_inst.flg_editable%TYPE;
                            l_total_avail_quantity alert_default.supply_soft_inst.total_avail_quantity%TYPE;
                            l_flg_preparing        alert_default.supply_soft_inst.flg_preparing%TYPE;
                            l_flg_countable        alert_default.supply_soft_inst.flg_countable%TYPE;


                        BEGIN

                            FOR i IN 1 .. l_a_supplies.count()
                            LOOP

                                SELECT dssi.quantity,
                                       dssi.id_unit_measure,
                                       dssi.flg_cons_type,
                                       dssi.flg_reusable,
                                       dssi.flg_editable,
                                       dssi.total_avail_quantity,
                                       dssi.flg_preparing,
                                       dssi.flg_countable
                                INTO l_quantity,
                                     l_id_unit_measure,
                                     l_flg_cons_type,
                                     l_flg_reusable,
                                     l_flg_editable,
                                     l_total_avail_quantity,
                                     l_flg_preparing,
                                     l_flg_countable
                                FROM alert_default.supply_soft_inst dssi
                                JOIN alert_default.supply ds ON ds.id_supply = dssi.id_supply
        
                                WHERE ds.id_content = l_a_supplies(i)
                                AND ds.flg_available = 'Y';
    
                                SELECT s.id_supply BULK COLLECT
                                INTO l_id_supply
                                FROM alert.supply s
                                WHERE s.id_content = l_a_supplies(i)
                                AND s.flg_available = 'Y';
    
                                FOR j IN 1 .. l_id_supply.count()
                                LOOP
        
                                    l_id_supply_soft_inst := alert.seq_supply_soft_inst.nextval;
        
                                    --INSERIR EM supply_soft_inst
                                    BEGIN
                                        INSERT INTO alert.supply_soft_inst
                                            (id_supply_soft_inst,
                                             id_supply,
                                             id_institution,
                                             id_software,
                                             id_professional,
                                             id_dept,
                                             quantity,
                                             id_unit_measure,
                                             flg_cons_type,
                                             flg_reusable,
                                             flg_editable,
                                             total_avail_quantity,
                                             flg_preparing,
                                             flg_countable)
                                        VALUES
                                            (l_id_supply_soft_inst,
                                             l_id_supply(j),
                                             " & i_institution & ",
                                             " & i_software & ",
                                             0,
                                             NULL,
                                             l_quantity,
                                             l_id_unit_measure,
                                             l_flg_cons_type,
                                             l_flg_reusable,
                                             l_flg_editable,
                                             l_total_avail_quantity,
                                             l_flg_preparing,
                                             l_flg_countable);
                                    EXCEPTION
                                        WHEN dup_val_on_index THEN
                                            continue;
                                    END;
               
                                END LOOP;
    
                            END LOOP;

                        END;"

        Dim cmd_insert_supply_soft As New OracleCommand(sql, i_conn)

        Try
            cmd_insert_supply_soft.CommandType = CommandType.Text
            cmd_insert_supply_soft.ExecuteNonQuery()
        Catch ex As Exception
            cmd_insert_supply_soft.Dispose()
            Return False
        End Try

        cmd_insert_supply_soft.Dispose()

        Return True

    End Function

    Function SET_SUPPLY_SUP_AREA(ByVal i_institution As Int64, ByVal i_software As Int16, ByVal i_sup_area As Int16, ByVal i_a_supplies() As supplies_default, ByVal i_conn As OracleConnection) As Boolean

        Dim sql As String = "DECLARE

                                 l_a_supplies table_varchar := table_varchar("

        For i As Integer = 0 To i_a_supplies.Count() - 1

            If (i < i_a_supplies.Count() - 1) Then

                sql = sql & "'" & i_a_supplies(i).id_content_supply & "', "

            Else

                sql = sql & "'" & i_a_supplies(i).id_content_supply & "');"

            End If

        Next

        sql = sql & "       l_id_supply table_varchar := table_varchar(); --Como existem supplies repetidos é necessário um table_varchar

                            l_id_sup_area alert_default.supply_sup_area.id_supply_area%TYPE;

                            l_a_supply_soft_inst table_number := table_number();

                        BEGIN

                            FOR i IN 1 .. l_a_supplies.count()
                            LOOP
    
                                SELECT ssi.id_supply_soft_inst 
                                BULK COLLECT
                                INTO l_a_supply_soft_inst
                                FROM alert.supply_soft_inst ssi
                                JOIN alert.supply s ON s.id_supply = ssi.id_supply
                                                AND s.flg_available = 'Y'
                                WHERE ssi.id_institution = " & i_institution & "
                                AND ssi.id_software = " & i_software & "
                                AND s.id_content = l_a_supplies(i);
    
                                FOR j IN 1 .. l_a_supply_soft_inst.count()
                                LOOP
                
                                    --INSERIR EM supply_sup_area
                                    BEGIN
                                        INSERT INTO alert.supply_sup_area
                                            (id_supply_area, id_supply_soft_inst, flg_available)
                                        VALUES
                                            (" & i_sup_area & ", l_a_supply_soft_inst(j), 'Y');
                                    EXCEPTION
                                        WHEN dup_val_on_index THEN
                                            continue;
                                    END;
        
                                END LOOP;
    
                            END LOOP;

                        END;"

        Dim cmd_insert_supply_sup_area As New OracleCommand(sql, i_conn)

        Try
            cmd_insert_supply_sup_area.CommandType = CommandType.Text
            cmd_insert_supply_sup_area.ExecuteNonQuery()
        Catch ex As Exception
            cmd_insert_supply_sup_area.Dispose()
            Return False
        End Try

        cmd_insert_supply_sup_area.Dispose()

        Return True

    End Function

    Function SET_SUPPLY_LOC_DEFAULT(ByVal i_institution As Int64, ByVal i_software As Int16, ByVal i_a_supplies() As supplies_default, ByVal i_conn As OracleConnection) As Boolean

        Dim sql As String = "DECLARE

                                 l_a_supplies table_varchar := table_varchar("

        For i As Integer = 0 To i_a_supplies.Count() - 1

            If (i < i_a_supplies.Count() - 1) Then

                sql = sql & "'" & i_a_supplies(i).id_content_supply & "', "

            Else

                sql = sql & "'" & i_a_supplies(i).id_content_supply & "');"

            End If

        Next

        sql = sql & "
                        l_id_supply table_varchar := table_varchar(); --Como existem supplies repetidos é necessário um table_varchar

                        l_a_supply_soft_inst table_number := table_number();

                        l_default_supply_soft_inst alert_default.supply_soft_inst.id_supply_soft_inst%TYPE;

                        FUNCTION record_exists
                        (
                            l_id_supply_location  IN alert.supply_location.id_supply_location%TYPE,
                            l_id_supply_soft_inst IN alert.supply_loc_default.id_supply_soft_inst%TYPE
                        ) RETURN BOOLEAN IS
    
                            l_count_rep INTEGER := 0;
    
                        BEGIN
    
                            SELECT COUNT(1)
                            INTO l_count_rep
                            FROM alert.supply_loc_default sld
                            WHERE sld.id_supply_location = l_id_supply_location
                            AND sld.id_supply_soft_inst = l_id_supply_soft_inst;
    
                            IF l_count_rep > 0
                            THEN
        
                                RETURN TRUE;
        
                            ELSE
        
                                RETURN FALSE;
        
                            END IF;
    
                        END record_exists;

                    BEGIN

                        FOR i IN 1 .. l_a_supplies.count()
                        LOOP
    
                            DECLARE
        
                                l_id_supply_location table_number := table_number();
                                l_flg_default        table_varchar := table_varchar();
        
                            BEGIN
        
                                SELECT ssi.id_supply_soft_inst BULK COLLECT
                                INTO l_a_supply_soft_inst
                                FROM alert.supply_soft_inst ssi
                                JOIN alert.supply s ON s.id_supply = ssi.id_supply
                                                AND s.flg_available = 'Y'
                                WHERE ssi.id_institution = " & i_institution & "
                                AND ssi.id_software = " & i_software & "
                                AND s.id_content = l_a_supplies(i);
        
                                SELECT dsld.id_supply_location, dsld.flg_default BULK COLLECT
                                INTO l_id_supply_location, l_flg_default
                                FROM alert_default.supply_loc_default dsld
                                JOIN alert_default.supply_soft_inst ssi ON ssi.id_supply_soft_inst = dsld.id_supply_soft_inst
                                JOIN alert_default.supply ds ON ds.id_supply = ssi.id_supply
                                WHERE ds.id_content = l_a_supplies(i)
                                AND ds.flg_available = 'Y';
        
                                FOR j IN 1 .. l_a_supply_soft_inst.count()
                                LOOP
                                    BEGIN
                                        FOR jj IN 1 .. l_id_supply_location.count()
                                        LOOP
                    
                                            BEGIN
                        
                                                IF NOT record_exists(l_id_supply_location(jj), l_a_supply_soft_inst(j))
                                                THEN
                                                        
                                                    INSERT INTO alert.supply_loc_default
                                                        (id_supply_location, id_supply_loc_default, id_supply_soft_inst, flg_default)
                                                    VALUES
                                                        (l_id_supply_location(jj), alert.seq_supply_loc_default.nextval, l_a_supply_soft_inst(j), l_flg_default(jj));
                            
                                                END IF;
                        
                                            EXCEPTION
                                                WHEN dup_val_on_index THEN
                                                    continue;
                            
                                            END;
                    
                                        END LOOP;
                
                                    EXCEPTION
                                        WHEN dup_val_on_index THEN
                                            continue;
                                    END;
            
                                END LOOP;
                            END;
                        END LOOP;
                    END;"

        Dim cmd_insert_supply_loc_default As New OracleCommand(sql, i_conn)

        Try
            cmd_insert_supply_loc_default.CommandType = CommandType.Text
            cmd_insert_supply_loc_default.ExecuteNonQuery()
        Catch ex As Exception
            cmd_insert_supply_loc_default.Dispose()
            Return False
        End Try

        cmd_insert_supply_loc_default.Dispose()

        Return True

    End Function

    Function DELETE_SUPPLY_SOFT_INST(ByVal i_institution As Int64, ByVal i_software As Integer, ByVal i_a_supplies() As supplies_default, ByVal i_sup_area As Int16, ByVal i_conn As OracleConnection) As Boolean

        Dim sql As String = "DECLARE

                                l_a_supply_types table_varchar := table_varchar("

        For i As Integer = 0 To i_a_supplies.Count() - 1

            If (i < i_a_supplies.Count() - 1) Then

                sql = sql & "'" & i_a_supplies(i).id_content_category & "', "

            Else

                sql = sql & "'" & i_a_supplies(i).id_content_category & "');"

            End If

        Next

        sql = sql & "l_a_supplies     table_varchar := table_varchar("

        For i As Integer = 0 To i_a_supplies.Count() - 1

            If (i < i_a_supplies.Count() - 1) Then

                sql = sql & "'" & i_a_supplies(i).id_content_supply & "', "

            Else

                sql = sql & "'" & i_a_supplies(i).id_content_supply & "');"

            End If

        Next

        sql = sql & "
                             l_id_supply_soft_inst table_number := table_number();

                          FUNCTION RECORD_EXISTS
                          (
                             ID_SUPPLY_SOFT_INST     IN  alert.supply_soft_inst.id_supply_soft_inst%type
    
                           ) RETURN BOOLEAN IS
    
                            l_count_rep        integer := 0;
    
                           BEGIN
           
                               Select count(1)
                                 into l_count_rep
                                 from ALERT.SUPPLY_SUP_AREA SSA
                                where SSA.ID_SUPPLY_SOFT_INST=ID_SUPPLY_SOFT_INST;
                             
                               IF l_count_rep > 0 then
            
                                         RETURN TRUE;
                 
                               ELSE
             
                                         RETURN FALSE;
                     
                               END IF;
    
                           END RECORD_EXISTS;

                        BEGIN

                            FOR i IN 1 .. l_a_supplies.count()
                            LOOP
    
                                SELECT ssi.id_supply_soft_inst BULK COLLECT
                                INTO l_id_supply_soft_inst
                                FROM alert.supply_soft_inst ssi
                                WHERE ssi.id_supply IN (
                                
                                                        SELECT s.id_supply
                                                        FROM alert.supply s
                                                        JOIN alert.supply_type st ON st.id_supply_type = s.id_supply_type
                                                                              AND st.flg_available = 'Y'
                                                        WHERE s.id_content = l_a_supplies(i)
                                                        AND st.id_content = l_a_supply_types(i))
                                AND ssi.id_institution = " & i_institution & "
                                AND ssi.id_software IN (0, " & i_software & ");
    
                                DELETE FROM alert.supply_sup_area ssa
                                WHERE ssa.id_supply_soft_inst MEMBER OF(l_id_supply_soft_inst)
                                AND ssa.id_supply_area = " & i_sup_area & ";
        
                                --Verificar se o registo exista para mais que uma área
                                --Se não exisitr, apaga na LOC_DEFAULT e SOFT_INST
                               FOR j IN 1 .. l_id_supply_soft_inst.count()
                                LOOP
        
                                    IF NOT record_exists(l_id_supply_soft_inst(j))
                                    THEN
            
                                        --1 Apagar de LOC_DEFAULT
            
                                        DELETE FROM alert.supply_loc_default sld
                                        WHERE sld.id_supply_soft_inst IN (l_id_supply_soft_inst(j));
            
                                        --2 Apagar de SOFT_INST
            
                                        DELETE FROM alert.supply_soft_inst ssi
                                        WHERE ssi.id_supply_soft_inst IN (l_id_supply_soft_inst(j));
            
                                    END IF;
        
                                END LOOP;
    
                            END LOOP;

                        END;"

        Dim cmd_supply_sup_area As New OracleCommand(sql, i_conn)

        ' Try
        cmd_supply_sup_area.CommandType = CommandType.Text
            cmd_supply_sup_area.ExecuteNonQuery()
            'Catch ex As Exception
            '   cmd_supply_sup_area.Dispose()
            '  Return False
            'End Try

            cmd_supply_sup_area.Dispose()
        Return True

    End Function

End Class
