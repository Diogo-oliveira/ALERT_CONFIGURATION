﻿Imports Oracle.DataAccess.Client
Public Class Medication_API

    Dim db_access_general As New General

    Public Structure ROUTES
        Public id_route As String
        Public desc_route As String
    End Structure

    Public Structure MED_SET_INSTRUCTIONS
        Public id_product As String
        Public id_std_presc_dir As Int64
        Public rank As Int64
        Public id_grant As Int64
        Public market As Int16
        Public market_desc As String
        Public software As Int16
        Public software_desc As String
        Public id_pick_list As Int16
        Public institution As Int64
    End Structure

    Function GET_LIST_PRODUCTS(ByVal i_institution As Int64, ByVal i_product_supplier As String, ByVal i_product_desc As String, ByRef i_dr As OracleDataReader) As Boolean

        DEBUGGER.SET_DEBUG("MEDICATION_API :: GET_LIST_PRODUCTS(" & i_institution & ", " & i_product_supplier & ", " & i_product_desc & ")")

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution)

        Dim sql As String = "SELECT ed.desc_lang_" & l_id_language & " AS prod_desc, p.id_product
                              FROM alert_product_mt.product p
                              JOIN alert_product_mt.product_medication pm
                                ON pm.id_product = p.id_product
                               AND pm.id_product_supplier = p.id_product_supplier
                              JOIN alert_product_mt.entity_description ed
                                ON ed.code_entity_description = p.code_product
                             WHERE p.id_product_supplier = '" & i_product_supplier & "'
                               AND upper(ed.desc_lang_" & l_id_language & ") LIKE upper('%" & i_product_desc & "%')
                             ORDER BY prod_desc ASC"

        Dim cmd As New OracleCommand(sql, Connection.conn)
        Try
            cmd.CommandType = CommandType.Text
            i_dr = cmd.ExecuteReader()
            cmd.Dispose()
            Return True
        Catch ex As Exception

            DEBUGGER.SET_DEBUG_ERROR_INIT("MEDICATION_API :: MEDICATION_API")
            DEBUGGER.SET_DEBUG(ex.Message)
            DEBUGGER.SET_DEBUG(sql)
            DEBUGGER.SET_DEBUG_ERROR_CLOSE()

            cmd.Dispose()
            Return False
        End Try
    End Function

    Function GET_PRODUCT_SUPPLIER(ByVal i_ID_INST As Int64) As String

        DEBUGGER.SET_DEBUG("MEDICATION :: GET_PRODUCT_SUPPLIER(" & i_ID_INST & ")")

        Dim l_id_product_supplier As String = ""

        Dim sql As String = " SELECT smi.id_supplier
                              FROM alert_product_mt.supplier_mkt_inst smi
                              JOIN market m
                                ON m.id_market = smi.id_market
                              JOIN institution i
                                ON i.id_market = m.id_market
                             WHERE smi.flg_available = 'Y'
                               AND i.id_institution = " & i_ID_INST & " and rownum=1"

        Dim cmd As New OracleCommand(sql, Connection.conn)
        cmd.CommandType = CommandType.Text

        Dim dr As OracleDataReader

        Try

            dr = cmd.ExecuteReader()

            While dr.Read()

                l_id_product_supplier = dr.Item(0)

            End While

            dr.Dispose()
            dr.Close()
            cmd.Dispose()

        Catch ex As Exception

            DEBUGGER.SET_DEBUG_ERROR_INIT("MEDICATION :: GET_PRODUCT_SUPPLIER")
            DEBUGGER.SET_DEBUG(ex.Message)
            DEBUGGER.SET_DEBUG(sql)
            DEBUGGER.SET_DEBUG_ERROR_CLOSE()

            dr.Dispose()
            dr.Close()
            cmd.Dispose()

        End Try

        Return l_id_product_supplier

    End Function

    Function GET_PRODUCT_OPTIONS(ByVal i_institution As Int64, ByVal i_software As Int16, ByVal i_id_product As String, ByVal i_product_supplier As String, ByVal i_id_pick_list As Int16, ByRef i_dr As OracleDataReader) As Boolean

        DEBUGGER.SET_DEBUG("MEDICATION_API :: GET_PRODUCT_OPTIONS(" & i_institution & ", " & i_software & ", " & i_id_product & ", " & i_product_supplier & ")")

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution)
        Dim l_market As Int16 = db_access_general.GET_INSTITUTION_MARKET(i_institution)

        Dim sql As String = "SELECT p.id_product,
                               nvl2(lpl.id_product, decode(p.flg_available,'Y','Y','N','N'), 'N') as available,
                               p.id_product_level,
                               decode(pm.id_product_med_type,2,'IV',1,'Non-IV') as MED_TYPE,
                               pm.flg_mix_with_fluid,
                               pm.flg_justify_expensive,
                               pm.flg_controlled_drug,
                               pm.flg_blood_derivative,
                               pm.flg_dopant,
                               pm.flg_narcotic,
                               eds.desc_lang_" & l_id_language & " as product_synonym
                          FROM alert_product_mt.product p
                          JOIN alert_product_mt.product_medication pm
                            ON pm.id_product = p.id_product
                           AND pm.id_product_supplier = p.id_product_supplier
                          LEFT JOIN alert_product_mt.lnk_product_synonym lps
                            ON lps.id_product = p.id_product
                           AND lps.id_product_supplier = p.id_product_supplier
                           AND lps.id_grant IN (SELECT g.id_grant
                                                  FROM alert_product_mt.v_cfg_grant g
                                                 WHERE g.market = " & l_market & "
                                                   AND g.institution = " & i_institution & "
                                                   AND g.software = " & i_software & "
                                                   AND g.ID_CONTEXT is null)
                          LEFT JOIN alert_product_mt.entity_description eds
                            ON eds.code_entity_description = lps.code_synonym
                          LEFT JOIN alert_product_mt.lnk_product_pick_list lpl
                            ON lpl.id_product = p.id_product
                           AND lpl.id_product_supplier = p.id_product_supplier
                           AND lpl.id_pick_list = " & i_id_pick_list & "
                         WHERE p.id_product_supplier = '" & i_product_supplier & "'
                           AND p.id_product = '" & i_id_product & "'"

        Dim cmd As New OracleCommand(sql, Connection.conn)
        Try
            cmd.CommandType = CommandType.Text
            i_dr = cmd.ExecuteReader()
            cmd.Dispose()
            Return True
        Catch ex As Exception

            DEBUGGER.SET_DEBUG_ERROR_INIT("MEDICATION_API :: GET_PRODUCT_OPTIONS")
            DEBUGGER.SET_DEBUG(ex.Message)
            DEBUGGER.SET_DEBUG(sql)
            DEBUGGER.SET_DEBUG_ERROR_CLOSE()

            cmd.Dispose()
            Return False
        End Try
    End Function

    Function SET_PARAMETERS(ByVal i_institution As Int64, ByVal i_software As Int16, ByVal i_id_product As String, ByVal i_id_product_supplier As String, i_flg_available As String, i_id_pick_list As Int16, i_id_product_level As Int16,
                            ByVal i_med_type As Int16, ByVal i_mix_fluid As String, ByVal i_justify_expensive As String, i_controlled_drug As String,
                            ByVal i_blood_derivate As String, ByVal i_dopant As String, ByVal i_narcotic As String,
                            ByVal i_product_synonym As String) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution)

        Dim sql As String = "UPDATE ALERT_PRODUCT_MT.PRODUCT P
                                 SET P.FLG_AVAILABLE='" & i_flg_available & "', p.id_product_level=" & i_id_product_level & "
                               WHERE P.ID_PRODUCT='" & i_id_product & "'
                                 AND P.ID_PRODUCT_SUPPLIER='" & i_id_product_supplier & "'"

        Dim cmd_update_product As New OracleCommand(sql, Connection.conn)

        Try
            cmd_update_product.CommandType = CommandType.Text
            cmd_update_product.ExecuteNonQuery()
        Catch ex As Exception
            cmd_update_product.Dispose()
            Return False
        End Try

        cmd_update_product.Dispose()

        If Not SET_LNK_PRODUCT_PICK_LIST(i_institution, i_id_product, i_id_product_supplier, i_id_pick_list) Then
            Return False
        End If

        sql = "UPDATE ALERT_PRODUCT_MT.PRODUCT_MEDICATION P
                                 SET P.FLG_MIX_WITH_FLUID = '" & i_mix_fluid & "',
                                   P.FLG_JUSTIFY_EXPENSIVE = '" & i_justify_expensive & "',
                                   P.FLG_CONTROLLED_DRUG = '" & i_controlled_drug & "',
                                   P.FLG_BLOOD_DERIVATIVE = '" & i_blood_derivate & "',
                                   P.FLG_DOPANT = '" & i_dopant & "',
                                   P.FLG_NARCOTIC = '" & i_narcotic & "',
                                   P.id_product_med_type = '" & i_med_type & "'
                               WHERE P.ID_PRODUCT='" & i_id_product & "'
                                 AND P.ID_PRODUCT_SUPPLIER='" & i_id_product_supplier & "'"

        Dim cmd_update_product_medication As New OracleCommand(sql, Connection.conn)

        Try
            cmd_update_product_medication.CommandType = CommandType.Text
            cmd_update_product_medication.ExecuteNonQuery()
        Catch ex As Exception
            cmd_update_product_medication.Dispose()
            Return False
        End Try

        If i_product_synonym <> "" Then

            If Not SET_PRODUCT_SYNONYM(i_institution, i_software, i_id_product, i_id_product_supplier, i_id_pick_list, i_product_synonym) Then

                Return False

            End If
        Else

            If Not DELETE_PRODUCT_SYNONYM(i_id_product, i_id_product_supplier, i_id_pick_list) Then

                Return False
            End If

        End If

        Return True

    End Function

    Function SET_LNK_PRODUCT_PICK_LIST(ByVal i_institution As Int64, ByVal i_id_product As String, ByVal i_id_product_supplier As String, i_id_pick_list As Int16) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution)

        Dim sql As String = "BEGIN
                                INSERT INTO alert_product_mt.lnk_product_pick_list
                                    (id_product, id_product_supplier, id_pick_list, rank)
                                VALUES
                                    ('" & i_id_product & "', '" & i_id_product_supplier & "', " & i_id_pick_list & ", 10);
                            EXCEPTION
                                WHEN dup_val_on_index THEN
                                    dbms_output.put_line('Repeated record');
                            END;"

        Dim cmd_insert_pl As New OracleCommand(sql, Connection.conn)

        Try
            cmd_insert_pl.CommandType = CommandType.Text
            cmd_insert_pl.ExecuteNonQuery()
        Catch ex As Exception
            cmd_insert_pl.Dispose()
            Return False
        End Try

        cmd_insert_pl.Dispose()

        Return True

    End Function

    Function SET_PRODUCT_SYNONYM(ByVal i_institution As Int64, ByVal i_software As Int16, ByVal i_id_product As String, ByVal i_id_product_supplier As String, i_id_pick_list As Int16, ByVal i_synonym As String) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution)
        Dim l_id_market As Int16 = db_access_general.GET_INSTITUTION_MARKET(i_institution)

        Dim sql As String = "DECLARE
                                l_grant NUMBER(24);
                            BEGIN

                                BEGIN
                                    SELECT g.id_grant
                                      INTO l_grant
                                      FROM alert_product_mt.v_cfg_grant g
                                     WHERE g.market = " & l_id_market & "
                                       AND g.institution = " & i_institution & "
                                       AND g.software = " & i_software & "
                                       AND g.id_context IS NULL
                                       AND rownum = 1;
                                EXCEPTION
                                    WHEN no_data_found THEN
                                        l_grant := alert_product_mt.pk_grants.set_by_soft_inst(i_context     => '',
                                                                                               i_prof        => profissional(0, " & i_institution & ",  " & i_software & "),
                                                                                               i_market      => " & l_id_market & ",
                                                                                               i_grant_order => 1);
                                END;

                                    INSERT INTO alert_product_mt.lnk_product_synonym
                                        (id_product, id_product_supplier, code_synonym, id_grant, id_pick_list)
                                    VALUES
                                        ('" & i_id_product & "', '" & i_id_product_supplier & "', 'LNK_PRODUCT_SYNONYM.CODE_SYNONYM." & i_id_product_supplier & "." & i_id_product & "', l_grant , " & i_id_pick_list & ");

                                    alert_product_mt.pk_product_utils.insert_into_entity_desc(" & l_id_language & ",
                                                                                              'LNK_PRODUCT_SYNONYM.CODE_SYNONYM." & i_id_product_supplier & "." & i_id_product & "',
                                                                                              NULL,
                                                                                              '" & i_synonym & "');
                                    pk_lucene_index_admin.sync_specific_index ('ALERT_PRODUCT_MT','ENTITY_DESCRIPTION'," & l_id_language & ");

                                EXCEPTION
                                    WHEN dup_val_on_index THEN
                                    alert_product_mt.pk_product_utils.insert_into_entity_desc(" & l_id_language & ",
                                                                                              'LNK_PRODUCT_SYNONYM.CODE_SYNONYM." & i_id_product_supplier & "." & i_id_product & "',
                                                                                              NULL,
                                                                                              '" & i_synonym & "');
                                    pk_lucene_index_admin.sync_specific_index ('ALERT_PRODUCT_MT','ENTITY_DESCRIPTION'," & l_id_language & ");
    
                                END;"

        Dim cmd_insert_syn As New OracleCommand(sql, Connection.conn)

        Try
            cmd_insert_syn.CommandType = CommandType.Text
            cmd_insert_syn.ExecuteNonQuery()
        Catch ex As Exception
            cmd_insert_syn.Dispose()
            Return False
        End Try

        cmd_insert_syn.Dispose()

        Return True

    End Function

    Function DELETE_PRODUCT_SYNONYM(ByVal i_id_product As String, ByVal i_id_product_supplier As String, ByVal i_id_pick_list As Int16) As Boolean

        Dim sql As String = "DELETE FROM alert_product_mt.lnk_product_synonym lps
                         WHERE lps.id_product_supplier = '" & i_id_product_supplier & "'
                           AND lps.id_product = '" & i_id_product & "'
                           AND lps.id_pick_list = " & i_id_pick_list

        Dim cmd_delete_syn As New OracleCommand(sql, Connection.conn)

        Try
            cmd_delete_syn.CommandType = CommandType.Text
            cmd_delete_syn.ExecuteNonQuery()
        Catch ex As Exception
            cmd_delete_syn.Dispose()
            Return False
        End Try

        cmd_delete_syn.Dispose()

        Return True

    End Function

    Function GET_PRODUCT_ROUTES(ByVal i_institution As Int64, ByVal i_product As String, ByVal i_product_supplier As String, ByRef i_dr As OracleDataReader) As Boolean

        DEBUGGER.SET_DEBUG("MEDICATION_API :: GET_PRODUCT_ROUTES(" & i_institution & ", " & i_product & ", " & i_product_supplier & ")")

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution)

        Dim sql As String = "SELECT r.id_route, ed.desc_lang_" & l_id_language & " AS desc_route, r.flg_default
                              FROM alert_product_mt.lnk_product_mkt_route r
                              JOIN alert_product_mt.route ro
                                ON ro.id_route = r.id_route
                               AND ro.id_route_supplier = r.id_product_supplier
                              JOIN alert_product_mt.entity_description ed
                                ON ed.code_entity_description = ro.code_route
                             WHERE r.id_product = '" & i_product & "'
                               AND r.id_product_supplier = '" & i_product_supplier & "'
                             ORDER BY desc_route ASC"

        Dim cmd As New OracleCommand(sql, Connection.conn)
        Try
            cmd.CommandType = CommandType.Text
            i_dr = cmd.ExecuteReader()
            cmd.Dispose()
            Return True
        Catch ex As Exception

            DEBUGGER.SET_DEBUG_ERROR_INIT("MEDICATION_API :: GET_PRODUCT_ROUTES")
            DEBUGGER.SET_DEBUG(ex.Message)
            DEBUGGER.SET_DEBUG(sql)
            DEBUGGER.SET_DEBUG_ERROR_CLOSE()

            cmd.Dispose()
            Return False
        End Try
    End Function

    Function GET_MARKET_ROUTES(ByVal i_institution As Int64, ByVal i_product_supplier As String, ByRef i_dr As OracleDataReader) As Boolean

        DEBUGGER.SET_DEBUG("MEDICATION_API :: GET_PRODUCT_ROUTES(" & i_institution & ", " & i_product_supplier & ")")

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution)

        Dim sql As String = "SELECT r.id_route, ed.desc_lang_" & l_id_language & " AS desc_route
                              FROM alert_product_mt.route r
                              JOIN alert_product_mt.entity_description ed
                                ON ed.code_entity_description = r.code_route
                             WHERE r.id_route_supplier = '" & i_product_supplier & "'
                             ORDER BY desc_route ASC"

        Dim cmd As New OracleCommand(sql, Connection.conn)
        Try
            cmd.CommandType = CommandType.Text
            i_dr = cmd.ExecuteReader()
            cmd.Dispose()
            Return True
        Catch ex As Exception

            DEBUGGER.SET_DEBUG_ERROR_INIT("MEDICATION_API :: GET_PRODUCT_ROUTES")
            DEBUGGER.SET_DEBUG(ex.Message)
            DEBUGGER.SET_DEBUG(sql)
            DEBUGGER.SET_DEBUG_ERROR_CLOSE()

            cmd.Dispose()
            Return False
        End Try
    End Function

    Function SET_PRODUCT_ROUTES(ByVal i_institution As Int64, ByVal i_id_product As String, ByVal i_id_product_supplier As String, ByVal i_routes() As String) As Boolean

        Dim l_id_market As Int16 = db_access_general.GET_INSTITUTION_MARKET(i_institution)

        Dim sql As String = " DECLARE
                                    l_id_routes table_varchar := table_varchar("

        For i As Integer = 0 To i_routes.Count - 1

            If i < i_routes.Count - 1 Then
                sql = sql & "'" & i_routes(i) & "'" & ", "
            Else
                sql = sql & "'" & i_routes(i) & "'"
            End If


        Next
        sql = sql & ");
                                BEGIN

                                    DELETE FROM alert_product_mt.lnk_product_mkt_route R
                                     WHERE R.ID_PRODUCT = '" & i_id_product & "' 
                                      AND R.ID_PRODUCT_SUPPLIER = '" & i_id_product_supplier & "'
                                      AND R.ID_MARKET = " & l_id_market & "
                                      AND R.ID_ROUTE NOT IN (SELECT * FROM TABLE(l_id_routes));

                                    FOR i IN 1 .. l_id_routes.count
                                    LOOP
    
                                        BEGIN        
                                            insert into alert_product_mt.lnk_product_mkt_route (ID_PRODUCT, ID_PRODUCT_SUPPLIER, ID_MARKET, ID_ROUTE, ID_ROUTE_SUPPLIER, ID_ROUTE_STATUS, FLG_DEFAULT, RANK, FLG_STD)
                                                                          values ('" & i_id_product & "', '" & i_id_product_supplier & "', " & l_id_market & ", l_id_routes(i), '" & i_id_product_supplier & "', 1, 'N', 1, 'U');
                                                    EXCEPTION
                                                        WHEN dup_val_on_index THEN
                                                            CONTINUE;
                                                    END;
                                    END LOOP;                                               
                                END;"

        Dim cmd_insert_routes As New OracleCommand(sql, Connection.conn)

        Try
            cmd_insert_routes.CommandType = CommandType.Text
            cmd_insert_routes.ExecuteNonQuery()
        Catch ex As Exception
            cmd_insert_routes.Dispose()
            Return False
        End Try

        cmd_insert_routes.Dispose()

        Return True

    End Function

    Function SET_ROUTE_DEFAULT(ByVal i_institution As Int64, ByVal i_id_product As String, ByVal i_id_product_supplier As String, ByVal i_routes As String) As Boolean

        Dim l_id_market As Int16 = db_access_general.GET_INSTITUTION_MARKET(i_institution)

        Dim sql As String = "BEGIN

                                UPDATE alert_product_mt.lnk_product_mkt_route l
                                   SET l.flg_default = 'N'
                                 WHERE l.id_product = '" & i_id_product & "'
                                   AND l.id_product_supplier = '" & i_id_product_supplier & "'
                                   AND l.id_market = " & l_id_market & ";

                                UPDATE alert_product_mt.lnk_product_mkt_route l
                                   SET l.flg_default = 'Y'
                                 WHERE l.id_product = '" & i_id_product & "'
                                   AND l.id_product_supplier = '" & i_id_product_supplier & "'
                                   AND l.id_market = " & l_id_market & "
                                   AND l.id_route = '" & i_routes & "';
                            END;"

        Dim cmd_route_default As New OracleCommand(sql, Connection.conn)

        Try
            cmd_route_default.CommandType = CommandType.Text
            cmd_route_default.ExecuteNonQuery()
        Catch ex As Exception
            cmd_route_default.Dispose()
            MsgBox(sql)
            Return False
        End Try

        cmd_route_default.Dispose()

        Return True

    End Function

    Function GET_MARKET_UM(ByVal i_institution As Int64, ByVal i_search_string As String, ByRef i_dr As OracleDataReader) As Boolean

        DEBUGGER.SET_DEBUG("MEDICATION_API :: GET_MARKET_UM(" & i_institution & ", " & i_search_string & ")")

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution)

        Dim sql As String = "SELECT u.id_unit_measure, pk_translation.get_translation(" & l_id_language & ", u.code_unit_measure) AS um_desc
                              FROM alert.unit_measure u
                             WHERE u.flg_available = 'Y'
                               AND pk_translation.get_translation(" & l_id_language & ", u.code_unit_measure) IS NOT NULL "

        If i_search_string = "" Then

            sql = sql & "                               
                             ORDER BY um_desc ASC"

        Else
            sql = sql & "
                               AND upper(pk_translation.get_translation(" & l_id_language & ", u.code_unit_measure)) LIKE upper('%" & i_search_string & "%')
                             ORDER BY um_desc ASC"
        End If


        Dim cmd As New OracleCommand(sql, Connection.conn)
        Try
            cmd.CommandType = CommandType.Text
            i_dr = cmd.ExecuteReader()
            cmd.Dispose()
            Return True
        Catch ex As Exception

            DEBUGGER.SET_DEBUG_ERROR_INIT("MEDICATION_API :: GET_MARKET_UM")
            DEBUGGER.SET_DEBUG(ex.Message)
            DEBUGGER.SET_DEBUG(sql)
            DEBUGGER.SET_DEBUG_ERROR_CLOSE()

            cmd.Dispose()
            Return False
        End Try
    End Function

    Function SET_PRODUCT_UM(ByVal i_institution As Int64, ByVal i_id_product As String, ByVal i_id_product_supplier As String, ByVal i_context As Int16, ByVal i_um() As Int64) As Boolean

        Dim sql As String = " DECLARE
                                   l_tbl_um table_number := table_number("

        For i As Integer = 0 To i_um.Count - 1

            If i < i_um.Count - 1 Then
                sql = sql & i_um(i) & ", "
            Else
                sql = sql & i_um(i)
            End If


        Next
        sql = sql & ");
                                BEGIN

                                        FOR i IN 1 .. l_tbl_um.count
                                        LOOP
                                            BEGIN
                                            INSERT INTO alert_product_mt.lnk_product_unit_measure
                                                (id_product, id_product_supplier, id_unit_measure, id_unit_measure_context, flg_default)
                                            VALUES
                                                ('" & i_id_product & "', '" & i_id_product_supplier & "', l_tbl_um(i), " & i_context & ", 'N');
                                            EXCEPTION
                                                WHEN DUP_VAL_ON_INDEX THEN
                                                      CONTINUE;
                                            END;
                                        END LOOP;
                                    END;"

        Dim cmd_insert_um As New OracleCommand(sql, Connection.conn)

        Try
            cmd_insert_um.CommandType = CommandType.Text
            cmd_insert_um.ExecuteNonQuery()
        Catch ex As Exception
            cmd_insert_um.Dispose()
            Return False
        End Try

        cmd_insert_um.Dispose()

        Return True

    End Function

    Function GET_PRODUCT_UM(ByVal i_institution As Int64, ByVal i_id_product As String, ByVal i_id_product_supplier As String, ByVal i_context As Int16, ByRef i_dr As OracleDataReader) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution)

        Dim sql As String = "       SELECT lum.id_unit_measure, pk_translation.get_translation(" & l_id_language & ", um.code_unit_measure) AS PROD_DESC, lum.flg_default
                                      FROM alert_product_mt.lnk_product_unit_measure lum
                                      JOIN alert.unit_measure um
                                        On um.id_unit_measure = lum.id_unit_measure
                                     WHERE lum.id_product = '" & i_id_product & "'
                                       And lum.id_product_supplier = '" & i_id_product_supplier & "'
                                       AND lum.id_unit_measure_context = " & i_context & "
                                       ORDER BY PROD_DESC ASC"

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

    Function DELETE_PRODUCT_UM(ByVal i_institution As Int64, ByVal i_id_product As String, ByVal i_id_product_supplier As String, ByVal i_context As Int16, ByVal i_unit_measure() As Int64) As Boolean

        Dim sql As String = " DECLARE
                                   l_tbl_um table_number := table_number("

        For i As Integer = 0 To i_unit_measure.Count - 1
            If i < i_unit_measure.Count - 1 Then
                sql = sql & i_unit_measure(i) & ", "
            Else
                sql = sql & i_unit_measure(i)
            End If
        Next
        sql = sql & ");
                                BEGIN
                                        FOR i IN 1 .. l_tbl_um.count
                                        LOOP
                                            DELETE FROM alert_product_mt.lnk_product_unit_measure lum
                                             WHERE lum.id_product = '" & i_id_product & "'
                                               AND lum.id_product_supplier = '" & i_id_product_supplier & "'
                                               AND lum.id_unit_measure_context = " & i_context & "
                                               AND lum.id_unit_measure = l_tbl_um(i);
                                        END LOOP;
                                    END;"

        Dim cmd_delete_um As New OracleCommand(sql, Connection.conn)

        Try
            cmd_delete_um.CommandType = CommandType.Text
            cmd_delete_um.ExecuteNonQuery()
        Catch ex As Exception
            cmd_delete_um.Dispose()
            Return False
        End Try

        cmd_delete_um.Dispose()

        Return True

    End Function

    Function SET_DEFAULT_UM(ByVal i_institution As Int64, ByVal i_id_product As String, ByVal i_id_product_supplier As String, ByVal i_context As Int16, ByVal i_unit_measure As Int64) As Boolean

        Dim sql As String = " BEGIN

                                UPDATE alert_product_mt.lnk_product_unit_measure lum
                                SET lum.flg_default='N'
                                 WHERE lum.id_product = '" & i_id_product & "'
                                   AND lum.id_product_supplier = '" & i_id_product_supplier & "'
                                   AND lum.id_unit_measure_context = " & i_context & ";

   
                                UPDATE alert_product_mt.lnk_product_unit_measure lum
                                SET lum.flg_default='Y'
                                 WHERE lum.id_product = '" & i_id_product & "'
                                   AND lum.id_product_supplier = '" & i_id_product_supplier & "'
                                   AND lum.id_unit_measure_context = " & i_context & "
                                   AND LUM.ID_UNIT_MEASURE = " & i_unit_measure & ";

                                END;  "

        Dim cmd_default_um As New OracleCommand(sql, Connection.conn)

        Try
            cmd_default_um.CommandType = CommandType.Text
            cmd_default_um.ExecuteNonQuery()
        Catch ex As Exception
            cmd_default_um.Dispose()
            Return False
        End Try

        cmd_default_um.Dispose()

        Return True

    End Function

    Function GET_ALL_FREQS(ByVal i_institution As Int64, ByVal i_software As Int16, ByRef i_dr As OracleDataReader) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution)
        Dim l_market As Int16 = db_access_general.GET_INSTITUTION_MARKET(i_institution)

        Dim sql As String = "SELECT v3.data AS data, v3.label AS label, v3.type AS TYPE, v3.rank AS rank
                              FROM (
                                    --frequecies from presc_list
                                    SELECT g.data AS data, g.label AS label, g.type AS TYPE, g.display_rank AS rank
                                      FROM TABLE(alert_product_tr.pk_product.get_presc_list_pipelined(i_lang                => " & l_id_language & ",
                                                                                                       i_session_data        => alert_product_mt.session_data(NULL,
                                                                                                                                             " & i_institution & ",
                                                                                                                                             " & i_software & ",
                                                                                                                                             " & l_market & "),
                                                                                                       i_id_description_type => 1,
                                                                                                       i_id_presc_list       => 3)) g
        
                                    UNION ALL
                                    --frequencies from presc_dir_frequency
                                    SELECT v2.id_presc_dir_frequency AS data, v2.frequency_desc AS label, 'V' AS TYPE, v2.rank AS rank
                                      FROM (SELECT /*+OPT_ESTIMATE(TABLE t ROWS=1)*/
                                              t.id_presc_dir_frequency,
                                              nvl(t.frequency_synonym,
                                                  alert_product_tr.pk_api_med_core_in.get_translation(i_lang      => " & l_id_language & ",
                                                                                                      i_code_mess => t.code_presc_dir_frequency)) frequency_desc,
                                              t.display_rank rank,
                                              t.flg_freq_type,
                                              t.flg_default,
                                              t.num_days,
                                              t.flg_prn,
                                              t.flg_normal,
                                              t.flg_iv
                                               FROM TABLE(alert_product_mt.pk_product_med.tf_get_frequencies_attributes(i_lang               => " & l_id_language & ",
                                                                                                                        i_session_data       => alert_product_mt.session_data(NULL,
                                                                                                                                                             " & i_institution & ",
                                                                                                                                                             " & i_software & ",
                                                                                                                                                             " & l_market & "),
                                                                                                                        i_tab_frequency_type => table_varchar('PDSTD'))) t) v2) v3
                             WHERE v3.type <> 'AI'
                               AND label IS NOT NULL
                             ORDER BY rank NULLS LAST, label, data"

        Dim cmd As New OracleCommand(sql, Connection.conn)
        ' Try
        cmd.CommandType = CommandType.Text
            i_dr = cmd.ExecuteReader()
            cmd.Dispose()
            Return True
        '  Catch ex As Exception

        ' DEBUGGER.SET_DEBUG_ERROR_INIT("MEDICATION_API :: GET_PRODUCT_OPTIONS")
        ' DEBUGGER.SET_DEBUG(ex.Message)
        ' DEBUGGER.SET_DEBUG(sql)
        ' DEBUGGER.SET_DEBUG_ERROR_CLOSE()

        'cmd.Dispose()
        'Return False
        ' End Try
    End Function

    Function GET_PRODUCT_DESC(ByVal i_institution As Int64, ByVal i_id_product As String, ByVal i_id_product_supplier As String) As String

        'DEBUGGER.SET_DEBUG("MEDICATION :: GET_PRODUCT_SUPPLIER(" & i_ID_INST & ")")

        Dim l_product_desc As String = ""
        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution)

        Dim sql As String = "   SELECT ed.desc_lang_" & l_id_language & "
                                  FROM alert_product_mt.product p
                                  JOIN alert_product_mt.entity_description ed
                                    ON ed.code_entity_description = p.code_product
                                 WHERE p.id_product = '" & i_id_product & "'
                                   AND p.id_product_supplier = '" & i_id_product_supplier & "'"

        Dim cmd As New OracleCommand(sql, Connection.conn)
        cmd.CommandType = CommandType.Text

        Dim dr As OracleDataReader

        Try
            dr = cmd.ExecuteReader()
            While dr.Read()
                l_product_desc = dr.Item(0)
            End While

            dr.Dispose()
            dr.Close()
            cmd.Dispose()

        Catch ex As Exception

            'DEBUGGER.SET_DEBUG_ERROR_INIT("MEDICATION :: GET_PRODUCT_SUPPLIER")
            'DEBUGGER.SET_DEBUG(ex.Message)
            'DEBUGGER.SET_DEBUG(sql)
            'DEBUGGER.SET_DEBUG_ERROR_CLOSE()

            dr.Dispose()
            dr.Close()
            cmd.Dispose()

        End Try

        Return l_product_desc

    End Function

    Function GET_ALL_INSTRUCTIONS(ByVal i_institution As Int64, ByVal i_software As Int64, ByVal i_id_product As String, ByVal i_id_product_supplier As String, ByVal i_pick_list As Int16, ByRef i_dr As OracleDataReader) As Boolean

        'DEBUGGER.SET_DEBUG("MEDICATION :: GET_PRODUCT_SUPPLIER(" & i_ID_INST & ")")

        Dim l_market As Int16 = db_access_general.GET_INSTITUTION_MARKET(i_institution)

        Dim sql As String = "   SELECT lsd.id_product,
                                       d.id_std_presc_directions,
                                       lsd.rank,
                                       lsd.id_grant,
                                       cfg.market,
                                       (SELECT (cgd.value || ' - ' || m.desc_market)
                                          FROM alert_product_mt.config_grant cg
                                          JOIN alert_product_mt.config_grant_det cgd
                                            ON cgd.id_grant = cg.id_grant
                                          JOIN alert_product_mt.config_universe cu
                                            ON cu.id_universe = cgd.id_universe
                                          JOIN market m
                                            ON m.id_market = cgd.value
                                         WHERE cg.id_grant = lsd.id_grant
                                           AND cu.column_name = 'market') AS MARKET_DESC,
                                       cfg.software,
                                       (SELECT (cgd.value || ' - ' || (decode(cgd.value, 0, 'ALL', s.name)))
                                          FROM alert_product_mt.config_grant cg
                                          JOIN alert_product_mt.config_grant_det cgd
                                            ON cgd.id_grant = cg.id_grant
                                          JOIN alert_product_mt.config_universe cu
                                            ON cu.id_universe = cgd.id_universe
                                          JOIN software s
                                            ON s.id_software = cgd.value
                                         WHERE cg.id_grant = lsd.id_grant
                                           AND cu.column_name = 'software') AS SOFTWARE_DESC,
                                       lsd.id_pick_list,
                                       (SELECT cgd.value
                                          FROM alert_product_mt.config_grant cg
                                          JOIN alert_product_mt.config_grant_det cgd
                                            ON cgd.id_grant = cg.id_grant
                                          JOIN alert_product_mt.config_universe cu
                                            ON cu.id_universe = cgd.id_universe
                                         WHERE cg.id_grant = lsd.id_grant
                                           AND cu.column_name = 'institution') AS INSTITUTION
                                  FROM alert_product_mt.std_presc_dir d
                                  JOIN alert_product_mt.lnk_product_std_presc_dir lsd
                                    ON lsd.id_std_presc_directions = d.id_std_presc_directions
                                  JOIN alert_product_mt.v_cfg_grant cfg
                                    ON cfg.id_grant = lsd.id_grant
                                 WHERE lsd.id_product IN ('" & i_id_product & "')
                                   AND lsd.id_product_supplier = '" & i_id_product_supplier & "'
                                   AND cfg.id_context = 'LNK_PRODUCT_STD_PRESC_DIR'
                                   AND cfg.institution IN (0, " & i_institution & ")
                                   AND cfg.market IN (0, " & l_market & ")
                                   AND cfg.software IN (0, " & i_software & ")
                                   AND lsd.id_pick_list IN (0, " & i_pick_list & ")
                                 ORDER BY cfg.market DESC, cfg.institution DESC, cfg.software DESC, lsd.id_pick_list DESC, lsd.rank ASC  "

        Dim cmd As New OracleCommand(sql, Connection.conn)
        Try
            cmd.CommandType = CommandType.Text
            i_dr = cmd.ExecuteReader()
            cmd.Dispose()
            Return True
        Catch ex As Exception

            '    DEBUGGER.SET_DEBUG_ERROR_INIT("MEDICATION_API :: MEDICATION_API")
            '   DEBUGGER.SET_DEBUG(ex.Message)
            '  DEBUGGER.SET_DEBUG(sql)
            ' DEBUGGER.SET_DEBUG_ERROR_CLOSE()

            cmd.Dispose()
            Return False
        End Try
    End Function

    Function GET_STD_PRESC_DIR(ByVal i_institution As Int64, ByVal i_id_std_presc_dir As Int64, ByRef i_dr As OracleDataReader) As Boolean

        'DEBUGGER.SET_DEBUG("MEDICATION :: GET_PRODUCT_SUPPLIER(" & i_ID_INST & ")")
        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution)
        Dim sql As String = "   SELECT d.id_std_presc_directions,
                                       d.flg_sos,
                                       d.id_sos_take_condition,
                                       d.sos_take_condition,
                                       ed_s.desc_lang_" & l_id_language & " AS admin_site,
                                       ed_am.desc_lang_" & l_id_language & " AS admin_method,
                                       d.notes,
                                       to_char(d.patient_instr_desc) AS patient_instructions
                                  FROM alert_product_mt.std_presc_dir d
                                  LEFT JOIN alert_product_mt.admin_method am
                                    ON am.id_admin_method = d.id_admin_method
                                  LEFT JOIN translation ed_am
                                    ON ed_am.code_translation = am.code_admin_method
                                  LEFT JOIN alert_product_mt.admin_site a
                                    ON a.id_admin_site = d.id_admin_site
                                  LEFT JOIN translation ed_s
                                    ON ed_s.code_translation = a.code_admin_site
                                 WHERE d.id_std_presc_directions =  " & i_id_std_presc_dir

        Dim cmd As New OracleCommand(sql, Connection.conn)
        Try
            cmd.CommandType = CommandType.Text
            i_dr = cmd.ExecuteReader()
            cmd.Dispose()
            Return True
        Catch ex As Exception

            '    DEBUGGER.SET_DEBUG_ERROR_INIT("MEDICATION_API :: MEDICATION_API")
            '   DEBUGGER.SET_DEBUG(ex.Message)
            '  DEBUGGER.SET_DEBUG(sql)
            ' DEBUGGER.SET_DEBUG_ERROR_CLOSE()

            cmd.Dispose()
            Return False
        End Try

    End Function
End Class

