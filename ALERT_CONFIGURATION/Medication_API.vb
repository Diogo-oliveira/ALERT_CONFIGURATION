Imports Oracle.DataAccess.Client
Public Class Medication_API

    Dim db_access_general As New General

    Public Structure ROUTES
        Public id_route As String
        Public desc_route As String
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

    Function SET_PARAMETERS(ByVal i_institution As Int64, ByVal i_id_product As String, ByVal i_id_product_supplier As String, i_flg_available As String, i_id_pick_list As Int16, i_id_product_level As Int16,
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

            If Not SET_PRODUCT_SYNONYM(i_institution, i_id_product, i_id_product_supplier, i_id_pick_list, i_product_synonym) Then

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

    Function SET_PRODUCT_SYNONYM(ByVal i_institution As Int64, ByVal i_id_product As String, ByVal i_id_product_supplier As String, i_id_pick_list As Int16, ByVal i_synonym As String) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution)

        Dim sql As String = "DECLARE
                                l_grant NUMBER(24);
                            BEGIN

                                BEGIN
                                    SELECT g.id_grant
                                      INTO l_grant
                                      FROM alert_product_mt.v_cfg_grant g
                                     WHERE g.market = 12
                                       AND g.institution = 2945
                                       AND g.software = 11
                                       AND g.id_context IS NULL
                                       AND rownum = 1;
                                EXCEPTION
                                    WHEN no_data_found THEN
                                        l_grant := alert_product_mt.pk_grants.set_by_soft_inst(i_context     => '',
                                                                                               i_prof        => profissional(0, 2945, 11),
                                                                                               i_market      => 12,
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
            MsgBox(sql)
            Return False
        End Try

        cmd_insert_routes.Dispose()

        Return True

    End Function


End Class


