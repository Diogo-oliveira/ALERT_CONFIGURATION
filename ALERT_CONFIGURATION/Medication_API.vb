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

    Public Structure MED_SET_DOSES
        Public dose_value_1 As Int64
        Public id_unit_dose_1 As Int64
        Public desc_unit_dose_1 As String
        Public dose_value_2 As Int64
        Public id_unit_dose_2 As Int64
        Public desc_unit_dose_2 As String
        Public dose_value_3 As Int64
        Public id_unit_dose_3 As Int64
        Public desc_unit_dose_3 As String
        Public dose_value_4 As Int64
        Public id_unit_dose_4 As Int64
        Public desc_unit_dose_4 As String
        Public dose_value_5 As Int64
        Public id_unit_dose_5 As Int64
        Public desc_unit_dose_5 As String
        Public dose_value_6 As Int64
        Public id_unit_dose_6 As Int64
        Public desc_unit_dose_6 As String
        Public dose_value_7 As Int64
        Public id_unit_dose_7 As Int64
        Public desc_unit_dose_7 As String
        Public id_product_component As String
        Public flg_updated As String
    End Structure

    Public Structure UM_INFO
        Public id_unit_measure As Int64
        Public unit_measure_desc As String
    End Structure

    Function GET_LIST_PRODUCTS(ByVal i_institution As Int64, ByVal i_product_supplier As String, ByVal i_product_desc As String, ByRef i_dr As OracleDataReader) As Boolean

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

            Dim l_parameters(1) As String
            l_parameters(0) = "prod_desc"
            l_parameters(1) = "id_product"

            DEBUGGER.SET_DEBUG("MEDICATION_API :: GET_LIST_PRODUCTS(" & i_institution & ", " & i_product_supplier & ", " & i_product_desc & ")")
            DEBUGGER.SET_RESPONSE("MEDICATION_API :: GET_LIST_PRODUCTS", l_parameters, i_dr)
            i_dr = cmd.ExecuteReader()

            cmd.Dispose()
            Return True
        Catch ex As Exception
        DEBUGGER.SET_DEBUG_ERROR_INIT("MEDICATION_API :: GET_LIST_PRODUCTS")
        DEBUGGER.SET_DEBUG(ex.Message)
        DEBUGGER.SET_DEBUG(Sql)
        DEBUGGER.SET_DEBUG_ERROR_CLOSE()
        cmd.Dispose()
        Return False
        End Try
    End Function

    Function GET_LIST_PRODUCTS_BY_ID(ByVal i_institution As Int64, ByVal i_product_supplier As String, ByVal i_id_product As String, ByRef i_dr As OracleDataReader) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution)

        Dim sql As String = "SELECT ed.desc_lang_" & l_id_language & " AS prod_desc, p.id_product
                              FROM alert_product_mt.product p
                              JOIN alert_product_mt.product_medication pm
                                ON pm.id_product = p.id_product
                               AND pm.id_product_supplier = p.id_product_supplier
                              JOIN alert_product_mt.entity_description ed
                                ON ed.code_entity_description = p.code_product
                             WHERE p.id_product_supplier = '" & i_product_supplier & "'
                               AND UPPER(p.id_product) LIKE UPPER('" & i_id_product & "')
                             ORDER BY prod_desc ASC"

        Dim cmd As New OracleCommand(sql, Connection.conn)
        Try
            cmd.CommandType = CommandType.Text
            i_dr = cmd.ExecuteReader()

            Dim l_parameters(1) As String
            l_parameters(0) = "prod_desc"
            l_parameters(1) = "id_product"

            DEBUGGER.SET_DEBUG("MEDICATION_API :: GET_LIST_PRODUCTS_BY_ID(" & i_institution & ", " & i_product_supplier & ", " & i_id_product & ")")
            DEBUGGER.SET_RESPONSE("MEDICATION_API :: GET_LIST_PRODUCTS_BY_ID", l_parameters, i_dr)
            i_dr = cmd.ExecuteReader()

            cmd.Dispose()
            Return True
        Catch ex As Exception
            DEBUGGER.SET_DEBUG_ERROR_INIT("MEDICATION_API :: GET_LIST_PRODUCTS_BY_ID")
            DEBUGGER.SET_DEBUG(ex.Message)
            DEBUGGER.SET_DEBUG(sql)
            DEBUGGER.SET_DEBUG_ERROR_CLOSE()
            cmd.Dispose()
            Return False
        End Try
    End Function

    Function GET_PRODUCT_SUPPLIER(ByVal i_ID_INST As Int64) As String

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

            Dim l_parameters(0) As String
            l_parameters(0) = "id_supplier"

            DEBUGGER.SET_DEBUG("MEDICATION :: GET_PRODUCT_SUPPLIER(" & i_ID_INST & ")")
            DEBUGGER.SET_RESPONSE("MEDICATION_API :: GET_PRODUCT_SUPPLIER", l_parameters, dr)
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
            cmd.Dispose()
        End Try

        Return l_id_product_supplier

    End Function

    Function GET_PRODUCT_MED_TYPE(ByVal i_id_product As String, ByVal i_id_product_supplier As String) As Int16

        Dim l_id_product_supplier As String = ""

        Dim sql As String = "   SELECT decode(pm.id_product_med_type, 2, 2, 1) med_type
                                 FROM alert_product_mt.product_medication pm
                                WHERE pm.id_product = '" & i_id_product & "'
                                  AND pm.id_product_supplier = '" & i_id_product_supplier & "'"

        Dim cmd As New OracleCommand(sql, Connection.conn)
        cmd.CommandType = CommandType.Text

        Dim dr As OracleDataReader

        Try
            dr = cmd.ExecuteReader()

            Dim l_parameters(0) As String
            l_parameters(0) = "med_type"

            DEBUGGER.SET_DEBUG("MEDICATION :: GET_PRODUCT_MED_TYPE(" & i_id_product & ", " & i_id_product_supplier & ")")
            DEBUGGER.SET_RESPONSE("MEDICATION_API :: GET_PRODUCT_MED_TYPE", l_parameters, dr)
            dr = cmd.ExecuteReader()

            While dr.Read()
                l_id_product_supplier = dr.Item(0)
            End While
            dr.Dispose()
            dr.Close()
            cmd.Dispose()
        Catch ex As Exception
            DEBUGGER.SET_DEBUG_ERROR_INIT("MEDICATION :: GET_PRODUCT_MED_TYPE")
            DEBUGGER.SET_DEBUG(ex.Message)
            DEBUGGER.SET_DEBUG(sql)
            DEBUGGER.SET_DEBUG_ERROR_CLOSE()
            cmd.Dispose()
        End Try

        Return l_id_product_supplier

    End Function

    Function GET_PRODUCT_N_STD_INSTRUCTIONS(ByVal i_id_product As String, ByVal i_id_product_supplier As String) As Int64

        Dim l_id_product_supplier As String = ""

        Dim sql As String = " SELECT COUNT(*)
                              FROM alert_product_mt.lnk_product_std_presc_dir d
                             WHERE d.id_product = '" & i_id_product & "'
                               AND d.id_product_supplier = '" & i_id_product_supplier & "'"

        Dim cmd As New OracleCommand(sql, Connection.conn)
        cmd.CommandType = CommandType.Text

        Dim dr As OracleDataReader

        Try
            dr = cmd.ExecuteReader()

            Dim l_parameters(0) As String
            l_parameters(0) = "COUNT"

            DEBUGGER.SET_DEBUG("MEDICATION :: GET_PRODUCT_N_STD_INSTRUCTIONS(" & i_id_product & ", " & i_id_product_supplier & ")")
            DEBUGGER.SET_RESPONSE("MEDICATION_API :: GET_PRODUCT_N_STD_INSTRUCTIONS", l_parameters, dr)
            dr = cmd.ExecuteReader()

            While dr.Read()
                l_id_product_supplier = dr.Item(0)
            End While
            dr.Dispose()
            dr.Close()
            cmd.Dispose()
        Catch ex As Exception
            DEBUGGER.SET_DEBUG_ERROR_INIT("MEDICATION :: GET_PRODUCT_N_STD_INSTRUCTIONS")
            DEBUGGER.SET_DEBUG(ex.Message)
            DEBUGGER.SET_DEBUG(sql)
            DEBUGGER.SET_DEBUG_ERROR_CLOSE()
            cmd.Dispose()
        End Try

        Return l_id_product_supplier

    End Function

    Function GET_PRODUCT_OPTIONS(ByVal i_institution As Int64, ByVal i_software As Int16, ByVal i_id_product As String, ByVal i_product_supplier As String, ByVal i_id_pick_list As Int16, ByRef i_dr As OracleDataReader) As Boolean

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

        Using cmd As New OracleCommand(sql, Connection.conn)

            Try
                cmd.CommandType = CommandType.Text
                Try
                    i_dr.Dispose()
                    i_dr = cmd.ExecuteReader()
                    cmd.Dispose()
                    cmd.Cancel()
                Catch ex As Exception
                    i_dr = cmd.ExecuteReader()
                    cmd.Dispose()
                    cmd.Cancel()
                End Try

                Dim l_parameters(10) As String
                l_parameters(0) = "id_product"
                l_parameters(1) = "available"
                l_parameters(2) = "id_product_level"
                l_parameters(3) = "MED_TYPE"
                l_parameters(4) = "flg_mix_with_fluid"
                l_parameters(5) = "flg_justify_expensive"
                l_parameters(6) = "flg_controlled_drug"
                l_parameters(7) = "flg_blood_derivative"
                l_parameters(8) = "flg_dopant"
                l_parameters(9) = "flg_narcotic"
                l_parameters(10) = "product_synonym"

                DEBUGGER.SET_DEBUG("MEDICATION_API :: GET_PRODUCT_OPTIONS(" & i_institution & ", " & i_software & ", " & i_id_product & ", " & i_product_supplier & ")")
                DEBUGGER.SET_RESPONSE("MEDICATION_API :: GET_PRODUCT_OPTIONS", l_parameters, i_dr)
                Try
                    i_dr.Dispose()
                    i_dr = cmd.ExecuteReader()
                    cmd.Dispose()
                    cmd.Cancel()
                Catch ex As Exception

                    i_dr = cmd.ExecuteReader()
                    cmd.Dispose()
                    cmd.Cancel()
                End Try

            Catch ex As Exception

                DEBUGGER.SET_DEBUG_ERROR_INIT("MEDICATION_API :: GET_PRODUCT_OPTIONS")
                DEBUGGER.SET_DEBUG(ex.Message)
                DEBUGGER.SET_DEBUG(sql)
                DEBUGGER.SET_DEBUG_ERROR_CLOSE()

                cmd.Dispose()
                Return False
            End Try

        End Using

        Return True

    End Function

    Function SET_PARAMETERS(ByVal i_institution As Int64, ByVal i_software As Int16, ByVal i_id_product As String, ByVal i_id_product_supplier As String, i_flg_available As String, i_id_pick_list As Int16, i_id_product_level As Int16,
                            ByVal i_med_type As Int16, ByVal i_mix_fluid As String, ByVal i_justify_expensive As String, i_controlled_drug As String,
                            ByVal i_blood_derivate As String, ByVal i_dopant As String, ByVal i_narcotic As String,
                            ByVal i_product_synonym As String) As Boolean

        DEBUGGER.SET_DEBUG("MEDICATION_API :: SET_PARAMETERS(" & i_institution & ", " & i_software & ", " & i_id_product & ", " & i_id_product_supplier &
         ", " & i_flg_available & ", " & i_id_pick_list & ", " & i_id_product_level &
        ", " & i_med_type & ", " & i_mix_fluid & ", " & i_justify_expensive & ", " & i_controlled_drug &
         ", " & i_blood_derivate & ", " & i_dopant & ", " & i_narcotic &
         i_product_synonym & ")")

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
            DEBUGGER.SET_DEBUG_ERROR_INIT("MEDICATION_API :: SET_PARAMETERS")
            DEBUGGER.SET_DEBUG(ex.Message)
            DEBUGGER.SET_DEBUG(sql)
            DEBUGGER.SET_DEBUG_ERROR_CLOSE()
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
            DEBUGGER.SET_DEBUG_ERROR_INIT("MEDICATION_API :: SET_PARAMETERS")
            DEBUGGER.SET_DEBUG(ex.Message)
            DEBUGGER.SET_DEBUG(sql)
            DEBUGGER.SET_DEBUG_ERROR_CLOSE()
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

        DEBUGGER.SET_DEBUG("MEDICATION_API :: SET_LNK_PRODUCT_PICK_LIST(" & i_institution & ", " & i_id_product & ", " & i_id_product_supplier & ", " & i_id_pick_list & ")")

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
            DEBUGGER.SET_DEBUG_ERROR_INIT("MEDICATION_API :: SET_LNK_PRODUCT_PICK_LIST")
            DEBUGGER.SET_DEBUG(ex.Message)
            DEBUGGER.SET_DEBUG(sql)
            DEBUGGER.SET_DEBUG_ERROR_CLOSE()
            cmd_insert_pl.Dispose()
            Return False
        End Try

        cmd_insert_pl.Dispose()

        Return True

    End Function

    Function SET_PRODUCT_SYNONYM(ByVal i_institution As Int64, ByVal i_software As Int16, ByVal i_id_product As String, ByVal i_id_product_supplier As String, i_id_pick_list As Int16, ByVal i_synonym As String) As Boolean

        DEBUGGER.SET_DEBUG("MEDICATION_API :: SET_PRODUCT_SYNONYM(" & i_institution & ", " & i_software & ", " & i_id_product & ", " & i_id_product_supplier & ", " & i_id_pick_list & ", " & i_synonym & ")")

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
            DEBUGGER.SET_DEBUG_ERROR_INIT("MEDICATION_API :: SET_PRODUCT_SYNONYM")
            DEBUGGER.SET_DEBUG(ex.Message)
            DEBUGGER.SET_DEBUG(sql)
            DEBUGGER.SET_DEBUG_ERROR_CLOSE()
            cmd_insert_syn.Dispose()
            Return False
        End Try

        cmd_insert_syn.Dispose()

        Return True

    End Function

    Function DELETE_PRODUCT_SYNONYM(ByVal i_id_product As String, ByVal i_id_product_supplier As String, ByVal i_id_pick_list As Int16) As Boolean

        DEBUGGER.SET_DEBUG("MEDICATION_API :: DELETE_PRODUCT_SYNONYM(" & i_id_product & ", " & i_id_product_supplier & ", " & i_id_pick_list & ")")

        Dim sql As String = "DELETE FROM alert_product_mt.lnk_product_synonym lps
                         WHERE lps.id_product_supplier = '" & i_id_product_supplier & "'
                           AND lps.id_product = '" & i_id_product & "'
                           AND lps.id_pick_list = " & i_id_pick_list

        Dim cmd_delete_syn As New OracleCommand(sql, Connection.conn)

        Try
            cmd_delete_syn.CommandType = CommandType.Text
            cmd_delete_syn.ExecuteNonQuery()
        Catch ex As Exception
            DEBUGGER.SET_DEBUG_ERROR_INIT("MEDICATION_API :: SET_PRODUCT_SYNONYM")
            DEBUGGER.SET_DEBUG(ex.Message)
            DEBUGGER.SET_DEBUG(sql)
            DEBUGGER.SET_DEBUG_ERROR_CLOSE()
            cmd_delete_syn.Dispose()
            Return False
        End Try

        cmd_delete_syn.Dispose()

        Return True

    End Function

    Function GET_PRODUCT_ROUTES(ByVal i_institution As Int64, ByVal i_product As String, ByVal i_product_supplier As String, ByRef i_dr As OracleDataReader) As Boolean

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

        Using cmd As New OracleCommand(sql, Connection.conn)

            Try
                cmd.CommandType = CommandType.Text
                Try
                    i_dr.Dispose()
                    i_dr = cmd.ExecuteReader()
                    cmd.Dispose()
                    cmd.Cancel()
                Catch ex As Exception
                    i_dr = cmd.ExecuteReader()
                    cmd.Dispose()
                    cmd.Cancel()
                End Try


                Dim l_parameters(2) As String
                l_parameters(0) = "id_route"
                l_parameters(1) = "desc_route"
                l_parameters(2) = "flg_default"
                DEBUGGER.SET_DEBUG("MEDICATION_API :: GET_PRODUCT_ROUTES(" & i_institution & ", " & i_product & ", " & i_product_supplier & ")")
                DEBUGGER.SET_RESPONSE("MEDICATION_API :: GET_PRODUCT_ROUTES", l_parameters, i_dr)
                Try
                    i_dr.Dispose()
                    i_dr = cmd.ExecuteReader()
                    cmd.Dispose()
                    cmd.Cancel()
                Catch ex As Exception
                    i_dr = cmd.ExecuteReader()
                    cmd.Dispose()
                    cmd.Cancel()
                End Try
            Catch ex As Exception
                DEBUGGER.SET_DEBUG_ERROR_INIT("MEDICATION_API :: GET_PRODUCT_ROUTES")
                DEBUGGER.SET_DEBUG(ex.Message)
                DEBUGGER.SET_DEBUG(sql)
                DEBUGGER.SET_DEBUG_ERROR_CLOSE()
                cmd.Dispose()
                Return False
            End Try
        End Using

        Return True

    End Function

    Function GET_MARKET_ROUTES(ByVal i_institution As Int64, ByVal i_product_supplier As String, ByRef i_dr As OracleDataReader) As Boolean

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

            DEBUGGER.SET_DEBUG("MEDICATION_API :: GET_MARKET_ROUTES(" & i_institution & ", " & i_product_supplier & ")")
            Dim l_parameters(1) As String
            l_parameters(0) = "id_route"
            l_parameters(1) = "desc_route"

            DEBUGGER.SET_RESPONSE("MEDICATION_API :: GET_MARKET_ROUTES", l_parameters, i_dr)
            i_dr = cmd.ExecuteReader()

            cmd.Dispose()
            Return True
        Catch ex As Exception
            DEBUGGER.SET_DEBUG_ERROR_INIT("MEDICATION_API :: GET_MARKET_ROUTES")
            DEBUGGER.SET_DEBUG(ex.Message)
            DEBUGGER.SET_DEBUG(sql)
            DEBUGGER.SET_DEBUG_ERROR_CLOSE()
            cmd.Dispose()
            Return False
        End Try
    End Function

    Function SET_PRODUCT_ROUTES(ByVal i_institution As Int64, ByVal i_id_product As String, ByVal i_id_product_supplier As String, ByVal i_routes() As String) As Boolean

        Dim l_debug_string As String = "MEDICATION_API :: SET_PRODUCT_ROUTES(" & i_institution & ", " & i_id_product & ", " & i_id_product_supplier & ", "
        l_debug_string = l_debug_string & DEBUGGER.GET_ARRAY_STRING(i_routes)
        l_debug_string = l_debug_string & ")"

        DEBUGGER.SET_DEBUG(l_debug_string)

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
            DEBUGGER.SET_DEBUG_ERROR_INIT("MEDICATION_API :: SET_PRODUCT_ROUTES")
            DEBUGGER.SET_DEBUG(ex.Message)
            DEBUGGER.SET_DEBUG(sql)
            DEBUGGER.SET_DEBUG_ERROR_CLOSE()
            cmd_insert_routes.Dispose()
            Return False
        End Try

        cmd_insert_routes.Dispose()

        Return True

    End Function

    Function SET_ROUTE_DEFAULT(ByVal i_institution As Int64, ByVal i_id_product As String, ByVal i_id_product_supplier As String, ByVal i_routes As String) As Boolean

        DEBUGGER.SET_DEBUG("MEDICATION_API :: SET_ROUTE_DEFAULT(" & i_institution & ", " & i_id_product & ", " & i_id_product_supplier & ", " & i_routes & ")")

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
            DEBUGGER.SET_DEBUG_ERROR_INIT("MEDICATION_API :: SET_ROUTE_DEFAULT")
            DEBUGGER.SET_DEBUG(ex.Message)
            DEBUGGER.SET_DEBUG(sql)
            DEBUGGER.SET_DEBUG_ERROR_CLOSE()
            cmd_route_default.Dispose()
            Return False
        End Try

        cmd_route_default.Dispose()

        Return True

    End Function

    Function GET_MARKET_UM(ByVal i_institution As Int64, ByVal i_id_software As Int64, ByVal i_search_string As String, ByVal i_id_product_supplier As String, ByRef i_dr As OracleDataReader) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution)

        Dim sql As String = "SELECT id_unit_measure, pk_translation.get_translation(" & l_id_language & ", code_unit_measure) AS um_desc
                                  FROM (SELECT v1.id_unit_measure, um.code_unit_measure
                                          FROM (SELECT DISTINCT lpum.id_unit_measure AS id_unit_measure
                                                  FROM alert_product_mt.lnk_product_unit_measure lpum
                                                 WHERE lpum.id_unit_measure_context = 1
                                                   AND lpum.id_product_supplier IN ('" & i_id_product_supplier & "')) v1 
                                          JOIN alert.unit_measure um
                                            ON um.id_unit_measure = v1.id_unit_measure
                                         WHERE um.flg_available = 'Y'
                                           AND um.id_unit_measure_type = 1015
        
                                        UNION ALL
        
                                        SELECT ume.id_unit_measure, um.code_unit_measure
                                          FROM alert_product_mt.lnk_um_supp_exceptions ume
                                          JOIN alert.unit_measure um
                                            ON um.id_unit_measure = ume.id_unit_measure
                                         WHERE ume.flg_available = 'Y'
                                           AND ume.id_unit_measure_context = 1015
                                           AND ume.id_supplier IN ('" & i_id_product_supplier & "')
        )"

        If i_search_string = "" Then

            sql = sql & "                               
                             ORDER BY um_desc ASC"

        Else
            sql = sql & "
                               where upper(pk_translation.get_translation(" & l_id_language & ", code_unit_measure)) LIKE upper('%" & i_search_string & "%')
                             ORDER BY um_desc ASC"
        End If


        Dim cmd As New OracleCommand(sql, Connection.conn)
        Try
            cmd.CommandType = CommandType.Text
            i_dr = cmd.ExecuteReader()

            DEBUGGER.SET_DEBUG("MEDICATION_API :: GET_MARKET_UM(" & i_institution & ", " & i_id_software & ", " & i_search_string & ", " & i_id_product_supplier & ")")
            Dim l_parameters(1) As String
            l_parameters(0) = "id_unit_measure"
            l_parameters(1) = "um_desc"

            DEBUGGER.SET_RESPONSE("MEDICATION_API :: GET_MARKET_UM", l_parameters, i_dr)
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

        Dim l_debug_string As String = "MEDICATION_API :: SET_PRODUCT_UM(" & i_institution & ", " & i_id_product & ", " & i_id_product_supplier & ", " & i_context & ", "

        l_debug_string = l_debug_string & DEBUGGER.GET_ARRAY_NUMBER(i_um) & ")"

        DEBUGGER.SET_DEBUG(l_debug_string)

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
            DEBUGGER.SET_DEBUG_ERROR_INIT("MEDICATION_API :: SET_PRODUCT_UM")
            DEBUGGER.SET_DEBUG(ex.Message)
            DEBUGGER.SET_DEBUG(sql)
            DEBUGGER.SET_DEBUG_ERROR_CLOSE()
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

        Using cmd As New OracleCommand(sql, Connection.conn)

            Try
                cmd.CommandType = CommandType.Text
                Try
                    i_dr.Dispose()
                    i_dr = cmd.ExecuteReader()
                    cmd.Dispose()
                    cmd.Cancel()
                Catch ex As Exception
                    i_dr = cmd.ExecuteReader()
                    cmd.Dispose()
                    cmd.Cancel()
                End Try

                DEBUGGER.SET_DEBUG("MEDICATION_API :: GET_PRODUCT_UM(" & i_institution & ", " & i_id_product & ", " & i_id_product_supplier & ", " & i_context & ")")

                Dim l_parameters(2) As String
                l_parameters(0) = "id_unit_measure"
                l_parameters(1) = "PROD_DESC"
                l_parameters(2) = "flg_default"

                DEBUGGER.SET_RESPONSE("MEDICATION_API :: GET_PRODUCT_UM", l_parameters, i_dr)
                Try
                    i_dr.Dispose()
                    i_dr = cmd.ExecuteReader()
                    cmd.Dispose()
                    cmd.Cancel()
                Catch ex As Exception
                    i_dr = cmd.ExecuteReader()
                    cmd.Dispose()
                    cmd.Cancel()
                End Try

            Catch ex As Exception
                DEBUGGER.SET_DEBUG_ERROR_INIT("MEDICATION_API :: GET_PRODUCT_UM")
                DEBUGGER.SET_DEBUG(ex.Message)
                DEBUGGER.SET_DEBUG(sql)
                DEBUGGER.SET_DEBUG_ERROR_CLOSE()
                cmd.Dispose()
                Return False
            End Try
        End Using

        Return True

    End Function

    Function DELETE_PRODUCT_UM(ByVal i_institution As Int64, ByVal i_id_product As String, ByVal i_id_product_supplier As String, ByVal i_context As Int16, ByVal i_unit_measure() As Int64) As Boolean

        Dim l_debug_string As String = "MEDICATION_API :: DELETE_PRODUCT_UM(" & i_institution & ", " & i_id_product & ", " & i_id_product_supplier & ", " & i_context & ", "

        l_debug_string = l_debug_string & DEBUGGER.GET_ARRAY_NUMBER(i_unit_measure) & ")"

        DEBUGGER.SET_DEBUG(l_debug_string)

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
            DEBUGGER.SET_DEBUG_ERROR_INIT("MEDICATION_API :: DELETE_PRODUCT_UM")
            DEBUGGER.SET_DEBUG(ex.Message)
            DEBUGGER.SET_DEBUG(sql)
            DEBUGGER.SET_DEBUG_ERROR_CLOSE()
            cmd_delete_um.Dispose()
            Return False
        End Try

        cmd_delete_um.Dispose()

        Return True

    End Function

    Function SET_DEFAULT_UM(ByVal i_institution As Int64, ByVal i_id_product As String, ByVal i_id_product_supplier As String, ByVal i_context As Int16, ByVal i_unit_measure As Int64) As Boolean

        DEBUGGER.SET_DEBUG("MEDICATION_API :: SET_DEFAULT_UM(" & i_institution & ", " & i_id_product & ", " & i_id_product_supplier & ", " & i_context & ", " & i_unit_measure & ")")

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
            DEBUGGER.SET_DEBUG_ERROR_INIT("MEDICATION_API :: SET_DEFAULT_UM")
            DEBUGGER.SET_DEBUG(ex.Message)
            DEBUGGER.SET_DEBUG(sql)
            DEBUGGER.SET_DEBUG_ERROR_CLOSE()
            cmd_default_um.Dispose()
            Return False
        End Try

        cmd_default_um.Dispose()

        Return True

    End Function

    Function GET_ALL_FREQS(ByVal i_institution As Int64, ByVal i_software As Int16, ByVal i_prn As String, ByRef i_dr As OracleDataReader) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution)
        Dim l_market As Int16 = db_access_general.GET_INSTITUTION_MARKET(i_institution)

        Dim sql As String = "SELECT v3.data, v3.label AS label, v3.type AS TYPE, v3.rank AS rank
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
                                                                                                                        i_tab_frequency_type => table_varchar('PDSTD'))) t
                                            WHERE t.id_market = " & l_market & " AND (t.flg_prn='" & i_prn & "' or ('" & i_prn & "'= 'N' AND t.flg_prn='A'))) v2) v3
                             WHERE v3.type <> 'AI'
                               AND label IS NOT NULL
                             ORDER BY rank NULLS LAST, label, data"

        Dim cmd As New OracleCommand(sql, Connection.conn)
        Try
            cmd.CommandType = CommandType.Text
            i_dr = cmd.ExecuteReader()

            DEBUGGER.SET_DEBUG("MEDICATION_API :: GET_ALL_FREQS(" & i_institution & ", " & i_software & ", " & i_prn & ")")
            Dim l_parameters(3) As String
            l_parameters(0) = "data"
            l_parameters(1) = "label"
            l_parameters(2) = "TYPE"
            l_parameters(3) = "rank"
            DEBUGGER.SET_RESPONSE("MEDICATION_API :: GET_ALL_FREQS", l_parameters, i_dr)
            i_dr = cmd.ExecuteReader()

            cmd.Dispose()
            Return True
        Catch ex As Exception
            DEBUGGER.SET_DEBUG_ERROR_INIT("MEDICATION_API :: GET_ALL_FREQS")
            DEBUGGER.SET_DEBUG(ex.Message)
            DEBUGGER.SET_DEBUG(sql)
            DEBUGGER.SET_DEBUG_ERROR_CLOSE()

            cmd.Dispose()
            Return False
        End Try
    End Function

    Function GET_PRODUCT_DESC(ByVal i_institution As Int64, ByVal i_id_product As String, ByVal i_id_product_supplier As String) As String

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

            DEBUGGER.SET_DEBUG("MEDICATION_API :: GET_PRODUCT_DESC(" & i_institution & ", " & i_id_product & ", " & i_id_product_supplier & ")")
            Dim l_parameters(0) As String
            l_parameters(0) = "prod_desc"

            DEBUGGER.SET_RESPONSE("MEDICATION_API :: GET_PRODUCT_DESC", l_parameters, dr)
            dr = cmd.ExecuteReader()

            While dr.Read()
                l_product_desc = dr.Item(0)
            End While

            dr.Dispose()
            dr.Close()
            cmd.Dispose()

        Catch ex As Exception
            DEBUGGER.SET_DEBUG_ERROR_INIT("MEDICATION_API :: GET_PRODUCT_DESC")
            DEBUGGER.SET_DEBUG(ex.Message)
            DEBUGGER.SET_DEBUG(sql)
            DEBUGGER.SET_DEBUG_ERROR_CLOSE()
            cmd.Dispose()
        End Try

        Return l_product_desc

    End Function

    Function COUNT_STD_INSTR_PICK_LIST(ByVal i_institution As Int64, ByVal i_software As Int64, ByVal i_id_product As String, ByVal i_id_product_supplier As String, ByVal i_pick_list As Int16) As Int64

        Dim l_market As Int16 = db_access_general.GET_INSTITUTION_MARKET(i_institution)
        Dim L_COUNT As Int64

        Dim sql As String = "SELECT COUNT(*)
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
                               AND lsd.id_pick_list IN (" & i_pick_list & ") "

        Dim cmd As New OracleCommand(sql, Connection.conn)
        cmd.CommandType = CommandType.Text

        Dim dr As OracleDataReader

        Try
            dr = cmd.ExecuteReader()

            DEBUGGER.SET_DEBUG("MEDICATION_API :: COUNT_STD_INSTR_PICK_LIST(" & i_institution & ", " & i_software & ", " & i_id_product & ", " & i_id_product_supplier & ", " & i_pick_list & ")")
            Dim l_parameters(0) As String
            l_parameters(0) = "COUNT"

            DEBUGGER.SET_RESPONSE("MEDICATION_API :: COUNT_STD_INSTR_PICK_LIST", l_parameters, dr)
            dr = cmd.ExecuteReader()

            While dr.Read()
                Try
                    L_COUNT = dr.Item(0)
                Catch ex As Exception
                    L_COUNT = 0
                End Try
            End While

            dr.Dispose()
            dr.Close()
            cmd.Dispose()

        Catch ex As Exception
            DEBUGGER.SET_DEBUG_ERROR_INIT("MEDICATION_API :: COUNT_STD_INSTR_PICK_LIST")
            DEBUGGER.SET_DEBUG(ex.Message)
            DEBUGGER.SET_DEBUG(sql)
            DEBUGGER.SET_DEBUG_ERROR_CLOSE()
            cmd.Dispose()
        End Try

        Return L_COUNT
    End Function

    Function GET_ALL_INSTRUCTIONS(ByVal i_institution As Int64, ByVal i_software As Int64, ByVal i_id_product As String, ByVal i_id_product_supplier As String, ByVal i_pick_list As Int16, ByRef i_dr As OracleDataReader) As Boolean

        Dim l_market As Int16 = db_access_general.GET_INSTITUTION_MARKET(i_institution)
        Dim l_id_pick_list As Int16 = 0

        If COUNT_STD_INSTR_PICK_LIST(i_institution, i_software, i_id_product, i_id_product_supplier, i_pick_list) > 0 Then
            l_id_pick_list = i_pick_list
        End If

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
                                   AND lsd.id_pick_list IN (" & l_id_pick_list & ")
                                 ORDER BY cfg.market DESC, cfg.institution DESC, cfg.software DESC, lsd.id_pick_list DESC, lsd.rank ASC  "

        Dim cmd As New OracleCommand(sql, Connection.conn)
        Try
            cmd.CommandType = CommandType.Text
            i_dr = cmd.ExecuteReader()

            DEBUGGER.SET_DEBUG("MEDICATION_API :: GET_ALL_INSTRUCTIONS(" & i_institution & ", " & i_software & ", " & i_id_product & ", " & i_id_product_supplier & ", " & i_pick_list & ")")
            Dim l_parameters(9) As String
            l_parameters(0) = "id_product"
            l_parameters(1) = "id_std_presc_directions"
            l_parameters(2) = "rank"
            l_parameters(3) = "id_grant"
            l_parameters(4) = "market"
            l_parameters(5) = "MARKET_DESC"
            l_parameters(6) = "software"
            l_parameters(7) = "SOFTWARE_DESC"
            l_parameters(8) = "id_pick_list"
            l_parameters(9) = "INSTITUTION"

            DEBUGGER.SET_RESPONSE("MEDICATION_API :: GET_ALL_INSTRUCTIONS", l_parameters, i_dr)
            i_dr = cmd.ExecuteReader()

            cmd.Dispose()
            Return True
        Catch ex As Exception
            DEBUGGER.SET_DEBUG_ERROR_INIT("MEDICATION_API :: GET_ALL_INSTRUCTIONS")
            DEBUGGER.SET_DEBUG(ex.Message)
            DEBUGGER.SET_DEBUG(sql)
            DEBUGGER.SET_DEBUG_ERROR_CLOSE()
            cmd.Dispose()
            Return False
        End Try
    End Function

    Function GET_STD_PRESC_DIR(ByVal i_institution As Int64, ByVal i_id_std_presc_dir As Int64, ByRef i_dr As OracleDataReader) As Boolean

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

            DEBUGGER.SET_DEBUG("MEDICATION_API :: GET_STD_PRESC_DIR(" & i_institution & ", " & i_id_std_presc_dir & ")")
            Dim l_parameters(7) As String
            l_parameters(0) = "id_std_presc_directions"
            l_parameters(1) = "flg_sos"
            l_parameters(2) = "id_sos_take_condition"
            l_parameters(3) = "sos_take_condition"
            l_parameters(4) = "admin_site"
            l_parameters(5) = "admin_method"
            l_parameters(6) = "notes"
            l_parameters(7) = "patient_instructions"

            DEBUGGER.SET_RESPONSE("MEDICATION_API :: GET_STD_PRESC_DIR", l_parameters, i_dr)
            i_dr = cmd.ExecuteReader()

            cmd.Dispose()
            Return True
        Catch ex As Exception
            DEBUGGER.SET_DEBUG_ERROR_INIT("MEDICATION_API :: GET_STD_PRESC_DIR")
            DEBUGGER.SET_DEBUG(ex.Message)
            DEBUGGER.SET_DEBUG(sql)
            DEBUGGER.SET_DEBUG_ERROR_CLOSE()
            cmd.Dispose()
            Return False
        End Try

    End Function

    Function GET_SOS_LIST(ByVal i_institution As Int64, ByVal i_software As Int64, ByRef i_dr As OracleDataReader) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution)
        Dim sql As String = "   SELECT DISTINCT *
                                  FROM (SELECT cr.id_cancel_reason AS id_cancel_reason,
                                               nvl2(rsi.desc_synonym, rsi.desc_synonym, pk_translation.get_translation(" & l_id_language & ", cr.code_cancel_reason)) reason_desc
                                          FROM alert.cancel_reason cr
                                          LEFT JOIN alert.reason_synonym_inst rsi
                                            ON rsi.id_reason = cr.id_cancel_reason
                                           AND rsi.id_institution = " & i_institution & "
                                          JOIN alert.reason_action_relation rar
                                            ON rar.id_reason = cr.id_cancel_reason
                                          JOIN alert.reason_action ra
                                            ON ra.id_action = rar.id_action
                                           AND ra.flg_type = 'O'
                                          JOIN alert.cancel_rea_soft_inst crsi
                                            ON crsi.id_cancel_reason = cr.id_cancel_reason
                                          JOIN alert.cancel_rea_area cra
                                            ON cra.id_cancel_rea_area = crsi.id_cancel_rea_area
                                         WHERE upper(cra.intern_name) = upper('PRN_DEFAULT_REASONS')
                                           AND crsi.id_software IN (0, " & i_software & " )
                                           AND crsi.id_institution IN (0, " & i_institution & ")
                                           AND crsi.flg_available = 'Y')
                                 ORDER BY reason_desc ASC "

        Dim cmd As New OracleCommand(sql, Connection.conn)
        Try
            cmd.CommandType = CommandType.Text
            i_dr = cmd.ExecuteReader()

            DEBUGGER.SET_DEBUG("MEDICATION_API :: GET_SOS_LIST(" & i_institution & ", " & i_software & ")")

            Dim l_parameters(1) As String
            l_parameters(0) = "id_cancel_reason"
            l_parameters(1) = "reason_desc"

            DEBUGGER.SET_RESPONSE("MEDICATION_API :: GET_SOS_LIST", l_parameters, i_dr)
            i_dr = cmd.ExecuteReader()

            cmd.Dispose()
            Return True
        Catch ex As Exception
            DEBUGGER.SET_DEBUG_ERROR_INIT("MEDICATION_API :: GET_SOS_LIST")
            DEBUGGER.SET_DEBUG(ex.Message)
            DEBUGGER.SET_DEBUG(sql)
            DEBUGGER.SET_DEBUG_ERROR_CLOSE()
            cmd.Dispose()
            Return False
        End Try

    End Function

    Function GET_ADMIN_METHOD_LIST(ByVal i_institution As Int64, ByVal i_id_route As String, ByVal i_id_route_supplier As String, ByRef i_dr As OracleDataReader) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution)
        Dim sql As String = "   SELECT lras.id_admin_method, pk_translation.get_translation(" & l_id_language & ", am.code_admin_method) AS admin_method_desc
                                  FROM alert_product_mt.lnk_route_admin_method lras
                                  JOIN alert_product_mt.admin_method am
                                    ON am.id_admin_method = lras.id_admin_method
                                 WHERE lras.id_route = '" & i_id_route & "'
                                   AND lras.id_route_supplier = '" & i_id_route_supplier & "'
                                   AND lras.flg_available = 'Y'
                                 ORDER BY admin_method_desc "

        Dim cmd As New OracleCommand(sql, Connection.conn)
        Try
            cmd.CommandType = CommandType.Text
            i_dr = cmd.ExecuteReader()

            DEBUGGER.SET_DEBUG("MEDICATION_API :: GET_ADMIN_METHOD_LIST(" & i_institution & ", " & i_id_route & ", " & i_id_route_supplier & ")")
            Dim l_parameters(1) As String
            l_parameters(0) = "id_admin_method"
            l_parameters(1) = "admin_method_desc"

            DEBUGGER.SET_RESPONSE("MEDICATION_API :: GET_ADMIN_METHOD_LIST", l_parameters, i_dr)
            i_dr = cmd.ExecuteReader()

            cmd.Dispose()
            Return True
        Catch ex As Exception
            DEBUGGER.SET_DEBUG_ERROR_INIT("MEDICATION_API :: GET_ADMIN_METHOD_LIST")
            DEBUGGER.SET_DEBUG(ex.Message)
            DEBUGGER.SET_DEBUG(sql)
            DEBUGGER.SET_DEBUG_ERROR_CLOSE()
            cmd.Dispose()
            Return False
        End Try

    End Function

    Function GET_ADMIN_SITE_LIST(ByVal i_institution As Int64, ByVal i_id_route As String, ByVal i_id_route_supplier As String, ByRef i_dr As OracleDataReader) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution)
        Dim sql As String = "  SELECT am.id_admin_site, pk_translation.get_translation(" & l_id_language & ", am.code_admin_site) AS admin_site_desc
                               FROM alert_product_mt.lnk_route_admin_site lras
                               JOIN alert_product_mt.admin_site am
                                 ON am.id_admin_site = lras.id_admin_site
                              WHERE lras.id_route = '" & i_id_route & "'
                                AND lras.id_route_supplier = '" & i_id_route_supplier & "'
                                order by admin_site_desc"

        Dim cmd As New OracleCommand(sql, Connection.conn)
        Try
            cmd.CommandType = CommandType.Text
            i_dr = cmd.ExecuteReader()

            DEBUGGER.SET_DEBUG("MEDICATION_API :: GET_ADMIN_SITE_LIST(" & i_institution & ", " & i_id_route & ", " & i_id_route_supplier & ")")
            Dim l_parameters(1) As String
            l_parameters(0) = "id_admin_site"
            l_parameters(1) = "admin_site_desc"

            DEBUGGER.SET_RESPONSE("MEDICATION_API :: GET_ADMIN_SITE_LIST", l_parameters, i_dr)
            i_dr = cmd.ExecuteReader()

            cmd.Dispose()
            Return True
        Catch ex As Exception
            DEBUGGER.SET_DEBUG_ERROR_INIT("MEDICATION_API :: GET_ADMIN_SITE_LIST")
            DEBUGGER.SET_DEBUG(ex.Message)
            DEBUGGER.SET_DEBUG(sql)
            DEBUGGER.SET_DEBUG_ERROR_CLOSE()
            cmd.Dispose()
            Return False
        End Try

    End Function

    Function GET_DURATION_UM(ByVal i_institution As Int64, ByRef i_dr As OracleDataReader) As Boolean

        Dim sql As String = "  SELECT lnk.id_unit_measure AS id,
                                       pk_translation.get_translation(2, um.code_unit_measure) as um_desc
                                  FROM alert_product_mt.lnk_main_um_umc lnk
                                  JOIN alert.unit_measure um
                                    ON um.id_unit_measure = lnk.id_unit_measure
                                   AND um.id_unit_measure_type = lnk.id_unit_measure_context
                                 WHERE lnk.id_unit_measure_context = 5
                                 ORDER BY um_desc"

        Dim cmd As New OracleCommand(sql, Connection.conn)
        Try
            cmd.CommandType = CommandType.Text
            i_dr = cmd.ExecuteReader()

            DEBUGGER.SET_DEBUG("MEDICATION_API :: GET_DURATION_UM(" & i_institution & ")")

            Dim l_parameters(1) As String
            l_parameters(0) = "id_unit_measure"
            l_parameters(1) = "um_desc"

            DEBUGGER.SET_RESPONSE("MEDICATION_API :: GET_DURATION_UM", l_parameters, i_dr)
            i_dr = cmd.ExecuteReader()

            cmd.Dispose()
            Return True
        Catch ex As Exception
            DEBUGGER.SET_DEBUG_ERROR_INIT("MEDICATION_API :: GET_DURATION_UM")
            DEBUGGER.SET_DEBUG(ex.Message)
            DEBUGGER.SET_DEBUG(sql)
            DEBUGGER.SET_DEBUG_ERROR_CLOSE()
            cmd.Dispose()
            Return False
        End Try

    End Function

    Function GET_STD_PRESC_DIR_ITEM(ByVal i_institution As Int64, ByVal i_id_product As String, ByVal i_id_pick_list As Int16, ByVal i_id_std_presc_dir As Int64, ByVal i_grant As Int64, ByVal i_rank As Int64, ByRef i_dr As OracleDataReader) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution)

        Dim sql As String = "SELECT *
                              FROM (SELECT pi.rank,
                                           pi.duration_value,
                                           pi.id_unit_duration,
                                           pk_translation.get_translation(2, um_duration.code_unit_measure) AS duration_desc,
                                           pi.num_executions,
                                           pi.id_dose,
                                           pi.dose_value,
                                           pi.id_unit_dose,
                                           pk_translation.get_translation(" & l_id_language & ", um.code_unit_measure) Unit_Measure,
                                           pi.id_presc_dir_frequency AS ID_FREQUENCY,
                                           nvl(tf.desc_lang_" & l_id_language & ", t_single.desc_lang_" & l_id_language & ") AS Frequency,
                                           rownum rn
                                      FROM alert_product_mt.std_presc_dir d
                                      LEFT JOIN alert_product_mt.lnk_product_std_presc_dir lsd
                                        ON lsd.id_std_presc_directions = d.id_std_presc_directions
                                      LEFT JOIN alert_product_mt.std_presc_dir_item pi
                                        ON pi.id_std_presc_directions = d.id_std_presc_directions
                                      LEFT JOIN alert_product_mt.presc_dir_frequency f
                                        ON f.id_presc_dir_frequency = pi.id_presc_dir_frequency
                                      LEFT JOIN translation tf
                                        ON tf.code_translation = f.code_presc_dir_frequency
                                      LEFT JOIN alert.unit_measure um
                                        ON um.id_unit_measure = pi.id_unit_dose
                                      LEFT JOIN alert.unit_measure um_duration
                                        ON um_duration.id_unit_measure = pi.id_unit_duration
                                      LEFT JOIN alert_product_mt.presc_list_item pli
                                        ON pli.id_presc_list_item = pi.id_presc_dir_frequency
                                      LEFT JOIN translation t_single
                                        ON t_single.code_translation = pli.code_presc_list_item
                                     WHERE lsd.id_product IN ('" & i_id_product & "')
                                       AND lsd.id_grant = " & i_grant & "
                                       AND lsd.id_pick_list = " & i_id_pick_list & "
                                       AND lsd.rank = " & i_rank & "
                                       and d.id_std_presc_directions = " & i_id_std_presc_dir & "
                                  order by rank asc)
                             WHERE rn < 8 " 'limitar a 7 instruções

        Dim cmd As New OracleCommand(sql, Connection.conn)
        Try
            cmd.CommandType = CommandType.Text
            i_dr = cmd.ExecuteReader()

            DEBUGGER.SET_DEBUG("MEDICATION_API :: GET_STD_PRESC_DIR_ITEM(" & i_institution & ", " & i_id_product & ", " & i_id_pick_list & ", " & i_id_std_presc_dir & ", " & i_grant & ", " & i_rank & ")")

            Dim l_parameters(11) As String
            l_parameters(0) = "rank"
            l_parameters(1) = "duration_value"
            l_parameters(2) = "id_unit_duration"
            l_parameters(3) = "duration_desc"
            l_parameters(4) = "num_executions"
            l_parameters(5) = "id_dose"
            l_parameters(6) = "dose_value"
            l_parameters(7) = "id_unit_dose"
            l_parameters(8) = "Unit_Measure"
            l_parameters(9) = "ID_FREQUENCY"
            l_parameters(10) = "Frequency"
            l_parameters(11) = "rownum"

            DEBUGGER.SET_RESPONSE("MEDICATION_API :: GET_STD_PRESC_DIR_ITEM", l_parameters, i_dr)
            i_dr = cmd.ExecuteReader()

            cmd.Dispose()
            Return True
        Catch ex As Exception
            DEBUGGER.SET_DEBUG_ERROR_INIT("MEDICATION_API :: GET_STD_PRESC_DIR_ITEM")
            DEBUGGER.SET_DEBUG(ex.Message)
            DEBUGGER.SET_DEBUG(sql)
            DEBUGGER.SET_DEBUG_ERROR_CLOSE()
            cmd.Dispose()
            Return False
        End Try
    End Function

    Function GET_STD_CREEN_TYPE(ByVal i_institution As Int64, ByVal i_id_product As String, ByVal i_id_product_supplier As String) As Int16

        Dim l_id_type_screen As Int16 = 0
        Dim sql As String = "SELECT DISTINCT (nvl2(c.id_product, 3, DECODE(pm.id_product_med_type,2,2,1)))
                              FROM alert_product_mt.product_medication pm
                              LEFT JOIN alert_product_mt.product_med_component c
                                ON c.id_product = pm.id_product
                               AND c.id_product_supplier = pm.id_product_supplier
                               AND pm.id_product_med_type = 2
                             WHERE pm.id_product = '" & i_id_product & "'
                               AND pm.id_product_supplier = '" & i_id_product_supplier & "'"

        Dim cmd As New OracleCommand(sql, Connection.conn)
        cmd.CommandType = CommandType.Text

        Dim dr As OracleDataReader

        Try
            dr = cmd.ExecuteReader()

            DEBUGGER.SET_DEBUG("MEDICATION_API :: GET_STD_CREEN_TYPE(" & i_institution & ", " & i_id_product & ", " & i_id_product_supplier & ")")

            Dim l_parameters(0) As String
            l_parameters(0) = "screen_type"

            DEBUGGER.SET_RESPONSE("MEDICATION_API :: GET_STD_CREEN_TYPE", l_parameters, dr)
            dr = cmd.ExecuteReader()

            While dr.Read()
                l_id_type_screen = dr.Item(0)
            End While

            dr.Dispose()
            dr.Close()
            cmd.Dispose()

        Catch ex As Exception
            DEBUGGER.SET_DEBUG_ERROR_INIT("MEDICATION_API :: GET_STD_CREEN_TYPE")
            DEBUGGER.SET_DEBUG(ex.Message)
            DEBUGGER.SET_DEBUG(sql)
            DEBUGGER.SET_DEBUG_ERROR_CLOSE()
            cmd.Dispose()
        End Try

        Return l_id_type_screen

    End Function

    Function GET_ID_GRANT(ByVal i_institution As Int64, ByVal i_software As String, ByVal i_context As String) As Int64

        Dim l_id_grant As Int64 = -1
        Dim l_market As Int16 = db_access_general.GET_INSTITUTION_MARKET(i_institution)

        Dim sql As String = "SELECT id_grant
                              FROM (SELECT id_grant, grant_order, market, institution, software, decode(grant_order, 99,99,rownum) AS rn
                                      FROM (SELECT *
                                              FROM (SELECT c.id_grant,
                                                           c.market,
                                                           c.institution,
                                                           c.software,
                                                           decode(c.grant_order, 0, 99, c.grant_order) AS grant_order
                                                             FROM alert_product_mt.v_cfg_grant c
                                                            WHERE c.id_context = '" & i_context & "'
                                                              AND c.market IN (0, " & l_market & ")
                                                              AND c.institution IN (0, " & i_institution & ")
                                                              AND c.software IN (0, " & i_software & "))
                                             ORDER BY market DESC, institution DESC, software DESC, grant_order ASC)
                                     ORDER BY rn ASC)
                             WHERE rownum = 1"

        Dim cmd As New OracleCommand(sql, Connection.conn)
        cmd.CommandType = CommandType.Text

        Dim dr As OracleDataReader

        Try
            dr = cmd.ExecuteReader()

            DEBUGGER.SET_DEBUG("MEDICATION_API :: GET_ID_GRANT(" & i_institution & ", " & i_software & ", " & i_context & ")")

            Dim l_parameters(0) As String
            l_parameters(0) = "id_grant"

            DEBUGGER.SET_RESPONSE("MEDICATION_API :: GET_ID_GRANT", l_parameters, dr)
            dr = cmd.ExecuteReader()

            While dr.Read()
                Try
                    l_id_grant = dr.Item(0)
                Catch ex As Exception
                    l_id_grant = -1
                End Try
            End While

            dr.Dispose()
            dr.Close()
            cmd.Dispose()

        Catch ex As Exception
            DEBUGGER.SET_DEBUG_ERROR_INIT("MEDICATION_API :: GET_STD_CREEN_TYPE")
            DEBUGGER.SET_DEBUG(ex.Message)
            DEBUGGER.SET_DEBUG(sql)
            DEBUGGER.SET_DEBUG_ERROR_CLOSE()
            cmd.Dispose()
        End Try

        Return l_id_grant

    End Function

    Function SET_ID_GRANT(ByVal i_institution As Int64, ByVal i_software As String, ByVal i_context As String) As Boolean

        DEBUGGER.SET_DEBUG("MEDICATION_API :: SET_ID_GRANT(" & i_institution & ", " & i_software & ", " & i_context & ")")

        Dim l_market As Int16 = db_access_general.GET_INSTITUTION_MARKET(i_institution)
        Dim sql As String = " DECLARE
                                    l_grant NUMBER;
                                BEGIN

                                    l_grant := alert_product_tr.pk_grants.set_by_soft_inst(i_context     => 'LNK_PRODUCT_STD_PRESC_DIR',
                                                                                           i_prof        => profissional(0, " & i_institution & ", " & i_software & "),
                                                                                           i_market      => " & l_market & ",
                                                                                           i_grant_order => 1);
                                END;  "

        Dim cmd_default_um As New OracleCommand(sql, Connection.conn)

        Try
            cmd_default_um.CommandType = CommandType.Text
            cmd_default_um.ExecuteNonQuery()
        Catch ex As Exception
            DEBUGGER.SET_DEBUG_ERROR_INIT("MEDICATION_API :: SET_ID_GRANT")
            DEBUGGER.SET_DEBUG(ex.Message)
            DEBUGGER.SET_DEBUG(sql)
            DEBUGGER.SET_DEBUG_ERROR_CLOSE()
            cmd_default_um.Dispose()
            Return False
        End Try

        cmd_default_um.Dispose()

        Return True

    End Function

    Function UPDATE_STD_PRESC_DIR(ByVal i_institution As Int64, ByVal i_id_product As String, ByVal i_id_product_supplier As String, ByVal i_id_std_presc_dir As Int64, ByVal i_id_grant As Int64, ByVal i_id_pick_list As Int16, ByVal i_id_new_std_presc_dir As Int64, ByVal i_rank As Int64, ByVal i_id_new_grant As Int64) As Boolean

        DEBUGGER.SET_DEBUG("MEDICATION_API :: UPDATE_STD_PRESC_DIR(" & i_institution & ", " & i_id_product & ", " & i_id_product_supplier & ", " & i_id_std_presc_dir & ", " & i_id_grant & ", " & i_id_pick_list & ", " & i_id_new_std_presc_dir & ", " & i_rank & ")")

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution)

        Dim sql As String = "UPDATE alert_product_mt.lnk_product_std_presc_dir d
                               SET d.rank = " & i_rank & ", d.id_std_presc_directions = " & i_id_new_std_presc_dir & ", d.id_grant = " & i_id_new_grant & "
                             WHERE d.id_product = '" & i_id_product & "'
                               AND d.id_std_presc_directions = " & i_id_std_presc_dir & "
                               AND d.id_product_supplier = '" & i_id_product_supplier & "'
                               AND d.id_grant = " & i_id_grant & "
                               AND d.id_pick_list =" & i_id_pick_list

        Dim cmd_update_std As New OracleCommand(sql, Connection.conn)

        Try
            cmd_update_std.CommandType = CommandType.Text
            cmd_update_std.ExecuteNonQuery()
        Catch ex As Exception
            DEBUGGER.SET_DEBUG_ERROR_INIT("MEDICATION_API :: UPDATE_STD_PRESC_DIR")
            DEBUGGER.SET_DEBUG(ex.Message)
            DEBUGGER.SET_DEBUG(sql)
            DEBUGGER.SET_DEBUG_ERROR_CLOSE()
            cmd_update_std.Dispose()
            Return False
        End Try

        cmd_update_std.Dispose()

        Return True

    End Function

    Function CREATE_STD_PRESC_DIR(ByVal i_institution As Int64, ByVal i_id_product As String, ByVal i_id_product_supplier As String, ByVal i_id_std_presc_dir As Int64, ByVal i_id_grant As Int64, ByVal i_id_pick_list As Int16, ByVal i_rank As Int64) As Boolean

        DEBUGGER.SET_DEBUG("MEDICATION_API :: CREATE_STD_PRESC_DIR(" & i_institution & ", " & i_id_product & ", " & i_id_product_supplier & ", " & i_id_std_presc_dir & ", " & i_id_grant & ", " & i_id_pick_list & ", " & i_rank & ")")

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution)

        Dim sql As String = "INSERT INTO alert_product_mt.lnk_product_std_presc_dir
                                    (id_product, id_product_supplier, id_std_presc_directions, rank, id_grant, id_pick_list)
                                VALUES
                                    ('" & i_id_product & "', '" & i_id_product_supplier & "', " & i_id_std_presc_dir & ", " & i_rank & ", " & i_id_grant & ", " & i_id_pick_list & ")"

        Dim cmd_update_std As New OracleCommand(sql, Connection.conn)

        Try
            cmd_update_std.CommandType = CommandType.Text
            cmd_update_std.ExecuteNonQuery()
        Catch ex As Exception
            DEBUGGER.SET_DEBUG_ERROR_INIT("MEDICATION_API :: CREATE_STD_PRESC_DIR")
            DEBUGGER.SET_DEBUG(ex.Message)
            DEBUGGER.SET_DEBUG(sql)
            DEBUGGER.SET_DEBUG_ERROR_CLOSE()
            cmd_update_std.Dispose()
            Return False
        End Try

        cmd_update_std.Dispose()
        Return True

    End Function

    Function CHECK_DUP_INSTRUCTIONS(ByVal i_institution As Int64, ByVal i_id_std_presc_dir As Int64) As Int64

        Dim l_count As Int64 = 0

        Dim sql As String = "SELECT COUNT(*)
                              FROM (SELECT *
                                      FROM alert_product_mt.lnk_product_std_presc_dir d
                                     WHERE d.id_std_presc_directions = " & i_id_std_presc_dir & "
        
                                    UNION ALL
        
                                    SELECT *
                                      FROM alert_product_mt.lnk_product_std_presc_dir d
                                     WHERE d.id_std_presc_directions = " & i_id_std_presc_dir & "
                                       AND d.id_pick_list = 0)"

        Dim cmd As New OracleCommand(sql, Connection.conn)
        cmd.CommandType = CommandType.Text

        Dim dr As OracleDataReader

        Try
            dr = cmd.ExecuteReader()

            DEBUGGER.SET_DEBUG("MEDICATION_API :: CHECK_DUP_INSTRUCTIONS(" & i_institution & ", " & i_id_std_presc_dir & ")")

            Dim l_parameters(0) As String
            l_parameters(0) = "count"

            DEBUGGER.SET_RESPONSE("MEDICATION_API :: CHECK_DUP_INSTRUCTIONS", l_parameters, dr)
            dr = cmd.ExecuteReader()

            While dr.Read()
                l_count = dr.Item(0)
            End While

            dr.Dispose()
            dr.Close()
            cmd.Dispose()

        Catch ex As Exception
            DEBUGGER.SET_DEBUG_ERROR_INIT("MEDICATION_API :: CHECK_DUP_INSTRUCTIONS")
            DEBUGGER.SET_DEBUG(ex.Message)
            DEBUGGER.SET_DEBUG(sql)
            DEBUGGER.SET_DEBUG_ERROR_CLOSE()
            cmd.Dispose()
        End Try

        Return l_count

    End Function

    Function GET_NEW_STD_INSTRUCTION_ID(ByVal i_institution As Int64) As Int64

        Dim l_id_std_instruction As Int64 = 0

        Dim sql As String = "select alert_product_mt.seq_std_presc_dir.nextval from dual"

        Dim cmd As New OracleCommand(sql, Connection.conn)
        cmd.CommandType = CommandType.Text

        Dim dr As OracleDataReader

        Try
            dr = cmd.ExecuteReader()

            DEBUGGER.SET_DEBUG("MEDICATION_API :: GET_NEW_STD_INSTRUCTION_ID(" & i_institution & ")")

            Dim l_parameters(0) As String
            l_parameters(0) = "seq_std_presc_dir"

            DEBUGGER.SET_RESPONSE("MEDICATION_API :: GET_NEW_STD_INSTRUCTION_ID", l_parameters, dr)
            dr = cmd.ExecuteReader()

            While dr.Read()
                l_id_std_instruction = dr.Item(0)
            End While

            dr.Dispose()
            dr.Close()
            cmd.Dispose()

        Catch ex As Exception
            DEBUGGER.SET_DEBUG_ERROR_INIT("MEDICATION_API :: GET_NEW_STD_INSTRUCTION_ID")
            DEBUGGER.SET_DEBUG(ex.Message)
            DEBUGGER.SET_DEBUG(sql)
            DEBUGGER.SET_DEBUG_ERROR_CLOSE()
            cmd.Dispose()
        End Try

        Return l_id_std_instruction

    End Function

    Function CREATE_STD_INSTRUCTION(ByVal i_institution As Int64, ByVal i_id_std_presc_directions As Int64, ByVal i_flg_sos As String, ByVal i_id_sos As Int16, ByVal i_sos_take_condition As String, ByVal i_notes As String, ByVal i_patient_instructions As String, ByVal i_id_admin_site As String, ByVal i_id_admin_method As String) As Boolean

        DEBUGGER.SET_DEBUG("MEDICATION_API :: CREATE_STD_INSTRUCTION(" & i_institution & ", " & i_id_std_presc_directions & ", " & i_flg_sos & ", " & i_id_sos & ", " & i_sos_take_condition & ", " & i_notes & ", " & i_patient_instructions & ", " & i_id_admin_site & ", " & i_id_admin_method & ")")

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution)

        Dim l_id_sos_take_condition As String = "NULL"

        If i_id_sos <> 19 Then
            l_id_sos_take_condition = "32"
        End If

        Dim sql As String = "INSERT INTO alert_product_mt.std_presc_dir
                                (id_std_presc_directions,
                                 flg_sos,
                                 id_sos,
                                 id_sos_take_condition,
                                 sos_take_condition,
                                 flg_free_text,
                                 notes,
                                 patient_instr_desc,
                                 id_admin_site,
                                 id_admin_method)
                            VALUES
                                (" & i_id_std_presc_directions & ", '" & i_flg_sos & "', " & i_id_sos & ", " & l_id_sos_take_condition & ", " & i_sos_take_condition & ", 'N', '" & i_notes & "',  '" & i_patient_instructions & "',  " & i_id_admin_site & ",  " & i_id_admin_method & ")"

        Dim cmd_create_std As New OracleCommand(sql, Connection.conn)

        Try
            cmd_create_std.CommandType = CommandType.Text
            cmd_create_std.ExecuteNonQuery()
        Catch ex As Exception

            Try
                sql = "INSERT INTO alert_product_mt.std_presc_dir
                                (id_std_presc_directions,
                                 flg_sos,
                                 id_sos,
                                 id_sos_take_condition,
                                 sos_take_condition,
                                 flg_free_text,
                                 notes,
                                 patient_instr_desc,
                                 id_admin_site,
                                 id_admin_method)
                            VALUES
                                ((SELECT MAX(id_std_presc_directions) + 1 FROM alert_product_mt.std_presc_dir), '" & i_flg_sos & "', " & i_id_sos & ", " & l_id_sos_take_condition & ", " & i_sos_take_condition & ", 'N', '" & i_notes & "',  '" & i_patient_instructions & "',  " & i_id_admin_site & ",  " & i_id_admin_method & ")"

                cmd_create_std.ExecuteNonQuery()
            Catch ex2 As Exception
                DEBUGGER.SET_DEBUG_ERROR_INIT("MEDICATION_API :: CREATE_STD_INSTRUCTION")
                DEBUGGER.SET_DEBUG(ex.Message)
                DEBUGGER.SET_DEBUG(sql)
                DEBUGGER.SET_DEBUG_ERROR_CLOSE()
                cmd_create_std.Dispose()
            End Try
            Return False
        End Try

        cmd_create_std.Dispose()

        Return True

    End Function

    Function CREATE_STD_PRESC_DIR_ITEM(ByVal i_institution As Int64, ByVal i_id_std_presc_directions As Int64, ByVal i_index_instructions As Int16, ByVal i_a_instructions() As String) As Boolean

        Dim l_debug_string As String = "MEDICATION_API :: CREATE_STD_PRESC_DIR_ITEM(" & i_institution & ", " & i_id_std_presc_directions & ", " & i_index_instructions & ", "

        l_debug_string = l_debug_string & DEBUGGER.GET_ARRAY_STRING(i_a_instructions) & ")"

        DEBUGGER.SET_DEBUG(l_debug_string)

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution)

        Dim l_id_recurrence As String

        If i_a_instructions(2) = 1 Then
            l_id_recurrence = 0
        Else
            l_id_recurrence = i_a_instructions(2)
        End If

        Dim sql As String = "BEGIN
                             INSERT INTO alert_product_mt.std_presc_dir_item
                                (id_std_presc_directions,
                                 id_std_presc_dir_item,
                                 id_recurrence,
                                 duration_value,
                                 id_unit_duration,
                                 num_executions,
                                 id_dose,
                                 dose_value,
                                 id_unit_dose,
                                 rank,
                                 id_presc_dir_frequency)
                            VALUES
                                (" & i_id_std_presc_directions & ", ALERT_PRODUCT_MT.SEQ_STD_PRESC_DIR_ITEM.NEXTVAL, " & l_id_recurrence & ", " & i_a_instructions(3) & ", " & i_a_instructions(4) & ", " & i_a_instructions(5) & ", 10, " & i_a_instructions(0) & ", " & i_a_instructions(1) & ", " & i_index_instructions & ", " & i_a_instructions(2) & ");
                              EXCEPTION
                                    WHEN dup_val_on_index THEN
                                        UPDATE alert_product_mt.std_presc_dir_item i
                                           SET i.id_recurrence          = " & i_a_instructions(2) & ",
                                               i.id_duration            = " & i_a_instructions(3) & ",
                                               i.id_unit_duration       = " & i_a_instructions(4) & ",
                                               i.num_executions         = " & i_a_instructions(5) & ",
                                               i.id_dose                = " & i_a_instructions(0) & ",
                                               i.dose_value             = " & i_a_instructions(1) & ",
                                               i.id_presc_dir_frequency = " & i_a_instructions(2) & "
                                         WHERE i.id_std_presc_directions = " & i_id_std_presc_directions & ";
                                END;"

        Dim cmd_create_std As New OracleCommand(sql, Connection.conn)

        Try
            cmd_create_std.CommandType = CommandType.Text
            cmd_create_std.ExecuteNonQuery()
        Catch ex As Exception
            DEBUGGER.SET_DEBUG_ERROR_INIT("MEDICATION_API :: CREATE_STD_PRESC_DIR_ITEM")
            DEBUGGER.SET_DEBUG(ex.Message)
            DEBUGGER.SET_DEBUG(sql)
            DEBUGGER.SET_DEBUG_ERROR_CLOSE()
            cmd_create_std.Dispose()
            Return False
        End Try

        cmd_create_std.Dispose()

        Return True

    End Function


    Function DELETE_STD_INSTRUCTION(ByVal i_institution As Int64, ByVal i_id_product As String, ByVal i_id_product_suplier As String, ByVal i_id_std_presc_directions As Int64, ByVal i_rank As Int64, ByVal i_id_grant As Int64, ByVal i_id_pick_list As Int16) As Boolean

        DEBUGGER.SET_DEBUG("MEDICATION_API :: DELETE_STD_INSTRUCTION(" & i_institution & ", " & i_id_std_presc_directions & ", " & i_id_product_suplier & ", " & i_id_std_presc_directions & ", " & i_rank & ", " & i_id_grant & ", " & i_id_pick_list & ") ")

        Dim sql As String = " DELETE FROM alert_product_mt.lnk_product_std_presc_dir d
                              WHERE d.id_product = '" & i_id_product & "'
                                AND d.id_product_supplier = '" & i_id_product_suplier & "'
                                AND d.id_std_presc_directions = " & i_id_std_presc_directions & "
                                AND d.rank = " & i_rank & "
                                AND d.id_grant = " & i_id_grant & "
                                AND d.id_pick_list = " & i_id_pick_list

        Dim cmd_delete_std As New OracleCommand(sql, Connection.conn)

        Try
            cmd_delete_std.CommandType = CommandType.Text
            cmd_delete_std.ExecuteNonQuery()
        Catch ex As Exception
            DEBUGGER.SET_DEBUG_ERROR_INIT("MEDICATION_API :: DELETE_STD_INSTRUCTION")
            DEBUGGER.SET_DEBUG(ex.Message)
            DEBUGGER.SET_DEBUG(sql)
            DEBUGGER.SET_DEBUG_ERROR_CLOSE()
            cmd_delete_std.Dispose()
            Return False
        End Try

        cmd_delete_std.Dispose()

        Return True

    End Function


    Function GET_STD_PRESC_DIR_ITEM_IV(ByVal i_institution As Int64, ByVal i_id_product As String, ByVal i_id_pick_list As Int16, ByVal i_id_std_presc_dir As Int64, ByVal i_grant As Int64, ByVal i_rank As Int64, ByRef i_dr As OracleDataReader) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution)

        Dim sql As String = "SELECT *
                              FROM (SELECT pi.rank,
                                           pi.duration_value,
                                           pi.id_unit_duration,
                                           pk_translation.get_translation(2, um_duration.code_unit_measure) AS duration_desc,
                                           pi.num_executions,
                                           pi.id_presc_dir_frequency AS ID_FREQUENCY,
                                           nvl(tf.desc_lang_" & l_id_language & ", t_single.desc_lang_" & l_id_language & ") AS Frequency,
                                           rownum rn,
                                           pi.id_std_presc_dir_item
                                      FROM alert_product_mt.std_presc_dir d
                                      LEFT JOIN alert_product_mt.lnk_product_std_presc_dir lsd
                                        ON lsd.id_std_presc_directions = d.id_std_presc_directions
                                      LEFT JOIN alert_product_mt.std_presc_dir_item pi
                                        ON pi.id_std_presc_directions = d.id_std_presc_directions
                                      LEFT JOIN alert_product_mt.presc_dir_frequency f
                                        ON f.id_presc_dir_frequency = pi.id_presc_dir_frequency
                                      LEFT JOIN translation tf
                                        ON tf.code_translation = f.code_presc_dir_frequency
                                      LEFT JOIN alert.unit_measure um
                                        ON um.id_unit_measure = pi.id_unit_dose
                                      LEFT JOIN alert.unit_measure um_duration
                                        ON um_duration.id_unit_measure = pi.id_unit_duration
                                      LEFT JOIN alert_product_mt.presc_list_item pli
                                        ON pli.id_presc_list_item = pi.id_presc_dir_frequency
                                      LEFT JOIN translation t_single
                                        ON t_single.code_translation = pli.code_presc_list_item
                                     WHERE lsd.id_product IN ('" & i_id_product & "')
                                       AND lsd.id_grant = " & i_grant & "
                                       AND lsd.id_pick_list = " & i_id_pick_list & "
                                       AND lsd.rank = " & i_rank & "
                                       and d.id_std_presc_directions = " & i_id_std_presc_dir & "
                                  order by rank asc)
                             WHERE rn = 1 "

        Dim cmd As New OracleCommand(sql, Connection.conn)
        Try
            cmd.CommandType = CommandType.Text
            i_dr = cmd.ExecuteReader()

            DEBUGGER.SET_DEBUG("MEDICATION_API :: GET_STD_PRESC_DIR_ITEM_IV(" & i_institution & ", " & i_id_product & ", " & i_id_pick_list & ", " & i_id_std_presc_dir & ", " & i_grant & ", " & i_rank & ")")

            Dim l_parameters(8) As String
            l_parameters(0) = "rank"
            l_parameters(1) = "duration_value"
            l_parameters(2) = "id_unit_duration"
            l_parameters(3) = "duration_desc"
            l_parameters(4) = "num_executions"
            l_parameters(5) = "ID_FREQUENCY"
            l_parameters(6) = "Frequency"
            l_parameters(7) = "rownum"
            l_parameters(8) = "id_std_presc_dir_item"

            DEBUGGER.SET_RESPONSE("MEDICATION_API :: GET_STD_PRESC_DIR_ITEM_IV", l_parameters, i_dr)
            i_dr = cmd.ExecuteReader()

            cmd.Dispose()
            Return True
        Catch ex As Exception
            DEBUGGER.SET_DEBUG_ERROR_INIT("MEDICATION_API :: GET_STD_PRESC_DIR_ITEM_IV")
            DEBUGGER.SET_DEBUG(ex.Message)
            DEBUGGER.SET_DEBUG(sql)
            DEBUGGER.SET_DEBUG_ERROR_CLOSE()
            cmd.Dispose()
            Return False
        End Try
    End Function

    Function GET_STD_PRESC_DIR_ITEM_SEQ(ByVal i_institution As Int64, ByVal i_id_std_presc_dir_item As Int64, ByRef i_dr As OracleDataReader) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution)

        Dim sql As String = "SELECT *
                              FROM (SELECT s.dose_value,
                                           s.id_unit_dose,
                                           pk_translation.get_translation(" & l_id_language & ", um.code_unit_measure) um_desc,
                                           s.duration_value,
                                           CASE
                                                WHEN s.id_unit_duration = 10374
                                                     AND s.duration_value >= 60 THEN
                                                 (s.duration_value - MOD(s.duration_value, 60)) / 60
                                                ELSE
                                                 NULL
                                            END duration_hours,
                                           CASE
                                                WHEN s.id_unit_duration = 10374
                                                     AND MOD(s.duration_value, 60) > 0 THEN
                                                 MOD(s.duration_value, 60)
                                                ELSE
                                                 NULL
                                            END duration_minutes,
                                           CASE
                                                WHEN s.id_rate <> 21 THEN
                                                 s.rate_value
                                                ELSE
                                                 NULL
                                            END rate_value,
                                           s.id_unit_rate,
                                           CASE
                                                WHEN s.id_rate <> 21
                                                     OR s.id_rate IS NULL THEN
                                                 pk_translation.get_translation(2, um_rate.code_unit_measure)
                                                ELSE
                                                 pk_translation.get_translation(2, 'PRESC_RATE.CODE_RATE.21')
                                            END rate_desc,
                                           rownum rn
                                      FROM alert_product_mt.std_presc_dir_item_seq s
                                      LEFT JOIN alert.unit_measure um
                                        ON um.id_unit_measure = s.id_unit_dose
                                      LEFT JOIN alert.unit_measure um_duration
                                        ON um_duration.id_unit_measure = s.id_unit_duration
                                      LEFT JOIN alert.unit_measure um_rate
                                        ON um_rate.id_unit_measure = s.id_unit_rate
                                     WHERE s.id_std_presc_dir_item = " & i_id_std_presc_dir_item & "
                                     ORDER BY s.id_item_seq ASC)
                             WHERE rn < 8 "

        Dim cmd As New OracleCommand(sql, Connection.conn)
        Try
            cmd.CommandType = CommandType.Text
            i_dr = cmd.ExecuteReader()

            DEBUGGER.SET_DEBUG("MEDICATION_API :: GET_STD_PRESC_DIR_ITEM_SEQ(" & i_institution & ", " & i_id_std_presc_dir_item & ")")

            Dim l_parameters(9) As String
            l_parameters(0) = "dose_value"
            l_parameters(1) = "id_unit_dose"
            l_parameters(2) = "um_desc"
            l_parameters(3) = "duration_value"
            l_parameters(4) = "duration_hours"
            l_parameters(5) = "duration_minutes"
            l_parameters(6) = "rate_value"
            l_parameters(7) = "id_unit_rate"
            l_parameters(8) = "rate_desc"
            l_parameters(9) = "rownum"

            DEBUGGER.SET_RESPONSE("MEDICATION_API :: GET_STD_PRESC_DIR_ITEM_SEQ", l_parameters, i_dr)
            i_dr = cmd.ExecuteReader()

            cmd.Dispose()
            Return True
        Catch ex As Exception
            DEBUGGER.SET_DEBUG_ERROR_INIT("MEDICATION_API :: GET_STD_PRESC_DIR_ITEM_SEQ")
            DEBUGGER.SET_DEBUG(ex.Message)
            DEBUGGER.SET_DEBUG(sql)
            DEBUGGER.SET_DEBUG_ERROR_CLOSE()
            cmd.Dispose()
            Return False
        End Try
    End Function

    Function GET_STD_PRESC_DIR_ITEM_SEQ_COMPLEX(ByVal i_institution As Int64, ByVal i_id_std_presc_dir_item As Int64, ByVal i_id_product_component As String, ByVal i_id_product_component_supp As String, ByRef i_dr As OracleDataReader) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution)

        Dim sql As String = "SELECT *
                              FROM (SELECT s.dose_value,
                                           s.id_unit_dose,
                                           pk_translation.get_translation(" & l_id_language & ", um.code_unit_measure) um_desc,
                                           rownum rn
                                      FROM alert_product_mt.std_presc_dir_item_seq s
                                      LEFT JOIN alert.unit_measure um
                                        ON um.id_unit_measure = s.id_unit_dose
                                      LEFT JOIN alert.unit_measure um_duration
                                        ON um_duration.id_unit_measure = s.id_unit_duration
                                      LEFT JOIN alert.unit_measure um_rate
                                        ON um_rate.id_unit_measure = s.id_unit_rate
                                     WHERE s.id_std_presc_dir_item = " & i_id_std_presc_dir_item & "
                                        AND S.ID_PRODUCT = '" & i_id_product_component & "'
                                        AND S.ID_PRODUCT_SUPPLIER= '" & i_id_product_component_supp & "'
                                     ORDER BY s.id_item_seq ASC)
                             WHERE rn < 8 "

        Dim cmd As New OracleCommand(sql, Connection.conn)
        Try
            cmd.CommandType = CommandType.Text
            i_dr = cmd.ExecuteReader()

            DEBUGGER.SET_DEBUG("MEDICATION_API :: GET_STD_PRESC_DIR_ITEM_SEQ_COMPLEX(" & i_institution & ", " & i_id_std_presc_dir_item & ", " & i_id_product_component & ", " & i_id_product_component_supp & ")")

            Dim l_parameters(3) As String
            l_parameters(0) = "dose_value"
            l_parameters(1) = "id_unit_dose"
            l_parameters(2) = "um_desc"
            l_parameters(3) = "rownum"

            DEBUGGER.SET_RESPONSE("MEDICATION_API :: GET_STD_PRESC_DIR_ITEM_SEQ_COMPLEX", l_parameters, i_dr)
            i_dr = cmd.ExecuteReader()

            cmd.Dispose()
            Return True
        Catch ex As Exception
            DEBUGGER.SET_DEBUG_ERROR_INIT("MEDICATION_API :: GET_STD_PRESC_DIR_ITEM_SEQ_COMPLEX")
            DEBUGGER.SET_DEBUG(ex.Message)
            DEBUGGER.SET_DEBUG(sql)
            DEBUGGER.SET_DEBUG_ERROR_CLOSE()
            cmd.Dispose()
            Return False
        End Try
    End Function

    Function GET_STD_PRESC_DIR_ADMIX_SEQ(ByVal i_institution As Int64, ByVal i_id_std_presc_dir_item As Int64, ByRef i_dr As OracleDataReader) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution)
        Dim sql As String = "SELECT *
                              FROM (SELECT s.duration_value,
                                           CASE
                                                WHEN s.id_unit_duration = 10374
                                                     AND s.duration_value >= 60 THEN
                                                 (s.duration_value - MOD(s.duration_value, 60)) / 60
                                                ELSE
                                                 NULL
                                            END duration_hours,
                                           CASE
                                                WHEN s.id_unit_duration = 10374
                                                     AND MOD(s.duration_value, 60) > 0 THEN
                                                 MOD(s.duration_value, 60)
                                                ELSE
                                                 NULL
                                            END duration_minutes,
                                           CASE
                                                WHEN s.id_rate <> 21 THEN
                                                 s.rate_value
                                                ELSE
                                                 NULL
                                            END rate_value,
                                           s.id_unit_rate,
                                           CASE
                                                WHEN s.id_rate <> 21
                                                     OR s.id_rate IS NULL THEN
                                                 pk_translation.get_translation(2, um_rate.code_unit_measure)
                                                ELSE
                                                 pk_translation.get_translation(2, 'PRESC_RATE.CODE_RATE.21')
                                            END rate_desc,
                                           rownum rn
                                      FROM alert_product_mt.Std_Presc_Dir_Admixture_Seq s
                                      LEFT JOIN alert.unit_measure um_duration
                                        ON um_duration.id_unit_measure = s.id_unit_duration
                                      LEFT JOIN alert.unit_measure um_rate
                                        ON um_rate.id_unit_measure = s.id_unit_rate
                                     WHERE s.id_std_presc_dir_item = " & i_id_std_presc_dir_item & "
                                     ORDER BY s.id_item_seq ASC)
                             WHERE rn < 8"

        Dim cmd As New OracleCommand(sql, Connection.conn)
        Try
            cmd.CommandType = CommandType.Text
            i_dr = cmd.ExecuteReader()

            DEBUGGER.SET_DEBUG("MEDICATION_API :: GET_STD_PRESC_DIR_ADMIX_SEQ(" & i_institution & ", " & i_id_std_presc_dir_item & ")")

            Dim l_parameters(6) As String
            l_parameters(0) = "duration_value"
            l_parameters(1) = "duration_hours"
            l_parameters(2) = "duration_minutes"
            l_parameters(3) = "rate_value"
            l_parameters(4) = "id_unit_rate"
            l_parameters(5) = "rate_desc"
            l_parameters(6) = "rownum"

            DEBUGGER.SET_RESPONSE("MEDICATION_API :: GET_STD_PRESC_DIR_ADMIX_SEQ", l_parameters, i_dr)
            i_dr = cmd.ExecuteReader()
            cmd.Dispose()
            Return True
        Catch ex As Exception
            DEBUGGER.SET_DEBUG_ERROR_INIT("MEDICATION_API :: GET_STD_PRESC_DIR_ADMIX_SEQ")
            DEBUGGER.SET_DEBUG(ex.Message)
            DEBUGGER.SET_DEBUG(sql)
            DEBUGGER.SET_DEBUG_ERROR_CLOSE()
            cmd.Dispose()
            Return False
        End Try
    End Function

    Function CREATE_STD_PRESC_DIR_ITEM_IV(ByVal i_institution As Int64, ByVal i_id_std_presc_directions As Int64, ByVal i_id_std_presc_dir_item As Int64, ByVal i_id_frequency As String, ByVal i_duration As String, ByVal i_unit_duration As String, ByVal i_num_exec As String) As Boolean

        DEBUGGER.SET_DEBUG("MEDICATION_API :: CREATE_STD_PRESC_DIR_ITEM_IV(" & i_institution & ", " & i_id_std_presc_directions & ", " & i_id_std_presc_dir_item & ", " & i_id_frequency & ", " & i_duration & ", " & i_unit_duration & ", " & i_num_exec & "())")

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution)

        Dim l_id_frequency As String = "NULL"
        Dim l_id_recurrence As String = "0"
        If i_id_frequency <> "" And i_id_frequency <> "1" Then
            l_id_frequency = i_id_frequency
            l_id_recurrence = i_id_frequency
        End If

        Dim l_id_duration As String = "NULL"
        If i_duration <> "" Then
            l_id_duration = i_duration
        End If

        Dim l_id_unit_duration As String = "NULL"
        If i_unit_duration <> "" Then
            l_id_unit_duration = i_unit_duration
        End If

        Dim l_num_execs As String = "NULL"
        If i_num_exec <> "" Then
            l_num_execs = i_num_exec
        End If

        Dim sql As String = "BEGIN
                                INSERT INTO alert_product_mt.std_presc_dir_item
                                    (id_std_presc_directions,
                                     id_std_presc_dir_item,
                                     id_recurrence,
                                     duration_value,
                                     id_unit_duration,
                                     num_executions,
                                     rank,
                                     id_presc_dir_frequency)
                                VALUES
                                    (" & i_id_std_presc_directions & ", " & i_id_std_presc_dir_item & ", " & l_id_recurrence & ", " & l_id_duration & ", " & l_id_unit_duration & ", " & l_num_execs & ", 1, " & l_id_frequency & ");
                              EXCEPTION
                                    WHEN dup_val_on_index THEN
                                        UPDATE alert_product_mt.std_presc_dir_item i
                                           SET i.id_recurrence          = " & l_id_recurrence & ",
                                               i.id_duration            = " & l_id_duration & ",
                                               i.id_unit_duration       = " & l_id_unit_duration & ",
                                               i.num_executions         = " & l_num_execs & ",
                                               i.id_presc_dir_frequency = " & l_id_frequency & "
                                         WHERE i.id_std_presc_directions = " & i_id_std_presc_directions & ";
                                END;"

        Dim cmd_create_std As New OracleCommand(sql, Connection.conn)

        Try
            cmd_create_std.CommandType = CommandType.Text
            cmd_create_std.ExecuteNonQuery()
        Catch ex As Exception
            DEBUGGER.SET_DEBUG_ERROR_INIT("MEDICATION_API :: CREATE_STD_PRESC_DIR_ITEM_IV")
            DEBUGGER.SET_DEBUG(ex.Message)
            DEBUGGER.SET_DEBUG(sql)
            DEBUGGER.SET_DEBUG_ERROR_CLOSE()
            cmd_create_std.Dispose()

            Return False
        End Try

        cmd_create_std.Dispose()

        Return True

    End Function

    Function CREATE_STD_PRESC_DIR_ITEM_SEQ(ByVal i_institution As Int64, ByVal i_id_std_presc_dir_ITEM As Int64, ByVal i_id_product_Supplier As String, ByVal i_item_seq As Int16, ByVal i_instructions() As String) As Boolean

        Dim l_debugg_string As String = "MEDICATION_API :: CREATE_STD_PRESC_DIR_ITEM_SEQ(" & i_institution & ", " & i_id_std_presc_dir_ITEM & ", " & i_id_product_Supplier & ", " & i_item_seq & ", "

        l_debugg_string = l_debugg_string & DEBUGGER.GET_ARRAY_STRING(i_instructions) & ")"

        DEBUGGER.SET_DEBUG(l_debugg_string)

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution)

        Dim sql As String = "BEGIN
                                INSERT INTO alert_product_mt.std_presc_dir_item_seq
                                    (id_std_presc_dir_item,
                                     id_item_seq,
                                     id_dose,
                                     dose_value,
                                     id_unit_dose,
                                     duration_value,
                                     id_unit_duration,
                                     id_rate,
                                     rate_value,
                                     id_unit_rate,
                                     id_product,
                                     id_product_supplier)
                                VALUES
                                    (" & i_id_std_presc_dir_ITEM & ", " & i_item_seq & ", 10, " & i_instructions(0) & ", " & i_instructions(1) & ", " & i_instructions(5) & ", " & i_instructions(6) & ", " & i_instructions(4) & ", " & i_instructions(2) & ", " & i_instructions(3) & ",  " & i_instructions(7) & ", '" & i_id_product_Supplier & "');
                                END;"

        Dim cmd_create_std As New OracleCommand(sql, Connection.conn)

        Try
            cmd_create_std.CommandType = CommandType.Text
            cmd_create_std.ExecuteNonQuery()
        Catch ex As Exception
            DEBUGGER.SET_DEBUG_ERROR_INIT("MEDICATION_API :: CREATE_STD_PRESC_DIR_ITEM_SEQ")
            DEBUGGER.SET_DEBUG(ex.Message)
            DEBUGGER.SET_DEBUG(sql)
            DEBUGGER.SET_DEBUG_ERROR_CLOSE()
            cmd_create_std.Dispose()

            Return False
        End Try

        cmd_create_std.Dispose()

        Return True

    End Function

    Function CREATE_STD_PRESC_DIR_ADMIXTURE_SEQ(i_id_std_presc_dir_ITEM As Int64, ByVal i_id_item_seq As Int64, ByVal i_id_rate As String, ByVal i_rate_value As String, ByVal i_id_unit_rate As String, ByVal i_duration_value As String, ByVal i_id_unit_duration As String) As Boolean

        Dim l_debugg_string As String = "MEDICATION_API :: CREATE_STD_PRESC_DIR_ADMIXTURE_SEQ(" & i_id_std_presc_dir_ITEM & ", " & i_id_item_seq & ", " & i_id_rate & ", " & i_rate_value & ", " & i_id_unit_rate & ", " & i_duration_value & ", " & i_id_unit_duration & ")"

        DEBUGGER.SET_DEBUG(l_debugg_string)

        Dim sql As String = "BEGIN
                                INSERT INTO alert_product_mt.std_presc_dir_admixture_seq
                                    (id_std_presc_dir_item, id_item_seq, id_rate, rate_value, id_unit_rate, duration_value, id_unit_duration)
                                VALUES
                                    (" & i_id_std_presc_dir_ITEM & ", " & i_id_item_seq & ", " & i_id_rate & ", " & i_rate_value & ", " & i_id_unit_rate & ", " & i_duration_value & ", " & i_id_unit_duration & ");
                             END;"

        Dim cmd_create_std As New OracleCommand(sql, Connection.conn)

        Try
            cmd_create_std.CommandType = CommandType.Text
            cmd_create_std.ExecuteNonQuery()
        Catch ex As Exception
            DEBUGGER.SET_DEBUG_ERROR_INIT("MEDICATION_API :: CREATE_STD_PRESC_DIR_ADMIXTURE_SEQ")
            DEBUGGER.SET_DEBUG(ex.Message)
            DEBUGGER.SET_DEBUG(sql)
            DEBUGGER.SET_DEBUG_ERROR_CLOSE()
            cmd_create_std.Dispose()

            Return False
        End Try

        cmd_create_std.Dispose()

        Return True

    End Function

    Function GET_NEW_STD_PRESC_DIR_ITEM_ID(ByVal i_institution As Int64) As Int64

        Dim l_id_std_instruction As Int64 = 0

        Dim sql As String = "select ALERT_PRODUCT_MT.SEQ_STD_PRESC_DIR_ITEM.NEXTVAL from dual"

        Dim cmd As New OracleCommand(sql, Connection.conn)
        cmd.CommandType = CommandType.Text

        Dim dr As OracleDataReader

        Try
            dr = cmd.ExecuteReader()

            DEBUGGER.SET_DEBUG("MEDICATION_API :: GET_NEW_STD_PRESC_DIR_ITEM_ID(" & i_institution & ")")

            Dim l_parameters(0) As String
            l_parameters(0) = "SEQ_STD_PRESC_DIR_ITEM"

            DEBUGGER.SET_RESPONSE("MEDICATION_API :: GET_NEW_STD_PRESC_DIR_ITEM_ID", l_parameters, dr)
            dr = cmd.ExecuteReader()

            While dr.Read()
                l_id_std_instruction = dr.Item(0)
            End While

            dr.Dispose()
            dr.Close()
            cmd.Dispose()

        Catch ex As Exception
            DEBUGGER.SET_DEBUG_ERROR_INIT("MEDICATION_API :: GET_NEW_STD_PRESC_DIR_ITEM_ID")
            DEBUGGER.SET_DEBUG(ex.Message)
            DEBUGGER.SET_DEBUG(sql)
            DEBUGGER.SET_DEBUG_ERROR_CLOSE()
            cmd.Dispose()
        End Try

        Return l_id_std_instruction

    End Function

    Function GET_COMPONENTS_LIST(ByVal i_institution As Int64, ByVal i_id_product As String, ByVal i_id_route_supplier As String, ByRef i_dr As OracleDataReader) As Boolean

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution)
        Dim sql As String = "  SELECT DISTINCT c.id_product_component,
                                                alert_product_tr.pk_product_utils.get_entity_desc(i_lang        => " & l_id_language & ",
                                                                                                  i_code_entity => nvl(lps.code_synonym, pc.code_product)) AS component_desc
                                  FROM alert_product_mt.product_medication pm
                                  LEFT JOIN alert_product_mt.product_med_component c
                                    ON c.id_product = pm.id_product
                                   AND c.id_product_supplier = pm.id_product_supplier
                                   AND pm.id_product_med_type = 2
                                  LEFT JOIN alert_product_mt.product pc
                                    ON pc.id_product = c.id_product_component
                                  LEFT JOIN alert_product_mt.lnk_product_synonym lps
                                    ON lps.id_product = pc.id_product
                                   AND lps.id_product_supplier = pc.id_product_supplier
                                 WHERE pm.id_product = '" & i_id_product & "'
                                   AND pm.id_product_supplier = '" & i_id_route_supplier & "'
                                   order by 2 asc "

        Dim cmd As New OracleCommand(sql, Connection.conn)
        Try
            cmd.CommandType = CommandType.Text
            i_dr = cmd.ExecuteReader()

            DEBUGGER.SET_DEBUG("MEDICATION_API :: GET_COMPONENTS_LIST(" & i_institution & ", " & i_id_product & ", " & i_id_route_supplier & ")")
            Dim l_parameters(1) As String
            l_parameters(0) = "id_product_component"
            l_parameters(1) = "component_desc"

            DEBUGGER.SET_RESPONSE("MEDICATION_API :: GET_COMPONENTS_LIST", l_parameters, i_dr)
            i_dr = cmd.ExecuteReader()

            cmd.Dispose()
            Return True
        Catch ex As Exception
            DEBUGGER.SET_DEBUG_ERROR_INIT("MEDICATION_API :: GET_COMPONENTS_LIST")
            DEBUGGER.SET_DEBUG(ex.Message)
            DEBUGGER.SET_DEBUG(sql)
            DEBUGGER.SET_DEBUG_ERROR_CLOSE()
            cmd.Dispose()
            Return False
        End Try

    End Function
End Class



