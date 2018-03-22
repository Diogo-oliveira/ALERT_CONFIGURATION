Imports Oracle.DataAccess.Client
Public Class Medication_API

    Dim db_access_general As New General

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
                               AND i.id_institution = " & i_ID_INST

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

    Function GET_PRODUCT_OPTIONS(ByVal i_institution As Int64, ByVal i_id_product As String, ByVal i_product_supplier As String, ByVal i_id_pick_list As Int16, ByRef i_dr As OracleDataReader) As Boolean

        DEBUGGER.SET_DEBUG("MEDICATION_API :: GET_PRODUCT_OPTIONS(" & i_institution & ", " & i_id_product & ", " & i_product_supplier & ")")

        Dim l_id_language As Int16 = db_access_general.GET_ID_LANG(i_institution)

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
                               eds.desc_lang_8 as product_synonym
                          FROM alert_product_mt.product p
                          JOIN alert_product_mt.product_medication pm
                            ON pm.id_product = p.id_product
                           AND pm.id_product_supplier = p.id_product_supplier
                          LEFT JOIN alert_product_mt.lnk_product_synonym lps
                            ON lps.id_product = p.id_product
                           AND lps.id_product_supplier = p.id_product_supplier
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

    Function SET_PARAMETERS(ByVal i_institution As Int64, ByVal i_id_product As String, ByVal i_id_product_supplier As String, i_flg_available As String, i_id_pick_list As Int16, i_id_product_level As Int16) As Boolean

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

End Class
