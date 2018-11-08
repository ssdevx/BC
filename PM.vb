Imports System.Data.OleDb
Module moduloPM
    Private cnConexion As OleDbConnection
    Private rs As OleDb.OleDbDataReader
    Dim da As OleDbDataAdapter
    Private strBD As String
    'Private intPeriodo As Integer = CInt(waBuroCredito.txtPeriodoActual.Text)
    Private intPeriodo As Integer
    Private dblCantidad As Double ' Suma cifra de control
    Dim rr As OleDbDataAdapter
    Dim oleComand As OleDb.OleDbCommand


    'strBD = Application.StartupPath & "\BC_PM.mdb"
    ' Programa que sirve para la creacion del archivo INTF del buró de Crédito
    ' Especificaciones:
    '   --  HD  Encabezado
    '   --  EM  Compania
    '   --  AC  Accionistas
    '   --  CR  Crédito
    '   --  DE  Detalles de credito
    '   --  AV Avales
    '   --  TS  Segmento de Cierre 

    Sub proAbrirArchivo()
        Dim strArchivo As String
        strArchivo = ""
        FileOpen(1, strArchivo, OpenMode.Binary)
    End Sub

    Sub proCerrarArchivo()
        FileClose(1)
    End Sub

    Sub proPrincipal()
        Dim dateFechaInicio As Date
        'Dim strConexion As String
        Dim intContador As Integer

        intPeriodo = CInt(waBuroCredito.txtPeriodoActual.Text)
        'strBD = "C:\Users\Usuario\Desktop\BC_Cinta\BC_PM.mdb"
        strBD = Application.StartupPath & "\BC_PM.mdb"
        cnConexion = New OleDb.OleDbConnection
        cnConexion.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source='" & strBD & "';Jet OLEDB:Database Password=;"
        cnConexion.Open()

        'oleComand = New OleDb.OleDbCommand("select * from qryCR_rfc", cnConexion)
        'oleComand = New OleDb.OleDbCommand("SELECT tbl_DE.DE_rfc As [CR_RFC],tbl_DE.DE_diasdevencimiento ,Sum(tbl_DE.DE_cantidad) As [DE_Cantidad],Count(*) As [CR_Experiencia],tbl_DE.fac_tipo_mon As [CR_Moneda],tbl_DE.FK_HD  FROM [tbl_DE] Group By tbl_DE.DE_rfc ,tbl_DE.DE_diasdevencimiento ,tbl_DE.fac_tipo_mon ,tbl_DE.FK_HD Order By tbl_DE.DE_rfc,tbl_DE.fac_tipo_mon,tbl_DE.DE_diasdevencimiento;", cnConexion)
        'oleComand.CommandType = CommandType.Text
        'rs = oleComand.ExecuteReader

        'da = New OleDb.OleDbDataAdapter("SELECT DISTINCT tbl_DE.DE_rfc As [CR_RFC],tbl_DE.DE_diasdevencimiento ,Sum(tbl_DE.DE_cantidad) As [DE_Cantidad],Count(*) As [CR_Experiencia],tbl_DE.fac_tipo_mon As [CR_Moneda],tbl_DE.FK_HD  FROM [tbl_DE] Group By tbl_DE.DE_rfc ,tbl_DE.DE_diasdevencimiento ,tbl_DE.fac_tipo_mon ,tbl_DE.FK_HD Order By tbl_DE.DE_rfc,tbl_DE.fac_tipo_mon,tbl_DE.DE_diasdevencimiento", cnConexion)
        da = New OleDb.OleDbDataAdapter("SELECT DISTINCT DE_RFC FROM tbl_DE WHERE FK_HD= @periodo", cnConexion)
        'da = New OleDb.OleDbDataAdapter("select * from qryCR_rfc", cnConexion)
        da.SelectCommand.Parameters.AddWithValue("@periodo", intPeriodo)

        Dim table As New DataTable()
        da.Fill(table)

        dateFechaInicio = Now
        'dblCantidad = 0
        intContador = 0
        For Each row As DataRow In table.Rows
            'proSegmento_EM(CStr(row("CR_rfc")))
            'proSegmento_CR(CStr(row("CR_rfc")))
            proSegmento_EM(CStr(row("DE_RFC")))
            proSegmento_CR(CStr(row("DE_RFC")))
            intContador += 1
        Next

        proSegmento_TS()
        'proCerrarArchivo()
        da.Dispose()
        cnConexion.Close()

        '   _____________________________________________________________________________________________________________________________

        cnConexion.Open()
        da = New OleDb.OleDbDataAdapter("select * from  tbl_TS where  FK_HD = @periodo", cnConexion)
        da.SelectCommand.Parameters.AddWithValue("@periodo", intPeriodo)
        Dim table1 As New DataTable()
        da.Fill(table1)

        Dim rw As DataRow
        rw = table1.Rows(0)


        waBuroCredito.txtCCMontos.Text = CStr(rw("TS_numerocompanias"))
        waBuroCredito.txtCCTotales.Text = CStr(rw("TS_cantidaddelsegmento"))

        waBuroCredito.txtCCMontos2.Text = CStr(intContador)
        waBuroCredito.txtCCTotales2.Text = CStr(dblCantidad)
        waBuroCredito.lblMensajes.Text = "Inició " & dateFechaInicio & " y terminó " & Now
        


        'While rs.Read
        '    MsgBox(rs(0).ToString)
        'End While






    End Sub

    'Public Sub proCifrasdeControl()
    '    Dim dateFechaInicio As Date
    '    dateFechaInicio = Now
    '    strBD = "C:\Users\Usuario\Desktop\BC_Cinta\BC_PM.mdb"
    '    cnConexion = New OleDb.OleDbConnection
    '    cnConexion.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source='" & strBD & "';Jet OLEDB:Database Password=;"
    '    cnConexion.Open()

    '    da = New OleDb.OleDbDataAdapter("select * from  tbl_TS where  FK_HD = 1", cnConexion)
    '    Dim table1 As New DataTable()
    '    da.Fill(table1)

    '    Dim rw As DataRow
    '    rw = table1.Rows(0)


    '    waBuroCredito.txtCCMontos.Text = CStr(rw("TS_numerocompanias"))
    '    waBuroCredito.txtCCTotales.Text = CStr(rw("TS_cantidaddelsegmento"))

    '    waBuroCredito.txtCCMontos2.Text = "" 'CStr(intContador)
    '    waBuroCredito.txtCCTotales2.Text = CStr(dblCantidad)
    '    waBuroCredito.lblMensajes.Text = "Inició " & dateFechaInicio & " y terminó " & Now
    '    'frmCifrasControl.txtCCMontos2 = CStr(intContador)
    '    'frmCifrasControl.txtCCTotales2 = CStr(dblCantidad)
    '    'frmCifrasControl.Show(vbModal)

    '    'DataRow(rw = table1.Rows(0))

    '    'fila = table1.Ro
    'End Sub


    '   --  G E N E R A R   S E G M E N T O S   [HD], EM, AC, CR, DE, AV, TS 
    '   --  Documento: ENTREGA DE BASE DE DATOS [informacion crediticia de personas morales y fisicas con actividad empresarial] Version 05


    '   --  HD  Encaabezado o Inicio
    Public Sub proSegmento_HD()
        Dim strCadena As String = ""
        Dim strFiller As String = "".PadLeft(52)

        'strBD = "C:\Users\Usuario\Desktop\BC_Cinta\BC_PM.mdb"
        strBD = Application.StartupPath & "\BC_PM.mdb"
        cnConexion = New OleDb.OleDbConnection
        cnConexion.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source='" & strBD & "';Jet OLEDB:Database Password=;"
        cnConexion.Open()

        oleComand = New OleDb.OleDbCommand("select * from tbl_HD where PK_HD = @periodo", cnConexion)
        da.SelectCommand.Parameters.AddWithValue("@periodo", intPeriodo)
        oleComand.CommandType = CommandType.Text
        rs = oleComand.ExecuteReader

        '   --  Suma longitudes de campos = 180 Bytes

        If rs.Read Then
            strCadena = "HDBNCPM"                                           '   --  Etiqueta Inicial
            strCadena += "00" & "0975"                                      '   --  Clave del BC para identificar el usuario
            strCadena += "01" & "0000"                                      '   --  Institucion anterior
            strCadena += "02" & "999"                                       '   --  Tipo Institucion (999 -> Comerciales)
            strCadena += "03" & "2"                                         '   --  Sumarizado para empresas comerciales
            strCadena += "04" & rs(1).ToString                              '   --  Checar que todas las opciones tenga longitud 8
            strCadena += "05" & rs(2).ToString                              '   --  PERIODO - Checar que todas las opciones tenga longitud 6
            strCadena += "06" & "05"                                        '   --  VERSION - Nuevo Etiqueta V5.0 Numero de version de formato de cinta
            strCadena += "07" & funRellenaconEspacios(75, "C CONECTIVIDAD") '   --  NOMBRE DEL OTORGANTE    - V5.0 Nombre corto asignado por el BC, 75 caracteres
            strCadena += "08" & strFiller                                   '   --  FILLER  - Campo reservado para uso futuro.
        End If
        rs.Close()
        cnConexion.Close()
        '   --  Relleno 
        Clipboard.SetText(strCadena)
    End Sub


    '   --  EM Datos Generales del Cliente.
    Public Sub proSegmento_EM(ByVal strCR_RFC As String)
        Dim strSQL As String = ""
        Dim strCadena As String = ""
        Dim strFiller As String = "".PadLeft(87)
        Dim str18 As String = "".PadLeft(18)
        Dim str02 As String = "".PadLeft(2)
        Dim str75 As String = "".PadLeft(75)
        Dim str25 As String = "".PadLeft(25)
        Dim str40 As String = "".PadLeft(40)
        Dim str08 As String = "".PadLeft(8)
        Dim str11 As String = "".PadLeft(11)
        Dim str30 As String = "".PadLeft(30)      '   --  Nueva Etiqueta V5.0
        Dim str150 As String = "".PadLeft(150)    '   --  Nueva Etiqueta V5.0

        'strBD = "C:\Users\Usuario\Desktop\BC_Cinta\BC_PM.mdb"
        strBD = Application.StartupPath & "\BC_PM.mdb"
        cnConexion = New OleDb.OleDbConnection
        cnConexion.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source='" & strBD & "';Jet OLEDB:Database Password=;"
        cnConexion.Open()

        'oleComand = New OleDb.OleDbCommand("Select top 1 * from tbl_EM where EM_RFC = '" & strCR_RFC & "' and FK_HD=1", cnConexion)
        'oleComand.CommandType = CommandType.Text
        'rs = oleComand.ExecuteReader

        da = New OleDb.OleDbDataAdapter("Select top 1 * from tbl_EM where EM_RFC = '" & strCR_RFC & "' and FK_HD= @periodo", cnConexion)
        da.SelectCommand.Parameters.AddWithValue("@periodo", intPeriodo)
        Dim table As New DataTable()
        da.Fill(table)
        For Each row As DataRow In table.Rows
            strCadena = "EMEM"                                                  '   --  Etiqueta Inicial
            strCadena += "00" & funRellenaconEspacios(13, CStr(row("EM_RFC")))  '   --  00  RFC [13]
            strCadena += "01" & str18                                           '   --  01  Cod Ciudadano (CURP en México) [18]
            strCadena += "02" & "0000000000"                                    '   --  02  Reservado para uso futuro [10]

            If CInt(row("EM_TIPOCLIENTE")) = 1 Or CInt(row("EM_TIPOCLIENTE")) = 4 Or CInt(row("EM_TIPOCLIENTE")) = 1 Then
                strCadena += "03" & funRellenaconEspacios(150, CStr(row("EM_COMPAÑIA")))
                strCadena += "04" & str30   '   --  04  Primer Nombre [30]
                strCadena += "05" & str30   '   --  05  Srgundo Nombre [30]
                strCadena += "06" & str25   '   --  06  Apellido Paterno [25]
                strCadena += "07" & str25   '   --  07  Apellido Materno [25]
            Else
                strCadena += "03" & str150

                If IsNothing(row("EM_NOMBRE1").ToString) Then
                    strCadena += "04" & str30
                Else
                    strCadena += "04" & funRellenaconEspacios(30, row("EM_NOMBRE1").ToString)           '   --  04  Primer Nombre [30]
                End If

                If IsNothing(row("EM_NOMBRE2").ToString) Then
                    strCadena += "05" & str30
                Else
                    strCadena += "05" & funRellenaconEspacios(30, row("EM_NOMBRE2").ToString)           '   --  05  Srgundo Nombre [30]
                End If

                If IsNothing(row("EM_APELLIDOPATERNO").ToString) Then
                    strCadena += "06" & str25
                Else
                    strCadena += "06" & funRellenaconEspacios(25, row("EM_APELLIDOPATERNO").ToString)   '   --  06  Apellido Paterno [25]
                End If

                If IsNothing(row("EM_APELLIDOMATERNO").ToString) Then
                    strCadena += "07" & str25
                Else
                    strCadena += "07" & funRellenaconEspacios(25, row("EM_APELLIDOMATERNO").ToString)   '   --  07  Apellido Materno [25]
                End If
            End If

            strCadena += "08" & str02                                                   '   --  08  Nacionalidad [2]
            strCadena += "09" & str02                                                   '   --  09  Calificacion de cartera [2]
            strCadena += "10" & "00008888888"                                           '   --  10  Actividad Economica [11]
            strCadena += "11" & "00000000000"                                           '   --  11  Actividad Economica [11]
            strCadena += "12" & "00000000000"                                           '   --  12  Actividad Economica [11]
            strCadena += "13" & funRellenaconEspacios(40, row("EM_DIRECCION1").ToString) '   --  13  Primera linea de direccion  [40]

            If IsNothing(row("EM_DIRECCION2").ToString) Then                             '   --  14  Segunda linea de direccion [40]
                strCadena += "14" & str40
            Else
                strCadena += "14" & funRellenaconEspacios(40, row("EM_DIRECCION2").ToString)
            End If

            strCadena += "15" & funRellenaconEspacios(60, row("EM_COLONIA").ToString)   '   --  15  Colonia o Poblacion  | 60
            strCadena += "16" & funRellenaconEspacios(40, row("EM_DELMUN").ToString)    '   --  16  Delegacion o municipio | 40
            strCadena += "17" & funRellenaconEspacios(40, row("EM_CIUDAD").ToString)    '   --  17  Ciudad | 40
            strCadena += "18" & funRellenaconEspacios(4, row("EM_ESTADO").ToString)     '   --  18  Estado para Dom Mexico | 4
            strCadena += "19" & funRellenaconEspacios(10, row("EM_CP").ToString)        '   --  19  Codigo Postal | 10
            strCadena += "20" & str11                           '   --  20  Numero de telefono | 11
            strCadena += "21" & str08                           '   --  21  Extensio | 8
            strCadena += "22" & str11                           '   --  22  Fax  | 11
            strCadena += "23" & row("EM_TIPOCLIENTE").ToString  '   --  23  Tipo de Cliente |    1
            strCadena += "24" & str40                           '   --  24  Nombre de Edo Ext   |   40
            strCadena += "25" & "MX"                            '   --  25  Pais de Origen |    2
            strCadena += "26" & "00000000"                      '   --  26  Clave de Consolidadcion |    8
            strCadena += "27" & strFiller                       '   --  27  Filler  | 87
        Next
        'Clipboard.SetText(strCadena)
        da.Dispose()
        cnConexion.Close()

        '   --  Suma longitudes de campos 800 Bytes.
        'While rs.Read()
        '    strCadena = "EMEM"                                                  '   --  Etiqueta Inicial
        '    strCadena += "00" & rs(2).ToString                                  '   --  00  RFC [13]
        '    strCadena += "01" & str18                                           '   --  01  Cod Ciudadano (CURP en México) [18]
        '    strCadena += "02" & "0000000000"                                    '   --  02  Reservado para uso futuro [10]

        '    '   --  Cuando es  Persona Moral (2), Fondo de Fideocimiso (3) o Gobierno (4)
        '    If CInt(rs(16)) = 1 Or CInt(rs(16)) = 4 Or CInt(rs(16)) = 3 Then
        '        strCadena += "03" & funRellenaconEspacios(150, rs(3).ToString)  '   --  03  Nombre compania [150]
        '        strCadena += "04" & str30                                       '   --  04  Primer Nombre [30]
        '        strCadena += "05" & str30                                       '   --  05  Srgundo Nombre [30]
        '        strCadena += "06" & str25                                       '   --  06  Apellido Paterno [25]
        '        strCadena += "07" & str25                                       '   --  07  Apellido Materno [25]
        '    Else    '   --  Cuando es PFAE (Persona Fisica con Actividad Empresarial) 1
        '        strCadena += "03" & str150                                      '   --  03  Nombre compania [150]     

        '        If IsNothing(rs(4).ToString) Then                               '   --  04  Primer Nombre [30]
        '            strCadena += "04" & str30
        '        Else
        '            strCadena += "04" & funRellenaconEspacios(30, rs(4).ToString)
        '        End If

        '        If IsNothing(rs(5).ToString) Then                               '   --  05  Srgundo Nombre [30]
        '            strCadena += "05" & str30
        '        Else
        '            strCadena += "05" & funRellenaconEspacios(30, rs(5).ToString)
        '        End If

        '        If IsNothing(rs(6).ToString) Then                               '   --  06  Apellido Paterno [25]
        '            strCadena += "06" & str25
        '        Else
        '            strCadena += "06" & funRellenaconEspacios(25, rs(6).ToString)
        '        End If

        '        If IsNothing(rs(7).ToString) Then                               '   --  07  Apellido Materno [25]
        '            strCadena += "07" & str25
        '        Else
        '            strCadena += "07" & funRellenaconEspacios(25, rs(7).ToString)
        '        End If
        '    End If

        '    strCadena += "08" & str02                                       '   --  08  Nacionalidad [2]
        '    strCadena += "09" & str02                                       '   --  09  Calificacion de cartera [2]
        '    strCadena += "10" & "00008888888"                               '   --  10  Actividad Economica [11]
        '    strCadena += "11" & "00000000000"                               '   --  11  Actividad Economica [11]
        '    strCadena += "12" & "00000000000"                               '   --  12  Actividad Economica [11]
        '    strCadena += "13" & funRellenaconEspacios(40, rs(8).ToString)   '   --  13  Primera linea de direccion  [40]

        '    If IsNothing(rs(9).ToString) Then                               '   --  14  Segunda linea de direccion [40]
        '        strCadena += "14" & str40
        '    Else
        '        strCadena += "14" & funRellenaconEspacios(40, rs(9).ToString)
        '    End If

        '    strCadena += "15" & funRellenaconEspacios(60, rs(12).ToString)  '   --  15  Colonia o Poblacion  | 60
        '    strCadena += "16" & funRellenaconEspacios(40, rs(13).ToString)  '   --  16  Delegacion o municipio | 40
        '    strCadena += "17" & funRellenaconEspacios(40, rs(14).ToString)  '   --  17  Ciudad | 40
        '    strCadena += "18" & funRellenaconEspacios(4, rs(15).ToString)   '   --  18  Estado para Dom Mexico | 4
        '    strCadena += "19" & funRellenaconEspacios(10, rs(10).ToString)  '   --  19  Codigo Postal | 10
        '    strCadena += "20" & str11                                       '   --  20  Numero de telefono | 11
        '    strCadena += "21" & str08                                       '   --  21  Extensio | 8
        '    strCadena += "22" & str11                                       '   --  22  Fax  | 11
        '    strCadena += "23" & rs(16).ToString                             '   --  23  Tipo de Cliente |    1
        '    strCadena += "24" & str40                                       '   --  24  Nombre de Edo Ext   |   40
        '    strCadena += "25" & "MX"                                        '   --  25  Pais de Origen |    2
        '    strCadena += "26" & "00000000"                                  '   --  26  Clave de Consolidadcion |    8
        '    strCadena += "27" & strFiller                                   '   --  27  Filler  | 87
        'End While
        'rs.Close()
        'cnConexion.Close()
        'Clipboard.SetText(strCadena)
        'MsgBox("Se ha copiado correctamente ")
    End Sub


    '   --  Pendiente....
    Public Sub proSegmento_AC()
        Dim strCadena As String = ""
        Dim strFiller As String = ""
        Dim str150 As String = ""       '   --  Nueva etiqueta V5.0
        Dim str11 As String = ""        '   --  Nueva etiqueta V5.0
        Dim str8 As String = ""         '   --  Nueva etiqueta V5.0
        Dim str40 As String = ""        '   --  Nueva etiqueta V5.0
        Dim str18 As String = ""        '   --  Nueva etiqueta V5.0

        ' strBD = "C:\Users\Usuario\Desktop\BC_Cinta\BC_PM.mdb"
        strBD = Application.StartupPath & "\BC_PM.mdb"
        cnConexion = New OleDb.OleDbConnection
        cnConexion.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source='" & strBD & "';Jet OLEDB:Database Password=;"
        cnConexion.Open()

        'oleComand = New OleDb.OleDbCommand("select * from tbl_AC where fk_hd = 1", cnConexion)
        'oleComand.CommandType = CommandType.Text
        'rs = oleComand.ExecuteReader

        da = New OleDb.OleDbDataAdapter("select * from tbl_AC where fk_hd = @periodo", cnConexion)
        da.SelectCommand.Parameters.AddWithValue("@periodo", intPeriodo)
        Dim table As New DataTable()
        da.Fill(table)
        For Each row As DataRow In table.Rows
            strCadena = "ACAC"
            strCadena += "00" & CStr(row("Accionista_RFC")) '   --  RFC Accionista
            strCadena += "01" & str18                       '   --  CURP (Opcional)
            strCadena += "02" & "0000000000"                '   --  Numero DUN (Opcional)
            strCadena += "03" & funRellenaconEspacios(150, CStr(row("Accionista_NombreCompañiaAccionista")))    '  --  Nombre del Accionista
            strCadena += "04" & funRellenaconEspacios(30, CStr(row("Accionista_Nombre1")))                      '   --  Nombre del accionista
            strCadena += "05" & funRellenaconEspacios(30, CStr(row("Accionista_Nombre2")))                      '   --  Nombre 2 del accionista
            strCadena += "06" & funRellenaconEspacios(25, CStr(row("Accionista_ApellidoPaterno")))              '   --  Apellido paterno
            strCadena += "07" & funRellenaconEspacios(25, CStr(row("Accionista_ApellidoMaterno")))              '   --  Apellido Materno
            strCadena += "08" & funRellenaconCeros(2, CLng(row("Accionista_Porcentaje")))                       '   --  Porcentaje
            strCadena += "09" & funRellenaconEspacios(40, CStr(row("Accionista_Domicilio1")))                   '   --  Direccion 1
            strCadena += "10" & str40                                                                           '   --  Direccion 2(opcional)
            strCadena += "11" & funRellenaconEspacios(60, CStr(row("Accionista_ColoniaPoblacion")))             '   --  Colonia /poblacion (opcional)
            strCadena += "12" & funRellenaconEspacios(40, CStr(row("Accionista_DelegacionMunicipio")))          '   --  Delegacion o Municipio
            strCadena += "13" & funRellenaconEspacios(40, CStr(row("Accionista_Ciudad")))                       '   --  Ciudad
            strCadena += "14" & funRellenaconEspacios(4, CStr(row("Accionista_Estado")))                        '   --  Estados de México (opcional)
            strCadena += "15" & funRellenaconEspacios(10, CStr(row("Accionista_CP")))                           '   --  Codigo postal
            strCadena += "16" & str11                               '   --  Teléfono (opcional)
            strCadena += "17" & str8                                '   --  Extensión opcional
            strCadena += "18" & str11                               '   --  Fax (opcional)
            strCadena += "19" & CInt(row("Accionista_TipoCliente")) '   --  Tipo de Cliente
            strCadena += "20" & str40                               '   --  Estado en el extranjero
            strCadena += "21" & "MX"                                '   --  MX Predeterminado
            strCadena += "22" & strFiller
        Next
        'ipboard.SetText(strCadena)
        da.Dispose()
        cnConexion.Close()

        'strCadena = ""
        'While rs.Read
        '    strCadena = "ACAC"
        '    strCadena += "00" & rs(2).ToString                              '   --  RFC Accionista
        '    strCadena += "01" & str18                                       '   --  CURP (Opcional)    
        '    strCadena += "02" & "0000000000"                                '   --  Numero DUN (Opcional)
        '    strCadena += "03" & funRellenaconEspacios(150, rs(3).ToString)  '   -- Nombre del Accionista (PM)
        '    strCadena += "04" & funRellenaconEspacios(30, rs(4).ToString)   '   --  Nombre del Accionista (PFAE)
        '    strCadena += "05" & funRellenaconEspacios(30, rs(5).ToString)   '   --  Nombre 2 del Accionista (PFAE)
        '    strCadena += "06" & funRellenaconEspacios(25, rs(6).ToString)   '   --  Apellido Paterno
        '    strCadena += "07" & funRellenaconEspacios(25, rs(7).ToString)   '   --  Apellido Materno
        '    strCadena += "08" & funRellenaconCeros(2, CLng(rs(8).ToString))    '   --  Porcentaje
        '    '   --  09
        '    '   --  10                                                                                                                          
        '    '   --  11
        '    strCadena += "12" & funRellenaconEspacios(40, rs(9).ToString)   '   --  Delegacion o Municipio
        '    strCadena += "13" & funRellenaconEspacios(40, rs(10).ToString)  '   --  Ciudad
        '    '   --  14  
        '    '   --  15
        '    strCadena += "16" & str11                                       '   --  Telefono (Opcional)
        '    strCadena += "17" & str8                                        '   --  Extension (Opcional)
        '    strCadena += "18" & str11                                       '   --  Fax (Opcional)
        '    strCadena += "19" & rs(11).ToString                             '   --  Tipo Cliente
        '    strCadena += "20" & str40                                       '   --  Estado en el extranjero
        '    strCadena += "21" & "MX"                                        '   --  MX Predeterminado
        '    strCadena += "22" & strFiller                                   '   --   Filler de 40
        '    'Clipboard.SetText(strCadena.ToString)
        '    MsgBox(strCadena)
        'End While
        'rs.Close()
        'cnConexion.Close()
    End Sub

    '   --  CR  Credito
    Public Sub proSegmento_CR(ByVal strRFC As String)
        Dim strCadena As String = ""
        Dim strFiller As String = "".PadLeft(40)
        Dim strFiller2 As String = "".PadLeft(53)
        Dim str25 As String = "".PadLeft(25)

        'strBD = "C:\Users\Usuario\Desktop\BC_Cinta\BC_PM.mdb"
        strBD = Application.StartupPath & "\BC_PM.mdb"
        cnConexion = New OleDb.OleDbConnection
        cnConexion.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source='" & strBD & "';Jet OLEDB:Database Password=;"
        cnConexion.Open()

        'oleComand = New OleDb.OleDbCommand("select * from qryCR where CR_rfc = '" & strRFC & "'" & " and FK_HD=1", cnConexion)
        'oleComand.CommandType = CommandType.Text
        'rs = oleComand.ExecuteReader
        '   --  Suma longitudes de campos  = 400 Bytes.
        'da = New OleDb.OleDbDataAdapter("select * from qryCR where CR_rfc = '" & strRFC & "'" & " and FK_HD=1", cnConexion)
        da = New OleDb.OleDbDataAdapter("SELECT tbl_DE.DE_rfc As [CR_RFC],tbl_DE.DE_diasdevencimiento ,Sum(tbl_DE.DE_cantidad) As [DE_Cantidad],Count(*) As [CR_Experiencia],tbl_DE.fac_tipo_mon As [CR_Moneda],tbl_DE.FK_HD  FROM [tbl_DE] where tbl_DE.DE_rfc = @rfc and FK_HD = @periodo Group By tbl_DE.DE_rfc ,tbl_DE.DE_diasdevencimiento ,tbl_DE.fac_tipo_mon ,tbl_DE.FK_HD Order By tbl_DE.DE_rfc,tbl_DE.fac_tipo_mon,tbl_DE.DE_diasdevencimiento", cnConexion)
        'da = New OleDb.OleDbDataAdapter("SELECT tbl_DE.DE_rfc As [CR_RFC],tbl_DE.DE_diasdevencimiento ,Sum(tbl_DE.DE_cantidad) As [DE_Cantidad],Count(*) As [CR_Experiencia],tbl_DE.fac_tipo_mon As [CR_Moneda],tbl_DE.FK_HD  FROM [tbl_DE] where tbl_DE.DE_rfc = '" & strRFC & "'" & " and FK_HD = " & CStr(intPeriodo), cnConexion)
        'da = New OleDb.OleDbDataAdapter("SELECT tbl_DE.DE_rfc As [CR_RFC],tbl_DE.DE_diasdevencimiento ,Sum(tbl_DE.DE_cantidad) As [DE_Cantidad],Count(*) As [CR_Experiencia],tbl_DE.fac_tipo_mon As [CR_Moneda],tbl_DE.FK_HD  FROM [tbl_DE] where tbl_DE.DE_rfc = @rfc and FK_HD = o) & "'" & " Group By tbl_DE.DE_rfc ,tbl_DE.DE_diasdevencimiento ,tbl_DE.fac_tipo_mon ,tbl_DE.FK_HD Order By tbl_DE.DE_rfc,tbl_DE.fac_tipo_mon,tbl_DE.DE_diasdevencimiento", cnConexion)

        da.SelectCommand.Parameters.AddWithValue("@rfc", strRFC)
        da.SelectCommand.Parameters.AddWithValue("@periodo", intPeriodo)

        Dim table As New DataTable()
        da.Fill(table)
        For Each row As DataRow In table.Rows
            strCadena = "CRCR"
            strCadena += "00" & funRellenaconEspacios(13, CStr(row("CR_rfc")))       '   --  rfc
            strCadena += "01" & funRellenaconCeros(6, CLng(row("CR_experiencia")))   '   --  experiencias
            strCadena += "02" & str25                        '   --  No aplica (25 espacios) numero de contrato
            strCadena += "03" & str25                        '   --  No aplica (25 espacios) número de contrato anterior
            strCadena += "04" & funRellenaconCeros(8, 0)     '   --  fecha de apertura (no requerido)
            strCadena += "05" & funRellenaconCeros(6, 0)     '   --  Plazo (no requerido)
            strCadena += "06" & funRellenaconCeros(4, 0)     '   --  Tipo de crédito (no aplica)
            strCadena += "07" & funRellenaconCeros(20, 0)    '   --  saldo inicial (no aplica)

            If Trim(CStr(row("CR_moneda"))) = "MN" Then strCadena += "08" & "001" '   --  Moneda Numerico 3
            If Trim(CStr(row("CR_moneda"))) = "USD" Then strCadena += "08" & "005" '   --  Moneda Numerico 3

            strCadena += "09" & funRellenaconCeros(4, 0)     '   --  Número de pagos (no aplica)
            strCadena += "10" & funRellenaconCeros(5, 0)     '   --  **** Modificacion v 5.0 'Frecuencia de Pagos (no aplica)
            strCadena += "11" & funRellenaconCeros(20, 0)    '   --  Importe de Pagos (no aplica)
            strCadena += "12" & funRellenaconCeros(8, 0)     '   --  Fecha última de pago (no aplica)
            strCadena += "13" & funRellenaconCeros(8, 0)     '   --  Fecha de reestructura (no aplica)
            strCadena += "14" & funRellenaconCeros(20, 0)    '   --  Pago en efectivo (no aplica)
            strCadena += "15" & funRellenaconCeros(8, 0)     '   --  Fecha de liquidación (no aplica)
            strCadena += "16" & funRellenaconCeros(20, 0)    '   --  Monto de la quita (no aplica)
            strCadena += "17" & funRellenaconCeros(20, 0)    '   --  Dación de pago (no aplica)
            strCadena += "18" & funRellenaconCeros(20, 0)    '   --  Quebranto o castigo
            strCadena += "19" & "    "       '   --  Clave de observación (no aplica)
            strCadena += "20" & " "          '    --  Especiales (no aplica)
            strCadena += "21" & funRellenaconCeros(8, 0)     '   --  **** Nuevo v5.0 Fecha de primer incumplimiento (no aplica)
            strCadena += "22" & funRellenaconCeros(20, 0)    '   --  **** Nuevo v5.0 Saldo insuluto principal (no aplica)
            strCadena += "23" & funRellenaconCeros(20, 0)    '   --  **** Nuevo v5.0 Crédito Máximo utilizado (no aplica)
            strCadena += "24" & funRellenaconCeros(8, 0)     '   --  **** Nuevo v5.0 Fecha de Ingreso (no aplica)
            strCadena += "25" & strFiller


            '   --  DE
            strCadena = ""
            '***********************************  150 bytes en total ******************************
            strCadena = "DEDE"
            strCadena += "00" & funRellenaconEspacios(13, CStr(row("CR_rfc")))           '   --  rfc falta alinear a la izquierda y rellenar
            strCadena += "01" & str25                                                    '    --  Numero de Contrato(no aplica)
            strCadena += "02" & funRellenaconCeros(3, CLng(row("DE_diasdevencimiento"))) '   --  Dias de Vencimiento
            strCadena += "03" & funRellenaconCeros(20, CLng(row("DE_cantidad")))   ' Cantidad
            strCadena += "04" & funRellenaconCeros(20, 0)       '**** Nuevo v5.0 Interes que corresponden al credito (opcional)
            strCadena += "05" & strFiller2
            dblCantidad = dblCantidad + CLng(row("DE_cantidad"))
            strCadena = ""

            'MsgBox(dblCantidad)
        Next
        da.Dispose()
        cnConexion.Close()
        'Clipboard.SetText(strCadena)

        'While rs.Read()
        '    strCadena = "CRCR"
        '    strCadena += "00" & funRellenaconEspacios(13, rs(0).ToString)    '   --  00  RFC del Acreditado  [13]
        '    strCadena += "01" & funRellenaconCeros(6, CLng(rs(1).ToString)) '   --  01  Num de Experiencias crediticias [6]
        '    strCadena += "02" & str25                                       '   --  02  Num de Credito o Contrato   [25]    
        '    strCadena += "03" & str25                                       '   --  03  Num de cuenta, credito o contrato anterior  [25]
        '    strCadena += "04" & funRellenaconCeros(8, 0)                    '   --  04  Fecha de apertura del credito   [8]
        '    strCadena += "05" & funRellenaconCeros(6, 0)                    '   --  05  Plazo en meses  [6]
        '    strCadena += "06" & funRellenaconCeros(4, 0)                    '   --  06  Tipos de crédito    [4]
        '    strCadena += "07" & funRellenaconCeros(20, 0)                   '   --  07  Monto Autorizado (saldo inicial)   [20]
        '    If Trim(CStr(rs(2))) = "MN" Then strCadena += "08" & "001"
        '    If Trim(CStr(rs(2))) = "USD" Then strCadena += "08" & "005" '   --  08  Moneda  [3]
        '    strCadena += "09" & funRellenaconCeros(4, 0)                    '   --  09  Num de pagos    [4]
        '    strCadena += "10" & funRellenaconCeros(5, 0)                    '   --  10  Frecuencia de pagos [5]
        '    strCadena += "11" & funRellenaconCeros(20, 0)                   '   --  11  Importe de pago [20]
        '    strCadena += "12" & funRellenaconCeros(8, 0)                    '   --  12  Fecha de ultimo pago    [8]
        '    strCadena += "13" & funRellenaconCeros(8, 0)                    '   --  13  Fecha de reestructura   [8]
        '    strCadena += "14" & funRellenaconCeros(20, 0)                   '   --  14  Pago Final para cierre de cuenta morosa (En efectivo) [20]
        '    strCadena += "15" & funRellenaconCeros(8, 0)                    '   --  15  Fecha de liquidacion    [8]
        '    strCadena += "16" & funRellenaconCeros(20, 0)                   '   --  16  Quita   [20]
        '    strCadena += "17" & funRellenaconCeros(20, 0)                   '   --  17  Dacion de pago  [20]
        '    strCadena += "18" & funRellenaconCeros(20, 0)                   '   --  18  Quebranto o Castigo [20]
        '    strCadena += "19" & "    "                                      '   --  19  Clave de observacion    [4]
        '    strCadena += "20" & " "                                         '   --  20  Marca para credito especial [1]
        '    strCadena += "21" & funRellenaconCeros(8, 0)                    '   --  21  Fecha primer incumplimiento [8]
        '    strCadena += "22" & funRellenaconCeros(20, 0)                   '   --  22  Saldo insoluto del principal [20]
        '    strCadena += "23" & funRellenaconCeros(20, 0)                   '   --  23  Crédito máximo utilizado    [20]
        '    strCadena += "24" & funRellenaconCeros(8, 0)                    '   --  24  Fecha ingreso a cartera vencida [8]
        '    strCadena += "25" & strFiller                                   '   --  25  Filler  [40]
        '   --  Escribir en TXT


        '   --  Suma longitudes de Campo = 150 Bytes
        'strCadena = ""
        'strCadena = "DEDE"
        'strCadena += "00" & funRellenaconEspacios(13, rs(1).ToString)   '   --  00  RFC del Acreditado  [13]
        'strCadena += "01" & str25                                       '   --  01  Numero de cuenta, credito o contrato [25]
        'strCadena += "02" & funRellenaconCeros(3, CLng(rs(3).ToString)) '   --  02  Numeros de dias vencidos    [3]
        'strCadena += "03" & funRellenaconCeros(20, CLng(rs(4).ToString)) '   --  03  Cantidad (saldo)    [20]
        'strCadena += "00" & funRellenaconCeros(20, 0)                   '   --  04  Interes [20]
        'strCadena += "00" & strFiller2                                  '   --  05  Filler  [53]
        'dblCantidad = dblCantidad + CDbl(rs(5))
        'End While
        'rs.Close()
        'cnConexion.Close()
        'Clipboard.SetText(strCadena)
    End Sub

    Public Sub proSegmento_AV() '   --  Aval, aun no implementado

        '   --  Suma de longitudes de campos 750 Bytes.

        '   --  00  RFC Del Aval [13]
        '   --  01  Cod Ciudadano (CURP en México) [18]
        '   --  02  Campo reservado [10]
        '   --  03  Nomnbre de la compania [150]
        '   --  04  Primer nombre [30]
        '   --  05  Segundo nombre [30]
        '   --  06  Apellido paterno [25]
        '   --  07  Apellido materno [25]
        '   --  08  Primer linea de direccion [40]
        '   --  09  Segunda linea de direccion [40]
        '   --  10  Colonia o poblacion [60]
        '   --  11  Delegacion o municipio [40]
        '   --  12  Ciudad [40]
        '   --  13  Nombe estado (Mexico) [4]
        '   --  14  Codigo Postal [10]
        '   --  15  Num de telefono [11]
        '   --  16  Extension Tel [8]
        '   --  17  Numero de fax [11]
        '   --  18  Tipo de aval [1]
        '   --  19  Nombre estado (extranjero) [40]
        '   --  20  Pais de Origen Domicilio [2]
        '   --  21  Filler [94] 
    End Sub

    Public Sub proSegmento_TS()
        Dim strCadena As String = ""
        Dim strFiller As String = "".PadLeft(53)

        ' strBD = "C:\Users\Usuario\Desktop\BC_Cinta\BC_PM.mdb"
        strBD = Application.StartupPath & "\BC_PM.mdb"
        cnConexion = New OleDb.OleDbConnection
        cnConexion.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source='" & strBD & "';Jet OLEDB:Database Password=;"
        cnConexion.Open()

        'oleComand = New OleDb.OleDbCommand("select * from tbl_TS where FK_HD=1", cnConexion)
        'oleComand.CommandType = CommandType.Text
        'rs = oleComand.ExecuteReader

        da = New OleDb.OleDbDataAdapter("select * from tbl_TS where FK_HD= @periodo", cnConexion)
        da.SelectCommand.Parameters.AddWithValue("@periodo", intPeriodo)
        Dim table As New DataTable()
        da.Fill(table)

        Dim rw As DataRow
        rw = table.Rows(0)

        'For Each row As DataRow In table.Rows
        strCadena = "TSTS"
        strCadena += "00" & funRellenaconCeros(7, CLng(rw("TS_numerocompanias")))
        strCadena += "01" & funRellenaconCeros(30, CLng(rw("TS_cantidaddelsegmento")))
        strCadena += "02" & strFiller
        'Next

        'While rs.Read()
        '    strCadena = "TSTS"
        '    strCadena += "00" & funRellenaconCeros(7, CLng(rs(1))) '   --  00  Numeros de compañias reportadas [7]
        '    strCadena += "01" & funRellenaconCeros(30, CLng(rs(2))) '   --  01  Total de "Cantidad de Saldo"    [30]
        '    strCadena += "02" & strFiller  '   --  02  Filler  [53]
        'End While
        da.Dispose()
        cnConexion.Close()
        ' Clipboard.SetText(strCadena)
    End Sub
End Module
