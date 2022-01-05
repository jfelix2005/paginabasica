
<%Response.AddHeader "Content-Type", "text/html;charset=ISO-8859-1"%>

<%
Response.Expires = -3000
Response.Buffer = True
Server.ScriptTimeout=360

'Date dd/mm/yyyy
Session.lcid= 2057 '= UK English
'On Error Resume Next
wFechadoc = "" ' Variable para validar formato de fecha del documento
wFechaIng = "" ' Variable para validar formato de fecha de ingreso
wFechaCese = ""
wNumeroRegistro = ""
wHoraRegistro = ""
wUnidad = ""
wCantidad = ""

%>

<%
Acc = Request("Acc")
Modal = 0
if Acc = "New" then
    Id_Unidad = Request("Id_Unidad")
    wId_Usuario= Request("Id_Usuario")

    wEmpresa = Request("Empresa")
    wGrupo = Request("Grupo")
    'w_cm_CargaMasiva_id = 3 '- ID de Estructura de Carga Masiva 
else
	'Variables
	Dim mySmartUpload, wObj,MiFSO
	Dim file,wErrorFile

	'Object creation
	'***************
    Set mySmartUpload = Server.CreateObject("aspSmartUpload.SmartUpload")

	'Upload
	'******
	mySmartUpload.MaxFileSize=50000000
	mySmartUpload.Upload
    Id_Unidad = mySmartUpload.Form("Id_Unidad")		
    wId_Usuario = mySmartUpload.Form("Id_Usuario")	
    
    wEmpresa = mySmartUpload.Form("Empresa")	
    wGrupo =  mySmartUpload.Form("wGrupo")
    wcodigo_ticket = mySmartUpload.Form("Incidencia")	

end if

if not isObject(oConn) then
    Set oConn = Server.CreateObject("ADODB.Connection")		
    strConnQuery = Application(wEmpresa)
    oConn.Open(strConnQuery)		
end if  
   Set wRsIncidencia = Server.CreateObject("ADODB.recordset")
	wSQL = "SELECT "
	wSQL = wSQL + " inc_incidencia_id, "
	wSQL = wSQL + " codigo_ticket"
	wSQL = wSQL + " FROM inc_incidencia "
	wSQL = wSQL + " Where fb_empleado_id =(select fb_empleado_id from sc_user where "    
	wSQL = wSQL + " sc_user_id = " & wId_Usuario & ")"
	wRsIncidencia.Open wSQL, oConn
   
'-- VARIABLES IMPORTANTES
wIP_Address = Request.ServerVariables("remote_addr")
wSession_Id = Session.SessionID
wMensajeErrorStore =""
strMensajeOK = ""  
'CargaExito = "0"
wMensajeError = ""

if Acc <> "New" then

wHoraArchivo = replace(mid(Time(),1,8),":","")+Id_Unidad
    '-- ARMA nombre del archivo que se grabara con el archivo UPLOAD
    NameFile = "" '"carga_almacenamiento_temporal.xls"
    wErrorFile = ""
    wtxtArchivoCSV = ""
	
    if mySmartUpload.Files.TotalBytes > 0 Then
        wtxtNameFile = mySmartUpload.Files.Item(1).FileName	
        wtxtTamanoFile = mySmartUpload.Files.Item(1).Size 
    end if	
    NameFile = wHoraArchivo & wtxtNameFile
   '-- Establece en una variable la ruta del archivo
    StrFile= Request.servervariables("APPL_PHYSICAL_PATH")  & "intrasolution_carga\Files\" & wEmpresa & "\" & NameFile

    ExtFile = ucase(right(wtxtNameFile,3))      
 
    If (mySmartUpload.Files.TotalBytes <= 500000000) and ExtFile = "XLS" Then
	    NameFile = wHoraArchivo & wtxtNameFile '"carga_almacenamiento_temporal.XLS"

	    StrFile= Request.servervariables("APPL_PHYSICAL_PATH")  & "intrasolution_carga\Files\" & wEmpresa & "\" & NameFile

        set MiFSO = Server.CreateObject("Scripting.FileSystemObject")
        i = 0
        For each file In mySmartUpload.Files
            i = i + 1
            If file.size > 0 Then
                   file.SaveAs(StrFile)                     
            End If 
        Next
    Else   
        wErrorFile="EXTENSION_ERROR"
    End if

    if wErrorFile = "" then
    
        Set mySmartUpload=Nothing
        set MiFSO =Nothing

        'Crea una conexion al Excel y recupera la data en un Recordset, dependiendo del caso replica el nombre del archivo guardado
        if ExtFile = "XLS" then
	        NameFile = wHoraArchivo & wtxtNameFile '"carga_almacenamiento_temporal.xls"
	        StrFile= Request.servervariables("APPL_PHYSICAL_PATH")  & "intrasolution_carga\Files\" & wEmpresa & "\" & NameFile
	        Set cnADODBConnection = Server.CreateObject("ADODB.Connection")
	        cnADODBConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & StrFile & ";" & "Extended Properties=""Excel 8.0;IMEX=1;HDR=YES;"""

	        Set objRS = Server.CreateObject("ADODB.Recordset")
	         objRS.ActiveConnection = cnADODBConnection
	         objRS.CursorType = 3 'Static cursor.
	         objRS.LockType = 2 'Pessimistic Lock.
            
	         sql = "Select * from [A1:Y3000]" 
	         objRS.Source = sql
	         objRS.Open
          
        end if     

        'Crea Conexion a la Base de Datos SQL
        if not isObject(oConn) then
            Set oConn = Server.CreateObject("ADODB.Connection")		
            strConnQuery = Application(wEmpresa)
            oConn.Open(strConnQuery)		
        end if  
                       
        
        'Carga data del excel a tabla temporal
        CargaExito = "3"

        Call CargaTablaTemporal()
        
        'Si Carga data no tiene errores continua
        	
        Acc = "New"
    End if 
end if

%>

<html>

<head>
<meta http-equiv="Content-Language" content="es-mx">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>eco2biz : Carga Masiva de Monitoreo</title>
<link rel="stylesheet" type="text/css" href="Estilos/IntraStyles.css">
<link href="estilo/estilo.css" rel="stylesheet" type="text/css" />

</head>

<body bgcolor="#ffffff">
<!--
Modificado por Jorge Felix
Comentario: 'Se agrego replace para evitar comillas simples en comentario
Fecha: 13/10/2021
-->

    <form Name="frmCargos" Action="Carga_Masiva_Lista_Objetos.asp" Method="POST"  ENCTYPE="multipart/form-data">
    <input type="Hidden" name="Code" value="<%=Code%>">
    <input type="Hidden" name="Acc" value="<%=Acc%>">
    <input type="Hidden" name="Id_Unidad" value="<%=Id_Unidad%>">
    <input type="Hidden" name="Id_Usuario" value="<%=wId_Usuario%>">
    <input type="Hidden" name="Empresa" value="<%=wEmpresa%>">
    <input type="Hidden" name="wGrupo" value="<%=wGrupo%>">
    <input type="Hidden" name="wcodigo_ticket" value="<%=wcodigo_ticket%>">
  
    <br>

 <table width=800 align=center>
    <tr>
    <td>
    <table border="0" cellspacing="0" cellpadding="0" width="600" bgcolor="#ffffff" class=dos>
					<tr>
				      <th colspan=2 height=30 align="Left" valign=middle class=Header>&nbsp;<b>CARGA MASIVA : OBJETOS MODIFICADOS</b></th>
					</tr>	
			    <tbody>   

			    <tr>
				      <td width="200" height=22 align="right" class=row1><b>Archivo Excel: </b></td>
				      <td class=row1 align=left>
				      <input TYPE="file" name="ArchivoFoto" SIZE="44%" class=txtcombo>
                      </td>
			    </tr>			   
	
			    <tr>
				    <td align=center Colspan="4" class=row1 height=30><input type="button" value="Cargar Archivo" id=button1 name=button1 onclick="javascript:EjecutaAccion('Load');"></td>
			    </tr>
                <tr>
				      <td colspan=2  height=22 align="center" class=row1>
                        <% 'If wRsCargas("Titulo_CargaMasiva") = "Punto_Monitoreo" Then %>
                            Plantilla de Ejemplo de Carga Masiva Temporal<a href="ayuda/plantilla_carga_masiva_objetos.xls" target="_blank" ><b> Descargar</b></a>
                        <% 'End If %>
				      </td>
			    </tr>
				<tbody>
	</table>
    </td>
    </tr>
</table>
</body>
</form>

</html>

<script LANGUAGE="JavaScript">

	function EjecutaAccion(Acc)
	{
        if (Acc=="Load")
        {
		    frmCargos.Acc.value=Acc; 
		    frmCargos.submit();
		 }
		 else{
            ajaxindicatorstop();
		    alert("No se puede grabar el archivo por que tiene errores. Corrijalos y vuelva a cargarlo.");
		 }   
	}

</script>

<%

Sub CargaTablaTemporal()
    NumCampos = 13

    'Recupero datos de usuario
    Set wRsUsuario = Server.CreateObject("ADODB.Recordset")
    wSQL = "select name, email from sc_user where is_deleted = 0 and sc_user_id = " & wId_Usuario
    wRsUsuario.Open wSQL, oConn
    
    'Codigo Autogenerado
    Set wRsCodigo = Server.CreateObject("ADODB.Recordset")
    wSQL = "select count(*) as contador, max(codigo) as codigo from pa_carga_lista_objetos where is_deleted = 0"
    wRsCodigo.Open wSQL, oConn   

    If  wRsCodigo("contador") > 0 Then
        wCodigo = cint(Right(wRsCodigo("codigo"),4))
        wCodigo = wCodigo + 1
        wCodigoAudit = "CM-" +  Right("0000" + cstr(wCodigo),4)
    Else 
        wCodigoAudit = "CM-0001"
    End if
    
    filename = StrFile
    colDato = ""

    wOperador = wRsUsuario("name")
    wEmail = wRsUsuario("email")
    wFecha = cstr(day(Now)) +"/"+ cstr(Month(Now)) +"/"+ cstr(year(Now))
    wHora = mid(Time(),1,5)
              
    if ExtFile = "XLS" then
        ' Cargo la tabla cabecera
        strSQL = "set dateformat dmy insert into pa_carga_lista_objetos(codigo,fecha,hora,operador,correo_operador,nombre_archivo,ruta_archivo,estado,fb_uea_pe_id,created,created_by, is_deleted)values('" & wCodigoAudit & "','" & wFecha & "','" & wHora & "','" &  wOperador & "','" & wEmail & "','" & NameFile & "','" & StrFile & "',1," & Id_Unidad & ",'" & wFecha & "'," & wId_Usuario & ",0)" 
        oConn.Execute strSQL

        ' Obtengo Id del ultimo Registro
        Set wRsUltimo = Server.CreateObject("ADODB.Recordset")
        wSQL = "select max(pa_carga_lista_objetos_id) as ultimo_id from pa_carga_lista_objetos where is_deleted = 0"
        wRsUltimo.Open wSQL, oConn
        wIdUltimo = wRsUltimo("ultimo_id")

        wSQL1 = "set dateformat dmy insert into pa_carga_lista_objetos_detalle(categoria_objeto,tipo_objeto,nombre_objeto,objeto_relacionado,evento_relacionado,cliente,codigo_jira_incidente,codigo_jira_pase,codigo_is_incidente,programador,tipo_cambio,comentario,comentario_cabecera,fb_uea_pe_id,estado,created,created_by,is_deleted, pa_carga_lista_objetos_id)"
        'Lee cada registro del excel para insertarlo en la tabla temporal
        NumRegistro = 0
        
        objRS.MoveFirst
        Do While Not objRS.Eof
	        ' Arma la cadena con los valores que se van a grabar en la tabla temporal
	        wSQL2 = ""
            wSQL = ""

            IF not isNull(objRs(0)) Then
                wNumeroRegistro = objRs(0)
            Else 
                 wNumeroRegistro = ""
            End If

            If IsNull(wNumeroRegistro) or wNumeroRegistro = "" Then 
                CargaExito = "1"
                exit do 
            End If

	            For i = 1 to NumCampos 'Cargo cada columna en la variable
                        If Not isNull(objRs(i-1))  Then
                            If i = 13 Then 'Se agrego replace para evitar comillas simples en comentario
                                colDato =  "'" + replace(cstr(objRS(i-1)),"'","''")+ "'" 
                            Else
                                colDato =  "'" + cstr(objRS(i-1)) + "'" 
                            End if
                        Else
                                colDato = "Null"
     
                        End If

                        If i = 13 Then
                            wSQL2 = wSQL2 + colDato 
                        Else
                            wSQL2 = wSQL2 + colDato + ","
                        End If
	            Next  
                    wSQL = wSQL1 + "values(" + wSQL2 + "," + cstr(Id_Unidad) + ",1,'" + cstr(wFecha) + "'," + wId_Usuario + ",0," + CStr(wIdUltimo) + ")"
response.write wSQL 
'response.end 
	            oConn.Execute wSQL
                objRS.MoveNext 

                CargaExito = "1"
            'End If
        Loop
        objRS.close  
        set objRS = nothing
        cnADODBConnection.close
        set cnADODBConnection = nothing
    End if
    If CargaExito="0" Then
        strSQL = "set dateformat dmy Delete from pa_carga_lista_objetos where pa_carga_lista_objetos_id =" & wIdUltimo
        oConn.Execute strSQL

        strSQL1 = "set dateformat dmy Delete from pa_carga_lista_objetos_detalle where pa_carga_lista_objetos_id =" & wIdUltimo
        oConn.Execute strSQL1
    End If
    If CargaExito = "1" Then 
        ' Recupero si hay error en data para cargar
        'wValidaError = 1 ' momentaneo
        Set wRsProcesaCarga = Server.CreateObject("ADODB.RecordSet") 

        strSQL = "pr_pa_Procesa_Carga_Masiva_objetos " &  wIdUltimo & "," & wId_Usuario & "," & Id_Unidad

        wRsProcesaCarga.Open strSQL, oConn

        wValidaError = wRsProcesaCarga("valida_error")

            If wValidaError = 1 Then
                CargaExito = "0"
                wMensajeError = "Errores encontrados se han registrado en tabla revision"
                Response.Write wMensajeError
                wError = "1"
            Else
                CargaExito = "1"
                filename = StrFile 

                Set fso = Server.CreateObject("Scripting.FileSystemObject")
                if (fso.FileExists(filename)) then
                    fso.DeleteFile filename,true
                    'Response.Write "<font size=2 color=blue>Borrado el fichero " & filename & " </font>"
                    Response.Write "<font size=2 color=black>Se ha registrado con exito </font>"
                else
                    Response.Write "<font size=2 color=blue>No existe el fichero " & filename & " </font>"
                end if

            End If
        wRsProcesaCarga.close
    End If
End Sub

function esValidoFecha(cadena) 
      set expReg = New RegExp
      expReg.Pattern = "^(0?[1-9]|[12][0-9]|3[01])[\/](0?[1-9]|1[012])[/\\/](19|20)\d{2}$"
      esValidoFecha = expReg.Test(cadena) and len(cadena) = 10
      set expReg = nothing
end function
function esValidoHora(cadena) 
      set expReg = New RegExp
      expReg.Pattern = "^([01]?[0-9]|2[0-3]):[0-5][0-9](:[0-5][0-9])?$"
      'expReg.Pattern = "^([01]?[0-9]|2[0-3]):[0-5][0-9](:[0-5][0-9])\s?(?:[aApP](\.?)[mM]\1)?$"
      esValidoHora = expReg.Test(cadena) and len(cadena) = 8
      set expReg = nothing
end function
function esValidoNumero(cadena) 
      set expReg = New RegExp
      expReg.Pattern = "^[1-9]\d*(\.\d+)?$"
      esValidoNumero = expReg.Test(cadena) and len(cadena) > 0
      set expReg = nothing
end function

%>
<script type="text/javascript">

    if ("<%=wError%>" == "2") { 
        alert("Formato fecha incorrecto. Por favor cargue un archivo el formato correcto");
    }
     if ("<%=wError%>" == "1") { 
        alert("No se pudo cargar tabla residuos ");
    } 
    if ("<%=CargaExito%>" == "1") {
        alert("CARGA DE ARCHIVO EXITOSA");
    }
    if ("<%=CargaExito%>" == "0") {
        var msg_error = "NO SE PUDO CARGAR DATOS: " + "<%=wMensajeError%>";
        alert(msg_error);
    }
    
</script>