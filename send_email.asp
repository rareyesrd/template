<%@language=vbscript%>
<html>
<head><title>Azul Travel SRL | Agencias de Viajes | Santo Domingo | RD</title><head>
<body bgcolor="white" topmargin="0" leftmargin="0">
<H3 align="center"><font face="Verdana">Servicio de Correo Electrónico para <%=Request.Form("compania")%></font></H3>
<hr noshade color="#000080" width="744"><font face="Verdana">
<%
	''<!-- #include file="inicio.htm" -->
	'''' *****----- VARIABLES PARA EL ENVIO DE MAIL -----*****
	Dim emisor, receptor, smtpserver

	emisor = "info@azultravel.net; azultravelrd@gmail.com;"

	receptor = "info@azultravel.net; azultravelrd@gmail.com;"

	smtpserver = "10.20.0.6"

	'''' INICIA EL CUERPO DEL MENSAJE A ENVIAR
	mBody = "<html><body>"
	'''' cerifica a ver si lleva imagen
	if Request("logoImage") <> "" then
		mBody = mBody & "<img src=" & chr(34) & Request("logoImage") & chr(34) & ">" & vbCrlf
	end if
	mBody = mBody & "<font face=" & chr(34) & "Verdana" & chr(34) & _
		     " size=" & chr(34) & "4" & chr(34) & _
		     " width=" & chr(34) & "744" & chr(34) & _
		     " height=" & chr(34) & "72" &  chr(34) & _
		     " >Notificacion vía email </font></b></p><p><b><font face=" & chr(34)& "Verdana"& chr(34) & _
		     ">" & Request("titulo") & "</font></b></p><hr noshade color=" & chr(34) & "#93AAA2"& chr(34) & _
		     " size=" & chr(34) & "3" & chr(34) & "><font face="& chr(34)& "Verdana" & chr(34) & "><p>"

	''' NAVEGA SOBRE TODOS LOS DATOS DEL FORMULARIO CAPRURADO EN EL CLIENTE
	for each item in request.form
		if inStr("Titulo,mailDestino,mailNumero,From,Enviar,Enviar2", item ) <=0  then
		 	mBody = mBody & "-     <b>"& item &" :</b> "& request(item) & "<p>"
		end if
	next
	mBody = mBody & "</font></p></body></html>"
    ''' DECLARA EL CANAL DEL MENSAJE
	Dim objNewMail
	' PREPARA EL CANAL DEL MENSAJE
	Set objNewMail = Server.CreateObject("CDO.Message")

	''' VALIDA SI LA CONXION SE PUDO CREAR
	if isobject(objNewMail) then
	 	response.write "."
	else
		response.write ".."
	end if
  if isobject(objNewMail) then

	sch = "http://schemas.microsoft.com/cdo/configuration/"
	Set objConfig = CreateObject("CDO.Configuration")
	With objConfig.Fields
		.Item(sch & "sendusing") = 2 ' cdoSendUsingPort
		.Item(sch & "smtpserver") = smtpserver
		.update
	End With



		dim varNum '''' variable que indica el seleccionado si existen varios mail
		dim varstrDestino '''' declaracion del destino
		varNum = 0 ''' Es el valor por default
		if Request.form("mailDestino").count > 1 then
			varNum = Request("mailNumero")
			varstrDestino = Request.form("mailDestino").Item(varNum)
		else
			varstrDestino = Request.form("mailDestino")
		end if
	''' ENCABEZADO PARA TRABAJAR; A PARTIR DE ESTE SE ALIMENTARAN LOS PARAMETROS DEL MAIL
	WITH objNewMail
		Set .Configuration = objConfig
		.From =  emisor
		.To = receptor
		.Subject = "Notificación del formulario de contacto de su website"
		''''.TextBody = mBody
		.HTMLBody = mBody
		''' ENVIA EL MAIL
		call objNewMail.Send
		if err.number > 0 then
			varmensaje = "Actualmente el servicio de envío de correo no esta disponible. Por favor intente mas tarde."
		else
			varmensaje = "Su mensaje sera atendido en la mayor brevedad posible, gracias por preferirnos."
		end if
	end with
  else
		varmensaje = "Actualmente el servicio de envío de correo no esta disponible. Por favor intente mas tarde."
  end if
  Set objNewMail = Nothing
%>
<p align="center">Mensaje del Sistema: </p>
<p align="center"> <b> <%=varmensaje%></b>
</p></font>
<p>&nbsp;</p>
<p>&nbsp;</p>
<DIV><%=mBody%></DIV>
</html>