<% 
	@ CODEPAGE = 65001 
%>
<!DOCTYPE html>
<html lang="es">
<head>
	<meta charset="UTF-8">
	<title>Calendario de eventos</title>
	<link rel="stylesheet" href="../css/estilos.css">
</head>
<body>
	<%
		'Compruebo sesion y cookies para restringir usuarios y muestro encabezados de la web
		l_sesion = len(Session("cod"))
		l_recordar = len(request.cookies("recordar"))
			if l_sesion > 0 then
				if Session("cod") = 1 then
					codigo = Session("cod")
					response.write("<div class='header-admin'>")
					response.write("<h1>Calendario de eventos</h1><br>")
					response.write("<a href='administrador.asp'><b>Volver al panel de administración</b></a> | <a href='cerrar_sesion.asp'><b>Cerrar sesión</b></a>")
					response.write("</div>")
				elseif Session("cod") > 1 then
					codigo = Session("cod")
					response.write("<div class='header-admin'>")
					response.write("<h1>Calendario de eventos</h1><br>")
					response.write("<a href='cliente.asp'><b>Volver al panel de administración</b></a> | <a href='cerrar_sesion.asp'><b>Cerrar sesión</b></a>")
					response.write("</div>")
				else
					response.redirect("../index.asp")
				end if
			elseif l_recordar > 0 then
				if request.cookies("recordar") = 1 then
					codigo = request.cookies("recordar")
					response.write("<div class='header-admin'>")
					response.write("<h1>Calendario de eventos</h1><br>")
					response.write("<a href='administrador.asp'><b>Volver al panel de administración</b></a> | <a href='cerrar_sesion.asp'><b>Cerrar sesión</b></a>")
					response.write("</div>")
				elseif request.cookies("recordar") > 1 then
					codigo = request.cookies("recordar")
					response.write("<div class='header-admin'>")
					response.write("<h1>Calendario de eventos</h1><br>")
					response.write("<a href='cliente.asp'><b>Volver al panel de administración</b></a> | <a href='cerrar_sesion.asp'><b>Cerrar sesión</b></a>")
					response.write("</div>")
				else 
					response.redirect("../index.asp")
				end if
			else
				response.redirect("../index.asp")
			end if 
	%>
	<div class="alinea-centro">
		<p>Los <b>días con evento</b> aparecerán marcados en <span id="marcado"> rojo</span></p>
		<p><a href="ver_calendario.asp">Volver al día actual</a></p>
		<%
			'Se pone la fecha actual o la del mes que elija el usuario al cambiar con los enlaces
			if Request("datos") = "" then
			  datos = Date()
			else
			  datos = CDate(Request("datos"))
			end if

			mes = Month(datos)
			anio = Year(datos)

			'Con esta función saco el número de días que tiene un mes
			function dias_mes(mes, anio)
			  f_inicio = CDate("01/" & mes & "/" & anio)
			  f_fin = DateAdd("m", 1, f_inicio)
			  dias_mes = DateDiff("d", f_inicio, f_fin)
			end function

			'Función que busca en BD si el día actual (i) al hacer el calendario hay un evento
			'si es admin los muestra todos y si es un usuario le muestra sólo sus eventos 

			function dia_evento(i,mes,anio)
				f_formateada = cdate(i & "/" & mes & "/" & anio)

				if codigo = 1 then
					set conexion = server.createobject("ADODB.Connection")
					conexion.open("bd")
					con_evento = "SELECT fecha_evento FROM eventos WHERE fecha_evento LIKE '%"&f_formateada&"%'"
					set fechas = conexion.execute(con_evento)

					if fechas.EoF then
						dia_evento = Response.Write("<td bgcolor=#a9cce3><b>" & i & "</b></td>")
					else
						dia_evento = Response.Write("<td bgcolor=#ea0e00><b>" & i & "</b></td>")
					end if
					conexion.close
				else
					set conexion = server.createobject("ADODB.Connection")
					conexion.open("bd")
					con_evento = "SELECT fecha_evento FROM eventos WHERE fecha_evento LIKE '%"&f_formateada&"%' AND cliente="&codigo&""
					set fechas = conexion.execute(con_evento)

					if fechas.EoF then
						dia_evento = Response.Write("<td bgcolor=#a9cce3><b>" & i & "</b></td>")
					else
						dia_evento = Response.Write("<td bgcolor=#ea0e00><b>" & i & "</b></td>")
					end if
					conexion.close
				end if

			end function
			%>
			<!-- Cabecera del calendario -->
			<table class="calendario" align="center" border='solid 2px'>
			<tr bgcolor=#7fb3d5>
				<td></td>
				<td></td>
				<td><a href="ver_calendario.asp?datos=<%=DateAdd("m", -1, datos)%>"><b>Mes anterior</b></td>
				<td><b><%=MonthName(Month(datos)) & " " & Year(datos)%></b></td>
				<td><a href="ver_calendario.asp?datos=<%=DateAdd("m", 1, datos)%>"><b>Mes siguiente</b></td>
				<td></td>
				<td></td>
			</tr>
			<tr bgcolor=#7fb3d5>
				<td><b>Domingo</b></td><td><b>Lunes</b></td><td><b>Martes</b></td><td><b>Miércoles</b></td><td><b>Jueves</b></td><td><b>Viernes</b></td><td><b>Sábado</b></td>
			</tr>
			<%
			f_inicio = CDate("01/" & mes & "/" & anio)

			'Salto los primeros días del mes hasta el día que les corresponde
			for i = 1 to WeekDay(f_inicio)-1
			  if i = 1 then Response.Write "<tr>"
			  Response.write "<td> </td>"
			next

			'Aquí muestro el calendario
			for i = 1 to dias_mes(mes,anio)
			  datos = Cdate(( i & "/" & mes & "/" & anio))
			  if WeekDay(datos) = 1 then Response.Write "<tr>"
			  'Llamo a la función que me dice si hay evento un día en concreto 
			  'COMPRUEBO ANTES SI ES ADMIN(MUESTRO TODOS LOS EVENTOS) O UN CLIENTE (SÓLO MUESTRO LOS SUYOS)
			  dia_evento i,mes,anio
			  if WeekDay(datos) = 7 then Response.Write "</tr>"
			next

			'Vuelvo a colocar los días
			for j = WeekDay(Data)+1 to 7
			  Response.write "<td> </td>"
			  if j mod 7 = 0 then Response.Write "</tr>" 
			next
		%>
	</div>
</body>
</html>