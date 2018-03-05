<% @ CODEPAGE = 65001 %>
<!DOCTYPE html>
<html lang="es">
<head>
	<meta charset="UTF-8">
	<title>Buscar actividades</title>
	<link rel="stylesheet" href="../css/estilos.css">
</head>
<body>
	<%
		'Compruebo sesion y cookies para restringir usuarios y muestro encabezados de la web
		l_sesion = len(Session("cod"))
		l_recordar = len(request.cookies("recordar"))
			if l_sesion > 0 then
				if Session("cod") = 1 then
					response.write("<div class='header-admin'>")
					response.write("<h1>Buscar actividades</h1><br>")
					response.write("<a href='administrador.asp'><b>Volver al panel de administración</b></a> | <a href='cerrar_sesion.asp'><b>Cerrar sesión</b></a>")
					response.write("</div>")
				elseif Session("cod") > 1 then
					response.write("<div class='header-admin'>")
					response.write("<h1>Buscar actividades</h1><br>")
					response.write("<a href='cliente.asp'><b>Volver al panel de administración</b></a> | <a href='cerrar_sesion.asp'><b>Cerrar sesión</b></a>")
					response.write("</div>")
				else
					response.redirect("../index.asp")
				end if
			elseif l_recordar > 0 then
				if request.cookies("recordar") = 1 then
					response.write("<div class='header-admin'>")
					response.write("<h1>Buscar actividades</h1><br>")
					response.write("<a href='administrador.asp'><b>Volver al panel de administración</b></a> | <a href='cerrar_sesion.asp'><b>Cerrar sesión</b></a>")
					response.write("</div>")
				elseif request.cookies("recordar") > 1 then
					response.write("<div class='header-admin'>")
					response.write("<h1>Buscar actividades</h1><br>")
					response.write("<a href='cliente.asp'><b>Volver al panel de administración</b></a> | <a href='cerrar_sesion.asp'><b>Cerrar sesión</b></a>")
					response.write("</div>")
				else 
					response.redirect("../index.asp")
				end if
			else
				response.redirect("../index.asp")
			end if 

	%>
	<!-- Formulario para la búsqueda de actividades -->
	<div class="alinea-centro">
		<p><b>Introduzca el nombre de la actividad que desea buscar</b></p>
		<form action="#" method="post">
			<table>
				<tr>
					<td><input type="text" name="actividad"></td>
					<td><input type="submit" name="buscar" value="Buscar actividad"></td>
				</tr>
			</table>		
		</form>
		<%
			l_buscar = len(request.form("buscar"))

			if l_buscar > 0 then
				actividad = request.form("actividad")
				actividad = trim(actividad)
				response.write("<br><p>El criterio de búsqueda introducido ha sido <b>"&actividad&"</b></p>")
				
					set conexion = server.createobject("ADODB.Connection")
					conexion.open("bd")
					busqueda = "SELECT e.codigo, e.actividad, e.fecha_contrato, e.fecha_evento, a.Codigo, a.Nombre as nom_acti FROM eventos e, actividad a WHERE e.actividad=a.Codigo AND a.Nombre LIKE '%"&actividad&"%'"
					set datos = conexion.execute(busqueda)

					if datos.eof then
						response.write("<h3 class='error' align='center'>No se ha encontrado ningún evento con la actividad indicada</h3>")
					else
						response.write("<table border align='center'>")
						response.write("<tr bgcolor=#7fb3d5><td><b>Nombre de la actividad</b></td><td><b>Fecha de contrato</b></td><td><b>Fecha del evento</b></td></tr>")
						do while not datos.EoF
							response.write("<tr bgcolor=#a9cce3><td>"&datos("nom_acti")&"</td><td>"&datos("fecha_contrato")&"</td><td>"&datos("fecha_evento")&"</td></tr>")
							datos.movenext
						loop
					end if
					conexion.close
			end if
		%>
	</div>
</body>
</html>