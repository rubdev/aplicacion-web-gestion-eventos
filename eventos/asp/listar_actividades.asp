<% 
	@ CODEPAGE = 65001 
%>
<!DOCTYPE html>
<html lang="es">
<head>
	<meta charset="UTF-8">
	<title>Listado actividades</title>
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
					response.write("<h1>Listado de actividades</h1><br>")
					response.write("<a href='administrador.asp'><b>Volver al panel de administración</b></a> | <a href='cerrar_sesion.asp'><b>Cerrar sesión</b></a>")
					response.write("</div>")
				elseif Session("cod") > 1 then
					response.write("<div class='header-admin'>")
					response.write("<h1>Listado de actividades</h1><br>")
					response.write("<a href='cliente.asp'><b>Volver al panel de administración</b></a> | <a href='cerrar_sesion.asp'><b>Cerrar sesión</b></a>")
					response.write("</div>")
				else
					response.redirect("../index.asp")
				end if
			elseif l_recordar > 0 then
				if request.cookies("recordar") = 1 then
					response.write("<div class='header-admin'>")
					response.write("<h1>Listado de actividades</h1><br>")
					response.write("<a href='administrador.asp'><b>Volver al panel de administración</b></a> | <a href='cerrar_sesion.asp'><b>Cerrar sesión</b></a>")
					response.write("</div>")
				elseif request.cookies("recordar") > 1 then
					response.write("<div class='header-admin'>")
					response.write("<h1>Listado de actividades</h1><br>")
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
		<%
			'Se accebe a BD y muestran las actividades almacenados
			set conexion = server.createobject("ADODB.Connection")
			conexion.open("bd")
			consulta = "SELECT * FROM actividad ORDER BY Nombre ASC"
			set datos = conexion.execute(consulta)

			response.write("<table border align='center'>")
			response.write("<tr bgcolor=#7fb3d5><td><b>Código</b></td><td><b>Nombre</b></td><td><b>Descripción</b></td><td><b>Duración</b></td><td><b>Precio</b></td></tr>")
			do while not datos.EoF 
				response.write("<tr bgcolor=#a9cce3><td>"&datos("Codigo")&"</td><td>"&datos("Nombre")&"</td><td>"&datos("Descripcion")&"</td><td>"&datos("Duracion")&"</td><td>"&datos("Precio")&"</td></tr>")
				datos.movenext
			loop
			response.write("</table>")
			conexion.close
		%>
	</div>
</body>
</html>