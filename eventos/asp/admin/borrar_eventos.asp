<% 
	@ CODEPAGE = 65001 
%>
<!DOCTYPE html>
<html lang="es">
<head>
	<meta charset="UTF-8">
	<title>Borrar eventos</title>
	<link rel="stylesheet" href="../../css/estilos.css">
</head>
<body>
	<%
		l_sesion = len(Session("cod"))
		if l_sesion > 0 then
			codigo = Session("cod")
			if codigo <> 1 then
				response.redirect("../../index.asp")
			end if
		else
			recordar = request.cookies("recordar")
				if recordar <> "1" then
					response.redirect("../../index.asp")
				end if 
		end if
	%>
	<div class="header-admin">
		<h1>Borrar eventos</h1><br>
		<a href='../administrador.asp'><b>Volver al panel de administración</b></a> | <a href='../cerrar_sesion.asp'><b>Cerrar sesión</b></a>
	</div>
	<div class="alinea-centro">
		<p><b>Pulse en el botón 'borrar' asociado al evento</b></p>
		<%
			'Listo los eventos de la BD y le añado un boton a cada una para eliminarlas
			f_actual = date()

			set conexion = server.createobject("ADODB.Connection")
			conexion.open("bd")
			consulta = "SELECT e.codigo as cod, e.actividad, e.cliente, e.fecha_contrato, e.fecha_evento, a.Codigo, a.Nombre as nom_acti, c.Codigo, c.Nombre as nom_cli FROM eventos e, actividad a, cliente c WHERE e.actividad = a.Codigo AND e.cliente = c.Codigo AND e.fecha_evento > "&f_actual&" ORDER BY e.fecha_evento ASC"
			set datos = conexion.execute(consulta)

			'IMPORTANTE ==> if datos.eof then ==> si quieres comprobar antes que te ha devuelto datos la consulta

			response.write("<table border align='center'>")
			response.write("<tr bgcolor=#7fb3d5><td><b>Actividad</b></td><td><b>Cliente</b></td><td><b>Fecha de contrato</b></td><td><b>Fecha del evento</b></td><td><b>Borrar evento</b></td></tr>")
			do while not datos.EoF 
				'Compruebo que la fecha del evento no ha pasado ya
				if f_actual < datos("fecha_evento") then
					response.write("<tr bgcolor=#a9cce3><td>"&datos("nom_acti")&"</td><td>"&datos("nom_cli")&"</td><td>"&datos("fecha_contrato")&"</td><td>"&datos("fecha_evento")&"</td><td><form action='#' method='post'><input type='submit' name='borrar' value='Borrar'><input type='hidden' name='cod' value='"&datos("cod")&"'></form></td></tr>")
				end if
				datos.movenext
			loop
			response.write("</table>")
			conexion.close

			'Si pulsa en un botón borrar se ejecuta el siguiente código para borrar el evento de la BD
			l_borrar = len(request.form("borrar"))

			if l_borrar > 0 then
				codigo = cInt(request.form("cod"))
				conexion.open("bd")
				quita_codigo_actividad =""
				eliminar = "DELETE FROM eventos WHERE codigo="&codigo&""

				'Primero quito la asociación que tenga una actividad con el evento y después borro el evento
				conexion.execute eliminar, n

				if n = 1 then
					response.write("<h2 class='correcto' align='center>Evento borrado correctamente</h2>")
					response.addheader "REFRESH","2;URL=borrar_eventos.asp"
				else
					response.write("<h2 class='error' align='center>Se ha producido un error al eliminar el evento</h2>")
				end if

				conexion.close
			end if

		%>
</body>
</html>