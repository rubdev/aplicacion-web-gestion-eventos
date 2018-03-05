<% 
	@ CODEPAGE = 65001 
%>
<!DOCTYPE html>
<html lang="es">
<head>
	<meta charset="UTF-8">
	<title>Borrar actividades</title>
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
	<!-- cabecera -->
	<div class="header-admin">
		<h1>Borrar actividades</h1><br>
		<a href='../administrador.asp'><b>Volver al panel de administración</b></a> | <a href='../cerrar_sesion.asp'><b>Cerrar sesión</b></a>
	</div>
	
	<!-- Formulario para la inserción de clientes -->
	<div class="alinea-centro">
		<p><b>Pulse en el botón 'borrar' asociado a la actividad</b></p>
		<%
			'Listo las actividades de la BD y le añado un boton a cada una para eliminarlas
			set conexion = server.createobject("ADODB.Connection")
			conexion.open("bd")
			consulta = "SELECT * FROM actividad ORDER BY Nombre ASC"
			set datos = conexion.execute(consulta)

			response.write("<table border align='center'>")
			response.write("<tr bgcolor=#7fb3d5><td><b>Código</b></td><td><b>Nombre</b></td><td><b>Descripción</b></td><td><b>Duración</b></td><td><b>Precio</b></td><td><b>Borrar actividad</b></td></tr>")
			do while not datos.EoF 
				response.write("<tr bgcolor=#a9cce3><td>"&datos("Codigo")&"</td><td>"&datos("Nombre")&"</td><td>"&datos("Descripcion")&"</td><td>"&datos("Duracion")&"</td><td>"&datos("Precio")&"</td><td><form action='#' method='post'><input type='submit' name='borrar' value='Borrar'><input type='hidden' name='cod' value='"&datos("codigo")&"'></form></td></tr>")
				datos.movenext
			loop
			response.write("</table>")
			conexion.close

			'Si pulsa en un botón borrar se ejecuta el siguiente código para borrar los eventos planificados con la actividad de la BD si es que los tiene

			l_borrar = len(request.form("borrar"))

			if l_borrar > 0 then

				codigo = Cint(request.form("cod"))

				conexion.open("bd")
				borrar_ev_asociados = "DELETE FROM eventos WHERE actividad="&codigo&""
				response.write(borrar_ev_asociados)
				conexion.execute borrar_ev_asociados,n

				if n >= 1 then
					response.write("<h2 class='correcto' align='center>Los eventos planificados con la actividad se han eliminado de la BD</h2>")

					'Si se borran los asociados entonces borro la actividad de la BD
					borrar = "DELETE FROM actividad WHERE Codigo="&codigo&""
					conexion.execute borrar,num

					if num = 1 then
						response.write("<h2 class='correcto' align='center>Y la actividad también se ha correctamente</h2>")
						response.addheader "REFRESH","2;URL=borrar_actividades.asp"
					else
						response.write("<h2 class='error' align='center>Pero se ha producido un error al eliminar la actividad</h2>")
					end if
				else
					
					borrar = "DELETE FROM actividad WHERE Codigo="&codigo&""
					conexion.execute borrar,num

					if num = 1 then
						response.write("<h2 class='correcto' align='center>Y la actividad también se ha correctamente</h2>")
						response.addheader "REFRESH","2;URL=borrar_actividades.asp"
					else
						response.write("<h2 class='error' align='center>Pero se ha producido un error al eliminar la actividad</h2>")
					end if
				end if
				conexion.close
			end if
		%>
	</div>
</body>
</html>