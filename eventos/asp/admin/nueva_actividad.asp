<% 
	@ CODEPAGE = 65001 
%>
<!DOCTYPE html>
<html lang="es">
<head>
	<meta charset="UTF-8">
	<title>Añadir nueva actividad</title>
	<link rel="stylesheet" href="../../css/estilos.css">
</head>
<body>
	<%
		l_sesion = len(Session("cod"))
			if l_sesion > 0 then
				'usuario = Session("nombre")
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
		<h1>Añadir una nueva actividad</h1><br>
		<a href='../administrador.asp'><b>Volver al panel de administración</b></a> | <a href='../cerrar_sesion.asp'><b>Cerrar sesión</b></a>
	</div>
	
	<!-- Formulario para la inserción de actividades
		 Nombre,Descripcion,Duracion,Precio -->
	<div class="alinea-centro">
		<p><b>Rellena los siguientes campos para añadir una nueva actividad</b></p>
		<form action="#" method="post">
			<table>
				<tr>
					<td><b>Nombre</b></td>
					<td><input type="text" name="nombre"></td>
				</tr>
				<tr>
					<td><b>Descripción</b></td>
					<td><input type="text" name="descripcion"></td>
				</tr>
				<tr>
					<td><b>Duración (en horas)</b></td>
					<td><input type="number" name="duracion"></td>
				</tr>
				<tr>
					<td><b>Precio</b></td>
					<td><input type="text" name="precio"></td>
				</tr>
				<tr>
					<td><a href='../administrador.asp'><b>Cancelar</b></a></td>
					<td><input type="submit" name="enviar" value="Añadir actividad"></td>
				</tr>
			</table>
		</form>
	</div>
	<%
		'Se recogen los datos del formulario y se guardar en la BD
		l_env = len(request.form("enviar"))

		if l_env > 0 then
			nombre = request.form("nombre")
			descripcion = request.form("descripcion")
			duracion = request.form("duracion")
			precio = request.form("precio")

			set conexion = server.createobject("ADODB.Connection")
			conexion.open("bd")
			insertar = "INSERT INTO actividad (Nombre,Descripcion,Duracion,Precio) VALUES ('"&nombre&"','"&descripcion&"','"&duracion&"','"&precio&"')"
			
			conexion.execute insertar,n

			if n = 1 then
				response.write("<h2 class='correcto' align='center'>Actividad añadida correctamente</h2>")
			else
				response.write("<h2 class='error' align='center>Ha ocurrido un error al añadir la nueva actividad</h2>")
			end if
			conexion.close
		end if
	%>
</body>
</html>