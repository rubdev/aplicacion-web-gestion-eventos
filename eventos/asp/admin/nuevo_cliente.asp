<% 
	@ CODEPAGE = 65001 
%>
<!DOCTYPE html>
<html lang="es">
<head>
	<meta charset="UTF-8">
	<title>Añadir nuevo cliente</title>
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
		<h1>Añadir un nuevo cliente</h1><br>
		<a href='../administrador.asp'><b>Volver al panel de administración</b></a> | <a href='../cerrar_sesion.asp'><b>Cerrar sesión</b></a>
	</div>
	
	<!-- Formulario para la inserción de clientes -->
	<div class="alinea-centro">
		<p><b>Rellena los siguientes campos para añadir un nuevo cliente</b></p>
		<form action="#" method="post">
			<table>
				<tr>
					<td><b>Nombre</b></td>
					<td><input type="text" name="nombre"></td>
				</tr>
				<tr>
					<td><b>Teléfono</b></td>
					<td><input type="text" name="telefono"></td>
				</tr>
				<tr>
					<td><b>Dirección</b></td>
					<td><input type="text" name="direccion"></td>
				</tr>
				<tr>
					<td><b>Contraseña</b></td>
					<td><input type="text" name="contra"></td>
				</tr>
				<tr>
					<td><a href='../administrador.asp'><b>Cancelar</b></a></td>
					<td><input type="submit" name="enviar" value="Añadir cliente"></td>
				</tr>
			</table>
		</form>
	</div>
	<%
		'Se recogen los datos del formulario y se guardar en la BD
		l_env = len(request.form("enviar"))

		if l_env > 0 then
			nombre = request.form("nombre")
			telefono = request.form("telefono")
			direccion = request.form("direccion")
			contra = request.form("contra")

			set conexion = server.createobject("ADODB.Connection")
			conexion.open("bd")
			insertar = "INSERT INTO cliente (Nombre,Telefono,Direccion,Contra) VALUES ('"&nombre&"','"&telefono&"','"&direccion&"','"&contra&"')"
			conexion.execute insertar,n

			if n = 1 then
				response.write("<h2 class='correcto' align='center>Cliente añadido correctamente</h2>")
			else
				response.write("<h2 class='error' align='center>Ha ocurrido un error al añadir el nuevo cliente</h2>")
			end if
			
			conexion.close
		end if
	%>
</body>
</html>