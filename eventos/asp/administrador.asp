<% 
	@ CODEPAGE = 65001 
%>
<!DOCTYPE html>
<html lang="es">
<head>
	<meta charset="UTF-8">
	<title>Panel de administración de Administrador</title>
	<link rel="stylesheet" href="../css/estilos.css">
</head>
<body>
	<div class="header-admin">
		<%
			'compruebo si hay una sesión correcta
			'o alguna cookie válida guardada
			l_sesion = len(Session("cod"))
			if l_sesion > 0 then
				codigo = Session("cod")
				if codigo <> 1 then
					response.redirect("../index.asp")
				end if
			else
				recordar = request.cookies("recordar")
					if recordar <> "1" then
						response.redirect("../index.asp")
					end if 
			end if

			'Mensaje de bienvenida al admin
			response.write("<h1>Panel de administración de eventos</h1></br>")
			response.write("<b>Bienvenido Administrador</b> | <a href='cerrar_sesion.asp'><b>Cerrar sesión</b></a>")
		%>
	</div>
	<div class="acciones-admin" align="center">
		<div class="accion-admin">
			<p><b>Clientes</b></p>
			<table>
				<tr>
					<td><a href="admin/nuevo_cliente.asp"><b>Añadir un nuevo cliente</b></a></td>
				</tr>
				<tr>
					<td><a href="admin/listar_clientes.asp"><b>Listado de clientes</b></a></td>
				</tr>
			</table>
		</div>
		<div class="accion-admin">
			<p><b>Actividades</b></p>
			<table>
				<tr>
					<td><a href="admin/nueva_actividad.asp"><b>Añadir una nueva actividad</b></a></td>
				</tr>
				<tr>
					<td><a href="admin/borrar_actividades.asp"><b>Borrar actividades</b></a></td>
				</tr>
				<tr>
					<td><a href="buscar_actividad.asp"><b>Buscar actividad</b></a></td>
				</tr>
				<tr>
					<td><a href="listar_actividades.asp"><b>Listado de actividades</b></a></td>
				</tr>
			</table>
		</div>
		<div class="accion-admin">
			<p><b>Eventos</b></p>
			<table>
				<tr>
					<td><a href="admin/nuevo_evento.asp"><b>Añadir un nuevo evento</b></a></td>
				</tr>
				<tr>
					<td><a href="admin/borrar_eventos.asp"><b>Borrar eventos</b></a></td>
				</tr>
				<tr>
					<td><a href="buscar_eventos.asp"><b>Buscar eventos</b></a></td>
				</tr>
				<tr>
					<td><a href="ver_calendario.asp"><b>Ver calendario de eventos</b></a></td>
				</tr>
			</table>
		</div>
		
	</div>
</body>
</html>