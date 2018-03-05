<% 
	@ CODEPAGE = 65001 
%>
<!DOCTYPE html>
<html lang="es">
<head>
	<meta charset="UTF-8">
	<title>Panel de administración de cliente</title>
	<link rel="stylesheet" href="../css/estilos.css">
</head>
<body>
	<div class="header-admin">
		<%
			'compruebo si hay una sesión correcta
			'o alguna cookie válida guardada y muestro el mensaje de bienvenida al cliente
			l_sesion = len(Session("cod"))
			if l_sesion > 0 then
				codigo = Session("cod")
				if codigo < 2 then
					response.redirect("../index.asp")
				end if
				response.write("<h1>Panel de administración de cliente</h1></br>")
				response.write("<b>Bienvenid@ "&Session("Nombre")&"</b> | <a href='cerrar_sesion.asp'><b>Cerrar sesión</b></a>")
			else
				recordar = request.cookies("recordar")
				nombre = request.cookies("nombre")
					if recordar < 2 then
						response.redirect("../index.asp")
					end if 
				response.write("<h1>Panel de administración de cliente</h1></br>")
				response.write("<b>Bienvenid@ "&nombre&"</b> | <a href='cerrar_sesion.asp'><b>Cerrar sesión</b></a>")
			end if
		%>
	</div>
	<div class="acciones-admin" align="center">
		<div class="accion-admin">
			<p><b>Clientes</b></p>
			<table>
				<tr>
					<td><a href="client/ver_factura.asp"><b>Ver factura</b></a></td>
				</tr>
			</table>
		</div>
		<div class="accion-admin">
			<p><b>Actividades</b></p>
			<table>
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