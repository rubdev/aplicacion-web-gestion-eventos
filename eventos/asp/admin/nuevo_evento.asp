<% 
	@ CODEPAGE = 65001 
%>
<!DOCTYPE html>
<html lang="es">
<head>
	<meta charset="UTF-8">
	<title>Añadir nuevo evento</title>
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
		<h1>Añadir un nuevo evento</h1><br>
		<a href='../administrador.asp'><b>Volver al panel de administración</b></a> | <a href='../cerrar_sesion.asp'><b>Cerrar sesión</b></a>
	</div>
	
	<!-- Formulario para la inserción de eventos
		 actividad(id de activida),cliente,fecha_contrato,fecha_evento -->
	<div class="alinea-centro">
		<p><b>Rellena los siguientes campos para añadir un nuevo evento</b></p>
		<form action="#" method="post">
			<table>
				<tr>
					<td><b>Actividad</b></td>
					<td>
						<select name="actividad">
							<%
								'Cargo el nombre de todas las actividades y me quedo
								'con su id para después guardar todos los datos en la BD
								set conexion = server.createobject("ADODB.Connection")
								conexion.open("bd")
								actividades = "SELECT Codigo,Nombre FROM actividad ORDER BY Nombre ASC"
								set datos = conexion.execute(actividades)

								do while not datos.EoF
									response.write("<option value='"&datos("Codigo")&"'>"&datos("Nombre")&"</option>")
									datos.movenext
								loop

								conexion.close
							%>
						</select>
					</td>
				</tr>
				<tr>
					<td><b>Cliente</b></td>
					<td>
						<select name="cliente">
							<%
								'Cargo el nombre de todos los clientes y me quedo
								'con su id para después guardar todos los datos en la BD
								set conexion = server.createobject("ADODB.Connection")
								conexion.open("bd")
								actividades = "SELECT Codigo,Nombre FROM cliente ORDER BY Nombre ASC"
								set datos = conexion.execute(actividades)

								do while not datos.EoF
									if datos("Codigo") <> "1" then
										response.write("<option value='"&datos("Codigo")&"'>"&datos("Nombre")&"</option>")
									end if
									datos.movenext
								loop

								conexion.close
							%>
						</select>
					</td>
				</tr>
				<tr>
					<td><b>Fecha de evento</b></td>
					<td><input type="date" name="f_evento"></td>
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
			actividad = request.form("actividad")
			cliente = request.form("cliente")
			f_evento = request.form("f_evento")
			'Formateo la fecha de eventp introducida y compruebo que sean válida
			f_contrato = date()
			formateada_ev = Cdate(f_evento)
			if formateada_ev > f_contrato then
				'Si la fecha es correcta se procede a guardar el evento en la BD
				set conexion = server.createobject("ADODB.Connection")
				conexion.open("bd")
				insertar = "INSERT INTO eventos (actividad,cliente,fecha_contrato,fecha_evento) VALUES ('"&actividad&"','"&cliente&"','"&f_contrato&"','"&formateada_ev&"')"
				
				conexion.execute insertar,n

				if n = 1 then
					response.write("<h2 class='correcto' align='center'>Evento añadido correctamente</h2>")
				else
					response.write("<h2 class='error' align='center>Ha ocurrido un error al añadir el nuevo evento</h2>")
				end if
				conexion.close
			else
				response.write("<h3 class='error' align='center'>Error, la fecha debe ser posterior al día de hoy</h3>")
			end if
		end if
	%>
</body>
</html>