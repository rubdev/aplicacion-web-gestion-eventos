<% @ CODEPAGE = 65001 %>
<!DOCTYPE html>
<html lang="es">
<head>
	<meta charset="UTF-8">
	<title>Buscar Eventos</title>
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
					response.write("<h1>Buscar eventos</h1><br>")
					response.write("<a href='administrador.asp'><b>Volver al panel de administración</b></a> | <a href='cerrar_sesion.asp'><b>Cerrar sesión</b></a>")
					response.write("</div>")
				elseif Session("cod") > 1 then
					response.write("<div class='header-admin'>")
					response.write("<h1>Buscar eventos</h1><br>")
					response.write("<a href='cliente.asp'><b>Volver al panel de administración</b></a> | <a href='cerrar_sesion.asp'><b>Cerrar sesión</b></a>")
					response.write("</div>")
				else
					response.redirect("../index.asp")
				end if
			elseif l_recordar > 0 then
				if request.cookies("recordar") = 1 then
					response.write("<div class='header-admin'>")
					response.write("<h1>Buscar eventos</h1><br>")
					response.write("<a href='administrador.asp'><b>Volver al panel de administración</b></a> | <a href='cerrar_sesion.asp'><b>Cerrar sesión</b></a>")
					response.write("</div>")
				elseif request.cookies("recordar") > 1 then
					response.write("<div class='header-admin'>")
					response.write("<h1>Buscar eventos</h1><br>")
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
		<p><b>Rellene los parámetros por los que quiere realizar la búsqueda</b></p>
		<form action="#" method="post">
			<table>
				<tr>
					<td><b>Nombre de actividad</b></td>
					<td>
						<select name="actividad">
							<%
								'Listo las actividades almacenadas en bd
								set conexion = server.createobject("ADODB.Connection")
								conexion.open("bd")
								act = "SELECT Nombre FROM actividad"
								set datos = conexion.execute(act)
									response.write("<option value=''></option>")
								do while not datos.EoF
									response.write("<option value='"&datos("Nombre")&"'>"&datos("Nombre")&"</option>")
									datos.movenext
								loop

								conexion.close
							%>
						</select>
					</td>
				</tr>
				<tr>
					<td><b>Nombre de cliente</b></td>
					<td><input type="text" name="cliente"></td>
				</tr>
				<tr>
					<td><b>Fecha de contrato</b></td>
					<td><input type="date" name="f_contrato"></td>
				</tr>
				<tr>
					<td><b>Fecha de evento</b></td>
					<td><input type="date" name="f_evento"></td>
				</tr>
				<tr>
					<td><input type="submit" name="buscar" value="Buscar eventos"></td>
				</tr>
			</table>		
		</form>
		<%
			l_buscar = len(request.form("buscar"))

			if l_buscar > 0 then
				actividad = request.form("actividad")
				l_actividad = len(actividad)

				if l_actividad > 0 then 
					response.write("Busco por actividad")
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
				else
					cliente = request.form("cliente")
					l_cliente = len(cliente)
					if l_cliente > 0 then
						response.write("Busco por nombre de cliente")
						response.write("<br><p>El criterio de búsqueda introducido ha sido <b>"&cliente&"</b></p>")
					
						set conexion = server.createobject("ADODB.Connection")
						conexion.open("bd")
						busqueda = "SELECT e.codigo, e.actividad, e.cliente, e.fecha_contrato, e.fecha_evento, a.Codigo, a.Nombre as nom_acti, c.Codigo, c.Nombre as nom_cli FROM eventos e, actividad a, cliente c WHERE e.actividad=a.Codigo AND e.cliente=c.Codigo AND c.Nombre LIKE '%"&cliente&"%'"
						set datos = conexion.execute(busqueda)

						if datos.eof then
							response.write("<h3 class='error' align='center'>No se ha encontrado ningún evento con el cliente indicado</h3>")
						else
							response.write("<table border align='center'>")
							response.write("<tr bgcolor=#7fb3d5><td><b>Nombre del cliente</b></td><td><b>Nombre de la actividad</b></td><td><b>Fecha de contrato</b></td><td><b>Fecha del evento</b></td></tr>")
							do while not datos.EoF
								response.write("<tr bgcolor=#a9cce3><td>"&datos("nom_cli")&"</td><td>"&datos("nom_acti")&"</td><td>"&datos("fecha_contrato")&"</td><td>"&datos("fecha_evento")&"</td></tr>")
								datos.movenext
							loop
						end if
						conexion.close
					else
						f_contrato = request.form("f_contrato")
						l_con = len(f_contrato)

						if l_con > 0 then
							f_formateada = cdate(f_contrato)
							response.write("Busco por fecha contrato")
							response.write("<br><p>El criterio de búsqueda introducido ha sido <b>"&f_formateada&"</b></p>")
					
							set conexion = server.createobject("ADODB.Connection")
							conexion.open("bd")
							busqueda = "SELECT e.codigo, e.actividad, e.cliente, e.fecha_contrato, e.fecha_evento, a.Codigo, a.Nombre as nom_acti, c.Codigo, c.Nombre as nom_cli FROM eventos e, actividad a, cliente c WHERE e.actividad=a.Codigo AND e.cliente=c.Codigo AND e.fecha_contrato LIKE '%"&f_formateada&"%'"
							set datos = conexion.execute(busqueda)

							if datos.eof then
								response.write("<h3 class='error' align='center'>No se ha encontrado ningún evento en la fecha de contrato indicada</h3>")
							else
								response.write("<table border align='center'>")
								response.write("<tr bgcolor=#7fb3d5><td><b>Nombre del cliente</b></td><td><b>Nombre de la actividad</b></td><td><b>Fecha de contrato</b></td><td><b>Fecha del evento</b></td></tr>")
								do while not datos.EoF
									response.write("<tr bgcolor=#a9cce3><td>"&datos("nom_cli")&"</td><td>"&datos("nom_acti")&"</td><td>"&datos("fecha_contrato")&"</td><td>"&datos("fecha_evento")&"</td></tr>")
									datos.movenext
								loop
							end if
							conexion.close
						else
							f_evento = request.form("f_evento")
							l_ev = len(f_evento)

							if l_ev > 0 then
								f_formateada = cdate(f_evento)
								response.write("Busco por fecha de evento")
								response.write("<br><p>El criterio de búsqueda introducido ha sido <b>"&f_formateada&"</b></p>")
					
								set conexion = server.createobject("ADODB.Connection")
								conexion.open("bd")
								busqueda = "SELECT e.codigo, e.actividad, e.cliente, e.fecha_contrato, e.fecha_evento, a.Codigo, a.Nombre as nom_acti, c.Codigo, c.Nombre as nom_cli FROM eventos e, actividad a, cliente c WHERE e.actividad=a.Codigo AND e.cliente=c.Codigo AND e.fecha_evento LIKE '%"&f_formateada&"%'"
								set datos = conexion.execute(busqueda)

								if datos.eof then
									response.write("<h3 class='error' align='center'>No se ha encontrado ningún evento en la fecha de evento indicada</h3>")
								else
									response.write("<table border align='center'>")
									response.write("<tr bgcolor=#7fb3d5><td><b>Nombre del cliente</b></td><td><b>Nombre de la actividad</b></td><td><b>Fecha de contrato</b></td><td><b>Fecha del evento</b></td></tr>")
									do while not datos.EoF
										response.write("<tr bgcolor=#a9cce3><td>"&datos("nom_cli")&"</td><td>"&datos("nom_acti")&"</td><td>"&datos("fecha_contrato")&"</td><td>"&datos("fecha_evento")&"</td></tr>")
										datos.movenext
									loop
								end if
								conexion.close
							else
								response.write("<h3 class='error' align='center'>Debe introducir algún criterio de búsqueda</h3>")
							end if
						end if
					end if
				end if
			end if
		%>
	</div>
</body>
</html>