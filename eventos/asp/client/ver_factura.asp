<% @ CODEPAGE = 65001 %>
<!DOCTYPE html>
<html lang="es">
<head>
	<meta charset="UTF-8">
	<title>Ver facturas</title>
	<link rel="stylesheet" href="../../css/estilos.css">
</head>
<body>
	<%
		l_sesion = len(Session("cod"))
			if l_sesion > 0 then
				codigo = Session("cod")
				if codigo <= 1 then
					response.redirect("../../index.asp")
				end if
			else
				codigo = request.cookies("recordar")
					if codigo <= 1 then
						response.redirect("../../index.asp")
					end if 
			end if
	%>
	<!-- Cabecera -->
	<div class="header-admin">
		<h1>Listado de facturas</h1>
		<a href='../cliente.asp'><b>Volver al panel de administración</b></a> | <a href='../cerrar_sesion.asp'><b>Cerrar sesión</b></a>
	</div>
	<div class="alinea-centro">
		<%
			'Se accede a BD y muestran las facturas del cliente
			set conexion = server.createobject("ADODB.Connection")
			conexion.open("bd")
			consulta = "SELECT e.codigo, e.cliente, e.actividad, e.fecha_evento, a.Codigo, a.Nombre as nom_acti, a.Precio FROM eventos e, actividad a WHERE e.actividad = a.Codigo AND e.cliente="&codigo&""
			set datos = conexion.execute(consulta)

			total = 0 'variable encargada de guardar el precio total acumulado

			response.write("<table border align='center'>")
			response.write("<tr bgcolor=#7fb3d5><td><b>Actividad contratada</b></td><td><b>Fecha del evento</b></td><td><b>Precio</b></td></tr>")
			do while not datos.EoF 
				response.write("<tr bgcolor=#a9cce3><td>"&datos("nom_acti")&"</td><td>"&datos("fecha_evento")&"</td><td>"&datos("Precio")&"</td></tr>")
				pre = cint(datos("Precio"))
				total = total + pre
				datos.movenext
			loop
			response.write("</table><br>")

			response.write("<table border align='center'><tr bgcolor=#7fb3d5><td><b>Total gastado</b></td><td>"&total&" €</td></tr></table>")

			conexion.close
		%>
	</div>
</body>
</html>