<% 
	@ CODEPAGE = 65001 
%>
<!DOCTYPE html>
<html lang="es">
<head>
	<meta charset="UTF-8">
	<title>Listado clientes</title>
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
	<!-- Cabecera -->
	<div class="header-admin">
		<h1>Listado de clientes</h1>
		<a href='../administrador.asp'><b>Volver al panel de administración</b></a> | <a href='../cerrar_sesion.asp'><b>Cerrar sesión</b></a>
	</div>
	<div class="alinea-centro">
		<%
			'Se accebe a BD y muestran los clientes almacenados
			set conexion = server.createobject("ADODB.Connection")
			conexion.open("bd")
			consulta = "SELECT * FROM cliente ORDER BY Nombre ASC"
			set datos = conexion.execute(consulta)

			response.write("<table border align='center'>")
			response.write("<tr bgcolor=#7fb3d5><td><b>Código</b></td><td><b>Nombre</b></td><td><b>Teléfono</b></td><td><b>Dirección</b></td><td><b>Contraseña</b></td></tr>")
			do while not datos.EoF 
				response.write("<tr bgcolor=#a9cce3><td>"&datos("Codigo")&"</td><td>"&datos("Nombre")&"</td><td>"&datos("Telefono")&"</td><td>"&datos("Direccion")&"</td><td>"&datos("Contra")&"</td></tr>")
				datos.movenext
			loop
			response.write("</table>")
			conexion.close
		%>
	</div>
</body>
</html>