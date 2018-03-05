<% @ CODEPAGE = 65001 %>
<!DOCTYPE html>
<html lang="es">
<head>
	<meta charset="UTF-8">
	<title>Aplicación eventos ASP</title>
	<link rel="stylesheet" href="css/estilos.css">
</head>
<body>
	<%
		'Compruebo sesion y cookies para restringir usuarios y muestro encabezados de la web
		l_sesion = len(Session("cod"))
		l_recordar = len(request.cookies("recordar"))
			if l_sesion > 0 then
				response.write("Entro con sesión")
				if Session("cod") = 1 then
					response.redirect("asp/administrador.asp")
				elseif Session("cod") > 1 then
					response.redirect("asp/cliente.asp")
				else
					
				end if
			elseif l_recordar > 0 then
				response.write("Entro con cookies")
				if request.cookies("recordar") = 1 then
					response.redirect("asp/administrador.asp")
				elseif request.cookies("recordar") > 1 then
					response.redirect("asp/cliente.asp")
				else 
					
				end if
			else
				
			end if 

	%>
	<div class="alinea-centro">
		<h1>Aplicación eventos ASP</h1>
		<h2>Login de acceso</h2>
		<form action="#" method="post">
			<table>
				<tr>
					<td>
						<b>Usuario</b>
					</td>
					<td>
						<input type="text" name="usuario" autofocus>
					</td>
				</tr>
				<tr>
					<td>
						<b>Contraseña</b>
					</td>
					<td>
						<input type="password" name="pass">
					</td>
				</tr>
				<tr>
					<td>
						<b>Recordar <input type="checkbox" name="recordar"></b>
					</td>
					<td>
						<input type="submit" name="acceder" value="Acceder">
					</td>
				</tr>
			</table>
		</form>
	</div>
	<% 

		l_acceder = len(request.form("acceder"))

		if l_acceder > 0 then
			'recojo usuario y contraseña cuando se intenta acceder
			'response.write("-- intento acceder --")
			usuario = request.form("usuario")
			pass = request.form("pass")
			'response.write("<br>usu: "&usuario&" - pass: "&pass)
			'consulto en la BD si existe el usuario y contraseña indicados
			set conexion = server.createobject("ADODB.Connection")
			conexion.open("bd")
			consulta = "SELECT Nombre,Contra,Codigo from cliente WHERE Nombre='"&usuario&"' AND Contra='"&pass&"'"
			set datos = conexion.execute(consulta)

			'response.write(datos.EoF)
			if not datos.EoF  then
				'response.write("<br>hay un registro")
				'response.write("<br>Codigo del usuario: "&datos("codigo"))
				codigo = datos("Codigo")
				'response.write("<br>Cod:"&codigo)
				Session("cod")=codigo
				Session("nombre")=datos("Nombre")
				'response.write("<br>"&Session("nombre"))
				'compruebo si ha pulsado en recordar sesión para guardar una cookie
				l_check = len(request.form("recordar"))
				if l_check > 0 then
					response.write("<br>recuerdo sesión")
					response.cookies("recordar") = codigo
					response.cookies("nombre") = cstr(datos("Nombre"))
					response.cookies("recordar").expires = #02/10/2021#
					response.cookies("nombre").expires = #02/10/2021#
					if Session("cod") = "1" then
						response.redirect "asp/administrador.asp"
					else 
						response.redirect "asp/cliente.asp"
					end if	
				else
					if Session("cod") = "1" then
						response.redirect "asp/administrador.asp"
					else 
						response.redirect "asp/cliente.asp"
					end if	
				end if
			else
				response.write("<br>No hay coincidencias")
				response.write("<h2 class='error' align='center> El usuario o contraseña introducido no es correcto</h2>")
			end if
			conexion.close
		end if

	%>
</body>
</html>