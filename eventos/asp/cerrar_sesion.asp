<%
	@ CODEPAGE = 65001
%>
<!DOCTYPE html>
<html lang="es">
<head>
	<meta charset="UTF-8">
	<title>Cerrar sesión</title>
</head>
<body>
	<%
		codigo = request.cookies("recordar")
		response.cookies("recordar") = ""
		response.cookies("recordar").expires = #01/01/2018#
		response.cookies("nombre") = ""
		response.cookies("nombre").expires = #01/01/2018#
		Session.Abandon
		response.redirect("../index.asp")
	%>
</body>
</html>