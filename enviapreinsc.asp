<html>

<head>
<meta name="GENERATOR" content="Microsoft FrontPage 12.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="robots" content="index,follow">
<meta name="keywords" content="Ushuaia, Rio Grande, Tierra del Fuego, Vuelta, Enduro, Carrera, Off road, CrossWorld, Argentina, Fin del mundo" >
<meta name="description" content="XXVI Edicion de la Vuelta a Tierra del Fuego. Rio Grande - Ushuaia - Rio Grande. La carrera del Fin del mundo. Tierra del Fuego. Argentina."/>
<title>.:: Enduro del Fin del Mundo :: XXVI Vuelta a Tierra del Fuego ::.</title>
</head>

<body background="images/bg.jpg" topmargin="0">
<CENTER>
<script type="text/javascript"><!--
google_ad_client = "pub-6389487984360740";
/* 728x90, created 2/19/10 */
google_ad_slot = "3130797248";
google_ad_width = 728;
google_ad_height = 90;
//-->
</script>
<script type="text/javascript"
src="http://pagead2.googlesyndication.com/pagead/show_ads.js">
</script>

<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="450" height="281">
    <tr>
      <td height="47"><img border="0" src="images/1.jpg" width="766" height="47"></td>
    </tr>
    <tr>
      <td background="images/2.jpg" height="194">
      
      <%
Dim numero_azar
' Empezamos la funci�n Ramdomize.
randomize
' Buscamos el numero entre 1 y 5.
numero_azar = Int (Rnd*5) + 1
'Carga de imagen aleaotoria.
  %>  
      <img border="0" src="images/2_<%=numero_azar%>.jpg" width="766" height="194">
      
      
      
      
      </td>
    </tr>
    <tr>
      <td background="images/3.jpg" height="16">
      <map name="FPMap0">
      <area href="index.html" shape="rect" coords="6, 3, 131, 27">
      <area href="mapas.html" shape="rect" coords="145, 5, 254, 29">
      <area href="preinsc.html" shape="rect" coords="269, 3, 374, 26">
      <area href="sponsors.html" shape="rect" coords="390, 5, 495, 27">
      <area href="contacto.html" shape="rect" coords="509, 4, 619, 28">
      <area href="reglamento.html" shape="rect" coords="630, 4, 757, 27">
      </map>
      <font size="2" face="Verdana">
      <img border="0" src="images/menu.jpg" usemap="#FPMap0" width="766" height="39"></font></td>
    </tr>
    
    <%
    DIM humano
    humano = request("humano")
    if humano = "Si" then
       
    %>
    
    <%


DIM cuerpo

DIM asunto
asunto = "Pre-inscripci�n"

DIM categoria
categoria = request("categoria")

DIM nombre
nombre = request("nombre")

DIM nacimiento
nacimiento = request("nacimiento")

DIM edad
edad = request("edad")

DIM documento
documento = request("documento")

DIM domicilio
domicilio = request("domicilio")

DIM ciudad
ciudad = request.form("ciudad")

DIM pais
pais = request("pais")

DIM telefono
telefono = request("telefono")

DIM celular
celular = request("celular")

DIM email
email = request("email")

DIM vehiculo
vehiculo = request("vehiculo")

DIM marca
marca = request("marca")

DIM modelo
modelo = request("modelo")

DIM anio
anio = request("anio")

DIM cc
cc = request("cc")

DIM consulta
consulta = request("consulta")

DIM cdoHigh



cuerpo = VbCrLf
cuerpo = cuerpo & "Email generado desde la web www.endurofindelmundo.com.ar" & VbCrLf & VbCrLf
cuerpo = cuerpo & "Nombre = "  & nombre & VbCrLf
cuerpo = cuerpo & "Nacido el = "  & nacimiento & VbCrLf
cuerpo = cuerpo & "Edad = "  & edad & VbCrLf
cuerpo = cuerpo & "Documento Tipo y nro = "  & documento & VbCrLf
cuerpo = cuerpo & "Domicilio = "  & domicilio & VbCrLf
cuerpo = cuerpo & "Ciudad/Provincia = "  & ciudad & VbCrLf
cuerpo = cuerpo & "Pa�s= "  & pais & VbCrLf
cuerpo = cuerpo & "Telefono = "  & telefono & VbCrLf
cuerpo = cuerpo & "Celular = "  & celular & VbCrLf
cuerpo = cuerpo & "Email = "  & email & VbCrLf
cuerpo = cuerpo & "Categoria = "  & categoria & VbCrLf
cuerpo = cuerpo & "Veh�culo = "  & vehiculo & VbCrLf
cuerpo = cuerpo & "Marca = "  & marca & VbCrLf
cuerpo = cuerpo & "Modelo = "  & modelo & VbCrLf
cuerpo = cuerpo & "A�o = "  & anio & VbCrLf
cuerpo = cuerpo & "Cilindrada = "  & cc & VbCrLf & VbCrLf
cuerpo = cuerpo & "Consulta/Mensaje = "  & consulta & VbCrLf & VbCrLf

' Mando email
Dim objMail

Set objMail = CreateObject("CDONTS.NewMail")

objMail.From = email
objMail.To = "info@luxmedia.com.ar"
objMail.Subject = asunto
objMail.Body = cuerpo
objMail.importance = cdoHigh
objMail.Send
Set objMail = nothing



'response.redirect "gracias.asp"

%>
    
    <tr>
      <td background="images/3.jpg" height="408">
         <form method="POST" action="enviapreinsc.asp"> 
         <div align="center">
        <center>
        <table border="0" cellpadding="7" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="679" height="204">
          <tr>
            <td width="651" height="1" bgcolor="#FFFFCC">
        
         
            <p align="center"><b><font color="#800000" face="Verdana">
            <img border="0" src="images/LOGO-VUELTA-XXVI.gif" hspace="3"></font></b><p align="center"><b><font color="#000080" face="Verdana" size="4">
            Pre Inscripci�n enviada con �xito !</font></b><p align="center"><b>
            <font face="Verdana" size="4" color="#000080">Gracias !!
            
   <%else%>         
            <font face="Verdana" size="4" color="#000080">O contestaste mal la pregunta trampa o eres un asqueroso robot ...
            
         <%end if%>   
            
            
            </font></font></b><p align="center"><b>
            <font face="Verdana" size="4" color="#000080"><a href="http://www.anrdoezrs.net/click-3774875-10576254" target="_top" onmouseover="window.status='http://www.skype.com';return true;" onmouseout="window.status=' ';return true;">
<img src="http://www.tqlkg.com/image-3774875-10576254" width="468" height="60" alt="" border="0"/></a></font></b></td>
          </tr>
          </table>
        </center>
      </div>
         <p>&nbsp;</p>
         <p>&nbsp;</p>
      <p align="center">
 
         </form>
         <p align="center"><script type="text/javascript"><!--
google_ad_client = "pub-6389487984360740";
/* 728x90, creado 16/02/09 */
google_ad_slot = "8656157082";
google_ad_width = 728;
google_ad_height = 90;
//-->
         </script>
<script type="text/javascript"
src="http://pagead2.googlesyndication.com/pagead/show_ads.js">
         </script></td>
    </tr>
    <tr>
      <td background="images/4.jpg" height="171">
      <p align="center">&nbsp;</p>
      <p align="center"><a target="_blank" href="http://www.riogrande.gov.ar/">
      <img border="0" src="images/RG4.jpg"></a></p>
      <p align="center"><!-- Histats.com  START  -->
<a href="http://www.histats.com/es/" target="_blank" title="contador de visitas" ><script  type="text/javascript" language="javascript">
var s_sid = 613618;var st_dominio = 4;
var cimg = 451;var cwi =112;var che =61;
      </script></a>
<script  type="text/javascript" language="javascript" src="http://s11.histats.com/js9.js"></script>
<noscript>
      <a href="http://www.histats.com/es/" target="_blank">
<img  src="http://s103.histats.com/stats/0.gif?613618&1" alt="contador de visitas" border="0"></a></noscript>
<!-- Histats.com  END  -->&nbsp;
      <i><b><font color="#FFC500" face="Verdana" size="2">&nbsp;</font><font color="#800000" face="Verdana" size="2">Powered by </font></b></i>
            <span lang="es">
            
            
                  <a href="/luxmedia"><font color="#800000">
      <a href="http://www.crossworldtv.com.ar/luxmedia">
      <img border="0" src="images/logos/LM.jpg" width="80" height="56"></a></font></a></span></td>
    </tr>
  </table>
  </center>
</div>

</body>

</html>