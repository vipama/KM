<!--#include file="../../Config.inc.asp"-->
<form>
IP <input type="text" value="58.9.189." name="IP">
Date <input type="text" value="1/1/2006" name="Date">
Quanlity <input type="text" value="1" name="Quanlity">
<input type="submit">
</form>
<%for i=1 to request("Quanlity")
con.execute("insert into TabUniIP (IP,[Date]) values ('"&request("IP")&i&"',#"&request("Date")&"#)")

response.write "Add IP: "&request("IP")&i&" Date: "&request("Date")
next%>