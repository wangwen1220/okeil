<%
dbpath=server.mappath("managermmqasd.asp")  
Set conn=Server.CreateObject("ADODB.Connection")
'conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq="& dbpath
conn.open "provider=microsoft.jet.oledb.4.0;data source=" & dbpath 
%>