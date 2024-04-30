<%
option explicit
Server.scriptTimeout = 1
%>
<!--#include file="../src/aspZip.class.asp" -->
<%
dim zip, filepath

filepath = "test.zip"

set zip = new aspZip

zip.OpenArchieve(filepath)

zip.Add("..\src")
zip.Add(".\default.asp")

zip.CloseArchieve()

zip.ExtractTo(".\test")

set zip = nothing
%>
OK