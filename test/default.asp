<%
option explicit
Server.scriptTimeout = 10
%>
<!--#include file="../src/aspZip.class.asp" -->
<%
dim zip, filepath

filepath = "test.zip"

set zip = new aspZip

zip.OpenArquieve(filepath)

'zip.AddFile("../src")

'zip.CloseArquieve()

zip.Extract(".\test")

set zip = nothing
%>
OK