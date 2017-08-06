<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% Response.CodePage=65001%>
<% Response.Charset="UTF-8" %>
<%
'=++++++++++++++++++++++++++++++++++++++++=
on error resume next
dim conns,connstr,mdbs
'-----------  Access数据库连接字符串 -----------
    mdbs="data/KM_09_12_20.mdb"           '数据库文件目录
    Connstr="DRIVER=Microsoft Access Driver (*.mdb);DBQ="+server.mappath(mdbs)
'-----------  SQL数据库连接字符串    -----------
'	Connstr="driver={SQL Server};server=192.168.0.1,7788;database=KM_09_12_20;UID=sa;PWD=sa"
'---------------------------------------------
Set conn=Server.CreateObject("ADODB.Connection") 
    conn.Open connstr
'------------------------------------------==
'response.Write(server.mappath(mdbs))
'response.End()

 If Err Then
    Err.Clear
    Set Conns = Nothing
    Response.Write "数据库连接有误!请检查相应文件!"
    Response.End
 End If
'=++++++++++++++++++++++++++++++++++++++++=
%>
<!--#include file="function.asp"-->
<!--#include file="class_db.asp"-->