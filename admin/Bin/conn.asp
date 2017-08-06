<%
'=+++++++++++++++++连接系统部分数据库+++++++++++++++++++++=
On error resume next
dim conn,connstr,mdbs
'############  Access数据库连接字符串  ###########
    mdbs=Rpath&"../data/KM_09_12_20.mdb"           '数据库文件目录
    connstr="DRIVER=Microsoft Access Driver (*.mdb);DBQ="+server.mappath(mdbs)
	
'############  SQL数据库连接字符串    ############
'	Connstr="driver={SQL Server};server=192.168.0.1,7788;database=KM_09_12_20;UID=sa;PWD=sa"
'###############################################
set conn=Server.CreateObject("ADODB.Connection") 
    conn.Open connstr
'------------------------------------------==
 If err then
    err.clear
    set conn = Nothing
    response.Write "数据库连接有误!请检查相应文件!"
    response.end
 end If
'------------------------------------------==
%>