<!--#include file="md5.asp"-->
<!--#include file="Function.asp"-->

<%
'=+++++++++++++++++连接系统部分数据库+++++++++++++++++++++=
On error resume next
dim sysconn,sysconnstr,sysmdbs
'############  Access数据库连接字符串  ###########
    sysmdbs=Rpath&"Bin/TF_DB_10_03_03.ASA"           '数据库文件目录
    sysconnstr="DRIVER=Microsoft Access Driver (*.mdb);DBQ="+server.mappath(sysmdbs)

'############  SQL数据库连接字符串    ############
'	Connstr="driver={SQL Server};server=192.168.0.1,7788;database=KM_09_12_20;UID=sa;PWD=sa"
'###############################################
	
set sysconn=Server.CreateObject("ADODB.Connection") 
    sysconn.Open sysconnstr
'------------------------------------------==
 If err then
    err.clear
    set sysconn = Nothing
    response.Write "数据库连接有误!请检查相应文件!"
    response.end
 end If
'=++++++++++++++++++++++++++++++++++++++++=


'###### 读取配置信息  ###### 
set config=server.createobject("adodb.recordset") 
    exec="select * from sys_config" 
    config.open exec,sysconn,1,1
	if not config.eof then
	   c_title=config("title")
	   c_url  =config("url")
	   c_contact=config("contact")
	   c_tel=config("tel")
	   c_fax=config("fax")
	   c_mobile=config("mobile")
	   c_email=config("email")
	   c_qq=config("qq")
	   c_msn=config("msn")
	   c_address=config("address")
	   c_keywords=config("keywords")
	   c_description=config("description")
	   c_copyright=config("copyright")
	end if
	config.close
set config=nothing
%>

<!--#include file="conn.asp"-->