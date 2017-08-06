<!--#include file="Bin/Charset.asp"-->
<!--#include file="Bin/adminConn.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=c_title%></title>
<style type="text/css">
<!--
*{
	padding:0px;
	margin:0px;
	font-family:Verdana, Arial, Helvetica, sans-serif;
}
body {
	margin: 0px;
	background:#F7F7F7;
	font-size:12px;
}
input{
	vertical-align:middle;
}
img{
	border:none;
	vertical-align:middle;
}
a{
	color:#333333;
}
a:hover{
	color:#FF3300;
	text-decoration:none;
}
.main{
	width:500px;
	margin:100px auto 0px;
	border:4px solid #EEE;
	background:#FFF;
	padding: 12px;
padding-top:5px;
}

.main .title{
	width:100%;
	height:25px;
	margin:0px auto;
	background:url(images/login_toptitle.jpg) -10px 0px no-repeat;
	line-height:25px;
	font-size:14px;
	letter-spacing:2px;
	color:#F60;
	font-weight:bold;
	text-align: center;
}

.main .login{
	width:100%;
	margin:10px auto 0px;
	overflow:hidden;
}
.main .login .inputbox{
	width:280px;
	float:left;
	background:url(images/login_input_hr.gif) right center no-repeat;
}
.main .login .inputbox dl{
	width:280px;
	height:35px;
	clear:both;
}
.main .login .inputbox dl dt{
	float:left;
	width:80px;
	height:31px;
	line-height:31px;
	text-align:right;
	font-weight:bold;
}
.main .login .inputbox dl dd{
	width:195px;
	float:right;
	padding-top:1px;
}
.main .login .inputbox dl dd input{
	font-size:12px;
	font-weight:bold;
	border:1px solid #888;
	padding:4px;
	background:url(images/login_input_bg.gif) left top no-repeat;
}


.main .login .butbox{
	float:left;
	width:200px;
	margin-left:12px;
}
.main .login .butbox dl{
	width:200px;
}
.main .login .butbox dl dt{
	width:160px;
	height:41px;
	padding-top:2px;
}
.main .login .butbox dl dt input{
	width:98px;
	height:33px;
	background:url(images/login_submit.gif) no-repeat;
	border:none;
	cursor:pointer;
}
.main .login .butbox dl dd{
	height:21px;
	line-height:21px;
}
.main .login .butbox dl dd a{
	margin:5px;
}



.main .msg{
	width:560px;
	margin:10px auto;
	clear:both;
	line-height:17px;
	padding:6px;
	border:1px solid #FC9;
	background:#FFFFCC;
	color:#666;
}

.copyright{
	width:500px;
	text-align:right;
	margin:10px auto;
	font-size:10px;
	color:#999999;
}
.copyright a{
	font-weight:bold;
	color:#F63;
	text-decoration:none;
}
.copyright a:hover{
	color:#000;
}
-->
</style>
<script type="text/javascript" language="javascript">
<!--
	window.onload = function (){
		userid = document.getElementById("user_name");
		userid.focus();
	}
-->
</script>
</head>
<body>

	<div class="main">
		<div class="title">
			管理登陆
		</div>

		<div class="login">
<form name="log_in" method="post" action="">
            <div class="inputbox">
				<dl>
					<dt>账号：</dt>
					<dd><input type="text" name="user_name" id="user_name" size="20" onfocus="this.style.borderColor='#F93'" onblur="this.style.borderColor='#888'" />
					</dd>
				</dl>
				<dl>
					<dt>密码：</dt>
					<dd><input type="password" name="password" size="20" onfocus="this.style.borderColor='#F93'" onblur="this.style.borderColor='#888'" />
					</dd>
				</dl>

				<dl>
					<dt>验证码：</dt>
<dd>
<input name="verifycode" type="text" maxlength="5" size="20" onfocus="this.style.borderColor='#F93'" onblur="this.style.borderColor='#888'" style="background-image:url(Bin/code.asp); background-position:92% 40%; background-repeat:no-repeat"  />
</dd>
</dl>


          </div>
            <div class="butbox"><dl>
					<dt><input name="submit" type="submit" value=" "/>
					</dt>
			  </dl>
			</div>
		</form>
		</div>


	</div>

	<div class="copyright">
		Powered by <a href="#">天方v1.3</a>&nbsp;|&nbsp;<script src="http://s11.cnzz.com/stat.php?id=2058869&web_id=2058869" language="JavaScript"></script></div>



<script type="text/javascript">
var gaJsHost = (("https:" == document.location.protocol) ? "https://ssl." : "http://www.");
document.write(unescape("%3Cscript src='" + gaJsHost + "google-analytics.com/ga.js' type='text/javascript'%3E%3C/script%3E"));
</script>
<script type="text/javascript">
var gaJsHost = (("https:" == document.location.protocol) ? "https://ssl." : "http://www.");
document.write(unescape("%3Cscript src='" + gaJsHost + "google-analytics.com/ga.js' type='text/javascript'%3E%3C/script%3E"));
</script>
<script type="text/javascript">
try {
var pageTracker = _gat._getTracker("UA-15415383-2");
pageTracker._setDomainName("none");
pageTracker._setAllowLinker(true);
pageTracker._trackPageview();
} catch(err) {}</script>
</body>
</html>


<%
user_name=request.form("user_name")
password=md5(Request("password"))
'password=Request("password")
verifycode=request.form("verifycode")

if request.Form("submit")<>"" then

'####################################################
if user_name=""  then
response.Write("<script language=javascript>alert('用户名不能为空!');history.go(-1)</script>")
response.end
end if

if  password="" then
response.Write("<script language=javascript>alert('密码不能为空!');history.go(-1)</script>")
response.end
end if

if  verifycode="" then
response.Write("<script language=javascript>alert('验证码不能为空!');history.go(-1)</script>")
response.end
end if

if CStr(Session("kmcode"))<>cstr(Request.Form("VerifyCode")) then
response.Write("<script language=javascript>alert('验证码错误!');history.go(-1)</script>")
response.End
end if

    sql="select top 1 * from sys_admin where user_name='"&user_name&"' and password='"&password&"'"
set rs=sysconn.execute(sql)

if rs.eof or rs.bof then
   response.write "<script language=javascript>"
   response.write "alert('帐号或者密码错误，请重新输入!');"
   response.write "javascript:history.go(-1);"
   response.write "</script>"
else
   '#### 二重验证 ####-- 保证正常登陆
   if password<>rs("password") then
     response.write "<script language=javascript>"
     response.write "alert('帐号或者密码错误，请重新输入!');"
     response.write "javascript:history.go(-1);"
     response.write "</script>"
     response.End()
   else
	 session("user_id")     = rs("id")
     session("kmsys_admin") = user_name
	 session("adminname")   = rs("true_name")
	 session("user_power")  = rs("power")
	 session("user_editbox")= rs("editbox")

     response.redirect "system_admin.asp"
     response.End()
   end if
end if

'####################################################
end if
%>
