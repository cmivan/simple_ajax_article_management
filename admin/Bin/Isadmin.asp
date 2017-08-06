<%
'### 容错模式
 on error resume next
 
'### 验证是否已登陆
 if errTip="" then errTip="登陆超时或未登录,请重新登陆后再进行操作!"
 if session("kmsys_admin")="" then
    call backPage(errTip,Rpath&Rpath&"../",0)
 end if

'### 注销登陆
 if request.QueryString("login")="out" then
    session.Abandon()
    call backPage("注销成功!",Rpath&"../",0)
 end if
%>