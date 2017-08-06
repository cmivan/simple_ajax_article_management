<%
'× --------------------------------------
'× -------  用于返回热门最新推荐审核等  -----
'× --------------------------------------
  function getBool(num)
  	  if num="1" then
	     getBool =1
	  else
	     getBool =0
	  end if
  end function


'× --------------------------------------
'× ----------  返回提示信息 ---------------
'× --------------------------------------
  function backPage(backStr,backUrl,backType)
    back =""
	back =back&"<meta http-equiv=Content-Type content=text/html; charset=utf-8 />"
	back =back&"<link href='"&Rpath&"images/style.css' rel='stylesheet' type='text/css' />"
	
	if backType=0 then
	    'meta自动跳转到指定页面
        back =back&"<meta http-equiv=refresh content=1;url="&backUrl&">"
		back =back&"<body style=""overflow:hidden;"">"
	    back =back&"<br><TABLE border=0 align=center cellpadding=0 cellspacing=10 bgcolor=#FFFFFF><tr><td width=90% class=forumRow>"
		back =back&"<table width=100% border=0 align=center cellpadding=3 cellspacing=1 class=forMy><tr><td class=forumRow align=center>"
		back =back&backStr
		back =back&"</tr></table></td></tr></table>"
	elseif backType=1 then
	    'js弹出提示，返回指定页面
	    back =back&"<script language='javascript'>alert('"&backStr&"');window.location.href='"&backUrl&"';</script>"
	elseif backType=2 then
	    'js弹出提示，返回上一级
	    back =back&"<script language='javascript'>alert('"&backStr&"');history.back(1);</script>"
	elseif backType=3 then
	    'js弹出提示，返回指定页
	    back =back&"<script language='javascript'>window.location.href='"&backUrl&"';</script>"
	end if
	response.Write(back)
	response.End()
 end function
%>