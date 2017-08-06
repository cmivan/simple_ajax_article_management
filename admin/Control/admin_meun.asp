<!--#include file="../Bin/ConfigSys.asp"-->

<%
'程序模块相对目录
dim urlPath
    urlPath="../System/"

'######### 获取完整 URL ,用于多语言版本间切换 #########
dim urlStr
    'urlStr="中文,cn|英文,en|日语,jp"
	'urlStrs=split(urlStr,"|")
	
'///////////////////////////////////////////////
Function OnUrl(lag)
    onStr=lcase(Request.ServerVariables("URL"))
	if instr(onStr,"/"&lcase(lag)&"/")<>0 then
	   onStyle="class=on"
	end if
	OnUrl=onStyle
End Function

'///////////////////////////////////////////////
Function GetUrl(lag)     
On Error Resume Next     
Dim strTemp     
  If LCase(Request.ServerVariables("HTTPS"))="off"   Then     
     strTemp="http://"     
  Else     
     strTemp="https://"     
  End If     
   strTemp=strTemp&Request.ServerVariables("SERVER_NAME")     
If Request.ServerVariables("SERVER_PORT")<>80 Then strTemp=strTemp&":"&Request.ServerVariables("SERVER_PORT") 
  
   ul=lcase(Request.ServerVariables("URL"))
   for each urls in urlStrs
       urls=lcase(urls)
       urlss=split(urls,",")
       ul=replace(ul,"/"&urlss(1)&"/","/"&lag&"/")
   next
   strTemp=strTemp&"/"&ul
If Trim(Request.QueryString)<>"" Then strTemp = strTemp&"?"&Trim(Request.QueryString)     
   GetUrl=strTemp     
End Function
%>

<style>
body{
background-color:#9E9C92;
background-image:url(../images/left_top_bg.gif);
background-position:left top;
background-repeat:repeat-y;
overflow-x:hidden;
}
</style>

<base target="main">
<body>
<div class="menu">
<table width="100%" border="0">
<tr>
<%
'多语言切换按钮
for each lannav in urlStrs
    nav=split(lannav,",")
if ubound(nav)=1 then
%>
<td align="center" id="lannav"><a href="<%=GetUrl(nav(1))%>" <%=OnUrl(nav(1))%> target="_self" ><%=nav(0)%></a></td>
<%
end if
next
%>	
  </tr>
</table>


<%	
set rs=server.createobject("adodb.recordset") 
	exec="select * from sys_nav where type_id=0 and show=1 order by order_id asc,id desc" 
    rs.open exec,sysconn,1,1 
 if rs.eof then
	response.Write "&nbsp;暂无"&db_title&"分导航！"
 else
    dim i
	do while not rs.eof
	i=i+1
%>
<!-- Item Strat -->
<dl>
<%
if rs("url")<>"" then 
   thisUrl=urlPath&rs("url")
else
   thisUrl="javascript:void(0);"
end if
%>
<dt><a href="<%=thisUrl%>" id="imgmenu<%=i%>" onClick="showsubmenu(<%=i%>)" ><%=rs("title")%></a></dt>
<dd id="items2" style="display:block;">
<ul id="submenu<%=i%>" <%if i<>1 then%>style="display:none"<%else%>style="display:block"<%end if%>>
<%
dim db_showSet
    db_showSet=rs("db_showSet")
if db_showSet<>"" and ubound(split(db_showSet,"|"))=2 then 

'########## 情况3, 符合数组
   arr_nav(2)=""
   arr_nav=split(db_showSet,"|")
if arr_nav(0)<>"" and arr_nav(1)<>"" and arr_nav(2)<>"" and isnumeric(arr_nav(2)) then
    select case arr_nav(2)
		   case 2    '只有管理、添加菜单
%>
<li><a target="main" href="<%=urlPath&arr_nav(1)%>/article_manage.asp"><%=arr_nav(0)%>管理</a></li>
<li><a target="main" href="<%=urlPath&arr_nav(1)%>/article_edit.asp"><%=arr_nav(0)%>添加</a></li>
<%		  
		   case 3    '管理、添加、分类菜单具备
%>
<li><a target="main" href="<%=urlPath&arr_nav(1)%>/article_manage.asp"><%=arr_nav(0)%>管理</a></li>
<li><a target="main" href="<%=urlPath&arr_nav(1)%>/article_edit.asp"><%=arr_nav(0)%>添加</a></li>
<li><a target="main" href="<%=urlPath&arr_nav(1)%>/article_type.asp"><%=arr_nav(0)%>分类</a></li>
<%	
		   case else '只有管理菜单
%>
<li><a target="main" href="<%=urlPath&arr_nav(1)%>/article_manage.asp"><%=arr_nav(0)%>管理</a></li>
<%	   
    end select
end if
end if
%>


<%
set js=server.CreateObject("ADODB.Recordset")
	sql1="select * from sys_nav where type_id="&rs("id")&" and show=1 order by order_id asc,id desc"
	js.open sql1,sysconn,1,1
	do while not js.eof
%>
<%
if js("title")="-" then
'########## 情况1, "-" 为 分行符
%>
<div style="line-height:2px; height:2px; padding:0; margin:0;background-color:#FFFFFF;">&nbsp;</div>
<%
elseif left(js("title"),1)="#" then
'########## 情况2, "#company" 其中"#"为标识符 "company" 为表名
%>
<%
Tstr=js("title")
Tstr=replace(Tstr,"#","")
if Tstr<>"" then
	set ts=server.CreateObject("ADODB.Recordset")
	    tssql1="select * from "&Tstr&" order by order_id asc,id desc"
		ts.open tssql1,conn,1,1
		do while not ts.eof
%>
<li><a target="main" href="<%=urlPath&Tstr%>/article_edit.asp?id=<%=ts("id")%>"><%=ts("title")%></a></li>
<%
		ts.movenext
		loop
		ts.close
	set ts=nothing
end if
%>
<%else%>
<li><a target="main" href="<%="../"&js("url")%>"><%=js("title")%></a></li>
<%end if%>
<%
		js.movenext
		loop
		js.close
	set js=nothing
%>
</ul>
</dd>

</dl>
<%		
	rs.movenext
	loop
 end if
    rs.close
set rs=nothing
%>
<dl>
<dt><a>版权信息</a></dt>
<dd>
<ul>
<li>
<a href="#" target="_blank" style=" color:#333333;line-height:20px;padding:3px; padding-left:12px; margin:0; text-indent:0; background:none; border-left:#CCCCCC 1px solid; border-top:#CCCCCC 1px solid">
程序开发：天方。雨<br>
联系&nbsp;&nbsp;QQ：619835864<br>
官方网站：天方网络<br>
</a>
</li>
</ul>
</dd>
</dl>

<dl>
<dt><a href="?login=out" target="_top">退出管理</a></dt>
</dl>
<!-- Item End -->






<script>
var onTitle=1;
function showsubmenu(sid)
{
whichEl = eval("submenu" + sid);
imgmenu = eval("imgmenu" + sid);

for (i=1;i<=<%=i%>;i++){
eval("submenu" + i + ".style.display=\"none\";");
eval("imgmenu" + i + ".className=\"menu_title\";");
}

if (whichEl.style.display == "none"&&onTitle!=sid)
{
whichEl.style.display="block";
imgmenu.className="menu_title2";
}else{
whichEl.style.display="none";
imgmenu.className="menu_title";
sid=0;
}
onTitle=sid;          //记录当前栏目id
}
</script>
</div>
</body>
