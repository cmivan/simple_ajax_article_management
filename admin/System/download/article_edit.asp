﻿<!--配置文件-->
<!--#include file="article_config.asp"-->
<%
'///////////  处理提交数据部分 //////////
dim edit_id
    edit_id   =request("id")
 if edit_id="" or isnumeric(edit_id)=false then
    editStr="添加"
 else
    editStr="修改"
 end if
   
if request.Form("edit")="ok" then

Tf_title=request.form("Tf_title")
type_id=request.form("type_id")
'//系判断是否为数组，再判断是否为数字(否则失败)
type_ids=split(type_id,",")
if ubound(type_ids)=1 then
   if type_ids(0)<>"" and isnumeric(type_ids(0)) and type_ids(1)<>"" and isnumeric(type_ids(1)) then
	  typeB_id=type_ids(0)
	  typeS_id=type_ids(1)
	  session("type_id")=type_id   '记录本次操作分类
   else
	  response.Write("<script>alert('分类有误!请重新选择.');history.back(1);</script>")
	  response.End()
   end if
else
   if type_id<>"" and isnumeric(type_id) then
	  typeB_id=type_id
	  typeS_id=null
	  session("type_id")=type_id   '记录本次操作分类
   else
	  response.Write("<script>alert('分类有误!请重新选择.');history.back(1);</script>")
	  response.End()
   end if
end if
if isdate(Tf_add_data)=false then Tf_add_data=now()
if Tf_ok="" or isnumeric(Tf_ok)=false then Tf_ok=0
if Tf_hot="" or isnumeric(Tf_hot)=false then Tf_hot=0
if Tf_news="" or isnumeric(Tf_news)=false then Tf_news=0
if Tf_tj="" or isnumeric(Tf_tj)=false then Tf_tj=0


'///////////  写入数据部分 //////////
set rs=server.createobject("adodb.recordset") 
 if edit_id="" or isnumeric(edit_id)=false then
    exec="select * from "&db_table                       '判断，添加数据
	rs.open exec,conn,1,3
    rs.addnew
 else
	exec="select * from "&db_table&" where id="&edit_id  '判断，修改数据
    rs.open exec,conn,1,3
 end if
 if edit_id<>"" and isnumeric(edit_id) and rs.eof then
	response.Write("写入数据失败!")
 else
	
rs("title")=Tf_title
rs("typeS_id") =typeS_id
	  Tf_ok =1
else
   Tf_ok =0
end if
rs("ok")=Tf_ok
	  Tf_hot =1
else
   Tf_hot =0
end if
rs("hot")=Tf_hot
	  Tf_news =1
else
   Tf_news =0
end if
rs("news")=Tf_news
	  Tf_tj =1
else
   Tf_tj =0
end if
rs("tj")=Tf_tj

	end if
	rs.update
	rs.close
set rs=nothing
    call backPage(editStr&"操作成功!","article_manage.asp?typdB_id="&typdB_id&"&typdS_id="&typdS_id,0)
end if

'///////////  读取数据部分
   id=request.QueryString("id") 
if id<>"" and isnumeric(id) then
set rs=server.createobject("adodb.recordset") 
    exec="select * from "&db_table&" where id="&id
    rs.open exec,conn,1,1 
	if not rs.eof then
Tf_title=rs("title")
typeS_id  =rs("typeS_id")
'记录分类,用于分类下拉
if typeS_id="" or isnumeric(typeS_id)=false then
session("type_id")=typeB_id
else
session("type_id")=typeB_id&","&typeS_id
end if
if Tf_order_id="" or isnumeric(Tf_order_id)=false then Tf_order_id=0
if Tf_ok="" or isnumeric(Tf_ok)=false then Tf_ok=0
if Tf_hot="" or isnumeric(Tf_hot)=false then Tf_hot=0
if Tf_news="" or isnumeric(Tf_news)=false then Tf_news=0
if Tf_tj="" or isnumeric(Tf_tj)=false then Tf_tj=0
	end if
	rs.close
set rs=nothing
   
end if
%>


<script>
function CheckForm()
{

Tips(FormE.Tf_add_data);
FormE.Tf_add_data.focus();
return false;
}
}
///返回提示框
function Tips(ThisF){
  ThisF.style.background="#FF6600";
}
</script>
<body>
<TABLE width="800" border="0" align="center" cellpadding="0" cellspacing="10" bgcolor="#FFFFFF">
<TR><td class="forumRow">
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="forMy forMAargin"><tr><td align="center" class="forumRow">
<strong>&nbsp;&nbsp;<%=db_title%> <%=editStr%></strong>
</td></tr></table>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="forMy">
<form name="FormE" method="post" action="" onSubmit="return CheckForm(this);">
<tr><td align='right' class='forumrow'>标题：</td><td class='forumrow'>
<input type="text" name="Tf_title" id="Tf_title" value="<%=Tf_title%>" />
</td></tr>
<textarea name="Tf_note" style="width:100%" rows="4"><%=Tf_note%></textarea>


</td></tr>
<!--#include file="../../Class/articles_type.asp"-->


</td></tr>
<input name="Tf_small_pic" id="Tf_small_pic" type="text" value="<%=Tf_small_pic%>" size="45" maxlength="255" />
<input type=button value="浏览图片" onClick="showUploadDialog('image', 'FormE.Tf_small_pic', '')">



</td></tr>
<input type="text" name="Tf_add_data" id="Tf_add_data" value="<%=Tf_add_data%>" />
</td></tr>
<input type="text" name="Tf_order_id" id="Tf_order_id" value="<%=Tf_order_id%>" />
</td></tr>
<input type="checkbox" name="Tf_ok"<%if Tf_ok=1 then%> checked<%end if%> value="1">
</td></tr>
<input type="checkbox" name="Tf_hot"<%if Tf_hot=1 then%> checked<%end if%> value="1">
</td></tr>
<input type="checkbox" name="Tf_news"<%if Tf_news=1 then%> checked<%end if%> value="1">
</td></tr>
<input type="checkbox" name="Tf_tj"<%if Tf_tj=1 then%> checked<%end if%> value="1">
</td></tr>
<input type="text" name="Tf_toUrl" id="Tf_toUrl" value="<%=Tf_toUrl%>" />
</td></tr>
<tr>
<td width="88" align="right" valign="top" class="forumRow">
<input name="id" type="hidden" value="<%=id%>" />
<input name="edit" type="hidden" value="ok" />
</td><td class="forumRow">
<input name="submit" type="submit" value="确认<%=editStr%>" />
</td></tr>
</form>
</table></td></tr></table>
</body>
</html>