<!--配置文件-->
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

{接收数据}
{数据判断}

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
	
{写入数据}

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
{读取数据}
	end if
	rs.close
set rs=nothing
   
end if
%>


<script>
function CheckForm()
{
{表单判断}
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
{表单数据}
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