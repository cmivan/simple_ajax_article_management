<!--#include file="../Bin/ConfigSys.asp"-->
<%
'=========== 修改账号信息 ===============
if session("user_id")<>"" and request.Form("edit")="yes" then
    editbox  =request.form("editbox")
    user_name=request.form("user_name")
    user_name=replace(user_name,"'","")
    true_name=request.form("true_name")
    password =request.form("password")


 if editbox="" or isnumeric(editbox)=false then editbox=0

set rs=server.CreateObject("adodb.recordset")
    sql="select * from sys_admin where id="&session("user_id")
    rs.open sql,sysconn,1,3
 if not rs.eof then

    if user_name<>"" then rs("user_name")=user_name
	if true_name<>"" then rs("true_name")=true_name
	if password<>"" then rs("password")=md5(password)
	rs("editbox") =editbox
	session("user_editbox")= editbox

 end if
    rs.update 
    rs.close 
set rs=nothing
%>
<script language=javascript>
	alert("登陆帐号及密码修改成功，请牢记您帐号信息!");
	window.location.href="?";
</script>
<%end if%>



<script type="text/javascript">
function check(){
     if (document.password.user_name.value.length=="")
     {
           alert("用户名不能为空!");
           document.password.user_name.focus();
           return false;
     }

	 
     if (document.password.password.value!=document.password.password_again.value)
     {
           alert("两次输入密码不一致，请重新输入！");
           document.password.password.focus();
           return false;
     }
           return true;
}

</script>
<body>
<table width="450" border="0" align="center" cellpadding="0" cellspacing="10" bgcolor="#FFFFFF">
  <tr>
    <td valign="top" class="forumRow">
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="forMy forMAargin">
        <tr>
          <td align="center" class="forumRow">&nbsp;&nbsp;<strong>修改密码</strong></td>
        </tr>
</table>
	
<% 
set rs=server.createobject("adodb.recordset") 
    exec="select * from sys_admin where id="&session("user_id")
    rs.open exec,sysconn,1,1 
	if not rs.eof then
%>
      <form name="password" method="post" id="password" action="" onSubmit="return check()">
<table width="100%" border="0" align=center cellpadding=4 cellspacing=1 class="forMy">
          <tr>
            <td width="201" align="right" class="forumRow">账号：</td>
            <td width="366" class="forumRow"><input name="user_name" type="text" value="<%=rs("user_name")%>" size="30" disabled="disabled"/></td>
          </tr>
	 <tr>
            <td width="201" align="right" class="forumRow">称呼：</td>
            <td class="forumRow"><input name="true_name" type="text" value="<%=rs("true_name")%>" size="30" /></td>
          </tr>		  
          <tr>
            <td align="right" class="forumRow">密码：</td>
            <td class="forumRow"><input name="password" type="password" id="password" size="30" /></td>
          </tr>
          <tr>
            <td align="right" class="forumRow">确认密码：</td>
            <td class="forumRow"><input name="password_again" type="password" id="password_again" size="30" /></td>
          </tr>
          <tr>
            <td align="right" class="forumRow">编辑器：</td>
            <td class="forumRow">
			<%
			   editbox=rs("editbox")
			if editbox="" or isnumeric(editbox)=false then editbox=0
			%>
			<select name="editbox" id="editbox">
              <option value="0" <%if editbox=0 then%>selected<%end if%>>Ewebeditor&nbsp;-&nbsp;(常用)</option>
              <option value="1" <%if editbox=1 then%>selected<%end if%>>Fckeditor&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;-&nbsp;(支持IE8下运行)</option>
			  <option value="2" <%if editbox=2 then%>selected<%end if%>>Ewebeditor&nbsp;-&nbsp;(支持Word图片粘贴)</option>
<!--<option value="3" <%if editbox=3 then%>selected<%end if%>>163HtmlEditor</option>-->
            </select>
            </td>
          </tr>
          <tr>
            <td class="forumRow"><p class="submit">&nbsp;</p></td>
            <td align="left" class="forumRow"><span class="submit">
              <input name="reset" type="reset" value="重新填写" />
              <input name="submit" type="submit" value="修改密码" />
              <input name="edit" type="hidden" id="edit" value="yes" />
            </span></td>
          </tr>
        </table>
      </form>
<%
    end if
    rs.close
set rs=nothing
%>	</td>
  </tr>
</table>


</body>