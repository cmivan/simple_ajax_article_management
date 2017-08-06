<!--#include file="email_config.asp"-->
<br />

<table border="0" align=center cellpadding=0 cellspacing=10 bgcolor="#FFFFFF">
  <tr> 
    <td width="600" align="center" valign="top" class="forumRow">
<!--#include file="email_top.asp"-->
    <form name="adduser" method="post" action="EMAIL_AddUser.asp">
      <table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="forMy">
          <tr> 
            <td height="25" align="center" class="forumRaw"><b><font color="#000000">&nbsp;手动添加邮件地址</font></b></td>
          </tr>
          <tr> 
            <td align="center" bgcolor="#ECF5FF" class="forumRow"> 
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="forMy">
                <tr> 
                  <td align="right" class="forumRow"><font color="#000000">E-mail地址</font></td>
                  <td valign="top" class="forumRow"> <input name="email" type="text" class="bk" size="35" maxlength="50">                  </td>
                  <td valign="top" class="forumRow"><input type="submit" name="Submit" value="提交保存" class="Tips_bo" />
                  <input type="hidden" name="action" value="adduser" /></td>
                </tr>
              </table>            </td>
          </tr>
        </table>
      </form>
    </td>
  </tr>
</table>

<%
if request("action")<>"adduser" then
response.end
else
email=request("email")
mailnow=now()
conn.execute "insert into email (dateandtime,email) values ('"&mailnow&"','"&email&"')"
response.write"<SCRIPT language=JavaScript>alert('帐号添加成功，你可以继续添加！');"
response.write "</SCRIPT>"
end if
%>