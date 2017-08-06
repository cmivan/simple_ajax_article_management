<!--#include file="../Bin/ConfigSys.asp"-->

<%
'#### 修改配置信息 #######
if request.Form("submit")<>"" then
set config=server.createobject("adodb.recordset")
    exec="select * from sys_config"
    config.open exec,sysconn,1,3
    config("title")=request.form("title")
    config("url")=request.form("url") 
    config("contact")=request.form("contact")
    config("tel")=request.form("tel")
    config("fax")=request.form("fax")
    config("mobile")=request.form("mobile")
    config("email")=request.form("email")
    config("qq")=request.form("qq")
    config("msn")=request.form("msn")
    config("address")=request.form("address")
    config("keywords")=request.form("keywords") 
    config("description")=request.form("description") 
    config("copyright")=request.form("copyright") 
    config.update 
    config.close 
   '返回操作信息
    admin_config="ok"
end if
%>


<body>
 <table width="600" border="0" align="center" cellpadding="0" cellspacing="10" bgcolor="#FFFFFF">
  <form id="config" name="config" method="post" action="">
   <tr>
     <td class="forumRow"><table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="forMy forMAargin" >
       <tr>
         <td align="center" class="forumRow"><strong style="color:#FF6600">网站信息配置</strong></td>
         </tr>
     </table>
	 <table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="forMy forMAargin" >
       <tr>
         <td width="100" align="right" class="forumRow">网站名称：</td>
         <td class="forumRow"><input name="title" type="text" value="<%=c_title%>" /></td>
       </tr>
       <tr>
         <td align="right" class="forumRow">网址：</td>
         <td class="forumRow"><input name="url" type="text" value="<%=c_url%>" /></td>
       </tr>
       <tr>
         <td align="right" class="forumRow">联系人：</td>
         <td class="forumRow"><input name="contact" type="text" value="<%=c_contact%>" /></td>
       </tr>
       <tr>
         <td align="right" class="forumRow">电话：</td>
         <td class="forumRow"><input name="tel" type="text" value="<%=c_tel%>" /></td>
       </tr>
       <tr>
         <td align="right" class="forumRow">传真：</td>
         <td class="forumRow"><input name="fax" type="text" value="<%=c_fax%>" /></td>
       </tr>
       <tr>
         <td align="right" class="forumRow">手机：</td>
         <td class="forumRow"><input name="mobile" type="text" value="<%=c_mobile%>" /></td>
       </tr>
       <tr>
         <td align="right" class="forumRow">Email：</td>
         <td class="forumRow"><input name="email" type="text" value="<%=c_email%>" /></td>
       </tr>
       <tr>
         <td align="right" class="forumRow">QQ：</td>
         <td class="forumRow"><input name="qq" type="text" value="<%=c_qq%>" /></td>
       </tr>
       <tr>
         <td align="right" class="forumRow">MSN：</td>
         <td class="forumRow"><input name="msn" type="text" value="<%=c_msn%>" size="40" /></td>
       </tr>
       <tr>
         <td align="right" class="forumRow">公司地址：</td>
         <td class="forumRow"><input name="address" type="text" value="<%=c_address%>" size="40" /></td>
       </tr>
       <tr>
         <td align="right" class="forumRow">网站关键字：</td>
         <td class="forumRow"><input name="keywords" type="text" value="<%=c_keywords%>" size="40"/></td>
       </tr>
       <tr>
         <td align="right" class="forumRow">关键字描述：</td>
         <td class="forumRow"><input name="description" type="text" value="<%=c_description%>" size="40" /></td>
       </tr>
       <tr>
         <td align="right" class="forumRow">版权信息：</td>
         <td class="forumRow"><textarea name="copyright" cols="55" rows="6"><%=c_copyright%></textarea></td>
       </tr>
       <tr>
         <td class="forumRow">&nbsp;</td>
         <td class="forumRow"><span class="submit">
           <input name="submit" type="submit" value="确认修改" />
         </span></td>
       </tr>
     </table></td>
   </tr>
   </form>
</table>
 <%if admin_config="ok" then response.write "<script>alert('网站基本信息设置成功!');window.location.href='admin_config.asp';</script>"  %>
</body>