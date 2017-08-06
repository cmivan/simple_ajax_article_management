<!--#include file="email_config.asp"-->

<% if request("del")<>"" then 
   conn.Execute("delete * from email where id="&request("del"))
   end if %>
<SCRIPT LANGUAGE="JavaScript">
<!--//
function check()
{
   if (isNaN(go2to.page.value))
		alert("请正确填写转到页数！");
   else if (go2to.page.value=="") 
	     {
		alert("请输入转到页数！");
		 }
   else
		go2to.submit();
}
//-->
</SCRIPT>
<% 
     set rs=server.createobject("adodb.recordset") 
     if not isempty(request("page")) then   
     pagecount=cint(request("page"))   
     else   
     pagecount=1   
     end if

     if key="" then
     sql="select * from email order by id desc" 
     else
     sql="select * from email where dateandtime like '%"&key&"%' order by id desc" 
     end if

	 rs.open sql,conn,1,1   
     if rs.eof and rs.bof then   
     response.write"<BR>" 
	 response.write"=== 邮件列表为空，请在前台添加至少一个以上邮件地址，才可以管理！ ==="
     response.write"<BR><BR>"

	 response.end  
	 end if
     rs.pagesize=20
     if pagecount>rs.pagecount or pagecount<=0 then              
     pagecount=1              
     end if              
     rs.AbsolutePage=pagecount %>
<br />

<table border="0" align="center" cellpadding="0" cellspacing="10" bgcolor="#FFFFFF">
  <tr> 
    <td width="600" align="center" valign="top" class="forumRow">
<!--#include file="email_top.asp"-->
	
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="forMy forMAargin">
        <tr> 
          <% if rs.pagecount=1 then %>
          <td height="25" colspan="4"  class="forumRow"> <font class=ver>共有[<font color="#ff0000"><%=rs.recordcount%></font>]E-mail地址 以下是[<font color="red">1～<%=rs.recordcount%></font>]个&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="EMAIL_AddUser.asp">添加邮件地址</a></font></td>
        </tr>

        <%else%>
        <tr> 
          <td height="25" colspan="4" class="forumRow" ><font class=ver> 
            <% 
            page_start=(pagecount-1)*rs.pagesize
            if pagecount=1 then page_start=1
		    page_end=rs.pagesize*pagecount
		    if pagecount*rs.pagesize=>rs.recordcount then page_end=rs.recordcount end if%>
          共有[<font color="#ff0000"><%=rs.recordcount%></font>]E-mail地址 以下是[<font color="red"><%=page_start%>～<%=page_end%></font>]</font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="adduser.asp">添加邮件地址</a></td>
        </tr>
        <tr> 
          <td height="25" colspan="4" class="forumRow" > <% response.write"<form name=go2to form method=Post action='viewuser.asp?key="+key+"'>"
		     response.write "<font color=ffffff>◆&nbsp;</font>"                                              
		     if pagecount=1 then                                                        
			 response.write "<font class=ver>首页 上一页</font>&nbsp;"
			 else                                                        
             response.write "<a href=viewuser.asp?page=1&key="+key+"><font class=ver>首页</font></a>&nbsp;" 
	         response.write "<a href=viewuser.asp?page="+cstr(pagecount-1)+"&key="+key+"><font color='0000BE'>上一页</font></a>&nbsp;"
			 end if                                             
             if rs.PageCount-pagecount<1 then                                                        
             response.write "<font class=ver>下一页 尾页</font>"                                                    
			 else                                                        
             response.write "<a href=viewuser.asp?page="+cstr(pagecount+1)+"&key="+key+"><font class=ver>下一页</font></a>&nbsp;"
			 response.write "<a href=viewuser.asp?page="+cstr(rs.PageCount)+"&key="+key+"><font class=ver>尾页</font></a>"           
             end if 
			 response.write "<font class=ver>&nbsp;页次:<font class=ver>"&pagecount&"</font>/"&rs.pagecount&"页</font>" 
			response.write "<font class=ver> 转到第<input type='text' name='page' size=2 maxLength=3 style='font-size: 9pt; color:#00006A; position: relative; height: 18' value="&PageCount&">页</font>&nbsp;&nbsp;&nbsp;&nbsp;"                               
			response.write "<input class=button1 type='button' value='确 定' onclick=check() style='font-family: 宋体; font-size: 9pt; color: #000073; position: relative; height: 19'>" %> </td>
        </tr>
        <%end if%>
      </table>
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="forMy">
        <tr bgcolor="#C0C0C0"> 
          <td width="9%" align="center" class=forumRaw> 
         <font color="#000000">ID</font></td>
          <td width="31%" align="center" class="forumRaw"> 
          <font color="#000000">邮件地址</font></td>
          <td colspan="3" align="center" class=forumRaw> 
          <font color="#000000">用户名</font></td>
          <td  width="16%" align="center" class="forumRaw"> 
         <font color="#000000">- 删除 - </font>          </td>
        </tr>
        <% do while not rs.eof %>
        <tr bgcolor="#E3E3E3"> 
          <td width="9%"  height="25" align="center" bgcolor="#ECF5FF" class=forumRow> 
          <font color=000000><%=rs("id")%></font></td>
          <td align="center" bgcolor="#ECF5FF"  class=forumRow> 
          <font color="000000"> &nbsp;&nbsp;&nbsp;&nbsp;<a href="EMAIL_Send.asp?email=<%=rs("email")%>"><%=rs("email")%></a></font></td>
          <td colspan="3" align="center" bgcolor="#ECF5FF"  class=forumRow> 
          <font color="000000"> &nbsp;&nbsp;&nbsp;&nbsp;<%=rs("dateandtime")%></font></td>
          <td width="16%" align="center" bgcolor="#ECF5FF" class="forumRow" > 
          <font color="#000000">[<a href="?del=<%=rs("id")%>"><font class=ver>删除</font></a>]</font>          </td>
        </tr>
        <% i=i+1                                                                                                  
          rs.movenext                                                                                                  
          if i>=rs.PageSize then exit do 
		  loop                                                                    
          rs.close                                                                                                
          set rs=nothing                                                                                                
          conn.close                                                                                                
          set conn=nothing 
'end if
%>
      </table>
    <br> </td>
  </tr>
</table>
