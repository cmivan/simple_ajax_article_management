<!--配置文件-->
<!--#include file="article_config.asp"-->

<%
'///////////  处理提交数据部分 ////////// 
   edit_id   =request("id")
if edit_id="" or isnumeric(edit_id)=false then
   editStr="添加"
   else
   editStr="修改"
end if
   
if request.Form("edit")="ok" then
'///////////  写入数据部分 //////////
   set rs=server.createobject("adodb.recordset") 
    if edit_id="" or isnumeric(edit_id)=false then
       exec="select * from "&db_table                       '判断，添加数据
	   rs.open exec,sysconn,1,3
       rs.addnew
	else
	   exec="select * from "&db_table&" where id="&edit_id  '判断，修改数据
       rs.open exec,sysconn,1,3
	end if

	if edit_id<>"" and isnumeric(edit_id) and rs.eof then
	   response.Write("写入数据失败!")
	else

'——————接收数据
    title     =request.Form("title")
	toUrl     =request.Form("toUrl")
	note      =request.Form("note")
	add_data  =request.Form("add_data")
 if isdate(add_data)=false then add_data=now()
'——————排序
 if order_id="" or isnumeric(order_id)=false then order_id=0


	rs("title")    =title
	rs("toUrl") =toUrl
	rs("note")  =note
	rs("add_data")=add_data
    rs("order_id") =order_id
	

    response.Write("<script>alert('"&editStr&"操作成功!');window.location.href='article_manage.asp?typdB_id="&typdB_id&"&typdS_id="&typdS_id&"';</script>")
	end if
	rs.update
	rs.close
set rs=nothing

end if





'#################  读取数据部分 #################

   id=request.QueryString("id") 
if id<>"" and isnumeric(id) then
  '/// 当前状态:编辑
   edit_stat="edit"
   set rs=server.createobject("adodb.recordset") 
       exec="select * from "&db_table&" where id="&id
       rs.open exec,sysconn,1,1 
	   if not rs.eof then
	
	      title     =rs("title")
		  toUrl=rs("toUrl")
		  note =rs("note")
		  add_data=rs("add_data")
          order_id  =rs("order_id")
		  if order_id="" or isnumeric(order_id)=false then order_id=0

	   end if
	   rs.close
   set rs=nothing
   
else
  '/// 当前状态:添加
   edit_stat="add"
end if


  '/// 处理添加时间
  if isdate(add_data)=false then add_data=now()
%>


<body>
<TABLE width="660" border="0" align="center" cellpadding="0" cellspacing="10" bgcolor="#FFFFFF">
  <TR><td class="forumRow">
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="forMy forMAargin">
        <tr>
          <td align="center" class="forumRow"><strong>&nbsp;&nbsp;<%=db_title%> <%=editStr%></strong></td>
        </tr>
      </table>

<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="forMy">
        <form name="article_update" method="post" action="">
          <tr>
            <td width="13%" align="right" class="forumRow">名称：</td>
            <td width="87%" class="forumRow"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td width="36%"><input name="title" type="text" value="<%=title%>" size="40" /></td>
<td width="64%">&nbsp;</td>
</tr>
</table></td>
</tr>
		  
		  
		  
		  
<!--          <tr>
            <td align="right" class="forumRow">图标：</td>
            <td class="forumRow"><input name="small_pic" type="text" id="small_pic" value="<%=small_pic%>" size="45" maxlength="255" />
              <input type=button value="上传图片" onClick="showUploadDialog('image', 'article_update.small_pic', '')"></td>
          </tr>-->
		  
          <tr>
            <td align="right" class="forumRow"><%=db_title%>地址：</td>
            <td class="forumRow"><input name="toUrl" type="text" id="toUrl" value="<%=toUrl%>" size="45" maxlength="255" /> 
            (<span class="red">链接地址如: http://www.baidu.com</span>) </td>
          </tr>
		  
          <tr>
            <td align="right" valign="top" class="forumRow" style="padding-top:4px;">简要：</td>
            <td class="forumRow"><textarea name="note" cols="65" rows="4" id="note"><%=note%></textarea></td>
          </tr>
		  
          <tr style="display:none">
            <td align="right" valign="top" class="forumRow">排序：</td>
            <td class="forumRow"><input name="order_id" type="text" id="order_id" value="<%=order_id%>" size="10" /></td>
          </tr>
          <tr>
            <td align="right" valign="top" class="forumRow"><p class="submit">
                <input name="id" type="hidden" value="<%=id%>" />
				<input name="edit" type="hidden" value="ok" />
            </p></td>
            <td class="forumRow"><span class="submit">
              <input name="submit" type="submit" value="确认<%=editStr%>" />
            </span></td>
          </tr>
        </form>
      </table>
</td></TR></TABLE>
</body>
</html>