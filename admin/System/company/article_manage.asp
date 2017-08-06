<!--配置文件-->
<!--#include file="article_config.asp"-->
<body>
<TABLE width="600" border="0" align="center" cellpadding="0" cellspacing="10" bgcolor="#FFFFFF">
	<TR><td class="forumRow">

<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="forMy forMAargin">
			<Form name="search" method="get" action="">
			<TR>    
				<TD align="center" class="forumRow mainTitle"><%=db_title%> 管理列表</td>
				<TD width="300" align="center" class="forumRow">
				  <table border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td><input name="keyword" type="text" id="keyword" style="font-size: 9pt" size="25" /></td>
                  <td>&nbsp;</td>
                  <td width="70" align="center"><input name="submit" type="submit" value="<%=db_title%>搜索" align="absMiddle" width="43" height="18" border=0 />
                    <input name="typeB_id" type="hidden" id="typeB_id" style="font-size: 9pt" value="<%=typeB_id%>" size="25" />
                    <input name="typeS_id" type="hidden" id="typeS_id" style="font-size: 9pt" value="<%=typeS_id%>" size="25" /></td>
                </tr>
              </table></td>
			</TR>
			</FORM>
	  </TABLE>


<%
r_ok=request.QueryString("ok")
r_hot=request.QueryString("hot")
r_id=request.QueryString("r_id")

if r_id<>"" and isnumeric(r_id) then
   set rs=server.createobject("adodb.recordset") 
       exec="select * from "&db_table&" where id="&r_id
       rs.open exec,conn,1,3 
	   if not rs.eof then
	      
		  if r_ok<>"" and int(r_ok)=1 then
		     rs("ok")=0
		  elseif r_ok<>"" and int(r_ok)=0 then
		     rs("ok")=1
		  end if
		  
	      if r_hot<>"" and int(r_hot)=1 then
		     rs("hot")=0
		  elseif r_hot<>"" and int(r_hot)=0 then
		     rs("hot")=1
		  end if

	   rs.update
	   end if
	   rs.close
   set rs=nothing
end if
 
'---------------------------------------------
	
	'########## 处理接收到搜索字符的情况 #############
       keyword=request("keyword")
	if keyword<>"" then
	   keyword_sql =" and title like '%"&request("keyword")&"%'"
	   keyword_sql2=" where title like '%"&request("keyword")&"%'"
	end if
	
	'########## 定义排序字符 #############
	dim order_sql
	    order_sql=" order by order_id asc,id desc"
    
	if typeS_id<>"" and isnumeric(typeS_id)then
	   exec="select * from "&db_table&" where typeS_id="&typeS_id&keyword_sql&order_sql
	elseif typeB_id<>"" and isnumeric(typeB_id)then 
	   exec="select * from "&db_table&" where typeB_id="&typeB_id&keyword_sql&order_sql
	else
	   exec="select * from "&db_table&keyword_sql2&order_sql
	end if

   '### 用于热门，审核 按钮链接,批量移动修改等
	Dim FullUrl
  	    FullUrl="?typeB_id="&typeB_id&"&typeS_id="&typeS_id&"&page="&page&"&keyword="&keyword&""

set rs=server.createobject("adodb.recordset")
	rs.open exec,conn,1,1 
	if rs.eof then
%>
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="forMy forMAargin">
<tr>
<td align="center" bgcolor="#DBF1F7" class="forumRow">
暂无 <%=db_title%> 相应内容!
</td>
</tr>
</table>	
	
<%else%>
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="forMy forMAargin">
					<tr>
						<td width="40" align="center" bgcolor="#DBF1F7" class="forumRow">ID</td>
					  <td width="420" bgcolor="#DBF1F7" class="forumRow">&nbsp;<%=db_title%>&nbsp;名称\标题</td>
<%
'设置是否显示审核
if db_showOK=1 then
%>
<td width="18" align="center" bgcolor="#DBF1F7" class="forumRow">热</td>
						<td width="18" align="center" bgcolor="#DBF1F7" class="forumRow">审</td>
<%end if%>

						<td width="80" align="center" bgcolor="#DBF1F7" class="forumRow">录入时间</td>
						<td width="110" align="center" bgcolor="#DBF1F7" class="forumRow">修改操作</td>
					</tr>
	  </table>
<form name="news_category" method="post">
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="forMy forMAargin">
<%
	rs.PageSize =18 '每页记录条数
	iCount=rs.RecordCount '记录总数
	iPageSize=rs.PageSize
	maxpage=rs.PageCount 
	page=request("page")
	
	if Not IsNumeric(page) or page="" then
	   page=1
	else
	   page=cint(page)
	end if
	
	if page<1 then
	   page=1
	elseif  page>maxpage then
	   page=maxpage
	end if
	
	rs.AbsolutePage=Page
	if page=maxpage then
	   x=iCount-(maxpage-1)*iPageSize
	else
	   x=iPageSize
	end if	
	
For i=1 To x
%>
					
					<tr <%=onTable%>>
					  <td width="40" align="center" class="forumRow2"><%=rs("id")%></td>
					  <td width="420" class="forumRow2">&nbsp;<a href="article_edit.asp?id=<%=rs("id")%>">
					    <%if keyword<>"" then%>
					  <%=replace(rs("title"),keyword,"<span style='color:#FF0000'>"&keyword&"</span>")%>
					  <%else%>
					  <%=rs("title")%>
					  <%end if%>
					  </a></td>



<%
'设置是否显示审核
if db_showOK=1 then
%>
<td width="18" align="center" class="forumRow2">

						<%if rs("hot")=1 then%>
						<a href="<%=FullUrl%>&hot=1&r_id=<%=rs("id")%>" class="yes">√</a>
						<%else%>
						<a href="<%=FullUrl%>&hot=0&r_id=<%=rs("id")%>" class="no">×</a>
					  <%end if%>						</td>
						<td width="18" align="center" class="forumRow2">
						<%if rs("ok")=1 then%>
						<a href="<%=FullUrl%>&ok=1&r_id=<%=rs("id")%>" class="yes">√</a>
						<%else%>
						<a href="<%=FullUrl%>&ok=0&r_id=<%=rs("id")%>" class="no">×</a>
						<%end if%></td>
<%end if%>

					  <td width="80" align="center" class="forumRow2"><a title="<%=rs("add_data")%>"><%=year(rs("add_data"))%>-<%=month(rs("add_data"))%>-<%=day(rs("add_data"))%></a></td>
					  <td width="110" align="center" class="forumRow2"><input name="delete" type="button" value="删除" onClick="javascript:if(confirm('确定要删除此<%=db_title%>吗？删除后不可恢复!')){window.location.href='article_manage.asp?act=del&id=<%=rs("id")%>';}else{history.go(0);}" <%if db_showDel=1 then%>style="display:none"<%end if%> />
					    <input type="button" name="update" value="修改" onClick="window.location.href='article_edit.asp?id=<%=rs("id")%>' "/></td>
					</tr>
				<%
					rs.movenext
					next
				%>
		</table>
			
</form>
				
		<%
			end if
		%>

<!--#include file="../../Class/articles_paging.asp"-->

</td></TR></TABLE>

<script>
function editItem(num){
 for (i=1;i<num;i++){
    Eitem=document.getElementById("edit_item_"+i).checked;
    if(Eitem==''){
    document.getElementById("edit_item_"+i).checked="checked";
	}else{
    document.getElementById("edit_item_"+i).checked='';
	}
 }
}
</script>



<%
if errStr="ok" then
   if edits="del" then
      errStr="成功删除!"
   elseif edits="check" then
      errStr="成功审核!"
   elseif edits="not_check" then
      errStr="成功取消审核!"
   elseif edits="move" then
      errStr="成功移动!"
   end if
end if

if errStr<>"" then
   response.Write("<script>alert('"&errStr&"');window.location.href='"&FullUrl&"';</script>")
end if
%>

</body>
</html>
<%
if request("act")="del" and db_showDel<>1 then
   conn.execute("delete * from "&db_table&" where id="&id)
   response.Write "<script>alert('成功刪除"&db_title&"！');window.location.href='"&FullUrl&"';</script>" 
end if
%>