<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="forMy forMAargin">
<TR>
<TD colspan="3" align="left" class="forumRow">
 &nbsp;分类<img src="<%=Rpath%>Images/Ico/type_ico2.gif" width="12" height="12" /> | <a href="?"><strong>&nbsp;全部&nbsp;</strong></a>
 <%	
'### 搜索关键词 ####
    keyword=request("keyword")
set rs=server.createobject("adodb.recordset") 
	exec="select * from "&db_table&"_type where type_id=0 order by order_id asc" 
	rs.open exec,conn,1,1 
	do while not rs.eof
	'### 用于样式 ####
	if cstr(rs("id"))=cstr(typeB_id) then
	   styleB=" style=""font-weight:bold; color:#FF0000;"""
	   else
	   styleB=""
	end if
%>|<a href="?typeB_id=<%=rs("id")%>&keyword=<%=keyword%>" <%=styleB%>>&nbsp;&nbsp;<%=rs("title")%>&nbsp;&nbsp;</a><%
	rs.movenext
	loop
    rs.close
set rs=nothing
%></td>
</TR>

<%if typeB_id<>"" and isnumeric(typeB_id) then%>
<%	
set rs=server.createobject("adodb.recordset") 
	exec="select * from "&db_table&"_type where type_id="&typeB_id&" order by order_id asc" 
	rs.open exec,conn,1,1

if not rs.eof then
response.write "<TR><TD colspan=""3"" align=""left"" class=""forumRow"" style=""padding-left:50px;"">"
	do while not rs.eof
	'### 用于样式 ####
	if cstr(rs("id"))=cstr(typeS_id) then
	   styleB=" style=""color:#FF0000;"""
	   else
	   styleB=""
	end if
	
%>
&nbsp;- <a href="?typeB_id=<%=typeB_id%>&typeS_id=<%=rs("id")%>&keyword=<%=keyword%>" <%=styleB%>><%=rs("title")%></a>&nbsp;<%
	rs.movenext
	loop
response.write "</td></TR>"
end if
    rs.close
set rs=nothing
%>

<%end if%>
</TABLE>