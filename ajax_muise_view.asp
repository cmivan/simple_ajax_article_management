<!--#include file="data/conn.asp"-->
<%
dim db_table
    db_table="muise"
%>
<%
dim id,rsSql
    id=request.QueryString("id")
    if id<>"" and isnumeric(id) then
       rsSql="select top 1 * from "&db_table&" where id="&id&" order by order_id asc,id desc"
    else
       rsSql="select top 1 * from "&db_table&" order by order_id asc,id desc"
    end if
set rs=conn.execute(rsSql)
    if rs.eof then
%>
<div align="center" class="page_none">暂无相关信息!</div>
<%else%>
<table width="100%" border="0" cellpadding="2" cellspacing="1">
<tr>
<td align="center" bgcolor="#C3C3C3" style="font-size:14px; color:#666666; padding:5px;">
<%=rs("title")%><br />
<span class="view_info">Time:<%=rs("add_data")%>&nbsp;&nbsp;Hit:<%=get_one("article",rs("id"),"hit")%><%=rs("hits")%></span>
</td>
</tr>
<tr>
<td bgcolor="#D2D2D2">
<table width="100%" border="0" cellpadding="0" cellspacing="8">
  <tr>
    <td><%=rs("content")%></td>
    </tr>
</table>

</td>
</tr>
<tr>
  <td bgcolor="#C3C3C3">&nbsp;&nbsp;  [<a href="#">- top -</a>]</td>
</tr>
</table>
<%
    end if
set rs=nothing
%>