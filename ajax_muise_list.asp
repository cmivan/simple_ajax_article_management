<!--#include file="data/conn.asp"-->
<%
dim db_table
    db_table="muise"
%>
<%
dim typeB_id,keyword,rsSql
    typeB_id=request.QueryString("typeB_id")
    keyword =request("keyword")
    keyword =replace(keyword,"'","")
    if keyword<>"" then
       keywordStr1=" and title like '%"&keyword&"%'"
       keywordStr2=" where title like '%"&keyword&"%'"
    end if

    if typeB_id<>"" and isnumeric(typeB_id) then
       rsSql="select * from "&db_table&" where typeB_id="&typeB_id&keywordStr1&" order by order_id asc,id desc"
    else
       rsSql="select * from "&db_table&keywordStr2&" order by order_id asc,id desc"
    end if

set rs=server.createobject("adodb.recordset")
	rs.open rsSql,conn,1,1 
    if rs.eof then
%>
<div align="center" class="page_none">暂无相关信息!</div>
<%
    else


	rs.PageSize =10 '每页记录条数
	iCount   =rs.RecordCount '记录总数
	iPageSize=rs.PageSize
	maxpage  =rs.PageCount 
	page     =request("page")
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
<div <%=onTable%> class="list_title"><a href='javascript:gAjax("<%=db_table%>_view.asp?id=","list",<%=rs("id")%>);'><%=key(rs("title"),keyword)%></a>
<span class="info">(Time:<%=rs("add_data")%>&nbsp;&nbsp;Hit:<%=rs("hits")%>)</span>
</div>
<div class="list_content"><%=cutStr(rs("content"),150)%></div>
<%
	rs.movenext
next

'分页子程序
'生成上一页下一页链接
    Dim query, a, x, temp
    action = "http://" & Request.ServerVariables("HTTP_HOST") & Request.ServerVariables("SCRIPT_NAME")
    query = Split(Request.ServerVariables("QUERY_STRING"), "&")
   
    For Each x In query
        a = Split(x, "=")
        If StrComp(a(0), "page", vbTextCompare) <> 0 Then
            temp = temp & a(0) & "=" & a(1) & "&"
        End If
    Next
%>
<div class="list_paging">
<form method=get onsubmit=""document.location = '" & action & "?" & temp & "page='+ this.page.value;return false;"">
<%if page<=1 then%>
首页 上一页
<%else%>   
<a class="liti" href='javascript:gAjax("<%=db_table%>_list.asp?<%=temp%>page=","list",1);'>首页</A>
<a class="liti" href='javascript:gAjax("<%=db_table%>_list.asp?<%=temp%>page=","list",<%=Page-1%>);'>上一页</A>
<%end if%>
	
<%if page>=maxpage then%>
下一页 尾页
<%else%>   
<a class="liti" href='javascript:gAjax("<%=db_table%>_list.asp?<%=temp%>page=","list",<%=Page+1%>);'>下一页</A>
<a class="liti" href='javascript:gAjax("<%=db_table%>_list.asp?<%=temp%>page=","list",<%=maxpage%>);'>尾页</A>    
<%end if%>
页次：<%=page%>/<%=maxpage%>页
共 <%=iCount%> 条记录
</form>   
</div>
<%
    end if
set rs=nothing
%>