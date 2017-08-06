<%
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
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="forMy">
<form method=get onsubmit=""document.location = '" & action & "?" & temp & "page='+ this.page.value;return false;"">
<TR>
<TD align=center class="forumRow">&nbsp;&nbsp;
<%if page<=1 then%>
首页 上一页
<%else%>   
<A class="liti" HREF="?<%=temp%>page=1">首页</A>
<A class="liti" HREF="?<%=temp%>page=<%=Page-1%>">上一页</A>
<%end if%>
	
<%if page>=maxpage then%>
下一页 尾页
<%else%>   
<A class="liti" HREF="?<%=temp%>page=<%=Page+1%>">下一页</A>
<A class="liti" HREF="?<%=temp%>page=<%=maxpage%>">尾页</A>    
<%end if%>

	
	
页次：<%=page%>/<%=maxpage%>页
共 <%=iCount%> 条记录
转 
<INPUT CLASS=wenbenkuang TYEP=TEXT NAME=page SIZE=2 Maxlength=5 VALUE="<%=page%>">页<INPUT CLASS=go-wenbenkuang type=submit value=转到></TD>   
</TR></form>   
</table>
