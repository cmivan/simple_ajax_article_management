<!--#include file="ajax_top.asp"-->
<%
dim db_table
    db_table="article"
response.Redirect("ajax_article.asp")
%>

<div class="mainWidth" id="ajax_main">
<div class="main_left">
<div class="main_left_box">
<div>
<input name="keyword" onkeyup='gAjax("list.asp?keyword=","list",this.value);' type="text" class="search_bg" id="keyword" value="<%=keyword%>" size="10" maxlength="10"/>
</div>
<div class="leftNav">
<%
set rs=conn.execute("select * from article_type where type_id=0 order by order_id asc,id desc")
    do while not rs.eof
%>
<a href='javascript:gAjax("list.asp?typeB_id=","list",<%=rs("id")%>);'><%=rs("title")%></a>
<%
    rs.movenext
    loop
set rs=nothing
%>
</div>
<div class="clear"></div>
</div>
<!--
<div class="clear_padding"></div>
<div class="main_left_box">&nbsp;</div>
-->
</div>

<div class="main_right">
<div class="main_right_border">
<div id="ajax_list"></div>
</div></div>
<div class="clear"></div>
</div>


<script>
gAjax("list.asp?typeB_id=","list","");
//初始化页面内容
</script>
</body>
</html>

