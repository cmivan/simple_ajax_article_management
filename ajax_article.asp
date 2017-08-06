<!--#include file="ajax_top.asp"-->
<%
dim db_table
    db_table="article"
%>
<div class="mainWidth" id="ajax_main">
<div class="main_left">
<div class="main_left_box">
<div>
<input name="keyword" onkeyup='gAjax("<%=db_table%>_list.asp?keyword=","list",this.value);' type="text" class="search_bg" id="keyword" value="<%=keyword%>" size="10" maxlength="10"/>
</div>
<div class="leftNav">
<%
set rs=conn.execute("select * from "&db_table&"_type where type_id=0 order by order_id asc,id desc")
    do while not rs.eof
%>
<a href='javascript:gAjax("<%=db_table%>_list.asp?typeB_id=","list",<%=rs("id")%>);'><%=rs("title")%></a>
<%
    rs.movenext
    loop
set rs=nothing
%>
</div>
<div class="clear"></div>
</div>
<div class="clear_padding"></div>
<!--
<div class="main_left_box" style=" padding:0">
<a href="#"><img src="images/logo/1.jpg" width="110" border="0" /></a>
<a href="#"><img src="images/logo/2.jpg" width="110" border="0" /></a>
<a href="#"><img src="images/logo/3.jpg" width="110" border="0" /></a>
<a href="#"><img src="images/logo/4.jpg" width="110" border="0" /></a>
<a href="#"><img src="images/logo/5.jpg" width="110" border="0" /></a>
</div>
<div class="clear_padding"></div>
<div class="main_left_box">&nbsp;</div>
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
<div class="mainWidth">&nbsp;</div>

<script>
gAjax("<%=db_table%>_list.asp?typeB_id=","list","");  //初始化页面内容
</script>
</body>
</html>

