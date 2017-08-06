<!--#include file="ajax_top.asp"-->
<%
dim db_table
    db_table="muise"
    db_num  =4*3
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
</div>

<div class="main_right">
<div id="mk_show_box">
<%
set rs=server.CreateObject("adodb.recordset")
	sql="select top "&db_num&" * from photo"
	rs.open sql,conn,1,1
for i=1 to db_num
if rs.eof then
%>
<div id="mk_pic_item_<%=i%>">
<a style=" background:none;width:128px; height:152px; line-height:152px;"></a>
</div>
<%
else
%>
<div id="mk_pic_item_<%=i%>">
<a href="javascript:isshow('<%=i%>');"><img src="<%=rs("small_pic")%>" width="128" height="128" border="0" /><br><span><%=cutStr(rs("title"),8)%></span></a>
</div>
<%
    rs.movenext
end if
next
	  rs.close  
  set rs=nothing
%>

<div id="mk_pic_item_border_1">&nbsp;</div>
<div id="mk_pic_item_border_2">&nbsp;</div>
<div id="mk_pic_show"></div>
<div id="mk_pic_info"></div>
<div id="mk_flash_box">
  <object id="flash" classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,19,0" width="460" height="630">
    <param name="movie" value="images/change.swf" />
    <param name="quality" value="high" />
	<param name="wmode" value="transparent">
    <embed src="images/change.swf" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" width="460" height="630"></embed>
  </object>
</div>
</div>
<div class="clear"></div>
</div>
<div class="clear"></div>
</div>
<div class="mainWidth">&nbsp;</div>
<script language="javascript" src="photo_js.asp?show_Num=<%=db_num%>"></script>
<script language="javascript">isshow('2');  //初始化显示 </script>
</body>
</html>


