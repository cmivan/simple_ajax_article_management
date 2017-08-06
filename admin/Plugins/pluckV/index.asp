<!--#include file="../../Bin/Charset.asp"-->
<%Rpath="../../"%>
<!--#include file="../../Bin/adminConn.asp"-->
<!--#include file="../../Bin/Isadmin.asp"-->
<!--#include file="../../Bin/head.asp"--> 

<%response.Buffer=true%>
<body>
<div style="margin:auto; width:715px;">
<div id="left">
  <div class="left_body forumRow">
 &nbsp;&nbsp;+ <a href="?rel=add">添加采集规则</a>    + <a href="?">管理采集规则</a>
  </div>
</div>
	<div id="right">
	  <div class="right_body" style="overflow:hidden;">
	    <!--#include file="pluck.asp"-->
	  </div>
	</div>
	<div class="clear"></div>
</div>
</body>