<br />
<table border="0" align=center cellpadding=0 cellspacing=10 bgcolor="#FFFFFF">
  <tr>
    <td width="600" class="forumRow">

<!--#include file="VOTE_TOP.ASP"-->

<%id=trim(request("id"))
if request("actions")="1" then
     if request("choice")="1" then
	    choi=false
     else
	    choi=true
     end if
	 if request("current")="1" then
	    cu=true
     else
	    cu=false
     end if
     topic=trim(request("topic"))
     if topic="" or len(topic)>200 then
	    response.write "<script language='javascript'>alert('主题内容不能为空或大于指定字数!');history.go(-1);</script>"
     else
	    set rs=server.CreateObject("adodb.recordset")
		sql="select * from vote_title where id="&id
		rs.open sql,connstr,1,3
		rs("title")=topic
		rs("current")=cu
		rs("choice")=choi
		rs.update
		rs.close
		set rs=nothing
		response.write "<script language='javascript'>alert('主题修改成功!');history.go(-1);</script>"

     end if
else
set rs=server.CreateObject("adodb.recordset")
sql="select * from vote_title where id="&request("id")
rs.open sql,connstr
%>
<form name="form1" method="post" action="">
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="forMy">
    <tr>
      <td width="24%" height="45" class="forumRow"><div align="center">主题内容</div></td>
      <td width="76%" valign="bottom" class="forumRow"><textarea name="topic" cols="40" rows="4" id="topic"><%=rs("title")%>
</textarea>
      小于 200字</td>
    </tr>
    <tr>
      <td valign="top" class="forumRow"><div align="center">多选/单选</div></td>
      <td valign="bottom" class="forumRow">
	  <%if rs("choice")=true then%>
	  <input name="choice" type="radio" value="1" >
        单选
          <input type="radio" name="choice" value="2" checked>
      多选
	  <%else%>
	  <input name="choice" type="radio" value="1" checked>
        单选
          <input type="radio" name="choice" value="2" >
      多选
	  <%end if%></td>
    </tr>
    <tr>
      <td valign="bottom" class="forumRow">&nbsp;</td>
      <td valign="bottom" class="forumRow"><input type="submit" name="Submit" value="提交保存" />
&nbsp;&nbsp;
<input type="reset" name="Submit2" value="清除重来" />
<input name="actions" type="hidden" id="actions" value="1" /></td>
    </tr>
  </table>
 
</form>
<%end if%>
	</td>
  </tr>
  
</table>