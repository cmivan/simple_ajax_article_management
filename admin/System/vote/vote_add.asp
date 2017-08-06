<br />
<!--上面的articles_head.asp文件中的样式已经足够，不过为了方便在dreamwear中操纵-->
<link href="_kmsysfile/images/style.css" rel="stylesheet" type="text/css" />
<link href="../_kmsysfile/images/style.css" rel="stylesheet" type="text/css" />
<link href="../../_kmsysfile/images/style.css" rel="stylesheet" type="text/css" />
<table border="0" align="center" cellpadding="0" cellspacing="10" bgcolor="#FFFFFF">
  <tr>
    <td width="600" class="forumRow">
<!--#include file="VOTE_TOP.ASP"-->

<%
if request("actions")="1" then
   title=trim(request("titles"))
   if request("current")="1" then
      cu=true
   else
      cu=false
   end if
   if request("choice")="1" then
      choi=false
   else
      choi=true
   end if
   if title="" or len(title)>200 then
      response.write "<script language='javascript'>alert('主题内容不能为空或大于指定字数!');history.go(-1);</script>"
   else
      set rs=server.CreateObject("adodb.recordset")
	  sql="select * from vote_title where id=null"
	  rs.open sql,connstr,1,3
	  rs.addnew
	  rs("title")=title
	  rs("current")=cu
	  rs("choice")=choi
	  rs.update
	  response.write "<script language='javascript'>alert('主题添加成功！请进入编辑页面添加选项');window.location.href='VOTE_ADMIN.ASP';</script>"
   end if
end if
%>

<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="forMy">
<form name="form1" method="post" action="">
    <tr>
      <td height="14" colspan="2" align="center" class="forumRaw"><span style="font-weight: bold">添加投票标题</span></td>
    </tr>
	
    <tr>
      <td width="26%" height="14" class="forumRow"><div align="center">标题内容</div></td>
      <td width="74%" class="forumRow"><input name="titles" type="text" id="titles" value="" size="30" />
      小于200字符</td>
    </tr>
    <tr>
      <td width="26%" height="7" class="forumRow"><div align="center">单选/多选</div></td>
      <td class="forumRow"><input name="choice" type="radio" value="1" checked>
        单选
        <input type="radio" name="choice" value="2">
      多选</td>
    </tr>
    <tr>
      <td class="forumRow"><div align="center">&nbsp;</div></td>
      <td class="forumRow"><input name="actions" type="hidden" id="actions" value="1" />
        <input type="submit" name="Submit" value="提交保存" />
&nbsp;&nbsp;
<input type="reset" name="Submit2" value="清除重来" /></td>
    </tr>
</form>
</table>



	</td>
  </tr>
  
</table>