<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title><%=c_title%></title>
<link href="<%=Rpath%>_kmsysfile/images/style.css" rel="stylesheet" type="text/css" />
</head>
<script language="javascript">
//文件上传，弹出的窗口 js
function showUploadDialog(s_Type, s_Link, s_Thumbnail){
var arr = showModalDialog("../_kmsysfile/fckupLoad/dialog/i_upload.htm?style=popup&type="+s_Type+"&link="+s_Link+"&thumbnail="+s_Thumbnail, window, "dialogWidth:0px;dialogHeight:0px;help:no;scroll:no;status:no");}
</script>



<script>
<%
'### 管理页面，鼠标经过表格样式
 onTable=" style=""background-color:#EFEFE0;"" onmouseover=""cursor_('on',this);"" onmouseout=""cursor_('',this);"""
%>

//定义鼠标移过样式
var mcolor;
//用于记忆原来的颜色
function cursor_(ctype,objs){
    if (ctype=="on"){
	   mcolor=objs.style.backgroundColor;
	   objs.style.backgroundColor="#ffffff";
	}else{
	   objs.style.backgroundColor=mcolor;
	}
}


//防止刷新的
function document.onkeydown()       
{ 
if (event.keyCode==116){
window.parent.frames["main"].location.reload();
event.keyCode = 0;
return false;
}}
</script>
