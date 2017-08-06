<%
'=++++++++++++++++++++++++++++++++++++++++=
'on error resume next
dim conns,connstr,mdbs
'-----------  Access数据库连接字符串 -----------
    mdbs="db.mdb"           '数据库文件目录
    Connstr="DRIVER=Microsoft Access Driver (*.mdb);DBQ="+server.mappath(mdbs)
'-----------  SQL数据库连接字符串    -----------
'	Connstr="driver={SQL Server};server=192.168.0.1,7788;database=KM_09_12_20;UID=sa;PWD=sa"
'---------------------------------------------
Set conn=Server.CreateObject("ADODB.Connection") 
    conn.Open connstr
'------------------------------------------==
 If Err Then
    Err.Clear
    Set Conns = Nothing
    Response.Write "数据库连接有误!请检查相应文件!"
    Response.End
 End If
'=++++++++++++++++++++++++++++++++++++++++=



if request.Form("save")="ok" then
'////////////////////////////////
set rs=conn.execute("select * from div")
    do while not rs.eof
    moveDivs=request.Form("move_"&rs("id"))
	moveDiv=split(moveDivs,",")
	if ubound(moveDiv)=1 then
	   moveTop =moveDiv(0)
	   moveLeft=moveDiv(1)
	'----------------------
   set rsDiv=server.createobject("adodb.recordset") 
       exec="select * from div where id="&rs("id")                     '判断，添加数据
	   rsDiv.open exec,conn,1,3
	   if not rsDiv.eof then
	      rsDiv("Top")    =moveTop
	      rsDiv("Left")   =moveLeft
	      rsDiv.update
	   end if
	   rsDiv.close
   set rsDiv=nothing
   '----------------------
    end if
    rs.movenext
    loop
set rs=nothing
'////////////////////////////////
end if
%>
<style type="text/css">
<!--
body,td,th,form {
	font-size: 12px;
	margin:0;
}
a:link {
	text-decoration: none;
	color: #000000;
}
a:visited {
	text-decoration: none;
	color: #000000;
}
a:hover {
	text-decoration: none;
	color: #666666;
}
a:active {
	text-decoration: none;
	color: #666666;
}
div{
border:#CCCCCC 1px solid;
background-color:#999999;}
#Box{
background-color:#0099CC;
border:#999999 1px dotted}
-->
</style>

<SCRIPT LANGUAGE="JavaScript">
<!--
var isCatchDiv = false;
var dragX = 0;
var dragY = 0;
var divMove;
//function loadDiv(){
//divMove.style.top=document.body.scrollTop;
//divMove.style.left=document.body.scrollLeft;
//} 
function hideDiv(){
divMove.style.visibility = "hidden";
}
function showDiv(){
//loadDiv();
//divMove.style.visibility = "visible";
}
function catchDiv(obj){
divMove = obj;
BoxDiv  = document.getElementById("Box");
BoxDiv.style.display='block';

isCatchDiv = true;
var x=event.x+document.body.scrollLeft;
var y=event.y+document.body.scrollTop;
dragX=x-divMove.style.pixelLeft;
dragY=y-divMove.style.pixelTop;
BoxDiv.setCapture();
divMove.style.filter='alpha(opacity=50)';
BoxDiv.style.filter='alpha(opacity=50)';
BoxDiv.style.left=divMove.style.pixelLeft;
BoxDiv.style.top=divMove.style.pixelTop;
BoxDiv.style.height=divMove.style.height;
BoxDiv.style.width=divMove.style.width;
document.onmousemove = moveDiv;
}

function releaseDiv(){
isCatchDiv = false;
BoxDiv.releaseCapture();

divMove.style.left=BoxDiv.style.left;
divMove.style.top =BoxDiv.style.top;
divMove.style.filter='alpha(opacity=100)';
BoxDiv.style.display='none';

//-------------------------
document.getElementById("move_"+divMove.id).value=BoxDiv.style.top+","+BoxDiv.style.left;
document.onmousemove = null;
form1.submit();
}
function moveDiv(){
if(isCatchDiv){
BoxDiv.style.left = event.x+document.body.scrollLeft-dragX;
BoxDiv.style.top  = event.y+document.body.scrollTop-dragY;
}
}
window.onload = showDiv;
//-->
</SCRIPT>

<div style="border:#333333 1px solid;">
<form id="form1" name="form1" method="post" action="">
<%
set rs=conn.execute("select * from div")
    do while not rs.eof
%>
<input onchange="submit();" name="move_<%=rs("id")%>" id="move_<%=rs("tip")%>" type="text" value="<%=rs("top")%>,<%=rs("left")%>" style="width:100px;" size="20" />
<%
    rs.movenext
    loop
set rs=nothing
%>
<input type="hidden" name="save" value="ok" />
<input type="submit" name="Submit" value=" 保存 " />
</form>
</div>

<div style="margin:100px;position:relative;">
<div id="Box" style="position:absolute;cursor:move; z-index:9999; display:none" onMouseUp="releaseDiv()"></div>
<%
set rs=conn.execute("select * from div")
    do while not rs.eof
%>
<div id="<%=rs("tip")%>" onMouseDown="catchDiv(this);" style="position:absolute;cursor:move;width:<%=rs("width")%>;height:<%=rs("height")%>;top:<%=rs("top")%>;left:<%=rs("left")%>;"><%=rs("title")%></div>
<%
    rs.movenext
    loop
set rs=nothing
%>


</div>

