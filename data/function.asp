<%
'### 鼠标经过表格样式
    onTable=" onmouseover=""cursor_('on',this);"" onmouseout=""cursor_('',this);"""

function key(str,keyword)
    dim keywords
    keywords="<span style='color:#ff0000'>"&keyword&"</span>"
    key=replace(str,keyword,keywords)
end function

'***********************************
'    通用调用单条信息
'***********************************
function get_one(dbTable,gid,gfield)
     'gid未指定值的情况下，自动接收参数
	  if gid=0 then gid=request.QueryString("id")
  set get_one_rs=server.CreateObject("adodb.recordset")
      if gid="" or isnumeric(gid)=false then
	     get_one_sql="select top 1 * from "&dbTable&" order by order_id asc,id desc"
	  else
	  	 get_one_sql="select * from "&dbTable&" where id="&int(gid)
	  end if
	  get_one_rs.open get_one_sql,conn,1,3
	  if not get_one_rs.eof then
	     if gfield="hit" then
		    if get_one_rs("hits")="" or get_one_rs("hits")=null or isnumeric(get_one_rs("hits"))=false then get_one_rs("hits")=0
	        get_one_rs("hits")=get_one_rs("hits")+1        '累计点击
			get_one_rs.update
			else
			get_one=get_one_rs(gfield)  '读取指定字段
		end if
	  end if
	  get_one_rs.close  
  set get_one_rs=nothing
end function


'***********************************
'    过滤HTML字符
'***********************************
Function RemoveHTML(strHTML) 
   Dim objRegExp, Match, Matches 
   Set objRegExp = New Regexp 
       objRegExp.IgnoreCase = True 
       objRegExp.Global = True 
       '取闭合的<> 
       objRegExp.Pattern = "<.+?>" 
       '进行匹配 
   Set Matches = objRegExp.Execute(strHTML) 
       '遍历匹配集合，并替换掉匹配的项目 
   For Each Match in Matches 
       strHtml=Replace(strHTML,Match.Value,"") 
   Next
       RemoveHTML=strHTML 
   Set objRegExp = Nothing 
End Function


'***********************************
'    截取指定长度字符
'***********************************
function cutStr(str,lens)
dim i
i=1
y=0
txt=RemoveHTML(str)
for i=1 to len(txt)
j=mid(txt,i,1)
if ascw(j)>=0 and ascw(j)<=127 then '汉字外的其他符号,如:!@#,数字,大小写英文字母
y=y+0.5
else '汉字
y=y+1
end if
if -int(-y) >= lens then '截取长度
txt = left(txt,i)&"..."
exit for
end if
next
cutStr=txt
end function


'<><><><><><><><><><><><><><><><><><><><><><><><><><><><>
'<><><><><> 用于返回带参数的Url值
'<><><><><> 调用示例:Call ReUrl("value=5&page=34")
'<><><><><><><><><><><><><><><><><><><><><><><><><><><><>
function ReUrl(keys)
    dim UrlStr,UrlStrs,UrlKey,Uqs,ul,uls,rekey,revalue,key
        key=lcase(keys)
        UrlStr=lcase(request.QueryString)
        UrlStrs="&"&UrlStr&"&"
        if keys<>"" then
           key=split(keys,"&")
           for each ul in key
               uls     = split(ul,"=")
               rekey   = uls(0)
               revalue = ul
               if instr(UrlStrs,"&"&rekey&"=")>0 then
                  UrlStr=replace(UrlStr,rekey&"="&request.QueryString(rekey),revalue)
               else
                  if UrlStr<>"" then Uqs="&"
                  UrlStr=UrlStr&Uqs&revalue
               end if
           next
        end if
    ReUrl="?"&UrlStr
end function
%>





