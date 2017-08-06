<%
'--操作数据库类getList
Class getList
 Public  rs       '--数据集操作对象
 Private rs_query '--数据库连接字符串

 Private Sub Class_Initialize  '--类初始化
  Set rs = Server.CreateObject("Adodb.Recordset")    '--
      rs_query = ""
 End Sub

'---打开数据库记录集 
 Public Sub cOpen(db_table,num,paging)
    '########## 定义排序字符 #############
       typeB_id=request.QueryString("typeB_id")
       typeS_id=request.QueryString("typeS_id")
	   keyword =request("keyword")
   
	   order_sql =" and ok=1 order by order_id asc,id desc"
	   order_sql2=" ok=1 order by order_id asc,id desc"
	   
	   if paging then
		  topNum=" *"
	   else
          topNum=" top "&num&" *"
	   end if

	'########## 处理接收到搜索字符的情况 #############
	if keyword<>"" then
	   keyword_sql =" and title like '%"&request("keyword")&"%'"
	   keyword_sql2=" title like '%"&request("keyword")&"%' and"
	end if
	
	if typeS_id<>"" and isnumeric(typeS_id) then
	   rs_query="select"&topNum&" from "&db_table&" where typeS_id="&typeS_id&keyword_sql&order_sql
	elseif typeB_id<>"" and isnumeric(typeB_id)then 
	   rs_query="select"&topNum&" from "&db_table&" where typeB_id="&typeB_id&keyword_sql&order_sql
	else
	   rs_query="select"&topNum&" from "&db_table&" where"&keyword_sql2&order_sql2
	end if
       rs.open rs_query,conn,1,1
	   
'///是否分页 --------------------------
    if not rs.eof and paging then
       rs.PageSize =num         '每页记录条数
	   iCount   =rs.RecordCount '记录总数
	   iPageSize=rs.PageSize
	   maxpage  =rs.PageCount 
	   page     =request("page")
	   if Not IsNumeric(page) or page="" then
	      page=1
	   else
	      page=cint(page)
	   end if
	   if page<1 then
	      page=1
	   elseif  page>maxpage then
	      page=maxpage
	   end if
	   rs.AbsolutePage=Page
	   if page=maxpage then
	      x=iCount-(maxpage-1)*iPageSize
	   else
	      x=iPageSize
	   end if 
	end if
 End Sub

 
 Public Function cNext()   '---数据库集合移动到下一条
  rs.movenext
 End Function
 
 

 Public Function cCount()  '---数据库集合总条数
  cCount = rs.recordcount
 End Function
 
 Public Function cEof()    '---判断数据库记录是否已经到最后一条
  if rs.eof then
     cEof = true
  else
     cEof = false
  end if  
 End Function
 
 
'---关闭数据库 
 Public Function cClose()
      rs.close
  set rs = nothing
 End Function

End Class
%>
