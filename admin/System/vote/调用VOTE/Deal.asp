<%DB_PATH="../"%>
<!--#include file="../Inc/Conn.asp"-->

<%
if isempty(session("go")) then  'ʹ��session��¼�Ƿ���ͶƱ�� ʹ��ʱ���ױ�ˢ
   choice=request("choice")
   if isempty(choice) then
response.write"<script language='javascript'>alert('����ûѡ��ͶƱ���ݣ�');history.go(-1);</script>"
   else
      set rs=server.CreateObject("adodb.recordset")
      sql="update vote_choice set num=num+1 where id in("&choice&")"
      rs.open sql,connstr
      session("go")=1
      response.write "<script language='javascript'>alert('ͶƱ�ɹ�');</script>"
      response.write "<meta http-equiv='refresh' content='1;url=Show.asp?id="&request("id")&"'>"
   end if
else
   response.write "<script language='javascript'>alert('���Ѿ�Ͷ��Ʊ�ˣ�');history.go(-1)</script>"
end if
   
%>