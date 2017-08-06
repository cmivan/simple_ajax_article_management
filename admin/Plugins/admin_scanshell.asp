<!--#include file="../Bin/ConfigSys.asp"-->

<%
'设置密码
    PASSWORD = "security"
dim Report,RePath
    session("pig")=1
	
if request.QueryString("act")="login" then
	if request.Form("pwd") = PASSWORD then session("pig")=1
end if
%>
<style>td{padding-left:5px;}</style>
<body>
<table width="800" border="0" align="center" cellpadding="0" cellspacing="10" bgcolor="#FFFFFF">
  <tr>
    <td class="forumRow">
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="forMy forMAargin">
  <tr>
    <td align="center" class="forumRow"><span style="padding:5px;">本程序取自<a href="http://www.0x54.org" target="_blank">雷客图ASP站长安全助手</a>的ASP木马查找和可疑文件搜索功能
    </span></td>
    </tr>
</table>
<%If Session("pig") <> 1 then%>
<form name="form1" method="post" action="?act=login">
  <div align="center">Password: 
    <input name="pwd" type="password" size="15" value="<%=PASSWORD%>"> 
    <input type="submit" name="Submit" value="提交">
  </div>
</form>
<%
else
	if request.QueryString("act")<>"scan" then
%>
<form action="?act=scan" method="post" name="form1">
  <table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="forMy" >
                    <tr>
                      <td width="15%" align="right" class="forumRow">检查的路径：</td>
                      <td width="85%" class="forumRow"><input name="path" type="text" style="border:1px solid #999" value="." size="20" />
                      根目录相对路径<a href="#" onClick="form1.path.value='*'">“*”</a>;整网站路径<a href="#" onClick="form1.path.value='\'">“\”</a>;程序所在目录<a href="#" onClick="form1.path.value='.'">“.”</a></td>
                    </tr>
                    <tr>
                      <td align="right" class="forumRow">你要干什么：</td>
                      <td class="forumRow"><input name="radiobutton" type="radio" value="sws" checked>
查ASP木马
  <input type="radio" name="radiobutton" value="sf">
搜索符合条件之文件</td>
                    </tr>
                    <tr>
                      <td width="110" align="right" class="forumRow">查找内容：</td>
                      <td align="left" class="forumRow"><input name="Search_Content" type="text" id="Search_Content" style="border:1px solid #999" size="20">
                      * 要查找的字符串，不填就只进行日期检查</td>
                    </tr>
                    <tr>
                      <td align="right" class="forumRow">修改日期：</td>
                      <td align="left" class="forumRow"><input name="Search_Date" type="text" style="border:1px solid #999" value="<%=Left(Now(),InStr(now()," ")-1)%>" size="20">
                        * 多个日期用;隔开，任意日期填写<a href="#" onClick="javascript:form1.Search_Date.value='ALL'">ALL</a></td>
                    </tr>
                    <tr>
                      <td align="right" class="forumRow">文件类型：</td>
                      <td align="left" class="forumRow"><input name="Search_FileExt" type="text" style="border:1px solid #999" value="*" size="20">
                        * 类型之间用,隔开，*表示所有类型 </td>
                    </tr>
                    <tr>
                      <td align="left" class="forumRow">&nbsp;</td>
                      <td align="left" class="forumRow"><input name="submit" type="submit" style="background:#fff;border:1px solid #999;padding:2px 2px 0px 2px;border-width:1px 3px 1px 3px;" value=" 开始扫描 " /></td>
                    </tr>
                  </table>
      </form>
<%
	else
		server.ScriptTimeout = 600
		if request.Form("path")="" then
			response.Write("No Hack")
			response.End()
		end if
		if request.Form("path")="\" then
			TmpPath = Server.MapPath("\")
		elseif request.Form("path")="." then
			TmpPath = Server.MapPath(".")
		else
			TmpPath = Server.MapPath("\")&"\"&request.Form("path")
		end if
		timer1 = timer
		Sun = 0
		SumFiles = 0
		SumFolders = 1
		If request.Form("radiobutton") = "sws" Then
			DimFileExt = "asp,cer,asa,cdx"
			Call ShowAllFile(TmpPath)
		Else
			If request.Form("path") = "" or request.Form("Search_Date") = "" or request.Form("Search_FileExt") = "" Then
				response.Write("缉捕条件不完全，恕难从命<br><br><a href='javascript:history.go(-1);'>请返回重新输入</a>")
				response.End()
			End If
			DimFileExt = request.Form("Search_fileExt")
			Call ShowAllFile2(TmpPath)
		End If
%>
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="forMy">
  <tr>
    <td align="center" class="forumRow"> Scan WebShell -- ASPSecurity For Hacking  </tr>
  <tr>
    <td valign="top" class="forumRow" style="padding:5px;font-size:12px">
      <table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="forMy forMAargin" >
<tr><td colspan="7" class="forumRow"><span class="forumRow">扫描完毕！一共检查文件夹<font color="#FF0000"><%=SumFolders%></font>个，文件<font color="#FF0000"><%=SumFiles%></font>个，发现可疑点<font color="#FF0000"><%=Sun%></font>个 </span></td>
		      </tr>
			
			 <tr>
			 
			 
<%If request.Form("radiobutton") = "sws" Then%>
			   <td width="18%" class="forumRow">文件相对路径</td>
			   <td width="22%" class="forumRow">特征码</td>
			   <td width="44%" class="forumRow">描述</td>
			   <td width="16%" class="forumRow">创建/修改时间</td>
<%else%>   
			   <td width="50%" class="forumRow">文件相对路径</td>
			   <td width="25%" class="forumRow">文件创建时间</td>
			   <td width="25%" class="forumRow">修改时间</td>
<%end if%>
		      </tr>

			 <%=Report%>
		    </table>
			
<%
timer2 = timer
thetime=cstr(int(((timer2-timer1)*10000 )+0.5)/10)
%>
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="forMy forMAargin">
  <tr>
    <td align="center" class="forumRow">本页执行共用了 <%=thetime%> 毫秒</td>
  </tr></table>
<%
	end if
end if
%>
</td>
  </tr></table>	</td>
  </tr>
</table>
</body>
<%

'遍历处理path及其子目录所有文件
Sub ShowAllFile(Path)
	Set FSO = CreateObject("Scripting.FileSystemObject")
	if not fso.FolderExists(path) then exit sub
	Set f = FSO.GetFolder(Path)
	Set fc2 = f.files
	For Each myfile in fc2
		If CheckExt(FSO.GetExtensionName(path&"\"&myfile.name)) Then
			Call ScanFile(Path&Temp&"\"&myfile.name, "")
			SumFiles = SumFiles + 1
		End If
	Next
	Set fc = f.SubFolders
	For Each f1 in fc
		ShowAllFile path&"\"&f1.name
		SumFolders = SumFolders + 1
    Next
	Set FSO = Nothing
End Sub



Function getPath(path)
	dim ArrPaths,thisPath,TempPath
	    TempPath=path
	    ArrPaths=split(TempPath,"\")
	 if ubound(ArrPaths)>0 then getPath=ArrPaths(ubound(ArrPaths))
End Function

'检测文件
Sub ScanFile(FilePath, InFile)
	If InFile <> "" Then
		Infiles = "<font color=red>该文件被<a href=""http://"&Request.Servervariables("server_name")&"/"&tURLEncode(InFile)&""" target=_blank>"& InFile & "</a>文件包含执行</font>"
	End If
	Set FSOs = CreateObject("Scripting.FileSystemObject")
	on error resume next
	set ofile = fsos.OpenTextFile(FilePath)
	filetxt = Lcase(ofile.readall())
	If err Then Exit Sub end if
	if len(filetxt)>0 then
		'特征码检查
		'T_FilePath=FilePath
		'T_FilePath=replace(T_FilePath,"\","/")
		
	
	
		
'<><><><><><><><><><><><><><><><><><><><><><><><>
	dim ArrPaths,thisPath,TempPath
	    TempPath=FilePath
	    ArrPaths=split(TempPath,"\")
	 if ubound(ArrPaths)>0 then
		for i=0 to ubound(ArrPaths)-1
		    thisPath=thisPath&ArrPaths(i)&"\"
		next
	 end if
	if RePath<>thisPath then
	   RePath=thisPath
	   Report = Report&"<tr><td colspan='4' class='forumRaw'>"&thisPath&"</td></tr>"
	end if
'<><><><><><><><><><><><><><><><><><><><><><><><>


	
		filetxt = vbcrlf & filetxt
		temp = "<a href="""&tURLEncode(replace(replace(FilePath,server.MapPath("./")&"\","",1,1,1),"\","/"))&""" target=_blank>"&getPath(FilePath)&"</a>"
			'Check "WScr"&DoMyBest&"ipt.Shell"
			If instr( filetxt, Lcase("WScr"&DoMyBest&"ipt.Shell") ) or Instr( filetxt, Lcase("clsid:72C24DD5-D70A"&DoMyBest&"-438B-8A42-98424B88AFB8") ) then
				Report = Report&"<tr><td class='forumRow'>"&temp&"</td><td class='forumRow'>WScr"&DoMyBest&"ipt.Shell 或者 clsid:72C24DD5-D70A"&DoMyBest&"-438B-8A42-98424B88AFB8</td><td class='forumRow'><font color=red>危险组件，一般被ASP木马利用</font>"&infiles&"</td><td class='forumRow'>"&GetDateCreate(filepath)&"<br>"&GetDateModify(filepath)&"</td></tr>"
				Sun = Sun + 1
			End if
			'Check "She"&DoMyBest&"ll.Application"
			If instr( filetxt, Lcase("She"&DoMyBest&"ll.Application") ) or Instr( filetxt, Lcase("clsid:13709620-C27"&DoMyBest&"9-11CE-A49E-444553540000") ) then
				Report = Report&"<tr><td class='forumRow'>"&temp&"</td><td class='forumRow'>She"&DoMyBest&"ll.Application 或者 clsid:13709620-C27"&DoMyBest&"9-11CE-A49E-444553540000</td><td class='forumRow'><font color=red>危险组件，一般被ASP木马利用</font>"&infiles&"</td><td class='forumRow'>"&GetDateCreate(filepath)&"<br>"&GetDateModify(filepath)&"</td></tr>"
				Sun = Sun + 1
			End If
			'Check .Encode
			Set regEx = New RegExp
			regEx.IgnoreCase = True
			regEx.Global = True
			regEx.Pattern = "\bLANGUAGE\s*=\s*[""]?\s*(vbscript|jscript|javascript).encode\b"
			If regEx.Test(filetxt) Then
				Report = Report&"<tr><td class='forumRow'>"&temp&"</td><td class='forumRow'>(vbscript|jscript|javascript).Encode</td><td class='forumRow'><font color=red>似乎脚本被加密了</font>"&infiles&"</td><td class='forumRow'>"&GetDateCreate(filepath)&"<br>"&GetDateModify(filepath)&"</td></tr>"
				Sun = Sun + 1
			End If
			'Check my ASP backdoor :(
			regEx.Pattern = "\bEv"&"al\b"
			If regEx.Test(filetxt) Then
				Report = Report&"<tr><td class='forumRow'>"&temp&"</td><td class='forumRow'>Ev"&"al</td><td class='forumRow'>e"&"val()函数可以执行任意ASP代码，被一些后门利用。其形式一般是：ev"&"al(X)<br>但是javascript代码中也可以使用，有可能是误报。"&infiles&"</td><td class='forumRow'>"&GetDateCreate(filepath)&"<br>"&GetDateModify(filepath)&"</td></tr>"
				Sun = Sun + 1
			End If
			'Check exe&cute backdoor
			regEx.Pattern = "[^.]\bExe"&"cute\b"
			If regEx.Test(filetxt) Then
				Report = Report&"<tr><td class='forumRow'>"&temp&"</td><td class='forumRow'>Exec"&"ute</td><td class='forumRow'><font color=red>e"&"xecute()函数可以执行任意ASP代码，被一些后门利用。其形式一般是：ex"&"ecute(X)</font><br>"&infiles&"</td><td class='forumRow'>"&GetDateCreate(filepath)&"<br>"&GetDateModify(filepath)&"</td></tr>"
				Sun = Sun + 1
			End If
			'----------------------Start Update 200605031-----------------------------
			'Check .Create&TextFile and .OpenText&File
			regEx.Pattern = "\.(Open|Create)TextFile\b"
			If regEx.Test(filetxt) Then
				Report = Report&"<tr><td class='forumRow'>"&temp&"</td><td class='forumRow'>.CreateTextFile|.OpenTextFile</td><td class='forumRow'>使用了FSO的CreateTextFile|OpenTextFile函数读写文件"&infiles&"</td><td class='forumRow'>"&GetDateCreate(filepath)&"<br>"&GetDateModify(filepath)&"</td></tr>"
				Sun = Sun + 1
			End If
			'Check .SaveT&oFile
			regEx.Pattern = "\.SaveToFile\b"
			If regEx.Test(filetxt) Then
				Report = Report&"<tr><td class='forumRow'>"&temp&"</td><td class='forumRow'>.SaveToFile</td><td class='forumRow'>使用了Stream的SaveToFile函数写文件"&infiles&"</td><td class='forumRow'>"&GetDateCreate(filepath)&"<br>"&GetDateModify(filepath)&"</td></tr>"
				Sun = Sun + 1
			End If
			'Check .&Save
			regEx.Pattern = "\.Save\b"
			If regEx.Test(filetxt) Then
				Report = Report&"<tr><td class='forumRow'>"&temp&"</td><td class='forumRow'>.Save</td><td class='forumRow'>使用了XMLHTTP的Save函数写文件"&infiles&"</td><td class='forumRow'>"&GetDateCreate(filepath)&"<br>"&GetDateModify(filepath)&"</td></tr>"
				Sun = Sun + 1
			End If
			'------------------              End           ----------------------------
			Set regEx = Nothing
			
		'Check include file
		Set regEx = New RegExp
		regEx.IgnoreCase = True
		regEx.Global = True
		regEx.Pattern = "<!--\s*#include\s*file\s*=\s*"".*"""
		Set Matches = regEx.Execute(filetxt)
		For Each Match in Matches
			tFile = Replace(Mid(Match.Value, Instr(Match.Value, """") + 1, Len(Match.Value) - Instr(Match.Value, """") - 1),"/","\")
			If Not CheckExt(FSOs.GetExtensionName(tFile)) Then
				Call ScanFile( Mid(FilePath,1,InStrRev(FilePath,"\"))&tFile, replace(FilePath,server.MapPath("\")&"\","",1,1,1) )
				SumFiles = SumFiles + 1
			End If
		Next
		Set Matches = Nothing
		Set regEx = Nothing
		
		'Check include virtual
		Set regEx = New RegExp
		regEx.IgnoreCase = True
		regEx.Global = True
		regEx.Pattern = "<!--\s*#include\s*virtual\s*=\s*"".*"""
		Set Matches = regEx.Execute(filetxt)
		For Each Match in Matches
			tFile = Replace(Mid(Match.Value, Instr(Match.Value, """") + 1, Len(Match.Value) - Instr(Match.Value, """") - 1),"/","\")
			If Not CheckExt(FSOs.GetExtensionName(tFile)) Then
				Call ScanFile( Server.MapPath("\")&"\"&tFile, replace(FilePath,server.MapPath("\")&"\","",1,1,1) )
				SumFiles = SumFiles + 1
			End If
		Next
		Set Matches = Nothing
		Set regEx = Nothing
		
		'Check Server&.Execute|Transfer
		Set regEx = New RegExp
		regEx.IgnoreCase = True
		regEx.Global = True
		regEx.Pattern = "Server.(Exec"&"ute|Transfer)([ \t]*|\()"".*"""
		Set Matches = regEx.Execute(filetxt)
		For Each Match in Matches
			tFile = Replace(Mid(Match.Value, Instr(Match.Value, """") + 1, Len(Match.Value) - Instr(Match.Value, """") - 1),"/","\")
			If Not CheckExt(FSOs.GetExtensionName(tFile)) Then
				Call ScanFile( Mid(FilePath,1,InStrRev(FilePath,"\"))&tFile, replace(FilePath,server.MapPath("\")&"\","",1,1,1) )
				SumFiles = SumFiles + 1
			End If
		Next
		Set Matches = Nothing
		Set regEx = Nothing
		
		'Check Server&.Execute|Transfer
		Set regEx = New RegExp
		regEx.IgnoreCase = True
		regEx.Global = True
		regEx.Pattern = "Server.(Exec"&"ute|Transfer)([ \t]*|\()[^""]\)"
		If regEx.Test(filetxt) Then
			Report = Report&"<tr><td class='forumRow'>"&temp&"</td><td class='forumRow'>Server.Exec"&"ute</td><td class='forumRow'><font color=red>不能跟踪检查Server.e"&"xecute()函数执行的文件。请管理员自行检查</font><br>"&infiles&"</td><td class='forumRow'>"&GetDateCreate(filepath)&"<br>"&GetDateModify(filepath)&"</td></tr>"
			Sun = Sun + 1
		End If
		Set Matches = Nothing
		Set regEx = Nothing

		'Check RunatScript
		Set XregEx = New RegExp
		XregEx.IgnoreCase = True
		XregEx.Global = True
		XregEx.Pattern = "<scr"&"ipt\s*(.|\n)*?runat\s*=\s*""?server""?(.|\n)*?>"
		Set XMatches = XregEx.Execute(filetxt)
		For Each Match in XMatches
			tmpLake2 = Mid(Match.Value, 1, InStr(Match.Value, ">"))
			srcSeek = InStr(1, tmpLake2, "src", 1)
			If srcSeek > 0 Then
				srcSeek2 = instr(srcSeek, tmpLake2, "=")
				For i = 1 To 50
					tmp = Mid(tmpLake2, srcSeek2 + i, 1)
					If tmp <> " " and tmp <> chr(9) and tmp <> vbCrLf Then
						Exit For
					End If
				Next
				If tmp = """" Then
					tmpName = Mid(tmpLake2, srcSeek2 + i + 1, Instr(srcSeek2 + i + 1, tmpLake2, """") - srcSeek2 - i - 1)
				Else
					If InStr(srcSeek2 + i + 1, tmpLake2, " ") > 0 Then tmpName = Mid(tmpLake2, srcSeek2 + i, Instr(srcSeek2 + i + 1, tmpLake2, " ") - srcSeek2 - i) Else tmpName = tmpLake2
					If InStr(tmpName, chr(9)) > 0 Then tmpName = Mid(tmpName, 1, Instr(1, tmpName, chr(9)) - 1)
					If InStr(tmpName, vbCrLf) > 0 Then tmpName = Mid(tmpName, 1, Instr(1, tmpName, vbcrlf) - 1)
					If InStr(tmpName, ">") > 0 Then tmpName = Mid(tmpName, 1, Instr(1, tmpName, ">") - 1)
				End If
				Call ScanFile( Mid(FilePath,1,InStrRev(FilePath,"\"))&tmpName , replace(FilePath,server.MapPath("\")&"\","",1,1,1))
				SumFiles = SumFiles + 1
			End If
		Next
		Set Matches = Nothing
		Set regEx = Nothing

		'Check Crea"&"teObject
		Set regEx = New RegExp
		regEx.IgnoreCase = True
		regEx.Global = True
		regEx.Pattern = "CreateO"&"bject[ |\t]*\(.*\)"
		Set Matches = regEx.Execute(filetxt)
		For Each Match in Matches
			If Instr(Match.Value, "&") or Instr(Match.Value, "+") or Instr(Match.Value, """") = 0 or Instr(Match.Value, "(") <> InStrRev(Match.Value, "(") Then
				Report = Report&"<tr><td class='forumRow'>"&temp&"</td><td class='forumRow'>Creat"&"eObject</td><td class='forumRow'>Crea"&"teObject函数使用了变形技术。可能是误报"&infiles&"</td><td class='forumRow'>"&GetDateCreate(filepath)&"<br>"&GetDateModify(filepath)&"</td></tr>"
				Sun = Sun + 1
				exit sub
			End If
		Next
		Set Matches = Nothing
		Set regEx = Nothing
	end if
	set ofile = nothing
	set fsos = nothing
End Sub

'检查文件后缀，如果与预定的匹配即返回TRUE
Function CheckExt(FileExt)
	If DimFileExt = "*" Then CheckExt = True
	Ext = Split(DimFileExt,",")
	For i = 0 To Ubound(Ext)
		If Lcase(FileExt) = Ext(i) Then 
			CheckExt = True
			Exit Function
		End If
	Next
End Function

Function GetDateModify(filepath)
	Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.GetFile(filepath) 
	s = f.DateLastModified 
	set f = nothing
	set fso = nothing
	GetDateModify = s
End Function

Function GetDateCreate(filepath)
	Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.GetFile(filepath) 
	s = f.DateCreated 
	set f = nothing
	set fso = nothing
	GetDateCreate = s
End Function

Function tURLEncode(Str)
	temp = Replace(Str, "%", "%25")
	temp = Replace(temp, "#", "%23")
	temp = Replace(temp, "&", "%26")
	tURLEncode = temp
End Function

Sub ShowAllFile2(Path)
	Set FSO = CreateObject("Scripting.FileSystemObject")
	if not fso.FolderExists(path) then exit sub
	Set f = FSO.GetFolder(Path)
	Set fc2 = f.files
	For Each myfile in fc2
		If CheckExt(FSO.GetExtensionName(path&"\"&myfile.name)) Then
			Call IsFind(Path&"\"&myfile.name)
			SumFiles = SumFiles + 1
		End If
	Next
	Set fc = f.SubFolders
	For Each f1 in fc
		ShowAllFile2 path&"\"&f1.name
		SumFolders = SumFolders + 1
    Next
	Set FSO = Nothing
End Sub

Sub IsFind(thePath)
	theDate = GetDateModify(thePath)
	on error resume next
	theTmp = Mid(theDate, 1, Instr(theDate, " ") - 1)
	if err then exit Sub
	
	xDate = Split(request.Form("Search_Date"),";")
	
	If request.Form("Search_Date") = "ALL" Then ALLTime = True
	
	For i = 0 To Ubound(xDate)
		If theTmp = xDate(i) or ALLTime = True Then 
			If request("Search_Content") <> "" Then
				Set FSOs = CreateObject("Scripting.FileSystemObject")
				set ofile = fsos.OpenTextFile(thePath, 1, false, -2)
				filetxt = Lcase(ofile.readall())
				If Instr( filetxt, LCase(request.Form("Search_Content"))) > 0 Then
					temp = "<a href=""http://"&Request.Servervariables("server_name")&"/"&tURLEncode(Replace(replace(thePath,server.MapPath("\")&"\","",1,1,1),"\","/"))&""" target=_blank>"&replace(thePath,server.MapPath("\")&"\","",1,1,1)&"</a>"
					Report = Report&"<tr><td class='forumRow'>"&temp&"</td><td class='forumRow'>"&GetDateCreate(thePath)&"</td><td class='forumRow'>"&theDate&"</td></tr>"
					Sun = Sun + 1
					Exit Sub
				End If
				ofile.close()
				Set ofile = Nothing
				Set FSOs = Nothing
			Else
				temp = "<a href=""http://"&Request.Servervariables("server_name")&"/"&tURLEncode(Replace(replace(thePath,server.MapPath("\")&"\","",1,1,1),"\","/"))&""" target=_blank>"&replace(thePath,server.MapPath("\")&"\","",1,1,1)&"</a>"
				Report = Report&"<tr><td class='forumRow'>"&temp&"</td><td class='forumRow'>"&GetDateCreate(thePath)&"</td><td class='forumRow'>"&theDate&"</td></tr>"
				Sun = Sun + 1
				Exit Sub
			End If
		End If
	Next
	
End Sub
%>