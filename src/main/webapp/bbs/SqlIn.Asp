<%
'--------版权说明------------------
'SQL通用防注入程序 V3.0
'2.0强化版，对代码做了一点优化，加入自动封注入者Ip的功能！^_^
'3.0版，加入后台登陆查看注入记录功能，方便网站管理员查看非法记录，以及删除以前的记录，是否对入侵者Ip解除封锁！
'火狐技术联盟 Neeao URL:http://www.wrsky.com
'Neeao站点：http://www.neeao.com 
'Mail:neeaocn@Gamil.com

'--------数据库连接部分--------------
dim dbkillSql,killSqlconn,connkillSql
'dbkillSql="../INC/SqlIn.mdb" 
'On Error Resume Next
'Set killSqlconn = Server.CreateObject("ADODB.Connection")
'connkillSql="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(dbkillSql)
'killSqlconn.Open connkillSql
'If Err Then
'	err.Clear
'	Set killSqlconn = Nothing
'	Response.Write "数据库连接出错，请检查连接字串。"
'	Response.End
'End If


'--------定义部份------------------
Dim Fy_Post,Fy_Get,Fy_In,Fy_Inf,Fy_Xh,Fy_db,Fy_dbstr,Kill_IP,WriteSql
'自定义需要过滤的字串,用 "|" 分隔
Fy_In = " 1=1|''='| and '|exec|insert |select |delete |update |count(|mid(| master|truncate|char(|declare"
Kill_IP=1
WriteSql=True			
'----------------------------------


Fy_Inf = split(Fy_In,"|")
'--------POST部份------------------
If Request.Form<>"" Then
	For Each Fy_Post In Request.Form
		For Fy_Xh=0 To Ubound(Fy_Inf)
			'response.write (LCase(Request.Form(Fy_Post)) & "--" & Fy_Inf(Fy_Xh) & Instr(LCase(Request.Form(Fy_Post)),Fy_Inf(Fy_Xh)) & "<br>")
			If Instr(LCase(Request.Form(Fy_Post)),Fy_Inf(Fy_Xh))<>0 Then
				If WriteSql=True Then
					conn.Execute("insert into Tbl_SqlIn(Sqlin_IP,SqlIn_Web,SqlIn_FS,SqlIn_CS,SqlIn_SJ) values('"&Request.ServerVariables("REMOTE_ADDR")&"','"&Request.ServerVariables("URL")&"','POST','"&Fy_Post&"','"&replace(Request.Form(Fy_Post),"'","''")&"')")
					conn.close
					Set conn = Nothing
				End If
				Response.Write "<Script Language=JavaScript>alert('系统提示你↓\n\n请不要在参数中包含非法字符尝试注入！');</Script>"
				Response.Write "非法操作！系统做了如下记录↓<br>"
				Response.Write "操作ＩＰ："&Request.ServerVariables("REMOTE_ADDR")&"<br>"
				Response.Write "操作时间："&Now&"<br>"
				'Response.Write "操作页面："&Request.ServerVariables("URL")&"<br>"
				'Response.Write "提交方式：ＰＯＳＴ<br>"
				'Response.Write "提交参数："&Fy_Post&"<br>"
				'Response.Write "提交数据："&Request.Form(Fy_Post)
				Response.End
			End If
		Next
	Next
End If
'----------------------------------

'--------GET部份-------------------
If Request.QueryString<>"" Then
	For Each Fy_Get In Request.QueryString
		For Fy_Xh=0 To Ubound(Fy_Inf)
			If Instr(LCase(Request.QueryString(Fy_Get)),Fy_Inf(Fy_Xh))<>0 Then
				If WriteSql=True Then
					conn.Execute("insert into Tbl_SqlIn(Sqlin_IP,SqlIn_Web,SqlIn_FS,SqlIn_CS,SqlIn_SJ) values('"&Request.ServerVariables("REMOTE_ADDR")&"','"&Request.ServerVariables("URL")&"','GET','"&Fy_Get&"','"&replace(Request.QueryString(Fy_Get),"'","''")&"')")
				conn.close
				Set conn = Nothing
				End If
				Response.Write "<Script Language=JavaScript>alert('系统提示你↓\n\n请不要在参数中包含非法字符尝试注入！');</Script>"
				Response.Write "非法操作！系统做了如下记录↓<br>"
				Response.Write "操作ＩＰ："&Request.ServerVariables("REMOTE_ADDR")&"<br>"
				Response.Write "操作时间："&Now&"<br>"
				'Response.Write "操作页面："&Request.ServerVariables("URL")&"<br>"
				'Response.Write "提交方式：ＧＥＴ<br>"
				'Response.Write "提交参数："&Fy_Get&"<br>"
				'Response.Write "提交数据："&Request.QueryString(Fy_Get)
				Response.End
			End If
		Next
	Next
End If

If Kill_IP=1 Then
	Dim Sqlin_IP,rsKill_IP,Kill_IPsql
	Sqlin_IP=Request.ServerVariables("REMOTE_ADDR")
	Kill_IPsql="select Sqlin_IP from Tbl_SqlIn where Sqlin_IP='"&Sqlin_IP&"' and kill_ip=1"
	Set rsKill_IP=conn.execute(Kill_IPsql)
	If Not(rsKill_IP.eof or rsKill_IP.bof) and Err=0 Then
		Response.write "<Script Language=JavaScript>alert('系统提示你↓\n\n你的Ip已经被本系统自动锁定！\n\n如想访问本站请和管理员联系！');</Script>"
	Response.End
	End If
	rsKill_IP.close
	
End If
%>