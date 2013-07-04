<!--#include file="conn.asp"-->
<%zt=request("zt")
if zt="" then zt=request("zt")
if instr(zt,"回复:")=0 then
	zt="回复:"+zt
end if
jbm=request.querystring("jbm")
if jbm<>"" then session("tbjbm") = jbm
jbm = session("tbjbm")
if zt<>"" then	session("tbyzt")=zt%>
<html>
<head>
<script language='javascript'>
function jc(zt,ly,tp,yzt)
{
if(zt=="")
{
alert("主题不能为空");
document.tj.zt.focus();
return false;
}
if(zt!="" && ly.length==0 && tp.length==7 && zt==yzt)
	{
	alert("不能发空信息!");
	document.tj.ly.focus();
	return false;
	}
}
</script>

<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link rel="stylesheet" href="style1.css" type="text/css">
<title><%=right(zt,len(zt)-3)%></title>
<script type="text/javascript">
var gaJsHost = (("https:" == document.location.protocol) ? "https://ssl." : "http://www.");
//document.write(unescape("%3Cscript src='" + gaJsHost + "google-analytics.com/ga.js' type='text/javascript'%3E%3C/script%3E"));
document.write(unescape("%3Cscript src='script/ga.js' type='text/javascript'%3E%3C/script%3E"));
</script>
<script type="text/javascript">
try {
var pageTracker = _gat._getTracker("UA-8383834-1");
pageTracker._trackPageview();
} catch(err) {}</script>
</head>
<%cz=request.querystring("cz")
id=request.querystring("id")
ykft=request.querystring("ykft")
ztid=request.querystring("ztid")
yc=request.querystring("yc")
ly=request.form("ly")
ys=request.form("ys")
tp=request.form("tp")
if tp="http://" or len(tp)<6 then
	tp=""
else
	tp="<img src="+tp+" alt=孙老师课堂-贴图>"
end if
yhm=session("tbyhm")
set rs=server.createobject("adodb.recordset")
page =int(Request.QueryString("page"))

if ztid<>"" then
	exec="select * from Tbl_BBSLT where id="+ztid
	rs.open exec,conn,1,1
	if (not rs.EOF) and (not rs.BOF)  then
		dj=rs("dj")+1
		hf=rs("hf")+1
	else
		dj=1
		hf=1
	end if
	rs.close
	if cz<>"tj" and yc="1" then
		exec="update Tbl_BBSLT set dj="+cstr(dj)+" where id="+ztid
		conn.execute exec
	end if
end if

if trim(zt)<>"" and cz="tj" then
	IPRule()
	if yhm="" then yhm=request.servervariables("remote_addr")
	zt=xrzh(zt)
	ly=xrzh(ly)
	exec="insert into Tbl_BBSLT (yhm,zt,ly,zid,sj) values ('"+yhm+"','"+ys+zt+"</font>"+"','"+ys+ly+"</font>"+"<br>"+tp+"','"+ztid+"','"+cstr(now())+ "')"
	conn.execute exec
	exec="update Tbl_BBSLT set hf="+cstr(hf)+",zhhf='"+yhm+"',zhsj='"+cstr(now())+"' where id="+ztid
	conn.execute exec
	conn.close
	if trim(ly)<>"" and 0 then
		set msg = Server.CreateOBject( "JMail.Message" )
		msg.Logging = true
		msg.silent = true
		msg.Charset="gb2312"
		msg.MailServerUserName=SerName
		msg.MailServerPassWord=PassWord
		msg.MailDomain=MailDomain
		msg.From = MailFrom
		msg.FromName = MailFromName
		msg.AddRecipient MailTo, MailToName
		msg.Subject = "贴吧－有人 " & zt
		htmlbody="主题：" & zt & "<br>"
		htmlbody=htmlbody & "留言：" & ly & "<br>"
		htmlbody=htmlbody & "用户名：" & yhm & "<br>"
		htmlbody=htmlbody & "IP：" & request.servervariables("remote_addr")
		msg.HTMLBody=htmlbody
		msg.Send(MailDomain)
	end if
	
	response.redirect("bbs_tj.asp?zt="+zt+"&ztid="+ztid)
end if
'yyzm = yzm()
if id<>"" and cz="sc" and session("tbjb")="radio" then
	exec="delete from Tbl_BBSLT where id="+id
	conn.execute exec
	exec="update Tbl_BBSLT set hf="+cstr(hf-2)+" where id="+ztid
	conn.execute exec
	conn.close
	response.redirect("bbs_tj.asp?zt="+zt+"&ztid="+ztid+"&page="+cstr(page))
end if
%>
<body leftmargin="0" topmargin="0">
<table width="100%" border="0" cellspacing="0" cellpadding="0" background="images/top-back.gif" align="center">
  <tr background="http://www.teachersun.com/images/top-back.gif"> 
    <td height="74" align="center" width="180"><a href="http://www.teachersun.com" title="孙老师课堂--职称英语、中小学生英语"><img src="http://www.teachersun.com/images/logo.jpg" width="85" height="78" border="0" title="孙老师课堂，职称英语，网上课堂"></a></td>
	<td align="center" height="74"><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000"  codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0" width="468" height="60" id=OBJECT1 title="孙老师课堂--职称英语、中小学生英语">
              <param name=movie value="http://www.teachersun.com/images/banner.swf">
              <param name=quality value=high>
              <embed src="http://www.teachersun.com/images/banner.swf" quality=high pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="468" height="60"> </embed>
            </object></td>
	<td align="center"><table height="37" cellspacing="0" cellpadding="0">
                      <TR>
			<TD align=left nowrap="nowrap"><FONT color=#FFA816><B>www.teachersun.com</B></FONT></TD>
		  </TR>
		  <TR>
			<TD align=left nowrap="nowrap"><FONT color=#989898><B>TEL：13621363312</B></FONT></TD>
		  </TR>
		  <TR>
			<TD align=left nowrap="nowrap"><FONT color=#989898><B>E-mail：slskt@263.net</B></FONT></TD>
		  </TR>
            </table></td>
  </tr>
</table>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" align="center">  
<tr>
  <td> 
    <%if ztid<>"" then	
	exec="select * from Tbl_BBSLT where id="+ztid
	rs.open exec,conn,1,1
	if not rs.eof then ztcz="1"
	zt=xszh(rs("zt"))
	ly=xszh(rs("ly"))
	zzxs=rs("yhm")
	if instr(zzxs,".")>0 then zzxs=left(zzxs,instrrev(zzxs,"."))+"*"
	zhhf=rs("sj")%>
    <table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#EFF9FE">
	<tr><td width="3%" align="center"></td><td><b><%=zt%></b></td></tr>
	<tr><td width="3%"></td><td><%=ly%></td></tr>
	<tr bgcolor="#CCD8E6"><td width="3%"></td>
        <td><b>作者: <%=zzxs%>&nbsp;&nbsp;<%=zhhf%></b> </td>
      </tr>
	</table>
	<%exec="select * from Tbl_BBSLT where zid='"+ztid+"' order by id"
	rs.close
	rs.open exec,conn,1,1
	rs.PageSize =30
	If page < 1 Then    page = 1
	If page > rs.PageCount Then    page = rs.PageCount
	if not  rs.eof then  rs.AbsolutePage =page
	For i = 1 To rs.PageSize
	If rs.EOF Then	 Exit For
	ysj="1"
	zt=replace(rs("zt"),chr(13),"<br>")
	if len(zt)>27 then
		hzt=left(zt,len(zt)-7)
		hzt=right(hzt,len(hzt)-20)
	else
		hzt=zt
	end if
	zt=xszh(zt)
	zzxs=rs("yhm")
	if instr(zzxs,".")>0 then zzxs=left(zzxs,instrrev(zzxs,"."))+"*"
	ly=xszh(rs("ly"))%>
	<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#EFF9FE">
	<tr><td width="3%" align="center"><%=rs.pagesize*(page-1)+cstr(i)%></td><td><b><%=zt%></b></td></tr>
	<tr><td width="3%"></td><td><%=ly%></td></tr>
	<tr bgcolor="#CCD8E6"><td width="3%"></td>
        <td><b>作者: <%=zzxs%>&nbsp;&nbsp;<%=rs("sj")%></b> 
          <%if session("tbjb")="radio" then%>
          &nbsp;&nbsp; &nbsp;&nbsp;
		  <a href="bbs_tj.asp?id=<%=rs("id")%>&ztid=<%=ztid%>&cz=sc" onClick="javascript:return confirm('确定删除此贴吗?')">删除</a>
          <%end if%>
        </td>
      </tr>
	</table>
	<%rs.movenext
next
if ysj="1" then%>
<center>
	每页<font color="#FF0000"><%=rs.pagesize%></font>条回复
	<br>
	第<%for i=1 to rs.PageCount%>
      <a href="bbs_tj.asp?ztid=<%=ztid%>&cz=ck&zt=<%=session("tbyzt")%>&page=<%=i%>"> 
    <%if i=page then
		response.write i
	else
		response.write  "["+cstr(i)+"]"
	end if%>
      </a> 
      <%next%>页
</center>
<%end if
rs.close
end if%>
<form name="tj"  method="post" action="bbs_tj.asp?cz=tj&ztid=<%=ztid%>">
<a name=1></a>
<table>
	<tr><td>标题</td><td><input type="text" size="50" name="zt" maxlength="200" value="<%=session("tbyzt")%>"></td></tr>
	<tr><td> 字体颜色</td><td><select name="ys">
	<option value="<font color=#000000>">黑</option>
	<option value="<font color=#FF0000>"><font color=#FF0000>红</font></option>
	<option value="<font color=#ff9900>"><font color=#ff9900>橙</font></option>
	<option value="<font color=#0000FF>"><font color=#0000FF>蓝</font></option>
	<option value="<font color=#FF00FF>"><font color=#FF00FF>紫</font></option>
	<option value="<font color=#C0C0C0>"><font color=#C0C0C0>灰</font></option>
	<option value="<font color=#FFFFFF>"><font color=#FFFFFF>白</font></option>
    </select>
	</td></tr>
	<tr><td valign="top">内容</td><td><textarea  name="ly" rows="5" cols="60"></textarea></td></tr>
	<tr><td>图片链接</td><td><input type="text" name="tp" value="http://" size="50"></td></tr>
	<tr><td></td><td> 
	<input type="submit" name="ok" value="提   交" <%'if (yhm="" and ykft="0") or ztcz<>"1" or yc<>"1" then response.write"disabled"%> onClick="javascript:return jc(zt.value,ly.value,tp.value,'<%=session("tbyzt")%>');">
	&nbsp;&nbsp;&nbsp;&nbsp;
	<input type="reset" name="Submit2" value="清   除"></td></tr>
</table>
</form>
</body>
</html>