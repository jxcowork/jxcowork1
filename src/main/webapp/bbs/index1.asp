<!--#include file="conn.asp"-->
<html>
<head>
<meta name="google-site-verification" content="DLmI_0dEUp1qHCiGdzyusTqvMGmF4LFBPulEbbDKDRw" />
<%
set rs=server.createobject("adodb.recordset")
jbm="����ʦ����-����"
yxft=1
%>
<title><%=jbm%></title>
<link rel="stylesheet" href="style1.css" type="text/css">
<link rel="icon" href="http://www.teachersun.com/Images/icon.ico" type="image/x-icon" />
<link rel="shortcut icon" href="http://www.teachersun.com/Images/icon.ico" type="image/x-icon" />
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
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
<script language="javascript">
function check(zt)
{
if(zt=="" || zt.length<3)
	{
	alert("���ⲻ��Ϊ��!���������!");
	document.de.zt.focus();
	return false;
	}
}
</script>
<%
cz=request.querystring("cz")
id=request.querystring("id")
cz=request.querystring("cz")
zd=request.querystring("zd")
ly=request.form("ly")
ys=request.form("ys")
zt=request.form("zt")
tp=request.form("tp")
page =int(Request.QueryString("page"))
if tp="http://" or len(tp)<6 then
	tp=""
else
	tp="<img src="+tp+" alt=����ʦ����-��ͼ>"
end if
tp=xrzh(tp)
ly=xrzh(ly)
zt=xrzh(zt)
yhm=request.querystring("yhm")
if yhm="" then   yhm=session("tbyhm")
if cz="exit" then 
	session("tbyhm")=""
	session("tbjb")=""
	response.redirect("index.asp")
end if

if zt<>"" and cz="tj" then
	IPRule()
	if yhm="" then yhm=request.servervariables("remote_addr")
	zt=xrzh(zt)
	ly=xrzh(ly)
	exec="insert into Tbl_BBSLT (yhm,zt,ly,zid,zhhf,sj,zhsj) values ('"+yhm+"','"+ys+zt+"</font>"+"','"+ys+ly+"</font>"+"<br>"+tp+"','s','"+yhm+"','" + cstr(now()) + "', '" + cstr(now()) + "')"
	
	conn.execute exec
	conn.close
	
	if ly<>"" and 0 then
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
		msg.Subject = "���ɣ����˷���" & zt
		htmlbody="���⣺" & zt & "<br>"
		htmlbody=htmlbody & "���ԣ�" & ly & "<br>"
		htmlbody=htmlbody & "�û�����" & yhm & "<br>"
		htmlbody=htmlbody & "IP��" & request.servervariables("remote_addr")
		msg.HTMLBody=htmlbody
		msg.Send(MailDomain)
	end if
	response.redirect("index.asp")
end if
if id<>"" and cz="sc" and session("tbjb")="radio" then
	exec="delete from Tbl_BBSLT where id="+id+" or zid='"+id+"'"
	conn.execute exec
	conn.close
	response.redirect("index.asp?page="+cstr(page))
end if
if id<>"" and cz="zd" and session("tbjb")="radio" then
	if zd="0" then
		zd="1"
	else
		zd="0"
	end if
	exec="update Tbl_BBSLT set zd='"+zd+"' where id="+id
	conn.execute exec
	conn.close
	response.redirect("index.asp")
end if%>
<body leftmargin="0" topmargin="0">
<table width="100%" border="0" cellspacing="0" cellpadding="0" background="images/top-back.gif" align="center">
  <tr background="http://www.teachersun.com/images/top-back.gif"> 
    <td height="74" align="right" width="180"><a href="http://www.teachersun.com" title="ְ��Ӣ�������Ͽ��ã�����ʦ���ð�����ְ�ƿ��ԡ�����ʦӢ���������Ӣ���Сѧ��Ӣ���Сѧ��Ӣ�￼�ԡ�Ӣ�ｲ�����������Ӣ�ｲ��"><img src="http://www.teachersun.com/images/logo.jpg" width="85" height="78" border="0" title="����ʦ���ã�ְ��Ӣ����Ͽ���"></a>&nbsp;&nbsp;&nbsp;&nbsp;</td>
	<td align="center" height="74"><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000"  codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0" width="468" height="60" id=OBJECT1 title="ְ��Ӣ�������Ͽ��ã�����ʦ���ð�����ְ�ƿ��ԡ�����ʦӢ���������Ӣ���Сѧ��Ӣ���Сѧ��Ӣ�￼�ԡ�Ӣ�ｲ�����������Ӣ�ｲ��">
              <param name=movie value="http://www.teachersun.com/images/banner.swf">
              <param name=quality value=high>
              <embed src="http://www.teachersun.com/images/banner.swf" quality=high pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="468" height="60"> </embed>
            </object></td>
	<td align="center"><table height="37" cellspacing="0" cellpadding="0">
        <TR>
			<TD align=left nowrap="nowrap"><FONT color=#FFA816><B>www.teachersun.com</B></FONT></TD>
		  </TR>
		  <TR>
			<TD align=left nowrap="nowrap"><FONT color=#989898><B>TEL��13621363312</B></FONT></TD>
		  </TR>
		  <TR>
			<TD align=left nowrap="nowrap"><FONT color=#989898><B>E-mail��slskt@263.net</B></FONT></TD>
		  </TR>
            </table>
			<table height="37" cellspacing="0" cellpadding="0">
              <tr>
                <td height="33" align=center><a href="http://www.teachersun.com" title="ְ��Ӣ�������Ͽ��ã�����ʦ���ð�����ְ�ƿ��ԡ�����ʦӢ���������Ӣ���Сѧ��Ӣ���Сѧ��Ӣ�￼�ԡ�Ӣ�ｲ�����������Ӣ�ｲ��"><font color="#FE6700"><b>����ҳ</b></font></a></td>
              </tr>
            </table></td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td><!-- SiteSearch Google -->
<table border="0" bgcolor="#ffffff">
<form method="get" action="http://www.google.com/custom" target="google_window">
<tr><td nowrap="nowrap" valign="top" align="left" height="32">
<a href="http://www.google.com/">
<img src="http://www.google.com/logos/Logo_25wht.gif" border="0" alt="Google" align="middle"></img></a>
</td>
<td nowrap="nowrap">
<input type="hidden" name="domains" value="www.teachersun.com"></input>
<label for="sbi" style="display: none">�������������ִ�</label>
<input type="text" name="q" size="31" maxlength="255" value="" id="sbi"></input>
<label for="sbb" style="display: none">�ύ������</label>
<input type="submit" name="sa" value="����" id="sbb"></input>
<input type="radio" name="sitesearch" value="www.teachersun.com" id="ss0"></input>
<label for="ss0" title="�������� http://www.teachersun.com"><font size="-1" color="#000000">����ʦ����</font></label>
<input type="radio" name="sitesearch" value="bbs.teachersun.com" checked id="ss1"></input>
<label for="ss1" title="���� ���� http://bbs.teachersun.com"><font size="-1" color="#000000">����</font></label>
<input type="hidden" name="client" value="pub-8617701222786522"></input>
<input type="hidden" name="forid" value="1"></input>
<input type="hidden" name="channel" value="6236647609"></input>
<input type="hidden" name="ie" value="GB2312"></input>
<input type="hidden" name="oe" value="GB2312"></input>
<input type="hidden" name="safe" value="active"></input>
<input type="hidden" name="cof" value="GALT:#008000;GL:1;DIV:#336699;VLC:663399;AH:center;BGC:FFFFFF;LBGC:336699;ALC:0000FF;LC:0000FF;T:000000;GFNT:0000FF;GIMP:0000FF;FORID:1"></input>
<input type="hidden" name="hl" value="zh-CN"></input></td></tr>
</form>
</table>
<!-- SiteSearch Google --></td>
  </tr>
</table>
<table width="100%" border="1" cellspacing="0" cellpadding="0" bgcolor="#7894AF" align="center">
  <tr bgcolor="#eeeeee">
    <td height="25" bgcolor="#0B6E81" bordercolordark="#ffffff" bordercolorlight="#000000">&nbsp;&nbsp;&nbsp;&nbsp;<font color="#FFFFFF">��ӭ����&nbsp;&nbsp;����ʦ��������!</font></td>
    <td height="25" bgcolor="#0B6E81" bordercolordark="#ffffff" bordercolorlight="#000000" align="right"><a href="http://www.teachersun.com" target="_blank" title="ְ��Ӣ�������Ͽ��ã�����ʦ���ð�����ְ�ƿ��ԡ�����ʦӢ���������Ӣ���Сѧ��Ӣ���Сѧ��Ӣ�￼�ԡ�Ӣ�ｲ�����������Ӣ�ｲ��"><font color="#FFFFFF">����&nbsp;&nbsp;����ʦ����</font></a></td>
  </tr>
</table>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1">
<tr>
<td height="400" valign="top">
<%
'sql="select count(*) as ooo from Tbl_BBSLT where zid='s' or zid is null"
'rs.open sql,conn,1,1
zts=300		'rs("ooo")
'rs.close
exec="select top 300 * from Tbl_BBSLT where zid='s' or zid is null order by zd,zhsj desc,id desc"
rs.open exec,conn,1,1
rs.PageSize =30
If page < 1 Then page = 1
If page > rs.PageCount Then page = rs.PageCount
if not  rs.eof then  rs.AbsolutePage =page%>
	  <table width="100%" border="0" cellpadding="0" cellspacing="1" align="center">
        <tr bgcolor="#8CDADA">
	      <td align="center" width="4%" height="25">���</td>
	      <td align="center" width="5%" height="25">���</td>
	      <td align="center" width="6%" height="25">����</td>
          <td align="center" width="54%" height="25">����</td>
          <td align="center" width="9%" height="25">����</td>
          <td align="center" width="22%" height="25">���ظ�</td>
        </tr>
	<%For i = 1 To rs.PageSize
	If rs.EOF Then	 Exit For
	zt=rs("zt")
	zd=rs("zd")
	if len(zt)>27 then
		hzt=left(zt,len(zt)-7)
		hzt=right(hzt,len(hzt)-20)
	else
		hzt=zt
	end if
	zzxs = rs("yhm")
	if instr(zzxs,".")>0 then zzxs=left(zzxs,instrrev(zzxs,"."))+"*"
	zhhf=rs("zhhf")
	zhsj=rs("zhsj")
	zhhf=xszh(zhhf)
	if instr(zhhf,".")>0 then zhhf=left(zhhf,instrrev(zhhf,"."))+"*"
	zt=xszh(zt)
	if len(zt)>50 then zt=left(zt,50)+"..."%>
	<tr 
	<%if hys="1" then
	response.write "bgcolor='#8CDADA'"
	hys="0"
	else
    response.write "bgcolor='#EFF9FE'"
	hys="1"
	end if%>>
	      <td align="center" width="4%" height="25"><%=rs.pagesize*(page-1)+cstr(i)%></td>
	      <td align="center" width="5%" height="25"><%=cstr(rs("dj"))%></td>
	      <td align="center" width="6%" height="25"><%=cstr(rs("hf"))%></td>
	      <td width="54%" height="25">&nbsp;&nbsp;&nbsp;
		  <a href="bbs_tj.asp?ztid=<%=rs("id")%>&ykft=<%=yxft%>&jbm=<%=jbm%>&yc=1&cz=ck&zt=<%=hzt%>" target="_blank" title="<%=rs("zt")%>">
		  <%=zt%>
          <%if zd="0" then%>
          <font color="#FF0000">[�ö�]</font> 
          <%end if%>
          </a> 
          <%if session("tbjb")="radio" then%>
          &nbsp;&nbsp; <a href="index.asp?id=<%=rs("id")%>&cz=sc" onClick="javascript:return confirm('ȷ��ɾ��������?')"><font color="#FF0000">ɾ��</font></a> 
          &nbsp;|&nbsp;&nbsp;<a href="index.asp?id=<%=rs("id")%>&cz=zd&zd=<%=zd%>" onClick="javascript:return confirm('�ö�ת����?')"><font color="#FF0000">�ö�ת��</font></a>
          <%end if%>
          </td>
	      <td align="center" width="9%" height="25"><%=zzxs%></td>
	      <td align="center" width="22%" height="25"><%=FormatDateTime(cstr(zhsj), 2)%></td>
	</tr>	
	<%rs.movenext
	next%>
	</table>
<% if 0 then %>
<script type="text/javascript"><!--
google_ad_client = "pub-8617701222786522";
google_ad_width = 728;
google_ad_height = 90;
google_ad_format = "728x90_as";
google_ad_type = "text_image";
google_ad_channel = "";
//--></script>
<script type="text/javascript"
  src="http://pagead2.googlesyndication.com/pagead/show_ads.js">
</script>
<% end if%>
	<center>
        <br>
        ���<font color="#ff0000"><%=zts%></font>������<br>
        &nbsp;&nbsp; ��
        <%for i=1 to rs.PageCount%>
        <a href="index.asp?lx=<%=lx%>&page=<%=i%>"> 
        <%if i=page then
			response.write "<font color=#ff0000>" + cstr(i) + "</font>"
		else
			response.write  "[<font color=#ff0000>"+cstr(i)+"</font>]"
		end if
		if i mod 20 =0 then
			response.write "<br>"
		end if
		%></a> 
        <%next%>
        ҳ&nbsp;&nbsp;<a href="bbs.asp?lx=<%=lx%>&page=<%=page%>" target="_blank"><font color=#ff0000>�����������</font></a></font>
</center>
<%rs.close%>
<form name="de"  method="post" action="index.asp?cz=tj">
    <table width="599" id="tb1">
	<tr><td>�� ��</td><td><input type="text" size="50" name="zt" maxlength="200"></td></tr>
	<tr><td>������ɫ</td><td>
	<select name="ys">
	<option value="<font color=#000000>">��</option>
	<option value="<font color=#FF0000>"><font color=#FF0000>��</font></option>
	<option value="<font color=#ff9900>"><font color=#ff9900>��</font></option>
	<option value="<font color=#0000FF>"><font color=#0000FF>��</font></option>
	<option value="<font color=#FF00FF>"><font color=#FF00FF>��</font></option>
	<option value="<font color=#C0C0C0>"><font color=#C0C0C0>��</font></option>
	<option value="<font color=#FFFFFF>"><font color=#FFFFFF>��</font></option>
	</select></td></tr>
	<tr><td valign="top">����</td><td><textarea  name="ly" rows="5" cols="60"></textarea></td>
	</tr>
	<tr><td>ͼƬ����</td><td><input type="text" name="tp" size="55" maxlength="100" value="http://"></td>
	</tr>
	<tr><td align="center" colspan="2"><%if session("tbyhm")="" then%>
	<font color="#FF0000">Ŀǰ������������</font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
	<a href="tb_dl.asp">��¼</a>&nbsp;|&nbsp; <a href="tb_zc.asp?jbm=<%=jbm%>">ע��</a> 
	<%else%>
	<font color="#FF0000"><%=session("tbyhm")%></font>��Ȩ��:&nbsp;&nbsp; 
	<a href="tb_zc.asp?cz=xg&jbm=<%=jbm%>">�޸���Ϣ</a> &nbsp;&nbsp;|&nbsp;&nbsp; 
	<a href="index.asp?cz=exit">�˳�</a> 
	<%end if%>
	</td>
	</tr>
	<tr><td></td><td>
	<input type="submit" name="Submit" value="��   ��" <%if yhm="" and (yxft="0" or yxft=0) then response.write"disabled"%> onClick="javascript:return check(zt.value);">
	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	<input type="reset" name="Submit2" value="��   ��">
	</td></tr>
</table>
</form>
<%conn.close
set conn=nothing%>
</td></tr></table>
<table width="100%" border="0" cellspacing="1" cellpadding="0" align="center"> 
	<tr bgcolor="#95ADD2" >
    <td height="25" align="center"><a href='http://www.teachersun.com' target="_blank" title="ְ��Ӣ�������Ͽ��ã�����ʦ���ð�����ְ�ƿ��ԡ�����ʦӢ���������Ӣ���Сѧ��Ӣ���Сѧ��Ӣ�￼�ԡ�Ӣ�ｲ�����������Ӣ�ｲ��">����ʦ����-����</a></td></tr>
</table>
</body>
</html>
<script language="javascript">
function ly()
{
	tb1.style.display='block';
}
</script>
<script type="text/javascript" src="http://lead.soperson.com/10029612/10032695.js">
</script>
