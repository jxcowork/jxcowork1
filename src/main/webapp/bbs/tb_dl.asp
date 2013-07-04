<!--#include file="conn.asp"-->
<html>
<head>
<%jbm=request.querystring("jbm")
if jbm<>"" then session("tbjbm") = jbm
jbm = session("tbjbm")%>
<link rel="stylesheet" href="style.css" type="text/css">
<title>孙老师课堂贴吧-登录</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312"></head>
<body>
<table width="100%" border="0" cellspacing="0" cellpadding="0" background="images/top-back.gif" align="center">
  <tr background="http://www.teachersun.com/images/top-back.gif"> 
    <td height="74" align="center" width="90"><a href="http://www.teachersun.com" title="孙老师课堂--职称英语、中小学生英语"><img src="http://www.teachersun.com/images/logo.jpg" width="85" height="78" border="0" title="孙老师课堂，职称英语，网上课堂"></a></td>
	<td align="left" height="74"><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000"  codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0" width="468" height="60" id=OBJECT1 title="孙老师课堂--职称英语、中小学生英语">
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

<table width="100%" border="1" cellspacing="0" cellpadding="0" bgcolor="#7894AF" align="center">
  <tr bgcolor="#eeeeee">
    <td height="25" bgcolor="#0B6E81" bordercolordark="#ffffff" bordercolorlight="#000000">&nbsp;&nbsp;&nbsp;&nbsp;<font color="#FFFFFF">欢迎来到&nbsp;&nbsp;孙老师课堂贴吧!</font></td>
  </tr>
</table>
<p>
<%yhm=request("yhm")
pass=request("pass")
dz=request("dz")
pass=setpwd(pass)
if yhm<>"" and pass<>"" then
	set rs=server.createobject("adodb.recordset")
	exec="select * from Tbl_BBSUSER where yhm='"+yhm+"' and pass='"+ pass +"'"
	rs.open exec,conn,1,1
	if rs.eof or rs.bof then%>
	<script language="JavaScript" type="text/javascript">
		alert("登陆失败，请核实您的用户名及密码是否正确！");
		window.location="tb_dl.asp";
	</script>
<%	'response.redirect "tb_dl.asp"
	else
		session("tbpass")=rs("pass")
		session("tbyhm")=rs("yhm")
		session("tbjb")=rs("jb")
		if pass<>session("tbpass") then	response.redirect "tb_dl.asp"
		response.redirect "index.asp"
	end if
	rs.close
	conn.close
	
end if%>
<br>
</p>
<form name="from" method="post" action="tb_dl.asp">
<input type="hidden" name="dz" size="10" value="<%=dz%>" maxlength="10">
  <table align="center">
    <tr>
	  <td>用户名</td>
      <td><input type="text" name="yhm" size="10" maxlength="10" class="tablefirst"></td>
	</tr>
	<tr>
	  <td>密 码</td>
      <td><input type="password" name="pass" size="10" maxlength="10" class="tablefirst"></td>
	</tr>
</table>
<table cellpadding="4" cellspacing="0" align="center">
	<tr>
	<td align="center">
	    <input type=submit onClick="javascript:return dl(yhm.value,pass.value);" value="登   陆">
&nbsp;&nbsp;&nbsp;&nbsp;
	    <input type=button onClick="javascript:window.location='tb_zc.asp?cz=zc';" value="注   册">
      </td>
	</tr>
</table>
</form>
</center>
</body>
</htm>
<script language='javascript'>
function dl(yhm,pass)
{
	if (yhm=="")
	{alert('输入用户名!');
	document.from.yhm.focus();
	return false;}
	if (pass=="")
	{alert('输入密码!');
	document.from.pass.focus();
	return false;}}
</script>