<!--#include file="conn.asp"-->
<html>
<head>
<%jbm=request.querystring("jbm")
if jbm<>"" then session("tbjbm") = jbm
jbm = session("tbjbm")%>
<link rel="stylesheet" href="style.css" type="text/css">
<title>����ʦ��������-��¼</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312"></head>
<body>
<table width="100%" border="0" cellspacing="0" cellpadding="0" background="images/top-back.gif" align="center">
  <tr background="http://www.teachersun.com/images/top-back.gif"> 
    <td height="74" align="center" width="90"><a href="http://www.teachersun.com" title="����ʦ����--ְ��Ӣ���Сѧ��Ӣ��"><img src="http://www.teachersun.com/images/logo.jpg" width="85" height="78" border="0" title="����ʦ���ã�ְ��Ӣ����Ͽ���"></a></td>
	<td align="left" height="74"><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000"  codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0" width="468" height="60" id=OBJECT1 title="����ʦ����--ְ��Ӣ���Сѧ��Ӣ��">
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
            </table></td>
  </tr>
</table>

<table width="100%" border="1" cellspacing="0" cellpadding="0" bgcolor="#7894AF" align="center">
  <tr bgcolor="#eeeeee">
    <td height="25" bgcolor="#0B6E81" bordercolordark="#ffffff" bordercolorlight="#000000">&nbsp;&nbsp;&nbsp;&nbsp;<font color="#FFFFFF">��ӭ����&nbsp;&nbsp;����ʦ��������!</font></td>
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
		alert("��½ʧ�ܣ����ʵ�����û����������Ƿ���ȷ��");
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
	  <td>�û���</td>
      <td><input type="text" name="yhm" size="10" maxlength="10" class="tablefirst"></td>
	</tr>
	<tr>
	  <td>�� ��</td>
      <td><input type="password" name="pass" size="10" maxlength="10" class="tablefirst"></td>
	</tr>
</table>
<table cellpadding="4" cellspacing="0" align="center">
	<tr>
	<td align="center">
	    <input type=submit onClick="javascript:return dl(yhm.value,pass.value);" value="��   ½">
&nbsp;&nbsp;&nbsp;&nbsp;
	    <input type=button onClick="javascript:window.location='tb_zc.asp?cz=zc';" value="ע   ��">
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
	{alert('�����û���!');
	document.from.yhm.focus();
	return false;}
	if (pass=="")
	{alert('��������!');
	document.from.pass.focus();
	return false;}}
</script>