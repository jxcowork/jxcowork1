<!--#include file="conn.asp"-->
<htm>

<head>
<script language='javascript'>
function xgjc(mm)
{
if(mm=="")
	{alert('���벻��Ϊ��!');
	document.de.mm.focus();
	return false;}
if(mm.indexOf("'")>=0 ||mm.indexOf("<")>=0 ||mm.indexOf(">")>=0 ||mm.indexOf(".")>=0)
{
	alert('����������!');
	document.de.mm.focus();
	return false;
}}

function zcjc(yhm,mm)
{
if(yhm=="")
	{alert('�û�������Ϊ��!');
	document.zc.yhm.focus();
	return false;}
/*var aa ="";
aa = yhm.substring(0,1);
if((aa<='z'&&aa>='a')||(aa<='Z'&&aa>='A'))
{}
else
	{alert('��λ������ĸ!');
	return false;}
*/
if(yhm.indexOf("'")>=0 ||yhm.indexOf("<")>=0 ||yhm.indexOf(">")>=0 ||yhm.indexOf(".")>=0)
	{alert('�û�������������!');
	return false;}

if(mm=="")
	{alert('���벻��Ϊ��!');
	document.zc.mm.focus();
	return false;}
if(mm.indexOf("'")>=0 ||mm.indexOf("<")>=0 ||mm.indexOf(">")>=0 ||mm.indexOf(".")>=0)
	{alert('����������!');
	document.zc.mm.focus();
	return false;}
}
</script>
<%jbm=request.querystring("jbm")
if jbm<>"" then session("tbjbm") = jbm
jbm = session("tbjbm")%>
<link rel="stylesheet" href="style.css" type="text/css">
<title>ע��</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
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
<%cz=request.querystring("cz")
mm=request.form("mm")
yhm=request.form("yhm")
if cz="xgxg" and mm<>"" and yhm<>"" then
	exec="update Tbl_BBSUSER set pass='"+setpwd(mm)+"' where yhm='"+yhm+"'"
	conn.execute exec
	conn.close%>
	<script language='javascript'>
		alert('��ϲ,�޸ĳɹ�!');
		window.location="index.asp";
	</script>
<%end if
if cz="zccg" and mm<>"" and yhm<>"" then
''�ж��Ƿ��ظ�
	set rs=server.createobject("adodb.recordset")
	exec="select * from Tbl_BBSUSER where yhm='"+yhm+"'"
	rs.open exec,conn,1,1 
	if rs.eof and instr(yhm,"��")=0 and instr(yhm,"����Ա")=0 then
		yhm=xrzh(yhm)
		exec="insert into Tbl_BBSUSER (yhm,pass,jb) values ('"+yhm+"','"+setpwd(mm)+"',1)"
		conn.execute exec
		conn.close
		session("tbyhm")=yhm%>
		  <script language='javascript'>
		alert('��ϲ,����ɹ�!');
		window.location="index.asp";
		</script>
      <%else%>
		  <script language='javascript'>
		alert('��Ǹ,���û����Ѿ��������ˣ��뻻������������!');
		window.location="tb_zc.asp";
		</script>
      <%end if
end if%>
</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<%if cz="" then%>
<form name="zc" action="tb_zc.asp?cz=zccg" method="post">
<table align="center"  align="center"> 
  <tr>
    <td align="center" >�û���</td>
	<td><input name="yhm" type="text" size="10" maxlength="10"></td>
</tr>
<tr>
    <td align="center">��  ��</td>
<td><input name="mm" type="password" size="10" maxlength="10">
</td></tr>
<tr><td align="center" colspan="2">
      <input type="submit" name="sdfsd" value="ע   ��" onClick="javascript:return zcjc(yhm.value,mm.value);">
&nbsp;&nbsp;&nbsp;&nbsp;
      <input type="reset" name="dsfsdf" value="��   ��">   
</td></tr>
</table>
</form>
<%else%>
<form name="de" action="tb_zc.asp?cz=xgxg" method="post">
<table align="center"  align="center">
<tr>
    <td align="center" >�û���:</td>
<td><input name="yhm" type="text" size="10" maxlength="10" value="<%=session("tbyhm")%>" readonly=""></td>
</tr>
<tr>
    <td align="center">������:</td>
<td><input name="mm" type="password" size="10" maxlength="10" value="">
</td></tr>
<tr><td align="center" colspan="2">
      <input type="submit" name="sdfsd" value="��   ��" onClick="javascript:return xgjc(yhm.value,mm.value);">
&nbsp;&nbsp;&nbsp;&nbsp;
      <input type="reset" name="dsfsdf" value="��   ��">   
</td></tr>
</table>
</form>
<%end if%>
</body>
</htm>