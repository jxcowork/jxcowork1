<!--#include file="conn.asp"-->
<htm>

<head>
<script language='javascript'>
function xgjc(mm)
{
if(mm=="")
	{alert('密码不能为空!');
	document.de.mm.focus();
	return false;}
if(mm.indexOf("'")>=0 ||mm.indexOf("<")>=0 ||mm.indexOf(">")>=0 ||mm.indexOf(".")>=0)
{
	alert('不能有乱码!');
	document.de.mm.focus();
	return false;
}}

function zcjc(yhm,mm)
{
if(yhm=="")
	{alert('用户名不能为空!');
	document.zc.yhm.focus();
	return false;}
/*var aa ="";
aa = yhm.substring(0,1);
if((aa<='z'&&aa>='a')||(aa<='Z'&&aa>='A'))
{}
else
	{alert('首位不是字母!');
	return false;}
*/
if(yhm.indexOf("'")>=0 ||yhm.indexOf("<")>=0 ||yhm.indexOf(">")>=0 ||yhm.indexOf(".")>=0)
	{alert('用户名不能有乱码!');
	return false;}

if(mm=="")
	{alert('密码不能为空!');
	document.zc.mm.focus();
	return false;}
if(mm.indexOf("'")>=0 ||mm.indexOf("<")>=0 ||mm.indexOf(">")>=0 ||mm.indexOf(".")>=0)
	{alert('不能有乱码!');
	document.zc.mm.focus();
	return false;}
}
</script>
<%jbm=request.querystring("jbm")
if jbm<>"" then session("tbjbm") = jbm
jbm = session("tbjbm")%>
<link rel="stylesheet" href="style.css" type="text/css">
<title>注册</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
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
<%cz=request.querystring("cz")
mm=request.form("mm")
yhm=request.form("yhm")
if cz="xgxg" and mm<>"" and yhm<>"" then
	exec="update Tbl_BBSUSER set pass='"+setpwd(mm)+"' where yhm='"+yhm+"'"
	conn.execute exec
	conn.close%>
	<script language='javascript'>
		alert('恭喜,修改成功!');
		window.location="index.asp";
	</script>
<%end if
if cz="zccg" and mm<>"" and yhm<>"" then
''判断是否重复
	set rs=server.createobject("adodb.recordset")
	exec="select * from Tbl_BBSUSER where yhm='"+yhm+"'"
	rs.open exec,conn,1,1 
	if rs.eof and instr(yhm,"孙")=0 and instr(yhm,"管理员")=0 then
		yhm=xrzh(yhm)
		exec="insert into Tbl_BBSUSER (yhm,pass,jb) values ('"+yhm+"','"+setpwd(mm)+"',1)"
		conn.execute exec
		conn.close
		session("tbyhm")=yhm%>
		  <script language='javascript'>
		alert('恭喜,申请成功!');
		window.location="index.asp";
		</script>
      <%else%>
		  <script language='javascript'>
		alert('抱歉,此用户名已经有人用了，请换名后重新申请!');
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
    <td align="center" >用户名</td>
	<td><input name="yhm" type="text" size="10" maxlength="10"></td>
</tr>
<tr>
    <td align="center">密  码</td>
<td><input name="mm" type="password" size="10" maxlength="10">
</td></tr>
<tr><td align="center" colspan="2">
      <input type="submit" name="sdfsd" value="注   册" onClick="javascript:return zcjc(yhm.value,mm.value);">
&nbsp;&nbsp;&nbsp;&nbsp;
      <input type="reset" name="dsfsdf" value="重   置">   
</td></tr>
</table>
</form>
<%else%>
<form name="de" action="tb_zc.asp?cz=xgxg" method="post">
<table align="center"  align="center">
<tr>
    <td align="center" >用户名:</td>
<td><input name="yhm" type="text" size="10" maxlength="10" value="<%=session("tbyhm")%>" readonly=""></td>
</tr>
<tr>
    <td align="center">新密码:</td>
<td><input name="mm" type="password" size="10" maxlength="10" value="">
</td></tr>
<tr><td align="center" colspan="2">
      <input type="submit" name="sdfsd" value="修   改" onClick="javascript:return xgjc(yhm.value,mm.value);">
&nbsp;&nbsp;&nbsp;&nbsp;
      <input type="reset" name="dsfsdf" value="重   置">   
</td></tr>
</table>
</form>
<%end if%>
</body>
</htm>