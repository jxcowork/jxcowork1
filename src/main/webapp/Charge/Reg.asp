<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!-- #include file="../INC/Ini.asp"-->
<!-- #include file="../INC/conn.asp"-->
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>����ʦ����--����ע��</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="KEYWORDS" content="����ʦ,Ӣ�ｲ��,Ӣ��,���,ְ��Ӣ��,ְ�ƿ���,��������Ӣ��,��ѧ��Ӣ��">
<meta name="DESCRIPTION" content="���Ͽ���ע�᣻����ʦ���ð�����ְ��Ӣ�ְ�ƿ��ԡ���������Ӣ���ѧ��Ӣ���ѧ��Ӣ�￼�ԡ�Сѧ��Ӣ�Сѧ��Ӣ�￼�ԡ�Ӣ�ｲ��">
<link rel="stylesheet" type="text/css" href="../CSS/style.css">
<script language="JavaScript" type="text/javascript">
    function VerifyInput()
    {
		//E-MAIL�ʼ���ַ���
		if (form1.Email){
			if (form1.Email.value == ""){
				alert("����Email��ַ�ǲ���������!!");
				form1.Email.focus();
				return (false);
			}
			var re = new RegExp("^([A-Za-z0-9_|-]+[.]*[A-Za-z0-9_|-]+)+@[A-Za-z0-9|-]+([.][A-Za-z0-9|-]+)*[.][A-Za-z0-9]+$","ig");
			{if (!re.test(form1.Email.value))
				{alert("���ĵ����ʼ���ʽ�������!!");
				form1.Email.focus();
				return false;
				}
			}
		}
		if (form1.UName.value == ""){
			alert("���������ǲ���������!!");
			form1.UName.focus();
			return (false);
		}

		if (form1.Tel.value == ""){
			alert("���ĵ绰�ǲ���������!!");
			form1.Tel.focus();
			return (false);
		}

		if (form1.Addr.value == ""){
			alert("���ĵ�ַ�ǲ���������!!");
			form1.Addr.focus();
			return (false);
		}

		SelectL=false
		for (i=0; i< form1.LCode.length; i++)
		{
			if (form1.LCode(i).checked==true)
			{
				SelectL=true;
				break;
			}
			
		}
		if(SelectL == false)
		{
			alert("����û��ѡ��γ�!!");
			form1.LCode(0).focus();
			return false;
		} 
		return true;

    }
</script>
</head>
<body>
<table width="758" border="0" cellspacing="0" cellpadding="0" align="center">
<tr>
  <td><!-- #include file="../SSI/Head.asp"--></td>
</tr>
<tr>
  <td align="center"><table width="758" border="0" cellspacing="0" cellpadding="0" align="center">
      <TR>
        <TD height=15>&nbsp;</TD>
      </TR>
      <tr>
        <td align="center"><table width=758 border="0" cellspacing="1" bgcolor="#e1e1e1">
            <tr bgcolor="#FFFFFF">
              <td align=center class="TitleR">*****  �������� *****</td>
            </tr>
            <tr bgcolor="#FFFFFF">
              <td align=left> 1����д������Ϣ<br>
                2��ѡ��γ�<br>
                3������������ύ����ť<br>
                4����վ����ʾ��ϸ�ĸ��ʽ��ͬʱ���������䷢��һ�ݱ����ɹ��ż�����������ϸ��<a href="http://www.teachersun.com/Charge/AccountPage.asp"  target="_blank"><font color="red">���ʽ</font></a>�� </td>
            </tr>
            <tr bgcolor="#FFFFFF">
              <td>&nbsp;</td>
            </tr>
            <form action="SubReg.asp" method="post" id="form1" onSubmit="return(VerifyInput())">
              <tr bgcolor="#FFFFFF">
                <td><table width=758 border="0" cellspacing="0">
                    <tr bgcolor="#FFFFFF">
                      <td>Email��</td>
                      <td><input name="Email" id="Email" type="text" maxlength="50">
                        &nbsp;&nbsp;<font color="#FF0000">�������ͨ�������䷢�͵�¼���롣������޸ģ���������д)</font></td>
                    </tr>
                    <tr bgcolor="#FFFFFF">
                      <td>������</td>
                      <td><input name="UName" type="text" maxlength="10">
                        &nbsp;&nbsp;<font color="#FF0000">��������޸ģ���������д)</font></td>
                    </tr>
                    <tr bgcolor="#FFFFFF">
                      <td>�绰��</td>
                      <td><input name="Tel" type="text" maxlength="20">
                      &nbsp;&nbsp;<font color="#FF0000">������)</font></td>
                    </tr>
                    <tr bgcolor="#FFFFFF">
                      <td>��ַ��</td>
                      <td><input name="Addr" type="text" size="80" maxlength="50">
                      &nbsp;&nbsp;<font color="#FF0000">������)</font></td>
                    </tr>
                    <tr bgcolor="#FFFFFF">
                      <td>�ʱࣺ</td>
                      <td><input name="Post" type="text" maxlength="6"></td>
                    </tr>
                    <tr bgcolor="#FFFFFF">
                      <td>���ǴӺδ���<br>֪������վ�ģ�</td>
                      <td valign="middle"><input name="fromknow" type="radio" value="�ٶ�">�ٶ�&nbsp;&nbsp;&nbsp;&nbsp;<input name="fromknow" type="radio" value="Google">Google&nbsp;&nbsp;&nbsp;&nbsp;<input name="fromknow" type="radio" value="������������">������������&nbsp;&nbsp;&nbsp;&nbsp;<input name="fromknow" type="radio" value="���ѽ���">���ѽ���&nbsp;&nbsp;&nbsp;&nbsp;<input name="fromknow" type="radio" value="��ѧԱ">��ѧԱ&nbsp;&nbsp;&nbsp;&nbsp;<input name="fromknow" type="radio" value="��վ����">��վ����&nbsp;&nbsp;&nbsp;&nbsp;<input name="fromknow" type="radio" value="����">����</td>
                    </tr>
                    <tr>
                      <td colspan="2" align="center"><input name="Submit" type="submit" value="�����ύ" class="b2"></td>
                    </tr>
                  </table></td>
              </tr>
              <tr bgcolor="#FFFFFF" height="60">
                <td align=center class="TitleR"><a href="http://www.teachersun.com/Charge/AccountPage.asp"  target="_blank"><font color="red">*****  ����ʦ���û���˺� *****</font></a></td>
              </tr>
              <tr bgcolor="#FFFFFF">
                <td width="1%" nowrap="nowrap"><strong>��ѡ�γ̣�</strong><br><font color="blue" style="font-size:12px; line-height:18px">�����ϲ������ظ����Σ�Ϊ��������ͬѧ��ѧϰ����Ч�������ο�ʱ�������ơ�</font></td>
              </tr>
              <tr>
                <td valign="top"><table width="100%" cellSpacing=1 cellPadding=0 align="center" bgcolor="#e1e1e1">
                    <% if 0 then%>
                    <tr bgcolor="#FFFFFF">
                      <td colspan="7" height="30"></td>
                    </tr>
                    <TR>
                      <TD colspan="7" align="center" bgcolor="#f1f2f4"><font style="font-size:14px; line-height:18px"><strong>ְ��Ӣ����/ģ��/Ѻ���</strong></font></TD>
                    </TR>
                    <TR>
                      <TD colspan="7" align="left" bgcolor="#ffffff"><font color="FF0000" style="font-size:14px; line-height:18px">���μӡ�ְ��Ӣ����/ģ��/Ѻ��ࡷδͨ��ѧԱ������������ض����Żݣ���</font></TD>
                    </TR>
                    <TR bgcolor="fafafa">
                      <TD align=middle height="20">�γ�����</TD>
                      <TD align=middle>�γ�����</TD>
                      <TD align=middle>ѧʱ</TD>
                      <TD align=middle>����ѧʱ</TD>
                      <TD align=middle>ԭ��</TD>
                      <TD align=middle>�Żݼ�</TD>
                      <TD align=middle>ѡ��</TD>
                    </TR>
                    <%	
	sql="select * from Tbl_Lesson where LCode in ('1C00', '2C00', '3C00') order by left(LCode,2)"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,1
	i=1
	while not rs.eof
	if i mod 2 =0 then
		BgColor="#fafafa"
	else
		BgColor="#ffffff"
	end if
	%>
                    <TR bgcolor="<%=BgColor%>">
                      <TD align="center" height=20">&nbsp;<%=rs("LName")%></TD>
                      <TD align="center" nowrap="nowrap"><%=rs("LContent")%></TD>
                      <TD align="center"><%=rs("LHour")%>��ʱ</TD>
                      <TD align="center">����
                        <!--%=rs("WHour")%>Сʱ--></TD>
                      <TD align="center"><s>��<%=rs("LFee")%></s></TD>
                      <TD align="center"><font color='red'>��
                        <%
		  if rs("LSale")=0 then
		  		response.Write("��")
		  else
		  		response.Write(""&rs("LSale")&"</font>")
		  end if
		  %></TD>
                      <TD align="center"><%if rs("FilePath")<>"0" then%>
                        <INPUT id=LCode type="checkbox" value=<%=rs("LCode")%> name=LCode>
                        <%else%>
                        &nbsp;
                        <%end if%></TD>
                    </TR>
                    <%rs.movenext
		i=i+1
	wend
	rs.close()
	%>
                    <tr bgcolor="#FFFFFF">
                      <td colspan="7" align="center"><font color="#FF0000" style="font-size:12px; line-height:18px">������ְ��Ӣ���ڿ��Ժ�����رգ�</font></td>
                    </tr>
                    <tr bgcolor="#FFFFFF">
                      <td colspan="7" align="center"><font color="#FF0000" style="font-size:12px; line-height:18px">�����Żݼ�<%=SaleEndDate%>ǰ��Ч���Ի�ת������Ϊ׼����</font></td>
                    </tr>
                    <%end if%>

                    <%if 0 then%>
                    <tr bgcolor="#FFFFFF">
                      <td colspan="7" height="30"></td>
                    </tr>
                    <TR>
                      <TD colspan="7" align="center" bgcolor="#f1f2f4"><font style="font-size:12px; line-height:18px"><strong>ְ��Ӣ��רҵ��</strong></font></TD>
                    </TR>
                    <TR>
                      <TD colspan="7" align="left" bgcolor="#ffffff"><font color="FF0000" style="font-size:14px; line-height:18px">���μӡ�ְ��Ӣ��רҵ�ࡷδͨ��ѧԱ������������ض����Żݣ���</font></TD>
                    </TR>

                    <TR height="20" bgcolor="fafafa">
                      <TD align=middle>�γ�����</TD>
                      <TD align=middle>�γ�����</TD>
                      <TD align=middle>ѧʱ</TD>
                      <TD align=middle>����ѧʱ</TD>
                      <TD align=middle>ԭ��</TD>
                      <TD align=middle>�Żݼ�</TD>
                      <TD align=middle>ѡ��</TD>
                    </TR>
                    <%	
	sql="select * from Tbl_Lesson where LCode in ('1B00', '2B00', '3B00') order by left(LCode,2)"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,1
	i=1
	while not rs.eof
	if i mod 2 =0 then
		BgColor="#fafafa"
	else
		BgColor="#ffffff"
	end if
	%>
                    <TR height="20" bgcolor="<%=BgColor%>">
                      <TD align="center">&nbsp;<%=rs("LName")%></TD>
                      <TD align="center" nowrap="nowrap"><%=rs("LContent")%></TD>
                      <TD align="center"><%=rs("LHour")%>��ʱ</TD>
                      <TD align="center">����
                        <!--%=rs("WHour")%>Сʱ--></TD>
                      <TD align="center"><s>��<%=rs("LFee")%></s></TD>
                      <TD align="center"><font color='red'>��
                        <%
		  if rs("LSale")=0 then
		  		response.Write("��")
		  else
		  		response.Write(""&rs("LSale")&"</font>")
		  end if
		  %></TD>
                      <TD align="center"><%if rs("FilePath")<>"0" then%>
                        <INPUT id=LCode type="checkbox" value=<%=rs("LCode")%> name=LCode>
                        <%else%>
                        &nbsp;
                        <%end if%></TD>
                    </TR>
                    <%rs.movenext
		i=i+1
	wend
	rs.close()
	%>
					<%if 0 then%>
                    <tr bgcolor="#FFFFFF">
                      <td colspan="7" align="center"><font color="#FF0000" style="font-size:12px; line-height:18px">���˿γ�����һ����ְ��Ӣ��Ӧ�Լ��ɡ���</font></td>
                    </tr>
                    <%end if%>
                    <tr bgcolor="#FFFFFF">
                      <td colspan="7" align="center"><font color="#FF0000" style="font-size:12px; line-height:18px">������ְ��Ӣ���ڿ��Ժ�����رգ�</font></td>
                    </tr>
                    <tr bgcolor="#FFFFFF">
                      <td colspan="7" align="center"><font color="#FF0000" style="font-size:12px; line-height:18px">�����Żݼ�<%=SaleEndDate%>ǰ��Ч���Ի�ת������Ϊ׼����</font></td>
                    </tr>
                    <%end if%>

                    <% if 1 then%>
                    <tr bgcolor="#FFFFFF">
                      <td colspan="7" height="30"></td>
                    </tr>
                    <TR>
                      <TD colspan="7" align="center" bgcolor="#f1f2f4"><font style="font-size:14px; line-height:18px"><strong>ְ��Ӣ��ȫ��</strong></font></TD>
                    </TR>
                    <TR>
                      <TD colspan="7" align="left" bgcolor="#ffffff"><font color="#FF0000">���μӡ�ְ��Ӣ��ȫ�̰ࡷδͨ��ѧԱ��ȫ������ض�һ�Σ�<br>
����ض�������<br>
1.ע���������ʵ����ע�ᣬ����μӹ��ҿ��ԡ�<br>
2.������2013����ҹ��������ڼ䣬�������ṩ���Ĳ�ѯ��ʽ����ʵδͨ��ѧԱ������ע��һ�¡�<br>
3.����ʦ�ر���ʾ��ȫ�̰��ͬѧ���������һû��������2013��10��31��ǰ��������ض������ڲ��ٰ���<br>
4.��������Ч�ɼ����������أ������©д���������׵ȸ���������ɵ���Ч�ɼ��������˿���߰�������ض���</font></TD>
                    </TR>
                    <TR height="20" bgcolor="fafafa">
                      <TD align=middle>�γ�����</TD>
                      <TD align=middle>�γ�����</TD>
                      <TD align=middle>ѧʱ</TD>
                      <TD align=middle>����ѧʱ</TD>
                      <TD align=middle>ԭ��</TD>
                      <TD align=middle>�Żݼ�</TD>
                      <TD align=middle>ѡ��</TD>
                    </TR>
                    <%	
	sql="select * from Tbl_Lesson where LCode in ('1000', '2000', '3000') order by left(LCode,2)"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,1
	i=1
	while not rs.eof
	if i mod 2 =0 then
		BgColor="#fafafa"
	else
		BgColor="#ffffff"
	end if
	%>
                    <TR height="20" bgcolor="<%=BgColor%>">
                      <TD align="center">&nbsp;<%=rs("LName")%></TD>
                      <TD align="center" nowrap="nowrap"><%=rs("LContent")%></TD>
                      <TD align="center"><%=rs("LHour")%>��ʱ</TD>
                      <TD align="center">����
                        <!--%=rs("WHour")%>Сʱ--></TD>
                      <TD align="center"><s>��<%=rs("LFee")%></s></TD>
                      <TD align="center"><font color='red'>��
                        <%
		  if rs("LSale")=0 then
		  		response.Write("��")
		  else
		  		response.Write(""&rs("LSale")&"</font>")
		  end if
		  %></TD>
                      <TD align="center"><%if rs("FilePath")<>"0" and 1 then%>
                        <INPUT id=LCode type="checkbox" value=<%=rs("LCode")%> name=LCode>
                        <%else%>
                        &nbsp;
                        <%end if%></TD>
                    </TR>
                    <%rs.movenext
		i=i+1
	wend
	rs.close()
	%>
					<%if 0 then%>
                    <tr bgcolor="#FFFFFF">
                      <td colspan="7" align="center"><font color="#FF0000" style="font-size:12px; line-height:18px">���˿γ�����һ����ְ��Ӣ��Ӧ�Լ��ɡ���</font></td>
                    </tr>
                    <%end if%>
                    <tr bgcolor="#FFFFFF">
                      <td colspan="7" align="center"><font color="#FF0000" style="font-size:12px; line-height:18px">������ְ��Ӣ���ڿ��Ժ�����رգ�</font></td>
                    </tr>
                    <tr bgcolor="#FFFFFF">
                      <td colspan="7" align="center"><font color="#FF0000" style="font-size:12px; line-height:18px">�����Żݼ�<%=SaleEndDate%>ǰ��Ч���Ի�ת������Ϊ׼����</font></td>
                    </tr>
                    <%end if%>
                    <% if 1 then%>
                    <tr bgcolor="#FFFFFF">
                      <td colspan="7" height="30"></td>
                    </tr>
                    <TR>
                      <TD colspan="7" align="center" bgcolor="#f1f2f4"><font color="red" style="font-size:16px; line-height:20px"><strong>ְ��Ӣ�����ȫ���˷Ѱ�</strong></font></TD>
                    </TR>
                    <TR>
                      <TD colspan="7" align="left" bgcolor="#ffffff">���μӡ�ְ��Ӣ�����ȫ���˷Ѱࡷδͨ��ѧԱ��ȫ���˷ѣ�<br>
ȫ���˷�������<br>
1.ע���������ʵ����ע�ᣬ����μӹ��ҿ��ԡ�<br>
2.������2013����ҹ��������ڼ䣬�������ṩ���Ĳ�ѯ��ʽ����ʵδͨ��ѧԱ������ע��һ�¡�<br>
3.����ʦ�ر���ʾ��ȫ�̰��ͬѧ���������һû��������2013��10��31��ǰ��������ض������ڲ��ٰ���<br>
4.��������Ч�ɼ����������أ������©д���������׵ȸ���������ɵ���Ч�ɼ��������˿���߰�������ض���</TD>
                    </TR>
                    <TR height="20" bgcolor="fafafa">
                      <TD align=middle>�γ�����</TD>
                      <TD align=middle>�γ�����</TD>
                      <TD align=middle>ѧʱ</TD>
                      <TD align=middle>����ѧʱ</TD>
                      <TD align=middle>ԭ��</TD>
                      <TD align=middle>�Żݼ�</TD>
                      <TD align=middle>ѡ��</TD>
                    </TR>
                    <%	
	sql="select * from Tbl_Lesson where LCode in ('1H00', '2H00', '3H00') order by left(LCode,2)"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,1
	i=1
	while not rs.eof
	if i mod 2 =0 then
		BgColor="#fafafa"
	else
		BgColor="#ffffff"
	end if
	%>
                    <TR height="20" bgcolor="<%=BgColor%>">
                      <TD align="center">&nbsp;<%=rs("LName")%></TD>
                      <TD align="center" nowrap="nowrap"><%=rs("LContent")%></TD>
                      <TD align="center"><%=rs("LHour")%>��ʱ</TD>
                      <TD align="center">����
                        <!--%=rs("WHour")%>Сʱ--></TD>
                      <TD align="center"><s>��<%=rs("LFee")%></s></TD>
                      <TD align="center"><font color='red'>��
                        <%
		  if rs("LSale")=0 then
		  		response.Write("��")
		  else
		  		response.Write(""&rs("LSale")&"</font>")
		  end if
		  %></TD>
                      <TD align="center"><%if rs("FilePath")<>"0" and 1 then%>
                        <INPUT id=LCode type="checkbox" value=<%=rs("LCode")%> name=LCode>
                        <%else%>
                        &nbsp;
                        <%end if%></TD>
                    </TR>
                    <%rs.movenext
		i=i+1
	wend
	rs.close()
	%>
					<%if 1 then%>
                    <tr bgcolor="#FFFFFF">
                      <td colspan="7" align="center"><font color="#FF0000" style="font-size:12px; line-height:18px">���˿γ̿γ̿�ͨ�󣬿�������ػ�����רҵ���ֵ����ֽ̲ġ���</font></td>
                    </tr>
                    <%end if%>
                    <tr bgcolor="#FFFFFF">
                      <td colspan="7" align="center"><font color="#FF0000" style="font-size:12px; line-height:18px">������ְ��Ӣ���ڿ��Ժ�����رգ�</font></td>
                    </tr>
                    <tr bgcolor="#FFFFFF">
                      <td colspan="7" align="center"><font color="#FF0000" style="font-size:12px; line-height:18px">�����Żݼ�<%=SaleEndDate%>ǰ��Ч���Ի�ת������Ϊ׼����</font></td>
                    </tr>
                    <%end if%>
                    <% if 0 then%>
                    <tr bgcolor="#FFFFFF">
                      <td colspan="7" height="30"></td>
                    </tr>
                    <TR>
                      <TD colspan="7" align="center" bgcolor="#f1f2f4"><font style="font-size:14px; line-height:18px"><strong>ְ��Ӣ��߼ۺ���������</strong></font></TD>
                    </TR>
                    <TR>
                      <TD colspan="7" align="left" bgcolor="#ffffff">
<!--H3 align="center"><font color="#0000FF"><strong>���߼ۺ��������ࡱ����ʱ����̣��˿γ�ֹͣ����</strong></font></H3-->
<H4 align="left" style="; font-size:16px"><font color="#000000">���γ��շ�<font color="#FF0000"><strong><%=HHZCFee%>Ԫ</strong></font>��������Բ���<font color="#FF0000"><strong>ȫ���˷ѡ�</strong></font>����ʦ������ͬѧ���˰࣬��Ϊ��У����ͨ����<font color="#FF0000"><strong>���ߣ�</strong></font></font></H4>
<font color="#FF0000">ȫ���˷�������</font><BR><font color="#000000">1.ע���������ʵ����ע�ᣬ����μӹ��ҿ��ԡ�<br>2.������2013����ҹ��������ڼ䣬�������ṩ���Ĳ�ѯ��ʽ����ʵδͨ��ѧԱ������ע��һ�¡�<br>
3.����ʦ�ر���ʾ��ȫ�̰��ͬѧ���������һû��������2013��10��31��ǰ�����˷ѡ����ڲ��ٰ���<br>
4.��������Ч�ɼ����������أ������©д���������׵ȸ���������ɵ���Ч�ɼ��������˿</font></TD>
                    </TR>
                    <TR height="20" bgcolor="fafafa">
                      <TD align=middle>�γ�����</TD>
                      <TD align=middle>�γ�����</TD>
                      <TD align=middle>ѧʱ</TD>
                      <TD align=middle>����ѧʱ</TD>
                      <TD align=middle>ԭ��</TD>
                      <TD align=middle>�Żݼ�</TD>
                      <TD align=middle>ѡ��</TD>
                    </TR>
                    <%	
	sql="select * from Tbl_Lesson where LCode in ('1D00', '2D00', '3D00') order by left(LCode,2)"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,1
	i=1
	while not rs.eof
	if i mod 2 =0 then
		BgColor="#fafafa"
	else
		BgColor="#ffffff"
	end if
	%>
                    <TR height="20" bgcolor="<%=BgColor%>">
                      <TD align="center">&nbsp;<%=rs("LName")%></TD>
                      <TD align="center" nowrap="nowrap"><%=rs("LContent")%></TD>
                      <TD align="center"><%=rs("LHour")%>��ʱ</TD>
                      <TD align="center">����
                        <!--%=rs("WHour")%>Сʱ--></TD>
                      <TD align="center"><s>��<%=rs("LFee")%></s></TD>
                      <TD align="center"><font color='red'>��
                        <%
		  if rs("LSale")=0 then
		  		response.Write("��")
		  else
		  		response.Write(""&rs("LSale")&"</font>")
		  end if
		  %></TD>
                      <TD align="center"><%if rs("FilePath")<>"0" and 1 then%>
                        <INPUT id=LCode type="checkbox" value=<%=rs("LCode")%> name=LCode>
                        <%else%>
                        &nbsp;
                        <%end if%></TD>
                    </TR>
                    <%rs.movenext
		i=i+1
	wend
	rs.close()
	%>
                    <tr bgcolor="#FFFFFF">
                      <td colspan="7" align="center"><font color="#FF0000" style="font-size:12px; line-height:18px">���˿γ�����һ����ְ��Ӣ��Ӧ�Լ��ɡ���</font></td>
                    </tr>
                    <tr bgcolor="#FFFFFF">
                      <td colspan="7" align="center"><font color="#FF0000" style="font-size:12px; line-height:18px">������ְ��Ӣ���ڿ��Ժ�����رգ�</font></td>
                    </tr>
                    <%end if%>

                    <tr bgcolor="#FFFFFF">
                      <td colspan="7" height="30"></td>
                    </tr>
                    <TR>
                      <TD colspan="7" align="center" bgcolor="#f1f2f4"><font style="font-size:12px; line-height:18px"><strong>����Ӣ��</strong></font></TD>
                    </TR>
                    <TR height="20" bgcolor="fafafa">
                      <TD align=middle>�γ�����</TD>
                      <TD align=middle>�γ�����</TD>
                      <TD align=middle>ѧʱ</TD>
                      <TD align=middle>����ѧʱ</TD>
                      <TD align=middle>ԭ��</TD>
                      <TD align=middle>�Żݼ�</TD>
                      <TD align=middle>ѡ��</TD>
                    </TR>
                    <%	
	sql="select * from Tbl_Lesson where LCode in ('8000','8A00') order by left(LCode,2)"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,1
	i=1
	while not rs.eof
	if i mod 2 =0 then
		BgColor="#fafafa"
	else
		BgColor="#ffffff"
	end if
	%>
                    <TR height="20" bgcolor="<%=BgColor%>">
                      <TD align="center">&nbsp;<%=rs("LName")%></TD>
                      <TD align="center" nowrap="nowrap"><%=rs("LContent")%></TD>
                      <TD align="center"><%=rs("LHour")%>��ʱ</TD>
                      <TD align="center">����
                        <!--%=rs("WHour")%>Сʱ--></TD>
                      <TD align="center"><s>��<%=rs("LFee")%></s></TD>
                      <TD align="center"><font color='red'>��
                        <%
		  if rs("LSale")=0 then
		  		response.Write("��")
		  else
		  		response.Write(""&rs("LSale")&"</font>")
		  end if
		  %></TD>
                      <TD align="center"><%if rs("FilePath")<>"0" then%>
                        <INPUT id=LCode type="checkbox" value=<%=rs("LCode")%> name=LCode>
                        <%else%>
                        &nbsp;
                        <%end if%></TD>
                    </TR>
                    <%rs.movenext
		i=i+1
	wend
	rs.close()
	%>
                    <tr bgcolor="#FFFFFF">
                      <td colspan="7" align="center"><font color="#FF0000" style="font-size:12px; line-height:18px">�����γ���Ч��Ϊһ�꣨�Կγ̿�ͨ�����𣩣�</font></td>
                    </tr>
                    <tr bgcolor="#FFFFFF">
                      <td colspan="7" align="center"><font color="#FF0000" style="font-size:12px; line-height:18px">�����Żݼ�<%=SaleEndDate%>ǰ��Ч���Ի�ת������Ϊ׼����</font></td>
                    </tr>
                    <tr bgcolor="#FFFFFF">
                      <td colspan="7" height="30"></td>
                    </tr>
                    <TR>
                      <TD colspan="7" align="center" bgcolor="#f1f2f4"><font style="font-size:12px; line-height:18px"><strong>����һ�꼶Ӣ�С������</strong></font></TD>
                    </TR>
                    <TR height="20" bgcolor="fafafa">
                      <TD align=middle>�γ�����</TD>
                      <TD align=middle>�γ�����</TD>
                      <TD align=middle>ѧʱ</TD>
                      <TD align=middle>����ѧʱ</TD>
                      <TD align=middle>ԭ��</TD>
                      <TD align=middle>�Żݼ�</TD>
                      <TD align=middle>ѡ��</TD>
                    </TR>
                    <%	
	sql="select * from Tbl_Lesson where LCode in ('A000') order by left(LCode,2)"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,1
	i=1
	while not rs.eof
	if i mod 2 =0 then
		BgColor="#fafafa"
	else
		BgColor="#ffffff"
	end if
	%>
                    <TR height="20" bgcolor="<%=BgColor%>">
                      <TD align="center">&nbsp;<font color="blue"><%=rs("LName")%></font></TD>
                      <TD align="center" nowrap="nowrap"><%=rs("LContent")%></TD>
                      <TD align="center"><%=rs("LHour")%>��ʱ</TD>
                      <TD align="center">����
                        <!--%=rs("WHour")%>Сʱ--></TD>
                      <TD align="center"><s>��<%=rs("LFee")%></s></TD>
                      <TD align="center"><font color='red'>��
                        <%
		  if rs("LSale")=0 then
		  		response.Write("��")
		  else
		  		response.Write(""&rs("LSale")&"</font>")
		  end if
		  %></TD>
                      <TD align="center"><%if rs("FilePath")<>"0" then%>
                        <INPUT id=LCode type="checkbox" value=<%=rs("LCode")%> name=LCode>
                        <%else%>
                        &nbsp;
                        <%end if%></TD>
                    </TR>
                    <%rs.movenext
		i=i+1
	wend
	rs.close()
	%>
                    <tr bgcolor="#FFFFFF">
                      <td colspan="7" align="center"><font color="#FF0000" style="font-size:12px; line-height:18px">�����Ͼ�������Ӧ�γ̲̽ģ�</font></td>
                    </tr>
                    <tr bgcolor="#FFFFFF">
                      <td colspan="7" align="center"><font color="#FF0000" style="font-size:12px; line-height:18px">�����Żݼ�<% =SaleEndDate%>ǰ��Ч���Ի�ת������Ϊ׼����</font></td>
                    </tr>
                    <tr bgcolor="#FFFFFF">
                      <td colspan="7" height="30"></td>
                    </tr>
                    <TR>
                      <TD colspan="7" align="center" bgcolor="#f1f2f4"><font style="font-size:12px; line-height:18px"><strong>���ж��꼶Ӣ��</strong></font></TD>
                    </TR>
                    <TR height="20" bgcolor="fafafa">
                      <TD align=middle>�γ�����</TD>
                      <TD align=middle>�γ�����</TD>
                      <TD align=middle>ѧʱ</TD>
                      <TD align=middle>����ѧʱ</TD>
                      <TD align=middle>ԭ��</TD>
                      <TD align=middle>�Żݼ�</TD>
                      <TD align=middle>ѡ��</TD>
                    </TR>
                    <%	
	sql="select * from Tbl_Lesson where LCode in ('7000') order by left(LCode,2)"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,1
	i=1
	while not rs.eof
	if i mod 2 =0 then
		BgColor="#fafafa"
	else
		BgColor="#ffffff"
	end if
	%>
                    <TR height="20" bgcolor="<%=BgColor%>">
                      <TD align="center">&nbsp;<%=rs("LName")%></TD>
                      <TD align="center" nowrap="nowrap"><%=rs("LContent")%></TD>
                      <TD align="center"><%=rs("LHour")%>��ʱ</TD>
                      <TD align="center">����
                        <!--%=rs("WHour")%>Сʱ--></TD>
                      <TD align="center"><s>��<%=rs("LFee")%></s></TD>
                      <TD align="center"><font color='red'>��
                        <%
		  if rs("LSale")=0 then
		  		response.Write("��")
		  else
		  		response.Write(""&rs("LSale")&"</font>")
		  end if
		  %></TD>
                      <TD align="center"><%if rs("FilePath")<>"0" then%>
                        <INPUT id=LCode type="checkbox" value=<%=rs("LCode")%> name=LCode>
                        <%else%>
                        &nbsp;
                        <%end if%></TD>
                    </TR>
                    <%rs.movenext
		i=i+1
	wend
	rs.close()
	%>
                    <tr bgcolor="#FFFFFF">
                      <td colspan="7" height="30"></td>
                    </tr>
                    <TR>
                      <TD colspan="7" align="center" bgcolor="#f1f2f4"><font style="font-size:12px; line-height:18px"><strong>�п�Ӣ��������꼶��</strong></font></TD>
                    </TR>
                    <TR height="20" bgcolor="fafafa">
                      <TD align=middle>�γ�����</TD>
                      <TD align=middle>�γ�����</TD>
                      <TD align=middle>ѧʱ</TD>
                      <TD align=middle>����ѧʱ</TD>
                      <TD align=middle>ԭ��</TD>
                      <TD align=middle>�Żݼ�</TD>
                      <TD align=middle>ѡ��</TD>
                    </TR>
                    <%	
	sql="select * from Tbl_Lesson where LCode in ('5000', '5B00', '5C00') order by left(LCode,2)"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,1
	i=1
	while not rs.eof
	if i mod 2 =0 then
		BgColor="#fafafa"
	else
		BgColor="#ffffff"
	end if
	%>
                    <TR height="20" bgcolor="<%=BgColor%>">
                      <TD align="center">&nbsp;<%=rs("LName")%><%if rs("LCode")="5000" then response.write("<br><font color=red>���˿γ̰������пγ̣�</font>")%></TD>
                      <TD align="center" nowrap="nowrap"><%=rs("LContent")%></TD>
                      <TD align="center"><%=rs("LHour")%>��ʱ</TD>
                      <TD align="center">����
                        <!--%=rs("WHour")%>Сʱ--></TD>
                      <TD align="center"><s>��<%=rs("LFee")%></s></TD>
                      <TD align="center"><font color='red'>��
                        <%
		  if rs("LSale")=0 then
		  		response.Write("��")
		  else
		  		response.Write(""&rs("LSale")&"</font>")
		  end if
		  %></TD>
                      <TD align="center"><%if rs("FilePath")<>"0" then%>
                        <INPUT id=LCode type="checkbox" value=<%=rs("LCode")%> name=LCode>
                        <%else%>
                        &nbsp;
                        <%end if%></TD>
                    </TR>
                    <%rs.movenext
		i=i+1
	wend
	rs.close()
	%>
                    </tr>
                    <tr bgcolor="#FFFFFF">
                      <td colspan="7" align="center"><font color="#FF0000" style="font-size:12px; line-height:18px">�����Ͼ�������Ӧ�γ̲̽ģ�</font></td>
                    </tr>
                    <tr bgcolor="#FFFFFF">
                      <td colspan="7" align="center"><font color="#FF0000" style="font-size:12px; line-height:18px">�����γ���Ч��Ϊһ�꣨�Կγ̿�ͨ�����𣩣�</font></td>
                    </tr>
                    <tr bgcolor="#FFFFFF">
                      <td colspan="7" align="center"><font color="#FF0000" style="font-size:12px; line-height:18px">�����Żݼ�<%=SaleEndDate%>ǰ��Ч���Ի�ת������Ϊ׼����</font></td>
                    </tr>
                    <tr bgcolor="#FFFFFF">
                      <td colspan="7" height="30"></td>
                    </tr>
                    <TR>
                      <TD colspan="7" align="center" bgcolor="#f1f2f4"><font style="font-size:12px; line-height:18px"><strong>����ѧ��Ӣ��</strong></font></TD>
                    </TR>
                    <TR height="20" bgcolor="fafafa">
                      <TD align=middle>�γ�����</TD>
                      <TD align=middle>�γ�����</TD>
                      <TD align=middle>ѧʱ</TD>
                      <TD align=middle>����ѧʱ</TD>
                      <TD align=middle>ԭ��</TD>
                      <TD align=middle>�Żݼ�</TD>
                      <TD align=middle>ѡ��</TD>
                    </TR>
                    <%
	sql="select * from Tbl_Lesson where LCode in ('4000', '4B00', '4C00') order by left(LCode,2)"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,1
	dim BgColor
	i=1
	while not rs.eof
	if i mod 2 =0 then
		BgColor="#fafafa"
	else
		BgColor="#ffffff"
	end if
	%>
                    <TR height="20" bgcolor="<%=BgColor%>">
                      <TD align="center">&nbsp;<%=rs("LName")%><%if rs("LCode")="4000" then response.write("<br><font color=red>���˿γ̰������пγ̣�</font>")%></TD>
                      <TD align="center"><%=rs("LContent")%></TD>
                      <TD align="center"><%=rs("LHour")%>��ʱ</TD>
                      <TD align="center">����
                        <!--%=rs("WHour")%>Сʱ--></TD>
                      <TD align="center"><s>��<%=rs("LFee")%></s></TD>
                      <TD align="center"><font color='red'>��
                        <%
		  if rs("LSale")=0 then
		  		response.Write("��")
		  else
		  		response.Write(""&rs("LSale")&"</font>")
		  end if
		  %></TD>
                      <TD align="center"><%if rs("FilePath")<>"0" then%>
                        <INPUT id=LCode type="checkbox" value=<%=rs("LCode")%> name=LCode>
                        <%else%>
                        &nbsp;
                        <%end if%></TD>
                    </TR>
                    <%rs.movenext
		i=i+1
	wend
	rs.close()
	%>
                    <tr bgcolor="#FFFFFF">
                      <td colspan="7" align="center"><font color="#FF0000" style="font-size:12px; line-height:18px">�����γ���Ч��Ϊһ�꣨�Կγ̿�ͨ�����𣩣�</font></td>
                    </tr>
                    <tr bgcolor="#FFFFFF">
                      <td colspan="7" align="center"><font color="#FF0000" style="font-size:12px; line-height:18px">�����Żݼ�<%=SaleEndDate%>ǰ��Ч���Ի�ת������Ϊ׼����</font></td>
                    </tr>
                  </table></td>
                <!--tr>
                <td align="center"><span class="TitleR">ȫ���γ�Ϊ�ֳ�¼������γ̽��ȿ����µĿγ̣�</span></td>
              </tr-->
				<tr bgcolor="#FFFFFF">
				  <td align=left> <span class="Title3">ע������˵����</span><br>
				    1����д������Ϣ<br>
					2��ѡ��γ�<br>
					3������������ύ����ť<br>
					4����վ����ʾ��ϸ�ĸ��ʽ��ͬʱ���������䷢��һ�ݱ����ɹ��ż�����������ϸ�ĸ��ʽ�� </td>
				</tr>
              <tr bgcolor="#FFFFFF">
                <td align="center"><input name="Submit" type="submit" value="�����ύ" class="b2"></td>
              </tr>
            </form>
            <TR>
              <TD height=10 bgcolor="#FFFFFF">&nbsp;</TD>
            </TR>
          </table></td>
      </tr>
      <tr>
        <td><!-- #include file="../SSI/Bottom.asp"--></td>
      </tr>
    </table>
    <% 
	Call CloseConn()
	%>
<script type="text/javascript">
var _bdhmProtocol = (("https:" == document.location.protocol) ? " https://" : " http://");
document.write(unescape("%3Cscript src='" + _bdhmProtocol + "hm.baidu.com/h.js%3F8e1ebf318d56927010e89c9d0f4b269d' type='text/javascript'%3E%3C/script%3E"));
</script>
</body>
</html>