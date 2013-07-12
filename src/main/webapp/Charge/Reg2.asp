<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!-- #include file="../INC/Ini.asp"-->
<!-- #include file="../INC/conn.asp"-->
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>孙老师课堂--报名注册</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="KEYWORDS" content="孙老师,英语讲座,英语,免费,职称英语,职称考试,成人三级英语,中学生英语">
<meta name="DESCRIPTION" content="网上课堂注册；孙老师课堂包括：职称英语、职称考试、成人三级英语、中学生英语、中学生英语考试、小学生英语、小学生英语考试、英语讲座">
<link rel="stylesheet" type="text/css" href="../CSS/style.css">
<script language="JavaScript" type="text/javascript">
    function VerifyInput()
    {
		//E-MAIL邮件地址检测
		if (form1.Email){
			if (form1.Email.value == ""){
				alert("您的Email地址是不是忘填了!!");
				form1.Email.focus();
				return (false);
			}
			var re = new RegExp("^([A-Za-z0-9_|-]+[.]*[A-Za-z0-9_|-]+)+@[A-Za-z0-9|-]+([.][A-Za-z0-9|-]+)*[.][A-Za-z0-9]+$","ig");
			{if (!re.test(form1.Email.value))
				{alert("您的电子邮件格式有问题喔!!");
				form1.Email.focus();
				return false;
				}
			}
		}
		if (form1.UName.value == ""){
			alert("您的姓名是不是忘填了!!");
			form1.UName.focus();
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
			alert("您还没有选择课程!!");
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
              <td align=center class="TitleR">*****  报名流程 *****</td>
            </tr>
            <tr bgcolor="#FFFFFF">
              <td align=left> 1、填写个人信息<br>
                2、选择课程<br>
                3、点击“报名提交”按钮<br>
                4、网站将显示详细的付款方式，同时向您的邮箱发送一份报名成功信件，并附带详细的<a href="http://www.teachersun.com/Charge/AccountPage.asp"  target="_blank"><font color="red">付款方式</font></a>。 </td>
            </tr>
            <tr bgcolor="#FFFFFF">
              <td>&nbsp;</td>
            </tr>
            <form action="SubReg.asp" method="post" id="form1" onSubmit="return(VerifyInput())">
              <tr bgcolor="#FFFFFF">
                <td><table width=758 border="0" cellspacing="0">
                    <tr bgcolor="#FFFFFF">
                      <td>Email：</td>
                      <td><input name="Email" id="Email" type="text" maxlength="50">
                        &nbsp;&nbsp;<font color="#FF0000">（必填，将通过此邮箱发送登录密码。此项不能修改，请慎重填写)</font></td>
                    </tr>
                    <tr bgcolor="#FFFFFF">
                      <td>姓名：</td>
                      <td><input name="UName" type="text" maxlength="10">
                        &nbsp;&nbsp;<font color="#FF0000">（此项不能修改，请慎重填写)</font></td>
                    </tr>
                    <tr bgcolor="#FFFFFF">
                      <td>电话：</td>
                      <td><input name="Tel" type="text" maxlength="20"></td>
                    </tr>
                    <tr bgcolor="#FFFFFF">
                      <td>地址：</td>
                      <td><input name="Addr" type="text" size="80" maxlength="50"></td>
                    </tr>
                    <tr bgcolor="#FFFFFF">
                      <td>邮编：</td>
                      <td><input name="Post" type="text" maxlength="6"></td>
                    </tr>
                    <tr bgcolor="#FFFFFF">
                      <td>您是从何处得<br>知我们网站的：</td>
                      <td valign="middle"><input name="fromknow" type="radio" value="百度">百度&nbsp;&nbsp;&nbsp;&nbsp;<input name="fromknow" type="radio" value="Google">Google&nbsp;&nbsp;&nbsp;&nbsp;<input name="fromknow" type="radio" value="其他搜索引擎">其他搜索引擎&nbsp;&nbsp;&nbsp;&nbsp;<input name="fromknow" type="radio" value="朋友介绍">朋友介绍&nbsp;&nbsp;&nbsp;&nbsp;<input name="fromknow" type="radio" value="老学员">老学员&nbsp;&nbsp;&nbsp;&nbsp;<input name="fromknow" type="radio" value="网站链接">网站链接&nbsp;&nbsp;&nbsp;&nbsp;<input name="fromknow" type="radio" value="其他">其他</td>
                    </tr>
                    <tr>
                      <td colspan="2" align="center"><input name="Submit" type="submit" value="报名提交" class="b2"></td>
                    </tr>
                  </table></td>
              </tr>
              <tr bgcolor="#FFFFFF" height="60">
                <td align=center class="TitleR"><a href="http://www.teachersun.com/Charge/AccountPage.asp"  target="_blank"><font color="red">*****  孙老师课堂汇款账号 *****</font></a></td>
              </tr>
              <tr bgcolor="#FFFFFF">
                <td width="1%" nowrap="nowrap"><strong>所选课程：</strong><br><font color="blue" style="font-size:12px; line-height:18px">基础较差需多次重复听课，为了有利于同学们学习，有效期内听课课时不再限制。</font></td>
              </tr>
              <tr>
                <td valign="top"><table width="100%" cellSpacing=1 cellPadding=0 align="center" bgcolor="#e1e1e1">
                    <% if 1 then%>
                    <tr bgcolor="#FFFFFF">
                      <td colspan="7" height="30"></td>
                    </tr>
                    <TR>
                      <TD colspan="7" align="center" bgcolor="#f1f2f4"><font style="font-size:14px; line-height:18px"><strong>职称英语全程</strong></font></TD>
                    </TR>
                    <TR>
                      <TD colspan="7" align="left" bgcolor="#ffffff"><font color="#FF0000">凡参加《职称英语全程班》未通过学员，全部免费重读一次！<br>
免费重读条件：<br>
1.注册必须用真实姓名注册，必须参加国家考试。<br>
2.必须在2013年国家公布分数期间，向我们提供您的查询方式。核实未通过学员姓名与注册一致。<br>
3.孙老师特别提示：全程班的同学如果今年万一没过，请在2012年10月31日前办理免费重读。过期不再办理！</font></TD>
                    </TR>
                    <TR height="20" bgcolor="fafafa">
                      <TD align=middle>课程名称</TD>
                      <TD align=middle>课程内容</TD>
                      <TD align=middle>学时</TD>
                      <TD align=middle>网上学时</TD>
                      <TD align=middle>原价</TD>
                      <TD align=middle>优惠价</TD>
                      <TD align=middle>选课</TD>
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
                      <TD align="center"><%=rs("LHour")%>课时</TD>
                      <TD align="center">不限
                        <!--%=rs("WHour")%>小时--></TD>
                      <TD align="center"><s>￥<%=rs("LFee")%></s></TD>
                      <TD align="center"><font color='red'>￥
                        <%
		  if rs("LSale")=0 then
		  		response.Write("无")
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
                      <!--td colspan="7" align="center"><font color="#FF0000" style="font-size:12px; line-height:18px">－2012年课程4月6号报名开始－</font></td-->
                      <td colspan="7" align="center"><font color="#FF0000" style="font-size:12px; line-height:18px">－本期职称英语于考试后三天关闭－</font></td>
                    </tr>
                    <tr bgcolor="#FFFFFF">
                      <td colspan="7" align="center"><font color="#FF0000" style="font-size:12px; line-height:18px">－此优惠价<%=SaleEndDate%>前有效（以汇款及转款日期为准）－</font></td>
                    </tr>
                    <%end if%>
                    <%if 0 then%>
                    <tr bgcolor="#FFFFFF">
                      <td colspan="7" height="30"></td>
                    </tr>
                    <TR>
                      <TD colspan="7" align="center" bgcolor="#f1f2f4"><font style="font-size:12px; line-height:18px"><strong>职称英语专业班</strong></font></TD>
                    </TR>
                    <TR>
                      <TD colspan="7" align="left" bgcolor="#ffffff"><font color="FF0000" style="font-size:14px; line-height:18px">凡参加《职称英语专业班》未通过学员，不享受免费重读的优惠！！</font></TD>
                    </TR>

                    <TR height="20" bgcolor="fafafa">
                      <TD align=middle>课程名称</TD>
                      <TD align=middle>课程内容</TD>
                      <TD align=middle>学时</TD>
                      <TD align=middle>网上学时</TD>
                      <TD align=middle>原价</TD>
                      <TD align=middle>优惠价</TD>
                      <TD align=middle>选课</TD>
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
                      <TD align="center"><%=rs("LHour")%>课时</TD>
                      <TD align="center">不限
                        <!--%=rs("WHour")%>小时--></TD>
                      <TD align="center"><s>￥<%=rs("LFee")%></s></TD>
                      <TD align="center"><font color='red'>￥
                        <%
		  if rs("LSale")=0 then
		  		response.Write("无")
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
                      <td colspan="7" align="center"><font color="#FF0000" style="font-size:12px; line-height:18px">－本期职称英语于考试后三天关闭－</font></td>
                    </tr>
                    <tr bgcolor="#FFFFFF">
                      <td colspan="7" align="center"><font color="#FF0000" style="font-size:12px; line-height:18px">－此优惠价<%=SaleEndDate%>前有效（以汇款及转款日期为准）－</font></td>
                    </tr>
                    <%end if%>
                    <% if 0 then%>
                    <tr bgcolor="#FFFFFF">
                      <td colspan="7" height="30"></td>
                    </tr>
                    <TR>
                      <TD colspan="7" align="center" bgcolor="#f1f2f4"><font style="font-size:14px; line-height:18px"><strong>职称英语冲刺/模考/押题班</strong></font></TD>
                    </TR>
                    <TR>
                      <TD colspan="7" align="left" bgcolor="#ffffff"><font color="FF0000" style="font-size:14px; line-height:18px">凡参加《职称英语冲刺/模考/押题班》未通过学员，不享受免费重读的优惠！！</font></TD>
                    </TR>
                    <TR bgcolor="fafafa">
                      <TD align=middle height="20">课程名称</TD>
                      <TD align=middle>课程内容</TD>
                      <TD align=middle>学时</TD>
                      <TD align=middle>网上学时</TD>
                      <TD align=middle>原价</TD>
                      <TD align=middle>优惠价</TD>
                      <TD align=middle>选课</TD>
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
                      <TD align="center"><%=rs("LHour")%>课时</TD>
                      <TD align="center">不限
                        <!--%=rs("WHour")%>小时--></TD>
                      <TD align="center"><s>￥<%=rs("LFee")%></s></TD>
                      <TD align="center"><font color='red'>￥
                        <%
		  if rs("LSale")=0 then
		  		response.Write("无")
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
                      <td colspan="7" align="center"><font color="#FF0000" style="font-size:12px; line-height:18px">－本期职称英语于考试后三天关闭－</font></td>
                    </tr>
                    <tr bgcolor="#FFFFFF">
                      <td colspan="7" align="center"><font color="#FF0000" style="font-size:12px; line-height:18px">－此优惠价<%=SaleEndDate%>前有效（以汇款及转款日期为准）－</font></td>
                    </tr>
                    <%end if%>

                    <tr bgcolor="#FFFFFF">
                      <td colspan="7" height="30"></td>
                    </tr>
                    <TR>
                      <TD colspan="7" align="center" bgcolor="#f1f2f4"><font style="font-size:12px; line-height:18px"><strong>成人英语</strong></font></TD>
                    </TR>
                    <TR height="20" bgcolor="fafafa">
                      <TD align=middle>课程名称</TD>
                      <TD align=middle>课程内容</TD>
                      <TD align=middle>学时</TD>
                      <TD align=middle>网上学时</TD>
                      <TD align=middle>原价</TD>
                      <TD align=middle>优惠价</TD>
                      <TD align=middle>选课</TD>
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
                      <TD align="center"><%=rs("LHour")%>课时</TD>
                      <TD align="center">不限
                        <!--%=rs("WHour")%>小时--></TD>
                      <TD align="center"><s>￥<%=rs("LFee")%></s></TD>
                      <TD align="center"><font color='red'>￥
                        <%
		  if rs("LSale")=0 then
		  		response.Write("无")
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
                      <td colspan="7" align="center"><font color="#FF0000" style="font-size:12px; line-height:18px">－本课程有效期为一年（自课程开通日算起）－</font></td>
                    </tr>
                    <tr bgcolor="#FFFFFF">
                      <td colspan="7" align="center"><font color="#FF0000" style="font-size:12px; line-height:18px">－此优惠价<%=SaleEndDate%>前有效（以汇款及转款日期为准）－</font></td>
                    </tr>
                    <tr bgcolor="#FFFFFF">
                      <td colspan="7" height="30"></td>
                    </tr>
                    <TR>
                      <TD colspan="7" align="center" bgcolor="#f1f2f4"><font style="font-size:12px; line-height:18px"><strong>初中一年级英语（小升初）</strong></font></TD>
                    </TR>
                    <TR height="20" bgcolor="fafafa">
                      <TD align=middle>课程名称</TD>
                      <TD align=middle>课程内容</TD>
                      <TD align=middle>学时</TD>
                      <TD align=middle>网上学时</TD>
                      <TD align=middle>原价</TD>
                      <TD align=middle>优惠价</TD>
                      <TD align=middle>选课</TD>
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
                      <TD align="center"><%=rs("LHour")%>课时</TD>
                      <TD align="center">不限
                        <!--%=rs("WHour")%>小时--></TD>
                      <TD align="center"><s>￥<%=rs("LFee")%></s></TD>
                      <TD align="center"><font color='red'>￥
                        <%
		  if rs("LSale")=0 then
		  		response.Write("无")
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
                    <!--tr bgcolor="#FFFFFF">
                      <td colspan="7" align="left" height="50"><font color="#FF0000" style="font-size:12px; line-height:18px"><b>好消息：初一动漫课程特惠 仅120元/年   2010年4月15日前有效</b></font></td>
                    </tr-->
                    <tr bgcolor="#FFFFFF">
                      <td colspan="7" align="center"><font color="#FF0000" style="font-size:12px; line-height:18px">－此优惠价<% =SaleEndDate%>前有效（以汇款及转款日期为准）－</font></td>
                    </tr>
                    <tr bgcolor="#FFFFFF">
                      <td colspan="7" height="30"></td>
                    </tr>
                    <TR>
                      <TD colspan="7" align="center" bgcolor="#f1f2f4"><font style="font-size:12px; line-height:18px"><strong>初中二年级英语</strong></font></TD>
                    </TR>
                    <TR height="20" bgcolor="fafafa">
                      <TD align=middle>课程名称</TD>
                      <TD align=middle>课程内容</TD>
                      <TD align=middle>学时</TD>
                      <TD align=middle>网上学时</TD>
                      <TD align=middle>原价</TD>
                      <TD align=middle>优惠价</TD>
                      <TD align=middle>选课</TD>
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
                      <TD align="center"><%=rs("LHour")%>课时</TD>
                      <TD align="center">不限
                        <!--%=rs("WHour")%>小时--></TD>
                      <TD align="center"><s>￥<%=rs("LFee")%></s></TD>
                      <TD align="center"><font color='red'>￥
                        <%
		  if rs("LSale")=0 then
		  		response.Write("无")
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
                      <TD colspan="7" align="center" bgcolor="#f1f2f4"><font style="font-size:12px; line-height:18px"><strong>中考英语（初中三年级）</strong></font></TD>
                    </TR>
                    <TR height="20" bgcolor="fafafa">
                      <TD align=middle>课程名称</TD>
                      <TD align=middle>课程内容</TD>
                      <TD align=middle>学时</TD>
                      <TD align=middle>网上学时</TD>
                      <TD align=middle>原价</TD>
                      <TD align=middle>优惠价</TD>
                      <TD align=middle>选课</TD>
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
                      <TD align="center">&nbsp;<%=rs("LName")%><%if rs("LCode")="5000" then response.write("<br><font color=red>（此课程包含下列课程）</font>")%></TD>
                      <TD align="center" nowrap="nowrap"><%=rs("LContent")%></TD>
                      <TD align="center"><%=rs("LHour")%>课时</TD>
                      <TD align="center">不限
                        <!--%=rs("WHour")%>小时--></TD>
                      <TD align="center"><s>￥<%=rs("LFee")%></s></TD>
                      <TD align="center"><font color='red'>￥
                        <%
		  if rs("LSale")=0 then
		  		response.Write("无")
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
                      <td colspan="7" align="center"><font color="#FF0000" style="font-size:12px; line-height:18px">－本课程有效期为一年（自课程开通日算起）－</font></td>
                    </tr>
                    <tr bgcolor="#FFFFFF">
                      <td colspan="7" align="center"><font color="#FF0000" style="font-size:12px; line-height:18px">－此优惠价<%=SaleEndDate%>前有效（以汇款及转款日期为准）－</font></td>
                    </tr>
                    <tr bgcolor="#FFFFFF">
                      <td colspan="7" height="30"></td>
                    </tr>
                    <TR>
                      <TD colspan="7" align="center" bgcolor="#f1f2f4"><font style="font-size:12px; line-height:18px"><strong>高中学生英语</strong></font></TD>
                    </TR>
                    <TR height="20" bgcolor="fafafa">
                      <TD align=middle>课程名称</TD>
                      <TD align=middle>课程内容</TD>
                      <TD align=middle>学时</TD>
                      <TD align=middle>网上学时</TD>
                      <TD align=middle>原价</TD>
                      <TD align=middle>优惠价</TD>
                      <TD align=middle>选课</TD>
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
                      <TD align="center">&nbsp;<%=rs("LName")%><%if rs("LCode")="4000" then response.write("<br><font color=red>（此课程包含下列课程）</font>")%></TD>
                      <TD align="center"><%=rs("LContent")%></TD>
                      <TD align="center"><%=rs("LHour")%>课时</TD>
                      <TD align="center">不限
                        <!--%=rs("WHour")%>小时--></TD>
                      <TD align="center"><s>￥<%=rs("LFee")%></s></TD>
                      <TD align="center"><font color='red'>￥
                        <%
		  if rs("LSale")=0 then
		  		response.Write("无")
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
                      <td colspan="7" align="center"><font color="#FF0000" style="font-size:12px; line-height:18px">－本课程有效期为一年（自课程开通日算起）－</font></td>
                    </tr>
                    <tr bgcolor="#FFFFFF">
                      <td colspan="7" align="center"><font color="#FF0000" style="font-size:12px; line-height:18px">－此优惠价<%=SaleEndDate%>前有效（以汇款及转款日期为准）－</font></td>
                    </tr>
                  </table></td>
                <!--tr>
                <td align="center"><span class="TitleR">全部课程为现场录音，随课程进度开放新的课程！</span></td>
              </tr-->
				<tr bgcolor="#FFFFFF">
				  <td align=left> <span class="Title3">注册流程说明：</span><br>
				    1、填写个人信息<br>
					2、选择课程<br>
					3、点击“报名提交”按钮<br>
					4、网站将显示详细的付款方式，同时向您的邮箱发送一份报名成功信件，并附带详细的付款方式。 </td>
				</tr>
              <tr bgcolor="#FFFFFF">
                <td align="center"><input name="Submit" type="submit" value="报名提交" class="b2"></td>
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
</body>
</html>