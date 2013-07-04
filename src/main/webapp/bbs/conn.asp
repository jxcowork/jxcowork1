<%on error resume next
MailDomain="smtp.ym.163.com"
SerName="info@teachersun.com"
PassWord="in1234"
MailFrom="info@teachersun.com"
MailFromName="info"
MailTo="mingdai@netease.com"
MailToName="大米袋"


set conn=server.createobject("adodb.connection")
connstr="Provider=SQLOLEDB;Data Source=127.0.0.1;Initial Catalog=www_teachersun_com;User Id=teachersun;PASSWORD=J&G(*F$C&^V7hf9234fdf&FDD;"
connstr="DSN=tsun;Uid=teachersun;Pwd=J&G(*F$C&^V7hf9234fdf&FDD"
on error resume next
conn.open connstr
If Err Then
	err.Clear
	Set Conn = Nothing
	Response.Write "网站维护中，请稍候。"
	Response.End
End If

Function SetPwd(sPwd)
    sSql = ""
    For i = 1 To Len(sPwd)
        If bType = True Then
            sSql = sSql & Chr(-Asc(Mid(sPwd, i, 1)) - 10000)
        Else
            sSql = sSql & Chr(-Asc(Mid(sPwd, i, 1)) - 10000)
        End If
    Next
    SetPwd = sSql
End Function

Function IPRule()
	Dim IP, SearchIP
	IP=request.servervariables("remote_addr")
	SearchIP="222.46.18.34|221.218.189.154|221.11.66.82|210.245.147.241|89.97.244.74|83.181.236.174|210.205.32.147|220.207.253.231|60.210.198.249|222.171.154.73|218.28.9.218|60.190.240.68|60.190.240.73|123.165.84.206|58.242.161.164|80.143.194.87|60.190.240.76|221.12.48.10|195.225.178.38|194.8.74.158|125.120.187.65|125.120.176.171|125.118.169.150|125.34.54.71|200.63.42.136|195.225.178.17|195.225.178.35|222.132.169.227|91.214.46.18|213.5.64.19|221.5.67.201|213.5.64.20|74.105.156.34|221.2.98.206|200.192.89.231|94.127.21.116|60.211.184.29|202.165.220.202|213.5.64.211|210.76.97.227|218.21.233.198|210.82.89.29|60.211.231.243|220.194.47.104|58.243.250.40|58.243.254.66|119.148.160.138|210.83.81.73|220.163.156.225|220.163.148.68|220.163.156.45|220.181.94.226|213.186.122.2|123.126.68.20|202.85.216.219|220.181.51.36|49.81.140.132|218.30.103.147|173.199.114.171|173.199.116.27|180.153.240.69|173.199.114.147|218.78.247.110|173.199.119.59"
	if Instr(1, SearchIP, IP, 0) then response.redirect("index.asp")
	Dim SearchString, SearchChar
	SearchString ="travelers car insurance|1 loan resource|autoquote insurance|Unknown|None|乙醚|月赚|隐形耳机|免费打电话|群发|服饰|批发|早泄|风险管理|管道疏通|招商加盟|广告公司|代办证件|代辦證件|可确认后付款|出售|干粉砂浆|证件代办|达信论坛|两性|办证|刻章|管理系统|文凭|培训学校|百家乐|建筑论坛|极品美女|代发广告|斗地主|培训|综合商城|日赚|气枪|高温布|兼职|网上花店|加盟代理|宏发贷款|监理|代开发票|票据|软水器|有意者|q460985040|考试答案qq|瀛欒|代办|卡库|LLL|美女图片|1721613566|确保一次高分通过|私人侦探|378740331|火热订购中|95.88.155.48|隆力奇|GX账号|0731988|代孕妈妈|卵子|禄脴赂麓|孙骗子|卵巢|回复:</font>|鍥炲|锟截革拷|kubagouwu|53union|胃癌|商城建设"   ' 被搜索的字符串。
	SearchStr=split(SearchString,"|")
	SearchChar = zt   ' 要查找字符串 "P"。
	if len(zt)=8 and IsNumeric(zt) then response.redirect("index.asp")
	if SearchChar="回复:</font>" then response.redirect("index.asp")
	if SearchChar="" then response.redirect("index.asp")
	for i=0 to ubound(SearchStr)
		if Instr(1, SearchChar, SearchStr(i)) then response.redirect("index.asp")
	next
	'if Instr(1, SearchString, SearchChar, 0) then response.write("主题有字符串")
	SearchString ="博大考神|在线服务|早泄|群发|免费打电话|月赚|风险管理|针孔摄像机|交友|婚介|考勤机|闪存卡|蓝牙耳机|OEM|除湿机|变频器|安防|门禁|窃听器|出售|干粉砂浆|福彩双色球|[url=http://|监控摄像机|代办证件|促销台|翻译公司|旅游服务|Viagra|华章|斗地主|博彩|设备|[URL=http://|[URL]http://|月经|梅毒|709182385|打胎|nanpuhospital| sex |<a href=|无纺布口罩|.cn|炒股稳赚|私服|禄脴赂麓|LLL|brazil.mcneel.com|inthesetimes|www.easyvoa.com|78348|在考场上没收|526053166|810660651|Aloha|&euro;|531498614|孙骗子|xuniwu|胃癌|88952634"   ' 被搜索的字符串。
	SearchStr=split(SearchString,"|")
	SearchChar = ly
	if SearchChar="回复:</font>" then response.redirect("index.asp")
	if SearchChar="" then response.redirect("index.asp")
	for i=0 to ubound(SearchStr)
		if Instr(1, SearchChar, SearchStr(i)) then response.redirect("index.asp")
	next
End Function

Function Xszh(zh)
if zh<>"" then
	zh=replace(zh,"  ","&nbsp;&nbsp;")
	zh=replace(zh,"&euro;","'")
	zh=replace(zh,">","&gt;")
	zh=replace(zh,"<","&lt;")
	zh=replace(zh,chr(13),"<br>")
	zh=replace(zh,"&lt;br&gt;","<br>")
	zh=replace(zh,"&lt;f","<f")
	zh=replace(zh,"&lt;img","<img")
	zh=replace(zh,"孙老师课堂-贴图&gt;","孙老师课堂-贴图>")
	zh=replace(zh,"&lt;/f","</f")
	zh=replace(zh,"&lt;/F","</F")
	zh=replace(zh,"0&gt;","0>")
	zh=replace(zh,"F&gt;","F>")
	zh=replace(zh,"t&gt;","t>")
	xszh = zh
end if
End Function

Function Xrzh(zh)
if zh<>"" then
	zh=replace(zh,"'","&euro;")
	Xrzh = zh
end if
End Function
%>
<!--#Include File="SqlIn.Asp"-->
