<%on error resume next
MailDomain="smtp.ym.163.com"
SerName="info@teachersun.com"
PassWord="in1234"
MailFrom="info@teachersun.com"
MailFromName="info"
MailTo="mingdai@netease.com"
MailToName="���״�"


set conn=server.createobject("adodb.connection")
connstr="Provider=SQLOLEDB;Data Source=127.0.0.1;Initial Catalog=www_teachersun_com;User Id=teachersun;PASSWORD=J&G(*F$C&^V7hf9234fdf&FDD;"
connstr="DSN=tsun;Uid=teachersun;Pwd=J&G(*F$C&^V7hf9234fdf&FDD"
on error resume next
conn.open connstr
If Err Then
	err.Clear
	Set Conn = Nothing
	Response.Write "��վά���У����Ժ�"
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
	SearchString ="travelers car insurance|1 loan resource|autoquote insurance|Unknown|None|����|��׬|���ζ���|��Ѵ�绰|Ⱥ��|����|����|��й|���չ���|�ܵ���ͨ|���̼���|��湫˾|����֤��|���k�C��|��ȷ�Ϻ󸶿�|����|�ɷ�ɰ��|֤������|������̳|����|��֤|����|����ϵͳ|��ƾ|��ѵѧУ|�ټ���|������̳|��Ʒ��Ů|�������|������|��ѵ|�ۺ��̳�|��׬|��ǹ|���²�|��ְ|���ϻ���|���˴���|�귢����|����|������Ʊ|Ʊ��|��ˮ��|������|q460985040|���Դ�qq|孙�|����|����|LLL|��ŮͼƬ|1721613566|ȷ��һ�θ߷�ͨ��|˽����̽|378740331|���ȶ�����|95.88.155.48|¡����|GX�˺�|0731988|��������|����|»Ø¸´|��ƭ��|�ѳ�|�ظ�:</font>|回�|�ظ�|kubagouwu|53union|θ��|�̳ǽ���"   ' ���������ַ�����
	SearchStr=split(SearchString,"|")
	SearchChar = zt   ' Ҫ�����ַ��� "P"��
	if len(zt)=8 and IsNumeric(zt) then response.redirect("index.asp")
	if SearchChar="�ظ�:</font>" then response.redirect("index.asp")
	if SearchChar="" then response.redirect("index.asp")
	for i=0 to ubound(SearchStr)
		if Instr(1, SearchChar, SearchStr(i)) then response.redirect("index.asp")
	next
	'if Instr(1, SearchString, SearchChar, 0) then response.write("�������ַ���")
	SearchString ="������|���߷���|��й|Ⱥ��|��Ѵ�绰|��׬|���չ���|��������|����|���|���ڻ�|���濨|��������|OEM|��ʪ��|��Ƶ��|����|�Ž�|������|����|�ɷ�ɰ��|����˫ɫ��|[url=http://|��������|����֤��|����̨|���빫˾|���η���|Viagra|����|������|����|�豸|[URL=http://|[URL]http://|�¾�|÷��|709182385|��̥|nanpuhospital| sex |<a href=|�޷Ĳ�����|.cn|������׬|˽��|»Ø¸´|LLL|brazil.mcneel.com|inthesetimes|www.easyvoa.com|78348|�ڿ�����û��|526053166|810660651|Aloha|&euro;|531498614|��ƭ��|xuniwu|θ��|88952634"   ' ���������ַ�����
	SearchStr=split(SearchString,"|")
	SearchChar = ly
	if SearchChar="�ظ�:</font>" then response.redirect("index.asp")
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
	zh=replace(zh,"����ʦ����-��ͼ&gt;","����ʦ����-��ͼ>")
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
