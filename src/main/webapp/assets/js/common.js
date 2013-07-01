function setHome(obj, url) {
	try {
		obj.style.behavior = 'url(#default#homepage)';
		obj.setHomePage(url);
	} catch (e) {
		if (window.netscape) {
			try {
				netscape.security.PrivilegeManager.enablePrivilege("UniversalXPConnect");
			} catch (e) {
				alert("此操作被浏览器拒绝！\n请在浏览器地址栏输入“about:config”并回车\n然后将 [signed.applets.codebase_principal_support]的值设置为'true',双击即可。");
			}
			var prefs = Components.classes['@mozilla.org/preferences-service;1'].getService(Components.interfaces.nsIPrefBranch);
			prefs.setCharPref('browser.startup.homepage', vrl);
		}
	}
}
function addFavorite(url, name) {
	try {
		window.external.addFavorite(url, name);
	} catch (e) {
		try {
			window.sidebar.addPanel(name, url, "");
		} catch (e) {
			alert("加入收藏失败，请使用Ctrl+D进行添加");
		}
	}
}

function VerifyInput(form1)
{
	if(form1.PWD.value == "")
		{
		alert("您的密码是不是忘填了!!");
		form1.PWD.focus();
		return false;
		} 
	//E-MAIL邮件地址检测
	if (form1.Email){
		if (form1.Email.value == "info@teachersun.com" || form1.Email.value == "slskt@263.net" ){
			alert("请填写您报名注册时填写的的Email地址!!");
			form1.Email.focus();
			return (false);
		}
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

	return true;
}

function playImage(){
	var t = n = 0, count; 
	count=$("#banner_list a").length; 
	$("#banner_list a:not(:first-child)").hide(); 
	$("#banner_info").html($("#banner_list a:first-child").find("img").attr('alt')); 
	$("#banner_info").click(function(){window.open($("#banner_list a:first-child").attr('href'), "_blank")}); 
	$("#banner li").click(function() { 
		var i = $(this).text() - 1;//获取Li元素内的值，即1，2，3，4 
		n = i; 
		if (i >= count) return; 
		$("#banner_info").html($("#banner_list a").eq(i).find("img").attr('alt')); 
		$("#banner_info").unbind().click(function(){window.open($("#banner_list a").eq(i).attr('href'), "_blank")})
		$("#banner_list a").filter(":visible").fadeOut(500).parent().children().eq(i).fadeIn(1000); 
		$(this).css({"background":"#be2424",'color':'#000'}).siblings().css({"background":"#6f4f67",'color':'#fff'}); 
	});
	t = setInterval(function showAuto() 
	{ 
		n = n >=(count - 1) ? 0 : ++n; 
		$("#banner li").eq(n).trigger('click'); 
	} , 4000); 
	$("#banner").hover(function(){clearInterval(t)}, function(){t = setInterval(function showAuto() 
	{ 
		n = n >=(count - 1) ? 0 : ++n; 
		$("#banner li").eq(n).trigger('click'); 
	} , 4000);}); 
}
