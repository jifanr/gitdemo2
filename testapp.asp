<!--#include file="../inc/head_admin.asp" -->
<!--#include file="web_diaoyong/a2.asp"-->
<%
productcode = SafeRequest("productcode")
if isallowbuy(productcode,ErrMsg) = false then PopErr "对不起，您不允许申请该产品"

if SafeRequest("act")="update" then
	call update()
	else
	call list()
end if
%>
<!--#include file="../inc/end_admin.asp"-->

<%sub list()
	set rst=conn.execute("select * from sitetypecfg where sitetypeid='"&productcode&"' and cfgname='netcn_idc_name0' ")
	if not rst.eof then netcn_idc_name0 = rst("cfgvalue")
	rst.close

	set rst=conn.execute("select * from sitetypecfg where sitetypeid='"&productcode&"' and cfgname='netcn_idc_name1' ")
	if not rst.eof then netcn_idc_name1 = rst("cfgvalue")
	rst.close

	set rst=conn.execute("select * from sitetypecfg where sitetypeid='"&productcode&"' and cfgname='netcn_idc_name2' ")
	if not rst.eof then netcn_idc_name2 = rst("cfgvalue")
	rst.close

	set rst=conn.execute("select * from sitetypecfg where sitetypeid='"&productcode&"' and cfgname='netcn_idc_name3' ")
	if not rst.eof then netcn_idc_name3 = rst("cfgvalue")
	rst.close

	set rst=conn.execute("select * from sitetypecfg where sitetypeid='"&productcode&"' and cfgname='netcn_idc_name4' ")
	if not rst.eof then netcn_idc_name4 = rst("cfgvalue")
	rst.close

	set rst=conn.execute("select * from sitetypecfg where sitetypeid='"&productcode&"' and cfgname='netcn_idc_name5' ")
	if not rst.eof then netcn_idc_name5 = rst("cfgvalue")
	rst.close

	set rst=conn.execute("select * from sitetypecfg where sitetypeid='"&productcode&"' and cfgname='netcn_idc_name6' ")
	if not rst.eof then netcn_idc_name6 = rst("cfgvalue")
	rst.close

	set rst=conn.execute("select * from sitetypecfg where sitetypeid='"&productcode&"' and cfgname='netcn_idc_name7' ")
	if not rst.eof then netcn_idc_name7 = rst("cfgvalue")
	rst.close
	'11-08 万商机房设置   增加-->
	set rst=conn.execute("select * from sitetypecfg where sitetypeid='"&productcode&"' and cfgname='netcn_idc_name8' ")
	if not rst.eof then netcn_idc_name8 = rst("cfgvalue")
	rst.close
	set rst=conn.execute("select * from sitetypecfg where sitetypeid='"&productcode&"' and cfgname='netcn_idc_name9' ")
	if not rst.eof then netcn_idc_name9 = rst("cfgvalue")
	rst.close
	set rst=conn.execute("select * from sitetypecfg where sitetypeid='"&productcode&"' and cfgname='netcn_idc_name10' ")
	if not rst.eof then netcn_idc_name10 = rst("cfgvalue")
	rst.close
	set rst=conn.execute("select * from sitetypecfg where sitetypeid='"&productcode&"' and cfgname='netcn_idc_name11' ")
	if not rst.eof then netcn_idc_name11 = rst("cfgvalue")
	rst.close
	set rst=conn.execute("select * from sitetypecfg where sitetypeid='"&productcode&"' and cfgname='netcn_idc_name12' ")
	if not rst.eof then netcn_idc_name12 = rst("cfgvalue")
	rst.close
	set rst=conn.execute("select * from sitetypecfg where sitetypeid='"&productcode&"' and cfgname='netcn_idc_name13' ")
	if not rst.eof then netcn_idc_name13 = rst("cfgvalue")
	rst.close
	set rst=conn.execute("select * from sitetypecfg where sitetypeid='"&productcode&"' and cfgname='netcn_idc_name14' ")
	if not rst.eof then netcn_idc_name14 = rst("cfgvalue")
	rst.close
	set rst=conn.execute("select * from sitetypecfg where sitetypeid='"&productcode&"' and cfgname='netcn_idc_name15' ")
	if not rst.eof then netcn_idc_name15 = rst("cfgvalue")
	rst.close

	set rst=conn.execute("select * from sitetypecfg where sitetypeid='"&productcode&"' and cfgname in('netcn_idc_name16')")
	if not rst.eof then netcn_idc_name16 = rst("cfgvalue")
	rst.close
				
	'11-08 万商机房设置   增加-->
	set rst=conn.execute("select * from sitetypecfg where sitetypeid='"&productcode&"' and cfgname='netcn_idc_os1' ")
	if not rst.eof then netcn_idc_os1 = rst("cfgvalue")
	rst.close

	set rst=conn.execute("select * from sitetypecfg where sitetypeid='"&productcode&"' and cfgname='netcn_idc_os2' ")
	if not rst.eof then netcn_idc_os2 = rst("cfgvalue")
	rst.close
	'10-19 万网祥云主机系统类型添加
	set rst=conn.execute("select * from sitetypecfg where sitetypeid='"&productcode&"' and cfgname='netcn_idc_os3' ")
	if not rst.eof then netcn_idc_os3 = rst("cfgvalue")
	rst.close
	set rst=conn.execute("select * from sitetypecfg where sitetypeid='"&productcode&"' and cfgname='netcn_idc_os4' ")
	if not rst.eof then netcn_idc_os4 = rst("cfgvalue")
	rst.close
	set rst=conn.execute("select * from sitetypecfg where sitetypeid='"&productcode&"' and cfgname='netcn_idc_os5' ")
	if not rst.eof then netcn_idc_os5 = rst("cfgvalue")
	rst.close
	set rst=conn.execute("select * from sitetypecfg where sitetypeid='"&productcode&"' and cfgname='netcn_idc_os6' ")
	if not rst.eof then netcn_idc_os6 = rst("cfgvalue")
	rst.close

    '模板处理开始 ifend 20181205
	Dim moban
	Set moban = server.CreateObject("jf.jf_moban")
	moban.init jfo
	moban.usefile "app.moban.html"
	moban.add "productcode",productcode
	productcode_up = Getproductcode_up(SafeRequest("productcode"))
	set rs = conn.execute ("select * from winiis_product where productcode='" & productcode &"'")
	productname = rs("productname")
	pricetype = print_datalist_opt_name("pricetype",rs ("priceType"))
	servergrouplist = rs("servergrouplist")
	istry=rs("istry")
	levelcode = session("levelcode")
	lpjh = "productcode="& productcode &"&levelcode="& levelcode
	jiage = replace(whois_query(lpjh),vblf,"")
	moban.add "productname",productname
	moban.add  "pricetype",pricetype
	moban.add  "servergrouplist",servergrouplist
	moban.add  "istry",istry
	moban.add  "jiage",jiage
	'moban.update xunhuanming

    'If session("rootusername")= "admin" Then
	Dim jfcoupon 
	set jfcoupon = server.createobject("jf.jf_coupon")
	'jfcoupon.init jfo
	Set rsuco = jfcoupon.getUserCoupon(session("username"),productcode)
	If Not rsuco.eof and session("levelcode")="R004" then
    	moban.add "xian",false
    	Dim ico
    	ico = 0
    	xunhuanming = "quanlist"
    	Do While Not rsuco.eof
        	codt2 = Split(rsuco("codt2")," ")(0)
        	moban.add xunhuanming & ".usercouponId",rsuco("usercouponId") & ""
        	moban.add xunhuanming & ".coupon_jine",rsuco("coupon_jine") & ""
        	moban.add xunhuanming & ".codt2",codt2 & ""
        	moban.add xunhuanming & ".ico",ico & ""
        	moban.update xunhuanming
        	rsuco.movenext:ico = ico+1
    	Loop
    	Set jfcoupon = Nothing
	End If

	If session("isadmin")="Y" And SafeRequest("jfact")="bat" Then
	    moban.add "guanli",false
	End If 
	
	os = request("os")
	'部分产品代表win / linux 是反的
	if inarrayI(lcase(productcode),array("hk10")) <> -1 then 
		ArrayOScode_w = Split("1 w W"," ")
		ArrayOScode_l = Split("0 l L"," ")
		moban.add "xitong",false
		else 
		ArrayOScode_w = Split("0 w W"," ")
		ArrayOScode_l = Split("1 l L"," ")
	end if 
		  
	if os<>"" then 
		moban.add "os1",false
		If inarrayI(os,ArrayOScode_w) <> -1 Then 
			oso = ArrayOScode_w(0)
			moban.add "oso",oso
		ElseIf inarrayI(os,ArrayOScode_l) Then 
			oso = ArrayOScode_l(0)
			moban.add "oso",oso
		end if 
	end if 
		
	if netcn_idc_os1<>"" then
	    moban.add "netcn_idc_os1",false
	end if 
	 
	if netcn_idc_os2<>"" then
	    moban.add "netcn_idc_os2",false
	end if 
	 
	if netcn_idc_os3<>"" then
	    moban.add "netcn_idc_os3",false
	end if 
	   
	if netcn_idc_os4<>"" then
	    moban.add "netcn_idc_os4",false  
	end if 
	   
	if netcn_idc_os5<>"" then
	    moban.add "netcn_idc_os5",false    
	end if 
	     
	if netcn_idc_os6<>"" then
	    moban.add "netcn_idc_os6",false     
	end if 
	      
	If session("isadmin")="Y" And SafeRequest("jfact")="bat" Then
 	else
		moban.add "chktype",false
	End If

	istrue  = netcn_idc_name0<>"" or netcn_idc_name1<>"" or netcn_idc_name2<>"" or netcn_idc_name3<>"" _
	or netcn_idc_name4<>"" or netcn_idc_name5<>"" or netcn_idc_name6<>"" or netcn_idc_name7<>"" _
	Or netcn_idc_name8<>"" or netcn_idc_name9<>"" or netcn_idc_name10<>"" or netcn_idc_name11<>"" _
	or netcn_idc_name12<>"" or netcn_idc_name13<>"" or netcn_idc_name14<>"" _or netcn_idc_name15<>"" _
	Or netcn_idc_name16<>""
	
	if istrue then 
    	if netcn_idc_name0<>"" then
    	    moban.add "netcn_idc_name0",false
    	end if
    	if netcn_idc_name1<>"" then
    	    moban.add "netcn_idc_name1",false
    	end if
    	if netcn_idc_name2<>"" then
    	    moban.add "netcn_idc_name2",false
    	end if
    	if netcn_idc_name3<>"" then
    	    moban.add "netcn_idc_name3",false
    	end if
    	if netcn_idc_name4<>"" then
    	    moban.add "netcn_idc_name4",false
    	end if
    	if netcn_idc_name5<>"" then
    	    moban.add "netcn_idc_name5",false
    	end if
    	if netcn_idc_name6<>"" then
    	    moban.add "netcn_idc_name6",false
    	end if
    	if netcn_idc_name7<>"" then
    	    moban.add "netcn_idc_name7",false
    	end if
    	if netcn_idc_name8<>"" then
    	    moban.add "netcn_idc_name8",false
    	end if
    	if netcn_idc_name9<>"" then
    	    moban.add "netcn_idc_name9",false
    	end if
    	if netcn_idc_name10<>"" then
    	    moban.add "netcn_idc_name10",false
    	end if
    	if netcn_idc_name11<>"" then
    	    moban.add "netcn_idc_name11",false
    	end if
    	if netcn_idc_name12<>"" then
    	    moban.add "netcn_idc_name12",false
    	end if
    	if netcn_idc_name13<>"" then
    	    moban.add "netcn_idc_name13",false
    	end if
    	if netcn_idc_name14<>"" then
    	    moban.add "netcn_idc_name14",false
    	end if
    	if netcn_idc_name15<>"" then
    	    moban.add "netcn_idc_name15",false
    	end if
    	if netcn_idc_name16<>"" then
    	    moban.add "netcn_idc_name16",false
    	end if
	end if

	If pricetype="月" and 1=2 Then 
    	moban.add "xufeitixing",false
    	If SafeRequest("hkprice")&""<>"" Then 
    		moban.add "hkprice",SafeRequest("hkprice")
    	Else
    	End If
	End If 
	if Gsession("syscfg-onlyorder")="Y" Or session("isadmin")="Y" then
	    moban.add "xiadanmoshi",false
	end if
	'模板处理结束 ifend 20181205
	moban.output
end sub

sub update()
	usercouponId = trim(SafeRequest("usercouponId"))
	usercouponprice = 0
	If isnumeric(usercouponId) then
		Dim jfcoupon 
		set jfcoupon = server.createobject("jf.jf_coupon")
		jfcoupon.init jfo
		Set rsuco = jfcoupon.getUserCoupon(session("username"),productcode,usercouponId)
		If rsuco.eof Then 
			usercouponprice = 0
			usercouponId = ""
		Else
			usercouponprice = rsuco("coupon_jine")
		End If
		rsuco.close
	End If 
	if session("rootisadmin") = "Y" then 
		response.write "usercouponId:" & usercouponId & "<hr>"
		response.write "usercouponprice:" & usercouponprice & "<hr>"
		response.end
	end if 
	isbat = SafeRequest("isbat")
	Dim thei,endi,dom1,dom2
	If isbat="1" and session("isadmin")="Y" Then 
		server.ScriptTimeout=1800	
		rute = "goodsbus"
		domainname = LCase(SafeRequest("domainname"))
		pos1 = InStr(domainname,"{")
		pos2 = InStr(domainname,"}")
		iii = Mid(domainname,pos1+1,pos2-pos1-1)
		dom1 = Mid(domainname,1,pos1-1)
		dom2 = Mid(domainname,pos2+1)
		iiiarr = split(iii,"-")
		thei = cint(iiiarr(0))
		endi = cint(iiiarr(1))
	Else
        thei=0
        endi=0
		domainname = LCase(SafeRequest("domainname"))
		rute = SafeRequest("rute")
	End If 
	Dim i
	For i = thei To endi
		productcode=SafeRequest("productcode")
		If isbat="1" and session("isadmin")="Y" Then 
			domainname = dom1 & i & dom2
		Else
			domainname = LCase(SafeRequest("domainname"))
		End If 
		ret = CheckDomainPX(domainname)
		if ret <> "" then PopErr ret
		orderid = CreateUID(productcode,"主机")
		applytime=SafeRequest("applytime")
		idc=SafeRequest("idc")
		'Response.Write idc & "<br />"
		os=SafeRequest("os")
		if cint(applytime)<1 then PopErr "请申请时间不能小于1"
		set rst=conn.execute ("select * from winiis_allproduct where domainname='"&domainname&"' and productcode='"&productcode&"' ")
		if not rst.eof then
			PopErr "该域名的站点已存在,请绑定别的域名"
		end if
		rst.close
		set rst=conn.execute ("select * from winiis_web where websitename='"&domainname&"' ")
		if not rst.eof then
			PopErr "该域名的站点已存在,请绑定别的域名"
		end if
		rst.close
		set rst=conn.execute ("select * from winiis_ftp where ftpuser='"&ftpaccount&"' ")
		if not rst.eof then
			PopErr "该FTP帐号已存在,请选择别的FTP帐号名"
		end if
		rst.close
		set rsp = conn.execute("select * from winiis_product where productcode='"&productcode&"'")
		if rsp.eof then
			PopErr "产品类型["&productcode&"不存在"
		end if
		pricetype=rsp("pricetype")
		istry_p=rsp("istry")
		rsp.close
		'需要增加字段校验信息
		set rs=server.createobject("adodb.recordset")
		sql="select * from winiis_orderlist"
		rs.open sql,conn,1,3
		rs.addnew
		rs("OrderID") = OrderID
		rs("username")=session("username")
		if istry="Y" then
			rs("OrderType")="try"
			rs("price")=0
			rs("PayStatus")="Y"
		else
			rs("OrderType")="add"
			price = GetProductPriceByRegType(session("username"),productcode,applytime,"") 
			If usercouponId<>"" Then 
				rs("usercouponId") = usercouponId
			End If 
			If usercouponprice > 0 Then 
				If price >= usercouponprice Then 
					price= price - usercouponprice 
				Else
					msg = "优惠券金额大于订单金额，不能使用"
				End If
			End If 
			rs("price")=price
			rs("PayStatus")="N"
		end if
		rs("domainname") = domainname
		rs("productcode")= productcode
		rs("applytime")=applytime
		rs("pricetype")=pricetype
		'rs("PayStatus")="N"
		rs("HandleStatus")="N"
		rs("regtime")=now()
		rs("remark")=SafeRequest("remark")
		rs.update
		rs.close
		set rs=Nothing
		If usercouponId<>"" then
			conn.Execute "update winiis_user_coupon set useStatus='Y' where usercouponId=" & usercouponId
		End If
		sql = " insert into winiis_orderparam (OrderID,ParamName,ParamValue,ParamDesc) values ('"&OrderID&"','idc','"&idc&"','机房位置') "
		conn.execute (sql)
		sql =" insert into winiis_orderparam (OrderID,ParamName,ParamValue,ParamDesc) values ('"&OrderID&"','os','"&os&"','操作系统') "
		conn.execute (sql)
		sql =" insert into winiis_orderparam (OrderID,ParamName,ParamValue,ParamDesc) values ('"&OrderID&"','Domain','"&domainname&"','域名') "
		conn.execute (sql)
	next
	'如果会员帐上金额足够,则自动受理,并开通
	'call OrderHandle(OrderID)
	'添加订单不实时开通,张,2011-11-25
	if rute = "goodsbus" then
		response.write "<script>alert('该产品已成功添加到订单中');location.href='../userself/userinform.asp';</script>"
	else
	    %>
		<form name="orderform" action="../reg/Handle.asp" method="post">
		<input type="hidden" name="OrderID" id="OrderID" value="<%=OrderID%>" />
		</form>
		<script language="javascript">document.orderform.submit();</script>
	    <%
	end if
end Sub
%>