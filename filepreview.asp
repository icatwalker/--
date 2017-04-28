<%
'从文件中获取显示数据
function GetContent(filename)
   dim fs
   dim thisfile
   dim rdstr
   dim fullfilename
   rdstr=""
   if (left(filename,1)="\" or left(filename,1)="/") then
      fullfilename=Application("WebRoot") & filename
   else
      fullfilename=Application("WebRoot") & "\" & filename
   end if
   on error resume next
   set fs=createobject("scripting.filesystemobject")
   set thisfile=fs.opentextfile(fullfilename)
   rdstr=thisfile.readall
   thisfile.close
   GetContent=rdstr
end function

    logtablename="customs"
    '预览素材
    dim sql
    dim dat
	dim itemid
    dim templateid
	dim bkpic
	dim bkclr
	dim fontclr
	dim fontoption
	dim width
	dim height
	dim align
	dim alignstr
	dim circle
	dim circlestr
	dim delay
	dim delaystr
	dim scrollunit
	dim scrollunitstr
	dim fontname
	dim fontsize
	dim fontsizestr
	dim fontbold
	dim fontbold1
	dim fontitalic
	dim fontitalic1

	dim itemname
	dim contenttype
	dim path
	dim descript
	dim outdescript
	dim getfilecontent
	dim itemtype
	dim dispmsg
    dim charsetstr
	dim ttempurl
	dim tt
	dim loc

	charsetstr="<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"

    dispmsg=""
	bkpic=""
	bkclr="#0000FF"
	fontclr="#FFFFFF"
	fontoption="24"
	fontname=""
	fontsize=24
	fontsizestr=""
	fontbold=""
	fontbold1=""
	fontitalic=""
	fontitalic1=""
	width="100%"
	height="100%"
	align=1
	alignstr="align='center'"
	circle=1
	circlestr="loop=""-1"""
	delay=15
	delaystr="scrollDelay=""15"""
	scrollunit=1
    scrollunitstr="scrollAmount=""1"""

    itemtype=0
	if trim(request.querystring("itemtype"))<>"" then
	   itemtype=cint(trim(request.querystring("itemtype"))) mod 100
	end if
	itemid=trim(request.querystring("itemid"))
	if (itemid="") then
	   response.write "给出的素材标记有误！"
	   response.end
	end if

    templateid=trim(request.querystring("templateid"))
	outdescript=trim(request.querystring("descript"))
	if (outdescript<>"") then
	    if instr(outdescript,"&lt;")>0 then
			outdescript=replace(outdescript,"&lt;","<")
			outdescript=replace(outdescript,"&gt;",">")
		end if
	else
	end if
	set dat=server.CreateObject("ADODB.recordset")
    dat.CursorLocation=3
    set dat.ActiveConnection=Session("dbConnect")

    '首先得到模板信息
	if (templateid<>"0" and templateid<>"") then
       sql="select * from disptemplate where templateid=" & templateid
	   dat.open sql
	   if (not dat.eof) then
			bkpic=trim(dat("bkpic").value)
			bkclr="#" & right("000000" & Hex(dat("bkclr").value),6)
			fontclr="#" & right("000000" & Hex(dat("fontclr").value),6)
			fontoption=trim(dat("fontoption").value)
			loc=instr(fontoption,";")
			if (loc>0) then
			   fontsize=left(fontoption,loc-1)
			   if fontsize<>"" then
				   fontsize=clng(fontsize)
			   end if
			   fontoption=right(fontoption,len(fontoption)-loc)
			   loc=instr(fontoption,";")
			   if (loc>0) then
			      fontname=left(fontoption,loc-1)
			   else
			      fontname=fontoption
			   end if
			else
				if fontoption<>"" then
				   fontsize=clng(fontoption)
				end if
			end if
            if (instr(fontoption,",b,")>0) or (instr(fontoption,",B,")>0) then
			   fontbold="<B>"
			   fontbold1="</B>"
			end if
            if (instr(fontoption,",i,")>0) or (instr(fontoption,",I,")>0) then
			   fontitalic="<I>"
			   fontitalic1="</I>"
			end if
			width=trim(dat("width").value)
			height=trim(dat("height").value)
			align=clng(dat("align").value)
			if (align=0) then
			   alignstr="align='left'"
			elseif(align=2) then
			   alignstr="align='right'"
			end if
			circle=clng(dat("circle").value)
			if (circle=0) then
			   circlestr="loop=""1"""
			end if
			delay=clng(dat("delay").value)
			delaystr="scrollDelay=""" & delay & """"
			scrollunit=clng(dat("scrollunit").value)
			scrollunitstr="scrollAmount=""" & scrollunit & """"
	   end if
	   dat.close
	end if

    '建立数据库查询sql语句
    'sql="select itemname,itemid,contenttype,path,descript,itemtype from " & session(logtablename&"companyid") &"_menuitem where itemtype=" & itemtype & " and itemid=" & itemid
	sql="select itemname,itemid,contenttype,path,descript,itemtype from " & session(logtablename&"companyid") &"_menuitem where itemid=" & itemid
	dat.open sql
	if (not dat.eof) then
		itemname=trim(dat("itemname").value)
		itemtype=cint(dat("itemtype").value)
		if itemtype=1 then
		   contenttype=14
		else
		   contenttype=clng(dat("contenttype").value)
		end if
		path=trim(dat("path").value)
		descript=trim(dat("descript").value)
		if (descript="") or isnull(descript) then
		   descript=outdescript
		   if (descript="") then
		      descript=itemname
		   end if
		else
			if (descript<>"") then
				if instr(descript,"&lt;")>0 then
					descript=replace(descript,"&lt;","<")
					descript=replace(descript,"&gt;",">")
				end if
			end if
		end if
		dat.close
		set dat=nothing
		'生成对应的内容
		'第一行输出信息
		'response.write "<b><font color=white>素材名：" & itemname & "&nbsp;&nbsp;&nbsp;描述：" & descript & "&nbsp;&nbsp;&nbsp;类型："
		dispmsg="素材名：" & itemname & "\r\n描述：" & descript & "\r\n类型："
		'0-自适应,1-文本,2-网页,3-图片,4-通知(静态),5-通知(向上滚动文本),6-字幕(向左滚动文本),7-动画,8-Office文稿,9-音频,10-视频文件/网络视频/电视,11-操作系统自检测,12-专用应用程序...,13-远程命令,

		if (contenttype=12) then
		'对于12类型的，专门处理以下
		   dispmsg=dispmsg & "12-专用应用程序 "
		   tt=lcase(path)
		   if (instr(tt,".swf")>0) then
		      contenttype=7
		   elseif ((instr(tt,".ppt")>0) or (instr(tt,".pps")>0) or (instr(tt,".doc")>0)) then
		      contenttype=8
		   elseif (instr(tt,".txt")>0) then
		      contenttype=1
		   elseif ((instr(tt,".htm")>0) or (instr(tt,".asp")>0) or (instr(tt,".php")>0) or (instr(tt,".java")>0)) then
		      contenttype=2
		   elseif ((instr(tt,".bmp")>0) or (instr(tt,".jpg")>0) or (instr(tt,".gif")>0) or (instr(tt,".pnp")>0)) then
		      contenttype=3
		   elseif ((instr(tt,".asf")>0) or (instr(tt,".wmv")>0) or (instr(tt,".wma")>0) or (instr(tt,".mpg")>0) or (instr(tt,".avi")>0) or (instr(tt,".rm")>0) or (instr(tt,".dat")>0) or (instr(tt,".vob")>0)) then
              contenttype=10
		   end if
		end if
		'处理path,去掉其中的命令行参数部分
		'从后面查找 空格-
        loc=instr(path," -")
		if (loc>0) then
		   path=left(path,loc-1)
		end if

		Select Case contenttype
		case 0
			'response.write "0-自适应</font></b><br>"
			dispmsg=dispmsg & "0-自适应"
   			response.write "<html>" & chr(13) & chr(10)
   			response.write "<head>" & chr(13) & chr(10)
   			response.write charsetstr & chr(13) & chr(10)
   			response.write "<title>Digital Multi-Media Distributing System</title>" & chr(13) & chr(10)
   			response.write "<script language=""javascript"">" & chr(13) & chr(10)
   			response.write "<!--" & chr(13) & chr(10)
   			response.write "function displayinfo(dispstr)" & chr(13) & chr(10)
   			response.write "{" & chr(13) & chr(10)
   			response.write "   alert(dispstr);" & chr(13) & chr(10)
   			response.write "}" & chr(13) & chr(10)
   			response.write "-->" & chr(13) & chr(10)
   			response.write "</script>" & chr(13) & chr(10)
   			response.write "</head>" & chr(13) & chr(10)
   			response.write "<body background=""" & bkpic & """ bgcolor=""" & bkclr & """ ondblclick=""displayinfo('" & dispmsg & "');"">" & chr(13) & chr(10)
   			response.write "<table height=""100%"" width=""100%"">" & chr(13) & chr(10)
   			response.write "<tr height=""100%"" width=""100%"">" & chr(13) & chr(10)
   			response.write "<td valign=""middle""><p " & alignstr & "><span><font color=""" & fontclr & """ style=""font-size:" & fontsize & "pt"" face=""" & fontname & """ style=""line-height: 130%"">" & fontbold & fontitalic & descript & fontbold1 & fontitalic1 & "</font></span></p></td>" & chr(13) & chr(10)
   			response.write "</tr>" & chr(13) & chr(10)
   			response.write "</table>" & chr(13) & chr(10)
   			response.write "</body>" & chr(13) & chr(10)
   			response.write "</html>" & chr(13) & chr(10)
		case 1,14
			'response.write "1-文本</font></b><br>"
			if itemtype=1 then
			   dispmsg=dispmsg & "14-栏目"
			else
			   dispmsg=dispmsg & "1-文本"
			end if
			if (path<>"") then
			   getfilecontent=GetContent(path)
			end if
			if (getfilecontent<>"") then
			   descript=getfilecontent
			end if
			response.write "<html>" & chr(13) & chr(10)
			response.write "<head>" & chr(13) & chr(10)
			response.write charsetstr & chr(13) & chr(10)
			response.write "<title>Digital Multi-Media Distributing System</title>" & chr(13) & chr(10)
   			response.write "<script language=""javascript"">" & chr(13) & chr(10)
   			response.write "<!--" & chr(13) & chr(10)
   			response.write "function displayinfo(dispstr)" & chr(13) & chr(10)
   			response.write "{" & chr(13) & chr(10)
   			response.write "   alert(dispstr);" & chr(13) & chr(10)
   			response.write "}" & chr(13) & chr(10)
   			response.write "-->" & chr(13) & chr(10)
   			response.write "</script>" & chr(13) & chr(10)
			response.write "</head>" & chr(13) & chr(10)
			response.write "<body background=""" & bkpic & """ bgcolor=""" & bkclr & """ ondblclick=""displayinfo('" & dispmsg & "');"">" & chr(13) & chr(10)
			response.write "<table height=""100%"" width=""100%"">" & chr(13) & chr(10)
			response.write "<tr height=""100%"" width=""100%"">" & chr(13) & chr(10)
			response.write "<td valign=""middle"" " & alignstr & "><span><font color=""" & fontclr & """ style=""font-size:" & fontsize & "pt"" face=""" & fontname & """ style=""line-height: 130%"">" & fontbold & fontitalic & descript & fontbold1 & fontitalic1 & "</font></span></td>" & chr(13) & chr(10)
			response.write "</tr>" & chr(13) & chr(10)
			response.write "</table>" & chr(13) & chr(10)
			response.write "</body>" & chr(13) & chr(10)
			response.write "</html>" & chr(13) & chr(10)
		case 2
			'response.write "2-网页</font></b><br>"
			dispmsg=dispmsg & "2-网页"
			'直接转走
			if (path<>"") then
				response.Redirect(path)
				response.end
			end if
		case 3
			'response.write "3-图片</font></b><br>"
			dispmsg=dispmsg & "3-图片"
			response.write "<html>" & chr(13) & chr(10)
			response.write "<head>" & chr(13) & chr(10)
			response.write  charsetstr & chr(13) & chr(10)
			response.write "<title>Digital Multi-Media Distributing System</title>" & chr(13) & chr(10)
   			response.write "<script language=""javascript"">" & chr(13) & chr(10)
   			response.write "<!--" & chr(13) & chr(10)
   			response.write "function displayinfo(dispstr)" & chr(13) & chr(10)
   			response.write "{" & chr(13) & chr(10)
   			response.write "   alert(dispstr);" & chr(13) & chr(10)
   			response.write "}" & chr(13) & chr(10)
   			response.write "-->" & chr(13) & chr(10)
   			response.write "</script>" & chr(13) & chr(10)
			response.write "</head>" & chr(13) & chr(10)
			response.write "<body background=""" & bkpic & """ bgcolor=""" & bkclr & """ leftMargin=""0"" topMargin=""0"" rightMargin=""0"" bottomMargin=""0"" ondblclick=""displayinfo('" & dispmsg & "');"">" & chr(13) & chr(10)
      '2010-01-16增加div叠加显示功能
      if (mid(descript,1,1)="@") then
      '后面的一个字节数字给出开始位置在整个高度的十分之几0表示最顶,1表示十分之一高度
         ttempurl="top='5'"
         if (mid(descript,2,1)="#") then
            if (mid(descript,3,1)>="0" and mid(descript,3,1)<="9") then
               ttempurl="top='" & CInt(mid(descript,3,1))*10 & "%'"
               tt=right(descript,len(descript)-3)
            else
               tt=right(descript,len(descript)-2)
			end if
         else
            tt=right(descript,len(descript)-1)
         end if
         response.write "<div id=""mypic"" " & alignstr & " style=""position:absolute;" & ttempurl & ";width='100%';height='100%';""><font color=""" & fontclr & """ style=""font-size:" & fontsize & "pt"" face=""" & fontname & """ style=""line-height: 130%"">" & fontbold & fontitalic & tt & fontbold1 & fontitalic1 & "</font></div>" & chr(13) & chr(10)
      end if

			response.write "<center>" & chr(13) & chr(10)
			if (itemtype=0) then
			   response.write "<p " & alignstr & "><img border=""0"" src=""" & path & """></p>" & chr(13) & chr(10)
			else
			   response.write "<p " & alignstr & "><img border=""0"" src=""" & path & """ width=""" & width & """ height=""" & height & """></p>" & chr(13) & chr(10)
			end if
			response.write "</center>" & chr(13) & chr(10)
			response.write "</body>" & chr(13) & chr(10)
			response.write "</html>" & chr(13) & chr(10)
		case 4
			'response.write "4-通知(静态)</font></b><br>"
			dispmsg=dispmsg & "4-通知(静态)"
			if (path<>"") then
			   getfilecontent=GetContent(path)
			end if
			if (getfilecontent<>"") then
			   descript=getfilecontent
			end if
			response.write "<html>" & chr(13) & chr(10)
			response.write "<head>" & chr(13) & chr(10)
			response.write charsetstr & chr(13) & chr(10)
			response.write "<title>Digital Multi-Media Distributing System</title>" & chr(13) & chr(10)
   			response.write "<script language=""javascript"">" & chr(13) & chr(10)
   			response.write "<!--" & chr(13) & chr(10)
   			response.write "function displayinfo(dispstr)" & chr(13) & chr(10)
   			response.write "{" & chr(13) & chr(10)
   			response.write "   alert(dispstr);" & chr(13) & chr(10)
   			response.write "}" & chr(13) & chr(10)
   			response.write "-->" & chr(13) & chr(10)
   			response.write "</script>" & chr(13) & chr(10)
			response.write "</head>" & chr(13) & chr(10)
			response.write "<body background=""" & bkpic & """ bgcolor=""" & bkclr & """ ondblclick=""displayinfo('" & dispmsg & "');"">" & chr(13) & chr(10)
			response.write "<center>" & chr(13) & chr(10)
			response.write "<p " & alignstr & "><span><font color=""" & fontclr & """ style=""font-size:" & fontsize & "pt"" face=""" & fontname & """ style=""line-height: 130%"">" & fontbold & fontitalic & descript & fontbold1 & fontitalic1 & "</font></span></p>" & chr(13) & chr(10)
			response.write "</center>" & chr(13) & chr(10)
			response.write "</body>" & chr(13) & chr(10)
			response.write "</html>" & chr(13) & chr(10)
		case 5
			'response.write "5-通知(向上滚动文本)</font></b><br>"
			dispmsg=dispmsg & "5-通知(向上滚动文本)"
			if (path<>"") then
			   getfilecontent=GetContent(path)
			end if
			if (getfilecontent<>"") then
			   descript=getfilecontent
			end if
			response.write "<html>" & chr(13) & chr(10)
			response.write "<head>" & chr(13) & chr(10)
			response.write charsetstr & chr(13) & chr(10)
			response.write "<title>Digital Multi-Media Distributing System</title>" & chr(13) & chr(10)
   			response.write "<script language=""javascript"">" & chr(13) & chr(10)
   			response.write "<!--" & chr(13) & chr(10)
   			response.write "function displayinfo(dispstr)" & chr(13) & chr(10)
   			response.write "{" & chr(13) & chr(10)
   			response.write "   alert(dispstr);" & chr(13) & chr(10)
   			response.write "}" & chr(13) & chr(10)
   			response.write "-->" & chr(13) & chr(10)
   			response.write "</script>" & chr(13) & chr(10)
			response.write "</head>" & chr(13) & chr(10)
			response.write "<body background=""" & bkpic & """ bgcolor=""" & bkclr & """ ondblclick=""displayinfo('" & dispmsg & "');"">" & chr(13) & chr(10)
			response.write "<center>" & chr(13) & chr(10)
			response.write "<p>" & chr(13) & chr(10)            
			'response.write "<table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%"" height=""100%""><tr height=""100%"" width=""100%""><td valign=""middle"" " & alignstr & ">" & chr(13) & chr(10)
			response.write "<table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%"" height=""100%""><tr height=""100%"" width=""100%""><td valign=""middle"" align=""center"">" & chr(13) & chr(10)
			if (itemtype=0) then
   			   response.write "<MARQUEE direction=""up"" " & scrollunitstr & " " & delaystr & ">" & chr(13) & chr(10)
			   response.write "<table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%"" height=""100%""><tr height=""100%"" width=""100%""><td valign=""middle"" " & alignstr & ">" & chr(13) & chr(10)
			else
   			   response.write "<MARQUEE direction=""up"" " & scrollunitstr & " " & delaystr & " width=""" & width & """ height=""" & height & """>" & chr(13) & chr(10)
			   response.write "<table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%"" height=""100%""><tr height=""100%"" width=""100%""><td valign=""middle"" " & alignstr & " width=""" & width & """ height=""" & height & """>" & chr(13) & chr(10)
			end if
			response.write "<span><font color=""" & fontclr & """ style=""font-size:" & fontsize & "pt"" face=""" & fontname & """ style=""line-height: 130%"">" & fontbold & fontitalic & descript & fontbold1 & fontitalic1 & "</font></span>" & chr(13) & chr(10)			
			response.write "</td></tr></table>" & chr(13) & chr(10)
			response.write "</MARQUEE>" & chr(13) & chr(10)
			response.write "</td></tr></table>" & chr(13) & chr(10)
			response.write "</p>" & chr(13) & chr(10)
			response.write "</center>" & chr(13) & chr(10)
			response.write "</body>" & chr(13) & chr(10)
			response.write "</html>" & chr(13) & chr(10)
		case 6
			'response.write "6-字幕(向左滚动文本)</font></b><br>"
			dispmsg=dispmsg & "6-字幕(向左滚动文本)"
			if (path<>"") then
			   getfilecontent=GetContent(path)
			end if
			if (getfilecontent<>"") then
			   descript=getfilecontent
			end if
			response.write "<html>" & chr(13) & chr(10)
			response.write "<head>" & chr(13) & chr(10)
			response.write charsetstr & chr(13) & chr(10)
			response.write "<title>Digital Multi-Media Distributing System</title>" & chr(13) & chr(10)
   			response.write "<script language=""javascript"">" & chr(13) & chr(10)
   			response.write "<!--" & chr(13) & chr(10)
   			response.write "function displayinfo(dispstr)" & chr(13) & chr(10)
   			response.write "{" & chr(13) & chr(10)
   			response.write "   alert(dispstr);" & chr(13) & chr(10)
   			response.write "}" & chr(13) & chr(10)
   			response.write "-->" & chr(13) & chr(10)
   			response.write "</script>" & chr(13) & chr(10)
			response.write "</head>" & chr(13) & chr(10)
			response.write "<body background=""" & bkpic & """ bgcolor=""" & bkclr & """ ondblclick=""displayinfo('" & dispmsg & "');"">" & chr(13) & chr(10)
			response.write "<center>" & chr(13) & chr(10)
			response.write "<p>" & chr(13) & chr(10)
			response.write "<table border=""0"" width=""100%"" height=""100%""><tr width=""99%"" heigth=""99%""><td valign=""middle"" " & alignstr & ">" & chr(13) & chr(10)
			if (itemtype=0) then
   			   response.write "<MARQUEE direction=""left"" " & scrollunitstr & " " & delaystr & ">" & chr(13) & chr(10)
			else
   			   response.write "<MARQUEE direction=""left"" " & scrollunitstr & " " & delaystr & " width=""" & width & """ height=""" & height & """>" & chr(13) & chr(10)
			end if
			response.write "<span><font color=""" & fontclr & """ style=""font-size:" & fontsize & "pt"" face=""" & fontname & """ style=""line-height: 130%"">" & fontbold & fontitalic & descript & fontbold1 & fontitalic1 & "</font></span>" & chr(13) & chr(10)
			response.write "</MARQUEE>" & chr(13) & chr(10)
			response.write "</td></tr></table>" & chr(13) & chr(10)
			response.write "</p>" & chr(13) & chr(10)
			response.write "</center>" & chr(13) & chr(10)
			response.write "</body>" & chr(13) & chr(10)
			response.write "</html>" & chr(13) & chr(10)
		case 7
			'response.write "7-动画</font></b><br>"
			dispmsg=dispmsg & "7-动画"
			response.write "<html>" & chr(13) & chr(10)
			response.write "<head>" & chr(13) & chr(10)
			response.write charsetstr & chr(13) & chr(10)
			response.write "<title>Digital Multi-Media Distributing System</title>" & chr(13) & chr(10)
   			response.write "<script language=""javascript"">" & chr(13) & chr(10)
   			response.write "<!--" & chr(13) & chr(10)
   			response.write "function displayinfo(dispstr)" & chr(13) & chr(10)
   			response.write "{" & chr(13) & chr(10)
   			response.write "   alert(dispstr);" & chr(13) & chr(10)
   			response.write "}" & chr(13) & chr(10)
   			response.write "-->" & chr(13) & chr(10)
   			response.write "</script>" & chr(13) & chr(10)
			response.write "</head>" & chr(13) & chr(10)
			response.write "<body background=""" & bkpic & """ bgcolor=""" & bkclr & """ leftMargin=""0"" topMargin=""0"" rightMargin=""0"" bottomMargin=""0"" ondblclick=""displayinfo('" & dispmsg & "');"">" & chr(13) & chr(10)
      '2010-01-16增加div叠加显示功能
      if (mid(descript,1,1)="@") then
      '后面的一个字节数字给出开始位置在整个高度的十分之几0表示最顶,1表示十分之一高度
         ttempurl="top='5'"
         if (mid(descript,2,1)="#") then
            if (mid(descript,3,1)>="0" and mid(descript,3,1)<="9") then
               ttempurl="top='" & CInt(mid(descript,3,1))*10 & "%'"
               tt=right(descript,len(descript)-3)
            else
               tt=right(descript,len(descript)-2)
			end if
         else
            tt=right(descript,len(descript)-1)
         end if
         response.write "<div id=""myflash"" " & alignstr & " style=""position:absolute;" & ttempurl & ";width='100%';height='100%';""><font color=""" & fontclr & """ style=""font-size:" & fontsize & "pt"" face=""" & fontname & """ style=""line-height: 130%"">" & fontbold & fontitalic & tt & fontbold1 & fontitalic1 & "</font></div>" & chr(13) & chr(10)
      end if

			response.write "<center>" & chr(13) & chr(10)
			response.write "<div valign=""middle"" " & alignstr & ">" & chr(13) & chr(10)
			if (itemtype=0) then
			   response.write "<object classid=""clsid:D27CDB6E-AE6D-11cf-96B8-444553540000"" codebase=""http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0"" width=""100%"" height=""100%"">" & chr(13) & chr(10)
			else
			   response.write "<object classid=""clsid:D27CDB6E-AE6D-11cf-96B8-444553540000"" codebase=""http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0"" width=""" & width & """ height=""" & height & """>" & chr(13) & chr(10)
			end if
			response.write "<param name=""movie"" value=""" & path & """>" & chr(13) & chr(10)
			response.write "<param name=""quality"" value=""high"">" & chr(13) & chr(10)
			response.write "<param name=""wmode"" value=""transparent"">" & chr(13) & chr(10)
			response.write "<embed quality=""high"" pluginspage=""http://www.macromedia.com/go/getflashplayer"" type=""application/x-shockwave-flash""></embed>" & chr(13) & chr(10)
			response.write "</object>" & chr(13) & chr(10)
			response.write "</div>" & chr(13) & chr(10)
			response.write "</center>" & chr(13) & chr(10)
			response.write "</body>" & chr(13) & chr(10)
			response.write "</html>" & chr(13) & chr(10)
		case 8
			'response.write "8-Office文稿</font></b><br>"
			dispmsg=dispmsg & "8-Office文稿"
			'直接转走
			if (path<>"") then
				response.Redirect(path)
				response.end
			end if
		case 9
			'response.write "9-音频</font></b><br>"
			dispmsg=dispmsg & "9-音频"
			response.write "<html>" & chr(13) & chr(10)
			response.write "<head>" & chr(13) & chr(10)
			response.write charsetstr & chr(13) & chr(10)
			response.write "<title>Digital Multi-Media Distributing System</title>" & chr(13) & chr(10)
   			response.write "<script language=""javascript"">" & chr(13) & chr(10)
   			response.write "<!--" & chr(13) & chr(10)
   			response.write "function displayinfo(dispstr)" & chr(13) & chr(10)
   			response.write "{" & chr(13) & chr(10)
   			response.write "   alert(dispstr);" & chr(13) & chr(10)
   			response.write "}" & chr(13) & chr(10)
   			response.write "-->" & chr(13) & chr(10)
   			response.write "</script>" & chr(13) & chr(10)
			response.write "</head>" & chr(13) & chr(10)
			response.write "<body background=""" & bkpic & """ bgcolor=""" & bkclr & """ leftMargin=""0"" topMargin=""0"" rightMargin=""0"" bottomMargin=""0"" ondblclick=""displayinfo('" & dispmsg & "');"">" & chr(13) & chr(10)
			response.write "<center>" & chr(13) & chr(10)
			if (itemtype=0) then
			   response.write "<p valign=""middle"" " & alignstr & "><img border=""0"" dynsrc=""" & path & """ start=""fileopen"" " & circlestr & "></p>" & chr(13) & chr(10)
			else
			   response.write "<p valign=""middle"" " & alignstr & "><img border=""0"" dynsrc=""" & path & """ width=""" & width & """ height=""" & height & """ start=""fileopen"" " & circlestr & "></p>" & chr(13) & chr(10)
			end if
			response.write "</center>" & chr(13) & chr(10)
			response.write "</body>" & chr(13) & chr(10)
			response.write "</html>" & chr(13) & chr(10)
		case 10
			'response.write "10-视频文件/网络视频/电视</font></b><br>"
			dispmsg=dispmsg & "10-视频文件/网络视频/电视"
			response.write "<html>" & chr(13) & chr(10)
			response.write "<head>" & chr(13) & chr(10)
			response.write charsetstr & chr(13) & chr(10)
			response.write "<title>Digital Multi-Media Distributing System</title>" & chr(13) & chr(10)
   			response.write "<script language=""javascript"">" & chr(13) & chr(10)
   			response.write "<!--" & chr(13) & chr(10)
   			response.write "function displayinfo(dispstr)" & chr(13) & chr(10)
   			response.write "{" & chr(13) & chr(10)
   			response.write "   alert(dispstr);" & chr(13) & chr(10)
   			response.write "}" & chr(13) & chr(10)
   			response.write "-->" & chr(13) & chr(10)
   			response.write "</script>" & chr(13) & chr(10)
			response.write "</head>" & chr(13) & chr(10)
			response.write "<body background=""" & bkpic & """ bgcolor=""" & bkclr & """ leftMargin=""0"" topMargin=""0"" rightMargin=""0"" bottomMargin=""0"" ondblclick=""displayinfo('" & dispmsg & "');"">" & chr(13) & chr(10)
      '2010-01-16增加div叠加显示功能
      if (mid(descript,1,1)="@") then
      '后面的一个字节数字给出开始位置在整个高度的十分之几0表示最顶,1表示十分之一高度
         ttempurl="top='5'"
         if (mid(descript,2,1)="#") then
            if (mid(descript,3,1)>="0" and mid(descript,3,1)<="9") then
               ttempurl="top='" & CInt(mid(descript,3,1))*10 & "%'"
               tt=right(descript,len(descript)-3)
            else
               tt=right(descript,len(descript)-2)
			end if
         else
            tt=right(descript,len(descript)-1)
         end if
         response.write "<div id=""myvideo"" " & alignstr & " style=""position:absolute;" & ttempurl & ";width='100%';height='100%';""><font color=""" & fontclr & """ style=""font-size:" & fontsize & "pt"" face=""" & fontname & """ style=""line-height: 130%"">" & fontbold & fontitalic & tt & fontbold1 & fontitalic1 & "</font></div>" & chr(13) & chr(10)
      end if

			response.write "<center>" & chr(13) & chr(10)
           okurl=path
           dealit=false%>
	   <%if lcase(right(okurl,4))=".avi" then%>
		<object id="video" width="100%" height="100%" border="0" classid="clsid:CFCDAA03-8BE4-11cf-B84B-0020AFBBCCFA">
		<param name="ShowDisplay" value="0">
		<param name="ShowControls" value="1">
		<param name="AutoStart" value="1">
		<param name="AutoRewind" value="0">
		<param name="PlayCount" value="0">
		<param name="Appearance" value="0" value="">
		<param name="BorderStyle" value="0" value="">
                <%if (itemtype=0) then%>
		<param name="FileName" value="<%=okurl%>">
		<embed width="100%" height="100%" border="0" showdisplay="0" showcontrols="1" autostart="1" autorewind="0" playcount="0" filename="<%=okurl%>" src="<%=okurl%>">
                <%else%>
		<param name="MovieWindowHeight" value="<%=height%>">
		<param name="MovieWindowWidth" value="<%=width%>">
		<param name="FileName" value="<%=okurl%>">
		<embed width="100%" height="100%" border="0" showdisplay="0" showcontrols="1" autostart="1" autorewind="0" playcount="0" moviewindowheight="<%=height%>" moviewindowwidth="<%=width%>" filename="<%=okurl%>" src="<%=okurl%>">
                <%end if%>
		</embed>
		</object>                
	   <%dealit=true
           end if%>
           <%if lcase(right(okurl,4))=".mpg" then%>
		<object classid="clsid:05589FA1-C356-11CE-BF01-00AA0055595A" id="ActiveMovie1" width="<%=width%>" height="<%=height%>">
		<param name="Appearance" value="0">
		<param name="AutoStart" value="-1">
		<param name="AllowChangeDisplayMode" value="-1">
		<param name="AllowHideDisplay" value="0">
		<param name="AllowHideControls" value="-1">
		<param name="AutoRewind" value="-1">
		<param name="Balance" value="0">
		<param name="CurrentPosition" value="0">
		<param name="DisplayBackColor" value="0">
		<param name="DisplayForeColor" value="16777215">
		<param name="DisplayMode" value="0">
		<param name="Enabled" value="-1">
		<param name="EnableContextMenu" value="-1">
		<param name="EnablePositionControls" value="-1">
		<param name="EnableSelectionControls" value="0">
		<param name="EnableTracker" value="-1">
		<param name="Filename" value="<%=okurl%>" valuetype="ref">
		<param name="FullScreenMode" value="0">
		<param name="MovieWindowSize" value="0">
		<param name="PlayCount" value="1">
		<param name="Rate" value="1">
		<param name="SelectionStart" value="-1">
		<param name="SelectionEnd" value="-1">
		<param name="ShowControls" value="-1">
		<param name="ShowDisplay" value="-1">
		<param name="ShowPositionControls" value="0">
		<param name="ShowTracker" value="-1">
		<param name="Volume" value="-480">
		</object>
           <%dealit=true
           end if%>
	   <%if lcase(right(okurl,3))=".rm" then%>
		<OBJECT ID=video1 CLASSID="clsid:CFCDAA03-8BE4-11cf-B84B-0020AFBBCCFA" HEIGHT=<%=height%> WIDTH=<%=width%>>
		<param name="_ExtentX" value="9313">
		<param name="_ExtentY" value="7620">
		<param name="AUTOSTART" value="0">
		<param name="SHUFFLE" value="0">
		<param name="PREFETCH" value="0">
		<param name="NOLABELS" value="0">
		<param name="SRC" value="rtsp://<%=okurl%>">
		<param name="CONTROLS" value="ImageWindow">
		<param name="CONSOLE" value="Clip1">
		<param name="LOOP" value="0">
		<param name="NUMLOOP" value="0">
		<param name="CENTER" value="0">
		<param name="MAINTAINASPECT" value="0">
                <%if (itemtype=0) then%>
		<param name="BACKGROUNDCOLOR" value="#000000"><embed SRC type="audio/x-pn-realaudio-plugin" CONSOLE="Clip1" CONTROLS="ImageWindow" AUTOSTART="false">
                <%else%>
		<param name="BACKGROUNDCOLOR" value="#000000"><embed SRC type="audio/x-pn-realaudio-plugin" CONSOLE="Clip1" CONTROLS="ImageWindow" HEIGHT="<%=height%>"	WIDTH="<%=width%>" AUTOSTART="false">
                <%end if%>
		</OBJECT>
	   <%dealit=true
           end if%>
	   <%if (okurl="") or (lcase(right(okurl,4))=".wmv") then%>
		<object id="NSPlay" width=<%=width%> height=<%=height%> classid="CLSID:22d6f312-b0f6-11d0-94ab-0080c74c7e95" codebase="http://activex.microsoft.com/activex/controls/mplayer/en/nsmp2inf.cab#Version=6,4,5,715" standby="Loading Microsoft Windows Media Player components..." type="application/x-oleobject" align="left" hspace="5">
		<param name="AutoRewind" value=1>
		<param name="FileName" value="<%=okurl%>">
		<param name="ShowControls" value="1">
		<param name="ShowPositionControls" value="0">
		<param name="ShowAudioControls" value="1">
		<param name="ShowTracker" value="0">
		<param name="ShowDisplay" value="0">
		<param name="ShowStatusBar" value="0">
		<param name="ShowGotoBar" value="0">
		<param name="ShowCaptioning" value="0">
		<param name="AutoStart" value=1>
		<param name="Volume" value="-2500">
		<param name="AnimationAtStart" value="0">
		<param name="TransparentAtStart" value="0">
		<param name="AllowChangeDisplaySize" value="0">
		<param name="AllowScan" value="0">
		<param name="EnableContextMenu" value="0">
		<param name="ClickToPlay" value="0">
		</object>
	   <%dealit=true
           end if%>
	   <%if (lcase(right(okurl,4))=".wma") or (lcase(right(okurl,4))=".mp3") then%>
		<object classid="clsid:22D6F312-B0F6-11D0-94AB-0080C74C7E95" id="MediaPlayer1" > 
		<param name="Filename" value="<%=okurl%>">
		<param name="PlayCount" value="1">
		<param name="AutoStart" value="1">
		<param name="ClickToPlay" value="1">
		<param name="DisplaySize" value="0">
		<param name="EnableFullScreen Controls" value="1">
		<param name="ShowAudio Controls" value="1">
		<param name="EnableContext Menu" value="1">
		<param name="ShowDisplay" value="1">
		</object>
	   <%dealit=true
           end if%>
           <%if dealit=false then
			if (itemtype=0) then
			   response.write "<p valign=""middle"" " & alignstr & "><img border=""0"" dynsrc=""" & path & """ start=""fileopen"" " & circlestr & "></p>" & chr(13) & chr(10)
			else
			   response.write "<p valign=""middle"" " & alignstr & "><img border=""0"" dynsrc=""" & path & """ width=""" & width & """ height=""" & height & """ start=""fileopen"" " & circlestr & "></p>" & chr(13) & chr(10)
			end if
          end if
			response.write "</center>" & chr(13) & chr(10)
			response.write "</body>" & chr(13) & chr(10)
			response.write "</html>" & chr(13) & chr(10)
		case 11
			'response.write "11-操作系统自检测</font></b><br>"
			dispmsg=dispmsg & "11-操作系统自检测"
			response.write "<html>" & chr(13) & chr(10)
			response.write "<head>" & chr(13) & chr(10)
			response.write charsetstr & chr(13) & chr(10)
			response.write "<title>Digital Multi-Media Distributing System</title>" & chr(13) & chr(10)
   			response.write "<script language=""javascript"">" & chr(13) & chr(10)
   			response.write "<!--" & chr(13) & chr(10)
   			response.write "function displayinfo(dispstr)" & chr(13) & chr(10)
   			response.write "{" & chr(13) & chr(10)
   			response.write "   alert(dispstr);" & chr(13) & chr(10)
   			response.write "}" & chr(13) & chr(10)
   			response.write "-->" & chr(13) & chr(10)
   			response.write "</script>" & chr(13) & chr(10)
			response.write "</head>" & chr(13) & chr(10)
			response.write "<body background=""" & bkpic & """ bgcolor=""" & bkclr & """ ondblclick=""displayinfo('" & dispmsg & "');"">" & chr(13) & chr(10)
			response.write "<center>" & chr(13) & chr(10)
			response.write "<p>素材文件：" & path
			response.write "不支持该类型素材的预览</p>"
			response.write "</center>" & chr(13) & chr(10)
			response.write "</body>" & chr(13) & chr(10)
			response.write "</html>" & chr(13) & chr(10)
		case 12
			'response.write "12-专用应用程序</font></b><br>"
			dispmsg=dispmsg & "12-专用应用程序"
			response.write "<html>" & chr(13) & chr(10)
			response.write "<head>" & chr(13) & chr(10)
			response.write charsetstr & chr(13) & chr(10)
			response.write "<title>Digital Multi-Media Distributing System</title>" & chr(13) & chr(10)
   			response.write "<script language=""javascript"">" & chr(13) & chr(10)
   			response.write "<!--" & chr(13) & chr(10)
   			response.write "function displayinfo(dispstr)" & chr(13) & chr(10)
   			response.write "{" & chr(13) & chr(10)
   			response.write "   alert(dispstr);" & chr(13) & chr(10)
   			response.write "}" & chr(13) & chr(10)
   			response.write "-->" & chr(13) & chr(10)
   			response.write "</script>" & chr(13) & chr(10)
			response.write "</head>" & chr(13) & chr(10)
			response.write "<body background=""" & bkpic & """ bgcolor=""" & bkclr & """ ondblclick=""displayinfo('" & dispmsg & "');"">" & chr(13) & chr(10)
			response.write "<center>" & chr(13) & chr(10)
			response.write "<p>素材文件：" & path
			response.write "不支持该类型素材的预览</p>"
			response.write "</center>" & chr(13) & chr(10)
			response.write "</body>" & chr(13) & chr(10)
			response.write "</html>" & chr(13) & chr(10)
		case 13
			'response.write "13-远程指令</font></b><br>"
			dispmsg=dispmsg & "13-远程指令"
			response.write "<html>" & chr(13) & chr(10)
			response.write "<head>" & chr(13) & chr(10)
			response.write charsetstr & chr(13) & chr(10)
			response.write "<title>Digital Multi-Media Distributing System</title>" & chr(13) & chr(10)
   			response.write "<script language=""javascript"">" & chr(13) & chr(10)
   			response.write "<!--" & chr(13) & chr(10)
   			response.write "function displayinfo(dispstr)" & chr(13) & chr(10)
   			response.write "{" & chr(13) & chr(10)
   			response.write "   alert(dispstr);" & chr(13) & chr(10)
   			response.write "}" & chr(13) & chr(10)
   			response.write "-->" & chr(13) & chr(10)
   			response.write "</script>" & chr(13) & chr(10)
			response.write "</head>" & chr(13) & chr(10)
			response.write "<body background=""" & bkpic & """ bgcolor=""" & bkclr & """ ondblclick=""displayinfo('" & dispmsg & "');"">" & chr(13) & chr(10)
			response.write "<center>" & chr(13) & chr(10)
			response.write "<p>不支持该类型素材的预览</p>"
			response.write "</center>" & chr(13) & chr(10)
			response.write "</body>" & chr(13) & chr(10)
			response.write "</html>" & chr(13) & chr(10)
		Case Else
			'response.write contenttype & "-未知类型<br></font>"
			response.write "<html>" & chr(13) & chr(10)
			response.write "<head>" & chr(13) & chr(10)
			response.write charsetstr & chr(13) & chr(10)
			response.write "<title>Digital Multi-Media Distributing System</title>" & chr(13) & chr(10)
   			response.write "<script language=""javascript"">" & chr(13) & chr(10)
   			response.write "<!--" & chr(13) & chr(10)
   			response.write "function displayinfo(dispstr)" & chr(13) & chr(10)
   			response.write "{" & chr(13) & chr(10)
   			response.write "   alert(dispstr);" & chr(13) & chr(10)
   			response.write "}" & chr(13) & chr(10)
   			response.write "-->" & chr(13) & chr(10)
   			response.write "</script>" & chr(13) & chr(10)
			response.write "</head>" & chr(13) & chr(10)
			response.write "<body background=""" & bkpic & """ bgcolor=""" & bkclr & """ ondblclick=""displayinfo('" & dispmsg & "');"">" & chr(13) & chr(10)
			response.write "<center><p>" & chr(13) & chr(10)
			dispmsg=dispmsg & contenttype & "-未知类型</p>"
			response.write "</center>" & chr(13) & chr(10)
			response.write "</body>" & chr(13) & chr(10)
			response.write "</html>" & chr(13) & chr(10)
		End Select
    else
		dat.close
		set dat=nothing
		response.write "栏目或者节目单无法预览，或者给出的素材标记有误，没有查找到该记录！"
		response.end 
    end if
%>