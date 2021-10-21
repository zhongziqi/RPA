// 搜索关键词
Dim keyWords = ""
// 职位名称
Dim jobTitle = "招聘"
// 期望工作地
Dim expectAddress = "深圳"
// 最小年龄
Dim minAge = 18
// 最大年龄
Dim maxAge = 30
// 最少工作年限
Dim workExperience = ""
// 期望最少工资 (单位:k)
Dim expectMinSalary = 5
Dim expectMaxSalary = 12 

Dim hWeb = ""
Dim iRet = ""
Dim arrayData = ""
Dim bRet = ""
Dim sRet = ""
Dim isScroll = False
Dim scrapWord = ""
Dim page = 1 
Dim pageNum = 10
Dim totalPage = 5


hWeb = WebBrowser.Create("chrome","https://ehire.51job.com/navigate.aspx",30000,{"bContinueOnError":False,"iDelayAfter":300,"iDelayBefore":200,"sBrowserPath":"","sStartArgs":""})
#icon("@res:mhulerr5-sek1-vu7u-cbm7-ap6l58rlhu4t.png")
Mouse.Action({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"A","id":"MainMenuNew1_m3"}]},"left","click",10000,{"bContinueOnError":False,"iDelayAfter":300,"iDelayBefore":200,"bSetForeground":True,"sCursorPosition":"Center","iCursorOffsetX":0,"iCursorOffsetY":0,"sKeyModifiers":[],"sSimulate":"simulate","bMoveSmoothly":False})
#icon("@res:8ga78g9e-5kq3-fhvs-eqjl-qelh2c5a5679.png")
// Keyboard.InputText({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"INPUT","id":"search_keyword_txt"}]},keyWords,True,20,10000,{"bContinueOnError":False,"iDelayAfter":300,"iDelayBefore":800,"bSetForeground":True,"sSimulate":"message","bValidate":False,"bClickBeforeInput":False})
#icon("@res:v3uk074d-u21d-dnsu-cnd8-jj8oemagbnd3.png")
Keyboard.InputText({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"INPUT","id":"search_jobname_txt"}]},jobTitle,True,20,10000,{"bContinueOnError":False,"iDelayAfter":300,"iDelayBefore":800,"bSetForeground":True,"sSimulate":"message","bValidate":False,"bClickBeforeInput":False})
#icon("@res:vboufc8g-3uqr-4fop-ph92-0bfab2adi4pq.png")
Mouse.Action({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"A","id":"TopDegree_5|"}]},"left","click",10000,{"bContinueOnError":False,"iDelayAfter":300,"iDelayBefore":200,"bSetForeground":True,"sCursorPosition":"Center","iCursorOffsetX":0,"iCursorOffsetY":0,"sKeyModifiers":[],"sSimulate":"simulate","bMoveSmoothly":False})
#icon("@res:hdejojab-81ho-5lcv-a5di-1cq2g8a24crf.png")
Keyboard.InputText({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"INPUT","id":"search_expjobarea_txt"}]},expectAddress,True,20,10000,{"bContinueOnError":False,"iDelayAfter":300,"iDelayBefore":500,"bSetForeground":True,"sSimulate":"message","bValidate":False,"bClickBeforeInput":False})
#icon("@res:0pvg33uf-2l9k-2hl1-cjvu-k9s40sljcnfd.png")
Keyboard.InputText({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"INPUT","id":"search_agef_input"}]},minAge,True,20,10000,{"bContinueOnError":False,"iDelayAfter":300,"iDelayBefore":500,"bSetForeground":True,"sSimulate":"message","bValidate":False,"bClickBeforeInput":False})
#icon("@res:655ddamc-5v81-i2ti-pccc-rj5sk44tolj2.png")
Keyboard.InputText({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"INPUT","id":"search_aget_input"}]},maxAge,True,20,10000,{"bContinueOnError":False,"iDelayAfter":300,"iDelayBefore":500,"bSetForeground":True,"sSimulate":"message","bValidate":False,"bClickBeforeInput":False})
#icon("@res:2ikutibi-qqfr-uvr4-5h37-qsppb3g1imaq.png")
// Keyboard.InputText({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"INPUT","id":"search_wyf_input"}]},workExperience,True,20,10000,{"bContinueOnError":False,"iDelayAfter":300,"iDelayBefore":500,"bSetForeground":True,"sSimulate":"message","bValidate":False,"bClickBeforeInput":False})
#icon("@res:828d1ban-k3dv-b4jj-605n-dgj4pti49ce1.png")
Keyboard.InputText({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"INPUT","id":"search_expectsalary_input"}]},expectMinSalary,True,20,10000,{"bContinueOnError":False,"iDelayAfter":300,"iDelayBefore":500,"bSetForeground":True,"sSimulate":"message","bValidate":False,"bClickBeforeInput":False})
#icon("@res:op5vlhpk-m1b7-4a0o-5i82-usecpt7vmtf7.png")
Keyboard.InputText({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"INPUT","id":"search_expectsalaryto_input"}]},expectMaxSalary,True,20,10000,{"bContinueOnError":False,"iDelayAfter":300,"iDelayBefore":500,"bSetForeground":True,"sSimulate":"message","bValidate":False,"bClickBeforeInput":False})
#icon("@res:d7ntu118-9pu6-dbte-2akg-q881e4jkmjme.png")
Mouse.Action({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"I","parentid":"allShowdivHide","idx":4}]},"left","click",10000,{"bContinueOnError":False,"iDelayAfter":1000,"iDelayBefore":200,"bSetForeground":True,"sCursorPosition":"Center","iCursorOffsetX":0,"iCursorOffsetY":0,"sKeyModifiers":[],"sSimulate":"simulate","bMoveSmoothly":False})
#icon("@res:vmok8naj-crt9-n1g2-68ok-b6ff9ndrooju.png")
Mouse.Action({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"A","id":"search_jobstatus_a_1"}]},"left","click",10000,{"bContinueOnError":False,"iDelayAfter":300,"iDelayBefore":200,"bSetForeground":True,"sCursorPosition":"Center","iCursorOffsetX":0,"iCursorOffsetY":0,"sKeyModifiers":[],"sSimulate":"simulate","bMoveSmoothly":False})
#icon("@res:vch01qi6-5bik-6eq2-vdae-0ti8rdknkp9c.png")
Mouse.Action({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"I","parentid":"allShowdivHide","idx":6}]},"left","click",10000,{"bContinueOnError":False,"iDelayAfter":1000,"iDelayBefore":400,"bSetForeground":True,"sCursorPosition":"Center","iCursorOffsetX":0,"iCursorOffsetY":0,"sKeyModifiers":[],"sSimulate":"simulate","bMoveSmoothly":False})
#icon("@res:rrk39dhe-viav-4tpr-8c9k-r8svup9e42fd.png")
Mouse.Action({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"A","id":"search_jobterm_a_1"}]},"left","click",10000,{"bContinueOnError":False,"iDelayAfter":300,"iDelayBefore":200,"bSetForeground":True,"sCursorPosition":"Center","iCursorOffsetX":0,"iCursorOffsetY":0,"sKeyModifiers":[],"sSimulate":"simulate","bMoveSmoothly":False})
#icon("@res:hes7tem7-rbho-c0kv-gnk9-vkfihkq6ks09.png")
Mouse.Action({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"I","parentid":"allShowdivHide","idx":2}]},"left","click",10000,{"bContinueOnError":False,"iDelayAfter":1000,"iDelayBefore":200,"bSetForeground":True,"sCursorPosition":"Center","iCursorOffsetX":0,"iCursorOffsetY":0,"sKeyModifiers":[],"sSimulate":"simulate","bMoveSmoothly":False})
#icon("@res:ivk09r77-38aq-qpcd-83er-j1qsqkknc72d.png")
Mouse.Action({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"A","id":"search_rsmupdate_a_0"}]},"left","click",10000,{"bContinueOnError":False,"iDelayAfter":300,"iDelayBefore":200,"bSetForeground":True,"sCursorPosition":"Center","iCursorOffsetX":0,"iCursorOffsetY":0,"sKeyModifiers":[],"sSimulate":"simulate","bMoveSmoothly":False})
#icon("@res:ernlija9-od6i-3hon-3n8f-toqs6mb5dgn0.png")



// exit()

// 获取总页数
totalRet = Text.Get({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"SPAN","parentid":"form1","css-selector":"body>form>div>div>div>div>ul>li.Search_num-set>span"}]},10000,{"bContinueOnError":False,"iDelayAfter":2000,"iDelayBefore":2000,"bSetForeground":True})
totalPage = Cint(Regex.FindStr(totalRet,"(?<=/).+",0))
TracePrint(totalPage)
// 循环总页数
For x = 1 To totalPage Step 1
	Dim y = "'" & x & "'"  
	// 滚动到底部加载当前页面全部数据
	Do While isScroll= False 
		#icon("@res:og8donnq-6kor-99uu-vpuj-g81mru12fdj7.png")
		isScroll = UiElement.Exists({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"DIV","parentid":"search_resume_list"}]},{"bContinueOnError":False,"iDelayAfter":300,"iDelayBefore":200})
		Mouse.Wheel(50,"down", [],{"iDelayAfter":1000,"iDelayBefore":2000})
	Loop
	#icon("@res:0tn0cps8-6j2e-dqc6-u742-4j1kfn3sr8mj.png")
	Mouse.Action({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"A","css-selector":"body>form>div>div>div.position-list>div>a[title="& y &"]"}]},"left","click",10000,{"bContinueOnError":False,"iDelayAfter":300,"iDelayBefore":2000,"bSetForeground":True,"sCursorPosition":"Center","iCursorOffsetX":0,"iCursorOffsetY":0,"sKeyModifiers":[],"sSimulate":"simulate","bMoveSmoothly":False})
	// TracePrint("body>form>div>div>div.position-list>div>a[title="& y &"]")
	
	// 爬取当前页面数据
	arrayData = UiElement.DataScrap({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"DIV","id":"search_resume_list"}]},{"ExtractTable":0,"Columns":[{"selecors":[{"tag":"ul","index":0,"className":"ls","value":"ul.ls","prefix":""},{"tag":"li","value":"li","index":0,"prefix":">"},{"tag":"div","index":0,"className":"position-list-asd fl","value":"div.position-list-asd.fl","prefix":">"},{"tag":"div","index":0,"className":"","value":"div","prefix":">"},{"tag":"div","index":0,"className":"position-info clearfix","value":"div.position-info.clearfix","prefix":">"},{"tag":"div","index":0,"className":"fl","value":"div.fl","prefix":">"},{"tag":"div","index":0,"className":"position-info-con fl","value":"div.position-info-con.fl","prefix":">"},{"tag":"div","index":0,"className":"clearfix","value":"div.clearfix","prefix":">"},{"tag":"p","index":0,"className":"fl position-text","value":"p.fl.position-text","prefix":">"}],"props":["text"]},{"selecors":[{"tag":"ul","index":0,"className":"ls","value":"ul.ls","prefix":""},{"tag":"li","value":"li","index":0,"prefix":">"},{"tag":"div","index":0,"className":"position-list-asd fl","value":"div.position-list-asd.fl","prefix":">"},{"tag":"div","index":0,"className":"","value":"div","prefix":">"},{"tag":"div","index":0,"className":"position-info clearfix","value":"div.position-info.clearfix","prefix":">"},{"tag":"div","index":0,"className":"fl","value":"div.fl","prefix":">"},{"tag":"div","index":0,"className":"position-info-con fl","value":"div.position-info-con.fl","prefix":">"},{"tag":"div","index":0,"className":"clearfix","value":"div.clearfix","prefix":">"},{"tag":"div","index":0,"className":"fl position-li-id","value":"div.fl.position-li-id","prefix":">"},{"tag":"span","index":2,"className":"","value":"span:nth-child(2)","prefix":">"}],"props":["text"]},{"selecors":[{"tag":"ul","index":0,"className":"ls","value":"ul.ls","prefix":""},{"tag":"li","value":"li","index":0,"prefix":">"},{"tag":"div","index":0,"className":"position-list-asd fl","value":"div.position-list-asd.fl","prefix":">"},{"tag":"div","index":0,"className":"","value":"div","prefix":">"},{"tag":"div","index":0,"className":"position-info clearfix","value":"div.position-info.clearfix","prefix":">"},{"tag":"div","index":0,"className":"fl","value":"div.fl","prefix":">"},{"tag":"div","index":0,"className":"position-info-con fl","value":"div.position-info-con.fl","prefix":">"},{"tag":"ul","index":0,"className":"position-info-con1 clearfix mt10","value":"ul.position-info-con1.clearfix.mt10","prefix":">"}],"props":["text"]}]},{"objNextLinkElement":{"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"A","id":"pagerBottomNew_nextButton"}]},"iMaxNumberOfPage":1,"iMaxNumberOfResult":-1,"iDelayBetweenMS":2000,"bContinueOnError":False})
	iRet = Len(arrayData)
	TracePrint( iRet)
	
	// 循环当前页面数据
	For i = 1 To iRet Step 1 
		// 设置元素属性
		#icon("@res:pn127foe-fqcr-s8cp-04p3-van96p0qan9m.png")
		UiElement.SetAttribute({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"UL","parentid":"search_resume_list"}]},"id","ul_0",{"bContinueOnError":False,"iDelayAfter":1000,"iDelayBefore":1000})
		Try 3
			sRet = UiElement.GetAttribute({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"LI","parentid":"ul_0","idx":i-1}]},"id",{"bContinueOnError":False,"iDelayAfter":600,"iDelayBefore":1000})
		Catch aa
			TracePrint("没有获取到li的id属性")
		Else
			TracePrint("已获取到li元素的id属性")
		End Try
		TracePrint(sRet)
		
		// 获取点击按钮文本
		Try 3
			#icon("@res:trd49lra-ofp7-qhkm-u0ak-cp184fripkg3.png")
			scrapWord = Text.Get({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"A","parentid":sRet,"css-selector":"body>form>div>div>div>ul>li>div>div>div>div>a.free-hichat"}]},10000,{"bContinueOnError":False,"iDelayAfter":2000,"iDelayBefore":2000,"bSetForeground":True})
		Catch a
			TracePrint("没有获取到点击按钮的文本")
		Else
			TracePrint(scrapWord)
		End Try
		
		// 判断点击按钮文本为: "立即Hi聊" 则点击触发
		If scrapWord="立即Hi聊"
			#icon("@res:m4gv6ni2-ro9b-k9bq-jn3v-626o20oh5uvl.png")
			Mouse.Action({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"A","parentid":sRet }]},"left","click",10000,{"bContinueOnError":False,"iDelayAfter":300,"iDelayBefore":1000,"bSetForeground":True,"sCursorPosition":"Center","iCursorOffsetX":0,"iCursorOffsetY":0,"sKeyModifiers":[],"sSimulate":"simulate","bMoveSmoothly":False})
			
			// 关闭确认消耗窗口
			#icon("@res:nllcqajp-seqt-vj63-co4j-5hml322scfms.png")
			bRet = Text.Exists({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"DIV","parentid":"tip_msgbox2_content"}]},"聊过","instr",1,10000,{"bContinueOnError":False,"iDelayAfter":300,"iDelayBefore":200,"bSetForeground":True})
			If bRet=False
				
				Try 3
					#icon("@res:0qnrpsbi-kfuo-0jtf-eihi-fkqf6o20qddu.png")
					Mouse.Action({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"DIV","parentid":"chat_select_job","aaname":"确定"}]},"left","click",10000,{"bContinueOnError":False,"iDelayAfter":800,"iDelayBefore":1500,"bSetForeground":True,"sCursorPosition":"Center","iCursorOffsetX":0,"iCursorOffsetY":0,"sKeyModifiers":[],"sSimulate":"simulate","bMoveSmoothly":False})
				Catch b
					TracePrint(b)
				Else
					TracePrint("关闭消耗窗口成功")
				End Try
				Try 3
					#icon("@res:f31j0hu2-5biu-57j7-il57-8kdk9lm3d9oj.png")
					Mouse.Action({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"DIV","parentid":"tip_autobox1","aaname":"确定"}]},"left","click",10000,{"bContinueOnError":False,"iDelayAfter":800,"iDelayBefore":1500,"bSetForeground":True,"sCursorPosition":"Center","iCursorOffsetX":0,"iCursorOffsetY":0,"sKeyModifiers":[],"sSimulate":"simulate","bMoveSmoothly":False})
				Catch c
					TracePrint(c)
				Else
					TracePrint("关闭确认窗口")
				End Try
				Try 3
					#icon("@res:sgk9kbch-dbaj-kiv9-f4mv-64csv9dvvh78.png")
					Mouse.Action({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"IFRAME","id":"chatframe"},{"tag":"I","id":"chatclosediv"}]},"left","click",10000,{"bContinueOnError":False,"iDelayAfter":800,"iDelayBefore":1500,"bSetForeground":True,"sCursorPosition":"Center","iCursorOffsetX":0,"iCursorOffsetY":0,"sKeyModifiers":[],"sSimulate":"simulate","bMoveSmoothly":False})
				Catch qq
					TracePrint(qq)
				Else
					TracePrint("关闭聊天窗口")
				End Try
			Else
				
				#icon("@res:nhkh77mf-p4pt-g0d8-es8j-sjc7uv7dkh0o.png")
				Mouse.Action({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"IMG","parentid":"tip_msgbox2"}]},"left","click",10000,{"bContinueOnError":False,"iDelayAfter":300,"iDelayBefore":200,"bSetForeground":True,"sCursorPosition":"Center","iCursorOffsetX":0,"iCursorOffsetY":0,"sKeyModifiers":[],"sSimulate":"simulate","bMoveSmoothly":False})
				TracePrint("其他同事聊过了")
			End If
			
			// 关闭确认确认窗口
			// Mouse.Action({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"DIV","parentid":"li_1_308837212","css-selector":"body>form>div>div>div>ul>li>div>div>div>div","idx":1}]},"left","click",10000,{"bContinueOnError":False,"iDelayAfter":2000,"iDelayBefore":200,"bSetForeground":True,"sCursorPosition":"Center","iCursorOffsetX":0,"iCursorOffsetY":0,"sKeyModifiers":[],"sSimulate":"simulate","bMoveSmoothly":False})
			
			// 关闭聊天窗口
		Else
			TracePrint("当前候选人已经沟通过了")
			
			// Mouse.Action({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"A","parentid":sRet }]},"left","click",10000,{"bContinueOnError":False,"iDelayAfter":300,"iDelayBefore":1000,"bSetForeground":True,"sCursorPosition":"Center","iCursorOffsetX":0,"iCursorOffsetY":0,"sKeyModifiers":[],"sSimulate":"simulate","bMoveSmoothly":False})
			
		End If
	Next
Next
// exit()


// 滚动到底部加载全部数据

// 循环获取到的数据


