

Dim dictScrollPostion = ""
Dim arrayData = ""
Dim arrayDataHello = ""
Dim arrElement = ""
Dim sRet = ""
Dim objRect = ""
Dim bRet = False
Dim objRet = ""
Dim iRet = ""
Dim arrayRet = ""
Dim objExcelWorkBook = ""
Dim objDatatable = ""
Dim arrRet = ""
Dim hWeb = ""
Dim bRets = ""
Dim x = 1
Dim findName = ""
Dim workYear = ""
Dim expectJob = ""
Dim age = ""
Dim summary = ""
Dim iRetFather = ""
Dim iRetSon = ""
Dim helloFather =""
Dim helloSon = ""
Dim nowTime = ""
Dim midTime = ""
Dim midArray = []
Dim targetTime = ""
Dim leftTime = ""
Dim number = 0
Dim require_react = ""
Dim working = ""

hWeb = WebBrowser.Create("chrome","https://www.zhipin.com/web/boss/index",30000,{"bContinueOnError":False,"iDelayAfter":5000,"iDelayBefore":200,"sBrowserPath":"","sStartArgs":""})
#icon("@res:o57nj5kg-k7m3-2spb-h8eo-j9s6v05rmhni.png")
Mouse.Action({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"A","type":"desktop","parentid":"main","aaname":"推荐牛人                 "}]},"left","click",10000,{"bContinueOnError":False,"iDelayAfter":3000,"iDelayBefore":200,"bSetForeground":True,"sCursorPosition":"Center","iCursorOffsetX":0,"iCursorOffsetY":0,"sKeyModifiers":[],"sSimulate":"simulate","bMoveSmoothly":False})
#icon("@res:kfs5f2rm-qa4u-k2pf-pai2-7helqnv75sl9.png")
Mouse.Action({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"I","parentid":"recommendContent"}]},"left","click",10000,{"bContinueOnError":False,"iDelayAfter":1000,"iDelayBefore":200,"bSetForeground":True,"sCursorPosition":"Center","iCursorOffsetX":0,"iCursorOffsetY":0,"sKeyModifiers":[],"sSimulate":"simulate","bMoveSmoothly":False})
arrayData = UiElement.DataScrap({"html":[{"id":"recommendContent","parentid":"main","tag":"DIV"}],"wnd":[{"app":"iexplore","cls":"IEFrame","title":"*"},{"cls":"Internet Explorer_Server"}]},{"ExtractTable":0,"Columns":[{"selecors":[{"tag":"div","index":1,"className":"recommend-container page-container page-inner","value":"div:nth-child(1)","prefix":""},{"tag":"div","index":0,"className":"page-header","value":"div.page-header","prefix":">"},{"tag":"div","index":2,"className":"recommend-top-dropdown  ui-dropmenu ui-dropmenu-visible ui-dropmenu-label-arrow ui-dropmenu-drop-arrow","value":"div:nth-child(2)","prefix":">"},{"tag":"div","index":2,"className":"ui-dropmenu-list","value":"div:nth-child(2)","prefix":">"},{"tag":"div","index":0,"value":"div","prefix":">"},{"tag":"ul","index":2,"value":"ul:nth-child(2)","prefix":">"},{"tag":"li","value":"li","index":0,"prefix":">"}],"props":["text"]}]},{"objNextLinkElement":"","iMaxNumberOfPage":5,"iMaxNumberOfResult":-1,"iDelayBetweenMS":1000,"bContinueOnError":False})
#icon("@res:rff4b5he-mn01-lml7-kmol-agkfovr6evt5.png")
Mouse.Action({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"I","parentid":"recommendContent"}]},"left","click",10000,{"bContinueOnError":False,"iDelayAfter":1000,"iDelayBefore":200,"bSetForeground":True,"sCursorPosition":"Center","iCursorOffsetX":0,"iCursorOffsetY":0,"sKeyModifiers":[],"sSimulate":"simulate","bMoveSmoothly":False})
For i = 1 To 1 Step 1 
	#icon("@res:mm284vt1-2omf-u5re-3t24-5dh32ndns3ua.png")
	Mouse.Action({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"I","parentid":"recommendContent"}]},"left","click",10000,{"bContinueOnError":False,"iDelayAfter":1000,"iDelayBefore":200,"bSetForeground":True,"sCursorPosition":"Center","iCursorOffsetX":0,"iCursorOffsetY":0,"sKeyModifiers":[],"sSimulate":"simulate","bMoveSmoothly":False})
	TracePrint(i & ':i---')
	// ui-dropmenu-label
	#icon("@res:fo8nbpkf-e5ij-rj49-u0ca-haaiv75ipfaf.png")
	// Mouse.Action({"html":[{"parentid":"recommendContent","tag":"LI" ,"attrMap":{"css-selector":"div.ui-dropmenu-list>div>ul>li:nth-child("& i &")"}}],"wnd":[{"app":"iexplore","cls":"IEFrame","title":"*"},{"cls":"Internet Explorer_Server"}]},"left","click",10000,{"bContinueOnError":False,"iDelayAfter":5000,"iDelayBefore":1000,"bSetForeground":True,"sCursorPosition":"Center","iCursorOffsetX":0,"iCursorOffsetY":0,"sKeyModifiers":[],"sSimulate":"simulate","bMoveSmoothly":False})
	Mouse.Action({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"LI","parentid":"recommendContent","idx":2}]},"left","click",10000,{"bContinueOnError":False,"iDelayAfter":3000,"iDelayBefore":200,"bSetForeground":True,"sCursorPosition":"Center","iCursorOffsetX":0,"iCursorOffsetY":0,"sKeyModifiers":[],"sSimulate":"simulate","bMoveSmoothly":False})
	
	// 获取推荐的所有候选人数据--调试时注释
	Do While bRet= False 
		#icon("@res:default.png")
		bRet = Text.Exists({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"IFRAME","name":"recommendFrame"},{"tag":"DIV","parentid":"recommend-list"}]},"没有更多","instr",1,10000,{"bContinueOnError":False,"iDelayAfter":300,"iDelayBefore":200,"bSetForeground":True})
		Mouse.Wheel(50,"down", [],{"iDelayAfter":1000,"iDelayBefore":2000})
		//  TracePrint(bRet)
	Loop
	
	// 使用中的爬虫代码  
	arrayData = UiElement.DataScrap({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"IFRAME","name":"recommendFrame"},{"tag":"DIV","id":"recommend-list"}]},{"ExtractTable":0,"Columns":[{"selecors":[{"tag":"div","index":0,"className":"sec-content candidate-card","value":"div.sec-content.candidate-card","prefix":""},{"tag":"ul","index":0,"className":"recommend-card-list ul-less-height","value":"ul.recommend-card-list.ul-less-height","prefix":">"},{"tag":"li","value":"li","index":0,"prefix":">"},{"tag":"div","index":0,"className":"","value":"div","prefix":">"},{"tag":"div","index":0,"className":"geek-info-card less-height","value":"div.geek-info-card.less-height","prefix":">"},{"tag":"div","index":0,"className":"candidate-list-content","value":"div.candidate-list-content","prefix":">"},{"tag":"div","index":0,"className":"card-inner","value":"div.card-inner","prefix":">"},{"tag":"div","index":0,"className":"col-2","value":"div.col-2","prefix":">"},{"tag":"div","index":0,"className":"name","value":"div.name","prefix":">"}],"props":["text"]},{"selecors":[{"tag":"div","index":0,"className":"sec-content candidate-card","value":"div.sec-content.candidate-card","prefix":""},{"tag":"ul","index":0,"className":"recommend-card-list ul-less-height","value":"ul.recommend-card-list.ul-less-height","prefix":">"},{"tag":"li","value":"li","index":0,"prefix":">"},{"tag":"div","index":0,"className":"","value":"div","prefix":">"},{"tag":"div","index":0,"className":"geek-info-card less-height","value":"div.geek-info-card.less-height","prefix":">"},{"tag":"div","index":0,"className":"candidate-list-content","value":"div.candidate-list-content","prefix":">"},{"tag":"div","index":0,"className":"card-inner","value":"div.card-inner","prefix":">"},{"tag":"div","index":0,"className":"col-2","value":"div.col-2","prefix":">"},{"tag":"div","index":0,"className":"name","value":"div.name","prefix":">"},{"tag":"span","index":0,"className":"label-text","value":"span.label-text","prefix":">"}],"props":["text"]},{"selecors":[{"tag":"div","index":0,"className":"sec-content candidate-card","value":"div.sec-content.candidate-card","prefix":""},{"tag":"ul","index":0,"className":"recommend-card-list ul-less-height","value":"ul.recommend-card-list.ul-less-height","prefix":">"},{"tag":"li","value":"li","index":0,"prefix":">"},{"tag":"div","index":0,"className":"","value":"div","prefix":">"},{"tag":"div","index":0,"className":"geek-info-card less-height","value":"div.geek-info-card.less-height","prefix":">"},{"tag":"div","index":0,"className":"candidate-list-content","value":"div.candidate-list-content","prefix":">"},{"tag":"div","index":0,"className":"card-inner","value":"div.card-inner","prefix":">"},{"tag":"div","index":0,"className":"col-1","value":"div.col-1","prefix":">"},{"tag":"div","index":0,"className":"salary","value":"div.salary","prefix":">"}],"props":["text"]},{"selecors":[{"tag":"div","index":0,"className":"sec-content candidate-card","value":"div.sec-content.candidate-card","prefix":""},{"tag":"ul","index":0,"className":"recommend-card-list ul-less-height","value":"ul.recommend-card-list.ul-less-height","prefix":">"},{"tag":"li","value":"li","index":0,"prefix":">"},{"tag":"div","index":0,"className":"","value":"div","prefix":">"},{"tag":"div","index":0,"className":"geek-info-card less-height","value":"div.geek-info-card.less-height","prefix":">"},{"tag":"div","index":0,"className":"candidate-list-content","value":"div.candidate-list-content","prefix":">"},{"tag":"div","index":0,"className":"card-inner","value":"div.card-inner","prefix":">"},{"tag":"div","index":0,"className":"col-2","value":"div.col-2","prefix":">"},{"tag":"div","index":0,"className":"info-labels","value":"div.info-labels","prefix":">"}],"props":["text"]},{"selecors":[{"tag":"div","index":0,"className":"sec-content candidate-card","value":"div.sec-content.candidate-card","prefix":""},{"tag":"ul","index":0,"className":"recommend-card-list ul-less-height","value":"ul.recommend-card-list.ul-less-height","prefix":">"},{"tag":"li","value":"li","index":0,"prefix":">"},{"tag":"div","index":0,"className":"","value":"div","prefix":">"},{"tag":"div","index":0,"className":"geek-info-card less-height","value":"div.geek-info-card.less-height","prefix":">"},{"tag":"div","index":0,"className":"candidate-list-content","value":"div.candidate-list-content","prefix":">"},{"tag":"div","index":0,"className":"card-inner","value":"div.card-inner","prefix":">"},{"tag":"div","index":0,"className":"col-2","value":"div.col-2","prefix":">"},{"tag":"div","index":0,"className":"expect-box","value":"div.expect-box","prefix":">"}],"props":["text"]},{"selecors":[{"tag":"div","index":0,"className":"sec-content candidate-card","value":"div.sec-content.candidate-card","prefix":""},{"tag":"ul","index":0,"className":"recommend-card-list ul-less-height","value":"ul.recommend-card-list.ul-less-height","prefix":">"},{"tag":"li","value":"li","index":0,"prefix":">"},{"tag":"div","index":0,"className":"","value":"div","prefix":">"},{"tag":"div","index":0,"className":"geek-info-card less-height","value":"div.geek-info-card.less-height","prefix":">"},{"tag":"div","index":0,"className":"candidate-list-content","value":"div.candidate-list-content","prefix":">"},{"tag":"div","index":0,"className":"card-inner","value":"div.card-inner","prefix":">"},{"tag":"div","index":0,"className":"col-2","value":"div.col-2","prefix":">"},{"tag":"div","index":0,"className":"position-box position-advantage","value":"div.position-box.position-advantage","prefix":">"},{"tag":"div","index":0,"className":"advantage-new","value":"div.advantage-new","prefix":">"}],"props":["text"]},{"selecors":[{"tag":"div","index":0,"className":"sec-content candidate-card","value":"div.sec-content.candidate-card","prefix":""},{"tag":"ul","index":0,"className":"recommend-card-list ul-less-height","value":"ul.recommend-card-list.ul-less-height","prefix":">"},{"tag":"li","value":"li","index":0,"prefix":">"},{"tag":"div","index":0,"className":"","value":"div","prefix":">"},{"tag":"div","index":0,"className":"geek-info-card less-height","value":"div.geek-info-card.less-height","prefix":">"},{"tag":"div","index":0,"className":"candidate-list-content","value":"div.candidate-list-content","prefix":">"},{"tag":"div","index":0,"className":"card-inner","value":"div.card-inner","prefix":">"},{"tag":"div","index":0,"className":"col-3","value":"div.col-3","prefix":">"}],"props":["text"]}]},{"objNextLinkElement":"","iMaxNumberOfPage":5,"iMaxNumberOfResult":-1,"iDelayBetweenMS":1000,"bContinueOnError":False})
	
	
	iRetFather = Len(arrayData)
	TracePrint('共获取到的候选人数量为：'& iRetFather)
	
	iRetSon = Len(arrayData[0])
	// 打开一个空白excel 用于存储获取的候选人数据----调试时注释
	// objExcelWorkBook = Excel.OpenExcel('''C:\Users\Administrator\Desktop\boss.xlsx''',True,"Excel","","")
	
	
	// objExcelWorkBook = Excel.OpenExcel('''C:\Users\Administrator\Desktop\招呼历史.xlsx''',True,"Excel","","")
	For y = 1 To iRetFather Step 1
		arrayDataHistory =[]
		sRet = ""
		// 2. 姓名正则  排除含外包/外派人员
		findName = ""
		// 3. 工作年限 (大于等于三年)
		workYear = 0
		// 4. 求职期望(含前端/Javascript/HTML/web)
		expectJob =  ""
		// 5. 年龄（<=30）
		age =  ""
		// 6.求职简介
		summary = ""
		// 7. 在职-暂不考虑
		working  = ""
		// 8.要求掌握react
		require_react =  ""
		
		
		
		
		// 判断当前候选人是否被其他招聘顾问询问过
		#icon("@res:ebp4k866-lmjp-l18g-9tao-fdknaqovp8av.png")
		// Mouse.Action({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"IFRAME","name":"recommendFrame"},{"tag":"EM","parentid":"recommend-list","css-selector":"body>div>div>div>div>div>div>div>div>div>ul>li>div>div>div>div>em.iboss-goutongjilu","idx":y-1}]},"left","click",10000,{"bContinueOnError":False,"iDelayAfter":3000,"iDelayBefore":3000,"bSetForeground":True,"sCursorPosition":"Center","iCursorOffsetX":0,"iCursorOffsetY":0,"sKeyModifiers":[],"sSimulate":"simulate","bMoveSmoothly":False})
		// arrayDataHistory = UiElement.DataScrap({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"HTML"}]},{"ExtractTable":0,"Columns":[{"selecors":[{"tag":"div","index":0,"className":"dialog-wrap dialog-prop-default coop-record-wrap","value":"div.dialog-wrap.dialog-prop-default.coop-record-wrap","prefix":""},{"tag":"div","index":0,"className":"dialog-container","value":"div.dialog-container","prefix":">"},{"tag":"div","index":0,"className":"dialog-con","value":"div.dialog-con","prefix":">"},{"tag":"div","index":0,"className":"chat-record-content","value":"div.chat-record-content","prefix":">"},{"tag":"ul","index":0,"className":"record","value":"ul.record","prefix":">"},{"tag":"li","value":"li","index":0,"prefix":">"},{"tag":"div","index":0,"className":"","value":"div","prefix":">"},{"tag":"p","index":0,"className":"action","value":"p.action","prefix":">"}],"props":["text"]},{"selecors":[{"tag":"div","index":0,"className":"dialog-wrap dialog-prop-default coop-record-wrap","value":"div.dialog-wrap.dialog-prop-default.coop-record-wrap","prefix":""},{"tag":"div","index":0,"className":"dialog-container","value":"div.dialog-container","prefix":">"},{"tag":"div","index":0,"className":"dialog-con","value":"div.dialog-con","prefix":">"},{"tag":"div","index":0,"className":"chat-record-content","value":"div.chat-record-content","prefix":">"},{"tag":"ul","index":0,"className":"record","value":"ul.record","prefix":">"},{"tag":"li","value":"li","index":0,"prefix":">"},{"tag":"div","index":0,"className":"","value":"div","prefix":">"},{"tag":"p","index":0,"className":"operat","value":"p.operat","prefix":">"}],"props":["text"]}]},{"objNextLinkElement":"","iMaxNumberOfPage":5,"iMaxNumberOfResult":-1,"iDelayBetweenMS":1000,"bContinueOnError":False})
		// TracePrint(arrayData[y-1][0]&"为第"&y-1&"个候选人+------+"&arrayDataHistory[0][1])
		// // 当前时间
		// nowTime = Time.Now()
		// // 目标时间转换
		// midTime = Regex.Replace(arrayDataHistory[0][1],"-|:|\\s",",",0)
		// midArray = Split(midTime, ',') 
		// targetTime = Time.TimeSerial(CInt(midArray[0]),CInt(midArray[1]),CInt(midArray[2]),CInt(midArray[3]),CInt(midArray[4]),0)
		// // 计算差值
		// leftTime = Time.DateDiff("d",targetTime,nowTime)
		// // TracePrint(arrayData[y-1][0]&"为第"&y-1&"个候选人+------+最近一次打招呼时间："&leftTime)
		// arrRet = push(arrayData[y-1],leftTime)
		
		// #icon("@res:uhs4jjt2-ut6c-763u-uqiq-snutfinv099u.png")
		// Mouse.Action({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"I","css-selector":"body>div>div>div>a>i"}]},"left","click",10000,{"bContinueOnError":False,"iDelayAfter":300,"iDelayBefore":200,"bSetForeground":True,"sCursorPosition":"Center","iCursorOffsetX":0,"iCursorOffsetY":0,"sKeyModifiers":[],"sSimulate":"simulate","bMoveSmoothly":False})
		
		
		
		// TracePrint(arrayData[y-1])
		
		// arrayDataHello = UiElement.DataScrap({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"HTML"}]},{"ExtractTable":0,"Columns":[{"selecors":[{"tag":"div","index":0,"className":"dialog-wrap dialog-prop-default coop-record-wrap","value":"div.dialog-wrap.dialog-prop-default.coop-record-wrap","prefix":""},{"tag":"div","index":0,"className":"dialog-container","value":"div.dialog-container","prefix":">"},{"tag":"div","index":0,"className":"dialog-con","value":"div.dialog-con","prefix":">"},{"tag":"div","index":0,"className":"chat-record-content","value":"div.chat-record-content","prefix":">"},{"tag":"ul","index":0,"className":"record","value":"ul.record","prefix":">"},{"tag":"li","value":"li","index":0,"prefix":">"},{"tag":"div","index":0,"className":"","value":"div","prefix":">"},{"tag":"p","index":0,"className":"action","value":"p.action","prefix":">"}],"props":["text"]},{"selecors":[{"tag":"div","index":0,"className":"dialog-wrap dialog-prop-default coop-record-wrap","value":"div.dialog-wrap.dialog-prop-default.coop-record-wrap","prefix":""},{"tag":"div","index":0,"className":"dialog-container","value":"div.dialog-container","prefix":">"},{"tag":"div","index":0,"className":"dialog-con","value":"div.dialog-con","prefix":">"},{"tag":"div","index":0,"className":"chat-record-content","value":"div.chat-record-content","prefix":">"},{"tag":"ul","index":0,"className":"record","value":"ul.record","prefix":">"},{"tag":"li","value":"li","index":0,"prefix":">"},{"tag":"div","index":0,"className":"","value":"div","prefix":">"},{"tag":"p","index":0,"className":"operat","value":"p.operat","prefix":">"}],"props":["text"]}]},{"objNextLinkElement":"","iMaxNumberOfPage":5,"iMaxNumberOfResult":-1,"iDelayBetweenMS":1000,"bContinueOnError":False})
		
		// TracePrint(arrayDataHello &"arraydatahello---")
		
		// helloFather = Len(arrayDataHello)
		// helloSon = Len(arrayDataHello[0])
		
		// TracePrint(helloFather&"--"& helloSon &"历史招呼信息---")
		
		// For x = 1 To helloFather Step 1 
		
		// 	For k = 1 To helloSon Step 1 
		// 		Excel.WriteCell(objExcelWorkBook,"Sheet1",Chr(64+k)&x,arrayData[x-1][k-1],False)
		// 	Next
		// Next
		
		// For i = 1 To iRetSon Step 1 
		// [姓名, 活跃状态, 薪资范围, 年龄 工作年限, 求职期望, 求职简介, 历史工作公司 毕业学校]
		
		//1. 按钮文本是否为"继续沟通"
		#icon("@res:80me7dm2-gukr-6uo4-chh9-asd0pee45v95.png")
		sRet = UiElement.GetValue({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"IFRAME","name":"recommendFrame"},{"tag":"BUTTON","parentid":"recommend-list","idx":y-1}]},{"bContinueOnError":False,"iDelayAfter":300,"iDelayBefore":200})
		TracePrint(sRet& "--是否继续沟通")
		
		// 2. 姓名正则  排除含外包/外派人员
		findName = Regex.FindStr(arrayData[y-1][0],"外包|外派",0)
		TracePrint("源文本："&arrayData[y-1][0] &"匹配后显示："&findName& "--姓名")
		
		// 3. 工作年限 (大于等于三年)
		workYear = cint(DigitFromStr(Regex.FindStr(arrayData[y-1][3],'[0-9]+年',0)))
		TracePrint(workYear& "--工作年限")
		
		// 4. 求职期望(含前端/Javascript/HTML/web)
		expectJob =  Regex.FindStr(arrayData[y-1][4],'前端|(?i)javascript|html|(?i)web',0)
		TracePrint("源文本："& arrayData[y-1][4] &"匹配后显示："& expectJob & "--求职期望")
		
		// 5. 年龄（<=30）
		age =  cint(DigitFromStr(Regex.FindStr(arrayData[y-1][3],'[0-9]+岁',0)))
		TracePrint("源文本："& arrayData[y-1][3] &"匹配后显示："& age & "--年龄")
		
		// 6.求职简介
		summary = Regex.FindStr(arrayData[y-1][5],"外包|外派",0)
		TracePrint(summary& "--求职简介")
		
		// 7. 去除 "在职-暂不考虑"
		working = Regex.FindStr(arrayData[y-1][3],"在职-暂不考虑",0)
		TracePrint(working& "--求职简介")
		
		// 8.要求掌握react
		require_react = Regex.FindStr(arrayData[y-1][5],'(?i)react',0)
		TracePrint(findName&"--"&require_react& "--require_react")
		
		
		
		
		If sRet <> "继续沟通" And findName ="" And workYear >= 3 And expectJob <> "" And age <=30 And summary ="" And working ="" And require_react <> ""
			#icon("@res:hgpsadcd-cs3n-u360-4p6g-nbii4ii8tlkr.png")
			// Mouse.Hover({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"IFRAME","name":"recommendFrame"},{"tag":"BUTTON","parentid":"recommend-list","idx":y-1}]},10000,{"bContinueOnError":False,"iDelayAfter":300,"iDelayBefore":200,"bSetForeground":True,"sCursorPosition":"Center","iCursorOffsetX":0,"iCursorOffsetY":0,"sKeyModifiers":[],"sSimulate":"simulate","bMoveSmoothly":False})
			#icon("@res:a96v8b09-hmu1-vpf4-al11-lmevupf6sli4.png")
			Mouse.Action({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"IFRAME","name":"recommendFrame"},{"tag":"BUTTON","parentid":"recommend-list","idx":y-1}]},"left","click",10000,{"bContinueOnError":False,"iDelayAfter":300,"iDelayBefore":200,"bSetForeground":True,"sCursorPosition":"Center","iCursorOffsetX":0,"iCursorOffsetY":0,"sKeyModifiers":[],"sSimulate":"simulate","bMoveSmoothly":False})
			// TracePrint("已筛选候人选数量为："& y)
		Else
			// Break
			TracePrint(arrayData[y-1][0]&"：不符合条件----")
		End If
		// Excel.WriteCell(objExcelWorkBook,"Sheet1",Chr(64+i)&y,arrayData[y-1][i-1],False)
		// Next
	Next
	
	
	// For y = 1 To 1 Step 1 
	// 	#icon("@res:ir1stmts-aj2p-9vae-tir6-gnqogf566pdu.png")
	// 	Mouse.Hover({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"IFRAME","name":"recommendFrame"},{"tag":"BUTTON","parentid":"recommend-list","idx":2}]},10000,{"bContinueOnError":False,"iDelayAfter":300,"iDelayBefore":200,"bSetForeground":True,"sCursorPosition":"Center","iCursorOffsetX":0,"iCursorOffsetY":0,"sKeyModifiers":[],"sSimulate":"simulate","bMoveSmoothly":False})
	// Next
	#icon("@res:cn6nf0uv-eoel-aame-l4ie-13vnvh6ueslv.png")
	// Mouse.Action({"html":[{"parentid":"recommendContent","tag":"LI"}],"wnd":[{"app":"iexplore","cls":"IEFrame","title":"*"},{"cls":"Internet Explorer_Server"}]},"left","click",10000,{"bContinueOnError":False,"iDelayAfter":300,"iDelayBefore":200,"bSetForeground":True,"sCursorPosition":"Center","iCursorOffsetX":0,"iCursorOffsetY":0,"sKeyModifiers":[],"sSimulate":"simulate","bMoveSmoothly":False})
Next


