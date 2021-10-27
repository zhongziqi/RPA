```vbscript
// 搜索关键词
Dim arrSelItem = ""
Dim totalRet = ""
// 最小年龄
Dim minAge = 18
// 最大年龄
Dim maxAge = 30
// 最少工作年限
Dim workExperienceMin = 3
// 最多工作年限(可不修改)
Dim workExperienceMax = 15



// 期望最少工资 (单位:k)
Dim minSalary = 12
Dim maxSalary = 28 

Dim hWeb = ""
Dim iRet = ""
Dim arrayData = ""
Dim bRet = ""
Dim sRet = ""
Dim isScroll = False
Dim scrapWord = ""
Dim totalPage = 5
Dim freeRet = ""
Dim freeNum = ""
Dim findName = ""

hWeb = WebBrowser.Create("chrome","https://easy.lagou.com/dashboard/index.htm?from=c_index",30000,{"bContinueOnError":False,"iDelayAfter":300,"iDelayBefore":200,"sBrowserPath":"","sStartArgs":""})
#icon("@res:gs9r3ssj-gnmn-m4an-3r2j-ceq9lobbvqjo.png")
Mouse.Action({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"A","aaname":"找人才"}]},"left","click",10000,{"bContinueOnError":False,"iDelayAfter":300,"iDelayBefore":200,"bSetForeground":True,"sCursorPosition":"Center","iCursorOffsetX":0,"iCursorOffsetY":0,"sKeyModifiers":[],"sSimulate":"simulate","bMoveSmoothly":False})
#icon("@res:cl30j2pv-r8pq-mn5j-k34j-df58o0be47n9.png")
Keyboard.InputText({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"INPUT","id":"tagInput"}]},"web前端+react",True,300,10000,{"bContinueOnError":False,"iDelayAfter":1000,"iDelayBefore":200,"bSetForeground":True,"sSimulate":"simulate","bValidate":False,"bClickBeforeInput":True})
#icon("@res:c0ktg1k9-7dpo-ul3b-5mm4-kqk8pkk184ac.png")
Mouse.Action({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"BUTTON","parentid":"search-candidate"}]},"left","click",10000,{"bContinueOnError":False,"iDelayAfter":3000,"iDelayBefore":200,"bSetForeground":True,"sCursorPosition":"Center","iCursorOffsetX":0,"iCursorOffsetY":0,"sKeyModifiers":[],"sSimulate":"simulate","bMoveSmoothly":False})
// exit()
#icon("@res:faq8oebt-s08t-gsos-lm0g-rb06os0jkj4n.png")
Mouse.Action({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"DIV","parentid":"root","aaname":"最近在线"}]},"left","click",10000,{"bContinueOnError":False,"iDelayAfter":300,"iDelayBefore":200,"bSetForeground":True,"sCursorPosition":"Center","iCursorOffsetX":0,"iCursorOffsetY":0,"sKeyModifiers":[],"sSimulate":"simulate","bMoveSmoothly":False})
#icon("@res:u7d4v8s7-cm7u-le0n-mkog-kagn61jp8rnc.png")
Mouse.Action({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"DIV","parentid":"root","aaname":"专科及以上"}]},"left","click",10000,{"bContinueOnError":False,"iDelayAfter":300,"iDelayBefore":200,"bSetForeground":True,"sCursorPosition":"Center","iCursorOffsetX":0,"iCursorOffsetY":0,"sKeyModifiers":[],"sSimulate":"simulate","bMoveSmoothly":False})
#icon("@res:m8npm0mi-2fhq-8uor-jvu4-jirps58r2rcm.png")
Mouse.Action({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"DIV","parentid":"root","css-selector":"body>div>div>div>div>div>div>div>div>div>div>div>div>div>div"}]},"left","click",10000,{"bContinueOnError":False,"iDelayAfter":300,"iDelayBefore":200,"bSetForeground":True,"sCursorPosition":"Center","iCursorOffsetX":0,"iCursorOffsetY":0,"sKeyModifiers":[],"sSimulate":"simulate","bMoveSmoothly":False})
#icon("@res:169i9n5n-seq6-4ndg-p05d-92ot0j41n9ph.png")
UiElement.SetAttribute({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"DIV","parentid":"root","css-selector":"body>div>div>div>div>div>div>div>div>div>div>div>div>div:nth-child(3)>div>div>div"}]},"id","experience",{"bContinueOnError":False,"iDelayAfter":300,"iDelayBefore":200})
#icon("@res:6ib5pl1g-gcga-bkvj-ggsl-t83v2vgcqpiu.png")
Mouse.Action({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"LI","parentid":"experience","idx":workExperienceMin-1}]},"left","click",10000,{"bContinueOnError":False,"iDelayAfter":300,"iDelayBefore":200,"bSetForeground":True,"sCursorPosition":"Center","iCursorOffsetX":0,"iCursorOffsetY":0,"sKeyModifiers":[],"sSimulate":"simulate","bMoveSmoothly":False})
#icon("@res:uifictd4-fq4c-r0u2-q2gv-bt5k851q9i65.png")
Mouse.Action({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"DIV","parentid":"root","css-selector":"body>div>div>div>div>div>div>div>div>div>div>div>div>div>div","idx":1}]},"left","click",10000,{"bContinueOnError":False,"iDelayAfter":300,"iDelayBefore":200,"bSetForeground":True,"sCursorPosition":"Center","iCursorOffsetX":0,"iCursorOffsetY":0,"sKeyModifiers":[],"sSimulate":"simulate","bMoveSmoothly":False})
#icon("@res:169i9n5n-seq6-4ndg-p05d-92ot0j41n9ph.png")
UiElement.SetAttribute({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"DIV","parentid":"root","css-selector":"body>div>div>div>div>div>div>div>div>div>div>div>div>div:nth-child(4)>div>div>div"}]},"id","experienceMax",{"bContinueOnError":False,"iDelayAfter":300,"iDelayBefore":200})
#icon("@res:6ib5pl1g-gcga-bkvj-ggsl-t83v2vgcqpiu.png")
Mouse.Action({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"LI","parentid":"experienceMax","idx":workExperienceMax-2}]},"left","click",10000,{"bContinueOnError":False,"iDelayAfter":300,"iDelayBefore":500,"bSetForeground":True,"sCursorPosition":"Center","iCursorOffsetX":0,"iCursorOffsetY":0,"sKeyModifiers":[],"sSimulate":"simulate","bMoveSmoothly":False})
#icon("@res:ktnrr3pq-kvle-62qo-4sp8-frvoejau1h47.png")
Mouse.Action({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"DIV","parentid":"root","css-selector":"body>div>div>div>div>div>div>div>div>div>div>div>div>div>div","tabindex":"0","idx":2}]},"left","click",10000,{"bContinueOnError":False,"iDelayAfter":300,"iDelayBefore":200,"bSetForeground":True,"sCursorPosition":"Center","iCursorOffsetX":0,"iCursorOffsetY":0,"sKeyModifiers":[],"sSimulate":"simulate","bMoveSmoothly":False})
#icon("@res:169i9n5n-seq6-4ndg-p05d-92ot0j41n9ph.png")
UiElement.SetAttribute({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"DIV","parentid":"root","css-selector":"body>div>div>div>div>div>div>div>div>div:nth-child(7)>div>div>div>div:nth-child(3)>div>div>div"}]},"id","minSalary",{"bContinueOnError":False,"iDelayAfter":300,"iDelayBefore":200})
#icon("@res:6ib5pl1g-gcga-bkvj-ggsl-t83v2vgcqpiu.png")
Mouse.Action({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"LI","parentid":"minSalary","idx":minSalary-1}]},"left","click",10000,{"bContinueOnError":False,"iDelayAfter":300,"iDelayBefore":500,"bSetForeground":True,"sCursorPosition":"Center","iCursorOffsetX":0,"iCursorOffsetY":0,"sKeyModifiers":[],"sSimulate":"simulate","bMoveSmoothly":False})
#icon("@res:ga877g1q-nt0h-fifh-aat3-o4r7pa5pq4ja.png")
Mouse.Action({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"DIV","parentid":"root","css-selector":"body>div>div>div>div>div>div>div>div>div>div>div>div>div>div","tabindex":"0","idx":3}]},"left","click",10000,{"bContinueOnError":False,"iDelayAfter":300,"iDelayBefore":200,"bSetForeground":True,"sCursorPosition":"Center","iCursorOffsetX":0,"iCursorOffsetY":0,"sKeyModifiers":[],"sSimulate":"simulate","bMoveSmoothly":False})
#icon("@res:169i9n5n-seq6-4ndg-p05d-92ot0j41n9ph.png")
UiElement.SetAttribute({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"DIV","parentid":"root","css-selector":"body>div>div>div>div>div>div>div>div>div:nth-child(7)>div>div>div>div:nth-child(4)>div>div>div"}]},"id","maxSalary",{"bContinueOnError":False,"iDelayAfter":300,"iDelayBefore":200})
#icon("@res:6ib5pl1g-gcga-bkvj-ggsl-t83v2vgcqpiu.png")
Mouse.Action({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"LI","parentid":"maxSalary","idx":maxSalary-2}]},"left","click",10000,{"bContinueOnError":False,"iDelayAfter":300,"iDelayBefore":500,"bSetForeground":True,"sCursorPosition":"Center","iCursorOffsetX":0,"iCursorOffsetY":0,"sKeyModifiers":[],"sSimulate":"simulate","bMoveSmoothly":False})
#icon("@res:hbrf99s4-bljq-101c-22h2-8vhtpmhtsds2.png")
Mouse.Action({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"DIV","parentid":"root","css-selector":"body>div>div>div>div>div>div>div>div>div>div>div>div>div>div","tabindex":"0","idx":4}]},"left","click",10000,{"bContinueOnError":False,"iDelayAfter":300,"iDelayBefore":200,"bSetForeground":True,"sCursorPosition":"Center","iCursorOffsetX":0,"iCursorOffsetY":0,"sKeyModifiers":[],"sSimulate":"simulate","bMoveSmoothly":False})
#icon("@res:169i9n5n-seq6-4ndg-p05d-92ot0j41n9ph.png")
UiElement.SetAttribute({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"DIV","parentid":"root","css-selector":"body>div>div>div>div>div>div>div>div>div:nth-child(9)>div>div>div>div:nth-child(3)>div>div>div"}]},"id","minAge",{"bContinueOnError":False,"iDelayAfter":300,"iDelayBefore":200})
#icon("@res:6ib5pl1g-gcga-bkvj-ggsl-t83v2vgcqpiu.png")
Mouse.Action({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"LI","parentid":"minAge","idx":minAge-16}]},"left","click",10000,{"bContinueOnError":False,"iDelayAfter":300,"iDelayBefore":500,"bSetForeground":True,"sCursorPosition":"Center","iCursorOffsetX":0,"iCursorOffsetY":0,"sKeyModifiers":[],"sSimulate":"simulate","bMoveSmoothly":False})
#icon("@res:7431bmqu-3fl2-a0mm-9viq-q8gou7qcc9io.png")
Mouse.Action({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"DIV","parentid":"root","css-selector":"body>div>div>div>div>div>div>div>div>div>div>div>div>div>div","tabindex":"0","idx":5}]},"left","click",10000,{"bContinueOnError":False,"iDelayAfter":300,"iDelayBefore":200,"bSetForeground":True,"sCursorPosition":"Center","iCursorOffsetX":0,"iCursorOffsetY":0,"sKeyModifiers":[],"sSimulate":"simulate","bMoveSmoothly":False})
#icon("@res:169i9n5n-seq6-4ndg-p05d-92ot0j41n9ph.png")
UiElement.SetAttribute({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"DIV","parentid":"root","css-selector":"body>div>div>div>div>div>div>div>div>div:nth-child(9)>div>div>div>div:nth-child(4)>div>div>div"}]},"id","maxAge",{"bContinueOnError":False,"iDelayAfter":300,"iDelayBefore":200})
#icon("@res:6ib5pl1g-gcga-bkvj-ggsl-t83v2vgcqpiu.png")
Mouse.Action({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"LI","parentid":"maxAge","idx":maxAge-16}]},"left","click",10000,{"bContinueOnError":False,"iDelayAfter":300,"iDelayBefore":500,"bSetForeground":True,"sCursorPosition":"Center","iCursorOffsetX":0,"iCursorOffsetY":0,"sKeyModifiers":[],"sSimulate":"simulate","bMoveSmoothly":False})
#icon("@res:idnfcnsd-1u95-vcin-9je1-f08l403nkrd0.png")

// 获取总页数
totalPage = Text.Get({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"A","parentid":"root","css-selector":"body>div>div>div>div>div>div>div>div>ul>li:nth-last-child(2)>a"}]},10000,{"bContinueOnError":False,"iDelayAfter":2000,"iDelayBefore":2000,"bSetForeground":True})
TracePrint(totalPage)

// TracePrint(totalPage)
// 循环总页数
For x = 1 To totalPage Step 1
	Dim y = "'" & x & "'"  
	// 滚动到底部加载当前页面全部数据
	Do While isScroll= False 
		#icon("@res:2iq56mie-dre5-rtqm-7ff1-6ighrteohcbf.png")
		isScroll = UiElement.Exists({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"UL","parentid":"root"}]},{"bContinueOnError":False,"iDelayAfter":300,"iDelayBefore":200})
		Mouse.Wheel(50,"down", [],{"iDelayAfter":1000,"iDelayBefore":2000})
	Loop
	#icon("@res:0tn0cps8-6j2e-dqc6-u742-4j1kfn3sr8mj.png")
	TracePrint("body>div>div>div>div>div>div.search-content-container>div>div>ul.lg-pagination>li[title="& y &"]>a")
	
	Mouse.Action({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"A","css-selector":"body>div>div>div>div>div>div>div>div>ul>li[title="& y &"]>a"}]},"left","click",10000,{"bContinueOnError":False,"iDelayAfter":4000,"iDelayBefore":2000,"bSetForeground":True,"sCursorPosition":"Center","iCursorOffsetX":0,"iCursorOffsetY":0,"sKeyModifiers":[],"sSimulate":"simulate","bMoveSmoothly":False})
	
	// 爬取当前页面数据
	arrayData = UiElement.DataScrap({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"DIV","id":"root"}]},{"ExtractTable":0,"Columns":[{"selecors":[{"tag":"div","index":0,"className":"wide-container","value":"div.wide-container","prefix":""},{"tag":"div","index":0,"className":"wide-container-inner","value":"div.wide-container-inner","prefix":">"},{"tag":"div","index":5,"className":"","value":"div:nth-child(5)","prefix":">"},{"tag":"div","index":0,"className":"search-container","value":"div.search-container","prefix":">"},{"tag":"div","index":0,"className":"search-content-container","value":"div.search-content-container","prefix":">"},{"tag":"div","index":2,"className":"","value":"div:nth-child(2)","prefix":">"},{"tag":"div","index":1,"className":"","value":"div:nth-child(1)","prefix":">"},{"tag":"div","index":0,"className":"talent-list","value":"div.talent-list","prefix":">"},{"tag":"div","value":"div","index":0,"prefix":">"},{"tag":"div","index":0,"className":"talent-item","value":"div.talent-item","prefix":">"},{"tag":"div","index":0,"className":"talent-item-top","value":"div.talent-item-top","prefix":">"}],"props":["text"]},{"selecors":[{"tag":"div","index":0,"className":"wide-container","value":"div.wide-container","prefix":""},{"tag":"div","index":0,"className":"wide-container-inner","value":"div.wide-container-inner","prefix":">"},{"tag":"div","index":5,"className":"","value":"div:nth-child(5)","prefix":">"},{"tag":"div","index":0,"className":"search-container","value":"div.search-container","prefix":">"},{"tag":"div","index":0,"className":"search-content-container","value":"div.search-content-container","prefix":">"},{"tag":"div","index":2,"className":"","value":"div:nth-child(2)","prefix":">"},{"tag":"div","index":1,"className":"","value":"div:nth-child(1)","prefix":">"},{"tag":"div","index":0,"className":"talent-list","value":"div.talent-list","prefix":">"},{"tag":"div","value":"div","index":0,"prefix":">"},{"tag":"div","index":0,"className":"talent-item","value":"div.talent-item","prefix":">"},{"tag":"div","index":0,"className":"talent-item-content","value":"div.talent-item-content","prefix":">"},{"tag":"div","index":0,"className":"user-opt","value":"div.user-opt","prefix":">"},{"tag":"div","index":0,"className":"item-user","value":"div.item-user","prefix":">"},{"tag":"div","index":0,"className":"item-user-txt","value":"div.item-user-txt","prefix":">"},{"tag":"div","index":0,"className":"name","value":"div.name","prefix":">"}],"props":["text"]},{"selecors":[{"tag":"div","index":0,"className":"wide-container","value":"div.wide-container","prefix":""},{"tag":"div","index":0,"className":"wide-container-inner","value":"div.wide-container-inner","prefix":">"},{"tag":"div","index":5,"className":"","value":"div:nth-child(5)","prefix":">"},{"tag":"div","index":0,"className":"search-container","value":"div.search-container","prefix":">"},{"tag":"div","index":0,"className":"search-content-container","value":"div.search-content-container","prefix":">"},{"tag":"div","index":2,"className":"","value":"div:nth-child(2)","prefix":">"},{"tag":"div","index":1,"className":"","value":"div:nth-child(1)","prefix":">"},{"tag":"div","index":0,"className":"talent-list","value":"div.talent-list","prefix":">"},{"tag":"div","value":"div","index":0,"prefix":">"},{"tag":"div","index":0,"className":"talent-item","value":"div.talent-item","prefix":">"},{"tag":"div","index":0,"className":"talent-item-content","value":"div.talent-item-content","prefix":">"},{"tag":"div","index":0,"className":"user-opt","value":"div.user-opt","prefix":">"},{"tag":"div","index":0,"className":"item-user","value":"div.item-user","prefix":">"},{"tag":"div","index":0,"className":"item-user-txt","value":"div.item-user-txt","prefix":">"}],"props":["text"]},{"selecors":[{"tag":"div","index":0,"className":"wide-container","value":"div.wide-container","prefix":""},{"tag":"div","index":0,"className":"wide-container-inner","value":"div.wide-container-inner","prefix":">"},{"tag":"div","index":5,"className":"","value":"div:nth-child(5)","prefix":">"},{"tag":"div","index":0,"className":"search-container","value":"div.search-container","prefix":">"},{"tag":"div","index":0,"className":"search-content-container","value":"div.search-content-container","prefix":">"},{"tag":"div","index":2,"className":"","value":"div:nth-child(2)","prefix":">"},{"tag":"div","index":1,"className":"","value":"div:nth-child(1)","prefix":">"},{"tag":"div","index":0,"className":"talent-list","value":"div.talent-list","prefix":">"},{"tag":"div","value":"div","index":0,"prefix":">"},{"tag":"div","index":0,"className":"talent-item","value":"div.talent-item","prefix":">"},{"tag":"div","index":0,"className":"talent-item-content","value":"div.talent-item-content","prefix":">"},{"tag":"div","index":0,"className":"labels","value":"div.labels","prefix":">"}],"props":["text"]},{"selecors":[{"tag":"div","index":0,"className":"wide-container","value":"div.wide-container","prefix":""},{"tag":"div","index":0,"className":"wide-container-inner","value":"div.wide-container-inner","prefix":">"},{"tag":"div","index":5,"className":"","value":"div:nth-child(5)","prefix":">"},{"tag":"div","index":0,"className":"search-container","value":"div.search-container","prefix":">"},{"tag":"div","index":0,"className":"search-content-container","value":"div.search-content-container","prefix":">"},{"tag":"div","index":2,"className":"","value":"div:nth-child(2)","prefix":">"},{"tag":"div","index":1,"className":"","value":"div:nth-child(1)","prefix":">"},{"tag":"div","index":0,"className":"talent-list","value":"div.talent-list","prefix":">"},{"tag":"div","value":"div","index":0,"prefix":">"},{"tag":"div","index":0,"className":"talent-item","value":"div.talent-item","prefix":">"},{"tag":"div","index":0,"className":"talent-item-content","value":"div.talent-item-content","prefix":">"},{"tag":"div","index":0,"className":"info-list","value":"div.info-list","prefix":">"},{"tag":"div","index":0,"className":"item-info-work ","value":"div.item-info-work","prefix":">"},{"tag":"div","index":0,"className":"item-stick-container","value":"div.item-stick-container","prefix":">"},{"tag":"div","index":0,"className":"content","value":"div.content","prefix":">"}],"props":["text"]},{"selecors":[{"tag":"div","index":0,"className":"wide-container","value":"div.wide-container","prefix":""},{"tag":"div","index":0,"className":"wide-container-inner","value":"div.wide-container-inner","prefix":">"},{"tag":"div","index":5,"className":"","value":"div:nth-child(5)","prefix":">"},{"tag":"div","index":0,"className":"search-container","value":"div.search-container","prefix":">"},{"tag":"div","index":0,"className":"search-content-container","value":"div.search-content-container","prefix":">"},{"tag":"div","index":2,"className":"","value":"div:nth-child(2)","prefix":">"},{"tag":"div","index":1,"className":"","value":"div:nth-child(1)","prefix":">"},{"tag":"div","index":0,"className":"talent-list","value":"div.talent-list","prefix":">"},{"tag":"div","value":"div","index":0,"prefix":">"},{"tag":"div","index":0,"className":"talent-item","value":"div.talent-item","prefix":">"},{"tag":"div","value":"div","index":0,"prefix":">"},{"tag":"div","index":0,"className":"info-list","value":"div.info-list","prefix":">"},{"tag":"div","index":0,"className":"item-info-edu ","value":"div.item-info-edu","prefix":">"},{"tag":"div","value":"div","index":0,"prefix":">"},{"tag":"div","index":0,"className":"content","value":"div.content","prefix":">"}],"props":["text"]}]},{"objNextLinkElement":"","iMaxNumberOfPage":1,"iMaxNumberOfResult":-1,"iDelayBetweenMS":1000,"bContinueOnError":False})
	iRet = Len(arrayData)
	
	TracePrint(iRet)
	
	// 循环当前页面数据
	For i = 1 To iRet Step 1 
		#icon("@res:eq1cam7s-173e-7rp1-tu40-h1o7a9mt5q2l.png")
		// 0. 求职意向(eg:职意向：前端工程师深圳6k-8k\n4天前来过)
		// 1. 姓名正则  排除含外包/外派人员
		// 2. 工作年限等(eg:吴家卿\n1年工作经验大专21岁男)\
		// 3. 技能(eb前端开发VueReact)
		// 4. 历史任职公司(深圳华距同创WEB前端\nIT技术服务｜咨询)
		// 5. 所在学校(广西理工职业技术学院计算机科学与技术 | 大专)
		findName = Regex.FindStr(arrayData[i-1][1],"外包|外派",0)
		
		scrapWord = Text.Get({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"DIV","css-selector":"body>div>div>div>div>div>div>div>div>div>div>div>div>div>div>div.chat.operation","idx":i-1}]},10000,{"bContinueOnError":False,"iDelayAfter":2000,"iDelayBefore":2000,"bSetForeground":True})
		// TracePrint(scrapWord,"scrapword")
		If findName = "" And scrapWord ="打招呼"
			Mouse.Action({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"DIV","css-selector":"body>div>div>div>div>div>div>div>div>div>div>div>div>div>div>div.chat.operation","idx":i-1}]},"left","click",10000,{"bContinueOnError":False,"iDelayAfter":3000,"iDelayBefore":1000,"bSetForeground":True,"sCursorPosition":"Center","iCursorOffsetX":0,"iCursorOffsetY":0,"sKeyModifiers":[],"sSimulate":"simulate","bMoveSmoothly":False})			
			
			TracePrint("源文本："&arrayData[i-1][1] &"匹配后显示："&findName& "--姓名  符合条件")
			#icon("@res:8ioq9hpf-1kfp-j7i0-e09a-629hfi3sor92.png")
			Mouse.Action({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"BUTTON","aaname":"发送后留在此页"}]},"left","click",10000,{"bContinueOnError":False,"iDelayAfter":1000,"iDelayBefore":200,"bSetForeground":True,"sCursorPosition":"Center","iCursorOffsetX":0,"iCursorOffsetY":0,"sKeyModifiers":[],"sSimulate":"simulate","bMoveSmoothly":False})
			#icon("@res:9s4mqjeo-tq0d-1te1-amm0-qiq0o7030f1n.png")
			// Mouse.Hover({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"DIV","css-selector":"body>div>div>div>div>div>div>div>div>div>div>div>div>div>div>div.chat.operation","idx":i-1}]},10000,{"bContinueOnError":False,"iDelayAfter":300,"iDelayBefore":200,"bSetForeground":True,"sCursorPosition":"Center","iCursorOffsetX":0,"iCursorOffsetY":0,"sKeyModifiers":[],"sSimulate":"simulate","bMoveSmoothly":False})
		End If
		TracePrint("源文本："&arrayData[i-1][1] &"匹配后显示："&findName& "--姓名  符合条件")
		
	Next
Next


```

