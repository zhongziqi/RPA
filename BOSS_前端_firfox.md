```vbscript

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
Dim certificate =""
Dim active_status = ""

hWeb = WebBrowser.Create("firefox","https://www.zhipin.com/web/boss/index",30000,{"bContinueOnError":False,"iDelayAfter":5000,"iDelayBefore":200,"sBrowserPath":"","sStartArgs":""})
#icon("@res:9h98cr8i-r09n-2ndk-cgum-f3732ib84b9a.png")
// Mouse.Action({"wnd":[{"cls":"MozillaWindowClass","title":"*","app":"firefox"}],"html":[{"tag":"A","type":"desktop","parentid":"main","css-selector":"body>div>div>div>div>dl.menu-recommend>dt>a[ka='menu-geek-recommend']"}]},"left","click",10000,{"bContinueOnError":False,"iDelayAfter":3000,"iDelayBefore":200,"bSetForeground":True,"sCursorPosition":"Center","iCursorOffsetX":0,"iCursorOffsetY":0,"sKeyModifiers":[],"sSimulate":"simulate","bMoveSmoothly":False})
Mouse.Action({"wnd":[{"cls":"MozillaWindowClass","title":"*","app":"firefox"}],"html":[{"tag":"A","type":"desktop","parentid":"main","css-selector":"body>div>div>div>div>dl.menu-recommend>dt>a[ka='menu-geek-recommend']"}]},"left","click",10000,{"bContinueOnError":False,"iDelayAfter":3000,"iDelayBefore":200,"bSetForeground":True,"sCursorPosition":"Center","iCursorOffsetX":0,"iCursorOffsetY":0,"sKeyModifiers":[],"sSimulate":"simulate","bMoveSmoothly":False})


#icon("@res:6jq66rov-rvgl-3eea-o5ke-0qk1gdm0alds.png")
Mouse.Action({"wnd":[{"cls":"MozillaWindowClass","title":"*","app":"firefox"}],"html":[{"tag":"IFRAME","name":"recommendFrame"},{"tag":"I","css-selector":"body>div>div>div>div>div>div>div>div>div>div>div>i"}]},"left","click",10000,{"bContinueOnError":False,"iDelayAfter":1000,"iDelayBefore":200,"bSetForeground":True,"sCursorPosition":"Center","iCursorOffsetX":0,"iCursorOffsetY":0,"sKeyModifiers":[],"sSimulate":"simulate","bMoveSmoothly":False})
arrayData = UiElement.DataScrap({"wnd":[{"cls":"MozillaWindowClass","title":"*","app":"firefox"}],"html":[{"tag":"IFRAME","name":"recommendFrame"},{"tag":"DIV","id":"recommend-list"}]},{"ExtractTable":0,"Columns":[{"selecors":[{"tag":"div","index":0,"className":"sec-content candidate-card","value":"div.sec-content.candidate-card","prefix":""},{"tag":"ul","index":0,"className":"recommend-card-list ul-less-height","value":"ul.recommend-card-list.ul-less-height","prefix":">"},{"tag":"li","value":"li","index":0,"prefix":">"},{"tag":"div","index":0,"className":"","value":"div","prefix":">"},{"tag":"div","index":0,"className":"geek-info-card less-height","value":"div.geek-info-card.less-height","prefix":">"},{"tag":"div","index":0,"className":"candidate-list-content","value":"div.candidate-list-content","prefix":">"},{"tag":"div","index":0,"className":"card-inner","value":"div.card-inner","prefix":">"},{"tag":"div","index":0,"className":"col-2","value":"div.col-2","prefix":">"},{"tag":"div","index":0,"className":"name","value":"div.name","prefix":">"}],"props":["text"]},{"selecors":[{"tag":"div","index":0,"className":"sec-content candidate-card","value":"div.sec-content.candidate-card","prefix":""},{"tag":"ul","index":0,"className":"recommend-card-list ul-less-height","value":"ul.recommend-card-list.ul-less-height","prefix":">"},{"tag":"li","value":"li","index":0,"prefix":">"},{"tag":"div","index":0,"className":"","value":"div","prefix":">"},{"tag":"div","index":0,"className":"geek-info-card less-height","value":"div.geek-info-card.less-height","prefix":">"},{"tag":"div","index":0,"className":"candidate-list-content","value":"div.candidate-list-content","prefix":">"},{"tag":"div","index":0,"className":"card-inner","value":"div.card-inner","prefix":">"},{"tag":"div","index":0,"className":"col-2","value":"div.col-2","prefix":">"},{"tag":"div","index":0,"className":"name","value":"div.name","prefix":">"},{"tag":"span","index":0,"className":"label-text","value":"span.label-text","prefix":">"}],"props":["text"]},{"selecors":[{"tag":"div","index":0,"className":"sec-content candidate-card","value":"div.sec-content.candidate-card","prefix":""},{"tag":"ul","index":0,"className":"recommend-card-list ul-less-height","value":"ul.recommend-card-list.ul-less-height","prefix":">"},{"tag":"li","value":"li","index":0,"prefix":">"},{"tag":"div","index":0,"className":"","value":"div","prefix":">"},{"tag":"div","index":0,"className":"geek-info-card less-height","value":"div.geek-info-card.less-height","prefix":">"},{"tag":"div","index":0,"className":"candidate-list-content","value":"div.candidate-list-content","prefix":">"},{"tag":"div","index":0,"className":"card-inner","value":"div.card-inner","prefix":">"},{"tag":"div","index":0,"className":"col-1","value":"div.col-1","prefix":">"},{"tag":"div","index":0,"className":"salary","value":"div.salary","prefix":">"}],"props":["text"]},{"selecors":[{"tag":"div","index":0,"className":"sec-content candidate-card","value":"div.sec-content.candidate-card","prefix":""},{"tag":"ul","index":0,"className":"recommend-card-list ul-less-height","value":"ul.recommend-card-list.ul-less-height","prefix":">"},{"tag":"li","value":"li","index":0,"prefix":">"},{"tag":"div","index":0,"className":"","value":"div","prefix":">"},{"tag":"div","index":0,"className":"geek-info-card less-height","value":"div.geek-info-card.less-height","prefix":">"},{"tag":"div","index":0,"className":"candidate-list-content","value":"div.candidate-list-content","prefix":">"},{"tag":"div","index":0,"className":"card-inner","value":"div.card-inner","prefix":">"},{"tag":"div","index":0,"className":"col-2","value":"div.col-2","prefix":">"},{"tag":"div","index":0,"className":"info-labels","value":"div.info-labels","prefix":">"}],"props":["text"]},{"selecors":[{"tag":"div","index":0,"className":"sec-content candidate-card","value":"div.sec-content.candidate-card","prefix":""},{"tag":"ul","index":0,"className":"recommend-card-list ul-less-height","value":"ul.recommend-card-list.ul-less-height","prefix":">"},{"tag":"li","value":"li","index":0,"prefix":">"},{"tag":"div","index":0,"className":"","value":"div","prefix":">"},{"tag":"div","index":0,"className":"geek-info-card less-height","value":"div.geek-info-card.less-height","prefix":">"},{"tag":"div","index":0,"className":"candidate-list-content","value":"div.candidate-list-content","prefix":">"},{"tag":"div","index":0,"className":"card-inner","value":"div.card-inner","prefix":">"},{"tag":"div","index":0,"className":"col-2","value":"div.col-2","prefix":">"},{"tag":"div","index":0,"className":"expect-box","value":"div.expect-box","prefix":">"}],"props":["text"]},{"selecors":[{"tag":"div","index":0,"className":"sec-content candidate-card","value":"div.sec-content.candidate-card","prefix":""},{"tag":"ul","index":0,"className":"recommend-card-list ul-less-height","value":"ul.recommend-card-list.ul-less-height","prefix":">"},{"tag":"li","value":"li","index":0,"prefix":">"},{"tag":"div","index":0,"className":"","value":"div","prefix":">"},{"tag":"div","index":0,"className":"geek-info-card less-height","value":"div.geek-info-card.less-height","prefix":">"},{"tag":"div","index":0,"className":"candidate-list-content","value":"div.candidate-list-content","prefix":">"},{"tag":"div","index":0,"className":"card-inner","value":"div.card-inner","prefix":">"},{"tag":"div","index":0,"className":"col-2","value":"div.col-2","prefix":">"},{"tag":"div","index":0,"className":"position-box position-advantage","value":"div.position-box.position-advantage","prefix":">"},{"tag":"div","index":0,"className":"advantage-new","value":"div.advantage-new","prefix":">"}],"props":["text"]},{"selecors":[{"tag":"div","index":0,"className":"sec-content candidate-card","value":"div.sec-content.candidate-card","prefix":""},{"tag":"ul","index":0,"className":"recommend-card-list ul-less-height","value":"ul.recommend-card-list.ul-less-height","prefix":">"},{"tag":"li","value":"li","index":0,"prefix":">"},{"tag":"div","index":0,"className":"","value":"div","prefix":">"},{"tag":"div","index":0,"className":"geek-info-card less-height","value":"div.geek-info-card.less-height","prefix":">"},{"tag":"div","index":0,"className":"candidate-list-content","value":"div.candidate-list-content","prefix":">"},{"tag":"div","index":0,"className":"card-inner","value":"div.card-inner","prefix":">"},{"tag":"div","index":0,"className":"col-3","value":"div.col-3","prefix":">"}],"props":["text"]}]},{"objNextLinkElement":"","iMaxNumberOfPage":5,"iMaxNumberOfResult":-1,"iDelayBetweenMS":1000,"bContinueOnError":False})
#icon("@res:7gsnr4mf-bkui-2osh-ds5p-4pvrkohqv5h9.png")
Mouse.Action({"wnd":[{"cls":"MozillaWindowClass","title":"*","app":"firefox"}],"html":[{"tag":"IFRAME","name":"recommendFrame"},{"tag":"I","css-selector":"body>div>div>div>div>div>div>div>div>div>div>div>i"}]},"left","click",10000,{"bContinueOnError":False,"iDelayAfter":1000,"iDelayBefore":200,"bSetForeground":True,"sCursorPosition":"Center","iCursorOffsetX":0,"iCursorOffsetY":0,"sKeyModifiers":[],"sSimulate":"simulate","bMoveSmoothly":False})

Do While True
    bRet= False
    For i = 1 To 1 Step 1
        #icon("@res:ktuvajta-1jpm-6u4q-kf2k-satmaepiijag.png")
        Mouse.Action({"wnd":[{"cls":"MozillaWindowClass","title":"*","app":"firefox"}],"html":[{"tag":"IFRAME","name":"recommendFrame"},{"tag":"I","css-selector":"body>div>div>div>div>div>div>div>div>div>div>div>i"}]},"left","click",10000,{"bContinueOnError":False,"iDelayAfter":1000,"iDelayBefore":200,"bSetForeground":True,"sCursorPosition":"Center","iCursorOffsetX":0,"iCursorOffsetY":0,"sKeyModifiers":[],"sSimulate":"simulate","bMoveSmoothly":False})
        TracePrint(i & ':i---')
    ???#icon("@res:coceg44o-um5m-ja1k-2d04-ml9v7g8dtkds.png")
    ????Keyboard.InputText({"wnd":[{"cls":"MozillaWindowClass","title":"*","app":"firefox"}],"html":[{"tag":"INPUT","css-selector":"body>div>div>div>div>div>div>div>div>div>div>div.ui-dropmenu-list>div>input"}]},"前端",True,20,10000,{"bContinueOnError":False,"iDelayAfter":300,"iDelayBefore":500,"bSetForeground":True,"sSimulate":"message","bValidate":False,"bClickBeforeInput":False})
        // ui-dropmenu-label
        #icon("@res:cvn4vcp4-vo47-0pnj-dp69-h0qmdcn8k32e.png")
        // Mouse.Action({"html":[{"parentid":"recommendContent","tag":"LI" ,"attrMap":{"css-selector":"div.ui-dropmenu-list>div>ul>li:nth-child("& i &")"}}],"wnd":[{"app":"iexplore","cls":"IEFrame","title":"*"},{"cls":"Internet Explorer_Server"}]},"left","click",10000,{"bContinueOnError":False,"iDelayAfter":5000,"iDelayBefore":1000,"bSetForeground":True,"sCursorPosition":"Center","iCursorOffsetX":0,"iCursorOffsetY":0,"sKeyModifiers":[],"sSimulate":"simulate","bMoveSmoothly":False})
        Mouse.Action({"wnd":[{"cls":"MozillaWindowClass","title":"*","app":"firefox"}],"html":[{"tag":"IFRAME","name":"recommendFrame"},{"tag":"LI","css-selector":"body>div>div>div>div>div>div>div>div>div>div>div>ul>li","idx":0}]},"left","click",10000,{"bContinueOnError":False,"iDelayAfter":3000,"iDelayBefore":200,"bSetForeground":True,"sCursorPosition":"Center","iCursorOffsetX":0,"iCursorOffsetY":0,"sKeyModifiers":[],"sSimulate":"simulate","bMoveSmoothly":False})

        // 获取推荐的所有候选人数据--调试时注释
        Do While bRet= False
            #icon("@res:default.png")
            bRet = Text.Exists({"wnd":[{"cls":"MozillaWindowClass","title":"*","app":"firefox"}],"html":[{"tag":"IFRAME","name":"recommendFrame"},{"tag":"DIV","parentid":"recommend-list"}]},"没有更多","instr",1,10000,{"bContinueOnError":False,"iDelayAfter":300,"iDelayBefore":200,"bSetForeground":True})
            TracePrint( bRet)
            Mouse.Wheel(20,"down", [],{"iDelayAfter":1000,"iDelayBefore":1000})
            // TracePrint(bRet)
        Loop

        // 使用中的爬虫代码
        arrayData = UiElement.DataScrap({"wnd":[{"cls":"MozillaWindowClass","title":"*","app":"firefox"}],"html":[{"tag":"IFRAME","name":"recommendFrame"},{"tag":"DIV","id":"recommend-list"}]},{"ExtractTable":0,"Columns":[{"selecors":[{"tag":"div","index":0,"className":"sec-content candidate-card","value":"div.sec-content.candidate-card","prefix":""},{"tag":"ul","index":0,"className":"recommend-card-list ul-less-height","value":"ul.recommend-card-list.ul-less-height","prefix":">"},{"tag":"li","value":"li","index":0,"prefix":">"},{"tag":"div","index":0,"className":"","value":"div","prefix":">"},{"tag":"div","index":0,"className":"geek-info-card less-height","value":"div.geek-info-card.less-height","prefix":">"},{"tag":"div","index":0,"className":"candidate-list-content","value":"div.candidate-list-content","prefix":">"},{"tag":"div","index":0,"className":"card-inner","value":"div.card-inner","prefix":">"},{"tag":"div","index":0,"className":"col-2","value":"div.col-2","prefix":">"},{"tag":"div","index":0,"className":"name","value":"div.name","prefix":">"}],"props":["text"]},{"selecors":[{"tag":"div","index":0,"className":"sec-content candidate-card","value":"div.sec-content.candidate-card","prefix":""},{"tag":"ul","index":0,"className":"recommend-card-list ul-less-height","value":"ul.recommend-card-list.ul-less-height","prefix":">"},{"tag":"li","value":"li","index":0,"prefix":">"},{"tag":"div","index":0,"className":"","value":"div","prefix":">"},{"tag":"div","index":0,"className":"geek-info-card less-height","value":"div.geek-info-card.less-height","prefix":">"},{"tag":"div","index":0,"className":"candidate-list-content","value":"div.candidate-list-content","prefix":">"},{"tag":"div","index":0,"className":"card-inner","value":"div.card-inner","prefix":">"},{"tag":"div","index":0,"className":"col-2","value":"div.col-2","prefix":">"},{"tag":"div","index":0,"className":"name","value":"div.name","prefix":">"},{"tag":"span","index":0,"className":"label-text","value":"span.label-text","prefix":">"}],"props":["text"]},{"selecors":[{"tag":"div","index":0,"className":"sec-content candidate-card","value":"div.sec-content.candidate-card","prefix":""},{"tag":"ul","index":0,"className":"recommend-card-list ul-less-height","value":"ul.recommend-card-list.ul-less-height","prefix":">"},{"tag":"li","value":"li","index":0,"prefix":">"},{"tag":"div","index":0,"className":"","value":"div","prefix":">"},{"tag":"div","index":0,"className":"geek-info-card less-height","value":"div.geek-info-card.less-height","prefix":">"},{"tag":"div","index":0,"className":"candidate-list-content","value":"div.candidate-list-content","prefix":">"},{"tag":"div","index":0,"className":"card-inner","value":"div.card-inner","prefix":">"},{"tag":"div","index":0,"className":"col-1","value":"div.col-1","prefix":">"},{"tag":"div","index":0,"className":"salary","value":"div.salary","prefix":">"}],"props":["text"]},{"selecors":[{"tag":"div","index":0,"className":"sec-content candidate-card","value":"div.sec-content.candidate-card","prefix":""},{"tag":"ul","index":0,"className":"recommend-card-list ul-less-height","value":"ul.recommend-card-list.ul-less-height","prefix":">"},{"tag":"li","value":"li","index":0,"prefix":">"},{"tag":"div","index":0,"className":"","value":"div","prefix":">"},{"tag":"div","index":0,"className":"geek-info-card less-height","value":"div.geek-info-card.less-height","prefix":">"},{"tag":"div","index":0,"className":"candidate-list-content","value":"div.candidate-list-content","prefix":">"},{"tag":"div","index":0,"className":"card-inner","value":"div.card-inner","prefix":">"},{"tag":"div","index":0,"className":"col-2","value":"div.col-2","prefix":">"},{"tag":"div","index":0,"className":"info-labels","value":"div.info-labels","prefix":">"}],"props":["text"]},{"selecors":[{"tag":"div","index":0,"className":"sec-content candidate-card","value":"div.sec-content.candidate-card","prefix":""},{"tag":"ul","index":0,"className":"recommend-card-list ul-less-height","value":"ul.recommend-card-list.ul-less-height","prefix":">"},{"tag":"li","value":"li","index":0,"prefix":">"},{"tag":"div","index":0,"className":"","value":"div","prefix":">"},{"tag":"div","index":0,"className":"geek-info-card less-height","value":"div.geek-info-card.less-height","prefix":">"},{"tag":"div","index":0,"className":"candidate-list-content","value":"div.candidate-list-content","prefix":">"},{"tag":"div","index":0,"className":"card-inner","value":"div.card-inner","prefix":">"},{"tag":"div","index":0,"className":"col-2","value":"div.col-2","prefix":">"},{"tag":"div","index":0,"className":"expect-box","value":"div.expect-box","prefix":">"}],"props":["text"]},{"selecors":[{"tag":"div","index":0,"className":"sec-content candidate-card","value":"div.sec-content.candidate-card","prefix":""},{"tag":"ul","index":0,"className":"recommend-card-list ul-less-height","value":"ul.recommend-card-list.ul-less-height","prefix":">"},{"tag":"li","value":"li","index":0,"prefix":">"},{"tag":"div","index":0,"className":"","value":"div","prefix":">"},{"tag":"div","index":0,"className":"geek-info-card less-height","value":"div.geek-info-card.less-height","prefix":">"},{"tag":"div","index":0,"className":"candidate-list-content","value":"div.candidate-list-content","prefix":">"},{"tag":"div","index":0,"className":"card-inner","value":"div.card-inner","prefix":">"},{"tag":"div","index":0,"className":"col-2","value":"div.col-2","prefix":">"},{"tag":"div","index":0,"className":"position-box position-advantage","value":"div.position-box.position-advantage","prefix":">"},{"tag":"div","index":0,"className":"advantage-new","value":"div.advantage-new","prefix":">"}],"props":["text"]},{"selecors":[{"tag":"div","index":0,"className":"sec-content candidate-card","value":"div.sec-content.candidate-card","prefix":""},{"tag":"ul","index":0,"className":"recommend-card-list ul-less-height","value":"ul.recommend-card-list.ul-less-height","prefix":">"},{"tag":"li","value":"li","index":0,"prefix":">"},{"tag":"div","index":0,"className":"","value":"div","prefix":">"},{"tag":"div","index":0,"className":"geek-info-card less-height","value":"div.geek-info-card.less-height","prefix":">"},{"tag":"div","index":0,"className":"candidate-list-content","value":"div.candidate-list-content","prefix":">"},{"tag":"div","index":0,"className":"card-inner","value":"div.card-inner","prefix":">"},{"tag":"div","index":0,"className":"col-3","value":"div.col-3","prefix":">"}],"props":["text"]}]},{"objNextLinkElement":"","iMaxNumberOfPage":5,"iMaxNumberOfResult":-1,"iDelayBetweenMS":1000,"bContinueOnError":False})


        iRetFather = Len(arrayData)
        TracePrint('共获取到的候选人数量为：'& iRetFather)

        iRetSon = Len(arrayData[0])
        // 打开一个空白excel 用于存储获取的候选人数据----调试时注释
        // objExcelWorkBook = Excel.OpenExcel('''C:\Users\Administrator\Desktop\boss.xlsx''',True,"Excel","","")


        // objExcelWorkBook = Excel.OpenExcel('''C:\Users\Administrator\Desktop\招呼历史.xlsx''',True,"Excel","","")
        For y = 1 To iRetFather Step 1
            arrayDataHistory =[]
            sRet = ""
            // 2. 姓名正则 排除含外包/外派人员
            findName = ""
            // 3. 工作年限 (大于等于三年)
            workYear = 0
            // 4. 求职期望(含前端/Javascript/HTML/web)
            expectJob = ""
            // 5. 年龄（<=30）
            age = ""
            // 6.求职简介
            summary = ""
            // 7. 在职-暂不考虑
            working = ""
            // 8.要求掌握react
            require_react = ""
            // 9. 学历
            certificate = ""
            // 10. 活跃状态
            active_status = ""


            // 判断当前候选人是否被其他招聘顾问询问过
            #icon("@res:tsa5vmg6-jqmn-84hv-j1pi-4ep57tkmtib5.png")
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

            // For k = 1 To helloSon Step 1
            // Excel.WriteCell(objExcelWorkBook,"Sheet1",Chr(64+k)&x,arrayData[x-1][k-1],False)
            // Next
            // Next

            // For i = 1 To iRetSon Step 1
            // [姓名, 活跃状态, 薪资范围, 年龄 工作年限, 求职期望, 求职简介, 历史工作公司 毕业学校]

            //1. 按钮文本是否为"继续沟通"
            #icon("@res:80me7dm2-gukr-6uo4-chh9-asd0pee45v95.png")
            sRet = UiElement.GetValue({"wnd":[{"cls":"MozillaWindowClass","title":"*","app":"firefox"}],"html":[{"tag":"IFRAME","name":"recommendFrame"},{"tag":"BUTTON","type":"button","parentid":"recommend-list","idx":y-1}]},{"bContinueOnError":False,"iDelayAfter":300,"iDelayBefore":200})
            TracePrint(sRet& "--是否继续沟通")

            // 2. 姓名正则 排除含外包/外派人员
            findName = Regex.FindStr(arrayData[y-1][0],"外包|外派",0)
            TracePrint("源文本："&arrayData[y-1][0] &"匹配后显示："&findName& "--姓名")

            // 3. 工作年限 (大于等于三年)
            workYear = cint(DigitFromStr(Regex.FindStr(arrayData[y-1][3],'([0-9]+年)(?!应届)',0)))
            TracePrint(workYear& "--工作年限")

            // 4. 求职期望(含前端/Javascript/HTML/web)
            expectJob = Regex.FindStr(arrayData[y-1][4],'前端|(?i)javascript|html|(?i)web',0)
            TracePrint("源文本："& arrayData[y-1][4] &"匹配后显示："& expectJob & "--求职期望")

            // 5. 年龄（<=30）
            age = cint(DigitFromStr(Regex.FindStr(arrayData[y-1][3],'[0-9]+岁',0)))
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

            // 9. 学历（大专|本科|硕士）
            certificate = cint(DigitFromStr(Regex.FindStr(arrayData[y-1][3],'大专|本科|硕士',0)))
            TracePrint("源文本："& arrayData[y-1][3] &"匹配后显示："& certificate & "--学历")

            // 10. 活跃状态
            active_status = Regex.FindStr(arrayData[y-1][1],"刚刚|今日|3日|本周|2周",0)
            TracePrint(active_status& "--活跃状态")



            If sRet <> "继续沟通" And findName ="" And workYear >= 5 And expectJob <> "" And age <=30 And summary ="" And working ="" And require_react <> "" And certificate <>"" And active_status <>""
                #icon("@res:9dius7o2-0t3v-kb14-hk2u-7m62k27kvlf2.png")
                // Mouse.Hover({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"IFRAME","name":"recommendFrame"},{"tag":"BUTTON","parentid":"recommend-list","idx":y-1}]},10000,{"bContinueOnError":False,"iDelayAfter":300,"iDelayBefore":200,"bSetForeground":True,"sCursorPosition":"Center","iCursorOffsetX":0,"iCursorOffsetY":0,"sKeyModifiers":[],"sSimulate":"simulate","bMoveSmoothly":False})
                #icon("@res:a96v8b09-hmu1-vpf4-al11-lmevupf6sli4.png")
                Mouse.Action({"wnd":[{"cls":"MozillaWindowClass","title":"*","app":"firefox"}],"html":[{"tag":"IFRAME","name":"recommendFrame"},{"tag":"BUTTON","type":"button","parentid":"recommend-list","idx":y-1}]},"left","click",10000,{"bContinueOnError":False,"iDelayAfter":300,"iDelayBefore":200,"bSetForeground":True,"sCursorPosition":"Center","iCursorOffsetX":0,"iCursorOffsetY":0,"sKeyModifiers":[],"sSimulate":"simulate","bMoveSmoothly":False})
                // TracePrint("已筛选候人选数量为："& y)
            Else
                // Break
                TracePrint(arrayData[y-1][0]&"：不符合条件----")
            End If
            // Excel.WriteCell(objExcelWorkBook,"Sheet1",Chr(64+i)&y,arrayData[y-1][i-1],False)
            // Next
        Next


        // For y = 1 To 1 Step 1
        // #icon("@res:ir1stmts-aj2p-9vae-tir6-gnqogf566pdu.png")
        // Mouse.Hover({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"IFRAME","name":"recommendFrame"},{"tag":"BUTTON","parentid":"recommend-list","idx":2}]},10000,{"bContinueOnError":False,"iDelayAfter":300,"iDelayBefore":200,"bSetForeground":True,"sCursorPosition":"Center","iCursorOffsetX":0,"iCursorOffsetY":0,"sKeyModifiers":[],"sSimulate":"simulate","bMoveSmoothly":False})
        // Next
        #icon("@res:cn6nf0uv-eoel-aame-l4ie-13vnvh6ueslv.png")
        // Mouse.Action({"html":[{"parentid":"recommendContent","tag":"LI"}],"wnd":[{"app":"iexplore","cls":"IEFrame","title":"*"},{"cls":"Internet Explorer_Server"}]},"left","click",10000,{"bContinueOnError":False,"iDelayAfter":300,"iDelayBefore":200,"bSetForeground":True,"sCursorPosition":"Center","iCursorOffsetX":0,"iCursorOffsetY":0,"sKeyModifiers":[],"sSimulate":"simulate","bMoveSmoothly":False})
    Next


Loop






```







