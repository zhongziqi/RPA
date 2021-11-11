Dim dictRet = ""

Dim arrayData = ""

Dim hWeb = ""

Dim objExcelWorkBook = ""

Dim length = ""





hWeb = WebBrowser.Create("chrome","https://www.jianyu360.cn/jylab/supsearch/index.html?publishtime=thisyear",30000,{"bContinueOnError":False,"iDelayAfter":3000,"iDelayBefore":200,"sBrowserPath":"","sStartArgs":""})

dictRet = Dialog.UDFDialog("剑鱼excel导出工具",@res"1636539745009.json",{},{"iTimeout":0,"strTimoutClick":"ok","bInterruptTimeout":False})

// TracePrint(dictRet["文本框"])

\#icon("@res:ed1ggjb1-helk-43p2-1s7q-mbfvieduhu64.png")

Keyboard.InputText({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"INPUT","id":"searchinput"}]},dictRet["请输入你要搜索的关键词"],True,20,10000,{"bContinueOnError":False,"iDelayAfter":300,"iDelayBefore":500,"bSetForeground":True,"sSimulate":"message","bValidate":False,"bClickBeforeInput":False})

\#icon("@res:qhj2i6ca-uo05-1htr-rhd4-23gt45lb4op8.png")

Mouse.Action({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"INPUT","type":"button"}]},"left","click",10000,{"bContinueOnError":False,"iDelayAfter":300,"iDelayBefore":200,"bSetForeground":True,"sCursorPosition":"Center","iCursorOffsetX":0,"iCursorOffsetY":0,"sKeyModifiers":[],"sSimulate":"simulate","bMoveSmoothly":False})

\#icon("@res:th9kq4c3-gs1g-hnjd-pn8k-j05ca8qa32ks.png")

Mouse.Action({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"FONT","parentid":"searchInner","css-selector":"body>section>div>div>div>div>font","idx":72}]},"left","click",10000,{"bContinueOnError":False,"iDelayAfter":2000,"iDelayBefore":200,"bSetForeground":True,"sCursorPosition":"Center","iCursorOffsetX":0,"iCursorOffsetY":0,"sKeyModifiers":[],"sSimulate":"simulate","bMoveSmoothly":False})

\#icon("@res:977o2jik-t9v8-uhe4-bfp0-scobvk80rfkv.png")

Mouse.Action({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"BUTTON","id":"right-table"}]},"left","click",10000,{"bContinueOnError":False,"iDelayAfter":2000,"iDelayBefore":200,"bSetForeground":True,"sCursorPosition":"Center","iCursorOffsetX":0,"iCursorOffsetY":0,"sKeyModifiers":[],"sSimulate":"simulate","bMoveSmoothly":False})

arrayData = UiElement.DataScrap({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"TABLE","parentid":"searchInner","idx":1}]},{"ExtractTable":1,"Columns":[]},{"objNextLinkElement":"","iMaxNumberOfPage":5,"iMaxNumberOfResult":-1,"iDelayBetweenMS":1000,"bContinueOnError":False})



length = Len(arrayData)

length_ = Len(arrayData[0])

TracePrint(length)

Dialog.Notify("已获取到"& length &"条数据"  , "请知悉", "0")

objExcelWorkBook = Excel.OpenExcel('''C:\Users\Administrator\Desktop\剑鱼.xlsx''',True,"Excel","","")

For i = 1 To length Step 1 

  Excel.WriteCell(objExcelWorkBook,"Sheet1",Chr(64+1)&1,"序号",False)

  Excel.WriteCell(objExcelWorkBook,"Sheet1",Chr(64+2)&1,"项目名称",False)

  Excel.WriteCell(objExcelWorkBook,"Sheet1",Chr(64+3)&1,"公告类型",False)

  Excel.WriteCell(objExcelWorkBook,"Sheet1",Chr(64+4)&1,"预算(万元)",False)

  Excel.WriteCell(objExcelWorkBook,"Sheet1",Chr(64+5)&1,"招标单位",False)

  Excel.WriteCell(objExcelWorkBook,"Sheet1",Chr(64+6)&1,"开标日期",False)

  Excel.WriteCell(objExcelWorkBook,"Sheet1",Chr(64+7)&1,"中标单位",False)

  Excel.WriteCell(objExcelWorkBook,"Sheet1",Chr(64+8)&1,"中标金额(元)",False)

  Excel.WriteCell(objExcelWorkBook,"Sheet1",Chr(64+9)&1,"发布日期",False)

  

  For y = 1 To length_ Step 1 

​    

​    Excel.WriteCell(objExcelWorkBook,"Sheet1",Chr(64+y)&(i+1),arrayData[i-1][y-1],False)

  Next

Next

exit()









// 打开一个空白excel 用于存储获取的候选人数据----调试时注释