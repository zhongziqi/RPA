```vbscript
Dim dictScrollPostion = ""
Dim ObjSet = ""
Dim objPoint = ""
Dim waittingData = ""
Dim arrayDataHello = ""
Dim arrElement = ""
Dim sRet = 100
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
Dim targetTime = ""
Dim leftTime = ""
Dim number = 0
Dim work_status = ""
Dim cetification = ""
Dim minSalary = ""
Dim xx = 0

hWeb = WebBrowser.Create("chrome","https://rd6.zhaopin.com/",30000,{"bContinueOnError":False,"iDelayAfter":5000,"iDelayBefore":200,"sBrowserPath":"","sStartArgs":""})
#icon("@res:p3b5pnjn-jjbn-esk1-u7nd-g5d0ivtnl2rr.png")
Mouse.Action({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"A","aaname":"           ��ѡ�˲�-Normal     Created with Sketch.                 *"}]},"left","click",10000,{"bContinueOnError":False,"iDelayAfter":3000,"iDelayBefore":200,"bSetForeground":True,"sCursorPosition":"Center","iCursorOffsetX":0,"iCursorOffsetY":0,"sKeyModifiers":[],"sSimulate":"simulate","bMoveSmoothly":False})
#icon("@res:v30r7a67-rgeg-vvlk-98ce-g99d13o7hng8.png")
Mouse.Action({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"A","aaname":"ѡ��ְλ"}]},"left","click",10000,{"bContinueOnError":False,"iDelayAfter":1000,"iDelayBefore":200,"bSetForeground":True,"sCursorPosition":"Center","iCursorOffsetX":0,"iCursorOffsetY":0,"sKeyModifiers":[],"sSimulate":"simulate","bMoveSmoothly":False})
#icon("@res:cnpcq2b1-unkp-m2kk-8h9l-hhbi18rumaod.png")
Mouse.Action({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"DIV","css-selector":"body>div>div>div>div>div>div>div>div>div>div>div","idx":25}]},"left","click",10000,{"bContinueOnError":False,"iDelayAfter":1000,"iDelayBefore":200,"bSetForeground":True,"sCursorPosition":"Center","iCursorOffsetX":0,"iCursorOffsetY":0,"sKeyModifiers":[],"sSimulate":"simulate","bMoveSmoothly":False})
#icon("@res:03qemflm-e4ia-fspl-iqo6-0ifg7p365v9o.png")
Mouse.Action({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"DIV","parentid":"root","css-selector":"body>div>div>div>div>div>div>div>div>div","idx":4}]},"left","click",10000,{"bContinueOnError":False,"iDelayAfter":2000,"iDelayBefore":200,"bSetForeground":True,"sCursorPosition":"Center","iCursorOffsetX":0,"iCursorOffsetY":0,"sKeyModifiers":[],"sSimulate":"simulate","bMoveSmoothly":False})

// ui-dropmenu-label
Mouse.Wheel(3,"down", [],{"iDelayAfter":300,"iDelayBefore":200})
ObjSet=Set.Create()

// Do While True

// 	dictScrollPostion = WebBrowser.GetScroll(hWeb,{"bContinueOnError":False,"iDelayAfter":300,"iDelayBefore":200})
// 	iRet = Dialog.MsgBox(dictScrollPostion,"UiBot","0","1",0)
// Loop
// Do While True

// dictScrollPostion = WebBrowser.GetScroll(hWeb,{"bContinueOnError":False,"iDelayAfter":300,"iDelayBefore":200})
// iRet = Dialog.MsgBox(dictScrollPostion,"UiBot","0","1",0)
// Loop
// objPoint=Mouse.GetPos()
#icon("@res:187b87un-pmkr-j8rh-91e9-k3r3jfq78khc.png")
Do While sRet<>0
	xx =0 
	// ���¹����Լ��ظ�������
	// Do While xx < 4 
	// 	Delay(1000)
	// 	xx = xx + 1
	// 	#icon("@res:default.png")
	// 	Mouse.Wheel(50,"down", [],{"iDelayAfter":1000,"iDelayBefore":2000})
	// Loop
	// ��ȡʣ�������Ƽ���
	sRet = Text.Get({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"SPAN","css-selector":"body>div>div>div>div>div>div>div>div>div>div>span.chat-rights-number"}]},10000,{"bContinueOnError":False,"iDelayAfter":300,"iDelayBefore":1500,"bSetForeground":True})
	
	waittingData = UiElement.DataScrap({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"DIV","id":"root"}]},{"ExtractTable":0,"Columns":[{"selecors":[{"tag":"div","index":0,"className":"app-main","value":"div.app-main","prefix":""},{"tag":"div","index":0,"className":"app-main__content","value":"div.app-main__content","prefix":">"},{"tag":"div","index":0,"className":"talent-recommend","value":"div.talent-recommend","prefix":">"},{"tag":"div","index":0,"className":"recommend-list","value":"div.recommend-list","prefix":">"},{"tag":"div","index":0,"className":"","value":"div","prefix":">"},{"tag":"div","index":0,"className":"","value":"div","prefix":">"},{"tag":"div","index":1,"className":"","value":"div:nth-child(1)","prefix":">"},{"tag":"div","value":"div","index":0,"prefix":">"},{"tag":"div","index":0,"className":"","value":"div","prefix":">"},{"tag":"div","index":0,"className":"recommend-item resume-item","value":"div.recommend-item.resume-item","prefix":">"},{"tag":"div","index":0,"className":"recommend-item__inner resume-item__inner","value":"div.recommend-item__inner.resume-item__inner","prefix":">"},{"tag":"div","index":0,"className":"recommend-item__inner-content","value":"div.recommend-item__inner-content","prefix":">"},{"tag":"div","index":0,"className":"resume-item__content","value":"div.resume-item__content","prefix":">"},{"tag":"div","index":0,"className":"resume-item__basic","value":"div.resume-item__basic","prefix":">"},{"tag":"div","index":0,"className":"resume-item__basic-info","value":"div.resume-item__basic-info","prefix":">"},{"tag":"div","index":0,"className":"talent-basic-info","value":"div.talent-basic-info","prefix":">"},{"tag":"div","index":0,"className":"talent-basic-info__title","value":"div.talent-basic-info__title","prefix":">"},{"tag":"div","index":0,"className":"talent-basic-info__name","value":"div.talent-basic-info__name","prefix":">"},{"tag":"div","index":0,"className":"talent-basic-info__name--inner","value":"div.talent-basic-info__name--inner","prefix":">"}],"props":["text"]},{"selecors":[{"tag":"div","index":0,"className":"app-main","value":"div.app-main","prefix":""},{"tag":"div","index":0,"className":"app-main__content","value":"div.app-main__content","prefix":">"},{"tag":"div","index":0,"className":"talent-recommend","value":"div.talent-recommend","prefix":">"},{"tag":"div","index":0,"className":"recommend-list","value":"div.recommend-list","prefix":">"},{"tag":"div","index":0,"className":"","value":"div","prefix":">"},{"tag":"div","index":0,"className":"","value":"div","prefix":">"},{"tag":"div","index":1,"className":"","value":"div:nth-child(1)","prefix":">"},{"tag":"div","value":"div","index":0,"prefix":">"},{"tag":"div","index":0,"className":"","value":"div","prefix":">"},{"tag":"div","index":0,"className":"recommend-item resume-item","value":"div.recommend-item.resume-item","prefix":">"},{"tag":"div","index":0,"className":"recommend-item__inner resume-item__inner","value":"div.recommend-item__inner.resume-item__inner","prefix":">"},{"tag":"div","index":0,"className":"recommend-item__inner-content","value":"div.recommend-item__inner-content","prefix":">"},{"tag":"div","index":0,"className":"resume-item__content","value":"div.resume-item__content","prefix":">"},{"tag":"div","index":0,"className":"resume-item__basic","value":"div.resume-item__basic","prefix":">"},{"tag":"div","index":0,"className":"resume-item__basic-info","value":"div.resume-item__basic-info","prefix":">"},{"tag":"div","index":0,"className":"talent-basic-info","value":"div.talent-basic-info","prefix":">"},{"tag":"div","index":0,"className":"talent-basic-info__basic","value":"div.talent-basic-info__basic","prefix":">"},{"tag":"span","index":1,"className":"","value":"span:nth-child(1)","prefix":">"}],"props":["text"]},{"selecors":[{"tag":"div","index":0,"className":"app-main","value":"div.app-main","prefix":""},{"tag":"div","index":0,"className":"app-main__content","value":"div.app-main__content","prefix":">"},{"tag":"div","index":0,"className":"talent-recommend","value":"div.talent-recommend","prefix":">"},{"tag":"div","index":0,"className":"recommend-list","value":"div.recommend-list","prefix":">"},{"tag":"div","index":0,"className":"","value":"div","prefix":">"},{"tag":"div","index":0,"className":"","value":"div","prefix":">"},{"tag":"div","index":1,"className":"","value":"div:nth-child(1)","prefix":">"},{"tag":"div","value":"div","index":0,"prefix":">"},{"tag":"div","index":0,"className":"","value":"div","prefix":">"},{"tag":"div","index":0,"className":"recommend-item resume-item","value":"div.recommend-item.resume-item","prefix":">"},{"tag":"div","index":0,"className":"recommend-item__inner resume-item__inner","value":"div.recommend-item__inner.resume-item__inner","prefix":">"},{"tag":"div","index":0,"className":"recommend-item__inner-content","value":"div.recommend-item__inner-content","prefix":">"},{"tag":"div","index":0,"className":"resume-item__content","value":"div.resume-item__content","prefix":">"},{"tag":"div","index":0,"className":"resume-item__basic","value":"div.resume-item__basic","prefix":">"},{"tag":"div","index":0,"className":"resume-item__basic-info","value":"div.resume-item__basic-info","prefix":">"},{"tag":"div","index":0,"className":"talent-basic-info","value":"div.talent-basic-info","prefix":">"},{"tag":"div","index":0,"className":"talent-basic-info__basic","value":"div.talent-basic-info__basic","prefix":">"},{"tag":"span","index":3,"className":"","value":"span:nth-child(3)","prefix":">"}],"props":["text"]},{"selecors":[{"tag":"div","index":0,"className":"app-main","value":"div.app-main","prefix":""},{"tag":"div","index":0,"className":"app-main__content","value":"div.app-main__content","prefix":">"},{"tag":"div","index":0,"className":"talent-recommend","value":"div.talent-recommend","prefix":">"},{"tag":"div","index":0,"className":"recommend-list","value":"div.recommend-list","prefix":">"},{"tag":"div","index":0,"className":"","value":"div","prefix":">"},{"tag":"div","index":0,"className":"","value":"div","prefix":">"},{"tag":"div","index":1,"className":"","value":"div:nth-child(1)","prefix":">"},{"tag":"div","value":"div","index":0,"prefix":">"},{"tag":"div","index":0,"className":"","value":"div","prefix":">"},{"tag":"div","index":0,"className":"recommend-item resume-item","value":"div.recommend-item.resume-item","prefix":">"},{"tag":"div","index":0,"className":"recommend-item__inner resume-item__inner","value":"div.recommend-item__inner.resume-item__inner","prefix":">"},{"tag":"div","index":0,"className":"recommend-item__inner-content","value":"div.recommend-item__inner-content","prefix":">"},{"tag":"div","index":0,"className":"resume-item__content","value":"div.resume-item__content","prefix":">"},{"tag":"div","index":0,"className":"resume-item__basic","value":"div.resume-item__basic","prefix":">"},{"tag":"div","index":0,"className":"resume-item__basic-info","value":"div.resume-item__basic-info","prefix":">"},{"tag":"div","index":0,"className":"talent-basic-info","value":"div.talent-basic-info","prefix":">"},{"tag":"div","index":0,"className":"talent-basic-info__basic","value":"div.talent-basic-info__basic","prefix":">"},{"tag":"span","index":4,"className":"","value":"span:nth-child(4)","prefix":">"}],"props":["text"]},{"selecors":[{"tag":"div","index":0,"className":"app-main","value":"div.app-main","prefix":""},{"tag":"div","index":0,"className":"app-main__content","value":"div.app-main__content","prefix":">"},{"tag":"div","index":0,"className":"talent-recommend","value":"div.talent-recommend","prefix":">"},{"tag":"div","index":0,"className":"recommend-list","value":"div.recommend-list","prefix":">"},{"tag":"div","index":0,"className":"","value":"div","prefix":">"},{"tag":"div","index":0,"className":"","value":"div","prefix":">"},{"tag":"div","index":1,"className":"","value":"div:nth-child(1)","prefix":">"},{"tag":"div","value":"div","index":0,"prefix":">"},{"tag":"div","index":0,"className":"","value":"div","prefix":">"},{"tag":"div","index":0,"className":"recommend-item resume-item","value":"div.recommend-item.resume-item","prefix":">"},{"tag":"div","index":0,"className":"recommend-item__inner resume-item__inner","value":"div.recommend-item__inner.resume-item__inner","prefix":">"},{"tag":"div","index":0,"className":"recommend-item__inner-content","value":"div.recommend-item__inner-content","prefix":">"},{"tag":"div","index":0,"className":"resume-item__content","value":"div.resume-item__content","prefix":">"},{"tag":"div","index":0,"className":"resume-item__basic","value":"div.resume-item__basic","prefix":">"},{"tag":"div","index":0,"className":"resume-item__basic-info","value":"div.resume-item__basic-info","prefix":">"},{"tag":"div","index":0,"className":"talent-basic-info","value":"div.talent-basic-info","prefix":">"},{"tag":"div","index":0,"className":"talent-basic-info__extra","value":"div.talent-basic-info__extra","prefix":">"},{"tag":"span","index":2,"className":"","value":"span:nth-child(2)","prefix":">"}],"props":["text"]},{"selecors":[{"tag":"div","index":0,"className":"app-main","value":"div.app-main","prefix":""},{"tag":"div","index":0,"className":"app-main__content","value":"div.app-main__content","prefix":">"},{"tag":"div","index":0,"className":"talent-recommend","value":"div.talent-recommend","prefix":">"},{"tag":"div","index":0,"className":"recommend-list","value":"div.recommend-list","prefix":">"},{"tag":"div","index":0,"className":"","value":"div","prefix":">"},{"tag":"div","index":0,"className":"","value":"div","prefix":">"},{"tag":"div","index":1,"className":"","value":"div:nth-child(1)","prefix":">"},{"tag":"div","value":"div","index":0,"prefix":">"},{"tag":"div","index":0,"className":"","value":"div","prefix":">"},{"tag":"div","index":0,"className":"recommend-item resume-item","value":"div.recommend-item.resume-item","prefix":">"},{"tag":"div","index":0,"className":"recommend-item__inner resume-item__inner","value":"div.recommend-item__inner.resume-item__inner","prefix":">"},{"tag":"div","index":0,"className":"recommend-item__inner-content","value":"div.recommend-item__inner-content","prefix":">"},{"tag":"div","index":0,"className":"resume-item__content","value":"div.resume-item__content","prefix":">"},{"tag":"div","index":0,"className":"resume-item__basic","value":"div.resume-item__basic","prefix":">"},{"tag":"div","index":0,"className":"resume-item__basic-info","value":"div.resume-item__basic-info","prefix":">"},{"tag":"div","index":0,"className":"talent-basic-info","value":"div.talent-basic-info","prefix":">"},{"tag":"div","index":0,"className":"talent-basic-info__extra","value":"div.talent-basic-info__extra","prefix":">"},{"tag":"span","index":3,"className":"is-shrinkless","value":"span:nth-child(3)","prefix":">"}],"props":["text"]},{"selecors":[{"tag":"div","index":0,"className":"app-main","value":"div.app-main","prefix":""},{"tag":"div","index":0,"className":"app-main__content","value":"div.app-main__content","prefix":">"},{"tag":"div","index":0,"className":"talent-recommend","value":"div.talent-recommend","prefix":">"},{"tag":"div","index":0,"className":"recommend-list","value":"div.recommend-list","prefix":">"},{"tag":"div","index":0,"className":"","value":"div","prefix":">"},{"tag":"div","index":0,"className":"","value":"div","prefix":">"},{"tag":"div","index":1,"className":"","value":"div:nth-child(1)","prefix":">"},{"tag":"div","value":"div","index":0,"prefix":">"},{"tag":"div","index":0,"className":"","value":"div","prefix":">"},{"tag":"div","index":0,"className":"recommend-item resume-item","value":"div.recommend-item.resume-item","prefix":">"},{"tag":"div","index":0,"className":"recommend-item__inner resume-item__inner","value":"div.recommend-item__inner.resume-item__inner","prefix":">"},{"tag":"div","index":0,"className":"recommend-item__inner-content","value":"div.recommend-item__inner-content","prefix":">"},{"tag":"div","index":0,"className":"resume-item__content","value":"div.resume-item__content","prefix":">"},{"tag":"div","index":0,"className":"resume-item__basic","value":"div.resume-item__basic","prefix":">"},{"tag":"div","index":0,"className":"resume-item__basic-info","value":"div.resume-item__basic-info","prefix":">"},{"tag":"div","index":0,"className":"talent-basic-info","value":"div.talent-basic-info","prefix":">"},{"tag":"div","index":0,"className":"talent-basic-info__tags","value":"div.talent-basic-info__tags","prefix":">"}],"props":["text"]}]},{"objNextLinkElement":"","iMaxNumberOfPage":0,"iMaxNumberOfResult":-1,"iDelayBetweenMS":1000,"bContinueOnError":False})
	iRetFather = Len(waittingData)
	TracePrint(iRetFather)
	TracePrint('ʣ���Ƽ�����Ϊ��'& sRet)
	
	iRetSon = Len(waittingData[0])
	TracePrint(waittingData[0])
	// exit()
	
	// ��ȡҳ������
	For y = 1 To iRetFather Step 1
		
		// �����������¹���, �Ի�ȡʣ�������Ƽ���
		
		// 0. ��������  �ų������/������Ա
		#icon("@res:mksonr7f-mv64-lbop-8km1-sschtetjj2hn.png")
		// UiElement.SetAttribute({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"DIV","parentid":"root","css-selector":"body>div>div>div>div>div>div>div>div>div[role='listitem']","idx":y-1}]},"style","170px",{"bContinueOnError":False,"iDelayAfter":300,"iDelayBefore":1000})
		findName = ""
		// 1. ���䣨<=32��
		age =  ""
		// 2. ѧ��
		cetification = ""
		// 3. ��ְ-�����ҹ���    ��ְ-�����ҹ���
		work_status = ""
		// 4. ��ְ����(��רԱ/����)
		expectJob =  ""
		// 5.������Сֵ 1w
		minSalary = ""
		// 6.��ְ��ǩ
		summary = ""
		
		bRet =""
		
		TracePrint(waittingData[y-1][1])
		
		// 1. ���䣨<=32��
		age =  Cint(Regex.FindStr(waittingData[y-1][1],'[1-9]\\d{1}(?!\\d)',0))
		TracePrint("Դ�ı���"& waittingData[y-1][1] &"ƥ�����ʾ��"& age & "--����")
		
		// 2. ѧ������ר|���ƣ�
		cetification =  Regex.FindStr(waittingData[y-1][2],'��ר|����|˶ʿ',0)
		TracePrint("Դ�ı���"& waittingData[y-1][2] &"ƥ�����ʾ��"& cetification & "--ѧ��")
		
		// 3. ��ְ״̬����ְ-�����ҹ���|��ְ-�����ҹ�����
		work_status =  Regex.FindStr(waittingData[y-1][3],'��ְ-�����ҹ���|��ְ-�����ҹ���',0)
		TracePrint("Դ�ı���"& waittingData[y-1][3] &"ƥ�����ʾ��"& work_status & "--��ְ״̬")
		
		// // 3. �������� (���ڵ�������)
		// workYear = cint(DigitFromStr(Regex.FindStr(waittingData[y-1][3],'[0-9]+��',0)))
		// TracePrint(workYear& "--��������")
		
		// 4. ��ְ����(��רԱ/����)
		expectJob =  Regex.FindStr(waittingData[y-1][4],'����',0)
		TracePrint("Դ�ı���"& waittingData[y-1][4] &"ƥ�����ʾ��"& expectJob & "--��ְ����")
		
		// 5. ������Сֵ(10,000)
		arrRet =  Split(waittingData[y-1][5],"-")
		num =  Regex.FindStr(arrRet[1],'[1-9]\\d*\\.?\\d*',0)
		thousand =  Regex.FindStr(arrRet[1],'ǧ',0)
		// tenThousand =  Regex.FindStr(arrRet[0],'��',0)
		If thousand="ǧ"
			minSalary = (num)*1000
		Else
			If num = ""
				minSalary = 0
			Else  
				minSalary = (num)*10000
			End If
		End If
		TracePrint("Դ�ı���"& waittingData[y-1][5] &"ƥ�����ʾ��"& cint(minSalary) & "--������Сֵ")
		#icon("@res:vo60278g-n0nt-jhsh-vspf-ak92gpvq7n7e.png")
		
		// 6.��ְ��ǩ
		summary = Regex.FindStr(waittingData[y-1][6],"(?i)it|������|����|����|������Դ",0)
		TracePrint(summary& "--��ְ���")
		
		// // 7. ȥ�� "��ְ-�ݲ�����"
		// working = Regex.FindStr(waittingData[y-1][3],"��ְ-�ݲ�����",0)
		// TracePrint(working& "--��ְ���")
		
		// // 8.Ҫ������react
		// require_react = Regex.FindStr(waittingData[y-1][5],'(?i)react',0)
		// TracePrint(findName&"--"&require_react& "--require_react")
		
		
		If age <=35 And cetification <>"" And work_status <>"" And expectJob <> "" And minSalary <=25000
			#icon("@res:re2jjcba-o7n4-33s6-kl51-nlb3tk19iaok.png")
			// �����Է�ֹ��������ڵ��Եײ��˵���
			// Mouse.Wheel(3,"down", [],{"iDelayAfter":300,"iDelayBefore":200})
			#icon("@res:mvdjj6hf-ondr-t21s-tgst-il3tjb7rjbhl.png")
			// Text.Click({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"SPAN","aaname":"������","idx":y-1}]},"������","instr",1,"left","click",10000,{"bContinueOnError":False,"iDelayAfter":300,"iDelayBefore":200,"bSetForeground":True,"sCursorPosition":"Center","iCursorOffsetX":0,"iCursorOffsetY":0,"sKeyModifiers":[],"sSimulate":"simulate"})
			#icon("@res:hhn1qjss-oap0-g60g-oge8-4i351iaffhpg.png")
			#icon("@res:ja55k0t0-g208-0j5n-9rf7-4rhvejvsrj83.png")
			// Mouse.Hover({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"I","isleaf":"1","css-selector":"body>div>div>div>div>div>div>div>div>div>div>div>div>div>div>div>div>div>button>div>button>div>i.sk-chat","idx":y-1}]},10000,{"bContinueOnError":False,"iDelayAfter":300,"iDelayBefore":200,"bSetForeground":True,"sCursorPosition":"Center","iCursorOffsetX":0,"iCursorOffsetY":0,"sKeyModifiers":[],"sSimulate":"simulate","bMoveSmoothly":False})
			#icon("@res:a96v8b09-hmu1-vpf4-al11-lmevupf6sli4.png")
			// Mouse.Action({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"I","isleaf":"1","css-selector":"body>div>div>div>div>div>div>div>div>div>div>div>div>div>div>div>div>div>button>div>button>div>i.sk-chat","idx":y-1}]},"left","click",10000,{"bContinueOnError":False,"iDelayAfter":300,"iDelayBefore":200,"bSetForeground":True,"sCursorPosition":"Center","iCursorOffsetX":0,"iCursorOffsetY":0,"sKeyModifiers":[],"sSimulate":"simulate","bMoveSmoothly":False})
			
			
			//  Mouse.Hover({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"},{"cls":"Chrome_RenderWidgetHostHWND","title":"Chrome Legacy Window"}],"html":[{"tag":"DIV","parentid":"root","aaname":"         ������       ","idx":y-1}]},10000,{"bContinueOnError":False,"iDelayAfter":300,"iDelayBefore":1500,"bSetForeground":True,"sCursorPosition":"Center","iCursorOffsetX":0,"iCursorOffsetY":0,"sKeyModifiers":[],"sSimulate":"simulate","bMoveSmoothly":False})
			
		Else
			TracePrint(waittingData[y-1][0]&"������������----")
		End If
	Next
Loop



```

