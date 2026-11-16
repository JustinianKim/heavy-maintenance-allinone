'================================================================================
' 제주항공 SAP GUI Script: 객실 청소 및 소독 오더 자동 등록/릴리즈
'================================================================================

Option Explicit

Dim SapGuiAuto, application, connection, session
Dim args, i, item

Set args = WScript.Arguments

'--------------------------------------------------------------------------------
' 0. 단독 SAP 연결 체크 모드 (CHECK_SAP)
'--------------------------------------------------------------------------------
If args.Count >= 1 Then
    If UCase(args(0)) = "CHECK_SAP" Then
        If InitSapSession(False) Then
            WScript.Echo "STATUS:OK|SAP GUI 세션이 정상 연결되었습니다."
            WScript.Quit 0
        Else
            WScript.Echo "STATUS:ERROR|SAP GUI가 로그인되어 있지 않거나 스크립팅이 비활성화되어 있습니다."
            WScript.Quit 1
        End If
    End If
End If

'--------------------------------------------------------------------------------
' 1. SAP GUI 세션 초기화
'--------------------------------------------------------------------------------
If Not InitSapSession(True) Then
    WScript.Echo "STATUS:ERROR|SAP GUI 세션 연결 실패"
    WScript.Quit 1
End If

'--------------------------------------------------------------------------------
' 2. 매개변수 확인 및 실행
'--------------------------------------------------------------------------------
If args.Count >= 7 Then
    ' 파라미터 모드 (Python GUI 등 외부 호출):
    ' args(0): acNo (항공기 기번, 예: HL8031)
    ' args(1): workCenter (작업장, 예: GMP-ACM)
    ' args(2): reqDate (완료요청일, 예: 2027.09.02)
    ' args(3): userNo (발행인 사번, 예: 1511024)
    ' args(4): qmtxt (오더 내역 텍스트)
    ' args(5): ilart (액티비티 타입, 예: CLN / CHK)
    ' args(6): taskList (태스크리스트 번호, 예: 10630, 13870 또는 "-")
    
    Dim pTaskList
    pTaskList = args(6)
    If pTaskList = "-" Or pTaskList = "NONE" Then pTaskList = ""

    Dim result
    result = ProcessCleaningOrder(args(0), args(1), args(2), args(3), args(4), args(5), pTaskList)
    If result Then
        WScript.Echo "STATUS:SUCCESS|" & args(0) & "|" & args(5) & "|오더가 성공적으로 생성 및 릴리즈되었습니다."
        WScript.Quit 0
    Else
        WScript.Echo "STATUS:FAILED|" & args(0) & "|" & args(5) & "|오더 생성 처리 중 오류가 발생했습니다."
        WScript.Quit 1
    End If

Else
    ' 인자가 없을 때 기본 테스트 실행
    Dim tasks
    tasks = Array( _
        Array("HL8031", "GMP-ACM", "2027.09.02", "1511024", "Aircraft Disinfection and Disinsection", "CLN", "10630"), _
        Array("HL8031", "GMP-ACM", "2027.09.02", "1511024", "CABIN CLEANING CHK DUE", "CHK", "") _
    )

    For i = 0 To UBound(tasks)
        item = tasks(i)
        Call ProcessCleaningOrder(item(0), item(1), item(2), item(3), item(4), item(5), item(6))
    Next

    MsgBox "총 " & (UBound(tasks) + 1) & "건의 작업 오더 생성이 완료되었습니다.", vbInformation, "SAP 자동화 완료"
End If


'================================================================================
' SAP GUI 연결 헬퍼 함수
'================================================================================
Function InitSapSession(showMsgOnError)
    On Error Resume Next
    InitSapSession = False

    If Not IsObject(application) Then
        Set SapGuiAuto = GetObject("SAPGUI")
        If Err.Number <> 0 Or SapGuiAuto Is Nothing Then
            If showMsgOnError Then WScript.Echo "ERROR: SAP GUI가 실행 중이지 않습니다."
            Exit Function
        End If
        Set application = SapGuiAuto.GetScriptingEngine
        If Err.Number <> 0 Or application Is Nothing Then
            If showMsgOnError Then WScript.Echo "ERROR: SAP Scripting Engine을 가져올 수 없습니다."
            Exit Function
        End If
    End If

    If Not IsObject(connection) Then
        If application.Children.Count = 0 Then
            If showMsgOnError Then WScript.Echo "ERROR: 활성화된 SAP 연결(Connection)이 없습니다."
            Exit Function
        End If
        Set connection = application.Children(0)
    End If

    If Not IsObject(session) Then
        If connection.Children.Count = 0 Then
            If showMsgOnError Then WScript.Echo "ERROR: 활성화된 SAP 세션(Session)이 없습니다."
            Exit Function
        End If
        Set session = connection.Children(0)
    End If

    If IsObject(WScript) Then
        WScript.ConnectObject session,     "on"
        WScript.ConnectObject application, "on"
    End If

    session.findById("wnd[0]").maximize
    InitSapSession = True
    On Error GoTo 0
End Function


'================================================================================
' 코어 루틴: 개별 클리닝/소독 오더 생성 함수
'================================================================================
Function ProcessCleaningOrder(acNo, workCenter, reqDate, userNo, qmtxt, ilart, taskList)
    If UCase(Trim(qmtxt)) = "AIRCRAFT DISINFECTION AND DISINSECTION" Then
        ProcessCleaningOrder = ProcessDisinfectionOrder(acNo, workCenter, reqDate, userNo, taskList)
        Exit Function
    End If

    On Error Resume Next
    ProcessCleaningOrder = False
    Dim currentStep

    ' (1) 신규 통지 생성 트랜잭션(/nIW21) 진입
    currentStep = "1. 신규 통지 트랜잭션(/nIW21) 진입"
    session.findById("wnd[0]").maximize
    
    If session.findById("wnd[0]/usr/ctxtRIWO00-QMART", False) Is Nothing Then
        Dim autoBtn
        Set autoBtn = session.findById("wnd[0]/usr/btn%#AUTOTEXT008", False)
        If Not autoBtn Is Nothing Then
            autoBtn.press
        Else
            session.findById("wnd[0]/tbar[0]/okcd").text = "/nIW21"
            session.findById("wnd[0]").sendVKey 0
        End If
    End If

    ' 통지 유형 AR 입력
    currentStep = "1-1. 통지유형(AR) 입력"
    session.findById("wnd[0]/usr/ctxtRIWO00-QMART").text = "AR"
    session.findById("wnd[0]/usr/ctxtRIWO00-QMART").caretPosition = 2
    session.findById("wnd[0]").sendVKey 0

    If CheckSapError(currentStep) Then Exit Function

    ' (2) 탭1 (TAB01) 데이터 입력
    currentStep = "2. TAB01 기본 정보 입력"
    session.findById("wnd[0]/usr/subSCREEN_1:SAPLIQS0:1050/subNOTIF_TYPE:SAPLIQS0:1051/txtVIQMEL-QMTXT").text = "CABIN CLEANING CHK DUE"
    session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_1:SAPLIQS0:7322/subOBJEKT:SAPLIWO1:0100/ctxtRIWO1-TPLNR").text = acNo
    session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_3:SAPLIQS0:7326/ctxtVIQMEL-INGRP").text = "100"
    session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_3:SAPLIQS0:7326/ctxtRIWO00-GEWRK").text = workCenter
    session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_4:SAPLIQS0:7330/ctxtVIQMEL-LTRMN").text = reqDate
    session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_5:SAPLIQS0:7901/subUSER0001:SAPLXQQM:0101/ctxtVIQMEL-ZRUSRNO").text = userNo
    session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_5:SAPLIQS0:7901/subUSER0001:SAPLXQQM:0101/ctxtVIQMEL-ZRUSRNO").setFocus
    session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_5:SAPLIQS0:7901/subUSER0001:SAPLXQQM:0101/ctxtVIQMEL-ZRUSRNO").caretPosition = Len(userNo)
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_2:SAPLIQS0:7715/cntlTEXT/shellcont/shell").setSelectionIndexes 0,0

    If CheckSapError(currentStep) Then Exit Function

    ' (3) 탭2 (TAB02) 데이터 입력
    currentStep = "3. TAB02 Due Date 입력"
    session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02").select
    session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_1:SAPLIQS0:7901/subUSER0001:SAPLXQQM:0100/ctxtVIQMEL-ZDDATE_AC").text = reqDate
    session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_1:SAPLIQS0:7901/subUSER0001:SAPLXQQM:0100/ctxtVIQMEL-ZDDATE_AC").setFocus
    session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_1:SAPLIQS0:7901/subUSER0001:SAPLXQQM:0100/ctxtVIQMEL-ZDDATE_AC").caretPosition = 10

    If CheckSapError(currentStep) Then Exit Function

    ' (4) 오더 생성 버튼 클릭 및 팝업창(오더 유형) 확인 처리
    currentStep = "4. 오더 생성 버튼(btnXICON_ORDER) 클릭"
    session.findById("wnd[0]/usr/subSCREEN_1:SAPLIQS0:1050/btnXICON_ORDER").press
    session.findById("wnd[1]/tbar[0]/btn[0]").press

    If CheckSapError(currentStep) Then Exit Function

    ' (5) 액티비티 타입(ILART) 입력
    currentStep = "5. 액티비티 타입(ILART: CHK) 입력"
    session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/tabsTS_1100/tabpIHKZ/ssubSUB_AUFTRAG:SAPLCOIH:1120/subHEADER:SAPLCOIH:0154/ctxtCAUFVD-ILART").text = "CHK"
    session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/tabsTS_1100/tabpIHKZ/ssubSUB_AUFTRAG:SAPLCOIH:1120/subHEADER:SAPLCOIH:0154/ctxtCAUFVD-ILART").setFocus
    session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/tabsTS_1100/tabpIHKZ/ssubSUB_AUFTRAG:SAPLCOIH:1120/subHEADER:SAPLCOIH:0154/ctxtCAUFVD-ILART").caretPosition = 3
    session.findById("wnd[0]").sendVKey 0

    If CheckSapError(currentStep) Then Exit Function

    ' (6) 릴리즈(Release) 및 저장(Save)
    currentStep = "6. 오더 릴리즈(Release) 및 저장(Save)"
    session.findById("wnd[0]/tbar[1]/btn[25]").press ' Release
    session.findById("wnd[0]/tbar[0]/btn[11]").press ' Save
    session.findById("wnd[0]").resizeWorkingPane 199,35,false

    If CheckSapError(currentStep) Then Exit Function

    ' (7) 이전 화면 정리
    currentStep = "7. 화면 정리"
    Dim cleanBtn
    Set cleanBtn = session.findById("wnd[0]/usr/btn%#AUTOTEXT010", False)
    If Not cleanBtn Is Nothing Then
        cleanBtn.press
    End If
    
    Dim qmnumElem
    Set qmnumElem = session.findById("wnd[0]/usr/ctxtRIWO00-QMNUM", False)
    If Not qmnumElem Is Nothing Then
        qmnumElem.text = ""
        session.findById("wnd[0]/tbar[0]/btn[3]").press ' F3 뒤로가기
    End If

    ProcessCleaningOrder = True
    On Error GoTo 0
End Function
End Function

'================================================================================
' 표준 스크립트(Script1.vbs) 기반 소독(Disinfection) 오더 생성 함수
'================================================================================
Function ProcessDisinfectionOrder(acNo, workCenter, reqDate, userNo, taskList)
    On Error Resume Next
    ProcessDisinfectionOrder = False
    Dim currentStep

    currentStep = "1. 통지유형(AR) 입력"
    session.findById("wnd[0]").maximize
    If session.findById("wnd[0]/usr/ctxtRIWO00-QMART", False) Is Nothing Then
        Dim autoBtn
        Set autoBtn = session.findById("wnd[0]/usr/btn%#AUTOTEXT008", False)
        If Not autoBtn Is Nothing Then
            autoBtn.press
        Else
            session.findById("wnd[0]/tbar[0]/okcd").text = "/nIW21"
            session.findById("wnd[0]").sendVKey 0
        End If
    End If
    session.findById("wnd[0]/usr/ctxtRIWO00-QMART").text = "AR"
    session.findById("wnd[0]/usr/ctxtRIWO00-QMART").caretPosition = 2
    session.findById("wnd[0]").sendVKey 0
    If CheckSapError(currentStep) Then Exit Function

    currentStep = "2. TAB01 소독 기본 정보 입력 (표준 Script1.vbs)"
    session.findById("wnd[0]/usr/subSCREEN_1:SAPLIQS0:1050/subNOTIF_TYPE:SAPLIQS0:1051/txtVIQMEL-QMTXT").text = "Aircraft Disinfection and Disinsection"
    session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_1:SAPLIQS0:7322/subOBJEKT:SAPLIWO1:0100/ctxtRIWO1-TPLNR").text = acNo
    session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_3:SAPLIQS0:7326/ctxtVIQMEL-INGRP").text = "100"
    session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_3:SAPLIQS0:7326/ctxtRIWO00-GEWRK").text = workCenter
    session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_4:SAPLIQS0:7330/ctxtVIQMEL-LTRMN").text = reqDate
    session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_5:SAPLIQS0:7901/subUSER0001:SAPLXQQM:0101/ctxtVIQMEL-ZRUSRNO").text = userNo
    session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_5:SAPLIQS0:7901/subUSER0001:SAPLXQQM:0101/ctxtVIQMEL-ZRUSRNO").setFocus
    session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_5:SAPLIQS0:7901/subUSER0001:SAPLXQQM:0101/ctxtVIQMEL-ZRUSRNO").caretPosition = Len(userNo)
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_2:SAPLIQS0:7715/cntlTEXT/shellcont/shell").setSelectionIndexes 0,0
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_2:SAPLIQS0:7715/cntlTEXT/shellcont/shell").setSelectionIndexes 0,0
    If CheckSapError(currentStep) Then Exit Function

    currentStep = "3. TAB02 Due Date 입력"
    session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02").select
    session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_1:SAPLIQS0:7901/subUSER0001:SAPLXQQM:0100/ctxtVIQMEL-ZDDATE_AC").text = reqDate
    session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_1:SAPLIQS0:7901/subUSER0001:SAPLXQQM:0100/ctxtVIQMEL-ZDDATE_AC").setFocus
    session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_1:SAPLIQS0:7901/subUSER0001:SAPLXQQM:0100/ctxtVIQMEL-ZDDATE_AC").caretPosition = 10
    If CheckSapError(currentStep) Then Exit Function

    currentStep = "4. 오더 생성 및 CHK 액티비티 입력"
    session.findById("wnd[0]/usr/subSCREEN_1:SAPLIQS0:1050/btnXICON_ORDER").press
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/tabsTS_1100/tabpIHKZ/ssubSUB_AUFTRAG:SAPLCOIH:1120/subHEADER:SAPLCOIH:0154/ctxtCAUFVD-ILART").text = "CHK"
    session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/tabsTS_1100/tabpIHKZ/ssubSUB_AUFTRAG:SAPLCOIH:1120/subHEADER:SAPLCOIH:0154/ctxtCAUFVD-ILART").setFocus
    session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/tabsTS_1100/tabpIHKZ/ssubSUB_AUFTRAG:SAPLCOIH:1120/subHEADER:SAPLCOIH:0154/ctxtCAUFVD-ILART").caretPosition = 3
    session.findById("wnd[0]").sendVKey 0
    If CheckSapError(currentStep) Then Exit Function

    currentStep = "5. Task List(" & taskList & ") 및 작업장 적용"
    session.findById("wnd[0]/mbar/menu[2]/menu[0]/menu[4]").select
    session.findById("wnd[0]/usr/ctxtPN_PLNNR-LOW").text = taskList
    session.findById("wnd[0]/usr/ctxtPN_PLNNR-LOW").setFocus
    session.findById("wnd[0]/usr/ctxtPN_PLNNR-LOW").caretPosition = Len(taskList)
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/tabsTS_1100/tabpIHKZ/ssubSUB_AUFTRAG:SAPLCOIH:1120/subAVO:SAPLCOI0:0310/ctxtAFVGD-ARBPL").text = workCenter
    session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/tabsTS_1100/tabpIHKZ/ssubSUB_AUFTRAG:SAPLCOIH:1120/subAVO:SAPLCOI0:0310/ctxtAFVGD-ARBPL").setFocus
    session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/tabsTS_1100/tabpIHKZ/ssubSUB_AUFTRAG:SAPLCOIH:1120/subAVO:SAPLCOI0:0310/ctxtAFVGD-ARBPL").caretPosition = Len(workCenter)
    session.findById("wnd[0]/tbar[1]/btn[25]").press
    session.findById("wnd[0]/tbar[0]/btn[11]").press
    session.findById("wnd[0]").resizeWorkingPane 199,35,false
    If CheckSapError(currentStep) Then Exit Function

    Dim cleanBtn
    Set cleanBtn = session.findById("wnd[0]/usr/btn%#AUTOTEXT010", False)
    If Not cleanBtn Is Nothing Then cleanBtn.press
    Dim qmnumElem
    Set qmnumElem = session.findById("wnd[0]/usr/ctxtRIWO00-QMNUM", False)
    If Not qmnumElem Is Nothing Then
        qmnumElem.text = ""
        session.findById("wnd[0]/tbar[0]/btn[3]").press
    End If

    ProcessDisinfectionOrder = True
    On Error GoTo 0
End Function

Function CheckSapError(stepName)
    If Err.Number <> 0 Then
        WScript.Echo "ERR_DETAIL: [" & stepName & "] " & Err.Description & " (Code: " & Err.Number & ")"
        CheckSapError = True
        Err.Clear
    Else
        CheckSapError = False
    End If
End Function
