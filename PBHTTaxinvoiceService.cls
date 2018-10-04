VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PBHTTaxinvoiceService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'=================================================================================
' Class Module for base module for Popbill Hometax API SDK. It include base functionality for
' RESTful web service request and parse json result. It uses Linkhub class module
' to accomplish authentication APIs.
'
' This module uses advapi32.dll and crypt32.dllm. most of windows system has it
' on system32 by default.
'
' http://www.popbill.com
' Author : Jeong Yohan (code@linkhub.co.kr)
' Written : 2016-06-27
' Updated : 2017-08-30
' Contributor : Kim EunHye (code@linkhub.co.kr)
' Updated : 2018-10-04
'
' Thanks for your interest.
'=================================================================================
Private Const ServiceID_REAL = "POPBILL"
Private Const ServiceID_TEST = "POPBILL_TEST"
Private Const ServiceURL_REAL = "https://popbill.linkhub.co.kr"
Private Const ServiceURL_TEST = "https://popbill-test.linkhub.co.kr"
Private Const APIVersion = "1.0"

Private m_IsTest As Boolean
Private m_Linkhub As Linkhub
Private m_scope As New Collection
Private m_LastErrCode As Long
Private m_LastErrMessage As String
Private m_token_Dic

'전자(세금)계산서 형태
Public Enum KeyType
    SELL '매출
    BUY '매입
    TRUSTEE '수탁
End Enum

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Sub Class_Initialize()
    Set m_token_Dic = CreateObject("Scripting.Dictionary")
End Sub

Private Sub Class_Terminate()
    m_token_Dic.RemoveAll
    Set m_token_Dic = Nothing
End Sub

Private Function UniStrToUTF8(UniString) As Byte()
   ' Convert a Unicode string to a byte stream of UTF-8
   Dim BArray() As Byte
   Dim TempB() As Byte
   Dim i As Long
   Dim k As Long
   Dim TLen As Long
   Dim b1 As Byte
   Dim b2 As Byte
   Dim UTF16 As Long
   Dim j
   TLen = Len(UniString)  ' Obtain length of Unicode input string
   If TLen = 0 Then Exit Function  ' get out if there's nothing to convert
   k = 0
   For i = 1 To TLen
      ' Work out the UTF16 value of the Unicode character
      CopyMemory b1, ByVal StrPtr(UniString) + ((i - 1) * 2), 1
      CopyMemory b2, ByVal StrPtr(UniString) + ((i - 1) * 2) + 1, 1
      ' Combine the 2 bytes into the Unicode UTF-16
      UTF16 = b2  ' assign b2 to UTF16 before multiplying by 256 to avoid overflow
      UTF16 = UTF16 * 256 + b1
      ' Convert UTF-16 to 2 or 3 bytes of UTF-8
      TempB = ToUTF8(UTF16)
      ' Copy the resultant bytes to BArray
      For j = 0 To UBound(TempB)
         ReDim Preserve BArray(k)
         BArray(k) = TempB(j): k = k + 1
      Next
      ReDim TempB(0)
   Next
   UniStrToUTF8 = BArray  ' Return the resultant UTF-8 byte array
End Function

Private Function ToUTF8(ByVal UTF16 As Long) As Byte()
   ' Convert a 16bit UTF-16BE to 2 or 3 UTF-8 bytes
   Dim BArray() As Byte
   If UTF16 < &H80 Then
      ReDim BArray(0)  ' one byte UTF-8
      BArray(0) = UTF16  ' Use number as is
   ElseIf UTF16 < &H800 Then
      ReDim BArray(1)  ' two byte UTF-8
      BArray(1) = &H80 + (UTF16 And &H3F)  ' Least Significant 6 bits
      UTF16 = UTF16 \ &H40  ' Shift UTF16 number right 6 bits
      BArray(0) = &HC0 + (UTF16 And &H1F)  ' Use 5 remaining bits
   Else
      ReDim BArray(2)  ' three byte UTF-8
      BArray(2) = &H80 + (UTF16 And &H3F)  ' Least Significant 6 bits
      UTF16 = UTF16 \ &H40  ' Shift UTF16 number right 6 bits
      BArray(1) = &H80 + (UTF16 And &H3F)  ' Use next 6 bits
      UTF16 = UTF16 \ &H40  ' Shift UTF16 number right 6 bits again
      BArray(0) = &HE0 + (UTF16 And &HF)  ' Use 4 remaining bits
   End If
   ToUTF8 = BArray  ' Return UTF-8 bytes in an array
End Function
'최근 오류코드 확인
Public Property Get LastErrCode() As Long
    LastErrCode = m_LastErrCode
End Property
'최근오류메시지 확인
Public Property Get LastErrMessage() As String
    LastErrMessage = m_LastErrMessage
End Property
'테스트 플래그
Public Property Let IsTest(ByVal value As Boolean)
    m_IsTest = value
End Property
'초기화
Public Sub Initialize(LinkID As String, SecretKey As String)
    Set m_Linkhub = New Linkhub
    
    m_Linkhub.LinkID = LinkID
    m_Linkhub.SercetKey = SecretKey
    
    m_scope.Add "member"
    
    '홈택스 전자세금계산서 연계
    m_scope.Add "111"
    
    m_IsTest = False
    
End Sub
Private Function getSession_token(CorpNum As String) As String
    Dim refresh As Boolean
    
    refresh = False
    Dim m_Token As LinkhubToken
    
    Set m_Token = Nothing
    
    If m_token_Dic.Exists(CorpNum) Then
        Set m_Token = m_token_Dic.Item(CorpNum)
    End If
    
    If m_Token Is Nothing Then
        refresh = True
    Else
        Dim utcnow As String
        utcnow = m_Linkhub.GetTime()
        refresh = m_Token.expiration < utcnow
    End If
    
    If refresh Then
    
        Set m_Token = m_Linkhub.getToken(IIf(m_IsTest, ServiceID_TEST, ServiceID_REAL), CorpNum, m_scope)
        
        If m_Token Is Nothing Then
            Err.Raise m_Linkhub.LastErrCode, "POPBILL", m_Linkhub.LastErrMessage
            Exit Function
        End If
        
        If m_token_Dic.Exists(CorpNum) Then
            m_token_Dic.Remove (CorpNum)
        End If
        
        m_token_Dic.Add CorpNum, m_Token
        
    End If
    
    getSession_token = m_Token.session_token

End Function

Private Function httpGET(url As String, BearerToken As String, Optional UserID As String = "") As Variant
     
    Dim winhttp1
    
    Set winhttp1 = CreateObject("MSXML2.XMLHTTP.6.0")
    Call winhttp1.Open("GET", IIf(m_IsTest, ServiceURL_TEST, ServiceURL_REAL) + url, False)
    
    Call winhttp1.setRequestHeader("Authorization", "Bearer " + BearerToken)
    Call winhttp1.setRequestHeader("x-pb-version", APIVersion)
    Call winhttp1.setRequestHeader("Accept-Encoding", "gzip,deflate")
    
    If UserID <> "" Then
        Call winhttp1.setRequestHeader("x-pb-userid", UserID)
    End If
    
    
    winhttp1.Send
    
    Dim Response As String
    Response = winhttp1.responseText
       
    Dim parsedDic As Object
    
    If winhttp1.Status <> 200 Then
    
        Set parsedDic = m_Linkhub.parse(Response)
        
        Err.Raise parsedDic.Item("code"), "POPBILL", parsedDic.Item("message")
        
        Exit Function
        
    End If
    
    Set winhttp1 = Nothing
    
    Set httpGET = m_Linkhub.parse(Response)

End Function

Private Function httpPOST(url As String, BearerToken As String, postdata As String, Optional UserID As String) As Variant
    Set httpPOST = httpPOST_override(url, BearerToken, "", postdata, UserID)
    
End Function

Private Function httpPOST_override(url As String, BearerToken As String, override As String, postdata As String, Optional UserID As String) As Variant
     
    Dim winhttp1
    
    Set winhttp1 = CreateObject("MSXML2.XMLHTTP.6.0")
    Call winhttp1.Open("POST", IIf(m_IsTest, ServiceURL_TEST, ServiceURL_REAL) + url, False)
    Call winhttp1.setRequestHeader("x-pb-version", APIVersion)
    Call winhttp1.setRequestHeader("Accept-Encoding", "gzip,deflate")
    
    If BearerToken <> "" Then
        Call winhttp1.setRequestHeader("Authorization", "Bearer " + BearerToken)
    End If
    
    If override <> "" Then
        Call winhttp1.setRequestHeader("X-HTTP-Method-Override", override)
    End If
    
    If UserID <> "" Then
        Call winhttp1.setRequestHeader("x-pb-userid", UserID)
    End If
    
    Call winhttp1.setRequestHeader("Content-Type", "Application/json")
    
    winhttp1.Send (postdata)
        
    Dim Response As String
    Response = winhttp1.responseText
       
    Dim parsedDic As Object
    
    If winhttp1.Status <> 200 Then
    
        Set parsedDic = m_Linkhub.parse(Response)
        
        Err.Raise parsedDic.Item("code"), "POPBILL", IIf(IsNull(parsedDic.Item("message")), "Null Exception", parsedDic.Item("message"))
        
        Exit Function
        
    End If
    
    Set winhttp1 = Nothing
    
    Set httpPOST_override = m_Linkhub.parse(Response)

End Function

'연동회원 아이디 중복 확인
Public Function CheckID(IDString As String) As PBResponse
On Error GoTo ErrHandler
    Dim result As Variant
    
    Set result = httpGET("/IDCheck?ID=" + IDString, "")
    
    Set CheckID = New PBResponse
    CheckID.code = result.Item("code")
    CheckID.Message = result.Item("message")
    Exit Function
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    Set CheckID = Nothing
End Function
'연동회원 담당자 목록조회
Public Function ListContact(CorpNum As String, Optional UserID As String = "") As Collection
On Error GoTo ErrHandler
        
    Dim result As Variant
       
    Set result = httpGET("/IDs", getSession_token(CorpNum), UserID)
    
    Set ListContact = New Collection
    
    Dim T As Variant
    Dim info As PBContactInfo
    
    For Each T In result
        Set info = New PBContactInfo
        
        info.fromDictionary T
        
        ListContact.Add info
    Next
    
    Exit Function
    
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    Set ListContact = Nothing
End Function
'담당자 추가
Public Function RegistContact(CorpNum As String, ContactInfo As PBContactInfo, Optional UserID As String = "") As PBResponse
On Error GoTo ErrHandler
    Set RegistContact = Nothing
    
    Dim postdata As String
    Dim tmp
    
    Set tmp = CreateObject("Scripting.Dictionary")
    
    tmp.Add "id", ContactInfo.id
    tmp.Add "pwd", ContactInfo.pwd
    tmp.Add "personName", ContactInfo.personName
    tmp.Add "tel", ContactInfo.tel
    tmp.Add "hp", ContactInfo.hp
    tmp.Add "fax", ContactInfo.fax
    tmp.Add "email", ContactInfo.email
    tmp.Add "searchAllAllowYN", ContactInfo.searchAllAllowYN
    tmp.Add "mgrYN", ContactInfo.mgrYN
            
    postdata = m_Linkhub.toString(tmp)
    
    Dim result As Variant
    
    Set result = httpPOST("/IDs/New", getSession_token(CorpNum), postdata)
    
    Set RegistContact = New PBResponse
    
    RegistContact.code = result.Item("code")
    RegistContact.Message = result.Item("message")
    Exit Function
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
End Function
'담당자 정보 수정
Public Function UpdateContact(CorpNum As String, ContactInfo As PBContactInfo, UserID As String) As PBResponse
On Error GoTo ErrHandler
    Set UpdateContact = Nothing
    
    Dim postdata As String
    Dim tmp
    
    Set tmp = CreateObject("Scripting.Dictionary")
    
    tmp.Add "id", ContactInfo.id
    tmp.Add "personName", ContactInfo.personName
    tmp.Add "tel", ContactInfo.tel
    tmp.Add "hp", ContactInfo.hp
    tmp.Add "fax", ContactInfo.fax
    tmp.Add "email", ContactInfo.email
    tmp.Add "searchAllAllowYN", ContactInfo.searchAllAllowYN
    tmp.Add "mgrYN", ContactInfo.mgrYN
            
    postdata = m_Linkhub.toString(tmp)
    
    Dim result As Variant
    
    Set result = httpPOST("/IDs", getSession_token(CorpNum), postdata, UserID)
    
    Set UpdateContact = New PBResponse
    
    UpdateContact.code = result.Item("code")
    UpdateContact.Message = result.Item("message")
    Exit Function
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
End Function
'회사정보 수정
Public Function UpdateCorpInfo(CorpNum As String, CorpInfo As PBCorpInfo, Optional UserID As String = "") As PBResponse
On Error GoTo ErrHandler
    Set UpdateCorpInfo = Nothing
    
    Dim postdata As String
    Dim tmp
    
    Set tmp = CreateObject("Scripting.Dictionary")
    
    tmp.Add "ceoname", CorpInfo.ceoname
    tmp.Add "corpName", CorpInfo.corpName
    tmp.Add "addr", CorpInfo.addr
    tmp.Add "bizType", CorpInfo.bizType
    tmp.Add "bizClass", CorpInfo.bizClass
            
    postdata = m_Linkhub.toString(tmp)
    
    Dim result As Variant
    
    Set result = httpPOST("/CorpInfo", getSession_token(CorpNum), postdata, UserID)
    
    Set UpdateCorpInfo = New PBResponse
    
    UpdateCorpInfo.code = result.Item("code")
    UpdateCorpInfo.Message = result.Item("message")
    Exit Function
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
End Function
'회원잔액조회
Public Function GetBalance(CorpNum As String) As Double
    On Error GoTo ErrHandler
    
    GetBalance = m_Linkhub.GetBalance(getSession_token(CorpNum), IIf(m_IsTest, ServiceID_TEST, ServiceID_REAL))
    
    Exit Function
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    GetBalance = m_LastErrCode
End Function
'파트너 잔액조회
Public Function GetPartnerBalance(CorpNum As String) As Double
    On Error GoTo ErrHandler
    
    GetPartnerBalance = m_Linkhub.GetPartnerBalance(getSession_token(CorpNum), IIf(m_IsTest, ServiceID_TEST, ServiceID_REAL))
    
    Exit Function
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    GetPartnerBalance = m_LastErrCode
End Function

'파트너 포인트 충전 URL - 2017/08/30 추가
Public Function GetPartnerURL(CorpNum As String, TOGO As String) As String
    On Error GoTo ErrHandler
    
    GetParnterURL = m_Linkhub.GetPartnerURL(getSession_token(CorpNum), IIf(m_IsTest, ServiceID_TEST, ServiceID_REAL), TOGO)
    
    Exit Function
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    GetPartnerURL = m_LastErrCode
End Function

'팝빌 기본 URL
Public Function GetPopbillURL(CorpNum As String, UserID As String, TOGO As String) As String
On Error GoTo ErrHandler
    Dim result As Variant
    
    Set result = httpGET("/?TG=" + TOGO, getSession_token(CorpNum), UserID)
    
    GetPopbillURL = result.Item("url")
    Exit Function
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    GetPopbillURL = ""
End Function
'회원가입 여부
Public Function CheckIsMember(CorpNum As String, LinkID As String) As PBResponse
On Error GoTo ErrHandler
    Dim result As Variant
    
    Set result = httpGET("/Join?CorpNum=" + CorpNum + "&LID=" + LinkID, "")
    
    Set CheckIsMember = New PBResponse
    CheckIsMember.code = result.Item("code")
    CheckIsMember.Message = result.Item("message")
    Exit Function
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    Set CheckIsMember = Nothing
End Function
'회원가입
Public Function JoinMember(JoinInfo As PBJoinForm) As PBResponse
On Error GoTo ErrHandler
    Set JoinMember = Nothing
    
    Dim postdata As String
    Dim tmp
    
    Set tmp = CreateObject("Scripting.Dictionary")
    
    tmp.Add "LinkID", JoinInfo.LinkID
    tmp.Add "CorpNum", JoinInfo.CorpNum
    tmp.Add "CorpName", JoinInfo.corpName
    tmp.Add "CEOName", JoinInfo.ceoname
    tmp.Add "Addr", JoinInfo.addr
    tmp.Add "ZipCode", JoinInfo.ZipCode
    tmp.Add "BizClass", JoinInfo.bizClass
    tmp.Add "BizType", JoinInfo.bizType
    tmp.Add "ContactName", JoinInfo.ContactName
    tmp.Add "ContactEmail", JoinInfo.ContactEmail
    tmp.Add "ContactFAX", JoinInfo.ContactFAX
    tmp.Add "ContactHP", JoinInfo.ContactHP
    tmp.Add "ContactTEL", JoinInfo.ContactTEL
    tmp.Add "ID", JoinInfo.id
    tmp.Add "PWD", JoinInfo.pwd
    
    postdata = m_Linkhub.toString(tmp)
    
    Dim result As Variant
    
    Set result = httpPOST("/Join", "", postdata)
    
    Set JoinMember = New PBResponse
    
    JoinMember.code = result.Item("code")
    JoinMember.Message = result.Item("message")
    Exit Function
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
End Function
'회사정보 조회
Public Function GetCorpInfo(CorpNum As String, Optional UserID As String = "") As PBCorpInfo
On Error GoTo ErrHandler
    Dim result As Variant
       
    Set result = httpGET("/CorpInfo", getSession_token(CorpNum), UserID)
    
    Set GetCorpInfo = New PBCorpInfo
        
    GetCorpInfo.fromDictionary result
    Exit Function
    
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    Set GetCorpInfo = Nothing
End Function
'과금정보 확인
Public Function GetChargeInfo(CorpNum As String, Optional UserID As String = "") As PBChargeInfo
On Error GoTo ErrHandler
    Dim result As Variant
       
    Set result = httpGET("/HomeTax/Taxinvoice/ChargeInfo", getSession_token(CorpNum), UserID)
    
    Set GetChargeInfo = New PBChargeInfo
        
    GetChargeInfo.fromDictionary result
    Exit Function
    
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    Set GetChargeInfo = Nothing
End Function

'수집 요청
Public Function RequestJob(CorpNum As String, tiType As KeyType, DType As String, SDate As String, EDate As String, Optional UserID As String = "") As String
On Error GoTo ErrHandler
    Dim result As Variant
    
    Dim strKeyType As String
    
    strKeyType = IIf(tiType = TRUSTEE, "TRUSTEE", IIf(tiType = BUY, "BUY", "SELL"))
    
    Set result = httpPOST("/HomeTax/Taxinvoice/" + strKeyType + "?DType=" + DType + "&SDate=" + SDate + "&EDate=" + EDate, getSession_token(CorpNum), "", UserID)
    
    RequestJob = result.Item("jobID")
    Exit Function
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    RequestJob = ""
End Function

Public Function GetJobState(CorpNum As String, jobID As String, Optional UserID As String = "") As PBHTTaxinvoiceJobState
On Error GoTo ErrHandler
    If Not Len(jobID) = 18 Then
        Err.Raise -99999999, "POPBILL", "작업아이디(jobID)가 올바르지 않습니다."
    End If
        
    Dim result As Variant
        
    Set result = httpGET("/HomeTax/Taxinvoice/" + jobID + "/State", getSession_token(CorpNum), UserID)
    
    Set GetJobState = New PBHTTaxinvoiceJobState
        
    GetJobState.fromDictionary result
        
    Exit Function
    
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    Set GetJobState = Nothing
End Function

Public Function ListActiveJob(CorpNum As String, Optional UserID As String = "") As Collection
On Error GoTo ErrHandler
   
    Dim result As Variant
    
    Set result = httpGET("/HomeTax/Taxinvoice/JobList", getSession_token(CorpNum), UserID)
    
    Set ListActiveJob = New Collection
    
    Dim T As Variant
    Dim info As PBHTTaxinvoiceJobState
    
    For Each T In result
        Set info = New PBHTTaxinvoiceJobState
        
        info.fromDictionary T
        
        ListActiveJob.Add info
    
    Next
    
    Exit Function
    
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    Set ListActiveJob = Nothing
End Function

Public Function Search(CorpNum As String, jobID As String, tiType As Collection, taxType As Collection, purposeType As Collection, taxRegIDType As String, taxRegID As String, taxRegIDYN As String, page As Integer, perPage As Integer, order As String, Optional UserID As String = "") As PBHTTaxinvoiceSearch
On Error GoTo ErrHandler

    If Not Len(jobID) = 18 Then
        Err.Raise -99999999, "POPBILL", "작업아이디(jobID)가 올바르지 않습니다."
    End If
    
    Dim result As Variant
    
    Dim uri As String
    Dim i As Integer
       
    uri = "/HomeTax/Taxinvoice/" + jobID
        
    If tiType.count > 0 Then
        uri = uri + "?Type="
        For i = 1 To tiType.count
            If i <> tiType.count Then
                uri = uri + tiType(i) + ","
            Else
                uri = uri + tiType(i)
            End If
        Next
    End If
        
    If taxType.count > 0 Then
        uri = uri + "&TaxType="
        For i = 1 To taxType.count
            If i <> taxType.count Then
                uri = uri + taxType(i) + ","
            Else
                uri = uri + taxType(i)
            End If
        Next
        
    End If
    
    If purposeType.count > 0 Then
        uri = uri + "&PurposeType="
        For i = 1 To purposeType.count
            If i <> purposeType.count Then
                uri = uri + purposeType(i) + ","
            Else
                uri = uri + purposeType(i)
            End If
        Next
        
    End If
       
    uri = uri + "&TaxRegIDType=" + taxRegIDType
    
    If taxRegID <> "" Then
        uri = uri + "&TaxRegID=" + taxRegID
    End If
       
    
    If taxRegIDYN <> "" Then
        uri = uri + "&TaxRegIDYN=" + taxRegIDYN
    End If
    
    page = IIf(page < 1, 1, page)
    perPage = IIf(perPage < 1, 500, perPage)
    uri = uri + "&Page=" + CStr(page)
    uri = uri + "&PerPage=" + CStr(perPage)
    uri = uri + "&Order=" + order
    
    Set result = httpGET(uri, getSession_token(CorpNum), UserID)
    
    Set Search = New PBHTTaxinvoiceSearch
    
    Search.fromDictionary result
    
    Exit Function
    
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    Set Search = Nothing
End Function

Public Function Summary(CorpNum As String, jobID As String, tiType As Collection, taxType As Collection, purposeType As Collection, taxRegIDType As String, taxRegID As String, taxRegIDYN As String, Optional UserID As String = "") As PBHTTaxinvoiceSummary
On Error GoTo ErrHandler

    If Not Len(jobID) = 18 Then
        Err.Raise -99999999, "POPBILL", "작업아이디(jobID)가 올바르지 않습니다."
    End If
    
    
    
    Dim result As Variant
    
    Dim uri As String
    Dim i As Integer
       
    uri = "/HomeTax/Taxinvoice/" + jobID + "/Summary"
        
    If tiType.count > 0 Then
        uri = uri + "?Type="
        For i = 1 To tiType.count
            If i <> tiType.count Then
                uri = uri + tiType(i) + ","
            Else
                uri = uri + tiType(i)
            End If
        Next
    End If
        
    If taxType.count > 0 Then
        uri = uri + "&TaxType="
        For i = 1 To taxType.count
            If i <> taxType.count Then
                uri = uri + taxType(i) + ","
            Else
                uri = uri + taxType(i)
            End If
        Next
        
    End If
    
    If purposeType.count > 0 Then
        uri = uri + "&PurposeType="
        For i = 1 To purposeType.count
            If i <> purposeType.count Then
                uri = uri + purposeType(i) + ","
            Else
                uri = uri + purposeType(i)
            End If
        Next
        
    End If
       
    uri = uri + "&TaxRegIDType=" + taxRegIDType
    
    If taxRegID <> "" Then
        uri = uri + "&TaxRegID=" + taxRegID
    End If
        
    If taxRegIDYN <> "" Then
        uri = uri + "&TaxRegIDYN=" + taxRegIDYN
    End If
    
        
    Set result = httpGET(uri, getSession_token(CorpNum), UserID)
    
    Set Summary = New PBHTTaxinvoiceSummary
    
    Summary.fromDictionary result
    
    Exit Function
    
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    Set Summary = Nothing
End Function

Public Function GetTaxinvoice(CorpNum As String, ntsconfirmNum As String, Optional UserID As String = "") As PBHTTaxinvoice
On Error GoTo ErrHandler
    If Not Len(ntsconfirmNum) = 24 Then
        Err.Raise -99999999, "POPBILL", "국세청승인번호가 올바르지 않았습니다."
    End If
    
    Dim result As Variant
    
    Set result = httpGET("/HomeTax/Taxinvoice/" + ntsconfirmNum, getSession_token(CorpNum), UserID)
        
    Set GetTaxinvoice = New PBHTTaxinvoice
        
    GetTaxinvoice.fromDictionary result
        
    Exit Function
    
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    Set GetTaxinvoice = Nothing
End Function

Public Function GetXML(CorpNum As String, ntsconfirmNum As String, Optional UserID As String = "") As PBHTTaxinvoiceXML
On Error GoTo ErrHandler
    If Not Len(ntsconfirmNum) = 24 Then
        Err.Raise -99999999, "POPBILL", "국세청승인번호가 올바르지 않았습니다."
    End If
    
    Dim result As Variant
        
    Set result = httpGET("/HomeTax/Taxinvoice/" + ntsconfirmNum + "?T=xml", getSession_token(CorpNum), UserID)
    
    Set GetXML = New PBHTTaxinvoiceXML
        
    GetXML.fromDictionary result
        
    Exit Function
    
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    Set GetXML = Nothing
End Function

Public Function GetFlatRatePopUpURL(CorpNum As String, Optional UserID As String = "") As String
On Error GoTo ErrHandler
    Dim result As Variant
    
    Set result = httpGET("/HomeTax/Taxinvoice?TG=CHRG", getSession_token(CorpNum), UserID)
    
    GetFlatRatePopUpURL = result.Item("url")
    Exit Function
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    GetFlatRatePopUpURL = ""
End Function

Public Function GetCertificatePopUpURL(CorpNum As String, Optional UserID As String = "") As String
On Error GoTo ErrHandler
    Dim result As Variant
    
    Set result = httpGET("/HomeTax/Taxinvoice?TG=CERT", getSession_token(CorpNum), UserID)
    
    GetCertificatePopUpURL = result.Item("url")
    Exit Function
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    GetCertificatePopUpURL = ""
End Function

Public Function GetFlatRateState(CorpNum As String, Optional UserID As String = "") As PBHTTaxinvoiceFlatRate
On Error GoTo ErrHandler
        
    Dim result As Variant
        
    Set result = httpGET("/HomeTax/Taxinvoice/Contract", getSession_token(CorpNum), UserID)
    
    Set GetFlatRateState = New PBHTTaxinvoiceFlatRate
        
    GetFlatRateState.fromDictionary result
        
    Exit Function
    
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    Set GetFlatRateState = Nothing
End Function

Public Function GetCertificateExpireDate(CorpNum As String, Optional UserID As String = "") As String
On Error GoTo ErrHandler
    Dim result As Variant
    
    Set result = httpGET("/HomeTax/Taxinvoice/CertInfo", getSession_token(CorpNum), UserID)
    
    GetCertificateExpireDate = result.Item("certificateExpiration")
    Exit Function
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    GetCertificateExpireDate = ""
End Function

'홈택스 전자세금계산서 보기 팝업 URL
Public Function GetPopUpURL(CorpNum As String, ntsconfirmNum As String, Optional UserID As String = "") As String
On Error GoTo ErrHandler
    If Not Len(ntsconfirmNum) = 24 Then
        Err.Raise -99999999, "POPBILL", "국세청승인번호가 올바르지 않았습니다."
    End If

    Dim result As Variant
    
    Set result = httpGET("/HomeTax/Taxinvoice/" + ntsconfirmNum + "/PopUp", getSession_token(CorpNum), UserID)
    
    GetPopUpURL = result.Item("url")
    Exit Function
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    GetPopUpURL = ""
End Function


'홈택스 공인인증서 로그인 테스트
Public Function CheckCertValidation(CorpNum As String, Optional UserID As String = "") As PBResponse
On Error GoTo ErrHandler
    Dim result As Variant
    
    Set result = httpGET("/HomeTax/Taxinvoice/CertCheck", getSession_token(CorpNum), UserID)
    
    Set CheckCertValidation = New PBResponse
    CheckCertValidation.code = result.Item("code")
    CheckCertValidation.Message = result.Item("message")
    Exit Function
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    Set CheckCertValidation = Nothing
End Function

'부서사용자 계정등록
Public Function RegistDeptUser(CorpNum As String, DeptUserID As String, DeptUserPWD As String, Optional UserID As String = "") As PBResponse
On Error GoTo ErrHandler
    Set RegistDeptUser = Nothing
    
    Dim postdata As String
    Dim tmp
    
    Set tmp = CreateObject("Scripting.Dictionary")
    
    tmp.Add "id", DeptUserID
    tmp.Add "pwd", DeptUserPWD
            
    postdata = m_Linkhub.toString(tmp)
    
    Dim result As Variant
    
    Set result = httpPOST("/HomeTax/Taxinvoice/DeptUser", getSession_token(CorpNum), postdata, UserID)
    
    Set RegistDeptUser = New PBResponse
    
    RegistDeptUser.code = result.Item("code")
    RegistDeptUser.Message = result.Item("message")
    Exit Function
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
End Function

'부서사용자 등록정보 확인
Public Function CheckDeptUser(CorpNum As String, Optional UserID As String = "") As PBResponse
On Error GoTo ErrHandler
    Dim result As Variant
    
    Set result = httpGET("/HomeTax/Taxinvoice/DeptUser", getSession_token(CorpNum), UserID)
    
    Set CheckDeptUser = New PBResponse
    CheckDeptUser.code = result.Item("code")
    CheckDeptUser.Message = result.Item("message")
    Exit Function
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    Set CheckDeptUser = Nothing
End Function

'부서사용자 로그인 테스트
Public Function CheckLoginDeptUser(CorpNum As String, Optional UserID As String = "") As PBResponse
On Error GoTo ErrHandler
    Dim result As Variant
    
    Set result = httpGET("/HomeTax/Taxinvoice/DeptUser/Check", getSession_token(CorpNum), UserID)
    
    Set CheckLoginDeptUser = New PBResponse
    CheckLoginDeptUser.code = result.Item("code")
    CheckLoginDeptUser.Message = result.Item("message")
    Exit Function
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    Set CheckLoginDeptUser = Nothing
End Function

'부서사용자 등록정보 삭제
Public Function DeleteDeptUser(CorpNum As String, Optional UserID As String = "") As PBResponse
On Error GoTo ErrHandler
    Set DeleteDeptUser = Nothing
    
    Dim result As Variant
    
    Set result = httpPOST_override("/HomeTax/Taxinvoice/DeptUser", getSession_token(CorpNum), "DELETE", "", UserID)
    Set DeleteDeptUser = New PBResponse
    
    DeleteDeptUser.code = result.Item("code")
    DeleteDeptUser.Message = result.Item("message")
    Exit Function
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    Set DeleteDeptUser = Nothing
End Function
