VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMessageManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private mlstSubscribers   As CList

'
' Class error context
'
Private mlErr     As Long
Private msErrCtx  As String
Private msErr     As String

Private Sub ClearErr()
  mlErr = 0&
  msErr = ""
  msErrCtx = ""
End Sub

Private Sub SetErr(ByVal psErrCtx As String, ByVal plErr As Long, ByVal psErr As String)
  mlErr = plErr
  msErr = psErr
  msErrCtx = psErrCtx
End Sub

Public Property Get LastErr() As Long
  LastErr = mlErr
End Property

Public Property Get LastErrDesc() As String
  LastErrDesc = msErr
End Property

Public Property Get LastErrContext() As String
  LastErrContext = msErrCtx
End Property

Private Sub Class_Initialize()
  Const LOCAL_ERR_CTX As String = "Class_Initialize"
  On Error GoTo Class_Initialize_Err
  ClearErr
  
  Set mlstSubscribers = New CList
  mlstSubscribers.ArrayDefine Array("Topic", "ClientID", "IMessageReceiver"), Array(vbString, vbString, vbObject)
  
Class_Initialize_Exit:
  Exit Sub
Class_Initialize_Err:
  SetErr LOCAL_ERR_CTX, Err.Number, Err.Description
  Resume Class_Initialize_Exit
End Sub

Private Sub Class_Terminate()
  On Error Resume Next
  Set mlstSubscribers = Nothing
End Sub

Public Function IsSubscribed(ByVal psClientID As String, ByVal psTopic As String) As Boolean
  Dim iFind   As Long
  iFind = mlstSubscribers.Find(Array("Topic", "ClientID"), Array(psTopic, psClientID))
  If iFind > 0 Then
    IsSubscribed = True
  End If
End Function

Public Function Subscribe(ByRef pIIMessageReceiver As IMessageReceiver, ByVal psTopic As String) As Boolean
  Const LOCAL_ERR_CTX As String = "Subscribe"
  On Error GoTo Subscribe_Err
  ClearErr
  
  Dim sClientID As String
  
  sClientID = pIIMessageReceiver.ClientID
  
  If IsSubscribed(sClientID, psTopic) Then
    Subscribe = True
    Exit Function
  End If
  
  mlstSubscribers.AddValues psTopic, sClientID, pIIMessageReceiver
  If mlstSubscribers.Count > 2 Then
    mlstSubscribers.Sort "Topic;ClientID"
  End If
  
  Subscribe = True
  
Subscribe_Exit:
  Exit Function

Subscribe_Err:
  SetErr LOCAL_ERR_CTX, Err.Number, Err.Description
  Resume Subscribe_Exit
End Function

Public Function SubscribeMulti(ByRef pIIMessageReceiver As IMessageReceiver, pavTopics As Variant) As Boolean
  Const LOCAL_ERR_CTX As String = "SubscribeMulti"
  On Error GoTo SubscribeMulti_Err
  ClearErr
  
  Dim sClientID   As String
  Dim vTopic      As Variant
  Dim sTopic      As String
  Dim iSubscrCt   As Integer
  
  sClientID = pIIMessageReceiver.ClientID
  
  For Each vTopic In pavTopics
    If Not IsSubscribed(sClientID, vTopic) Then
      sTopic = vTopic
      mlstSubscribers.AddValues sTopic, sClientID, pIIMessageReceiver
      iSubscrCt = iSubscrCt + 1
    End If
  Next
  If iSubscrCt > 0 Then
    If mlstSubscribers.Count > 2 Then
      mlstSubscribers.Sort "Topic;ClientID"
    End If
  End If
  
  SubscribeMulti = True
  
SubscribeMulti_Exit:
  Exit Function

SubscribeMulti_Err:
  SetErr LOCAL_ERR_CTX, Err.Number, Err.Description
  Resume SubscribeMulti_Exit
End Function

'if psTopic = "" then unsubscribe from all topics
Public Sub Unsubscribe(ByVal psClientID As String, ByVal psTopic As String)
  Dim i             As Long
  Dim fFound        As Boolean
  
  On Error Resume Next
  
  If Len(psTopic) = 0 Then
    Do
      fFound = False
      For i = 1 To mlstSubscribers.Count
        If mlstSubscribers("ClientID", i) = psClientID Then
          fFound = True
          mlstSubscribers.Remove i
          Exit For
        End If
      Next i
    Loop Until fFound = False
  Else
    i = mlstSubscribers.Find(Array("Topic", "ClientID"), Array(psTopic, psClientID))
    If i > 0 Then
      mlstSubscribers.Remove i
    End If
  End If
End Sub

'A subscriber can return a value <> 0& to break the broadcase
Public Function Broadcast( _
    ByVal psClientID As String, _
    ByVal psTopic As String, _
    ByVal pvData As Variant, _
    Optional ByVal psTargetClientIDS As String _
  ) As Long
  Const LOCAL_ERR_CTX As String = "Broadcast"
  On Error GoTo Broadcast_Err
  ClearErr
  
  Dim iFind           As Long
  Dim i               As Long
  Dim iiMsgReceiver   As IMessageReceiver
  Dim lRet            As Long
  Dim sReceiverID     As String
  Dim iTargetIDCt     As Integer
  Dim asTargetID()    As String
  Dim fDoSend         As Boolean
  Dim k               As Integer
  
  iFind = mlstSubscribers.FindFirst("Topic", psTopic)
  If iFind <= 0 Then
    'nobody to notify
    Broadcast = 0&
    Exit Function
  End If
  
  If Len(psTargetClientIDS) > 0 Then
    iTargetIDCt = SplitString(asTargetID(), psTargetClientIDS, ";")
  End If
  
  For i = 1 To mlstSubscribers.Count
    If mlstSubscribers("Topic", i) = psTopic Then
      sReceiverID = mlstSubscribers("ClientID", i)
      If iTargetIDCt = 0 Then
        fDoSend = True
      Else
        fDoSend = False
        For k = 1 To iTargetIDCt
          If asTargetID(k) = sReceiverID Then
            fDoSend = True
          End If
        Next k
      End If
      If fDoSend Then
        On Error Resume Next
        Set iiMsgReceiver = mlstSubscribers("IMessageReceiver", i)
        If Err.Number = 0 Then
          On Error GoTo Broadcast_Err
          lRet = iiMsgReceiver.OnMessageReceived(psClientID, psTopic, pvData)
          'If a receiver returns anything else than 0, we exit the loop
          If lRet <> 0& Then Exit For
        End If
        On Error GoTo Broadcast_Err
      End If
    End If
  Next i
  
Broadcast_Exit:
  Broadcast = lRet
  Set iiMsgReceiver = Nothing
  Exit Function

Broadcast_Err:
  SetErr LOCAL_ERR_CTX, Err.Number, Err.Description
  Resume Broadcast_Exit
  Resume
End Function

