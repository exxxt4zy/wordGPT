Option Explicit

Sub DeleteShortcut()
    Dim DeleteControl As CommandBarControl
    For Each DeleteControl In Application.CommandBars("Text").Controls
       If DeleteControl.Caption = "Ask ChatGPT" Then
          DeleteControl.Delete
       End If
    Next DeleteControl
End Sub

Sub AddToShortcut()
    Dim Bar As CommandBar
    Dim NewControl As CommandBarButton
    Set Bar = Application.CommandBars("Text")
    Set NewControl = Bar.Controls.Add(Type:=msoControlButton, ID:=1, Temporary:=False)
    With NewControl
        .Caption = "Ask ChatGPT"
        .OnAction = "Ask"
        .Style = msoButtonIconAndCaption
    End With
End Sub


Private Sub Ask()

    Dim selection As selection
    Set selection = Application.selection
    Dim selectedText As String
    Dim rng As Word.Range
    
    selectedText = Replace(selection.text, ChrW$(13), "")
    
    Dim req As Object
    Set req = CreateObject("WinHttp.WinHttpRequest.5.1")

    req.Open "POST", "https://chatgpt.qquade.repl.co/chat", False
    req.SetRequestHeader "Content-Type", "application/json"
    
    Dim f
    Dim json As String
    
    Dim text As String
    
    text = selectedText

    json = "{""message"": """ & text & """}"
    
    req.Send json
    
    req.WaitForResponse
    
    Dim objHTML, objWin As Object
    Set objHTML = CreateObject("HTMLFile")
    Set objWin = objHTML.parentWindow
    objWin.execScript "var data = " & req.ResponseText & ";", "JScript"
    objWin.execScript "var response_msg = data.choices[0].message.content;", "JScript"

    Dim result As String
    result = objWin.response_msg
    
    selection.Range.InsertAfter Chr(10) & result & Chr(10)
    
End Sub


Private Sub document_open()
    'adds the right-click shortcut when the document opens
    Call AddToShortcut
End Sub
