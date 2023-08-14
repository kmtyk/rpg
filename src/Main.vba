Sub ボタン1_Click()
  Dim namae As String
  namae = InputBox("来てくれてありがとう!お名前は?")
  Dim toi1 As Long
  toi1 = InputBox("城を出た。どっちに行く?(右:1 左:2)")
  If toi1 = 1 Then
    MsgBox "モンスターが現れた。"
    Dim mon1 As Long
    mon1 = InputBox("どうする?(話しかける:1 逃げる:2 アイテム:3)")
    If mon1 = 1 Then
      MsgBox "モンスターが仲間に加わった!"
    Else
      If mon1 = 2 Then
        MsgBox namae & "は、うまく逃げ切ることができた。"
    Else
        MsgBox "アイテムを使った!"
        MsgBox namae & "は体力が完全に回復した!"
    End If
  End If
Else
    MsgBox namae & ("は魔法屋にたどり着いた。")
End If
End Sub
