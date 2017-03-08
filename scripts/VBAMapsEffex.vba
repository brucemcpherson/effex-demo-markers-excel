Option Explicit
'// test connecting to maps app with effex

Function storeBoss()
    Dim ec As cVBAMapsEffex
    Set ec = New cVBAMapsEffex
    '' the boss key comes from the console
    ec.setProperty "boss", JSONParse("{'key':'bx2ao-1qw-b74i7saaoc26'}")
    
End Function

'/**
' * pull and populate active sheet from effex
' */
Public Sub pullSheet()
'
'  // we'll need these are in the registry
    Dim ec As cVBAMapsEffex
    Set ec = New cVBAMapsEffex
    
    Dim keys As cJobject
    Set keys = ec.getProperty("keys")
    Debug.Assert isSomething(keys)
    
    Dim result As cJobject, values As Variant
    
    ' // get the data using the updater key
    Set result = ec.pullFromEffex(keys.toString("updater"), "effex-demo-markers")
    If (Not result.cValue("ok")) Then
        MsgBox ("couldnt pull data " + JSONStringify(result))
        Exit Sub
    End If

    '  // now turn data into sheet shaped values
    values = ec.unObjectify(result.child("value"))

    ' // the the active sheet
    ActiveSheet.Cells.ClearContents
    
    ' // dump the data
    If (arrayLength(values) > 0) Then
        ActiveSheet.Range("A1").Resize(arrayLength(values), UBound(values, 2) - LBound(values, 1) + 1) = values
    End If

End Sub


'/**
' * push active sheet to effex
' */
Public Sub pushSheet()
'  // we'll need these are in the registry
    Dim ec As cVBAMapsEffex
    Set ec = New cVBAMapsEffex
    Const demoUrl = "https://storage.googleapis.com/effex-console-static/demos/effex-demo-markers/index.html"
    Dim keys As cJobject
    Set keys = ec.getProperty("keys")
    Debug.Assert isSomething(keys)
    
    Dim result As cJobject, values As Variant
    values = ActiveSheet().UsedRange.value
    
    ' // set the data using the updater key
    Set result = ec.pushDataForUpdate(ec.objectify(values), keys, "effex-demo-markers")
    If (Not result.cValue("ok")) Then
        MsgBox ("couldnt push data " + JSONStringify(result))
        Exit Sub
    End If
    
    '  // just show what happened
    Debug.Print demoUrl + "?updater=" + result.toString("key") + "&item=" + result.toString("alias")

End Sub

'/**
' * makeing some keys to use for effex
' * @return {object} the keys
' */
Public Function generateKeys()
  '// the boss key is already here
    Dim ec As cVBAMapsEffex
    Set ec = New cVBAMapsEffex
    
    Dim boss As cJobject, keys As cJobject
    Set boss = ec.getProperty("boss")
    Debug.Assert isSomething(boss)
   
    '// make the other keys and store them
    Set keys = ec.makeKeys(boss.toString("key"))
    ec.setProperty "keys", keys
    Debug.Print keys.stringify
End Function
