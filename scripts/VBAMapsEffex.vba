Option Explicit
'// test connecting to maps app with effex
'// this example assumes that it can create, read , generate aliases and update items
'//
'// use this alias for all data items
Const alias = "effex-demo-markers"
Private ec As New cVBAMapsEffex
Private efx As New cEffexApiClientVBA
Const demoUrl = "https://storage.googleapis.com/effex-console-static/demos/effex-demo-markers/index.html"


Function storeBoss()

    '' the boss key comes from the console
    ec.setProperty "boss", JSONParse("{'key':'----yourbosskey from effex console----'}")
    
End Function
Function addItemsToMenu()
    With CommandBars("standard")
        '//clear up any previous references
        If (isinCollection(.Controls, "effex")) Then
            .Controls("effex").Delete
        End If
        With .Controls.add(msoControlPopup)
            .Caption = "effex"
            With .Controls.add(msoControlButton)
                .Caption = "create in store"
                .OnAction = "createFromSheet"
            End With
            With .Controls.add(msoControlButton)
                .Caption = "update store"
                .OnAction = "updateFromSheet"
            End With
            With .Controls.add(msoControlButton)
                .Caption = "pull from store"
                .OnAction = "pullSheet"
            End With
        End With
    End With
End Function

'/**
' * pull and populate active sheet from effex
' */
Public Sub pullSheet()
'
'  // we'll need these are in the registry

    
    Dim keys As cJobject
    Set keys = ec.getProperty("keys")
    Debug.Assert isSomething(keys)
    
    Dim result As cJobject, values As Variant
    
    ' // get the data using the updater key
    Set result = ec.pullFromEffex(keys.toString("updater"), alias)
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

Public Sub createFromSheet()
    pushItems (True)
End Sub
Public Sub updateFromSheet()
    pushItems (False)
End Sub
'/**
' * push active sheet to effex
' */
Public Sub pushItems(Optional create As Boolean = False)
'  // we'll need these are in the registry


    Dim keys As cJobject
    Set keys = ec.getProperty("keys")
    Debug.Assert isSomething(keys)
    
    Dim result As cJobject, values As Variant
    values = ActiveSheet().UsedRange.value
    
    ' // set the data using the updater key
    If (create) Then
        Set result = ec.pushDataForUpdate(ec.objectify(values), keys, alias)
    Else
        ' updating an existing item
        Set result = efx.update(ec.objectify(values), alias, keys.toString("updater"))
    End If

    If (Not result.cValue("ok")) Then
        MsgBox ("couldnt push data " + JSONStringify(result))
        Exit Sub
    End If
    
    '  // just show what happened
    If (create) Then
        InputBox "heres the link", "maps app", demoUrl + "?updater=" + result.toString("key") + "&item=" + result.toString("alias")
    End If
End Sub

'/**
' * makeing some keys to use for effex
' * @return {object} the keys
' */
Public Function generateKeys()
  '// the boss key is already here

    
    Dim boss As cJobject, keys As cJobject
    Set boss = ec.getProperty("boss")
    Debug.Assert isSomething(boss)
   
    '// make the other keys and store them
    Set keys = ec.makeKeys(boss.toString("key"))
    ec.setProperty "keys", keys
    Debug.Print keys.stringify
End Function
