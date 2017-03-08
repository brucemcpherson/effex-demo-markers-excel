Option Explicit
Private efx As cEffexApiClientVBA
'' / store stuff in the registry
Public Function getProperty(key As String) As cJobject
    
    Dim s As String
    s = GetSetting("xLiberation", _
            "effexmaps", _
            key)
    
    If (s <> vbNullString) Then Set getProperty = JSONParse(s)
    
End Function
Public Function setProperty(key As String, value As cJobject) As cJobject

    SaveSetting "xLiberation", _
            "effexmaps", _
            key, JSONStringify(value)

End Function
''  /**
'  * make a set of keys that can be stored to props service
'  * @param {string} boss the boss key to use
'  * @param {object|null} [params] any params to pass on to the API
'  * @return {object} containg {keySet:{}, results:{}}
'  */
Public Function makeKeys(boss As String, Optional params As cJobject) As cJobject
    
    Dim keys As cJobject, i As Long, a As Variant, result As cJobject
    Set keys = New cJobject
    
    '' // generate one of each kind of key
    With keys.Init(Nothing)
        a = Split("writer,reader,updater", ",")
        For i = LBound(a) To UBound(a)
            Set result = efx.generateKey(boss, CStr(a(i)), params)
            Debug.Assert result.cValue("ok")
            .add CStr(a(i)), result.toString("keys.1")
        Next i
    End With

    Set makeKeys = keys
End Function
'  /**
'  * pull data from sheet and change it back to sheet format
'  * @param {string} readerKey the key to use to read it with
'  * @param {string} itemId the item id or alias
'  */
Public Function pullValues(key As String, itemId As String) As Variant
    Dim result As cJobject
    Set result = Me.pullFromEffex(key, itemId)
    Debug.Assert result.cValue("ok")
    
    '    // now unobjectify data for writing to sheet
    Set pullValues = Me.unObjectify(result.child("value"))
    
End Function
'
'  /**
'  * transform  rows of objects to values
'  * @param {[]} data the data to transform
'  * @param {[[*]]} spreadsheetValues
'  */
Public Function unObjectify(data As cJobject) As Variant
    
    Dim a As Variant, heads As cJobject, job As cJobject, joc As cJobject
    Set heads = New cJobject
    
    '// take the headings from the properties of the first row
    With heads.Init(Nothing).addArray
        ReDim a(0 To data.children.Count, 0 To data.children(1).children.Count - 1)
        For Each joc In data.children(1).children
            .add joc.key
            a(0, joc.childIndex - 1) = joc.key
        Next joc
    End With
    
    ' and then convert the data
    For Each job In data.children
        For Each joc In heads.children
            a(job.childIndex, joc.childIndex - 1) = job.cValue(joc.key)
        Next joc
    Next job
    
    unObjectify = a
End Function

'  /**
'  * create a data item on effex for updating
'  * @param {string} key the key
'  * @param {string} id an alias or id to use
'  * @return {object} the final pull result
'  */
Public Function pullFromEffex(key As String, id As String) As cJobject
    Set pullFromEffex = efx.read(id, key)
End Function

'  /**
'  * push to effex
'  * @param {object} keys a set of keys
'  * @param {object} data the object to write
'  * @param {boolean} [alias] to write - if no alias, then no aliasing will be done
'  * @param {object} [params] any params for the API
'  * @return {object} write results
'  */
Public Function pushDataForUpdate(data As cJobject, keys As cJobject, Optional alias As String, Optional params As cJobject)
    Dim p As cJobject, result As cJobject
    Set p = efx.clone(params)
    
    ' // ensure that the updater has access if not already something else in params
    If (Not isSomething(p.childExists("updaters"))) Then
        p.add ("updaters"), keys.toString("updater")
    End If
    
    '    // now write the item
    Set result = efx.xwrite(data, keys.toString("writer"), "post", p)
    Set pushDataForUpdate = result
    If (Not result.cValue("ok")) Then
        Exit Function
    End If

    ' // if an alias is required, make one
    If (alias <> vbNullString) Then
      Set pushDataForUpdate = _
        efx.registerAlias(keys.toString("writer"), keys.toString("updater"), result.toString("id"), alias)
    End If

End Function

'  /**
'  * objectify a sheet
'  * @param {[[*]]} sheetValues the values
'  * @return {[object]} objectfied data
'  */
Public Function objectify(values As Variant) As cJobject
    Dim ob As cJobject, i As Long, j As Long, v As Variant
    '// make an empty one
    Set ob = efx.clone()

    With ob
        With .add.addArray
            For i = LBound(values) + 1 To UBound(values)
                With .add
                    For j = LBound(values, 2) To UBound(values, 2)
                        '// dont like nulls - make them empty strings
                        v = values(i, j)
                        If (IsEmpty(v)) Then
                            v = ""
                        End If
                        .add CStr(values(LBound(values), j)), v
                    Next j
                End With
            Next i
        End With
    End With
    Set objectify = ob
End Function

'
'
'

Private Sub Class_Initialize()
    Set efx = New cEffexApiClientVBA
    efx.setVerbose True
End Sub