Option Explicit
Private Sub testEfx()
    Dim efx As cEffexApiClientVBA, bossKey As String, result As cJobject, _
        writers As cJobject, readers As cJobject, updaters As cJobject, _
        keys As cJobject, someData As cJobject, data As cJobject, alias As cJobject, _
        otherAlias As cJobject, otherData As cJobject, updateAlias As cJobject, _
        writerAlias As cJobject
        
    
    '// set up client
    Set efx = New cEffexApiClientVBA
    efx.setProd().setVerbose (True)
    
    '// boss key comes from console /// replace this with your own
    bossKey = "bx2ao-1qw-b74i7saaoc26"
    
    '// check service is up
    Set result = efx.ping()
    assure isSomething(result) And result.cValue("ok"), "ping", result

    '  // give up if its not
    If (Not result.cValue("ok")) Then
        MsgBox ("service is not up:" + JSONStringify(result))
        Exit Sub
    End If

    '// get some service info
    Set result = efx.info()
    assure isSomething(result) And result.cValue("ok"), "info", result
  
    '  // check boss key hasnt expired
    Set result = efx.validateKey(bossKey)
    assure isSomething(result) And result.cValue("ok"), "validate bosskey", result
    
    '  // get 1 writer key expiringin in 5 minutes
    Set writers = efx.generateKey(bossKey, "writer", JSONParse("{'seconds':300}"))
    assure isSomething(writers) And writers.cValue("ok") And writers.child("keys").children.Count = 1, "writerkeys", writers
    
    '  // get 2 reader keys expiringin in 5 minutes
    Set readers = efx.generateKey(bossKey, "reader", JSONParse("{'seconds':300,'count':2}"))
    assure isSomething(readers) And readers.cValue("ok") And readers.child("keys").children.Count = 2, "readerkeys", readers

    '  // get 1 updater key expiringin in 5 minutes
    Set updaters = efx.generateKey(bossKey, "updater", JSONParse("{'seconds':300}"))
    assure isSomething(updaters) And updaters.cValue("ok") And updaters.child("keys").children.Count = 1, "updaterkeys", updaters


    '  // set the keys up as default so we dont have to bother specifying them later
    efx.setKeys JSONParse("{" & _
            "'updater':'" & updaters.toString("keys.1") & _
            "','writer':'" & writers.toString("keys.1") & _
            "','reader':'" & readers.toString("keys.1") & _
        "'}")

    '  // more convenient for later
    Set keys = efx.getKeys()
    
    '  //---reading and writing
    Set someData = JSONParse("{'name':'xyz','a':[1,2,3],'b':2000}")
    
    ' // write is a reserved word in vba .. have renamed as xwrite
    Set data = efx.xwrite(someData)
    assure isSomething(data) And data.cValue("ok"), "write-post", data

    '  // read it back with the same key
    Set result = efx.read(data.toString("id"), keys.toString("writer"))
    assure isSomething(result) And result.cValue("ok") And sameOb(result.child("value"), someData), "read", result

    ' // do it again, but use a GET for writing
    Set data = efx.xwrite(someData, keys.toString("writer"), "get")
    assure isSomething(data) And data.cValue("ok"), "write-get", data

    '  // read it back with the same key
    Set result = efx.read(data.toString("id"), keys.toString("writer"))
    assure isSomething(result) And result.cValue("ok") And sameOb(result.child("value"), someData), "read-get", result

    '  // now do all that with some text data
    Dim textData As String
    textData = "some text data"

    ' // write is a reserved word in vba .. have renamed as xwrite
    Set data = efx.xwrite(textData)
    assure isSomething(data) And data.cValue("ok"), "write-post-text", data

    '  // read it back with the same key
    Set result = efx.read(data.toString("id"), keys.toString("writer"))
    assure isSomething(result) And result.cValue("ok") And result.toString("value") = textData, "read-post-text", result

    ' // do it again, but use a GET for writing
    Set data = efx.xwrite(textData, keys.toString("writer"), "get")
    assure isSomething(data) And data.cValue("ok"), "write-get-text", data

    '  // read it back with the same key
    Set result = efx.read(data.toString("id"), keys.toString("writer"))
    assure isSomething(result) And result.cValue("ok") And result.toString("value") = textData, "read-get-text", result

    '  //-- assigning readers and updaters
    Set data = efx.xwrite(someData, keys.toString("writer"), "post", _
        JSONParse("{'readers':'" + keys.toString("reader") + _
        "','updaters':'" + keys.toString("updater") + "'}"))
    assure isSomething(data) And data.cValue("ok"), "write-for-others", data

    '  // read it back with a reader key
    Set result = efx.read(data.toString("id"))
    assure isSomething(result) And result.cValue("ok") And sameOb(result.child("value"), someData), "read-reader", result
    
    '  // read it back with a reader key
    Set result = efx.read(data.toString("id"), keys.toString("updater"))
    assure isSomething(result) And result.cValue("ok") And sameOb(result.child("value"), someData), "read-updater", result
    
    '  // update it
    Set result = efx.update(textData, data.toString("id"))
    assure isSomething(result) And result.cValue("ok"), "update", result

    '  // check it took - i'll just use the reader key
    Set result = efx.read(data.toString("id"))
    assure isSomething(result) And result.cValue("ok") And result.toString("value") = textData, "read-updated", result

    '  //----work with aliases
    Set alias = efx.registerAlias(keys.toString("writer"), keys.toString("reader"), data.toString("id"), "somename")
    assure isSomething(alias) And alias.cValue("ok"), "alias", alias

    '  // read it back with a reader key, using the data alias
    Set result = efx.read(alias.toString("alias"))
    assure isSomething(result) And result.cValue("ok") And result.toString("value") = textData, "assign-alias", result

    '  // write another rec to the same alias
    Set otherData = efx.xwrite(someData, keys.toString("writer"), "post", _
        JSONParse("{'readers':'" + keys.toString("reader") + _
        "','updaters':'" + keys.toString("updater") + "'}"))
    assure isSomething(otherData) And otherData.cValue("ok"), "write-for-others-alias", otherData

    '  // assign the alias to the new data
    Set otherAlias = efx.registerAlias(keys.toString("writer"), keys.toString("reader"), otherData.toString("id"), alias.toString("alias"))
    assure isSomething(otherAlias) And otherAlias.cValue("ok") And alias.toString("alias") = otherAlias.toString("alias"), "other-alias", otherAlias
    
    ' // read it back with a reader key, using the data alias
    Set result = efx.read(alias.toString("alias"))
    assure isSomething(result) And result.cValue("ok") And sameOb(result.child("value"), someData), "read-otheralias", result

    '  // assign and update alias to an updater key
    Set updateAlias = efx.registerAlias(keys.toString("writer"), keys.toString("updater"), otherData.toString("id"), alias.toString("alias"))
    assure isSomething(updateAlias) And updateAlias.cValue("ok") And alias.toString("alias") = updateAlias.toString("alias"), "updater-alias", updateAlias
    
    '  // update it
    Set result = efx.update(textData, alias.toString("alias"))
    assure isSomething(result) And result.cValue("ok"), "update-alias-write", result

    '  // check it took - i'll just use the reader key
    Set result = efx.read(alias.toString("alias"))
    assure isSomething(result) And result.cValue("ok") And result.toString("value") = textData, "read-update-alias", result

    '  // try deleting the underlying data item using the alias.. should fail because we didnt assign an alias to the writer key
    Set result = efx.remove(alias.toString("alias"))
    assure isSomething(result) And Not result.cValue("ok") And result.cValue("code") = 404, "remove-should-fail", result


    '  // now this should work
    Set writerAlias = efx.registerAlias(keys.toString("writer"), keys.toString("writer"), otherData.toString("id"), alias.toString("alias"))
    assure isSomething(writerAlias) And writerAlias.cValue("ok") And alias.toString("alias") = writerAlias.toString("alias"), "writer-alias", writerAlias
    
    Set result = efx.remove(alias.toString("alias"))
    assure isSomething(result) And result.cValue("ok"), "remove-should-work", result

    '  // check the underlying is gone
    Set result = efx.read(alias.toString("alias"), keys.toString("writer"))
    assure isSomething(result) And Not result.cValue("ok") And result.toString("code") = 404, "read-should-fail", result
    
    ' clean up those objects
    result.tearDown
    writerAlias.tearDown
    updateAlias.tearDown
    data.tearDown
    otherData.tearDown
    alias.tearDown
    otherAlias.tearDown
    someData.tearDown

End Sub
Private Function sameOb(a As cJobject, b As cJobject) As Boolean
    '// compare cjobjects
    '// need to drop the parent as stringify keeps it
    Dim ac As cJobject, bc As cJobject
    
    '// clone them & detach them
    Set ac = a.clone().sever
    Set bc = b.clone().sever

    sameOb = JSONStringify(ac) = JSONStringify(bc)
    
    '' dont need these any more
    ac.tearDown
    bc.tearDown

End Function
Private Function assure(b As Boolean, message As String, result As cJobject) As Boolean
    
    Dim verbose As Boolean, mess As String
    '// by default
    verbose = True
    If isSomething(result) Then
        mess = JSONStringify(result)
    End If
    
    If (Not b) Then
        Debug.Print "failed:" + message + " : result : " + mess
    End If
    
    If (b And verbose) Then
        Debug.Print "passed:" + message
    End If

End Function

'
'  // check the underlying is gone
'  var result = efx.read (alias.alias, keys.writer);
'  assure ( !result.ok,"read-should-fail", result);
'
'}
'
'

'