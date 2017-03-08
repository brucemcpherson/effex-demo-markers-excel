'/**
' * basic rough driver for effex api
' * @namespace EffexApiClient
' */
Option Explicit
Private Const DEV = "https://nodestuff-xlibersion.c9users.io"
Private Const PROD = "https://ephex-auth.appspot-preview.com"
Private axOptions As cJobject
Private defaultOptions As cJobject
Private browser As cBrowser
Private pKeys As cJobject

' create should be called before using the clas
Private Function axCreate(Optional options As cJobject = Nothing) As cEffexApiClientVBA
    
    Set axOptions = options
    ' we'll need an object so create one if none given
    If (options Is Nothing) Then
        Set axOptions = New cJobject
        axOptions.Init Nothing
    End If
    ' for chaining
    Set axCreate = Me
End Function
'get something from the store
Private Function axGet(url As String) As cJobject
    Set axGet = axExec(url)
End Function
'remove something from the store
Private Function axRemove(url As String) As cJobject
    Set axRemove = axExec(url, "delete")
End Function
'post something to the store
Private Function axPost(url As String, payload As Variant) As cJobject
    Set axPost = axExec(url, "post", payload)
End Function
' general execution of requests
Private Function axExec(url As String, Optional method As String = "get", Optional payload As Variant = Empty) As cJobject
    Dim u As String, options As cJobject, body As String, ob As cJobject, p As cJobject
    
    ' build up the url with the api base
    u = axOptions.cValue("baseURL") & url

    Set options = New cJobject
    options.Init(Nothing).add "method", method
    
    ' got to play around with the payload
    body = vbNullString
    If (Not IsEmpty(payload) And method = "post") Then
        If (TypeName(payload) = "cJobject") Then
            ' if the payload is an object, then we need to stringify it after adding a data key to it
            Set p = payload
            body = "{""data"":" + JSONStringify(p) + "}"
        Else
            body = "{""data"":""" + payload + """}"
        End If
    End If
    
    If (method = "post") Then
        Debug.Assert body <> vbNullString
        browser.httpPost u, body, True
    Else
        Debug.Assert body = vbNullString
        browser.httpGET u, , , , , , , method
    End If
    
    ' dont need this any more
    options.tearDown
    
    ' in case logging is needed
    If (axOptions.cValue("verbose")) Then
        Debug.Print method, u
    End If
    
    '// do the parse, and fail over if its garbled
    Set ob = JSONParse(browser.Text, , False)
    If (ob Is Nothing) Then
        Set ob = New cJobject
        With ob.Init(Nothing)
            .add "ok", False
            .add "content", browser.Text
        End With
    End If

    '// finally add a message about failure if needed
    If (Not browser.isOk) Then
        With ob
            .add "code", browser.successCode
            .add "content", browser.Text
            .add "ok", False
        End With
    End If

    Set axExec = ob
End Function
Public Function setKeys(keys As cJobject) As cEffexApiClientVBA
    Set pKeys = keys
    Set setKeys = Me
End Function
Public Function getKeys() As cJobject
    Set getKeys = pKeys
End Function
Public Function setDev() As cEffexApiClientVBA
    Set setDev = setBase(DEV)
End Function
Public Function setProd() As cEffexApiClientVBA
    Set setProd = setBase(PROD)
End Function
Public Function setVerbose(verbose) As cEffexApiClientVBA
    axOptions.add "verbose", verbose
    Set setVerbose = Me
End Function
Public Function setBase(base) As cEffexApiClientVBA
    axOptions.add "baseURL", base
    Set setBase = Me
End Function
Public Function clone(Optional ob As cJobject) As cJobject
    If (isSomething(ob)) Then
        Set clone = JSONParse(JSONStringify(ob))
    Else
        Set clone = New cJobject
        clone.Init Nothing
    End If
End Function
'  /**
'   * turns a params object into a url
'   * @param {object} params the params
'   * @return {string} the uri
'   */
Private Function makeParams(Optional params As cJobject = Nothing) As String
    Dim c As cStringChunker, sep As String, job As cJobject, joc As cJobject, s As cStringChunker
    Set c = New cStringChunker
    Set s = New cStringChunker
    
    If (isSomething(params)) Then
        sep = "?"
        For Each job In params.children
            ' if an array, then we need to join them
            If (job.hasChildren) Then
                For Each joc In job.children
                    s.add(joc.toString).add (",")
                Next joc
                s.chopWhile (",")
            Else
                s.add job.toString
            End If
            c.add(sep).add(job.key).add("=").uri (s.toString)
            sep = "&"
        Next job
    End If
    makeParams = c.toString
End Function

'  /**
'  * @param {string} boss the boss key
'  * @param {string} mode the type like writer/reader/updater
'  * @param {object} params the params
'  * @return {Promise} to the result
'  */
Public Function generateKey(boss As String, mode As String, Optional params As cJobject) As cJobject
    Set generateKey = axGet("/" + boss + "/" + mode + makeParams(params))
End Function
'  /**
'   * ping the service
'   * @return {object} "PONG"
'   */
Public Function ping() As cJobject
    Set ping = axGet("/ping")
End Function
'  /**
'   * info the service
'   * @return {object}
'   */
Public Function info() As cJobject
    Set info = axGet("/info")
End Function
'  /**
'   * quotas the service
'   * @return {object}
'   */
Public Function getQuotas() As cJobject
    Set getQuotas = axGet("/quotas")
End Function

'
'  /**
'  * update an item
'  * @param {string} id the item id
'  * @param {string} updater the updater key
'  * @param {object} data what to write
'  * @param {string} method the to use (post,get)
'  * @param {object} params the params
'  * @return {object} to the result
'  */
Public Function update(data As Variant, Optional id As String, _
    Optional updater As String, Optional method As String = "post", Optional params As cJobject) As cJobject
    
    Dim url As String, pob As cJobject, p As cJobject
    
    '// in case we need to modify
    Set p = clone(params)
    
       '// if we are using get and have data, need to build that into the params
    If (method = "get") Then
        If (TypeName(data) = "cJobject") Then
            Set pob = data
            p.add "data", JSONStringify(pob)
        Else
            p.add "data", data
        End If
    End If
    
    method = LCase(method)
    
    ' this fills in any default keys created by setKeys
    url = "/updater/" + checkKey("updater", updater) + "/" + checkKey("item", id) + makeParams(p)
    Set update = axExec(url, method, data)
    
    ' we dont need these anymore
    If (isSomething(p)) Then
        p.tearDown
    End If
    
    If (isSomething(pob)) Then
        p.tearDown
    End If
    
End Function

'
'   /**
'  * write an item (write is reserved in VBA, so named xwrite)
'  * @param {string} writer the writer key
'  * @param {object} data what to write
'  * @param {string} method the to use (post,get)
'  * @param {object} params the params
'  * @return {Promise} to the result
'  */
Public Function xwrite(data As Variant, _
    Optional writer As String, _
    Optional method As String, _
    Optional params As cJobject) As cJobject
    
    Dim url As String, p As cJobject, pob As cJobject
    
    '// in case we need to modify
    Set p = clone(params)
    
    '// sort out the default method
    If (method = vbNullString) Then
        method = "post"
    End If
    method = LCase(method)
    
    '// if we are using get and have data, need to build that into the params
    If (method = "get") Then
        If (TypeName(data) = "cJobject") Then
            Set pob = data
            p.add "data", JSONStringify(pob)
        Else
            p.add "data", data
        End If
    End If
        
    ' this fills in any default keys created by setKeys
    url = "/writer/" + checkKey("writer", writer) + makeParams(p)
    Set xwrite = axExec(url, method, data)
    
    ' we dont need these anymore
    
    If (isSomething(p)) Then
        p.tearDown
    End If
    
    If (isSomething(pob)) Then
        p.tearDown
    End If
    
    End Function

'  /**
'  * remove
'  * @param {string} id the item id
'  * @param {string} writer the writer key
'  * @param {object} params the params
'  * @return {Promise} to the result
'  */
Public Function remove(id As String, _
    Optional writer As String, Optional params As cJobject) As cJobject
    
    Set remove = axRemove("/writer/" + checkKey("writer", writer) + "/" + checkKey("item", id) + makeParams(params))
End Function
Public Function read(id As String, _
    Optional reader As String, Optional params As cJobject) As cJobject
    
    Set read = axGet("/reader/" + checkKey("reader", reader) + "/" + checkKey("item", id) + makeParams(params))
End Function
'
'  /**
'  * @param {string} coupon the coupon code
'  * @return {Promise} to the result
'  */
Public Function validateKey(coupon As String) As cJobject
    Set validateKey = axGet("/validate/" + coupon)
End Function
'  /**
'  * @param {string} id the item id
'  * @param {string} writer the writer key
'  * @param {string} key the key to assign the alias for
'  * @param {string} alias the alias to assign
'  * @param {object} params the params
'  * @return {Promise} to the result
'  */
Public Function registerAlias(writer As String, key As String, id As String, alias As String, _
    Optional params As cJobject) As cJobject
    
    Set registerAlias = axGet("/" + checkKey("writer", writer) + "/" _
        + key + "/alias/" + URLEncode(alias) + "/" + checkKey("item", id) + makeParams(params))
End Function
Private Function checkKey(keyType As String, value As String) As String
    Dim k As String
    k = value
    If (k = vbNullString) Then
        k = pKeys.toString(keyType)
    End If

    Debug.Assert k <> vbNullString
    checkKey = k
    
End Function


Private Sub Class_Initialize()
    'we will need this later
    Set browser = New cBrowser
    Set defaultOptions = New cJobject
    With defaultOptions.Init(Nothing)
        .add "baseURL", PROD
        .add "verbose", False
    End With
    axCreate defaultOptions

End Sub

Private Sub Class_Terminate()
    ' recover memory from double linked objects
    If isSomething(browser) Then
        browser.tearDown
    End If
    If isSomething(defaultOptions) Then
        defaultOptions.tearDown
    End If
End Sub