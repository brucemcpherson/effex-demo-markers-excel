# VBA Project: **effex-demo-markers-excel**
## VBA Module: **[cEffexApiClientVBA](/libraries/cEffexApiClientVBA.cls "source is here")**
### Type: ClassModule  

This procedure list for repo (effex-demo-markers-excel) was automatically created on 3/8/2017 8:42:33 AM by VBAGit.
For more information see the [desktop liberation site](http://ramblings.mcpher.com/Home/excelquirks/drivesdk/gettinggithubready "desktop liberation")

Below is a section for each procedure in cEffexApiClientVBA

---
VBA Procedure: **axCreate**  
Type: **Function**  
Returns: **[cEffexApiClientVBA](/libraries/cEffexApiClientVBA_cls.md "cEffexApiClientVBA")**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function axCreate(Optional options As cJobject = Nothing) As cEffexApiClientVBA*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
options|[cJobject](/libraries/cJobject_cls.md "cJobject")|True| Nothing|


---
VBA Procedure: **axGet**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function axGet(url As String) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
url|String|False||


---
VBA Procedure: **axRemove**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function axRemove(url As String) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
url|String|False||


---
VBA Procedure: **axPost**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function axPost(url As String, payload As Variant) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
url|String|False||
payload|Variant|False||


---
VBA Procedure: **axExec**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function axExec(url As String, Optional method As String = "get", Optional payload As Variant = Empty) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
url|String|False||
method|String|True| "get"|
payload|Variant|True| Empty|


---
VBA Procedure: **setKeys**  
Type: **Function**  
Returns: **[cEffexApiClientVBA](/libraries/cEffexApiClientVBA_cls.md "cEffexApiClientVBA")**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function setKeys(keys As cJobject) As cEffexApiClientVBA*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
keys|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||


---
VBA Procedure: **getKeys**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function getKeys() As cJobject*  

**no arguments required for this procedure**


---
VBA Procedure: **setDev**  
Type: **Function**  
Returns: **[cEffexApiClientVBA](/libraries/cEffexApiClientVBA_cls.md "cEffexApiClientVBA")**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function setDev() As cEffexApiClientVBA*  

**no arguments required for this procedure**


---
VBA Procedure: **setProd**  
Type: **Function**  
Returns: **[cEffexApiClientVBA](/libraries/cEffexApiClientVBA_cls.md "cEffexApiClientVBA")**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function setProd() As cEffexApiClientVBA*  

**no arguments required for this procedure**


---
VBA Procedure: **setVerbose**  
Type: **Function**  
Returns: **[cEffexApiClientVBA](/libraries/cEffexApiClientVBA_cls.md "cEffexApiClientVBA")**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function setVerbose(verbose) As cEffexApiClientVBA*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
verbose|Variant|False||


---
VBA Procedure: **setBase**  
Type: **Function**  
Returns: **[cEffexApiClientVBA](/libraries/cEffexApiClientVBA_cls.md "cEffexApiClientVBA")**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function setBase(base) As cEffexApiClientVBA*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
base|Variant|False||


---
VBA Procedure: **clone**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function clone(Optional ob As cJobject) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
ob|[cJobject](/libraries/cJobject_cls.md "cJobject")|True||


---
VBA Procedure: **makeParams**  
Type: **Function**  
Returns: **String**  
Return description: **the uri**  
Scope: **Private**  
Description: **turns a params object into a url**  

*Private Function makeParams(Optional params As cJobject = Nothing) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
params|[cJobject](/libraries/cJobject_cls.md "cJobject")|True| Nothing|the params


---
VBA Procedure: **generateKey**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: **to the result**  
Scope: **Public**  
Description: **@param {string} boss the boss key**  

*Public Function generateKey(boss As String, mode As String, Optional params As cJobject) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
boss|String|False||the boss key
mode|String|False||the type like writer/reader/updater
params|[cJobject](/libraries/cJobject_cls.md "cJobject")|True||the params


---
VBA Procedure: **ping**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: **"PONG"**  
Scope: **Public**  
Description: **ping the service**  

*Public Function ping() As cJobject*  

**no arguments required for this procedure**


---
VBA Procedure: **info**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: **'   */**  
Scope: **Public**  
Description: **info the service**  

*Public Function info() As cJobject*  

**no arguments required for this procedure**


---
VBA Procedure: **getQuotas**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: **'   */**  
Scope: **Public**  
Description: **quotas the service**  

*Public Function getQuotas() As cJobject*  

**no arguments required for this procedure**


---
VBA Procedure: **update**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: **to the result**  
Scope: **Public**  
Description: **update an item**  

*Public Function update(data As Variant, Optional id As String, Optional updater As String, Optional method As String = "post", Optional params As cJobject) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
data|Variant|False||what to write
id|String|True||the item id
updater|String|True||the updater key
method|String|True| "post"|the to use (post,get)
params|[cJobject](/libraries/cJobject_cls.md "cJobject")|True||the params


---
VBA Procedure: **xwrite**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: **to the result**  
Scope: **Public**  
Description: **write an item (write is reserved in VBA, so named xwrite)**  

*Public Function xwrite(data As Variant, Optional writer As String, Optional method As String, Optional params As cJobject) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
data|Variant|False||what to write
writer|String|True||the writer key
method|String|True||the to use (post,get)
params|[cJobject](/libraries/cJobject_cls.md "cJobject")|True||the params


---
VBA Procedure: **remove**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: **to the result**  
Scope: **Public**  
Description: **remove**  

*Public Function remove(id As String, Optional writer As String, Optional params As cJobject) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
id|String|False||the item id
writer|String|True||the writer key
params|[cJobject](/libraries/cJobject_cls.md "cJobject")|True||the params


---
VBA Procedure: **read**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function read(id As String, Optional reader As String, Optional params As cJobject) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
id|String|False||
reader|String|True||
params|[cJobject](/libraries/cJobject_cls.md "cJobject")|True||


---
VBA Procedure: **validateKey**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: **to the result**  
Scope: **Public**  
Description: **@param {string} coupon the coupon code**  

*Public Function validateKey(coupon As String) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
coupon|String|False||the coupon code


---
VBA Procedure: **registerAlias**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: **to the result**  
Scope: **Public**  
Description: **@param {string} id the item id**  

*Public Function registerAlias(writer As String, key As String, id As String, alias As String, Optional params As cJobject) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
writer|String|False||the writer key
key|String|False||the key to assign the alias for
id|String|False||the item id
alias|String|False||the alias to assign
params|[cJobject](/libraries/cJobject_cls.md "cJobject")|True||the params


---
VBA Procedure: **checkKey**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function checkKey(keyType As String, value As String) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
keyType|String|False||
value|String|False||


---
VBA Procedure: **Class_Initialize**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Sub Class_Initialize()*  

**no arguments required for this procedure**


---
VBA Procedure: **Class_Terminate**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Sub Class_Terminate()*  

**no arguments required for this procedure**
