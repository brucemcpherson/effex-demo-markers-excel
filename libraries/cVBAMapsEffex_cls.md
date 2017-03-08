# VBA Project: **effex-demo-markers-excel**
## VBA Module: **[cVBAMapsEffex](/libraries/cVBAMapsEffex.cls "source is here")**
### Type: ClassModule  

This procedure list for repo (effex-demo-markers-excel) was automatically created on 3/8/2017 8:40:45 AM by VBAGit.
For more information see the [desktop liberation site](http://ramblings.mcpher.com/Home/excelquirks/drivesdk/gettinggithubready "desktop liberation")

Below is a section for each procedure in cVBAMapsEffex

---
VBA Procedure: **getProperty**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function getProperty(key As String) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
key|String|False||


---
VBA Procedure: **setProperty**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function setProperty(key As String, value As cJobject) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
key|String|False||
value|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||


---
VBA Procedure: **makeKeys**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: **containg {keySet:{}, results:{}}**  
Scope: **Public**  
Description: **make a set of keys that can be stored to props service**  

*Public Function makeKeys(boss As String, Optional params As cJobject) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
boss|String|False||the boss key to use
params|[cJobject](/libraries/cJobject_cls.md "cJobject")|True||


---
VBA Procedure: **pullValues**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Public**  
Description: **pull data from sheet and change it back to sheet format**  

*Public Function pullValues(key As String, itemId As String) As Variant*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
key|String|False||
itemId|String|False||the item id or alias


---
VBA Procedure: **unObjectify**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Public**  
Description: **transform  rows of objects to values**  

*Public Function unObjectify(data As cJobject) As Variant*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
data|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||


---
VBA Procedure: **pullFromEffex**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: **the final pull result**  
Scope: **Public**  
Description: **create a data item on effex for updating**  

*Public Function pullFromEffex(key As String, id As String) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
key|String|False||the key
id|String|False||an alias or id to use


---
VBA Procedure: **pushDataForUpdate**  
Type: **Function**  
Returns: **Variant**  
Return description: **write results**  
Scope: **Public**  
Description: **push to effex**  

*Public Function pushDataForUpdate(data As cJobject, keys As cJobject, Optional alias As String, Optional params As cJobject)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
data|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||the object to write
keys|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||a set of keys
alias|String|True||
params|[cJobject](/libraries/cJobject_cls.md "cJobject")|True||


---
VBA Procedure: **objectify**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: ****  
Scope: **Public**  
Description: **objectify a sheet**  

*Public Function objectify(values As Variant) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
values|Variant|False||


---
VBA Procedure: **Class_Initialize**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Sub Class_Initialize()*  

**no arguments required for this procedure**
