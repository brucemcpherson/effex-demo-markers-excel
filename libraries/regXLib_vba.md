# VBA Project: **effex-demo-markers-excel**
## VBA Module: **[regXLib](/libraries/regXLib.vba "source is here")**
### Type: StdModule  

This procedure list for repo (effex-demo-markers-excel) was automatically created on 3/8/2017 8:42:35 AM by VBAGit.
For more information see the [desktop liberation site](http://ramblings.mcpher.com/Home/excelquirks/drivesdk/gettinggithubready "desktop liberation")

Below is a section for each procedure in regXLib

---
VBA Procedure: **rxString**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function rxString(sname As String, s As String, Optional ignorecase As Boolean = True) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
sname|String|False||
s|String|False||
ignorecase|Boolean|True| True|


---
VBA Procedure: **rxGroup**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function rxGroup(sname As String, s As String, group As Long, Optional ignorecase As Boolean = True) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
sname|String|False||
s|String|False||
group|Long|False||
ignorecase|Boolean|True| True|


---
VBA Procedure: **rxTest**  
Type: **Function**  
Returns: **Boolean**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function rxTest(sname As String, s As String, Optional ignorecase As Boolean = True) As Boolean*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
sname|String|False||
s|String|False||
ignorecase|Boolean|True| True|


---
VBA Procedure: **rxReplace**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function rxReplace(sname As String, sFrom As String, sTo As String, Optional ignorecase As Boolean = True) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
sname|String|False||
sFrom|String|False||
sTo|String|False||
ignorecase|Boolean|True| True|


---
VBA Procedure: **rxPattern**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function rxPattern(sname As String) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
sname|String|False||


---
VBA Procedure: **rxMakeRxLib**  
Type: **Function**  
Returns: **[cregXLib](/libraries/cregXLib_cls.md "cregXLib")**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function rxMakeRxLib(sname As String) As cregXLib*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
sname|String|False||
