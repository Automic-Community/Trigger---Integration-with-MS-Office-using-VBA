*Trigger - Integration with MS Office using VBA*
=============


We have created some VBA libraries (.bas and .cls) in order to ease the use of DUAS triggers for any external system development using VBA.
http://github.com/Automic-Community/Trigger---Integration-with-MS-Office-using-VBA

<!-- List of attached files -->
Contents of Solution Package:

						
								*Dollar_Universe-VBA_Trigger.zip
								
						


Documenation and Instructions
---

<div class="ipsType_textblock ipsPad_half description_content"><strong class="bbc"><span>Introduction</span></strong><br /><span><span><span>We have created some VBA libraries (.bas and .cls) in order to ease the use of DUAS triggers for any external system development using VBA.</span></span></span><br /><br /><strong class="bbc"><span><span><span>Prerequisites</span></span></span></strong><br /><span><span>In Office 2010, you will need to add the "Microsoft Scripting Runtime" and "Microsoft ActiveX Data Objects 2.8 library" as references via Tools &gt; References in the VBA editor.</span></span><br /><span><span><span>You will also need to </span></span></span><strong class="bbc">import those 2 external librarie</strong><span><span><span>s:</span></span></span>
<ul class="bbc">
<li><span><strong class="bbc">JSON</strong>.bas (from the <a class="bbc_url" title="External link" href="http://www.ediy.co.nz/vbjson-json-parser-library-in-vb6-xidc55680.html" rel="nofollow external">VBJson</a> library)</span></li>
<li><span><strong class="bbc">cStringBuilder</strong>.cls (from the <a class="bbc_url" title="External link" href="http://www.ediy.co.nz/vbjson-json-parser-library-in-vb6-xidc55680.html" rel="nofollow external">VBJson</a> library)</span></li>
</ul>
<strong class="bbc"><span>Quick Start</span></strong><br /><span>Use the full example file to try this example with Excel. </span><br /><span>If you need to include the code in your existing VBA project you can i<span><span>mport these modules:</span></span></span>
<ul class="bbc">
<li><span><strong class="bbc">DUASApi.bas</strong></span></li>
<li><span><strong class="bbc">DUASApiParam.cls</strong></span></li>
<li><span><strong class="bbc">DUASApiRespObj.cls</strong></span></li>
<li><span><strong class="bbc">DUASApiRespObjElmt.cls</strong></span></li>
</ul>
<strong class="bbc"><span>Description</span></strong><br /><span><span><span>Here is a sample of code showing how to use the libraries. </span></span></span><br /><span><span><span>In the example below :</span></span></span>
<ul class="bbc">
<li><span><span><span>We first login to the Dollar Universe API and retrieve the token</span></span></span></li>
<li><span><span><span>We send an event</span></span></span></li>
<li><span><span><span>We log out</span></span></span></li>
</ul>
<pre class="prettyprint lang-auto linenums:0 prettyprinted"><span class="typ">Sub</span><span class="pln"> </span><span class="typ">Button1_Click</span><span class="pun">()</span><span class="pln">
</span><span class="pun">...</span><span class="pln">
</span><span class="str">' Login to the DUAS API
Set resp = DUASApi.loginDUASApi(host, port, login, pwd)
DUASApi.displayResponseInMsgBox resp
If resp.status = "success" Then

token = resp.token

'</span><span class="pln"> </span><span class="typ">Triggers</span><span class="pln"> an </span><span class="kwd">event</span><span class="pln"> to DUAS
</span><span class="typ">Set</span><span class="pln"> resp </span><span class="pun">=</span><span class="pln"> </span><span class="typ">DUASApi</span><span class="pun">.</span><span class="pln">triggerEventDUASApi</span><span class="pun">(</span><span class="pln">host</span><span class="pun">,</span><span class="pln"> port</span><span class="pun">,</span><span class="pln"> token</span><span class="pun">,</span><span class="pln"> eventType</span><span class="pun">,</span><span class="pln"> </span><span class="kwd">params</span><span class="pun">)</span><span class="pln">
</span><span class="typ">DUASApi</span><span class="pun">.</span><span class="pln">displayResponseInMsgBox resp

</span><span class="str">' Logout from DUAS API
Set resp = DUASApi.logoutDUASApi(host, port, token)
DUASApi.displayResponseInMsgBox resp

End If
...
End Sub</span></pre>
<br /><span><strong class="bbc">List of functions</strong></span><br /><br /><span><span><span>Here's a list of all the functions from <strong class="bbc">DUASApi.bas</strong> that you can use:</span></span></span><br /><span><span><span><span><span><span> <br /><br /></span></span></span></span></span></span>
<pre class="prettyprint lang-auto linenums:0 prettyprinted"><span class="str">' Logins to the Dollar Universe API with a HTTP request
'</span><span class="pln">
</span><span class="str">' Parameters
'</span><span class="pln"> host </span><span class="typ">The</span><span class="pln"> hostname of your DUAS server
</span><span class="str">' port The port number of your DUAS server
'</span><span class="pln"> login login to connect to the </span><span class="typ">Dollar</span><span class="pln"> </span><span class="typ">Universe</span><span class="pln"> API
</span><span class="str">' pwd password to connect to the Dollar Universe API
'</span><span class="pln">
</span><span class="str">' Returns: a structure containing all the information of the Dollar Universe API response
'</span><span class="pln">
</span><span class="typ">Public</span><span class="pln"> </span><span class="typ">Function</span><span class="pln"> loginDUASApi</span><span class="pun">(</span><span class="typ">ByVal</span><span class="pln"> host </span><span class="typ">As</span><span class="pln"> </span><span class="typ">String</span><span class="pun">,</span><span class="pln"> </span><span class="typ">ByVal</span><span class="pln"> port </span><span class="typ">As</span><span class="pln"> </span><span class="typ">String</span><span class="pun">,</span><span class="pln"> _
</span><span class="typ">ByVal</span><span class="pln"> login </span><span class="typ">As</span><span class="pln"> </span><span class="typ">String</span><span class="pun">,</span><span class="pln"> </span><span class="typ">ByVal</span><span class="pln"> pwd </span><span class="typ">As</span><span class="pln"> </span><span class="typ">String</span><span class="pun">)</span><span class="pln"> </span><span class="typ">As</span><span class="pln"> </span><span class="typ">DUASApiRespObj</span></pre>
<br /><span><span><span><span><span> </span></span></span></span></span><br /><span><span><span>Logs you in to the Dollar Universe API and returns a DUASApiRespObj which contains the return status and the <strong class="bbc">authentication token</strong> if status is successfull</span></span></span><span><span><span><span> <br /><br /></span></span></span></span>
<pre class="prettyprint lang-auto linenums:0 prettyprinted"><span class="str">' Triggers an event on DUAS via the Dollar Universe API with an HTTP request
'</span><span class="pln">
</span><span class="str">' Parameters
'</span><span class="pln"> host </span><span class="typ">The</span><span class="pln"> hostname of your DUAS server
</span><span class="str">' port The port number of your DUAS server
'</span><span class="pln"> token </span><span class="typ">The</span><span class="pln"> authentication token of your </span><span class="typ">Dollar</span><span class="pln"> </span><span class="typ">Universe</span><span class="pln"> API session
</span><span class="str">' eventType The event type defined in your trigger object in DUAS
'</span><span class="pln"> </span><span class="kwd">params</span><span class="pln"> </span><span class="typ">The</span><span class="pln"> list of parameters </span><span class="kwd">defined</span><span class="pln"> </span><span class="kwd">in</span><span class="pln"> your trigger </span><span class="kwd">object</span><span class="pln"> </span><span class="kwd">in</span><span class="pln"> DUAS
</span><span class="str">'
'</span><span class="pln"> </span><span class="typ">Returns</span><span class="pun">:</span><span class="pln"> a structure containing all the information of the </span><span class="typ">Dollar</span><span class="pln"> </span><span class="typ">Universe</span><span class="pln"> API response
</span><span class="str">'
Public Function triggerEventDUASApi(ByVal host As String, ByVal port As String, _
ByVal token As String, ByVal eventType As String, ByVal params As Collection) As DUASApiRespObj</span></pre>
<br /><span><span><span>Sends an event to the Dollar Universe API: you need to provide the authentication token, the event type, and a list of optional parameters defined in your trigger.</span></span></span><br /><span><span></span></span><br />
<pre class="prettyprint lang-auto linenums:0 prettyprinted"><span class="str">'
'</span><span class="pln"> </span><span class="typ">Logout</span><span class="pln"> </span><span class="kwd">from</span><span class="pln"> the </span><span class="typ">Dollar</span><span class="pln"> </span><span class="typ">Universe</span><span class="pln"> API </span><span class="kwd">with</span><span class="pln"> an HTTP request
</span><span class="str">'
'</span><span class="pln"> </span><span class="typ">Parameters</span><span class="pln">
</span><span class="str">' host The hostname of your DUAS server
'</span><span class="pln"> port </span><span class="typ">The</span><span class="pln"> port number of your DUAS server
</span><span class="str">' token The authentication token of your Dollar Universe API session
'</span><span class="pln">
</span><span class="str">' Returns: a structure containing all the information of the Dollar Universe API response
'</span><span class="pln">
</span><span class="typ">Public</span><span class="pln"> </span><span class="typ">Function</span><span class="pln"> logoutDUASApi</span><span class="pun">(</span><span class="typ">ByVal</span><span class="pln"> host </span><span class="typ">As</span><span class="pln"> </span><span class="typ">String</span><span class="pun">,</span><span class="pln"> </span><span class="typ">ByVal</span><span class="pln"> port </span><span class="typ">As</span><span class="pln"> </span><span class="typ">String</span><span class="pun">,</span><span class="pln"> </span><span class="typ">ByVal</span><span class="pln"> token </span><span class="typ">As</span><span class="pln"> </span><span class="typ">String</span><span class="pun">)</span><span class="pln"> </span><span class="typ">As</span><span class="pln"> </span><span class="typ">DUASApiRespObj</span></pre>
<br /><span><span><span> </span></span></span><br /><span><span><span>Logs you out of the Dollar Unvierse API; it basically revokes the authentication token that you give as parameter</span></span></span><br /><span><span><span><span><span><span> <br /><br /><span></span></span></span></span></span></span></span>
<pre class="prettyprint lang-auto linenums:0 prettyprinted"><span class="str">'
'</span><span class="pln"> </span><span class="typ">Displays</span><span class="pln"> </span><span class="kwd">in</span><span class="pln"> a </span><span class="typ">MsgBox</span><span class="pln"> the structure containing all the information of an HTTP response </span><span class="kwd">from</span><span class="pln"> the </span><span class="typ">Dollar</span><span class="pln"> </span><span class="typ">Unvierse</span><span class="pln"> API
</span><span class="str">'
'</span><span class="pln"> </span><span class="typ">Parameters</span><span class="pln">
</span><span class="str">' resp structure containing all the information of the Dollar Universe API response
'</span><span class="pln">
</span><span class="typ">Public</span><span class="pln"> </span><span class="typ">Sub</span><span class="pln"> displayResponseInMsgBox</span><span class="pun">(</span><span class="typ">ByVal</span><span class="pln"> resp </span><span class="typ">As</span><span class="pln"> </span><span class="typ">DUASApiRespObj</span><span class="pun">)</span></pre>
<br /><span><span><span><span><span> </span></span></span></span></span><br /><span><span><span>Displays as a MsgBox a response object from the Dollar Universe API</span></span></span><br /><br /><span><span><span><span> <br /></span></span></span></span>
<pre class="prettyprint lang-auto linenums:0 prettyprinted"><span class="str">'
'</span><span class="pln"> </span><span class="typ">Displays</span><span class="pln"> </span><span class="kwd">and</span><span class="pln"> appends </span><span class="kwd">in</span><span class="pln"> an </span><span class="typ">Excel</span><span class="pln"> </span><span class="typ">Cell</span><span class="pln"> the structure containing all the information of an HTTP response </span><span class="kwd">from</span><span class="pln"> the </span><span class="typ">Dollar</span><span class="pln"> </span><span class="typ">Unvierse</span><span class="pln"> API
</span><span class="str">'
'</span><span class="pln"> </span><span class="typ">Parameters</span><span class="pln">
</span><span class="str">' resp structure containing all the information of the Dollar Universe API response
'</span><span class="pln"> </span><span class="typ">WorksheetName</span><span class="pln"> </span><span class="typ">The</span><span class="pln"> </span><span class="typ">WorkSheet</span><span class="pln"> name of your </span><span class="typ">Excel</span><span class="pln"> file
</span><span class="str">' row The row number of the Excel Cell in which you want to display the information
'</span><span class="pln"> col </span><span class="typ">The</span><span class="pln"> column number of the </span><span class="typ">Excel</span><span class="pln"> </span><span class="typ">Cell</span><span class="pln"> </span><span class="kwd">in</span><span class="pln"> which you want to display the information
</span><span class="str">'
Public Sub displayResponseAppendInCell(ByVal resp As DUASApiRespObj, _
ByVal WorksheetName As String, ByVal row As Integer, ByVal col As Integer)</span></pre>
<br /><br /><span><span><span>Displays and appends in an Excel Cell a response object from the Dollar Universe API</span></span></span><br /><br /><br /><strong class="bbc"><span>Licences</span></strong><br /><span><span><span>We rely on an open library under BSD license called </span></span></span><a class="bbc_url" title="External link" href="http://www.ediy.co.nz/vbjson-json-parser-library-in-vb6-xidc55680.html" rel="nofollow external">VBJson</a><span><span><span> used to parse the JSON responses from DUAS.</span></span></span></div>

Copyright and License
---

Broadcom does not support, maintain or warrant Solutions, Templates, Actions and any other content published on the Community and is subject to Broadcom Community [Terms and Conditions](https://community.broadcom.com/termsandconditions)


Questions or Need Help? 
---
Join the [Automic Community Integrations](https://community.broadcom.com/communities/community-home?CommunityKey=83e49dd4-b93e-464a-a343-2bb1e51c13ec) to discuss this integration.
