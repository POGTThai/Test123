<% 
    Option Explicit
	Dim mstrErr, mstrVar, mstrParam, key 
	Dim objRecord, objLookUp, dsaCheqDet1, dsaCheqDet2
	
%>
<!--#include file="../subroutine/conn.asp"-->
<!--#include file="../Shared/callsub.asp"-->
<SCRIPT language=JavaScript src="../Shared/callWindow.js"></SCRIPT>
<SCRIPT language=JavaScript src="../Shared/winOpener.js"></SCRIPT>
<%
	set objRecord=Server.CreateObject("USPConnect.clsHTML")	
	objRecord.mstrRPC = "SP.FRMARAPENQUIRYEN.ASP"
	objRecord.Initialize gobjDB    
	
	Set objLookUp = server.CreateObject("USPConnect.clsHTML")
	objLookUp.mstrRPC = objRecord.mstrRPC
	objLookUp.Initialize gobjdb
	
	Set dsaCheqDet1 = server.CreateObject("VPDsa20.DynamicArray")
	dsaCheqDet1.Create "", chr(255), chr(254), chr(253), chr(252)
	
	Set dsaCheqDet2 = server.CreateObject("VPDsa20.DynamicArray")
	dsaCheqDet2.Create "", chr(255), chr(254), chr(253), chr(252)
	
    gobjGeneral.DSTR(mstrVar,21) = objPage.ReadData("USERID")
    gobjGeneral.DSTR(mstrVar,22) = objPage.ReadData("GUID")
%>

<HTML>
<HEAD>
  <TITLE>THI_ARAP</TITLE>
  <LINK HREF="../css/style.css" REL="stylesheet" STYLE="text/css">
</HEAD>
<SCRIPT language="JavaScript" src="../Scripts/General.js"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript">
function checkInput()
{
	return true;
}
</SCRIPT>
<BODY>
<IFRAME ID="__hifSmartNav" NAME="__hifSmartNav" STYLE="display:none"></IFRAME>			
<FORM NAME="FORM" METHOD="POST" ACTION="ACENQCHQ001-1.ASP" __smartNavEnabled="true">
<%    
    objPage.StartForm  ' Call at beginning of form 
        
    If Not objPage.isPostBack Then
%>
<!-- #include file="../Shared/StdInc1-0.asp" -->
<%
		If Request.QueryString("__FROM") = "PREV" Then
			Key = objPage.ReadData("Key")
			CallSub objRecord, "18", mstrErr, mstrVar, mstrParam, Key
			If Not objPage.HandleUSPErrorAlert(mstrErr,"FORM") Then
				Response.Write "<SCRIPT language=JavaScript>"
				Response.Write "history.back()"
				Response.Write "</SCRIPT>"			
			End If
		Else
			objRecord.ArrayValue.Value(1,44) = "1000"
			objRecord.ArrayValue.Value(1,43) = "R"
			objRecord.ArrayValue.Value(1,63) = "N"
			objRecord.ArrayValue.Value(1,901) = "0"
		End If
		objPage.Record = objRecord.ArrayValue.Value
		'Get Page Caption
		objLookUp.ArrayValue.Value(1,1) = "ACENQCHQ001-1.ASP"
		objLookUp.ArrayValue.Value(1,3) = "CHEQUE.TYPE" & CHR(253) & "CHEQUE.STATUS"
		objLookUp.ArrayValue.Value(1,4) = "" & CHR(253) & ""
		objLookUp.ArrayValue.Value(1,5) = "CODE" & CHR(252) & "DESC" & CHR(253) & "CODE" & CHR(252) & "DESC"
%>		
<!-- #include file="../Shared/StdInc1-1.asp" -->
<%
		objLookUp.ArrayValue.Value(1,5,1) = "R" & chr(252) & "P"
		objLookUp.ArrayValue.Value(1,5,2) = objLookUp.ArrayValue.Value(1,1,26) & chr(252) & objLookUp.ArrayValue.Value(1,1,27)
		objLookUp.ArrayValue.Value(1,6,1) = "N" & chr(252) & "A"
		objLookUp.ArrayValue.Value(1,6,2) = objLookUp.ArrayValue.Value(1,1,28) & chr(252) & objLookUp.ArrayValue.Value(1,1,29)
		objPage.LookupData = objLookUp.ArrayValue.Value
		
		With dsaCheqDet1
			.Value (1,3) = "100%"  ' Width
		    .Value (1,4) = "1"	' Hides Select All button
		    .Value (1,5) = "1"	' Hides Select All button
		    .Value (1,6) = "1"	' Hides Disable Insert button
		    .Value (1,7) = "1"  ' Disable Delete
		    .Value (1,8) = "0"
		    .Value (1,9) = "1"  ' Disable Add
		    
		    ' Attributes
		    .Value (2,1) = "64"
		    .Value (2,2) = "65"
		    .Value (2,3) = "66"
		    .Value (2,4) = "67"
		    .Value (2,5) = "68"
		    .Value (2,6) = "69"
		    .Value (2,7) = "70"
		    .Value (2,8) = "71"
		    .Value (2,9) = "72"
		    .Value (2,10) = "73"
		    .Value (2,11) = "813"
		    
		    ' Data Types 1 - Numeric, 2 - Date
		    .Value (3,4,1) = "2"
		    .Value (3,4,6) = "2"
		    
		    ' Heading
		    .Value (4,1) = objLookUp.ArrayValue.Value(1,1,30)
		    .Value (4,2) = objLookUp.ArrayValue.Value(1,1,31)
			.Value (4,3) = objLookUp.ArrayValue.Value(1,1,32)
			.Value (4,4) = objLookUp.ArrayValue.Value(1,1,33)
			.Value (4,5) = objLookUp.ArrayValue.Value(1,1,35)
		    .Value (4,6) = objLookUp.ArrayValue.Value(1,1,36)
		    .Value (4,7) = objLookUp.ArrayValue.Value(1,1,37)
			.Value (4,8) = objLookUp.ArrayValue.Value(1,1,38)
			.Value (4,9) = objLookUp.ArrayValue.Value(1,1,39)
			.Value (4,10) = objLookUp.ArrayValue.Value(1,1,40)
			.Value (4,11) = objLookUp.ArrayValue.Value(1,1,34)     
		    ' Cell Width
		    .Value (5,1) = "10%"
		    .Value (5,2) = "10%"
		    .Value (5,3) = "12.5%"
		    .Value (5,4) = "8%"
		    .Value (5,5) = "5%"
		    .Value (5,6) = "5%"
		    .Value (5,7) = "10%"
		    .Value (5,8) = "10%"
		    .Value (5,9) = "12.5%"
		    .Value (5,10) = "8%"
		    .Value (5,11) = "5%"
		    
		    ' Obj Type
		    .Value (6,11) = "6"
		    
		    ' Alignment
		    .Value (9,1) = "1"
		    .Value (9,2) = "1"
		    .Value (9,3) = "1"
		    .Value (9,4) = "2"
		    .Value (9,5) = "1"
		    .Value (9,6) = "1"
		    .Value (9,7) = "1"
		    .Value (9,8) = "2"
		    .Value (9,9) = "1"
		    .Value (9,10) = "2"
		    .Value (9,11) = "3"

			.Value (14,1) = "1"
			.Value (14,2) = "1"
			.Value (14,3) = "1"
			.Value (14,4) = "1"
			.Value (14,5) = "1"
			.Value (14,6) = "1"
			.Value (14,7) = "1"
			.Value (14,8) = "1"
			.Value (14,9) = "1"
			.Value (14,10) = "1"
			.Value (14,11) = "1"

		End With
		
	End If
	'Refresh
	If objPage.IsPostBack Then
		objLookUp.ArrayValue.Value = objPage.LookupData
		objPage.TBLCapture dsaCheqDet1.Value, "CheqDet1"
		'objPage.TBLCapture dsaCheqDet2.Value, "CheqDet2"
		objRecord.ArrayValue.Value = objPage.Record 
				
		objRecord.ArrayValue.Value(1,43) = Request.Form("optType")
		objRecord.ArrayValue.Value(1,44) = Request.Form("txtMaxQuery")
		objRecord.ArrayValue.Value(1,45) = Request.Form("txtChqBankFr")
		objRecord.ArrayValue.Value(1,46) = Request.Form("txtChqBankTo")
		objRecord.ArrayValue.Value(1,47) = Request.Form("txtChqBrnFr")
		objRecord.ArrayValue.Value(1,48) = Request.Form("txtChqBrnTo")
		objRecord.ArrayValue.Value(1,49) = Request.Form("txtChqNbrFr")
		objRecord.ArrayValue.Value(1,50) = Request.Form("txtChqNbrTo")
		objRecord.ArrayValue.Value(1,51) = Request.Form("txtChqDateFr")
		objRecord.ArrayValue.Value(1,52) = Request.Form("txtChqDateTo")
		objRecord.ArrayValue.Value(1,53) = Request.Form("txtClrDateFr")
		objRecord.ArrayValue.Value(1,54) = Request.Form("txtClrDateTo")
		objRecord.ArrayValue.Value(1,55) = Request.Form("cbxChqTypeFr")
		objRecord.ArrayValue.Value(1,56) = Request.Form("cbxChqTypeTo")
		objRecord.ArrayValue.Value(1,57) = Request.Form("cbxChqSttFr")
		objRecord.ArrayValue.Value(1,58) = Request.Form("cbxChqSttTo")
		objRecord.ArrayValue.Value(1,59) = Request.Form("txtPayInFr")
		objRecord.ArrayValue.Value(1,60) = Request.Form("txtPayInTo")
		objRecord.ArrayValue.Value(1,61) = Request.Form("txtVoucDateFr")
		objRecord.ArrayValue.Value(1,62) = Request.Form("txtVoucDateTo")
		objRecord.ArrayValue.Value(1,63) = Request.Form("optPayIn")
		objRecord.ArrayValue.Value(1,814) = Request.Form("txtChqBankFrDesc1")
		objRecord.ArrayValue.Value(1,815) = Request.Form("txtChqBankToDesc1")
		objRecord.ArrayValue.Value(1,816) = Request.Form("txtChqBrnFrDesc1")
		objRecord.ArrayValue.Value(1,817) = Request.Form("txtChqBrnToDesc1")
		objRecord.ArrayValue.Value(1,818) = Request.Form("txtPayInFrDesc1")
		objRecord.ArrayValue.Value(1,819) = Request.Form("txtPayInToDesc1")
		
		objPage.Record = objRecord.ArrayValue.Value
	    objRecord.ArrayValue.Value = objPage.Record
	    
	    Select Case objPage.GetAction
			Case "Query"
				Dim li 
				objRecord.ArrayValue.Value(1,813) = ""
				For li = 64 to 73  ' This field number must be correspond with data column display
					objRecord.ArrayValue.Value(1,li) = ""
				Next 
				CallSub objRecord, 15, mstrErr, mstrVar, mstrParam, Key
				If objPage.HandleUSPErrorAlert(mstrErr,"FORM") Then
					objPage.Record = objRecord.ArrayValue.Value
				End If
			Case "VoucherEnq"
				objPage.WriteData "CALV","__SUB"
				If objRecord.ArrayValue.Value(1,43) = "R" Then
					objPage.WriteData "__PARAM","ACENQVOC001-1.ASP,RV,ENQ"
				Else
					objPage.WriteData "__PARAM","ACENQVOC001-1.ASP,PV,ENQ"
				End If
				Response.Write "<SCRIPT language=JavaScript>" & vbCrLf
				Response.Write "Window_Opener('" & objPage.GetArgument & "','" & objPage.EndPageData() & "')" & vbCrLf
				Response.Write "</SCRIPT>"
	    End Select
	End If
%>
<!-- #include file="../Shared/StdInc2.asp"-->
<table class="table-border-title" border="1" width="100%" cellspacing="1" cellpadding="1">
	<tr>
		<td><%=objLookUp.ArrayValue.Value(1,2)%></td>
	</tr>
</table>
<table class="table-border" width="100%" cellspacing="0" cellpadding="0">	
<tr>
	<th colspan=4><%=objLookUp.ArrayValue.Value(1,1,1)%></th>
</tr>
<tr>
	<td class="subHeader" colspan=4><%=objLookUp.ArrayValue.Value(1,1,2)%></td>
</tr>
<tr>
	<td class="Caption" <%=STD_WIDTH_3%>><%=objLookUp.ArrayValue.Value(1,1,3)%></td>
	<td class="TextBox" <%=STD_WIDTH_4%>><%objPage.OptionDisplay "optType", objRecord.ArrayValue.Value(1,43), objLookUp.ArrayValue.Value(1,5), "1", " style='width:auto'"%></td>
	<td class="Caption" <%=STD_WIDTH_3%>><%=objLookUp.ArrayValue.Value(1,1,4)%></td>
	<td class="TextBox" <%=STD_WIDTH_4%>><INPUT class="align-r" id=txtMaxQuery name=txtMaxQuery value="<%=objRecord.ArrayValue.Value(1,44)%>" onchange="this.value=NumberCheck(this.value,0,0,999999999999,'N')"></td>
</tr>
<tr>
	<td class="Caption" <%=STD_WIDTH_3%>><%=objLookUp.ArrayValue.Value(1,1,5)%></td>
	<td class="TextBox" <%=STD_WIDTH_4%>><INPUT id=txtChqBankFr name=txtChqBankFr value="<%=objRecord.ArrayValue.Value(1,45)%>" style="width:30%" onchange="callWindow('','BANK.GROUP','txtChqBankFr',this.value,'1,1','txtChqBankFrDesc,txtChqBankFrDesc1')">
	<INPUT id=txtChqBankFrDesc name=txtChqBankFrDesc value="<%=objRecord.ArrayValue.Value(1,814)%>" ReadOnly style="width:68%"></td>
	<td class="Caption" <%=STD_WIDTH_3%>><%=objLookUp.ArrayValue.Value(1,1,6)%></td>
	<td class="TextBox" <%=STD_WIDTH_4%>><INPUT id=txtChqBankTo name=txtChqBankTo value="<%=objRecord.ArrayValue.Value(1,46)%>" style="width:30%" onchange="callWindow('','BANK.GROUP','txtChqBankTo',this.value,'1,1','txtChqBankToDesc,txtChqBankToDesc1')">
	<INPUT id=txtChqBankToDesc name=txtChqBankToDesc value="<%=objRecord.ArrayValue.Value(1,815)%>" ReadOnly style="width:68%"></td>
</tr>
<tr>
	<td class="Caption" <%=STD_WIDTH_3%>><%=objLookUp.ArrayValue.Value(1,1,7)%></td>
	<td class="TextBox" <%=STD_WIDTH_4%>><INPUT id=txtChqBrnFr name=txtChqBrnFr value="<%=objRecord.ArrayValue.Value(1,47)%>" style="width:30%" onchange="callWindow('','CHEQUE.BANK','txtChqBrnFr',this.value,'1,1','txtChqBrnFrDesc,txtChqBrnFrDesc1')">
	<INPUT id=txtChqBrnFrDesc name=txtChqBrnFrDesc value="<%=objRecord.ArrayValue.Value(1,816)%>" ReadOnly style="width:68%"></td>
	<td class="Caption" <%=STD_WIDTH_3%>><%=objLookUp.ArrayValue.Value(1,1,8)%></td>
	<td class="TextBox" <%=STD_WIDTH_4%>><INPUT id=txtChqBrnTo name=txtChqBrnTo value="<%=objRecord.ArrayValue.Value(1,48)%>" style="width:30%" onchange="callWindow('','CHEQUE.BANK','txtChqBrnTo',this.value,'1,1','txtChqBrnToDesc,txtChqBrnToDesc1')">
	<INPUT id=txtChqBrnToDesc name=txtChqBrnToDesc value="<%=objRecord.ArrayValue.Value(1,817)%>" ReadOnly style="width:68%"></td>
</tr>
<tr>
	<td class="Caption" <%=STD_WIDTH_3%>><%=objLookUp.ArrayValue.Value(1,1,9)%></td>
	<td class="TextBox" <%=STD_WIDTH_4%>><INPUT id=txtChqNbrFr name=txtChqNbrFr value="<%=objRecord.ArrayValue.Value(1,49)%>"></td>
	<td class="Caption" <%=STD_WIDTH_3%>><%=objLookUp.ArrayValue.Value(1,1,10)%></td>
	<td class="TextBox" <%=STD_WIDTH_4%>><INPUT id=txtChqNbrTo name=txtChqNbrTo value="<%=objRecord.ArrayValue.Value(1,50)%>"></td>
</tr>
<tr>
	<td class="Caption" <%=STD_WIDTH_3%>><%=objLookUp.ArrayValue.Value(1,1,11)%></td>
	<td class="TextBox" <%=STD_WIDTH_4%>><INPUT id=txtChqDateFr name=txtChqDateFr value="<%=objRecord.ArrayValue.Value(1,51)%>" onchange="this.value=formatDate(this.value)"></td>
	<td class="Caption" <%=STD_WIDTH_3%>><%=objLookUp.ArrayValue.Value(1,1,12)%></td>
	<td class="TextBox" <%=STD_WIDTH_4%>><INPUT id=txtChqDateTo name=txtChqDateTo value="<%=objRecord.ArrayValue.Value(1,52)%>" onchange="this.value=formatDate(this.value)"></td>
</tr>
<tr>
	<td class="Caption" <%=STD_WIDTH_3%>><%=objLookUp.ArrayValue.Value(1,1,13)%></td>
	<td class="TextBox" <%=STD_WIDTH_4%>><INPUT id=txtClrDateFr name=txtClrDateFr value="<%=objRecord.ArrayValue.Value(1,53)%>" onchange="this.value=formatDate(this.value)"></td>
	<td class="Caption" <%=STD_WIDTH_3%>><%=objLookUp.ArrayValue.Value(1,1,14)%></td>
	<td class="TextBox" <%=STD_WIDTH_4%>><INPUT id=txtClrDateTo name=txtClrDateTo value="<%=objRecord.ArrayValue.Value(1,54)%>" onchange="this.value=formatDate(this.value)"></td>
</tr>
<tr>
	<td class="Caption" <%=STD_WIDTH_3%>><%=objLookUp.ArrayValue.Value(1,1,15)%></td>
	<td class="TextBox" <%=STD_WIDTH_4%>><%objPage.ComboDisplay "cbxChqTypeFr", objRecord.ArrayValue.Value(1,55), objLookUp.ArrayValue.Value(1,3), "", ""%></td>
	<td class="Caption" <%=STD_WIDTH_3%>><%=objLookUp.ArrayValue.Value(1,1,16)%></td>
	<td class="TextBox" <%=STD_WIDTH_4%>><%objPage.ComboDisplay "cbxChqTypeTo", objRecord.ArrayValue.Value(1,56), objLookUp.ArrayValue.Value(1,3), "", ""%></td>
</tr>
<tr>
	<td class="Caption" <%=STD_WIDTH_3%>><%=objLookUp.ArrayValue.Value(1,1,17)%></td>
	<td class="TextBox" <%=STD_WIDTH_4%>><%objPage.ComboDisplay "cbxChqSttFr", objRecord.ArrayValue.Value(1,57), objLookUp.ArrayValue.Value(1,4), "", ""%></td>
	<td class="Caption" <%=STD_WIDTH_3%>><%=objLookUp.ArrayValue.Value(1,1,18)%></td>
	<td class="TextBox" <%=STD_WIDTH_4%>><%objPage.ComboDisplay "cbxChqSttTo", objRecord.ArrayValue.Value(1,58), objLookUp.ArrayValue.Value(1,4), "", ""%></td>
</tr>
<tr>
	<td class="Caption" <%=STD_WIDTH_3%>><%=objLookUp.ArrayValue.Value(1,1,19)%></td>
	<td class="TextBox" <%=STD_WIDTH_4%>><INPUT id=txtPayInFr name=txtPayInFr value="<%=objRecord.ArrayValue.Value(1,59)%>" style="width:30%" onchange="callWindow('','BANK.MAST','txtPayInFr',this.value,'1,1','txtPayInFrDesc,txtPayInFrDesc1')">
	<INPUT id=txtPayInFrDesc name=txtPayInFrDesc value="<%=objRecord.ArrayValue.Value(1,818)%>" ReadOnly style="width:68%"></td>
	<td class="Caption" <%=STD_WIDTH_3%>><%=objLookUp.ArrayValue.Value(1,1,20)%></td>
	<td class="TextBox" <%=STD_WIDTH_4%>><INPUT id=txtPayInTo name=txtPayInTo value="<%=objRecord.ArrayValue.Value(1,60)%>" style="width:30%" onchange="callWindow('','BANK.MAST','txtPayInTo',this.value,'1,1','txtPayInToDesc,txtPayInToDesc1')">
	<INPUT id=txtPayInToDesc name=txtPayInToDesc value="<%=objRecord.ArrayValue.Value(1,819)%>" ReadOnly style="width:68%"></td>
</tr>
<tr>
	<td class="Caption" <%=STD_WIDTH_3%>><%=objLookUp.ArrayValue.Value(1,1,21)%></td>
	<td class="TextBox" <%=STD_WIDTH_4%>><INPUT id=txtVoucDateFr name=txtVoucDateFr value="<%=objRecord.ArrayValue.Value(1,61)%>" onchange="this.value=formatDate(this.value)"></td>
	<td class="Caption" <%=STD_WIDTH_3%>><%=objLookUp.ArrayValue.Value(1,1,22)%></td>
	<td class="TextBox" <%=STD_WIDTH_4%>><INPUT id=txtVoucDateTo name=txtVoucDateTo value="<%=objRecord.ArrayValue.Value(1,62)%>" onchange="this.value=formatDate(this.value)"></td>
</tr>
<tr>
	<td class="Caption" <%=STD_WIDTH_3%>><%=objLookUp.ArrayValue.Value(1,1,23)%></td>
	<td class="TextBox" <%=STD_WIDTH_4%>><%objPage.OptionDisplay "optPayIn", objRecord.ArrayValue.Value(1,63), objLookUp.ArrayValue.Value(1,6), "", " style='width:auto'"%></td>
	<td class="TextBox" colspan=2></td>
</tr>
<tr>
	<td class="subHeader" colspan=4><%=objLookUp.ArrayValue.Value(1,1,24)%></td>
</tr>
<tr>
	<td colspan=4 class="table-noborder">
	<%objPage.TBLDisplay dsaCheqDet1.Value, "CheqDet1"%>
	</td>
</tr>
<tr>
	<td colspan=4 class="table-noborder">
	<INPUT type="button" class="button" value="<%=objLookUp.ArrayValue.Value(1,1,25)%>" onclick="__doPostBack('','Query','')" style="width:70">
	</td>
</tr>
<tr>
	<td class="table-noborder" colspan=4>
<!-- #include file="../Shared/StdInc3.asp"-->
	</td>
</tr>
</table>
<INPUT id=txtChqBankFrDesc1 name=txtChqBankFrDesc1 type="hidden" value="<%=objRecord.ArrayValue.Value(1,814)%>">
<INPUT id=txtChqBankToDesc1 name=txtChqBankToDesc1 type="hidden" value="<%=objRecord.ArrayValue.Value(1,815)%>">
<INPUT id=txtChqBrnFrDesc1 name=txtChqBrnFrDesc1 type="hidden" value="<%=objRecord.ArrayValue.Value(1,816)%>">
<INPUT id=txtChqBrnToDesc1 name=txtChqBrnToDesc1 type="hidden" value="<%=objRecord.ArrayValue.Value(1,817)%>">
<INPUT id=txtPayInFrDesc1 name=txtPayInFrDesc1 type="hidden" value="<%=objRecord.ArrayValue.Value(1,818)%>">
<INPUT id=txtPayInToDesc1 name=txtPayInToDesc1 type="hidden" value="<%=objRecord.ArrayValue.Value(1,819)%>">
<!-- #include file="../Shared/BrowserDesc.asp" -->
<!-- #include file="../Shared/SearchXrf.asp" -->
<% objPage.EndForm "FORM"  ' Call at end of form %>
</FORM>
</BODY>
</HTML>
<%  
    Set objPage = Nothing
    Set objRecord = Nothing
    Set objLookUp = Nothing %>


