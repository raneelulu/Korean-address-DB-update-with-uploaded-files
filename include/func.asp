<%
''################################################
''	File Name : func.asp
''	Summary   : functions (included)
''################################################
%>
<%

CONST YOUR_DBIP = ""
CONST YOUR_DBID = ""
CONST YOUR_DBPW = ""
CONST YOUR_DBNAME = ""

strDataSourceName="Driver={SQL Server};Server=" & YOUR_DBIP & ";UID=" & YOUR_DBID & ";PWD=" & YOUR_DBPW & ";Database=" & YOUR_DBNAME
numCursorType=3
Set objDataConn = Server.CreateObject("ADODB.Connection")
objDataConn.Open strDataSourceName
Set cmdTemp = Server.CreateObject("ADODB.Command")
Set objRecordSet = Server.CreateObject("ADODB.Recordset")

Function ExecuteSQL(strQuery, objRecordSet)
	cmdTemp.CommandText = strQuery
	cmdTemp.CommandType = 1
	Set cmdTemp.ActiveConnection = objDataConn

	objRecordSet.Open cmdTemp, , 1, numCursorType
End Function

Function GetPMBdb()
	Dim FCDB, FCSTR
	Set	FCDB = Server.CreateObject("ADODB.Connection")
	FCSTR = "Provider=SQLOLEDB;Data Source=" & YOUR_DBIP &";Initial Catalog=" & YOUR_DBNAME & ";user ID=" & YOUR_DBID &";password="& YOUR_DBPW &";"
	FCDB.Open(FCSTR)
	
	Set GetPMBdb = FCDB
End Function

''==================================================================================================
'' string change
''==================================================================================================
Function checkvalue(fcKind,fcValue)
  fcValue = Rtrim(fcValue)

  If isnull(fcValue) Then
	fcValue = ""
  ElseIf fcValue = "" Then

  ElseIf fcKind = 1 Then
	fcValue = replace(fcValue,"'","`")
	fcValue = replace(fcValue,"&nbsp;&nbsp;",chr(32)+chr(32))
	fcValue = replace(fcValue,"--","-[@]-")
	fcValue = Replace(fcValue, ";", "--")
	fcValue = Replace(fcValue, "select", "")
	fcValue = Replace(fcValue, "create", "")
	fcValue = Replace(fcValue, "drop", "")
	fcValue = Replace(fcValue, "delete", "")
	fcValue = Replace(fcValue, "truncate", "")
	fcValue = Replace(fcValue, "update", "")
	fcValue = Replace(fcValue, "exec", "")
	fcValue = Replace(fcValue, "SELECT", "")
	fcValue = Replace(fcValue, "CREATE", "")
	fcValue = Replace(fcValue, "DROP", "")
	fcValue = Replace(fcValue, "DELETE", "")
	fcValue = Replace(fcValue, "TRUNCATE", "")
	fcValue = Replace(fcValue, "UPDATE", "")
	fcValue = Replace(fcValue, "EXEC", "")
	fcValue = Replace(fcValue, "Select", "")
	fcValue = Replace(fcValue, "Create", "")
	fcValue = Replace(fcValue, "Drop", "")
	fcValue = Replace(fcValue, "Delete", "")
	fcValue = Replace(fcValue, "Truncate", "")
	fcValue = Replace(fcValue, "Update", "")
	fcValue = Replace(fcValue, "Exec", "")
	fcValue = Replace(fcValue, "db_name()", "")
	fcValue = Replace(fcValue, "char(94)", "")
	fcValue = Replace(fcValue, "declare", "")
	fcValue = Replace(fcValue, "Declare", "")

  checkvalue = fcValue
End Function

Function DateFormat(strdate, format)
	strYear = Year(strdate)
	strMonth = Month(strdate)
	strDay = Day(strdate)
	strHour = Hour(strdate)
	strMinute = Minute(strdate)
	strSecond = Second(strdate)
	If strMonth < 10 Then strMonth = "0" & strMonth
	If strDay < 10 Then strDay = "0" & strDay
	If strHour < 10 Then strHour = "0" & strHour
	If strMinute < 10 Then strMinute = "0" & strMinute
	If strSecond < 10 Then strSecond = "0" & strSecond
	If format = "D" Then 
		DateFormat = strYear & "-" & strMonth & "-" & strDay & " " & strHour & ":" & strMinute
	End If 
End Function

Function AlertMessage(msg)
	response.write "<script language='javascript'>alert('" & msg & "');history.back();</script>"
	response.end
End Function

Set rootdb=GetPMBdb()
%>