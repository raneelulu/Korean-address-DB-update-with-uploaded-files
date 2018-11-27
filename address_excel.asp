<!--#include file="include/func.asp"-->
<!--#include file="include/top.asp"-->
<!--#include file="include/freeaspupload.asp"-->
<%
'' first set your database and table
'' existing..database (database and table before update)
'' newtable..for_address_update (new data table for update list)
%>
<%
    Function getAddr(FileName, SheetName)
        If Instr(FileName,"B")  THEN
            response.write "법정동 시작시간 : " & now & "<br/>"
            response.write "<tr id='Bdong'>"
            response.write "<td>법정동</td>"
        ElseIf Instr(FileName, "H") THEN
            response.write "행정동 시작시간 : " & now & "<br/>"
            response.write "<tr id='Hdong'>"
            response.write "<td>행정동</td>"
        End If
        'set the path that you saved your files in variable 'targetFile'
        targetFile = path_that_you_want_to_save & FileName
        connectString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source="&targetFile&"; Extended Properties=""Excel 12.0;HDR=YES;IMEX=1;""" 
		conn.Mode=1
        conn.Open connectString

        sql = "SELECT [시도명],[시군구명],[읍면동명] FROM [" & SheetName & "$]"
        Rs.Open sql,conn
        IF NOT RS.EOF THEN
            completesido = ""
            psido = ""
            pgugun = ""
            pdong = ""
            WHILE NOT RS.eof    'excel's first line
                sido = ""
                gugun = ""
                dong = ""
                i = 0
                FOR EACH Field IN RS.Fields
                    If Instr(Field.value,"출장") <> 0 THEN EXIT FOR
                    i = i + 1
                    If i = 1 THEN
                        sido = Field.value
                    ElseIf i = 2 THEN
                        gugun = Field.value
                    ElseIf i = 3 THEN
                        dong = Field.value
                        If dong <> "" THEN                          'O?O
                            If sido = "세종특별자치시" Then
                                gugun = "세종시"
                                showProgress 2,"'"&gugun&"'"          'gu and gugun show
                            End If
                            If psido = sido and pgugun = gugun and pdong = dong THEN
                            Else
                                sql = "select * from existing..database where sido='" & sido & "' and gugun='" & gugun & "' and dong='" & dong & "'"
                                Set Grs = rootdb.execute(sql)
                                If Grs.EOF then 'not exist at database
                                    sql = "insert into newtable..for_address_update (sido,gugun,dong,updated,exist) values('" & sido & "', '" & gugun & "', '" & dong & "', '" & inputDate & "', " & 0 & ")"
                                    rootdb.execute(sql)
                                Else 'exist at database
                                    sql = "update existing..database set exist=" & 1	& "where sido='" & sido & "' and gugun='" & gugun & "' and dong='" & dong & "'"
                                    rootdb.execute(sql)
                                End If
                                psido = sido
                                pgugun = gugun
                                pdong = dong
                            End If
                        Else
                            If gugun <> "" THEN                       'OOX
                                showProgress 2,"'"&gugun&"'"          'gu and gugun show
                            Else                                      'OXX
                                If sido <> completesido THEN
                                    showProgress 1,"'"&sido&"'"
                                    If completesido <> "" THEN
                                        If Instr(FileName,"B")  THEN
                                            showProgress 3,"'Bdong'"
                                        ElseIf Instr(FileName, "H") THEN
                                            showProgress 3,"'Hdong'"
                                        End If
                                    End If
                                    completesido = sido
                                End If
                            END If
                        END If
                    END If
                NEXT
                RS.movenext
            WEND
            If Instr(FileName,"B")  THEN
                showProgress 3,"'Bdong'"
                response.write "법정동 종료시간 : " & now & "<br/>"
            ElseIf Instr(FileName, "H") THEN
                showProgress 3,"'Hdong'"
                response.write "행정동 종료시간 : " & now & "<br/>"
            End If
            Rs.close
            conn.close
        END IF
        response.write("</tr>")
        Response.flush
    End Function

    Function getDeleteAddr()
        sql = "select * from existing..database where exist is null"
        Set Grs = rootdb.execute(sql)
        Do Until Grs.EOF
            sido = Grs("sido")
            gugun = Grs("gugun")
            dong = Grs("dong")
            updated = Grs("updated")
            sql = "insert into newtable..for_address_update (sido,gugun,dong,updated,exist) values('" & sido & "', '" & gugun & "', '" & dong & "', '" & updated & "', " & 1 & ")"
            rootdb.execute(sql)
            Grs.movenext
        Loop
        Grs.close
    End Function

    Function showProgress(cnt,str)
        response.write "<script>showProgress(" & cnt & ", " & str &")</script>"
        response.flush
    End Function
%>        

<!doctype html>
<html lang="ko">
<head>
<title>주소 디비 업데이트</title>
</head>
<body>
<div id="content">
    <h2>주소 업데이트</h2>
    <div>
        <h3>엑셀파일형식</h3>
        <div style="margin:20px">
            <table>
                <thead>
                    <tr>
                        <th style="width:20%">법정동코드(행정동코드)</th>
                        <th style="width:14%">시도명</th>
                        <th style="width:14%">시군구명</th>
                        <th style="width:14%">읍면동명</th>
                        <th style="width:12%">동리명</th>
                        <th style="width:14%">생성일자</th>
                        <th style="width:12%">말소일자</th>
                    </tr>
                </thead>
                <tbody>
                    <tr>
                        <td>1100000000</td>
                        <td>서울특별시</td>
                        <td></td>
                        <td></td>
                        <td></td>
                        <td>19880423</td>
                        <td></td>
                    </tr>
                    <tr>
                        <td>1111000000</td>
                        <td>서울특별시</td>
                        <td>종로구</td>
                        <td></td>
                        <td></td>
                        <td>19880423</td>
                        <td></td>
                    </tr>
                    <tr>
                        <td>1111010100</td>
                        <td>서울특별시</td>
                        <td>종로구</td>
                        <td>청운동</td>
                        <td></td>
                        <td>19880423</td>
                        <td></td>
                    </tr>
                </tbody>
            </table>
        </div>
    </div>
    <div id="file_upload" style="display:none;">
        <h3>파일을 첨부해주세요</h3>
        <form name="uploadform" method="post" ENCTYPE="multipart/form-data">
            <input type=hidden name="save" value="upload" />
        <div style="margin-top:8px">
            <input type="file" name="uploadfile1"/><br>
            <input type="file" name="uploadfile2"/><br>
        </div>
			<input type="submit" name="submit" id="submit" value="주소 업데이트하기" class="button" onclick="hide()"/>
        </form>
    </div>
<%
    'set the path that you want to save your files in variable 'UPLOADPATH'
    UPLOADPATH = path_that_you_want_to_save & FileName

	Set Upload = New FreeASPUpload    
    Upload.upload
	ks = Upload.UploadedFiles.keys
	blnWrongType = False

	If Upload.form("save") = "upload" Then
%>
<%
        sql = "update existing..database set exist=null"
        rootdb.execute(sql)
        sql = "delete from newtable..for_address_update"
        rootdb.execute(sql)

        response.buffer = true
        Set conn = Server.CreateObject("ADODB.Connection")  
        Set Rs = Server.CreateObject("ADODB.RecordSet")  

        savetimeout = server.ScriptTimeout
        Server.ScriptTimeout = 180000
%>
<style>
#sido {display:inline-block;width:100px}
#gugun {display:inline-block;width:140px}
</style>
<script>
function showProgress(cnt,str)  {
    if(cnt==1)   {
        document.getElementById("sido").innerHTML = str;
    }
    else if(cnt==2) {
        document.getElementById("gugun").innerHTML = str;
    }
    else if(cnt==3) {
        var td = document.createElement("td");
        td.innerHTML = "O";
        document.getElementById(str).appendChild(td);
    }
    else if(cnt==4) {
        document.getElementById("sido").innerHTML = "삭제 리스트 검색 중";
        document.getElementById("gugun").innerHTML = "";
    }
    return;
}
</script>
    <div>
        <h3>진행상황</h3>
        <div style="margin:20px">
            <span id="sido"></span>
            <span id="gugun"></span>
        </div>
    </div>
    <div>
        <h3>완료된 시도명</h3>
        <div id="complete" style="margin:20px">
            <table style="margin-top:20px">
                <thead>
                    <tr>
                        <th></th>
                        <th>서울</th>
                        <th>부산</th>
                        <th>대구</th>
                        <th>인천</th>
                        <th>광주</th>
                        <th>대전</th>
                        <th>울산</th>
                        <th>세종</th>
                        <th>경기</th>
                        <th>강원</th>
                        <th>충북</th>
                        <th>충남</th>
                        <th>전북</th>
                        <th>전남</th>
                        <th>경북</th>
                        <th>경남</th>
                        <th>제주</th>
                    </tr>
                </thead>
                <tbody>
<%
		'' file upload
		If (UBound(ks) <> -1) then
			num = 0
			for each fileKey in ks            
                downfilename = ""
                SheetName = ""

				aryFileName = Split(Upload.UploadedFiles(fileKey).FileName, ".")
				strFileExe = LCase(aryFileName(UBound(aryFileName)))

				If UBound(aryFileName) < 1 Then
					blnWrongType = True
				ElseIf strFileExe <> "xlsx" And strFileExe <> "xls" Then
					blnWrongType = True
				Else
					If InStr(UCase(G_UploadFileTypes), UCase(aryFileName(UBound(aryFileName)))) <= 0 Then 
						blnWrongType = True 
					End If 
				End If 
				If blnWrongType Then
					AlertMessage "업로드 불가능한 파일입니다(xlsx, xls 파일만 가능)"
				End If 

				If Instr(fileKey, "uploadfile") Then
                    For i = 0 to UBound(aryFileName) - 1
                        downfilename = downfilename & aryFileName(i) & "."
                        If IsNumeric(aryFileName(i)) Then
                            inputDate = aryFileName(i)
                        Else
                            SheetName = aryFileName(i)
                        End if
                    Next
					downfilename = downfilename & strFileExe
					Upload.SaveAs UPLOADPATH,  num, downfilename
                    getAddr downfilename, SheetName
				End If 
				num =  num + 1
			next
		End If

		If IsNull(downfilename) Or downfilename = "" Then
			AlertMessage "파일을 선택해주십시오."
		End If

%>

<%
        response.flush
        showProgress 4,"''"
        getDeleteAddr()
        Server.ScriptTimeout = savetimeout
        response.write "<script>window.location.replace('/address_updated.asp');</script>"
    Else
        response.write "<script>document.getElementById('file_upload').style.display='';</script>"
    END If
%>
                </tbody>
            </table>
        </div>
    </div>
</div>
</body>
</html>
