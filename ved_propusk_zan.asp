<meta content="text/html; charset=windows-1251" http-equiv="content-type" />
<link rel="shortcut icon" href="images/favicon.ico" />
<head>
	<meta content="text/html; charset=Windows-1251" http-equiv="content-type">
	<link rel="shortcut icon" href="images/favicon.ico" /> 
	<link rel="stylesheet" href="css/metro.css">
	<link rel="stylesheet" href="css/metro-colors.css">
	<link rel="stylesheet" href="css/metro-icons.css">
	<link rel="stylesheet" href="css/metro-responsive.css">
	<link rel="stylesheet" href="css/metro-rtl.css">
	<link rel="stylesheet" href="css/metro-student.css">
	<link rel="stylesheet" href="css/metro-schemes.css">
	<link rel="stylesheet" href="css/loaders.css">
	<script src="js/jquery-3.1.0.min.js"></script>
	<script src="js/metro.min.js"></script>
	<title>ИС "Электронный журнал". Ведомость пропуска занятий</title>
    <style>
        .v-text {
            writing-mode: vertical-lr;
            text-orientation: mixed;
            margin: 0 auto;
        }
        .form-table tbody td {
            padding: 0px;
        }
        .switch-original input[type="checkbox"]:checked ~ .check {
            background: #2086bf0;
        }

        @media print {
            #notPrintableArea {
                visibility: hidden;
				position: absolute;
				top: 0;
				left: 0;
				width: 0;
            }
            #printableArea {
                visibility: visible !important;
				position: absolute;
				left: 0;
				top: 0;
				width: 100vw;
            }
            body {
                font-family: Times, 'Times New Roman', serif;
				font-size: 12pt;
            }
            h1, h2, h3, h4, h5, h6, p {
                color: #000;
                font-family: Times, 'Times New Roman', serif;
            }
			.table thead {
				border-bottom: 0;
			}
			table {
				background-color: unset;
				width: 100vw;
			}
			.table.bordered th, .table.bordered td {
				border: 1px #000 solid;
			}
			.table.border {
				border: 1px #000 solid;
			}
			.table thead th, .table thead td {
				color: #000;
			}
			.table tbody td {
				padding: 0;
				height: 0.58cm;
				padding-left: 0.19cm;
				padding-right: 0.19cm;
			}
        }
    </style>
</head>
<body>
<div id="notPrintableArea">
<table class="table border" style='width:90%; margin-top: 15px;' align=center> 	
<tr>
<td>
    <!-- #include file="header.asp" -->
	<!-- #include file="pass_check.asp" -->
<%
'Защита от студентов
if session("user") = "" or session("user") = "Студент" or session("user") = 0 then response.Redirect ("404.asp")
%>

<%
function Log(value)
    response.Write("<script language=javascript>console.log('" & value & "'); </script>")
end function

query = request.querystring("query")

if query = 1 then

'Выполняем подключение к БД
Set con = Server.CreateObject("ADODB.Connection")
Set rs_all = Server.CreateObject("ADODB.RecordSet")
Set rs_sep = Server.CreateObject("ADODB.RecordSet")
Set rs_oct = Server.CreateObject("ADODB.RecordSet")
Set rs_nov = Server.CreateObject("ADODB.RecordSet")
Set rs_dec = Server.CreateObject("ADODB.RecordSet")
Set rs_sem1 = Server.CreateObject("ADODB.RecordSet")
Set rs_jan = Server.CreateObject("ADODB.RecordSet")
Set rs_feb = Server.CreateObject("ADODB.RecordSet")
Set rs_mar = Server.CreateObject("ADODB.RecordSet")
Set rs_apr = Server.CreateObject("ADODB.RecordSet")
Set rs_sem2 = Server.CreateObject("ADODB.RecordSet")
Set rs_total = Server.CreateObject("ADODB.RecordSet")
Set rs_groupSearch = Server.CreateObject("ADODB.RecordSet")

strdbpath=server.mappath("base.mdb")
con.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strdbpath

groupSearch = request.Form("groupSearch")
dateSearch = request.Form("date")
empty_cols = request.Form("empty-cols"): if empty_cols = "on" then empty_cols = true else empty_cols = false
dateFormatted = mid(dateSearch, 9, 2) + "." + mid(dateSearch, 6, 2) + "." + mid(dateSearch, 1, 4)

now_year       = cint(mid(dateSearch, 1, 4))
selected_group = cstr(groupSearch)

sql_all         = "SELECT tbl_group.id_group, tbl_group.group_name, tbl_student.id_student, Left(LTrim([student_fio]),InStr(LTrim([student_fio]) & '',' '))+[student_nam] AS student_fi, Sum(IIf([nb1]=True,1,IIf([nb2]=True,2,IIf([nb1]=False And [nb2]=False,0,'')))) AS nb, tbl_student.student_fio FROM (tbl_group INNER JOIN tbl_student ON tbl_group.id_group = tbl_student.group_name) INNER JOIN tbl_journal ON tbl_student.id_student = tbl_journal.student_name GROUP BY tbl_group.id_group, tbl_group.group_name, tbl_student.id_student, Left(LTrim([student_fio]),InStr(LTrim([student_fio]) & '',' '))+[student_nam], tbl_student.student_fio HAVING (((tbl_group.id_group)=" + selected_group + ")) ORDER BY tbl_student.student_fio;"
sql_sep         = "SELECT tbl_group.id_group, tbl_group.group_name, tbl_student.id_student, Left(LTrim([student_fio]),InStr(LTrim([student_fio]) & '',' '))+[student_nam] AS student_fi, Sum(IIf([nb1]=True,1,IIf([nb2]=True,2,IIf([nb1]=False And [nb2]=False,0,'')))) AS nb, tbl_student.student_fio FROM (tbl_group INNER JOIN tbl_student ON tbl_group.id_group = tbl_student.group_name) INNER JOIN tbl_journal ON tbl_student.id_student = tbl_journal.student_name WHERE (((tbl_journal.nb_type)='5' Or (tbl_journal.nb_type)='6' Or (tbl_journal.nb_type)='7' Or (tbl_journal.nb_type) Is Null) AND ((tbl_journal.edit_date) Between #9/1/" + cstr(now_year - 1) + "# And #9/30/" + cstr(now_year - 1) + "#)) GROUP BY tbl_group.id_group, tbl_group.group_name, tbl_student.id_student, Left(LTrim([student_fio]),InStr(LTrim([student_fio]) & '',' '))+[student_nam], tbl_student.student_fio HAVING (((tbl_group.id_group)=" + selected_group + ")) ORDER BY tbl_student.student_fio;"
sql_oct         = "SELECT tbl_group.id_group, tbl_group.group_name, tbl_student.id_student, Left(LTrim([student_fio]),InStr(LTrim([student_fio]) & '',' '))+[student_nam] AS student_fi, Sum(IIf([nb1]=True,1,IIf([nb2]=True,2,IIf([nb1]=False And [nb2]=False,0,'')))) AS nb, tbl_student.student_fio FROM (tbl_group INNER JOIN tbl_student ON tbl_group.id_group = tbl_student.group_name) INNER JOIN tbl_journal ON tbl_student.id_student = tbl_journal.student_name WHERE (((tbl_journal.nb_type)='5' Or (tbl_journal.nb_type)='6' Or (tbl_journal.nb_type)='7' Or (tbl_journal.nb_type) Is Null) AND ((tbl_journal.edit_date) Between #10/1/" + cstr(now_year - 1) + "# And #10/31/" + cstr(now_year - 1) + "#)) GROUP BY tbl_group.id_group, tbl_group.group_name, tbl_student.id_student, Left(LTrim([student_fio]),InStr(LTrim([student_fio]) & '',' '))+[student_nam], tbl_student.student_fio HAVING (((tbl_group.id_group)=" + selected_group + ")) ORDER BY tbl_student.student_fio;"
sql_nov         = "SELECT tbl_group.id_group, tbl_group.group_name, tbl_student.id_student, Left(LTrim([student_fio]),InStr(LTrim([student_fio]) & '',' '))+[student_nam] AS student_fi, Sum(IIf([nb1]=True,1,IIf([nb2]=True,2,IIf([nb1]=False And [nb2]=False,0,'')))) AS nb, tbl_student.student_fio FROM (tbl_group INNER JOIN tbl_student ON tbl_group.id_group = tbl_student.group_name) INNER JOIN tbl_journal ON tbl_student.id_student = tbl_journal.student_name WHERE (((tbl_journal.nb_type)='5' Or (tbl_journal.nb_type)='6' Or (tbl_journal.nb_type)='7' Or (tbl_journal.nb_type) Is Null) AND ((tbl_journal.edit_date) Between #11/1/" + cstr(now_year - 1) + "# And #11/30/" + cstr(now_year - 1) + "#)) GROUP BY tbl_group.id_group, tbl_group.group_name, tbl_student.id_student, Left(LTrim([student_fio]),InStr(LTrim([student_fio]) & '',' '))+[student_nam], tbl_student.student_fio HAVING (((tbl_group.id_group)=" + selected_group + ")) ORDER BY tbl_student.student_fio;"
sql_dec         = "SELECT tbl_group.id_group, tbl_group.group_name, tbl_student.id_student, Left(LTrim([student_fio]),InStr(LTrim([student_fio]) & '',' '))+[student_nam] AS student_fi, Sum(IIf([nb1]=True,1,IIf([nb2]=True,2,IIf([nb1]=False And [nb2]=False,0,'')))) AS nb, tbl_student.student_fio FROM (tbl_group INNER JOIN tbl_student ON tbl_group.id_group = tbl_student.group_name) INNER JOIN tbl_journal ON tbl_student.id_student = tbl_journal.student_name WHERE (((tbl_journal.nb_type)='5' Or (tbl_journal.nb_type)='6' Or (tbl_journal.nb_type)='7' Or (tbl_journal.nb_type) Is Null) AND ((tbl_journal.edit_date) Between #12/1/" + cstr(now_year - 1) + "# And #12/31/" + cstr(now_year - 1) + "#)) GROUP BY tbl_group.id_group, tbl_group.group_name, tbl_student.id_student, Left(LTrim([student_fio]),InStr(LTrim([student_fio]) & '',' '))+[student_nam], tbl_student.student_fio HAVING (((tbl_group.id_group)=" + selected_group + ")) ORDER BY tbl_student.student_fio;"
sql_sem1        = "SELECT tbl_group.id_group, tbl_group.group_name, tbl_student.id_student, Left(LTrim([student_fio]),InStr(LTrim([student_fio]) & '',' '))+[student_nam] AS student_fi, Sum(IIf([nb1]=True,1,IIf([nb2]=True,2,IIf([nb1]=False And [nb2]=False,0,'')))) AS nb, tbl_student.student_fio FROM (tbl_group INNER JOIN tbl_student ON tbl_group.id_group = tbl_student.group_name) INNER JOIN tbl_journal ON tbl_student.id_student = tbl_journal.student_name WHERE (((tbl_journal.nb_type)='5' Or (tbl_journal.nb_type)='6' Or (tbl_journal.nb_type)='7' Or (tbl_journal.nb_type) Is Null) AND ((tbl_journal.edit_date) Between #9/1/" + cstr(now_year - 1) + "# And #12/31/" + cstr(now_year - 1) + "#)) GROUP BY tbl_group.id_group, tbl_group.group_name, tbl_student.id_student, Left(LTrim([student_fio]),InStr(LTrim([student_fio]) & '',' '))+[student_nam], tbl_student.student_fio HAVING (((tbl_group.id_group)=" + selected_group + ")) ORDER BY tbl_student.student_fio;"
sql_jan         = "SELECT tbl_group.id_group, tbl_group.group_name, tbl_student.id_student, Left(LTrim([student_fio]),InStr(LTrim([student_fio]) & '',' '))+[student_nam] AS student_fi, Sum(IIf([nb1]=True,1,IIf([nb2]=True,2,IIf([nb1]=False And [nb2]=False,0,'')))) AS nb, tbl_student.student_fio FROM (tbl_group INNER JOIN tbl_student ON tbl_group.id_group = tbl_student.group_name) INNER JOIN tbl_journal ON tbl_student.id_student = tbl_journal.student_name WHERE (((tbl_journal.nb_type)='5' Or (tbl_journal.nb_type)='6' Or (tbl_journal.nb_type)='7' Or (tbl_journal.nb_type) Is Null) AND ((tbl_journal.edit_date) Between #1/1/" + cstr(now_year) + "# And #1/31/" + cstr(now_year) + "#)) GROUP BY tbl_group.id_group, tbl_group.group_name, tbl_student.id_student, Left(LTrim([student_fio]),InStr(LTrim([student_fio]) & '',' '))+[student_nam], tbl_student.student_fio HAVING (((tbl_group.id_group)=" + selected_group + ")) ORDER BY tbl_student.student_fio;"
sql_feb         = "SELECT tbl_group.id_group, tbl_group.group_name, tbl_student.id_student, Left(LTrim([student_fio]),InStr(LTrim([student_fio]) & '',' '))+[student_nam] AS student_fi, Sum(IIf([nb1]=True,1,IIf([nb2]=True,2,IIf([nb1]=False And [nb2]=False,0,'')))) AS nb, tbl_student.student_fio FROM (tbl_group INNER JOIN tbl_student ON tbl_group.id_group = tbl_student.group_name) INNER JOIN tbl_journal ON tbl_student.id_student = tbl_journal.student_name WHERE (((tbl_journal.nb_type)='5' Or (tbl_journal.nb_type)='6' Or (tbl_journal.nb_type)='7' Or (tbl_journal.nb_type) Is Null) AND ((tbl_journal.edit_date) Between #2/1/" + cstr(now_year) + "# And #2/28/" + cstr(now_year) + "#)) GROUP BY tbl_group.id_group, tbl_group.group_name, tbl_student.id_student, Left(LTrim([student_fio]),InStr(LTrim([student_fio]) & '',' '))+[student_nam], tbl_student.student_fio HAVING (((tbl_group.id_group)=" + selected_group + ")) ORDER BY tbl_student.student_fio;"
sql_mar         = "SELECT tbl_group.id_group, tbl_group.group_name, tbl_student.id_student, Left(LTrim([student_fio]),InStr(LTrim([student_fio]) & '',' '))+[student_nam] AS student_fi, Sum(IIf([nb1]=True,1,IIf([nb2]=True,2,IIf([nb1]=False And [nb2]=False,0,'')))) AS nb, tbl_student.student_fio FROM (tbl_group INNER JOIN tbl_student ON tbl_group.id_group = tbl_student.group_name) INNER JOIN tbl_journal ON tbl_student.id_student = tbl_journal.student_name WHERE (((tbl_journal.nb_type)='5' Or (tbl_journal.nb_type)='6' Or (tbl_journal.nb_type)='7' Or (tbl_journal.nb_type) Is Null) AND ((tbl_journal.edit_date) Between #3/1/" + cstr(now_year) + "# And #3/31/" + cstr(now_year) + "#)) GROUP BY tbl_group.id_group, tbl_group.group_name, tbl_student.id_student, Left(LTrim([student_fio]),InStr(LTrim([student_fio]) & '',' '))+[student_nam], tbl_student.student_fio HAVING (((tbl_group.id_group)=" + selected_group + ")) ORDER BY tbl_student.student_fio;"
sql_apr         = "SELECT tbl_group.id_group, tbl_group.group_name, tbl_student.id_student, Left(LTrim([student_fio]),InStr(LTrim([student_fio]) & '',' '))+[student_nam] AS student_fi, Sum(IIf([nb1]=True,1,IIf([nb2]=True,2,IIf([nb1]=False And [nb2]=False,0,'')))) AS nb, tbl_student.student_fio FROM (tbl_group INNER JOIN tbl_student ON tbl_group.id_group = tbl_student.group_name) INNER JOIN tbl_journal ON tbl_student.id_student = tbl_journal.student_name WHERE (((tbl_journal.nb_type)='5' Or (tbl_journal.nb_type)='6' Or (tbl_journal.nb_type)='7' Or (tbl_journal.nb_type) Is Null) AND ((tbl_journal.edit_date) Between #4/1/" + cstr(now_year) + "# And #4/30/" + cstr(now_year) + "#)) GROUP BY tbl_group.id_group, tbl_group.group_name, tbl_student.id_student, Left(LTrim([student_fio]),InStr(LTrim([student_fio]) & '',' '))+[student_nam], tbl_student.student_fio HAVING (((tbl_group.id_group)=" + selected_group + ")) ORDER BY tbl_student.student_fio;"
sql_sem2        = "SELECT tbl_group.id_group, tbl_group.group_name, tbl_student.id_student, Left(LTrim([student_fio]),InStr(LTrim([student_fio]) & '',' '))+[student_nam] AS student_fi, Sum(IIf([nb1]=True,1,IIf([nb2]=True,2,IIf([nb1]=False And [nb2]=False,0,'')))) AS nb, tbl_student.student_fio FROM (tbl_group INNER JOIN tbl_student ON tbl_group.id_group = tbl_student.group_name) INNER JOIN tbl_journal ON tbl_student.id_student = tbl_journal.student_name WHERE (((tbl_journal.nb_type)='5' Or (tbl_journal.nb_type)='6' Or (tbl_journal.nb_type)='7' Or (tbl_journal.nb_type) Is Null) AND ((tbl_journal.edit_date) Between #1/1/" + cstr(now_year) + "# And #4/30/" + cstr(now_year) + "#)) GROUP BY tbl_group.id_group, tbl_group.group_name, tbl_student.id_student, Left(LTrim([student_fio]),InStr(LTrim([student_fio]) & '',' '))+[student_nam], tbl_student.student_fio HAVING (((tbl_group.id_group)=" + selected_group + ")) ORDER BY tbl_student.student_fio;"
sql_total       = "SELECT tbl_group.id_group, tbl_group.group_name, tbl_student.id_student, Left(LTrim([student_fio]),InStr(LTrim([student_fio]) & '',' '))+[student_nam] AS student_fi, Sum(IIf([nb1]=True,1,IIf([nb2]=True,2,IIf([nb1]=False And [nb2]=False,0,'')))) AS nb, tbl_student.student_fio FROM (tbl_group INNER JOIN tbl_student ON tbl_group.id_group = tbl_student.group_name) INNER JOIN tbl_journal ON tbl_student.id_student = tbl_journal.student_name WHERE (((tbl_journal.nb_type)='5' Or (tbl_journal.nb_type)='6' Or (tbl_journal.nb_type)='7' Or (tbl_journal.nb_type) Is Null) AND ((tbl_journal.edit_date) Between #9/1/" + cstr(now_year - 1) + "# And #4/30/" + cstr(now_year) + "#)) GROUP BY tbl_group.id_group, tbl_group.group_name, tbl_student.id_student, Left(LTrim([student_fio]),InStr(LTrim([student_fio]) & '',' '))+[student_nam], tbl_student.student_fio HAVING (((tbl_group.id_group)=" + selected_group + ")) ORDER BY tbl_student.student_fio;"
sql_groupSearch = "SELECT id_group, group_name FROM tbl_group WHERE id_group=" + selected_group + " ORDER BY group_name;"

rs_all.Open sql_all, con, 3, 3
rs_sep.Open sql_sep, con, 3, 3
rs_oct.Open sql_oct, con, 3, 3
rs_nov.Open sql_nov, con, 3, 3
rs_dec.Open sql_dec, con, 3, 3
rs_sem1.Open sql_sem1, con, 3, 3
rs_jan.Open sql_jan, con, 3, 3
rs_feb.Open sql_feb, con, 3, 3
rs_mar.Open sql_mar, con, 3, 3
rs_apr.Open sql_apr, con, 3, 3
rs_sem2.Open sql_sem2, con, 3, 3
rs_total.Open sql_total, con, 3, 3
rs_groupSearch.Open sql_groupSearch, con, 3, 3

dim students(3000,14)
total_sep   = 0
total_oct   = 0
total_nov   = 0
total_dec   = 0
total_sem1  = 0
total_jan   = 0
total_feb   = 0
total_mar   = 0
total_apr   = 0
total_sem2  = 0
total_total = 0

sep_empty = false
oct_empty = false
nov_empty = false
dec_empty = false
jan_empty = false
feb_empty = false
mar_empty = false
apr_empty = false

Function InArray(theArray,theValue, i)
    dim fnd
    fnd = False
    For i = 0 to UBound(theArray)
        If theArray(i, 1) = theValue Then
            fnd = True
            Exit For
        End If
    Next
    InArray = fnd
End Function

if rs_all.RecordCount > 0 then
	ij = 1
	cnt = 1
	while rs_all.EOF <> true
        students(ij,1) = rs_all.Fields(2)
		students(ij,2) = rs_all.Fields(3)
		ij = ij + 1
		cnt = cnt + 1
		rs_all.MoveNext
	wend
elseif rs_sep.RecordCount <= 0 and empty_cols = false then
    sep_empty = true
end if

if rs_sep.RecordCount > 0 then
	ij = 1
    total_sep = 0
	while rs_sep.EOF <> true
		if rs_sep.Fields(4) > 0 and InArray(students, rs_sep.Fields(2), ij) = true then
            students(ij,3) = rs_sep.Fields(4)
        else
            students(ij,3) = ""
        end if
        total_sep = total_sep + rs_sep.Fields(4)
		ij = ij + 1
		rs_sep.MoveNext
	wend
elseif rs_sep.RecordCount <= 0 and empty_cols = false then
    sep_empty = true
end if

if rs_oct.RecordCount > 0 then
	ij = 1
	while rs_oct.EOF <> true
        if rs_oct.Fields(4) > 0 and InArray(students, rs_oct.Fields(2), ij) = true then
		    students(ij,4) = rs_oct.Fields(4)
        else
            students(ij,4) = ""
        end if
		total_oct = total_oct + rs_oct.Fields(4)
		ij = ij + 1
		rs_oct.MoveNext
	wend
elseif rs_oct.RecordCount <= 0 and empty_cols = false then
    oct_empty = true
end if

if rs_nov.RecordCount > 0 then
	ij = 1
	while rs_nov.EOF <> true
        if rs_nov.Fields(4) > 0 and InArray(students, rs_nov.Fields(2), ij) = true then
		    students(ij,5) = rs_nov.Fields(4)
        else
            students(ij,5) = ""
        end if
		total_nov = total_nov + rs_nov.Fields(4)
		ij = ij + 1
		rs_nov.MoveNext
	wend
elseif rs_nov.RecordCount <= 0 and empty_cols = false then
    nov_empty = true
end if

if rs_dec.RecordCount > 0 then
	ij = 1
	while rs_dec.EOF <> true
        if rs_dec.Fields(4) > 0 and InArray(students, rs_dec.Fields(2), ij) = true then
		    students(ij,6) = rs_dec.Fields(4)
        else
            students(ij,6) = ""
        end if
		total_dec = total_dec + rs_dec.Fields(4)
		ij = ij + 1
		rs_dec.MoveNext
	wend
elseif rs_dec.RecordCount <= 0 and empty_cols = false then
    dec_empty = true
end if

if rs_sem1.RecordCount > 0 then
	ij = 1
	while rs_sem1.EOF <> true
        if rs_sem1.Fields(4) > 0 and InArray(students, rs_sem1.Fields(2), ij) = true then
		    students(ij,7) = rs_sem1.Fields(4)
        else
            students(ij,7) = 0
        end if
		total_sem1 = total_sem1 + rs_sem1.Fields(4)
		ij = ij + 1
		rs_sem1.MoveNext
	wend
end if

if rs_jan.RecordCount > 0 then
	ij = 1
	while rs_jan.EOF <> true
        if rs_jan.Fields(4) > 0 and InArray(students, rs_jan.Fields(2), ij) = true then
		    students(ij,8) = rs_jan.Fields(4)
        else
            students(ij,8) = ""
        end if
		total_jan = total_jan + rs_jan.Fields(4)
		ij = ij + 1
		rs_jan.MoveNext
	wend
elseif rs_jan.RecordCount <= 0 and empty_cols = false then
    jan_empty = true
end if

if rs_feb.RecordCount > 0 then
	ij = 1
	while rs_feb.EOF <> true
        if rs_feb.Fields(4) > 0 and InArray(students, rs_feb.Fields(2), ij) = true then
		    students(ij,9) = rs_feb.Fields(4)
        else
            students(ij,9) = ""
        end if
		total_feb = total_feb + rs_feb.Fields(4)
		ij = ij + 1
		rs_feb.MoveNext
	wend
elseif rs_feb.RecordCount <= 0 and empty_cols = false then
    feb_empty = true
end if

if rs_mar.RecordCount > 0 then
	ij = 1
	while rs_mar.EOF <> true
        if rs_mar.Fields(4) > 0 and InArray(students, rs_mar.Fields(2), ij) = true then
		    students(ij,10) = rs_mar.Fields(4)
        else
            students(ij,10) = ""
        end if
		total_mar = total_mar + rs_mar.Fields(4)
		ij = ij + 1
		rs_mar.MoveNext
	wend
elseif rs_mar.RecordCount <= 0 and empty_cols = false then
    mar_empty = true
end if

if rs_apr.RecordCount > 0 then
	ij = 1
	while rs_apr.EOF <> true
        if rs_apr.Fields(4) > 0 and InArray(students, rs_apr.Fields(2), ij) = true then
		    students(ij,11) = rs_apr.Fields(4)
        else
            students(ij,11) = ""
        end if
		total_apr = total_apr + rs_apr.Fields(4)
		ij = ij + 1
		rs_apr.MoveNext
	wend
elseif rs_apr.RecordCount <= 0 and empty_cols = false then
    apr_empty = true
end if

if rs_sem2.RecordCount > 0 then
	ij = 1
	while rs_sem2.EOF <> true
        if rs_sem2.Fields(4) > 0 and InArray(students, rs_sem2.Fields(2), ij) = true then
		    students(ij,12) = rs_sem2.Fields(4)
        else
            students(ij,12) = 0
        end if
		total_sem2 = total_sem2 + rs_sem2.Fields(4)
		ij = ij + 1
		rs_sem2.MoveNext
	wend
end if

if rs_total.RecordCount > 0 then
	ij = 1
	while rs_total.EOF <> true
        if rs_total.Fields(4) > 0 and InArray(students, rs_total.Fields(2), ij) = true then
		    students(ij,13) = rs_total.Fields(4)
        else
            students(ij,13) = 0
        end if
		total_total = total_total + rs_total.Fields(4)
		ij = ij + 1
		rs_total.MoveNext
	wend
end if

rs_sep.Close
rs_oct.Close
rs_nov.Close
rs_dec.Close
rs_sem1.Close
rs_jan.Close
rs_feb.Close
rs_mar.Close
rs_apr.Close
rs_sem2.Close
rs_total.Close
%>

<p style="margin: 0; padding: 0; float: left;"><a href="ved_propusk_zan.asp?query=0" ><button class="button"><span class="icon mif-undo"></span> Вернуться назад</button></a></p><p style="margin: 0; padding: 0; float: right;"><button class="button primary" onclick="window.print()"><span class="icon mif-printer"></span> Печать</button></p><br><br><br>
</div>
<div id=printableArea>
<center>
<%

%>
<h4><b>Пропуски занятий <%=cstr(rs_groupSearch.Fields(1)) %> на <%=dateFormatted %> г. (по неуважительной причине)</b></h4>

<% if cnt > 1 then %>
<table class="table striped hovered cell-hovered border bordered" id="ved_propusk_zan">
<thead align=center style="font-weight: bold">
<tr>
    <th style="margin: 0;text-align: center;">№</th>
    <th style="margin: 0;text-align: center;">Фамилия</th>
    <% if not sep_empty then %><th><p class="v-text">Сентябрь</th>
    <% end if: if not oct_empty then  %>
    <th><p class="v-text">Октябрь</th>
    <% end if: if not nov_empty then  %>
    <th><p class="v-text">Ноябрь</th>
    <% end if: if not dec_empty then  %>
    <th><p class="v-text">Декабрь</th>
    <% end if %>
    <th style="width:1;"><p style="text-align:center;">Итого<br>за 1<br>семестр</th>
    <% if not jan_empty then %>
    <th><p class="v-text">Январь</th>
    <% end if: if not feb_empty then %>
    <th><p class="v-text">Февраль</th>
    <% end if: if not mar_empty then %>
    <th><p class="v-text">Март</th>
    <% end if: if not apr_empty then %>
    <th><p class="v-text">Апрель</th>
    <% end if %>
    <th style="width:1;"><p style="text-align:center;">Итого<br>за 2<br>семестр</th>
    <th style="width:1;"><p style="text-align:center;">Всего</th>
</tr>
</thead>
<tbody align=center>
<%
for i = 1 to cnt 'Запрос выполняется для каждого студента
	if students(i,1) > 0 then 'Если пользователь в массиве имеет свой ID
		'Подготавливаем данные
		response.Write("<tr>")
		
		'Рисуем таблицу
		                      response.Write("<td>" & i & "</td>") ' №
                              response.Write("<td style='text-align:left; min-width:225px;'>" & students(i,2) & "</td>") ' Фамилия
        if not sep_empty then response.Write("<td>" & students(i,3) & "</td>") ' Сентябрь
        if not oct_empty then response.Write("<td>" & students(i,4) & "</td>") ' Октябрь
        if not nov_empty then response.Write("<td>" & students(i,5) & "</td>") ' Ноябрь
        if not dec_empty then response.Write("<td>" & students(i,6) & "</td>") ' Декабрь
                              response.Write("<td style='font-weight: bold;'>" & students(i,7) & "</td>") ' 1 сем
        if not jan_empty then response.Write("<td>" & students(i,8) & "</td>") ' Январь
        if not feb_empty then response.Write("<td>" & students(i,9) & "</td>") ' Февраль
        if not mar_empty then response.Write("<td>" & students(i,10) & "</td>")' Март
        if not apr_empty then response.Write("<td>" & students(i,11) & "</td>")' Апрель
                              response.Write("<td style='font-weight: bold;'>" & students(i,12) & "</td>")' 2 сем
                              response.Write("<td style='font-weight: bold;'>" & students(i,13) & "</td>")' Всего
		
		response.Write("</tr>")
	end if
next
response.Write("<tr>")
                      response.Write("<td colspan=2 style='text-align: center; font-weight: bold;'>Итого:</td>")
if not sep_empty then response.Write("<td>" & total_sep & "</td>")
if not oct_empty then response.Write("<td>" & total_oct & "</td>")
if not nov_empty then response.Write("<td>" & total_nov & "</td>")
if not dec_empty then response.Write("<td>" & total_dec & "</td>")
                      response.Write("<td style='font-weight: bold;'>" & total_sem1 & "</td>")
if not jan_empty then response.Write("<td>" & total_jan & "</td>")
if not feb_empty then response.Write("<td>" & total_feb & "</td>")
if not mar_empty then response.Write("<td>" & total_mar & "</td>")
if not apr_empty then response.Write("<td>" & total_apr & "</td>")
                      response.Write("<td style='font-weight: bold;'>" & total_sem2 & "</td>")
                      response.Write("<td style='font-weight: bold;'>" & total_total & "</td>")
response.Write("</tr>")
%>
</tbody>
</table>
</div>
<div id=notPrintableArea>
<% else %>
<br>
<h2><b>Во время выполнения запроса произошла ошибка!</b></h1>
<h4 style="width: 65%">В базе данных не были обнаружены записи о студентах, либо их слишком мало!</h4>
<% end if %>
<%
elseif query = 0 then
today_input = mid(date(), 7, 4) + "-" + mid(date(), 4, 2) + "-" + mid(date(), 1, 2)
%>
<form action="ved_propusk_zan.asp?query=1" method="post">
<center>
<h4 style="width: 75%;">Вы собираетесь сформировать ведомость пропусков занятий (по неуважительным причинам) по группам в разрезе месяцев и семестров.</h4>
<h5>Выберие группу и скоректируйте дату, если нужно, и нужно ли вам отображать пустые колонки.</h5>
<table class="form-table" style="width: 400px;">
    <tr>
        <td align=right>Группа:  </td>
        <td>
            <div class="input-control text" style="width: 150px">
            <select name="groupSearch" required>
                <option disabled selected value>Выберите...</option>
		            <% 
                    Set rs_group = Server.CreateObject("ADODB.RecordSet")
                    Set con = Server.CreateObject("ADODB.Connection")
                    strdbpath=server.mappath("base.mdb")
                    con.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strdbpath
                    sql_group = "SELECT id_group, group_name FROM tbl_group ORDER BY group_name;"
                    rs_group.Open sql_group, con, 3, 3
                    Set objId   = rs_group.Fields("id_group")
                    Set objName = rs_group.Fields("group_name")
                    do until rs_group.EOF
                    %>
                <option value="<% response.write(objId) %>"><% response.write(objName) %></option> 
                    <%
                    rs_group.MoveNext
                    loop
                    rs_group.Close 
                    con.Close
                    %>
	        </select>
            </div>
        </td>
    </tr>
    <tr>
        <td align=right>Дата:  </td>
        <td>
            <div class="input-control text" style="width: 150px">
	            <input type="date" name="date" value="<%= today_input %>" style="font-family: 'Segoe UI', 'Open Sans', sans-serif, serif; font-size: 0.875rem;">
            </div>
        </td>
    </tr>
    <tr style="height:49px;" title="Переключает отображение пустых колонок, т.е. месяцев в которых отсутсвуют полностью отметки о пропусках по неуважительной причине.">
        <td align=right>Пустые колонки:  </td>
        <td>
            <label class="switch-original">
                <input type="checkbox" name="empty-cols" checked>
                <span class="check"></span>
            </label>
        </td>
    </tr>
</table>

<script>
    function ShowHideMiniLoader(operation){
        if (operation == "hide"){
            document.getElementById("mini_loader").style.display = "none";
        } else {
            document.getElementById("mini_loader").style.display = "block";    
        }
    }
</script>
<button type=submit class="button primary">Подтвердить</button> <a href="group_change.asp"></a>
</center>
</form>
<%
end if
%>
</center>
<br>
<center>

<a href="group_change.asp?go=1"><button type="button" class="button subinfo">Вернуться к выбору группы</button><br>
<a href="help/02_12.asp" ><button class="button success"><span class="icon mif-info"></span> Помощь</button></a>
<a href="exit.asp"><button class="button danger" ><span class="icon mif-exit"></span> Выход</button></a>
</center>
</td>
</tr>
</tbody>
</table>
	  
<br>
</td>
</tr>
</table>
</div>
</body>
</html>