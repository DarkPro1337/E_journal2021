<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01//EN" "http://www.w3.org/TR/html4/strict.dtd">
<% 
'---------------------------------------------------
'Защита от SQL-инъекции 
'---------------------------------------------------
all= request.QueryString
for i=1 to len(all)
if len(all)-6 > 0 then
if mid(all,i,6)="SELECT" or mid(all,i,6)="INSERT" or mid(all,i,6)="DELETE" or mid(all,i,5)="UNION" then 
response.redirect("enter_error.html")
end if
end if
if len (all)-3 > 0 then
if mid(all,i,3)="AND" or mid(all,i,3)="XOR" then
response.Redirect ("enter_error.html")
end if
end if
if mid(all,i,1)=";" then response.Redirect ("enter_error.html")
next
'------------------------------------------------------
on error resume next
id = request.QueryString ("id_stud")
Set Conn = Server.CreateObject("ADODB.Connection") 
Set RS1 = Server.CreateObject("ADODB.Recordset")
Set RS2 = Server.CreateObject("ADODB.Recordset")
Set RS3 = Server.CreateObject("ADODB.Recordset")
Set RS4 = Server.CreateObject("ADODB.Recordset")
Set RS5 = Server.CreateObject("ADODB.Recordset")
Set RS6 = Server.CreateObject("ADODB.Recordset")
'Set RS_1 = Server.CreateObject("ADODB.Recordset")
'Set RS_2 = Server.CreateObject("ADODB.Recordset")
strDBPath = Server.MapPath("base.mdb")
Conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath
'strSQL1 = "SELECT tbl_student.student_fio, tbl_student.student_number FROM  tbl_student WHERE (((tbl_student.id_student)="&id&"))"

'strSQL2 = "SELECT Count(tbl_journal.sobytie)*2 AS [Count-sobytie] FROM tbl_student INNER JOIN tbl_journal ON tbl_student.id_student = tbl_journal.student_name WHERE (((tbl_student.id_student)="&id&") AND ((tbl_journal.sobytie)=6))"

'strSQL2 = "SELECT Count(tbl_journal.sobytie)*2 AS [Count-sobytie] FROM tbl_student INNER JOIN tbl_journal ON tbl_student.id_student = tbl_journal.student_name WHERE (((tbl_student.id_student)="&id&") AND ((tbl_journal.sobytie)=6))"
'strSQL3 = "SELECT Avg(tbl_journal.sobytie) AS [Avg-sobytie], tbl_disc.disc_name FROM (tbl_disc INNER JOIN tbl_plan ON tbl_disc.ID_disc = tbl_plan.disc_name) INNER JOIN (tbl_zan INNER JOIN (tbl_student INNER JOIN tbl_journal ON tbl_student.id_student = tbl_journal.student_name) ON tbl_zan.id_zan = tbl_journal.zan) ON tbl_plan.ID_plan = tbl_zan.disc_name WHERE (((tbl_student.id_student)="&id&") AND ((tbl_journal.sobytie)<6)) GROUP BY tbl_disc.disc_name"
'strSQL4 = "SELECT tbl_journal.sobytie AS [sobytie], tbl_disc.disc_name FROM (tbl_disc INNER JOIN tbl_plan ON tbl_disc.ID_disc = tbl_plan.disc_name) INNER JOIN (tbl_zan INNER JOIN (tbl_student INNER JOIN tbl_journal ON tbl_student.id_student = tbl_journal.student_name) ON tbl_zan.id_zan = tbl_journal.zan) ON tbl_plan.ID_plan = tbl_zan.disc_name WHERE ((tbl_student.id_student)="&id&")"
'strSQL5 = "SELECT tbl_journal.sobytie AS sobytie, tbl_disc.disc_name FROM ((tbl_disc INNER JOIN tbl_plan ON tbl_disc.ID_disc = tbl_plan.disc_name) INNER JOIN tbl_zan ON tbl_plan.ID_plan = tbl_zan.disc_name) INNER JOIN (tbl_student INNER JOIN tbl_journal ON tbl_student.id_student = tbl_journal.student_name) ON tbl_zan.id_zan = tbl_journal.zan WHERE ((tbl_student.id_student)="&id&")"
'response.write 1
sql1 = "SELECT tbl_student.student_fio, tbl_student.student_number FROM  tbl_student WHERE (((tbl_student.id_student)="&id&"))"
RS1.Open sql1, Conn, 3, 3
'response.write 2
sql2 = "SELECT Count(tbl_journal.nb2)*2 AS [Count-sobytie] FROM tbl_student INNER JOIN tbl_journal ON tbl_student.id_student = tbl_journal.student_name WHERE (((tbl_student.id_student)="&id&") AND ((tbl_journal.nb2)=true))"
RS2.Open sql3, Conn, 3, 3
'response.write 3
sql3 = "SELECT Count(tbl_journal.nb1) AS [Count-sobytie] FROM tbl_student INNER JOIN tbl_journal ON tbl_student.id_student = tbl_journal.student_name WHERE (((tbl_student.id_student)="&id&") AND ((tbl_journal.nb1)=true))"
RS3.Open sql3, Conn, 3, 3
'response.write 4
'count_passes = RS2.fields(0) + RS3.fields(0)
'response.write 5
sql4 = "SELECT Avg(tbl_journal.sobytie) AS [Avg-sobytie], tbl_disc.disc_name FROM (tbl_disc INNER JOIN tbl_plan ON tbl_disc.ID_disc = tbl_plan.disc_name) INNER JOIN (tbl_zan INNER JOIN (tbl_student INNER JOIN tbl_journal ON tbl_student.id_student = tbl_journal.student_name) ON tbl_zan.id_zan = tbl_journal.zan) ON tbl_plan.ID_plan = tbl_zan.disc_name WHERE (((tbl_student.id_student)="&id&") AND ((tbl_journal.sobytie)<6) AND ((tbl_plan.Semestr)='" & session("sem") & "')) GROUP BY tbl_disc.disc_name"
RS4.Open sql4, Conn, 3, 3
'response.write 6
sql5 = "SELECT tbl_journal.sobytie, tbl_disc.disc_name, tbl_journal.sobytie_old FROM (tbl_disc INNER JOIN tbl_plan ON tbl_disc.ID_disc = tbl_plan.disc_name) INNER JOIN (tbl_zan INNER JOIN (tbl_student INNER JOIN tbl_journal ON tbl_student.id_student = tbl_journal.student_name) ON tbl_zan.id_zan = tbl_journal.zan) ON tbl_plan.ID_plan = tbl_zan.disc_name WHERE ((tbl_student.id_student)="&id&") AND ((tbl_plan.Semestr)='" & session("sem") & "')"
RS5.Open sql5, Conn, 3, 3
'response.write 7
sql6 = "SELECT tbl_journal.sobytie, tbl_disc.disc_name, tbl_journal.nb1, tbl_journal.nb2, tbl_journal.sobytie_old FROM ((tbl_disc INNER JOIN tbl_plan ON tbl_disc.ID_disc = tbl_plan.disc_name) INNER JOIN tbl_zan ON tbl_plan.ID_plan = tbl_zan.disc_name) INNER JOIN (tbl_student INNER JOIN tbl_journal ON tbl_student.id_student = tbl_journal.student_name) ON tbl_zan.id_zan = tbl_journal.zan WHERE ((tbl_student.id_student)="&id&") AND ((tbl_plan.Semestr)='" & session("sem") & "')"
RS6.Open sql6, Conn, 3, 3
'response.write 8
'RS debug
'response.write (RS6.recordcount&"-RS6 ")
'response.write (RS5.recordcount&"-RS5 ")
'response.write (RS4.recordcount&"-RS4 ")
'response.write (RS3.recordcount&"-RS3 ")
'response.write (RS2.recordcount&"-RS2 ")
'response.write (RS1.recordcount&"-RS1 ")






%>
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01//EN" "http://www.w3.org/TR/html4/strict.dtd">
<html>
	<head>
		<meta content="text/html; charset=Windows-1251" http-equiv="content-type">
		<link rel="shortcut icon" href="images/favicon.ico" /> 
		<link rel="stylesheet" href="css/metro.css">
		<link rel="stylesheet" href="css/metro-colors.css">
		<link rel="stylesheet" href="css/metro-icons.css">
		<link rel="stylesheet" href="css/metro-responsive.css">
		<link rel="stylesheet" href="css/metro-rtl.css">
		<link rel="stylesheet" href="css/metro-schemes.css">
		<link rel="stylesheet" href="css/metro-student.css">
		<link rel="stylesheet" href="css/button.css">
		<script src="js/jquery-3.1.0.min.js"></script>
		<script src="js/metro.min.js"></script>
		<title>ИС "Электронный журнал". Данные о студенте <%=rs1.Fields(0)%></title>
		<style>
			title {
				display: block !important;
				color: #f00;
				font-size: 20px;
				padding: 14px;
			}
		</style>
	</head>
	
	<body>
	
	
		<body>
	
		<table class="table border" style='width:90%; margin-top: 15px;' align=center> 
			
			<tr>
			
				<td>
<%
'rs2.Open strSQL2, Conn, adOpenStatic
'rs3.Open strSQL3, Conn, 3, 3
'rs_1.Open strSQL4, Conn, 3, 3
'rs_2.Open strSQL5, Conn, 3, 3

dl_per=500 'Объявление переменной dl_per (максимальное количество элементов на одной строке)
old_dl=dl_per
%>
<!-- #include file="header.asp" -->
<!-- #include file="pass_check.asp" -->
<%
output_rating = false
output_usp_and_kach = true

'----------------------------------------------------
'Подключение БД
'----------------------------------------------------
Set Conn = Server.CreateObject("ADODB.Connection") 
Set rs_perpod = Server.CreateObject("ADODB.Recordset") 
Set rs_student = Server.CreateObject("ADODB.Recordset")
Set rs_rating_q = Server.CreateObject("ADODB.Recordset")
Set rs_rating_dop = Server.CreateObject("ADODB.Recordset")
strDBPath = Server.MapPath("base.mdb")
on error resume next
Conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath
if not err = 0 then response.redirect("bd_error.asp")
dim current_operation
current_operation = ""
current_operation = current_operation & request.form("select_operation")
strSQL_student = "SELECT tbl_student.id_student, tbl_student.student_fio, tbl_student.student_nam, tbl_student.student_otch, tbl_group.group_name, tbl_spec.id_spec, tbl_spec.spec_name FROM tbl_spec INNER JOIN (tbl_group INNER JOIN tbl_student ON tbl_group.id_group = tbl_student.group_name) ON tbl_spec.id_spec = tbl_group.spec WHERE ((tbl_student.id_student) = " + id + ")"
rs_student.Open strSQL_student, Conn, adOpenStatic
dim student_inf()
ReDim student_inf(11)
student_inf(0) = rs_student.fields(0) 'id студента
for i_sym = 1 to len(rs_student.fields(1))
	if mid(rs_student.fields(1), i_sym, 1) = " " then
		if not i_sym = 1 then
			student_inf(1) = mid(rs_student.fields(1), 1, i_sym - 1) 'фамилия студента
			exit for
		end if
	end if
next
student_inf(2) = rs_student.fields(2) 'имя студента
student_inf(3) = rs_student.fields(3) 'отчество студента
student_inf(4) = rs_student.fields(4) 'группа студента
student_inf(5) = rs_student.fields(5) 'id специальности студента
student_inf(6) = rs_student.fields(6) 'специальность студента
student_inf(7) = 0 'Место в рейтинге учащегося среди группы
student_inf(8) = 0 'Место в рейтинге учащегося среди специальности
student_inf(9) = 0 'Место в рейтинге учащегося среди всех учащихся
student_inf(10) = 0 'Успеваемость учащегося
student_inf(11) = 0 'Качество успеваемости учащегося
rs_student.Close
select case student_inf(5)
	case 2 'Эксплуатация транспортного электрооборудования и средств автоматики
		student_status = "cadet"
		student_string = "Курсант"
	case 3 'Эксплуатация внутренних водных путей
		student_status = "cadet"
		student_string = "Курсант"
	case 4 'Судовождение
		student_status = "cadet"
		student_string = "Курсант"
	case 5 'Экономика и бухгалтерский учет
		student_status = "student"
		student_string = "Студент"
	case 8 'Техническое обслуживание и ремонт автомобильного транспорта
		student_status = "student"
		student_string = "Студент"
	case 10 'Информационные системы
		student_status = "student"
		student_string = "Студент"
	case 11 'Судовождение (программа углубленной подготовки)
		student_status = "cadet"
		student_string = "Курсант"
end select
%>
<h2 style="text-align:center;"><%response.Write(student_string & " " &  rs1.Fields(0)& " ") 
response.Write ("группа "&session("gr"))%></h2>
<%
'Подсчёт места в рейтинге
if output_rating = true then
dim rating(), rating_gr()
for i_rating = 1 to 3
	select case i_rating
		case 1
			strSQL_rating = "SELECT id_student, student_fio, tbl_group.group_name, tbl_student.rating_old, tbl_student.student_number FROM tbl_group INNER JOIN tbl_student ON tbl_group.id_group = tbl_student.group_name WHERE tbl_group.group_name = '" & student_inf(4) & "' ORDER BY student_fio"
		case 2
			strSQL_rating = "SELECT id_student, student_fio, tbl_group.group_name, tbl_student.rating_old, tbl_student.student_number, tbl_spec.id_spec FROM tbl_spec INNER JOIN (tbl_group INNER JOIN tbl_student ON tbl_group.id_group = tbl_student.group_name) ON tbl_spec.id_spec = tbl_group.spec WHERE tbl_spec.id_spec = " & student_inf(5) & " ORDER BY student_fio"
		case 3
			strSQL_rating = "SELECT id_student, student_fio, tbl_group.group_name, tbl_student.rating_old, tbl_student.student_number FROM tbl_group INNER JOIN tbl_student ON tbl_group.id_group = tbl_student.group_name ORDER BY student_fio"
	end select
	rs_rating_q.open strSQL_rating, Conn, 3
	strSqlF="SELECT * FROM tbl_group"
	rs_rating_dop.open strSqlF, Conn, 3, 3
	Erase rating
	Erase rating_gr
	n_stud = rs_rating_q.RecordCount
	n_gr = rs_rating_dop.recordcount
	rs_rating_dop.close
	redim rating (n_stud, 7), rating_gr (n_gr,9)
	for i = 1 to n_stud
		rating(i,6) = rs_rating_q.Fields(0)' id студента
		rating(i,1) = rs_rating_q.Fields(1)' ФИО студента
		rating(i,7) = rs_rating_q.Fields(4)' Студ.номер
		rating(i,2) = rs_rating_q.Fields(2)' Группа
		rating(i,3) = rs_rating_q.Fields(3)' Рейтинг за прошлый год
		rating(i,4) = 0
		rating(i,5) = 0
		rs_rating_q.MoveNext
	next
	rs_rating_q.Close
	rs_rating_q.open "sum_rat_pos_new", Conn, 3
	n_pos = rs_rating_q.RecordCount
	for i=1 to n_pos
		stud_id = rs_rating_q.Fields(3)
		for j=1 to n_stud
			if  rating (j,6)= stud_id then rating(j,4) = rs_rating_q.Fields(2) 'Рейтинг посещаемости
		next
		rs_rating_q.MoveNext
	next
	rs_rating_q.close
	rs_rating_q.open "sum_rating_usp", Conn, 3
	n_usp = rs_rating_q.RecordCount
	for i=1 to n_usp
		stud_id = rs_rating_q.Fields(5)
		for j=1 to n_stud
			if  rating (j,6)= stud_id then rating(j,5) = rs_rating_q.Fields(4) 'Рейтинг успеваемости
		next
		rs_rating_q.MoveNext
	next
	rs_rating_q.close
	'-------------------------
	'Сортировка
	'-------------------------
	for k=1 to n_stud
		for i=1 to n_stud - 1
			if (rating(i,3)+rating(i,4)+rating(i,5)) < (rating(i+1,3)+rating(i+1,4)+rating(i+1,5)) then
				for j=1 to 7
					z=rating(i,j)
					rating(i,j)=rating(i+1,j)
					rating(i+1,j) = z
				next
			end if
		next
	next
	for i=1 to n_stud
		if rating(i,6) = session("id_student") then
			select case i_rating
				case 1
					student_inf(7) = i
					n_stud_group = n_stud
				case 2
					student_inf(8) = i
					n_stud_spec = n_stud
				case 3
					student_inf(9) = i
					n_stud_all = n_stud
			end select
			exit for
		end if
	next
next
end if

'Создание массива с оценками и подсчёт качества с успеваемостью
Set rs_disc = Server.CreateObject("ADODB.RecordSet")
Set rs_count_disc = Server.CreateObject("ADODB.RecordSet")
Set rs_ocen = Server.CreateObject("ADODB.RecordSet")
Set rs_prop = Server.CreateObject("ADODB.RecordSet")
Set rs_kontr = Server.CreateObject("ADODB.RecordSet")
Set rs_itog = Server.CreateObject("ADODB.RecordSet")
dim current_semestr_old
dim current_semestr
current_semestr = session("sem")
current_semestr = current_semestr & request.form("select_semestr")
current_semestr_old = ""
current_semestr_old = current_semestr_old & request.form("select_semestr_old")
if current_semestr = "" or current_operation = "reset" then
    if Month(Date())=>9 then
        current_semestr_old = 1
		current_semestr = 1
    else
        current_semestr_old = 2
		current_semestr = 2
    end if
end if
dim current_data_first
dim current_data_second
current_data_first = ""
current_data_second = ""
current_data_first = current_data_first & request.form("select_data_first")
current_data_second = current_data_second & request.form("select_data_second")
if (current_data_first = "" or current_data_second = "") or (not current_semestr = current_semestr_old) or current_operation = "reset" then
    onpost = false
	if current_semestr = 1 then
        if month(now()) >= 9 and month(now()) <= 12 then
            current_data_first = "09/01/" & Year(now)
            current_data_second = month(now()) & "/" & day(now()) & "/" & Year(now)
			current_day_first = 1
			current_month_first = 9
			current_year_first = Year(now)
			current_day_second = day(now())
			current_month_second = month(now())
			current_year_second = Year(now)
        else
            current_data_first = "09/01/" & Year(now) - 1
            current_data_second = "12/31/" & Year(now) - 1
			current_day_first = 1
			current_month_first = 9
			current_year_first = Year(now) - 1
			current_day_second = 31
			current_month_second = 12
			current_year_second = Year(now) - 1
        end if
    else
        if month(now()) >= 9 and month(now()) <= 12 then
            current_data_first = "01/01/" & Year(now) + 1
            current_data_second = "08/31/" & Year(now) + 1
			current_day_first = 1
			current_month_first = 1
			current_year_first = Year(now) + 1
			current_day_second = 31
			current_month_second = 8
			current_year_second = Year(now) + 1
        else
            current_data_first = "01/01/" & Year(now)
            current_data_second = month(now()) & "/" & day(now()) & "/" & Year(now)
			current_day_first = 1
			current_month_first = 1
			current_year_first = Year(now)
			current_day_second = day(now())
			current_month_second = month(now())
			current_year_second = Year(now)
        end if
    end if
	select case current_month_first
		case 1
			current_month_first_ru = "января"
		case 2
			current_month_first_ru = "февраля"
		case 3
			current_month_first_ru = "марта"
		case 4
			current_month_first_ru = "апреля"
		case 5
			current_month_first_ru = "мая"
		case 6
			current_month_first_ru = "июня"
		case 7
			current_month_first_ru = "июля"
		case 8
			current_month_first_ru = "августа"
		case 9
			current_month_first_ru = "сентября"
		case 10
			current_month_first_ru = "октября"
		case 11
			current_month_first_ru = "ноября"
		case 12
			current_month_first_ru = "декабря"
	end select
	select case current_month_second
		case 1
			current_month_second_ru = "января"
		case 2
			current_month_second_ru = "февраля"
		case 3
			current_month_second_ru = "марта"
		case 4
			current_month_second_ru = "апреля"
		case 5
			current_month_second_ru = "мая"
		case 6
			current_month_second_ru = "июня"
		case 7
			current_month_second_ru = "июля"
		case 8
			current_month_second_ru = "августа"
		case 9
			current_month_second_ru = "сентября"
		case 10
			current_month_second_ru = "октября"
		case 11
			current_month_second_ru = "ноября"
		case 12
			current_month_second_ru = "декабря"
	end select
	current_data_first_ru = current_day_first & " " & current_month_first_ru & " " & current_year_first & " г."
	current_data_second_ru = current_day_second & " " & current_month_second_ru & " " & current_year_second & " г."
else
	onpost = true
	current_data_first_ru = current_data_first
	current_data_second_ru = current_data_second
	for i_sym = 1 to len(current_data_first)
		if mid(current_data_first, i_sym, 1) = " " then
			for j_sym = i_sym + 1 to len(current_data_first)
				if mid(current_data_first, j_sym, 1) = " " then
					current_day_first = mid(current_data_first, 1, i_sym - 1)
					select case mid(current_data_first, i_sym + 1, j_sym - len(current_day_first) - 2)
						case "января"
							current_month_first = 1
						case "февраля"
							current_month_first = 2
						case "марта"
							current_month_first = 3
						case "апреля"
							current_month_first = 4
						case "мая"
							current_month_first = 5
						case "июня"
							current_month_first = 6
						case "июля"
							current_month_first = 7
						case "августа"
							current_month_first = 8
						case "сентября"
							current_month_first = 9
						case "октября"
							current_month_first = 10
						case "ноября"
							current_month_first = 11
						case "декабря"
							current_month_first = 12
					end select
					current_year_first = mid(current_data_first, j_sym + 1, 4)
					current_data_first = current_month_first & "/" & current_day_first & "/" & current_year_first
					exit for
				end if
			next
			exit for
		end if
	next
	for i_sym = 1 to len(current_data_second)
		if mid(current_data_second, i_sym, 1) = " " then
			for j_sym = i_sym + 1 to len(current_data_second)
				if mid(current_data_second, j_sym, 1) = " " then
					current_day_second = mid(current_data_second, 1, i_sym - 1)
					select case mid(current_data_second, i_sym + 1, j_sym - len(current_day_second) - 2)
						case "января"
							current_month_second = 1
						case "февраля"
							current_month_second = 2
						case "марта"
							current_month_second = 3
						case "апреля"
							current_month_second = 4
						case "мая"
							current_month_second = 5
						case "июня"
							current_month_second = 6
						case "июля"
							current_month_second = 7
						case "августа"
							current_month_second = 8
						case "сентября"
							current_month_second = 9
						case "октября"
							current_month_second = 10
						case "ноября"
							current_month_second = 11
						case "декабря"
							current_month_second = 12
					end select
					current_year_second = mid(current_data_second, j_sym + 1, 4)
					current_data_second = current_month_second & "/" & current_day_second & "/" & current_year_second
					exit for
				end if
			next
			exit for
		end if
	next
end if
'Рекордсет с дисциплинами по группе и по семестру
sql_disc="SELECT tbl_disc.disc_name, tbl_group.group_name, tbl_plan.Semestr, tbl_disc.disc_ab, tbl_plan.kol_chas, tbl_plan.id_plan, tbl_user.user_fio, tbl_plan.control_form FROM tbl_user INNER JOIN (tbl_group INNER JOIN (tbl_disc INNER JOIN tbl_plan ON tbl_disc.ID_disc = tbl_plan.disc_name) ON tbl_group.id_group = tbl_plan.gr_name) ON tbl_user.id_user = tbl_plan.Prepod_name WHERE (((tbl_group.group_name)='" & student_inf(4) & "') AND ((tbl_plan.Semestr)='" & current_semestr & "')) ORDER BY tbl_disc.disc_name"
'---------------------------------------------------------------------------------------------------------------------------
'Заполнение массива информацией о дисциплинах (id, Название, Аббревиатура, Преподаватель, кол-во часов, выдано часов, успеваемость, качество, количество обязательных работ)
'---------------------------------------------------------------------------------------------------------------------------
rs_disc.open sql_disc, conn, 3
dim disc()
	n_disc=rs_disc.RecordCount
redim disc(n_disc,19)
for i=1 to n_disc
	disc(i,1)=rs_disc.Fields(5)'id plan
	disc(i,2)=rs_disc.Fields(0)'дисциплина
	disc(i,3)=rs_disc.Fields(3)'сок. дисциплина
	disc(i,4)=rs_disc.Fields(6)'преподаватель
	disc(i,5)=rs_disc.Fields(4)'кол-во часов
	disc(i,15)=rs_disc.Fields(7)'форма контроля
	disc(i,10)= "" 'средний балл учащегося
	disc(i,11)= "" 'итоговая оценка учащегося
	disc(i,12)= 0 'количетство пропусков по дисциплине
	disc(i,13)= 0 'количетство уважительных пропусков по дисциплине
	disc(i,14)= 0 'количетство неужавительных пропусков по дисциплине
	disc(i,17)= 0 'Качество знаний учащегося
	disc(i,18)= 0 'Качество знаний группы
	disc(i,19)= 0 'Успеваемость группы
	sql_count_disc="SELECT Count([tbl_zan]![disc_name])*2 AS CountDisc FROM tbl_plan INNER JOIN tbl_zan ON tbl_plan.ID_plan = tbl_zan.disc_name GROUP BY tbl_plan.ID_plan HAVING (((tbl_plan.ID_plan)="& disc(i,1) & "))"
	rs_count_disc.open sql_count_disc, conn, 3
	if rs_count_disc.EOF<>true and rs_count_disc.BOF<>true then disc(i,6)=rs_count_disc.Fields(0) else disc(i,6)=0
	rs_count_disc.Close
	sql_count_disc="SELECT Count([tbl_zan]![disc_name])*2 AS CountDisc, tbl_plan.ID_plan, Count(tbl_zan.kontr) AS [Count-kontr] FROM tbl_plan INNER JOIN tbl_zan ON tbl_plan.ID_plan = tbl_zan.disc_name WHERE (((tbl_zan.kontr)=True)) GROUP BY tbl_plan.ID_plan HAVING (((tbl_plan.ID_plan)="& disc(i,1) & "))"
	rs_count_disc.open sql_count_disc, conn, 3
	if rs_count_disc.EOF<>true and rs_count_disc.BOF<>true then disc(i,9)=rs_count_disc.Fields(2) else disc(i,9)=0
	rs_count_disc.Close
	rs_disc.MoveNext
next
rs_disc.Close
'---------------------------------------------------------------------------------------------------------------------------
'Заполнение массива информацией об оценках
'---------------------------------------------------------------------------------------------------------------------------
dim avg_ocen()
redim avg_ocen(n_disc,3)
count_2 = 0
count_3 = 0
count_4 = 0
count_5 = 0
for j=1 to n_disc
	flag_itog = 0
	'---------------------------
	'Определение - есть ли итоговая оценка по дисциплине
	'---------------------------
	strSQL_itog = "SELECT tbl_itog.id_itog, tbl_itog.itog_disc, tbl_itog.itog_stud, tbl_itog.itog_oc, tbl_itog.Author FROM tbl_itog WHERE (((tbl_itog.itog_disc)="&disc(j,1)&") AND ((tbl_itog.itog_stud)="& student_inf(0) &"))"
	rs_itog.Open strSQL_itog, conn, 3
	if rs_itog.eof <> true and rs_itog.BOF <> true then 
		avg_ocen(j,1)= rs_itog.Fields(3)
		avg_ocen(j,3) = rs_itog.Fields(3)
	end if
	flag_itog = 1
	strSQL_ocen = "SELECT tbl_student.id_student, tbl_plan.ID_plan, Avg(tbl_journal.sobytie) AS [Avg-sobytie] FROM (tbl_plan INNER JOIN tbl_zan ON tbl_plan.ID_plan = tbl_zan.disc_name) INNER JOIN (tbl_student INNER JOIN tbl_journal ON tbl_student.id_student = tbl_journal.student_name) ON tbl_zan.id_zan = tbl_journal.zan WHERE (((tbl_journal.sobytie)<6)) GROUP BY tbl_student.id_student, tbl_plan.ID_plan HAVING (((tbl_student.id_student)="& student_inf(0) &") AND ((tbl_plan.ID_plan)="&disc(j,1)&"))"
	strSQL_prop = "SELECT tbl_student.id_student, tbl_plan.ID_plan, (Count(tbl_journal.sobytie))*2 AS [Count-sobytie] FROM (tbl_plan INNER JOIN tbl_zan ON tbl_plan.ID_plan = tbl_zan.disc_name) INNER JOIN (tbl_student INNER JOIN tbl_journal ON tbl_student.id_student = tbl_journal.student_name) ON tbl_zan.id_zan = tbl_journal.zan WHERE (((tbl_journal.sobytie)=6)) GROUP BY tbl_student.id_student, tbl_plan.ID_plan HAVING (((tbl_student.id_student)="& student_inf(0) &") AND ((tbl_plan.ID_plan)="& disc(j,1) &"));"
	rs_prop.Open strSQL_prop, conn, 3
	rs_ocen.Open strSQL_ocen, conn, 3
		flag=0
		if disc(j,9)>0 then
			strSQL_kontr = "SELECT tbl_student.id_student, tbl_plan.ID_plan, tbl_journal.sobytie FROM (tbl_plan INNER JOIN tbl_zan ON tbl_plan.ID_plan = tbl_zan.disc_name) INNER JOIN (tbl_student INNER JOIN tbl_journal ON tbl_student.id_student = tbl_journal.student_name) ON tbl_zan.id_zan = tbl_journal.zan WHERE (((tbl_student.id_student)="& student_inf(0) &") AND ((tbl_plan.ID_plan)="&disc(j,1)&") AND ((tbl_journal.sobytie)<6) AND ((tbl_zan.kontr)=True))"
			rs_kontr.Open strSQL_kontr, conn, 3
			for k=1 to disc(j,9)
				if rs_kontr.bof = true or rs_kontr.eof = true then flag=1 else if rs_kontr.fields(2)=6 or rs_kontr.fields(2)<2.5 then flag = 1
				if rs_kontr.bof<>true and rs_kontr.eof <>true then rs_kontr.movenext
			next
			rs_kontr.Close  
		end if  
		if flag=1 then
			if avg_ocen(j,1) = "" then avg_ocen (j,1) = "н/а"
			if flag_itog = 1 then avg_ocen(j,3) = avg_ocen(j,3) & " (н/а)" else avg_ocen(j,3) = "н/а"
			avg_ocen(j,2) = "Не выполнены обязательные практические или контрольные работы"
		else
			if avg_ocen(j,3) = "" then avg_ocen(j,3) = round(rs_ocen.Fields (2),1)
			if avg_ocen(j,1) = "" then avg_ocen (j,1) = round(rs_ocen.Fields (2),1)
		end if
		rs_ocen.close
		rs_prop.close
		rs_itog.Close
		'---вывод зачета
		if avg_ocen(j,1) = "7" then
			avg_ocen(j,3)= "зач"
		else
			if avg_ocen(j,3)="" then 
				avg_ocen(j,3)= "&nbsp;"
			else
				if avg_ocen(j,1) = "н/а" then
					count_2=count_2 + 1
				else
					if avg_ocen(j,1)<=2.5 then count_2 = count_2 + 1
					if avg_ocen(j,1) >2.5 and avg_ocen(j,1)<3.5 then count_3 = count_3 + 1
					if avg_ocen(j,1)>=3.5 and avg_ocen(j,1)<4.5 then count_4 = count_4 + 1
					if avg_ocen(j,1)>=4.5 then count_5 = count_5 + 1
				end if
			end if
	end if
next
student_inf(10) = round((count_3 + count_4 + count_5)/(count_2 + count_3 + count_4 + count_5)*100,0) 'Успеваемость учащегося
student_inf(11) = round((count_4 + count_5)/(count_2 + count_3 + count_4 + count_5)*100,0) 'Качество успеваемости учащегося

'Расчёт качества и успеваемости учащихся группы
strSQL_student="SELECT student_fio, tbl_student.id_student, tbl_student.student_number FROM tbl_group INNER JOIN tbl_student ON tbl_group.id_group = tbl_student.group_name WHERE ((tbl_group.group_name)='" & student_inf(4) &"') ORDER BY student_fio"
'---------------------------------------------------------------------------------------------------------------------------
'Заполнение массива фамилиями студентов текущей группы (id, Фамилия И.О., количество 2, 3, 4, 5, Успеваемость, Качество, цвет заливки ячеек в ведомости)
'---------------------------------------------------------------------------------------------------------------------------
rs_student.open strSQL_student, conn, 3
dim student()
n_stud = cint(rs_student.recordcount)
redim student(n_stud,3)
dim avg_ocen_group()
redim avg_ocen_group(n_stud,n_disc,1)
dim sum_prop()
redim sum_prop(n_stud,n_disc)
for i=1 to n_stud
	student(i,1)=rs_student.Fields(1)'id
	student(i,2)=rs_student.Fields(0)'FIO
	student(i,10)=rs_student.Fields(2)
	rs_student.MoveNext
next
rs_student.Close
for i=1 to n_stud
	count_2 = 0
	count_3 = 0
	count_4 = 0
	count_5 = 0
	for j=1 to n_disc
		'---------------------------
		'Определение - есть ли итоговая оценка по дисциплине
		'---------------------------
		strSQL_itog = "SELECT tbl_itog.id_itog, tbl_itog.itog_disc, tbl_itog.itog_stud, tbl_itog.itog_oc, tbl_itog.Author FROM tbl_itog WHERE (((tbl_itog.itog_disc)="&disc(j,1)&") AND ((tbl_itog.itog_stud)="&student(i,1)&"))"
		rs_itog.Open strSQL_itog, conn, 3
		if rs_itog.eof <> true and rs_itog.BOF <> true then
			avg_ocen_group(i,j,1)= rs_itog.Fields(3)
		else
			strSQL_ocen = "SELECT tbl_student.id_student, tbl_plan.ID_plan, Avg(tbl_journal.sobytie) AS [Avg-sobytie] FROM (tbl_plan INNER JOIN tbl_zan ON tbl_plan.ID_plan = tbl_zan.disc_name) INNER JOIN (tbl_student INNER JOIN tbl_journal ON tbl_student.id_student = tbl_journal.student_name) ON tbl_zan.id_zan = tbl_journal.zan WHERE (((tbl_journal.sobytie)<6) AND ((tbl_zan.date) Between #" & current_data_first & "# AND #" & current_data_second & "#)) GROUP BY tbl_student.id_student, tbl_plan.ID_plan HAVING (((tbl_student.id_student)="&student(i,1)&") AND ((tbl_plan.ID_plan)="&disc(j,1)&"))"
			strSQL_prop = "SELECT tbl_student.id_student, tbl_plan.ID_plan, (Count(tbl_journal.sobytie))*2 AS [Count-sobytie] FROM (tbl_plan INNER JOIN tbl_zan ON tbl_plan.ID_plan = tbl_zan.disc_name) INNER JOIN (tbl_student INNER JOIN tbl_journal ON tbl_student.id_student = tbl_journal.student_name) ON tbl_zan.id_zan = tbl_journal.zan WHERE (((tbl_journal.sobytie)=6) AND ((tbl_zan.date) Between #" & current_data_first & "# AND #" & current_data_second & "#)) GROUP BY tbl_student.id_student, tbl_plan.ID_plan HAVING (((tbl_student.id_student)="& student(i,1) &") AND ((tbl_plan.ID_plan)="& disc(j,1) &"));"
			rs_prop.Open strSQL_prop, conn, 3
			rs_ocen.Open strSQL_ocen, conn, 3
			if rs_prop.BOF<>true and rs_prop.EOF <>true then sum_prop (i,j) = rs_prop.Fields(2) else sum_prop (i,j)=0
				flag=0
				if disc(j,9)>0 then
					strSQL_kontr = "SELECT tbl_student.id_student, tbl_plan.ID_plan, tbl_journal.sobytie FROM (tbl_plan INNER JOIN tbl_zan ON tbl_plan.ID_plan = tbl_zan.disc_name) INNER JOIN (tbl_student INNER JOIN tbl_journal ON tbl_student.id_student = tbl_journal.student_name) ON tbl_zan.id_zan = tbl_journal.zan WHERE (((tbl_student.id_student)="&student(i,1)&") AND ((tbl_plan.ID_plan)="&disc(j,1)&") AND ((tbl_journal.sobytie)<6) AND ((tbl_zan.kontr)=True) AND ((tbl_zan.date) Between #" & current_data_first & "# AND #" & current_data_second & "#))"
					rs_kontr.Open strSQL_kontr, conn, 3
					for k=1 to disc(j,9)
						if rs_kontr.bof = true or rs_kontr.eof = true then flag=1 else if rs_kontr.fields(2)=6 or rs_kontr.fields(2)<=2 then flag = 1
						if rs_kontr.bof<>true and rs_kontr.eof <>true then rs_kontr.movenext
					next
					rs_kontr.Close  
				end if  
				if flag=1 then
					avg_ocen_group(i,j,1) = "н/а"
				else
					if rs_ocen.BOF<>true and rs_ocen.EOF <>true then 
						if sum_prop(i,j)>=disc(j,5)*0.3 then 
							avg_ocen_group(i,j,1) = "н/а" 
						else 
							avg_ocen_group(i,j,1)=round(rs_ocen.Fields (2),1)
						end if
					end if
				end if
			rs_ocen.close
			rs_prop.close
		end if
		rs_itog.Close
		if avg_ocen_group(i,j,1)="" then 
			avg_ocen_group(i,j,1)= "&nbsp;"
		else
			if avg_ocen_group(i,j,1) = "н/а" then
				count_2=count_2 + 1
			else
				if avg_ocen_group(i,j,1)<=2.5 then count_2 = count_2 + 1
				if avg_ocen_group(i,j,1) >2.5 and avg_ocen_group(i,j,1)<3.5 then count_3 = count_3 + 1
				if avg_ocen_group(i,j,1)>=3.5 and avg_ocen_group(i,j,1)<4.5 then count_4 = count_4 + 1
				if avg_ocen_group(i,j,1)>=4.5 then count_5 = count_5 + 1
			end if
		end if
	next
	if count_2 > 0 then neusp=neusp+1
	if student(i,8)= "100%" then usp=usp+1
next
for j=1 to n_disc
	d_count_2=0
	d_count_3=0
	d_count_4=0
	d_count_5=0
	for i=1 to n_stud
		if avg_ocen_group(i,j,1)<>"&nbsp;" then 
			if avg_ocen_group(i,j,1) = "н/а" then
				d_count_2 = d_count_2 + 1
			else
				if avg_ocen_group(i,j,1)<=2.5  or avg_ocen_group(i,j,1)="н/а" then d_count_2 = d_count_2 + 1
				if avg_ocen_group(i,j,1) > 2.5 and avg_ocen_group(i,j,1)<=3.49 then d_count_3 = d_count_3 + 1
				if avg_ocen_group(i,j,1)>=3.5 and avg_ocen_group(i,j,1)<=4.49 then d_count_4 = d_count_4 + 1
				if avg_ocen_group(i,j,1)>=4.5 then d_count_5 = d_count_5 + 1
			end if
		end if
	next
	disc(j,19) = round(((d_count_3 + d_count_4 + d_count_5)/(d_count_2 + d_count_3 + d_count_4 + d_count_5))*100,0)
	disc(j,18) = round(((d_count_4 + d_count_5)/(d_count_2 + d_count_3 + d_count_4 + d_count_5))*100,0)
next

'Формирование массива со всеми оценками
Set rs_avg = Server.CreateObject("ADODB.Recordset")
Set rs_sob = Server.CreateObject("ADODB.Recordset")
sql_sob = "SELECT tbl_journal.sobytie, tbl_disc.disc_name, tbl_journal.nb1, tbl_journal.nb2, tbl_journal.nb_type, tbl_journal.nb_comment, tbl_zan.date, tbl_zan.prim_student, tbl_zan.zan_type, tbl_journal.sobytie_old FROM ((tbl_disc INNER JOIN tbl_plan ON tbl_disc.ID_disc = tbl_plan.disc_name) INNER JOIN tbl_zan ON tbl_plan.ID_plan = tbl_zan.disc_name) INNER JOIN (tbl_student INNER JOIN tbl_journal ON tbl_student.id_student = tbl_journal.student_name) ON tbl_zan.id_zan = tbl_journal.zan WHERE (((tbl_student.id_student)="& student_inf(0) &") AND ((tbl_plan.Semestr)='" & current_semestr & "') AND ((tbl_zan.date) Between #" & current_data_first & "# AND #" & current_data_second & "#)) ORDER BY tbl_journal.edit_date ASC"
rs_sob.Open sql_sob, Conn, 3
dim sob()
n_sob = rs_sob.recordcount
redim sob(n_disc, n_sob, 6)
for i_disc = 1 to n_disc
	strSQL_itog = "SELECT tbl_itog.id_itog, tbl_itog.itog_disc, tbl_itog.itog_stud, tbl_itog.itog_oc, tbl_itog.Author FROM tbl_itog WHERE (((tbl_itog.itog_disc)="& disc(i_disc,1) &") AND ((tbl_itog.itog_stud)="& student_inf(0) &"))"
	rs_itog.Open strSQL_itog, conn, 3
	if rs_itog.eof <> true and rs_itog.BOF <> true then
		if disc(i_disc,9) = "Зачет" then
			if rs_itog.Fields(3) >= 3 then
				disc(i_disc,11) = "зач"
			else
				disc(i_disc,11) = "н/з"
			end if
		else
			if rs_itog.Fields(3) = 1 then disc(i_disc,11) = "н/а" else disc(i_disc,11) = rs_itog.Fields(3)
		end if
	end if
	rs_itog.Close
	sql_avg = "SELECT Avg(tbl_journal.sobytie) AS [Avg-sobytie], tbl_disc.disc_name FROM (tbl_disc INNER JOIN tbl_plan ON tbl_disc.ID_disc = tbl_plan.disc_name) INNER JOIN (tbl_zan INNER JOIN (tbl_student INNER JOIN tbl_journal ON tbl_student.id_student = tbl_journal.student_name) ON tbl_zan.id_zan = tbl_journal.zan) ON tbl_plan.ID_plan = tbl_zan.disc_name WHERE (((tbl_student.id_student)="& student_inf(0) &") AND ((tbl_journal.sobytie)<6) AND ((tbl_plan.Semestr)='" & current_semestr & "') AND ((tbl_plan.ID_plan) = "& disc(i_disc,1) &") AND ((tbl_zan.date) Between #" & current_data_first & "# AND #" & current_data_second & "#)) GROUP BY tbl_disc.disc_name"
	rs_avg.Open sql_avg, Conn, 3
	rs_avg.movefirst

	'if rs_avg.recordcount > 0 then
		disc(i_disc,10) = round(rs_avg.Fields(0),2) 'Средний балл
		rs_avg.close
		k_nb = 0
		k_nb_uvag = 0
		k_nb_neuvag = 0
		count_2 = 0
		count_3 = 0
		count_4 = 0
		count_5 = 0
		for i_sob=1 to rs_sob.recordcount
			if rs_sob.Fields(1) = disc(i_disc, 2) then 
				if rs_sob.Fields(3)=true then  
					sob(i_disc, i_sob, 6) = "2нб"
					sob(i_disc, i_sob, 1) = rs_sob.Fields(4) 'Тип нб
					sob(i_disc, i_sob, 2) = rs_sob.Fields(5) 'Описание нб
					sob(i_disc, i_sob, 3) = rs_sob.Fields(6) 'Дата занятия
					k_nb = k_nb + 2
					select case sob(i_disc, i_sob, 1)
						case 1, 2, 3, 4
							k_nb_uvag = k_nb_uvag + 2
						case 5, 6, 7
							k_nb_neuvag = k_nb_neuvag + 2
					end select
					if not rs_sob.Fields(9) = "" and not rs_sob.Fields(0) = "6" and not rs_sob.Fields(0) = "8" then
						sob(i_disc, i_sob, 0) = round(rs_sob.Fields(0),2)
						if sob(i_disc, i_sob, 0) < 2.50 then count_2 = count_2 + 1
						if sob(i_disc, i_sob, 0) >= 2.50 and sob(i_disc, i_sob, 0) < 3.50 then count_3 = count_3 + 1
						if sob(i_disc, i_sob, 0) >= 3.50 and sob(i_disc, i_sob, 0) < 4.50 then count_4 = count_4 + 1
						if sob(i_disc, i_sob, 0) >= 4.50 and sob(i_disc, i_sob, 0) < 6 then count_5 = count_5 + 1
					end if
				elseif  rs_sob.Fields(2)=true then
					sob(i_disc, i_sob, 6) = "1нб"
					sob(i_disc, i_sob, 1) = rs_sob.Fields(4) 'Тип нб
					sob(i_disc, i_sob, 2) = rs_sob.Fields(5) 'Описание нб
					sob(i_disc, i_sob, 3) = rs_sob.Fields(6) 'Дата занятия
					k_nb = k_nb + 1
					select case sob(i_disc, i_sob, 1)
						case 1, 2, 3, 4
							k_nb_uvag = k_nb_uvag + 1
						case 5, 6, 7
							k_nb_neuvag = k_nb_neuvag + 1
					end select
					if not rs_sob.Fields(9) = "" and not rs_sob.Fields(0) = "6" and not rs_sob.Fields(0) = "8" then
						sob(i_disc, i_sob, 0) = round(rs_sob.Fields(0),2)
						if sob(i_disc, i_sob, 0) < 2.50 then count_2 = count_2 + 1
						if sob(i_disc, i_sob, 0) >= 2.50 and sob(i_disc, i_sob, 0) < 3.50 then count_3 = count_3 + 1
						if sob(i_disc, i_sob, 0) >= 3.50 and sob(i_disc, i_sob, 0) < 4.50 then count_4 = count_4 + 1
						if sob(i_disc, i_sob, 0) >= 4.50 and sob(i_disc, i_sob, 0) < 6 then count_5 = count_5 + 1
					end if
				else
					if rs_sob.Fields(0) = 7 then
						sob(i_disc, i_sob, 0) = "зач"
						sob(i_disc, i_sob, 3) = rs_sob.Fields(6) 'Дата занятия
					else
						sob(i_disc, i_sob, 0) = round(rs_sob.Fields(0),2) 'Оценка
						sob(i_disc, i_sob, 3) = rs_sob.Fields(6) 'Дата занятия
						sob(i_disc, i_sob, 4) = rs_sob.Fields(7) 'Описание занятия
						sob(i_disc, i_sob, 5) = rs_sob.Fields(8) 'Тип занятия
						if sob(i_disc, i_sob, 0) < 2.50 then count_2 = count_2 + 1
						if sob(i_disc, i_sob, 0) >= 2.50 and sob(i_disc, i_sob, 0) < 3.50 then count_3 = count_3 + 1
						if sob(i_disc, i_sob, 0) >= 3.50 and sob(i_disc, i_sob, 0) < 4.50 then count_4 = count_4 + 1
						if sob(i_disc, i_sob, 0) >= 4.50 and sob(i_disc, i_sob, 0) < 6 then count_5 = count_5 + 1
					end if
				end if
			end if
			rs_sob.movenext
		next
		disc(i_disc,12) = k_nb 'Количетсво нб по дисциплине
		disc(i_disc,13) = k_nb_uvag 'Количетсво уважительных нб по дисциплине
		disc(i_disc,14) = k_nb_neuvag 'Количетсво неуважительных нб по дисциплине
		disc(i_disc,17) = round(((count_4 + count_5)/(count_2 + count_3 + count_4 + count_5))*100,0)
		rs_sob.movefirst
	'end if
next

'Формирование массива с задолженностями
Set rs_zadol = Server.CreateObject("ADODB.RecordSet")
'---------------------------------------------------------------------------------------------------------------------------
'Заполнение массива информацией о задолженностях 
'---------------------------------------------------------------------------------------------------------------------------
strSQL_zadol = "SELECT tbl_zan.date, tbl_zan.id_zan, tbl_zan.kontr, tbl_zan.prim, tbl_user.user_fio, tbl_zan.prim_student, tbl_zan.zan_type, tbl_disc.disc_name, tbl_disc.ID_disc FROM tbl_user INNER JOIN ((tbl_group INNER JOIN (tbl_disc INNER JOIN tbl_plan ON tbl_disc.ID_disc = tbl_plan.disc_name) ON tbl_group.id_group = tbl_plan.gr_name) INNER JOIN tbl_zan ON tbl_plan.ID_plan = tbl_zan.disc_name) ON tbl_user.id_user = tbl_plan.Prepod_name WHERE (((tbl_plan.Semestr)='" & current_semestr & "') AND ((tbl_group.group_name)='" & student_inf(4) & "') and ((tbl_zan.zan_type) between 2 and 4) AND ((tbl_zan.date) Between #" & current_data_first & "# AND #" & current_data_second & "#)) ORDER BY tbl_disc.disc_name, tbl_zan.date;" ' and ((tbl_zan.kontr) = true)
rs_zadol.Open strSQL_zadol, conn, 3
dim dat()
n_dat = cint(rs_zadol.recordcount)
redim dat(10, n_dat)
for i=1 to n_dat
	dat(1, i) = rs_zadol.Fields (1) 'id_zan
	dat(2, i) = left(cstr(rs_zadol.Fields (0)),5) 'Дата проведения занятия
	dat(3, i) = rs_zadol.Fields (2) 'Контрольная работа
	dat(5, i) = rs_zadol.Fields (3) 'Примечание
	dat(6, i) = rs_zadol.Fields (4) 'Автор
	dat(7, i) = rs_zadol.Fields (5) 'Примечание студент
	dat(8, i) = rs_zadol.Fields (6) 'Тип занятия
	dat(9, i) = rs_zadol.Fields (7) 'Название занятия
	dat(10, i) = rs_zadol.Fields (8) 'id
	rs_zadol.MoveNext
next
rs_zadol.Close
'---------------------------------------------------------------------------------------------------------------------------
'Заполнение массива информацией об оценках
'---------------------------------------------------------------------------------------------------------------------------	
strSQL_ocen = "SELECT tbl_group.group_name, tbl_plan.Semestr, tbl_disc.disc_name, tbl_student.id_student, tbl_journal.sobytie, tbl_zan.id_zan, tbl_plan.ID_plan, tbl_user.user_fio, tbl_journal.edit, tbl_journal.edit_date FROM tbl_user INNER JOIN ((((tbl_group INNER JOIN (tbl_disc INNER JOIN tbl_plan ON tbl_disc.ID_disc = tbl_plan.disc_name) ON tbl_group.id_group = tbl_plan.gr_name) INNER JOIN tbl_student ON tbl_group.id_group = tbl_student.group_name) INNER JOIN tbl_zan ON tbl_plan.ID_plan = tbl_zan.disc_name) INNER JOIN tbl_journal ON (tbl_zan.id_zan = tbl_journal.zan) AND (tbl_student.id_student = tbl_journal.student_name)) ON tbl_user.id_user = tbl_journal.Author WHERE (((tbl_group.group_name)='" & student_inf(4) &"') AND ((tbl_plan.Semestr)='" & current_semestr &"') AND ((tbl_zan.date) Between #" & current_data_first & "# AND #" & current_data_second & "#));"
rs_ocen.Open strSQL_ocen, conn, 3
dim sob_dolg()
redim sob_dolg(n_dat, 3)
while not rs_ocen.EOF
	for j=1 to n_dat
		if rs_ocen.Fields(3) = student_inf(0) and rs_ocen.fields(5) = dat (1, j) then 
			sob_dolg(j, 0) = rs_ocen.fields(4)
		end if
	next
	rs_ocen.MoveNext
Wend
rs_ocen.close
%>
<table class="table striped hovered cell-hovered border bordered">
	<thead>
		<tr>
			<th class="sortable-column" style="min-width: 380px;">Дисциплина</th>
			<th class="sortable-column">Оценки</th>
		</tr>
	</thead>
	<tbody>
		<%
		for i_disc = 1 to n_disc
			sob_vid_name = "Теоретическая работа"
			response.write("<tr><td class='sortable-column'>" & disc(i_disc, 2) & "</td><td>")
			for i_sob = 1 to n_sob
				if not sob(i_disc, i_sob, 0) = "" then
					if not sob(i_disc, i_sob, 4) = "" then
						select case sob(i_disc, i_sob, 5)
							case 1
								sob_vid_name = "Теоретическая работа"
							case 2
								sob_vid_name = "Практическая работа"
							case 3
								sob_vid_name = "Тест"
							case 4
								sob_vid_name = "Лабораторная работа"
							case 5
								sob_vid_name = "Самостоятельная работа"
						end select
					end if
					if sob(i_disc, i_sob, 0) = "1нб" or sob(i_disc, i_sob, 0) = "2нб" then
					elseif sob(i_disc, i_sob, 0) < 2.5 then
						%><div class='assess two' title='Дата: <%=sob(i_disc, i_sob, 3)%><%if not sob_vid_name = "" then response.write("&#10;Вид работы: " & sob_vid_name)%><%if not sob(i_disc, i_sob, 4) = "" then response.write("&#10;Описание: " & sob(i_disc, i_sob, 4))%>'><%=sob(i_disc, i_sob, 0)%></div><%
					elseif sob(i_disc, i_sob, 0) >= 2.5 and sob(i_disc, i_sob, 0) < 3.5 then
						%><div class='assess three' title='Дата: <%=sob(i_disc, i_sob, 3)%><%if not sob_vid_name = "" then response.write("&#10;Вид работы: " & sob_vid_name)%><%if not sob(i_disc, i_sob, 4) = "" then response.write("&#10;Описание: " & sob(i_disc, i_sob, 4))%>'><%=sob(i_disc, i_sob, 0)%></div><%
					elseif sob(i_disc, i_sob, 0) >= 3.5 and sob(i_disc, i_sob, 0) < 4.5 then
						%><div class='assess four' title='Дата: <%=sob(i_disc, i_sob, 3)%><%if not sob_vid_name = "" then response.write("&#10;Вид работы: " & sob_vid_name)%><%if not sob(i_disc, i_sob, 4) = "" then response.write("&#10;Описание: " & sob(i_disc, i_sob, 4))%>'><%=sob(i_disc, i_sob, 0)%></div><%
					elseif sob(i_disc, i_sob, 0) >= 4.5 then
						%><div class='assess five' title='Дата: <%=sob(i_disc, i_sob, 3)%><%if not sob_vid_name = "" then response.write("&#10;Вид работы: " & sob_vid_name)%><%if not sob(i_disc, i_sob, 4) = "" then response.write("&#10;Описание: " & sob(i_disc, i_sob, 4))%>'><%=sob(i_disc, i_sob, 0)%></div><%
					end if
				end if
				if not sob(i_disc, i_sob, 6) = "" then
					if sob(i_disc, i_sob, 6) = "1нб" or sob(i_disc, i_sob, 6) = "2нб" then
						select case sob(i_disc, i_sob, 1)
							case 1, 2, 3, 4
								select case sob(i_disc, i_sob, 1)
									case 1
										nb_type_name = "Больничный"
									case 2
										nb_type_name = "Дежурство"
									case 3
										nb_type_name = "Освобождение (заявление)"
									case 4
										nb_type_name = "Другое (уважительная)"
								end select
								%><div class='assess uvnb' title='Дата: <%=sob(i_disc, i_sob, 3)%>&#10;Причина: <%=nb_type_name%><%if not sob(i_disc, i_sob, 2) = "" and not sob(i_disc, i_sob, 2) = "-" then response.write("&#10;Описание: " & sob(i_disc, i_sob, 2))%>'><%=sob(i_disc, i_sob, 6)%></div><%
							case 5, 6, 7
								select case sob(i_disc, i_sob, 1)
									case 5
										nb_type_name = "Прогул"
									case 6
										nb_type_name = "Удалён с занятия"
									case 7
										nb_type_name = "Другое (не уважительная)"
								end select
								%><div class='assess nb' title='Дата: <%=sob(i_disc, i_sob, 3)%>&#10;Причина: <%=nb_type_name%><%if not sob(i_disc, i_sob, 2) = "" and not sob(i_disc, i_sob, 2) = "-" then response.write("&#10;Описание: " & sob(i_disc, i_sob, 2))%>'><%=sob(i_disc, i_sob, 6)%></div><%
							case else
								nb_type_name = "Не указана"
								%><div class='assess nullnb' title='Дата: <%=sob(i_disc, i_sob, 3)%>&#10;Причина: <%=nb_type_name%><%if not sob(i_disc, i_sob, 2) = "" and not sob(i_disc, i_sob, 2) = "-" then response.write("&#10;Описание: " & sob(i_disc, i_sob, 2))%>'><%=sob(i_disc, i_sob, 6)%></div><%
						end select
					end if
				end if
			next
			response.write("</td></tr>")
		next
		%>
	</tbody>
</table>
            <br>
			<center>
			<a href="disc_change.asp?go=1" class="button subinfo" >Вернуться к выбору дисциплин и ведомостей</a><br>
			<a href="group_change.asp?go=1" class="button subinfo" >Вернуться к выбору группы</a><br>
			<a href="help/03_9.asp" id="help" class="icon mif-info"> Помощь</a>
			<a href="exit.asp" id="exit" class="mif-exit"> Выход</a>
			</center>
            </td>
          </tr>
        </tbody>
      </table>
	  
      <br>
      </td>
    </tr>
</table>
</body>
</html>
