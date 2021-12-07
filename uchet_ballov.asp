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
	<title>ИС "Электронный журнал". Учёт поощирительных баллов</title>
    <style>
        summary 
        {
            border-color: #2086bf;
            font-size:larger;
            -webkit-touch-callout: none; /* iOS Safari */
            -webkit-user-select: none; /* Safari */
            -khtml-user-select: none; /* Konqueror HTML */
            -moz-user-select: none; /* Old versions of Firefox */
            -ms-user-select: none; /* Internet Explorer/Edge */
            user-select: none;
            padding: 0.3rem 1rem;
            height: 2.125rem;
            text-align: center;
            background-color: #ffffff;
            border: 1px #d9d9d9 solid;
            color: #262626;
            cursor: pointer;
            display: inline-block;
            outline: none;
            font-size: .875rem;
            margin: .15625rem 0;
            position: relative;
            vertical-align: middle;
        }
        details[open] summary ~ * {
          animation: sweep .5s ease-in-out;
        }

        @keyframes sweep {
          0%    {opacity: 0; margin-top: -10px}
          100%  {opacity: 1; margin-top: 0px}
        }
        
    </style>
</head>
<body>
<body>
<table class="table border" style='width:90%; margin-top: 15px;' align=center> 	
<tr>
<td>
    <!-- #include file="header.asp" -->
    <!-- #include file="pass_check.asp" -->
<%
    if session("user") = "" or session("user") = "Студент" or session("user") = 0 then response.Redirect ("404.asp")
    if request.QueryString("addNew") = 1 then ' Форма на добавление записи в БД
    group = session("gr")%>
    <center>
    <a href="uchet_ballov.asp" ><button class="button subinfo"><span class="icon mif-undo"></span> Вернуться назад</button></a>
    <h3>Добавление записи о поощрениях к группе <%=group%></h3>
    <%
    'Выполняем подключение к БД
    Set con = Server.CreateObject("ADODB.Connection")
    Set rs  = Server.CreateObject("ADODB.RecordSet")
    strdbpath=server.mappath("base.mdb")
    con.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strdbpath
    %>
    <form name="addNew" action="?addNew=2" method="post" style="margin-bottom: 0">
	    <%
        strSQL = "SELECT TOP 1 tbl_group.id_group, tbl_group.group_name FROM tbl_group WHERE (((tbl_group.group_name)='" + group + "'));"
        rs.Open strSQL, con
        %>
        <input type="hidden" value="<%=rs.Fields(0)%>" name="gr" />
        <%
        rs.Close
        %>
        <div class="input-control text" style="width: 150px">
	    <select name="student" required>
            <option disabled selected value>Студент</option>
		    <%
            strSQL = "SELECT tbl_group.id_group, tbl_group.group_name, tbl_student.id_student, tbl_student.student_fio, IIf(IsNull([score_value]),0,[score_value]) AS score_val FROM tbl_group INNER JOIN (tbl_student LEFT JOIN tbl_score ON tbl_student.id_student = tbl_score.score_student) ON tbl_group.id_group = tbl_student.group_name WHERE (((tbl_group.group_name)='" + group + "')) ORDER BY tbl_student.student_fio;"
            rs.Open strSQL, con
            set objId = rs.Fields(2)
            set objName = rs.Fields(3)
            set objValue = rs.Fields(4)
            do until rs.EOF %>
           <option value="<%response.write(objId)%>"><%response.write(objName)%> (<%response.Write(objValue)%>)</option> 
            <% rs.MoveNext
               loop
               rs.Close %>
	    </select>
        </div>
        <div class="input-control text" style="width: 80px">
            <input name="value" type="number" step="0.25" placeholder="Баллы" required>
        </div>
        <br />
        <div class="input-control text" style="width: 235px">
        <textarea name="comment" rows="10" cols="45" placeholder="Комментарий" maxlength="254"></textarea>
	    </div>
        <br /><br /><br /><br /><br />
	    <button type="submit" class="button primary"><span class="icon mif-pencil"></span>  Добавить запись</button> <button type="reset" class="button danger"><span class="icon mif-undo"></span>  Сброс</button>
        <br /><br />
    </form>
    </center>
<%
    elseif request.QueryString("addNew") = 2 then ' Запрос на добавление записи в БД
    group = request.Form("gr")
    stud  = request.Form("student")
    value = request.Form("value")
    comment = request.Form("comment")
    prepod = session("user")

    today = Now()
    dd = mid(today, 1, 2)
    mm = mid(today, 4, 2)
    yyyy = mid(today, 7, 4)
    nowDate = mm + "/" + dd + "/" + yyyy

    'Создаем подключение для запросов
    strdbpath=server.mappath("base.mdb")
    strDbConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="& strdbpath & ";"
    Set objConn = Server.CreateObject("ADODB.Connection")
    objConn.Open(strDbConnection)
    strSQL = "INSERT INTO tbl_score (score_student, score_value, score_author, score_date, score_comment) VALUES (" & stud & ", " & value & ", " & prepod & ", #" & nowDate & "#, '" & comment & "');" 'Подготавливаем запрос
    'Выполняем запрос
	objConn.Execute(strSQL)
    response.Redirect("uchet_ballov.asp")

    elseif request.QueryString("edit") > 0 then ' Форма на редактирование комментария записи в БД
    record = request.QueryString("edit")%>
    <center>
    <button class="button subinfo" onClick="window.history.go(-1)"><span class="icon mif-undo"></span> Вернуться назад</button>
    <h3>Редактирования комментария записи №<%=record%></h3>
    <%
    'Выполняем подключение к БД
    Set con = Server.CreateObject("ADODB.Connection")
    Set rs  = Server.CreateObject("ADODB.RecordSet")
    strdbpath=server.mappath("base.mdb")
    con.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strdbpath
    %>
    <form name="editExec" action="?editExec=1" method="post" style="margin-bottom: 0">
	    <%
        strSQL = "SELECT tbl_score.id_score, tbl_student.id_student, tbl_student.student_fio, tbl_group.id_group, tbl_group.group_name, tbl_score.score_value, tbl_score.score_comment FROM tbl_group INNER JOIN (tbl_student INNER JOIN tbl_score ON tbl_student.id_student = tbl_score.score_student) ON tbl_group.id_group = tbl_student.group_name WHERE (((tbl_score.id_score)=" + record + "));"
        rs.Open strSQL, con
        %>
        <input name="record" type="hidden" value="<%=record%>" />
        <div class="input-control text" style="width: 150px">
	        <input name="student" value="<%=rs.Fields("student_fio")%>" readonly disabled />
        </div>
        <div class="input-control text" style="width: 80px">
            <input name="value" value="<%=rs.Fields("score_value")%>" type="text" readonly disabled />
        </div>
        <br />
        <div class="input-control text" style="width: 235px">
        <textarea name="comment" rows="10" cols="45" placeholder="Комментарий" maxlength="254" required><%=rs.Fields("score_comment")%></textarea>
	    </div>
        <%
        rs.Close
        %>
        <br /><br /><br /><br /><br />
	    <button type="submit" class="button primary"><span class="icon mif-pencil"></span>  Редактировать запись</button> <button type="reset" class="button danger"><span class="icon mif-undo"></span>  Сброс</button>
    </form>
    </center>
    <br />
<%
    elseif request.QueryString("editExec") = 1 then ' Запрос на редактирование комментария в БД
    record = request.Form("record")
    comment = request.Form("comment")

    'Создаем подключение для запросов
    strdbpath=server.mappath("base.mdb")
    strDbConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="& strdbpath & ";"
    Set objConn = Server.CreateObject("ADODB.Connection")
    objConn.Open(strDbConnection)
    strSQL = "UPDATE tbl_score SET score_comment = '" + comment + "' WHERE id_score = " + record + ";" 'Подготавливаем запрос
    'Выполняем запрос
	objConn.Execute(strSQL)
    response.Redirect("uchet_ballov.asp?history=1")

    elseif request.QueryString("history") = 1 then ' Просмотр истории записей из БД (без фильтрации)
    group = session("gr")%>
    <center>
    <a href="uchet_ballov.asp" ><button class="button subinfo"><span class="icon mif-undo"></span> Вернуться назад</button></a>
    <h3>Просмотр истории записей о поощрениях группы <%=group%></h3>

    <%
    'Выполняем подключение к БД
    Set con = Server.CreateObject("ADODB.Connection")
    Set rs  = Server.CreateObject("ADODB.RecordSet")
    Set rs2  = Server.CreateObject("ADODB.RecordSet")
    strdbpath=server.mappath("base.mdb")
    con.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strdbpath

    strSQL = "SELECT tbl_score.id_score, tbl_group.id_group, tbl_group.group_name, tbl_student.id_student, tbl_student.student_fio, tbl_score.score_value, tbl_user.user_fio, tbl_score.score_date, tbl_score.score_comment FROM tbl_group INNER JOIN (tbl_user INNER JOIN (tbl_student INNER JOIN tbl_score ON tbl_student.id_student = tbl_score.score_student) ON tbl_user.id_user = tbl_score.score_author) ON tbl_group.id_group = tbl_student.group_name WHERE (((tbl_group.group_name)='" + group + "')) ORDER BY tbl_score.score_date DESC;"
    rs.Open strSQL, con, 3, 3

    dim history(5000,6) 'Массив на 5000 записей истории
    if rs.RecordCount > 0 then
	    ij = 1
	    cnt = 1
	    while rs.EOF <> true
		    history(ij,1) = rs.Fields("id_score")      'ID
		    history(ij,2) = rs.Fields("student_fio")   'ФИО Студента
		    history(ij,3) = rs.Fields("score_value")   'Баллы
            history(ij,4) = rs.Fields("user_fio")      'ФИО Преподавателя
            history(ij,5) = rs.Fields("score_comment") 'Комментарий
            history(ij,6) = rs.Fields("score_date")    'Дата
            ij = ij + 1
            cnt = cnt + 1
            rs.MoveNext
	    wend
    end if
    if cnt > 0 then
    %>
    <form name="filter" action="uchet_ballov.asp?history=2" method="post" style="margin-bottom: 0">
    Фильтрация по студенту
	    <div class="input-control text" style="width: 185px">
	    <select name="studentFilter" required>
            <option disabled selected value>Выберите...</option>
		    <% rs2.Open "SELECT tbl_group.id_group, tbl_group.group_name, tbl_student.id_student, tbl_student.student_fio FROM tbl_group INNER JOIN tbl_student ON tbl_group.id_group = tbl_student.group_name WHERE (((tbl_group.group_name)='" + group + "')) ORDER BY tbl_student.student_fio;", con
             set objId = rs2.Fields("id_student")
             set objName = rs2.Fields("student_fio")
             do until rs2.EOF %>
           <option value="<% response.write(objId) %>" <% if cstr(objName) = cstr(group) then response.Write(" selected") %>><% response.write(objName) %></option> 
            <% rs2.MoveNext
               loop
               rs2.Close %>
	    </select>
	    </div>
        <button type="submit" class="button primary"><span class="icon mif-filter"></span>  Выбрать</button> <a href="?history=1"><button type="button" class="button danger"><span class="icon mif-undo"></span>  Сброс</button></a><br><br>
    </form>
    <table class="table striped hovered cell-hovered border bordered" style="width: 100%">
    <thead align=center style="font-weight: bold">
    <tr><th>ID</th><th>Студент</th><th>Баллы</th><th>Преподаватель</th><th>Комментарий</th><th>Дата</th><th>Редактирование</th></tr>
    </thead>
    <tbody align=center>
    <%
        for i = 1 to cnt 'Запрос выполняется для каждого студента
	        if history(i,1) > 0 then 'Если пользователь в массиве имеет свой ID
		        'Подготавливаем данные
                if history(i,3) < 0 then
                response.Write("<tr style='background-color: #4390df;color: white;'>")
                elseif history(i,3) > 0 then
                response.Write("<tr style='background-color: #60a917;color: white;'>")
                else
                response.Write("<tr>")
                end if
                
		        'Рисуем таблицу
		        response.Write("<td style='width:1px'>" & history(i,1) & "</td>") 'ID
                response.Write("<td style='width:25%'>" & history(i,2) & "</td>") 'ФИО Студента
                response.Write("<td style='width:1px'>" & history(i,3) & "</td>") 'Баллы
                response.Write("<td style='width:1px'>" & history(i,4) & "</td>") 'ФИО Преподавателя
                response.Write("<td style='width:50%'>" & history(i,5) & "</td>") 'Комментарий
                response.Write("<td style='width:1px'>" & history(i,6) & "</td>") 'Дата
                response.Write("<td style='width:15%'><a href='?edit=" & history(i,1) & "'><button class='button subinfo'><span class='icon mif-pencil'></span> Изменить</button></a></td>") 'Кнопка редактирования

		        response.Write("</tr>")
	        end if
        next
    %>
    </tbody>
    </table>
    <% elseif cnt <= 0 or cnt = "" then %>
    <br>
    <h4><b>Записей не найдено!</b></h4>
    <% end if %>
    </center>
    <br />
<%
    elseif request.QueryString("history") = 2 then ' Просмотр истории записей из БД (без фильтрации)
    group = session("gr")%>
    <center>
    <a href="uchet_ballov.asp" ><button class="button subinfo"><span class="icon mif-undo"></span> Вернуться назад</button></a>
    <h3>Просмотр истории записей о поощрениях студента группы <%=group%></h3>

    <%
    'Выполняем подключение к БД
    Set con = Server.CreateObject("ADODB.Connection")
    Set rs  = Server.CreateObject("ADODB.RecordSet")
    Set rs2  = Server.CreateObject("ADODB.RecordSet")
    strdbpath=server.mappath("base.mdb")
    con.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strdbpath
    studentFilter = request.Form("studentFilter")
    strSQL = "SELECT tbl_score.id_score, tbl_group.id_group, tbl_group.group_name, tbl_student.id_student, tbl_student.student_fio, tbl_score.score_value, tbl_user.user_fio, tbl_score.score_date, tbl_score.score_comment FROM tbl_group INNER JOIN (tbl_user INNER JOIN (tbl_student INNER JOIN tbl_score ON tbl_student.id_student = tbl_score.score_student) ON tbl_user.id_user = tbl_score.score_author) ON tbl_group.id_group = tbl_student.group_name WHERE (((tbl_group.group_name)='" + group + "') AND ((tbl_student.id_student)=" + studentFilter + ")) ORDER BY tbl_score.score_date DESC;"
    rs.Open strSQL, con, 3, 3
    %>
    <form name="filter" action="uchet_ballov.asp?history=2" method="post" style="margin-bottom: 0">
    Фильтрация по студенту
	    <div class="input-control text" style="width: 185px">
	    <select name="studentFilter" required>
            <option disabled selected value>Выберите...</option>
		    <% rs2.Open "SELECT tbl_group.id_group, tbl_group.group_name, tbl_student.id_student, tbl_student.student_fio FROM tbl_group INNER JOIN tbl_student ON tbl_group.id_group = tbl_student.group_name WHERE (((tbl_group.group_name)='" + group + "')) ORDER BY tbl_student.student_fio;", con
             set objId = rs2.Fields("id_student")
             set objName = rs2.Fields("student_fio")
             do until rs2.EOF %>
           <option value="<% response.write(objId) %>" <% if cstr(objId) = cstr(studentFilter) then response.Write(" selected") %>><% response.write(objName) %></option> 
            <% rs2.MoveNext
               loop
               rs2.Close %>
	    </select>
	    </div>
        <button type="submit" class="button primary"><span class="icon mif-filter"></span>  Выбрать</button> <a href="?history=1"><button type="button" class="button danger"><span class="icon mif-undo"></span>  Сброс</button></a><br><br>
    </form>
    <%
    dim history_filtered(3000,6) 'Массив на 5000 записей истории
    if rs.RecordCount > 0 then
	    ij = 1
	    cnt = 1
	    while rs.EOF <> true
		    history_filtered(ij,1) = rs.Fields("id_score")      'ID
		    history_filtered(ij,2) = rs.Fields("student_fio")   'ФИО Студента
		    history_filtered(ij,3) = rs.Fields("score_value")   'Баллы
            history_filtered(ij,4) = rs.Fields("user_fio")      'ФИО Преподавателя
            history_filtered(ij,5) = rs.Fields("score_comment") 'Комментарий
            history_filtered(ij,6) = rs.Fields("score_date")    'Дата
            ij = ij + 1
            cnt = cnt + 1
            rs.MoveNext
	    wend
    end if
    if cnt > 0 then
    %>
    <table class="table striped hovered cell-hovered border bordered" style="width: 100%">
    <thead align=center style="font-weight: bold">
    <tr><th>ID</th><th>Студент</th><th>Баллы</th><th>Преподаватель</th><th>Комментарий</th><th>Дата</th><th>Редактировать</th></tr>
    </thead>
    <tbody align=center>
    <%
        for i = 1 to cnt 'Запрос выполняется для каждого студента
	        if history_filtered(i,1) > 0 then 'Если пользователь в массиве имеет свой ID
		        'Подготавливаем данные
                if history_filtered(i,3) < 0 then
                response.Write("<tr style='background-color: #4390df;color: white;'>")
                elseif history_filtered(i,3) > 0 then
                response.Write("<tr style='background-color: #60a917;color: white;'>")
                else
                response.Write("<tr>")
                end if

		        'Рисуем таблицу
		        response.Write("<td style='width:1px'>" & history_filtered(i,1) & "</td>") 'ID
                response.Write("<td style='width:25%'>" & history_filtered(i,2) & "</td>") 'ФИО Студента
                response.Write("<td style='width:1px'>" & history_filtered(i,3) & "</td>") 'Баллы
                response.Write("<td style='width:1px'>" & history_filtered(i,4) & "</td>") 'ФИО Преподавателя
                response.Write("<td style='width:50%'>" & history_filtered(i,5) & "</td>") 'Комментарий
                response.Write("<td style='width:1px'>" & history_filtered(i,6) & "</td>") 'Дата
                response.Write("<td style='width:15%'><a href='?edit=" & history_filtered(i,1) & "'><button class='button subinfo'><span class='icon mif-pencil'></span> Изменить</button></a></td>") 'Кнопка редактирования

		        response.Write("</tr>")
	        end if
        next
    %>
    </tbody>
    </table>
    <% elseif cnt <= 0 or cnt = "" then %>
    <br>
    <h4><b>Записей не найдено!</b></h4>
    <% end if %>
    </center>
    <br />
<%
    else ' Основная страница
    group = session("gr")

    'Выполняем подключение к БД
    Set con = Server.CreateObject("ADODB.Connection")
    Set rs  = Server.CreateObject("ADODB.RecordSet")
    strdbpath=server.mappath("base.mdb")
    con.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strdbpath

    strSQL = "SELECT tbl_group.id_group, tbl_group.group_name, tbl_student.id_student, tbl_student.student_fio, Sum(tbl_score.score_value) AS [sum_score_value] FROM tbl_group INNER JOIN (tbl_student LEFT JOIN tbl_score ON tbl_student.id_student = tbl_score.score_student) ON tbl_group.id_group = tbl_student.group_name GROUP BY tbl_group.id_group, tbl_group.group_name, tbl_student.id_student, tbl_student.student_fio HAVING (((tbl_group.group_name)='" + group + "')) ORDER BY tbl_student.student_fio;"
    rs.Open strSQL, con, 3, 3

    dim students(3000,3) 'Массив на 3000 студентов
    if rs.RecordCount > 0 then
	    ij = 1
	    cnt = 1
	    while rs.EOF <> true
		    students(ij,1) = rs.Fields("id_student") 'ID
		    students(ij,2) = rs.Fields("student_fio") 'ФИО
		    if rs.Fields("sum_score_value") <> "" then students(ij,3) = rs.Fields("sum_score_value") else students(ij,3) = 0 'Текущая сумма баллов
            ij = ij + 1
            cnt = cnt + 1
            rs.MoveNext
	    wend
    end if
%>  
    <center>
    <a href="?addNew=1" ><button class="button primary"><span class="icon mif-pencil"></span> Новая запись</button></a>
    <a href="?history=1" ><button class="button subinfo"><span class="icon mif-history"></span> Просмотр истории записей</button></a>
    <h3>Сумма поощрительных баллов студентов группы <%=group%></h3>
    <table class="table striped hovered cell-hovered border bordered" style="width: 50vw">
    <thead align=center style="font-weight: bold">
    <tr><th>№</th><th>ФИО</th><th>Сумма баллов</th></tr>
    </thead>
    <tbody align=center>
<%
    for i = 1 to cnt 'Запрос выполняется для каждого студента
	    if students(i,1) > 0 then 'Если пользователь в массиве имеет свой ID
		    'Подготавливаем данные
		    if students(i,3) > 0 then
            response.Write("<tr style='background-color: #60a917;color: white;'>")
            elseif students(i,3) = 0 then
            response.Write("<tr>")
            elseif students(i,3) < 0 then
            response.Write("<tr style='background-color: #4390df;color: white;'>")
            end if
		
		    'Рисуем таблицу
		    response.Write("<td style='width:1px'>" & i & "</td>")
            response.Write("<td style='width:50%'>" & students(i,2) & "</td>")
            response.Write("<td style='width:50%'>" & students(i,3) & "</td>")

		    response.Write("</tr>")
	    end if
    next
%>
</center>
</tbody>
</table>
<br>
<% end if %>
<center>
<a href="disc_change.asp?go=1"><button type="button" class="button subinfo">Вернуться к выбору дисциплин и ведомостей</button><br>
<a href="group_change.asp?go=1"><button type="button" class="button subinfo">Вернуться к выбору группы</button><br>
<a href="help/03_10.asp" ><button class="button success"><span class="icon mif-info"></span> Помощь</button></a>
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
</body>
</html>