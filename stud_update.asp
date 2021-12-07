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
	<title>ИС "Электронный журнал". Перевод студентов</title>
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
'Защита от студентов
if session("user") = "" or session("user") = "Студент" or session("user") = 0 then response.Redirect ("404.asp")

confirm = request.querystring("confirm")

if confirm = 1 then
%>
<center>
<h4>Вы собираетесь перевести каждого студента каждой группы на курс выше, вы уверены?</h4>
<script>
    function ShowHideMiniLoader(operation){
        if (operation == "hide"){
            document.getElementById("mini_loader").style.display = "none";
        } else {
            document.getElementById("mini_loader").style.display = "block";    
        }
    }
</script>
<a href="stud_update.asp?confirm=0"><button class="button primary" onclick="ShowHideMiniLoader('show')">Подтвердить<span class="mif-spinner3 mif-ani-spin" id="mini_loader" style="color: #fff; display: none; float: right; margin-left: 5px;"></span></button></a> <a href="group_change.asp"><button class="button danger">Вернуться назад</button></a>
</center>
<%
function Log(value)
    response.Write("<script language=javascript>console.log('" & value & "'); </script>")
end function
dim group_missmatch
group_missmatch = 0
function group_miss()
    group_missmatch = group_missmatch + 1
end function

elseif confirm = 0 then

'Выполняем подключение к БД
Set con=Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.RecordSet")
Set rs2 = Server.CreateObject("ADODB.RecordSet")
Set rs3 = Server.CreateObject("ADODB.RecordSet")
Set rs4 = Server.CreateObject("ADODB.RecordSet")
Set rs5 = Server.CreateObject("ADODB.RecordSet")

strdbpath=server.mappath("base.mdb")
con.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strdbpath
groupSearch = request.querystring("groupSearch")
if groupSearch = "" then
    rs.Open "SELECT tbl_student.id_student, tbl_student.student_fio, tbl_group.id_group FROM tbl_group INNER JOIN tbl_student ON tbl_group.id_group = tbl_student.group_name;", con, 3, 3
else
   rs.Open "SELECT tbl_student.id_student, tbl_student.student_fio, tbl_group.id_group FROM tbl_group INNER JOIN tbl_student ON tbl_group.id_group = tbl_student.group_name WHERE tbl_group.id_group = " & groupSearch & ";", con, 3, 3
end if
if len(request.Form) > 0 then
    if request.Form("addNewGroup") = "1" then
        group_name = request.Form("group_name")
        group_ruk = request.Form("group_ruk")
        spec = request.Form("spec")
        today = Now()
        dd = mid(today, 1, 2)
        mm = mid(today, 4, 2)
        yyyy = mid(today, 7, 4)
        nowDate = mm + "/" + dd + "/" + yyyy
        'Создаем подключение для запросов
        strDbConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="& strdbpath & ";"
        Set objConn = Server.CreateObject("ADODB.Connection")
        objConn.Open(strDbConnection)
        strSQL = "INSERT INTO tbl_group (group_name, group_ruk, spec, nb_edit_data) VALUES ('" & group_name & "', " & group_ruk & ", " & spec & ", #" & nowDate & "#);" 'Подготавливаем запрос
        'Выполняем запрос
	    objConn.Execute(strSQL)
        response.Redirect("stud_update.asp?confirm=0")
    elseif request.Form("useBackup") = "1" then
        Dim fso, curdir
        curdir = Server.MapPath(".")
        Set fso = CreateObject("Scripting.FileSystemObject")
        fso.CopyFile curdir + "/base_backup.mdb", curdir + "/base.mdb"
        fso.DeleteFile curdir + "/base_backup.mdb"
        set fso = nothing
        response.Redirect("stud_update.asp?confirm=0")
    end if
end if

'Формируем массив
dim students(3000,5) 'Массив на 3000 студентов
if rs.RecordCount > 0 then
	ij = 1
	cnt = 1
	while rs.EOF <> true
		students(ij,1) = rs.Fields(0) 'ID
		students(ij,2) = rs.Fields(1) 'ФИО
		students(ij,3) = rs.Fields(2) 'Текущая группа
        students(ij,4) = 0 'Новая группа
        rs3.Open "SELECT tbl_student.id_student, tbl_student.student_fio, tbl_group.id_group, tbl_group.group_name FROM tbl_group INNER JOIN tbl_student ON tbl_group.id_group = tbl_student.group_name WHERE (((tbl_group.id_group)=" & students(ij,3) &"));", con, 3
            'ГРУППЫ СПЕЦИАЛЬНОСТИ АТ
            groupFrom = "11-АТ" ' Группа из которой переводим
            groupTo   = "21-АТ" ' Группа в которую переводим
            if rs3.RecordCount > 0 AND rs3.Fields(3) = groupFrom then
                rs4.Open "SELECT tbl_group.id_group, tbl_group.group_name FROM tbl_group WHERE (((tbl_group.group_name)='" & groupTo & "'));", con, 3
                if rs4.RecordCount > 0 then
                    students(ij,4) = rs4.Fields(0)
                elseif rs4.RecordCount = 0 then
                    students(ij,4) = -2
                    students(ij,5) = groupTo
                    group_miss()
                end if
                rs4.Close
            end if

            groupFrom = "21-АТ" ' Группа из которой переводим
            groupTo   = "31-АТ" ' Группа в которую переводим
            if rs3.RecordCount > 0 AND rs3.Fields(3) = groupFrom then
                rs4.Open "SELECT tbl_group.id_group, tbl_group.group_name FROM tbl_group WHERE (((tbl_group.group_name)='" & groupTo & "'));", con, 3
                if rs4.RecordCount > 0 then
                    students(ij,4) = rs4.Fields(0)
                elseif rs4.RecordCount = 0 then
                    students(ij,4) = -2
                    students(ij,5) = groupTo
                    group_miss()
                end if
                rs4.Close
            end if

            groupFrom = "31-АТ" ' Группа из которой переводим
            groupTo   = "41-АТ" ' Группа в которую переводим
            if rs3.RecordCount > 0 AND rs3.Fields(3) = groupFrom then
                rs4.Open "SELECT tbl_group.id_group, tbl_group.group_name FROM tbl_group WHERE (((tbl_group.group_name)='" & groupTo & "'));", con, 3
                if rs4.RecordCount > 0 then
                    students(ij,4) = rs4.Fields(0)
                elseif rs4.RecordCount = 0 then
                    students(ij,4) = -2
                    students(ij,5) = groupTo
                    group_miss()
                end if
                rs4.Close
            end if

            groupFrom = "41-АТ" ' Группа из которой переводим
            if rs3.Fields(3) = groupFrom then
                students(ij,4) = -1
            end if
            '/АТ
            
            ' ГРУППЫ СПЕЦИАЛЬНОСТИ ВП
            groupFrom = "11-ВП" ' Группа из которой переводим
            groupTo   = "21-ВП" ' Группа в которую переводим
            if rs3.RecordCount > 0 AND rs3.Fields(3) = groupFrom then
                rs4.Open "SELECT tbl_group.id_group, tbl_group.group_name FROM tbl_group WHERE (((tbl_group.group_name)='" & groupTo & "'));", con, 3
                if rs4.RecordCount > 0 then
                    students(ij,4) = rs4.Fields(0)
                elseif rs4.RecordCount = 0 then
                    students(ij,4) = -2
                    students(ij,5) = groupTo
                    group_miss()
                end if
                rs4.Close
            end if

            groupFrom = "21-ВП" ' Группа из которой переводим
            groupTo   = "31-ВП" ' Группа в которую переводим
            if rs3.RecordCount > 0 AND rs3.Fields(3) = groupFrom then
                rs4.Open "SELECT tbl_group.id_group, tbl_group.group_name FROM tbl_group WHERE (((tbl_group.group_name)='" & groupTo & "'));", con, 3
                if rs4.RecordCount > 0 then
                    students(ij,4) = rs4.Fields(0)
                elseif rs4.RecordCount = 0 then
                    students(ij,4) = -2
                    students(ij,5) = groupTo
                    group_miss()
                end if
                rs4.Close
            end if

            groupFrom = "31-ВП" ' Группа из которой переводим
            groupTo   = "41-ВП" ' Группа в которую переводим
            if rs3.RecordCount > 0 AND rs3.Fields(3) = groupFrom then
                rs4.Open "SELECT tbl_group.id_group, tbl_group.group_name FROM tbl_group WHERE (((tbl_group.group_name)='" & groupTo & "'));", con, 3
                if rs4.RecordCount > 0 then
                    students(ij,4) = rs4.Fields(0)
                elseif rs4.RecordCount = 0 then
                    students(ij,4) = -2
                    students(ij,5) = groupTo
                    group_miss()
                end if
                rs4.Close
            end if

            groupFrom = "41-ВП" ' Группа из которой переводим
            if rs3.Fields(3) = groupFrom then
                students(ij,4) = -1
            end if
            '/ВП

            ' ГРУППЫ СПЕЦИАЛЬНОСТИ ИС
            groupFrom = "11-ИС" ' Группа из которой переводим
            groupTo   = "21/12-ИС" ' Группа в которую переводим
            if rs3.RecordCount > 0 AND rs3.Fields(3) = groupFrom then
                rs4.Open "SELECT tbl_group.id_group, tbl_group.group_name FROM tbl_group WHERE (((tbl_group.group_name)='" & groupTo & "'));", con, 3
                if rs4.RecordCount > 0 then
                    students(ij,4) = rs4.Fields(0)
                elseif rs4.RecordCount = 0 then
                    students(ij,4) = -2
                    students(ij,5) = groupTo
                    group_miss()
                end if
                rs4.Close
            end if

            groupFrom = "21/12-ИС" ' Группа из которой переводим
            groupTo   = "31/22-ИС" ' Группа в которую переводим
            if rs3.RecordCount > 0 AND rs3.Fields(3) = groupFrom then
                rs4.Open "SELECT tbl_group.id_group, tbl_group.group_name FROM tbl_group WHERE (((tbl_group.group_name)='" & groupTo & "'));", con, 3
                if rs4.RecordCount > 0 then
                    students(ij,4) = rs4.Fields(0)
                elseif rs4.RecordCount = 0 then
                    students(ij,4) = -2
                    students(ij,5) = groupTo
                    group_miss()
                end if
                rs4.Close
            end if

            groupFrom = "31/22-ИС" ' Группа из которой переводим
            groupTo   = "41/32-ИС" ' Группа в которую переводим
            if rs3.RecordCount > 0 AND rs3.Fields(3) = groupFrom then
                rs4.Open "SELECT tbl_group.id_group, tbl_group.group_name FROM tbl_group WHERE (((tbl_group.group_name)='" & groupTo & "'));", con, 3
                if rs4.RecordCount > 0 then
                    students(ij,4) = rs4.Fields(0)
                elseif rs4.RecordCount = 0 then
                    students(ij,4) = -2
                    students(ij,5) = groupTo
                    group_miss()
                end if
                rs4.Close
            end if

            groupFrom = "41/32-ИС" ' Группа из которой переводим
            if rs3.Fields(3) = groupFrom then
                students(ij,4) = -1
            end if
            '/ИС

            ' ГРУППЫ СПЕЦИАЛЬНОСТИ СВ
            groupFrom = "11-СВ" ' Группа из которой переводим
            groupTo   = "21-СВ" ' Группа в которую переводим
            if rs3.RecordCount > 0 AND rs3.Fields(3) = groupFrom then
                rs4.Open "SELECT tbl_group.id_group, tbl_group.group_name FROM tbl_group WHERE (((tbl_group.group_name)='" & groupTo & "'));", con, 3
                if rs4.RecordCount > 0 then
                    students(ij,4) = rs4.Fields(0)
                elseif rs4.RecordCount = 0 then
                    students(ij,4) = -2
                    students(ij,5) = groupTo
                    group_miss()
                end if
                rs4.Close
            end if

            groupFrom = "12-СВ" ' Группа из которой переводим
            groupTo   = "22-СВ" ' Группа в которую переводим
            if rs3.RecordCount > 0 AND rs3.Fields(3) = groupFrom then
                rs4.Open "SELECT tbl_group.id_group, tbl_group.group_name FROM tbl_group WHERE (((tbl_group.group_name)='" & groupTo & "'));", con, 3
                if rs4.RecordCount > 0 then
                    students(ij,4) = rs4.Fields(0)
                elseif rs4.RecordCount = 0 then
                    students(ij,4) = -2
                    students(ij,5) = groupTo
                    group_miss()
                end if
                rs4.Close
            end if

            groupFrom = "21-СВ" ' Группа из которой переводим
            groupTo   = "31-СВ" ' Группа в которую переводим
            if rs3.RecordCount > 0 AND rs3.Fields(3) = groupFrom then
                rs4.Open "SELECT tbl_group.id_group, tbl_group.group_name FROM tbl_group WHERE (((tbl_group.group_name)='" & groupTo & "'));", con, 3
                if rs4.RecordCount > 0 then
                    students(ij,4) = rs4.Fields(0)
                elseif rs4.RecordCount = 0 then
                    students(ij,4) = -2
                    students(ij,5) = groupTo
                    group_miss()
                end if
                rs4.Close
            end if

            groupFrom = "22-СВ" ' Группа из которой переводим
            groupTo   = "32-СВ" ' Группа в которую переводим
            if rs3.Fields(3) = groupFrom then
                rs4.Open "SELECT tbl_group.id_group, tbl_group.group_name FROM tbl_group WHERE (((tbl_group.group_name)='" & groupTo & "'));", con, 3
                if rs4.RecordCount > 0 then
                    students(ij,4) = rs4.Fields(0)
                elseif rs4.RecordCount = 0 then
                    students(ij,4) = -2
                    students(ij,5) = groupTo
                    group_miss()
                end if
                rs4.Close
            end if

            groupFrom = "31-СВ" ' Группа из которой переводим
            groupTo   = "41-СВ" ' Группа в которую переводим
            if rs3.RecordCount > 0 AND rs3.Fields(3) = groupFrom then
                rs4.Open "SELECT tbl_group.id_group, tbl_group.group_name FROM tbl_group WHERE (((tbl_group.group_name)='" & groupTo & "'));", con, 3
                if rs4.RecordCount > 0 then
                    students(ij,4) = rs4.Fields(0)
                elseif rs4.RecordCount = 0 then
                    students(ij,4) = -2
                    students(ij,5) = groupTo
                    group_miss()
                end if
                rs4.Close
            end if

            groupFrom = "32-СВ" ' Группа из которой переводим
            groupTo   = "42-СВ" ' Группа в которую переводим
            if rs3.RecordCount > 0 AND rs3.Fields(3) = groupFrom then
                rs4.Open "SELECT tbl_group.id_group, tbl_group.group_name FROM tbl_group WHERE (((tbl_group.group_name)='" & groupTo & "'));", con, 3
                if rs4.RecordCount > 0 then
                    students(ij,4) = rs4.Fields(0)
                elseif rs4.RecordCount = 0 then
                    students(ij,4) = -2
                    students(ij,5) = groupTo
                    group_miss()
                end if
                rs4.Close
            end if

            groupFrom = "41-СВ" ' Группа из которой переводим
            if rs3.Fields(3) = groupFrom then
                students(ij,4) = -1
            end if

            groupFrom = "42-СВ" ' Группа из которой переводим
            groupTo   = "52-СВ" ' Группа в которую переводим
            if rs3.RecordCount > 0 AND rs3.Fields(3) = groupFrom then
                rs4.Open "SELECT tbl_group.id_group, tbl_group.group_name FROM tbl_group WHERE (((tbl_group.group_name)='" & groupTo & "'));", con, 3
                if rs4.RecordCount > 0 then
                    students(ij,4) = rs4.Fields(0)
                elseif rs4.RecordCount = 0 then
                    students(ij,4) = -2
                    students(ij,5) = groupTo
                    group_miss()
                end if
                rs4.Close
            end if

            groupFrom = "52-СВ" ' Группа из которой переводим
            if rs3.Fields(3) = groupFrom then
                students(ij,4) = -1
            end if
            '/СВ

            ' ГРУППЫ СПЕЦИАЛЬНОСТИ ЭР
            groupFrom = "11-ЭР" ' Группа из которой переводим
            groupTo   = "21-ЭР" ' Группа в которую переводим
            if rs3.RecordCount > 0 AND rs3.Fields(3) = groupFrom then
                rs4.Open "SELECT tbl_group.id_group, tbl_group.group_name FROM tbl_group WHERE (((tbl_group.group_name)='" & groupTo & "'));", con, 3
                if rs4.RecordCount > 0 then
                    students(ij,4) = rs4.Fields(0)
                elseif rs4.RecordCount = 0 then
                    students(ij,4) = -2
                    students(ij,5) = groupTo
                    group_miss()
                end if
                rs4.Close
            end if

            groupFrom = "21-ЭР" ' Группа из которой переводим
            groupTo   = "31-ЭР" ' Группа в которую переводим
            if rs3.RecordCount > 0 AND rs3.Fields(3) = groupFrom then
                rs4.Open "SELECT tbl_group.id_group, tbl_group.group_name FROM tbl_group WHERE (((tbl_group.group_name)='" & groupTo & "'));", con, 3
                if rs4.RecordCount > 0 then
                    students(ij,4) = rs4.Fields(0)
                elseif rs4.RecordCount = 0 then
                    students(ij,4) = -2
                    students(ij,5) = groupTo
                    group_miss()
                end if
                rs4.Close
            end if

            groupFrom = "31-ЭР" ' Группа из которой переводим
            groupTo   = "41-ЭР" ' Группа в которую переводим
            if rs3.RecordCount > 0 AND rs3.Fields(3) = groupFrom then
                rs4.Open "SELECT tbl_group.id_group, tbl_group.group_name FROM tbl_group WHERE (((tbl_group.group_name)='" & groupTo & "'));", con, 3
                if rs4.RecordCount > 0 then
                    students(ij,4) = rs4.Fields(0)
                elseif rs4.RecordCount = 0 then
                    students(ij,4) = -2
                    students(ij,5) = groupTo
                    group_miss()
                end if
                rs4.Close
            end if

            groupFrom = "41-ЭР" ' Группа из которой переводим
            if rs3.Fields(3) = groupFrom then
                students(ij,4) = -1
            end if
            '/ЭР
            
        rs3.Close
		ij = ij + 1
		cnt = cnt + 1
		rs.MoveNext
	wend
end if

if len(request.Form) > 0 then
    if request.Form("update") = "1" then
        'Резервное копирование БД, хранится в корне ИС на сервере
        dim fs
        set fs=Server.CreateObject("Scripting.FileSystemObject")
        dim CurrentDirectory
        CurrentDirectory = Server.MapPath(".")
        fs.CopyFile CurrentDirectory + "/base.mdb", CurrentDirectory + "/base_backup.mdb"
        set fs = nothing

        'Создаем подключение для запросов
        strDbConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="& strdbpath & ";"
        Set objConn = Server.CreateObject("ADODB.Connection")
        objConn.Open(strDbConnection)

        strSQL = "DELETE FROM tbl_itog" 'Чистим таблицу итог
        objConn.Execute(strSQL)
        strSQL = "DELETE FROM tbl_journal" 'Чистим таблицу журнал
        objConn.Execute(strSQL)
        strSQL = "DELETE FROM tbl_journal_del" 'Чистим таблицу журнал
        objConn.Execute(strSQL)
        strSQL = "DELETE FROM tbl_zan" 'Чистим таблицу журнал
        objConn.Execute(strSQL)

        for i = 1 to cnt 'Запрос выполняется для каждого студента
	        if students(i,1) > 0 then 'Если пользователь в массиве имеет свой ID
                if students(i, 4) > 0 then 'Если группа студента не относится к отчислению, т.е. -1 то тогда переносим его на следующий курс
                    strSQL = "UPDATE tbl_group INNER JOIN tbl_student ON tbl_group.id_group = tbl_student.group_name SET tbl_student.group_name = " & students(i,4) & " WHERE (((tbl_student.id_student)=" & students(i,1) & "));" 'Подготавливаем запрос
                    objConn.Execute(strSQL)
                elseif students(i,4) = -1 then
                    strSQL = "DELETE FROM tbl_student WHERE (((tbl_student.id_student)=" & students(i,1) & "));" 'Подготавливаем запрос
                    objConn.Execute(strSQL)
                end if
            end if
        next

        response.Redirect("stud_update.asp?confirm=0") ' Избежание повторных переносов студентов из-за обновления страницы после POST запроса, сложная хрень, не пытайтесь вникнуть.
    end if
end if
%>
<center>
<% if group_missmatch > 0 then %>
<h1 style="color: Red; font-weight: bold">Внимание!</h1>
<h4>При выполнении данного запроса были найдены ошибки несоответсвия следующих групп студентов, следует добавить отсутсвующие группы в БД!</h2>
<p><a href=#missmatch>Количество ошибок: <% response.Write(group_missmatch) %></a></p>


<details style="margin-bottom: 15px">
<summary style="background: #2086bf; color: #fff; border-color: #2086bf; box-shadow: rgb(0 0 0 / 20%) 0px 3px 5px;"><span class="icon mif-wrench"></span> Добавить отсутсвующую группу в БД</summary>

<form name="addNew" action="stud_update.asp" method="post" style="margin-bottom: 0">
Название группы
    <div class="input-control text" style="width: 75px">
    <input type=text name="group_name" placeholder="11-АБ" required pattern="[0-9]{2}-[А-ЯЁ]{2}">
    </div>
	<div class="input-control text" style="width: 150px">
	<select name="group_ruk"  required>
        <option disabled selected value>Преподователь</option>
		<% rs5.Open "tbl_user", con
         set objId = rs5.Fields(0)
         set objName = rs5.Fields(1)
         do until rs5.EOF %>
       <option value="<% response.write(objId) %>"><% response.write(objName) %></option> 
        <% rs5.MoveNext
           loop
           rs5.Close %>
	</select>
    </div>
    <div class="input-control text" style="width: 150px">
    <select name="spec"  required>
        <option disabled selected value>Описание</option>
		<% rs5.Open "tbl_spec", con
         set objId = rs5.Fields(0)
         set objName = rs5.Fields(2)
         do until rs5.EOF %>
       <option value="<% response.write(objId) %>"><% response.write(objName) %></option> 
        <% rs5.MoveNext
           loop
           rs5.Close %>
	</select>
	</div>
    
    <input type="hidden" name="addNewGroup" value=1 />
	<button type="submit" class="button primary"><span class="icon mif-pencil"></span>  Добавить запись</button> <button type="reset" class="button danger"><span class="icon mif-undo"></span>  Сброс</button>
</form>

</details>
<% elseif group_missmatch = 0 and groupSearch = "" then %>
<form name="studUpdate" action="stud_update.asp" method="post" style="margin-bottom: 0; width: 50vw">
<h4 style="color: #60a917;"><span class="icon mif-checkmark"></span> Ошибок несоотвтетсвия в БД не было обнаружено, всё готово к переводу студентов, выполнить?</h4>
<h6 style="color: Gray;">В результате выполнения этого действия будет создана резервная копия БД (base_backup.mdb в корне ИС на сервере), в случае возникновения ошибок в работе Электронного журнала можно будет вернуться к предыдущей копии БД.</h6>
<input type="hidden" name="update" value=1 />
<button type="submit" class="button success"><span class="icon mif-magic-wand"></span>  Перевести студентов</button><br><br>
</form>
<%
elseif groupSearch <> "" then
%>
<h6 style="color: Gray;">Вы находитесь в режиме фильтрации, для возврата к переводу студентов сбросьте фильтр нажатием кнопки "Сброс".</h6>
<%
end if 

'Проверяем, существует ли копия БД в корне ИС на сервере
dim f
set fs=Server.CreateObject("Scripting.FileSystemObject")
CurrentDirectory = Server.MapPath(".")
if fs.FileExists(CurrentDirectory + "/base_backup.mdb") then
    set f=fs.GetFile(CurrentDirectory + "/base_backup.mdb")
    %>
    <form name="backup" action="stud_update.asp" method="post" style="margin-bottom: 0; width: 50vw">
    <h4 style="color: #2086bf;"><span class="icon mif-notification"></span> Была обнаружена резеврная копия БД от <% Response.Write(f.DateLastModified) %></h4>
    <input type="hidden" name="useBackup" value=1 />
    <button type="submit" class="button primary"><span class="icon mif-undo"></span>  Вернуться на предыдующую копию БД</button><br><br>
    </form>
    <%
end if
set f=nothing
set fs = nothing
%>

<form name="filter" action="stud_update.asp" method="get" style="margin-bottom: 0">
Фильтрация по группе
	<div class="input-control text" style="width: 125px">
	<select name="groupSearch" required>
        <option disabled selected value>Выберите...</option>
		<% rs5.Open "SELECT tbl_group.id_group, tbl_group.group_name FROM tbl_group ORDER BY tbl_group.group_name;", con
         set objId = rs5.Fields("id_group")
         set objName = rs5.Fields("group_name")
         do until rs5.EOF %>
       <option value="<% response.write(objId) %>" <% if cstr(objId) = cstr(groupSearch) then response.Write(" selected") %>><% response.write(objName) %></option> 
        <% rs5.MoveNext
           loop
           rs5.Close %>
	</select>
	</div>
    <button type="submit" class="button primary"><span class="icon mif-filter"></span>  Выбрать</button> <a href="?confirm=0"><button type="button" class="button danger"><span class="icon mif-undo"></span>  Сброс</button></a><br><br>
</form>
<table class="table striped hovered cell-hovered border bordered" style="width: 50vw">
<thead align=center style="font-weight: bold">
<tr><th>ID</th><th>ФИО</th><th>Текущая группа</th><th>Новая группа</th></tr>
</thead>
<tbody align=center>
<%
for i = 1 to cnt 'Запрос выполняется для каждого студента
	if students(i,1) > 0 then 'Если пользователь в массиве имеет свой ID
		'Подготавливаем данные
		response.Write("<tr>")
		
		'Рисуем таблицу
		response.Write("<td>" & students(i,1) & "</td>")
        response.Write("<td>" & students(i,2) & "</td>")
        
        response.Write("<td>")
        rs2.Open "SELECT tbl_group.id_group, tbl_group.group_name FROM tbl_group WHERE (((tbl_group.id_group)=" & students(i,3) & "));", con
        set objId = rs2.Fields(0)
        set objName = rs2.Fields(1)
        do until rs2.EOF
            response.write(objName)
            rs2.MoveNext
        loop
        rs2.Close
        response.Write("</td>")
        
        if students(i,4) = -1 then
            response.Write("<td style='color: green;'>Отчислен</td>")
        elseif students(i,4) = -2 then
            response.Write("<td title='Эта группа отсутсвует в БД!'><a name='missmatch' style='color: red; font-weight: bold'>" + students(i,5) + "</a></td>")
        else
        response.Write("<td>")
        rs2.Open "SELECT tbl_group.id_group, tbl_group.group_name FROM tbl_group WHERE (((tbl_group.id_group)=" & students(i,4) & "));", con
        set objId = rs2.Fields(0)
        set objName = rs2.Fields(1)
        do until rs2.EOF
            response.write(objName)
            rs2.MoveNext
        loop
        rs2.Close
        response.Write("</td>")
        end if

		
		response.Write("</tr>")
	end if
next
%>
</tbody>
</table>
</center>
<% end if %>
<br>
<center>
<a name="end"></a>
<a href="group_change.asp?go=1"><button type="button" class="button subinfo">Вернуться к выбору группы</button><br>
<a href="help/03_2.asp" ><button class="button success"><span class="icon mif-info"></span> Помощь</button></a>
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