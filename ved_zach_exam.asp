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
	<script src="js/jquery-3.1.0.min.js"></script>
	<script src="js/metro.min.js"></script>
	<title>ИС "Электронный журнал". Печать зачётно-экзаменационных ведомостей</title>
    <style>
        input, select {
            font-family: 'Segoe UI', 'Open Sans', sans-serif, serif; 
            font-size: 0.875rem;
        }
        .table tbody td {
            padding: 4px;
        }
        .switch, .switch-original {
            margin: 0 .125rem 0 0;
        }
        table .main-table td {
            text-align: center;
            font-size: 12pt;
        }
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
			table .main-table {
                border-collapse: collapse;
            }
            table .main-table td {
                text-align: center;
                border: solid 1px black;
            }
            table .info-table {
                border-collapse: collapse;
            }
            table .main-table, td .main-table, th .main-table {
                border: 1px solid black;
            }
            tr.numbers {
                text-align: center;
                font-size: 11pt;
                height: min-content;
            }
            table .main-table td {
                text-align: center;
                font-size: 12pt;
            }
        }
    </style>
</head>
<body>
<div id=notPrintableArea>
<table class="table border" style='width:90%; margin-top: 15px;' align=center> 	
<tr>
<td>
    <!-- #include file="header.asp" -->
    <!-- #include file="pass_check.asp" -->
<%
'Защита от студентов
if session("user") = "" or session("user") = "Студент" or session("user") = 0 then response.Redirect ("404.asp")

query = request.querystring("query")

if query = 0 then

today_input = mid(date(), 7, 4) + "-" + mid(date(), 4, 2) + "-" + mid(date(), 1, 2)
today_date  = mid(date(), 1, 2) + "." + mid(date(), 4, 2) + "." + mid(date(), 7, 4)

Set con = Server.CreateObject("ADODB.Connection")
Set rs1 = Server.CreateObject("ADODB.RecordSet")
Set rs2 = Server.CreateObject("ADODB.RecordSet")
Set rs3 = Server.CreateObject("ADODB.RecordSet")
strdbpath=server.mappath("base.mdb")
con.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strdbpath

function Log(value)
    response.Write("<script language=javascript>console.log('" & value & "'); </script>")
end function
%>
<script>
    $(window).scroll(function () {
        sessionStorage.scrollTop = $(this).scrollTop();
    });
    $(document).ready(function () {
        if (sessionStorage.scrollTop != "undefined") {
            $(window).scrollTop(sessionStorage.scrollTop);
        }
    });
</script>
<center>
<form action="ved_zach_exam.asp?query=0" method="post" onsubmit="this.form.submit()">
<table style="font-family: 'Segoe UI', 'Open Sans', sans-serif, serif; font-size: 0.875rem;">
    <tr>
        <td align=right>Дата:</td>
        <td><div class="input-control text" style="width: 140px"><input type="date" name="date" value="<% if request.form("date") = "" then response.write(today_input) else response.write(request.form("date")) %>" onchange="this.form.submit()"></div></td>
	</tr>
	<tr>
		<%
        groupSearch = request.Form("groupSearch")
        if groupSearch = "" then
            strSQL = "SELECT TOP 1 tbl_group.id_group, tbl_group.group_name FROM tbl_group GROUP BY tbl_group.id_group, tbl_group.group_name ORDER BY tbl_group.group_name;"
            rs2.Open strSQL, con
            groupSearch = rs2.Fields("id_group")
            rs2.Close()
        end if

        semestr = request.Form("semestr")
        if semestr = "" then semestr = 1

        hideBrackets = request.Form("hideBrackets")
        multiplePrepod = request.Form("multiplePrepod")

        discSearch = request.Form("discSearch")
        if discSearch = "" then
            strSQL = "SELECT TOP 1 tbl_group.id_group, tbl_plan.Semestr, tbl_disc.ID_disc, tbl_disc.disc_name FROM tbl_disc INNER JOIN ((tbl_group INNER JOIN tbl_student ON tbl_group.id_group = tbl_student.group_name) INNER JOIN tbl_plan ON tbl_group.id_group = tbl_plan.gr_name) ON tbl_disc.ID_disc = tbl_plan.disc_name GROUP BY tbl_group.id_group, tbl_plan.Semestr, tbl_disc.ID_disc, tbl_disc.disc_name HAVING (((tbl_group.id_group)=" + cstr(groupSearch) + ") AND ((tbl_plan.Semestr)='" + cstr(semestr) + "'));"
            rs2.Open strSQL, con
            discSearch = rs2.Fields("ID_disc")
            rs2.Close()
        end if

        strSQL = "SELECT tbl_group.id_group, tbl_plan.Semestr, tbl_disc.ID_disc, tbl_plan.control_form FROM (tbl_group INNER JOIN (tbl_disc INNER JOIN tbl_plan ON tbl_disc.ID_disc = tbl_plan.disc_name) ON tbl_group.id_group = tbl_plan.gr_name) INNER JOIN tbl_student ON tbl_group.id_group = tbl_student.group_name GROUP BY tbl_group.id_group, tbl_plan.Semestr, tbl_disc.ID_disc, tbl_plan.control_form HAVING (((tbl_group.id_group)=" + cstr(groupSearch) + ") AND ((tbl_plan.Semestr)='" + cstr(semestr) + "') AND ((tbl_disc.ID_disc)=" + cstr(discSearch) + "));"
        rs2.Open strSQL, con
        if rs2.EOF = false then
        control = rs2.Fields("control_form")
        end if
        rs2.Close()
        
        if len(request.Form("prepod")) = 0 then
            strSQL = "SELECT tbl_group.id_group, tbl_plan.Semestr, tbl_disc.ID_disc, tbl_plan.control_form, tbl_user.id_user, tbl_user.user_fio FROM tbl_user INNER JOIN (tbl_group INNER JOIN (tbl_disc INNER JOIN tbl_plan ON tbl_disc.ID_disc = tbl_plan.disc_name) ON tbl_group.id_group = tbl_plan.gr_name) ON tbl_user.id_user = tbl_plan.Prepod_name WHERE (((tbl_group.id_group)=" + cstr(groupSearch) + ") AND ((tbl_plan.Semestr)='" + cstr(semestr) + "') AND ((tbl_disc.ID_disc)=" + cstr(discSearch) + ") AND ((tbl_plan.control_form)='" + cstr(control) + "'));"
            rs2.Open strSQL, con
            if rs2.EOF = false then
            prepod = rs2.Fields("id_user")
            end if
            rs2.Close()
        elseif len(request.Form("prepod")) > 0 then
            prepod = request.Form("prepod")
        end if

        prepod2enable = request.Form("prepod2enable")
        prepod3enable = request.Form("prepod3enable")

        if len(request.Form("prepod2")) = 0 then
            strSQL = "SELECT tbl_user.id_user, Left([tbl_user].[user_fio],Len([tbl_user].[user_fio])-4)+[tbl_user].[user_io] AS full_fio, tbl_user.user_fio, tbl_user.user_io FROM tbl_user WHERE (((tbl_user.user_fio) Not Like '%/%' And (tbl_user.user_fio) Not Like 'Стрекаловский Н.В.') AND ((tbl_user.user_io) Not Like '%-%' And (tbl_user.user_io) Not Like '%Вакансия%' And (tbl_user.user_io) Not Like '%.%')) ORDER BY tbl_user.user_fio;"
            rs2.Open strSQL, con
            if rs2.EOF = false then
            prepod2 = rs2.Fields("id_user")
            end if
            rs2.Close()
        elseif len(request.Form("prepod2")) > 0 then
            prepod2 = request.Form("prepod2")
        end if

        if len(request.Form("prepod3")) = 0 then
            strSQL = "SELECT tbl_user.id_user, Left([tbl_user].[user_fio],Len([tbl_user].[user_fio])-4)+[tbl_user].[user_io] AS full_fio, tbl_user.user_fio, tbl_user.user_io FROM tbl_user WHERE (((tbl_user.user_fio) Not Like '%/%' And (tbl_user.user_fio) Not Like 'Стрекаловский Н.В.') AND ((tbl_user.user_io) Not Like '%-%' And (tbl_user.user_io) Not Like '%Вакансия%' And (tbl_user.user_io) Not Like '%.%')) ORDER BY tbl_user.user_fio;"
            rs2.Open strSQL, con
            if rs2.EOF = false then
            prepod3 = rs2.Fields("id_user")
            end if
            rs2.Close()
        elseif len(request.Form("prepod3")) > 0 then
            prepod3 = request.Form("prepod3")
        end if
        
        Response.Write("<h4 style='color: #c93d37'>Данная страница находится в стадии активной разработки, и её работа может быть нестабильна!</h4>")
        Response.Write("<h5>В случае возникновения проблем с отображением записей попробуйте нажать кнопку 'Обновить'</h5>")

        %>
		<td align=right>Группа:</td>
		<td style="width: 310px;">
            <div class="input-control text" style="width: 125px">
            <select name="groupSearch" id="groupSearch" onchange="this.form.submit()">
		       <%
               strSQL = "SELECT tbl_group.id_group, tbl_group.group_name FROM (tbl_spec INNER JOIN tbl_group ON tbl_spec.id_spec = tbl_group.spec) INNER JOIN (tbl_disc INNER JOIN tbl_plan ON tbl_disc.ID_disc = tbl_plan.disc_name) ON tbl_group.id_group = tbl_plan.gr_name GROUP BY tbl_group.id_group, tbl_group.group_name ORDER BY tbl_group.group_name;" 
               rs1.Open strSQL, con
               set objId = rs1.Fields("id_group")
               set objName = rs1.Fields("group_name")
               do until rs1.EOF %>
               <option value="<% response.write(objId) %>" <% if cstr(objId) = cstr(groupSearch) then response.Write("selected") %>><% response.write(objName) %></option> 
                <% rs1.MoveNext
                   loop
                   rs1.Close %>
	        </select>
            </div>
        </td>
	</tr>
	<tr>
        <td align=right>Семестр:</td>
        <td>
            <div class="input-control text" style="width: 50px">
            <select name="semestr" onchange="this.form.submit()">
                <option value="1" <% if semestr = 1 then response.Write(" selected") else if semestr = "" then response.Write(" selected") %>>1</option>
                <option value="2" <% if semestr = 2 then response.Write(" selected") %>>2</option>
            </select>
            </div>
        </td>
	</tr>
    <tr>
        <% if groupSearch > 0 AND semestr > 0 then %>
		<td align=right>Дисциплина:</td>
		<td>
            <div class="input-control text">
            <select name="discSearch" style="width: 300px;" onchange="this.form.submit()">
		        <% 
                 rs1.Open "SELECT tbl_group.id_group, tbl_plan.Semestr, tbl_disc.ID_disc, tbl_disc.disc_name FROM (tbl_spec INNER JOIN tbl_group ON tbl_spec.id_spec = tbl_group.spec) INNER JOIN (tbl_disc INNER JOIN tbl_plan ON tbl_disc.ID_disc = tbl_plan.disc_name) ON tbl_group.id_group = tbl_plan.gr_name GROUP BY tbl_group.id_group, tbl_plan.Semestr, tbl_disc.ID_disc, tbl_disc.disc_name HAVING tbl_group.id_group=" + cstr(groupSearch) + " AND tbl_plan.Semestr='" + cstr(semestr) + "';", con
                 set objId = rs1.Fields("ID_disc")
                 set objName = rs1.Fields("disc_name")
                 do until rs1.EOF %>
               <option value="<% response.write(objId) %>" <% if cstr(objId) = cstr(discSearch) then response.Write("selected") %>><% response.write(objName) %></option> 
                <% rs1.MoveNext
                   loop
                   rs1.Close %>
	        </select>
            </div>
        </td>
        <% end if %>
	</tr>
    <tr title="Скрывает содержимое скобок из строки дисциплины.">
        <% if groupSearch > 0 AND semestr > 0 then %>
        <td align=right>Скрыть скобки:</td>
        <td>
            <label class='switch'><input type='checkbox' name="hideBrackets" <% if hideBrackets = "on" then response.Write("checked") %> onchange="this.form.submit()"><span class='check'></span></label>
        </td>
        <% end if %>
    </tr>
    <tr>
        <% if groupSearch > 0 AND semestr > 0 AND discSearch > 0 then %>
		<td align=right>Форма контроля:</td>
		<td>
            <div class="input-control text">
            <select name="control" style="width: 300px;" onchange="this.form.submit()">
		        <%
                strSQL = "SELECT tbl_group.id_group, tbl_plan.Semestr, tbl_disc.ID_disc, tbl_plan.control_form FROM (tbl_spec INNER JOIN tbl_group ON tbl_spec.id_spec = tbl_group.spec) INNER JOIN (tbl_disc INNER JOIN tbl_plan ON tbl_disc.ID_disc = tbl_plan.disc_name) ON tbl_group.id_group = tbl_plan.gr_name GROUP BY tbl_group.id_group, tbl_plan.Semestr, tbl_disc.ID_disc, tbl_plan.control_form HAVING (((tbl_group.id_group)=" + cstr(groupSearch) + ") AND ((tbl_plan.Semestr)='" + cstr(semestr) + "') AND ((tbl_disc.ID_disc)=" + cstr(discSearch) + "));"
                rs1.Open strSQL, con
                if rs1.EOF = false then
                set objId = rs1.Fields("control_form")
                set objName = rs1.Fields("control_form")
                do until rs1.EOF %>
                <%
                if objId = "ДЗ" then 
                    objName = "Дифференцированный зачет"
                elseif objId = "ТК" then 
                    objName = "Текущий контроль за семестр"
                end if
                %>
               <option value="<% response.write(objId) %>"<% if cstr(objId) = cstr(control) then response.Write("selected") %>><% response.write(objName) %></option> 
                <% rs1.MoveNext
                   loop
                end if
                   rs1.Close %>
	        </select>
            </div>
        </td>
        <% end if %>
	</tr>
    <tr>
        <% if groupSearch > 0 AND semestr > 0 AND discSearch > 0 AND control <> "" then %>
		<td align=right>Преподователь:</td>
		<td>
            <div class="input-control text">
            <select name="prepod" style="width: 300px;" onchange="this.form.submit()">
		        <%
                rs1.Open "SELECT tbl_plan.Semestr, tbl_disc.ID_disc, tbl_user.id_user, LEFT(tbl_user.user_fio, len(tbl_user.user_fio) - 4) + tbl_user.user_io AS full_fio FROM((tbl_disc INNER JOIN tbl_plan ON tbl_disc.ID_disc = tbl_plan.disc_name) INNER JOIN tbl_user ON tbl_plan.Prepod_name = tbl_user.id_user) GROUP BY tbl_plan.Semestr, tbl_disc.ID_disc, tbl_user.id_user, LEFT(tbl_user.user_fio, len(tbl_user.user_fio) - 4) + tbl_user.user_io HAVING (NOT (LEFT(tbl_user.user_fio, len(tbl_user.user_fio) - 4) + tbl_user.user_io LIKE '%/%') AND NOT (LEFT(tbl_user.user_fio, len(tbl_user.user_fio) - 4) + tbl_user.user_io LIKE 'ВакаВакансия') AND NOT (LEFT(tbl_user.user_fio, len(tbl_user.user_fio) - 4) + tbl_user.user_io LIKE 'Стрекаловский Н.В.')) AND (tbl_plan.Semestr = '" + cstr(semestr) + "') AND (tbl_disc.ID_disc = " + cstr(discSearch) + ");", con
                set objId = rs1.Fields("id_user")
                set objName = rs1.Fields("full_fio")
                if rs1.RecordCount = 1 then prepod = objId
                do until rs1.EOF %>
               <option value="<% response.write(objId) %>" <% if cstr(objId) = cstr(prepod) then response.Write(" selected") %>><% response.write(objName) %></option> 
                <% rs1.MoveNext
                   loop
                   rs1.Close %>
	        </select>
            </div>
        </td>
        <% end if %>
	</tr>
    <tr title="Позволяет выбрать до трёх преподавателей в ведомость.">
        <% if groupSearch > 0 AND semestr > 0 AND discSearch > 0 AND control <> "" then %>
        <td align=right>Несколько преподавателей:</td>
        <td>
            <label class='switch'><input type='checkbox' name="multiplePrepod" <% if multiplePrepod = "on" then response.Write("checked") %> onchange="this.form.submit()"><span class='check'></span></label>
        </td>
        <% end if %>
    </tr>
    <% if multiplePrepod = "on" then %>
    <tr>
        <td align=right>Переключение преподавателей:</td>
        <td>
            <label class="input-control checkbox">
                <input type="checkbox" name="prepod2enable" <% if prepod2enable = "on" then response.Write(" checked") %> onchange="this.form.submit()">
                <span class="check"></span>
                <span class="caption"> </span>
            </label>
            <label class="input-control checkbox">
                <input type="checkbox" name="prepod3enable" <% if prepod3enable = "on" then response.Write(" checked") %> onchange="this.form.submit()">
                <span class="check"></span>
                <span class="caption"> </span>
            </label>
        </td>
    </tr>
    <tr>
        <% if groupSearch > 0 AND semestr > 0 AND discSearch > 0 AND control <> "" then %>
		<td align=right>Преподователь (2):</td>
		<td>
            <div class="input-control text">
            <select name="prepod2" style="width: 300px;" onchange="this.form.submit()">
		        <%
                rs1.Open "SELECT tbl_user.id_user, Left([tbl_user].[user_fio],Len([tbl_user].[user_fio])-4)+[tbl_user].[user_io] AS full_fio, tbl_user.user_fio, tbl_user.user_io FROM tbl_user WHERE (((tbl_user.user_fio) Not Like '%/%' And (tbl_user.user_fio) Not Like 'Стрекаловский Н.В.') AND ((tbl_user.user_io) Not Like '%-%' And (tbl_user.user_io) Not Like '%Вакансия%' And (tbl_user.user_io) Not Like '%.%')) ORDER BY tbl_user.user_fio;", con
                set objId = rs1.Fields("id_user")
                set objName = rs1.Fields("full_fio")
                if rs1.RecordCount = 1 then prepod = objId
                do until rs1.EOF %>
               <option value="<% response.write(objId) %>" <% if cstr(objId) = cstr(prepod2) then response.Write(" selected") %>><% response.write(objName) %></option> 
                <% rs1.MoveNext
                   loop
                   rs1.Close %>
	        </select>
            </div>
        </td>
        <% end if %>
	</tr>
    <tr>
        <% if groupSearch > 0 AND semestr > 0 AND discSearch > 0 AND control <> "" then %>
		<td align=right>Преподователь (3):</td>
		<td>
            <div class="input-control text">
            <select name="prepod3" style="width: 300px;" onchange="this.form.submit()">
		        <%
                rs1.Open "SELECT tbl_user.id_user, Left([tbl_user].[user_fio],Len([tbl_user].[user_fio])-4)+[tbl_user].[user_io] AS full_fio, tbl_user.user_fio, tbl_user.user_io FROM tbl_user WHERE (((tbl_user.user_fio) Not Like '%/%' And (tbl_user.user_fio) Not Like 'Стрекаловский Н.В.') AND ((tbl_user.user_io) Not Like '%-%' And (tbl_user.user_io) Not Like '%Вакансия%' And (tbl_user.user_io) Not Like '%.%')) ORDER BY tbl_user.user_fio;", con
                set objId = rs1.Fields("id_user")
                set objName = rs1.Fields("full_fio")
                if rs1.RecordCount = 1 then prepod = objId
                do until rs1.EOF %>
               <option value="<% response.write(objId) %>" <% if cstr(objId) = cstr(prepod3) then response.Write(" selected") %>><% response.write(objName) %></option> 
                <% rs1.MoveNext
                   loop
                   rs1.Close %>
	        </select>
            </div>
        </td>
        <% end if %>
	</tr>
    <% end if %>
</table>
<br>

<button class="button success" type=submit onload="this.form.submit()"><span class="icon mif-loop2"></span> Обновить</button>
<a href="ved_zach_exam.asp?query=0"><button class="button danger" type=button><span class="icon mif-undo"></span> Сбросить</button></a>

<% if groupSearch > 0 AND semestr > 0 AND discSearch > 0 AND control <> "" then %>
<%
str_date = mid(request.Form("date"), 9, 2) + "." + mid(request.Form("date"), 6, 2) + "." +  mid(request.Form("date"), 1, 4)
if request.Form("date") = "" then str_date = today_date

strSQL = "SELECT TOP 1 id_group, group_name FROM tbl_group WHERE id_group=" + cstr(groupSearch) + ";"
rs2.Open strSQL, con
if rs2.EOF = false then
str_groupSearch = rs2.Fields("group_name")
end if
rs2.Close()

strSQL = "SELECT TOP 1 ID_disc, disc_name FROM tbl_disc WHERE ID_disc=" + cstr(discSearch) + ";"
rs2.Open strSQL, con
if rs2.EOF = false then
    str_discSearch = rs2.Fields("disc_name")
    if instr(str_discSearch, ")") - instr(str_discSearch, "(") > 0 AND hideBrackets = "on" then
        str_discSearch = left(str_discSearch, len(str_discSearch) - (instr(str_discSearch, ")") - instr(str_discSearch, "(")) - 2)
    end if
end if
rs2.Close()

if control = "ДЗ" then
    str_control = "Дифференцированный зачет"
elseif control = "ТК" then
    str_control = "Текущий контроль за семестр"
else
    str_control = control
end if

strSQL = "SELECT TOP 1 tbl_user.id_user, Left([user_fio],Len([user_fio])-4)+[user_io] AS full_fio FROM tbl_user WHERE tbl_user.id_user=" + cstr(prepod) + ";"
rs2.Open strSQL, con
if rs2.EOF = false then
str_prepod = rs2.Fields("full_fio")
end if
rs2.Close()

if prepod2enable = "on" then
    strSQL = "SELECT TOP 1 tbl_user.id_user, Left([user_fio],Len([user_fio])-4)+[user_io] AS full_fio FROM tbl_user WHERE tbl_user.id_user=" + cstr(prepod2) + ";"
    rs2.Open strSQL, con
    if rs2.EOF = false then
    str_prepod2 = rs2.Fields("full_fio")
    end if
    rs2.Close()
end if

if prepod3enable = "on" then
    strSQL = "SELECT TOP 1 tbl_user.id_user, Left([user_fio],Len([user_fio])-4)+[user_io] AS full_fio FROM tbl_user WHERE tbl_user.id_user=" + cstr(prepod3) + ";"
    rs2.Open strSQL, con
    if rs2.EOF = false then
    str_prepod3 = rs2.Fields("full_fio")
    end if
    rs2.Close()
end if

rs3.Open "SELECT tbl_group.id_group, tbl_plan.Semestr, tbl_student.id_student, tbl_disc.ID_disc, tbl_group.group_name, tbl_disc.disc_name, tbl_student.student_fio, tbl_student.student_nam, tbl_student.student_otch FROM (tbl_group INNER JOIN (tbl_disc INNER JOIN tbl_plan ON tbl_disc.ID_disc = tbl_plan.disc_name) ON tbl_group.id_group = tbl_plan.gr_name) INNER JOIN tbl_student ON tbl_group.id_group = tbl_student.group_name WHERE (((tbl_group.id_group)=" + cstr(groupSearch) + ") AND ((tbl_plan.Semestr)='" + cstr(semestr) + "') AND ((tbl_disc.ID_disc)=" + cstr(discSearch) + ")) ORDER BY tbl_student.student_fio;", con, 3, 3
'Формируем массив
dim students(3000,4) 'Массив на 3000 студентов
if rs3.RecordCount > 0 then
	ij = 1
	cnt = 1
	while rs3.EOF <> true
		students(ij,1) = rs3.Fields("id_student") 'ID студента
		students(ij,2) = left(ltrim(rs3.Fields("student_fio")), instr(ltrim(rs3.Fields("student_fio")) & "", " ")) + rs3.Fields("student_nam") + " " + rs3.Fields("student_otch") 'ФИО студента (без проблов слева)
		students(ij,3) = true 'Включён ли студент в ведомость
        if right(str_groupSearch, 2) = "ИС" and mid(str_groupSearch, 3, 1) = "/" then if mid(rs3.Fields("student_fio"), 1, 1) <> " " then students(ij,4) = "checked" else students(ij,4) = ""
		ij = ij + 1
		cnt = cnt + 1
		rs3.MoveNext
	wend
end if
%>
</form>
<form action="ved_zach_exam.asp?query=1" method=post>

<details style="margin-bottom: 15px">
<summary style="background: #2086bf; color: #fff; border-color: #2086bf; box-shadow: rgb(0 0 0 / 20%) 0px 3px 5px;"><span class="icon mif-pencil"></span> Редактировать передаваемые данные</summary>
<br>
<div class="input-control text" style="width: 125px"><input type="text" name="selectedDate" value="<%=str_date%>"></div> 
<div class="input-control text" style="width: 125px"><input type="text" name="groupSearch" value="<%=str_groupSearch%>"></div> 
<div class="input-control text" style="width: 32.5px"><input type="text" name="semestr" value="<%=semestr%>"></div><br>
<div class="input-control text" style="width: 300px"><input type="text" name="discSearch" value="<%=str_discSearch%>"></div><br>
<div class="input-control text" style="width: 300px"><input type="text" name="control" value="<%=str_control%>"></div><br>
<div class="input-control text" style="width: 300px"><input type="text" name="prepod" value="<%=str_prepod%>"></div><br>
<% if multiplePrepod = "on" then %>
<% if prepod2enable = "on" then %>
<div class="input-control text" style="width: 300px"><input type="text" name="prepod2" value="<%=str_prepod2%>"></div><br>
<% end if %>
<% if prepod3enable = "on" then %>
<div class="input-control text" style="width: 300px"><input type="text" name="prepod3" value="<%=str_prepod3%>"></div><br>
<% end if %>
<% end if %>
</details>

<table class="table striped hovered cell-hovered border bordered" style="width: 35vw">
<thead align=center style="font-weight: bold">
<tr><th>ID</th><th>Ф.И.О. экзаменующегося</th><th style="width:1%" title="Определяет, включён ли студент в ведомость">ВКЛ</th><% if right(str_groupSearch, 2) = "ИС" and mid(str_groupSearch, 3, 1) = "/" then %><th style="width:1%" title="Определяет, находится ли студент в подгруппе">ПГ</th><% end if %></tr>
</thead>
<tbody align=center>
<%
'Создаем таблицу
for i = 1 to cnt
	if students(i,1) > 0 then
		response.Write("<tr>")
		'Рисуем таблицу
		response.Write("<td>" & students(i,1) & "</td>")
        response.Write("<td ><p style='text-align:left; margin:0;'>" & students(i,2) & "<p></td>")
        response.Write("<td title='Определяет, включён ли студент в ведомость'><label class='switch'><input type='checkbox' name='students' value='" + students(i,2) + "' checked><span class='check'></span></label></td>")
        if right(str_groupSearch, 2) = "ИС" and mid(str_groupSearch, 3, 1) = "/" then response.Write("<td title='Определяет, находится ли студент в подгруппе'><label class='switch'><input type='checkbox' name='subgroup' value='" + students(i,2) + "' " + students(i,4) + "><span class='check'></span></label></td>")
        response.Write("</tr>")
	end if
next
%>
</tbody>
</table>
<br />
<button class="button success" type=submit><span class="icon mif-cogs"></span> Сформировать ведомость</button>
</form>
</center>
<% end if %>
<%
elseif query = 1 then 

' Распределение у групп с подгруппами
' Не пытайтесь это понять, я сам не понимаю как это работает, магия, бл*ть

' Вытаскиваем данные с формы
Set main = Request.Form("students")
Set subb = Request.Form("subgroup")
' Формируем массивы на основе принятых данных
ReDim students1(main.Count - 1)
ReDim students2(subb.Count - 1)

' Заполняем масивы
For i = 1 To main.Count
    students1(i - 1) = main(i)
Next
For i = 1 To subb.Count
    students2(i - 1) = subb(i)
Next

' Объявляем справочники
Set students3 = Server.CreateObject("Scripting.Dictionary")
Set students4 = Server.CreateObject("Scripting.Dictionary")

' Заполняем данными из первого массива первый справочник
For Each strFirst In students1
    Call students3.Add(strFirst, 0)
Next
' Определяем участников из подгруппы во второй справочник
For Each strSecond In students2
    If students3.Exists(strSecond) Then Call students4.Add(strSecond, 0)
Next
' Определяем участников основной группы обратно в первый справочник
For Each strSecond In students1
    If students4.Exists(strSecond) Then Call students3.Remove(strSecond)
Next

' Перегоняем данные из справочников в массив, так удобнее в плане вывода на страницу
ReDim students5(students3.count)
ReDim students6(students4.count)
a1 = students3.Keys
a2 = students4.Keys

for i=0 to students3.Count - 1
    students5(i) = a1(i)
next

for i=0 to students4.Count - 1
    students6(i) = a2(i)
next

' Отладочная инфа, вывод сырых данных, которые получаются после обработки словарей
'Response.Write Join(students3.Keys, "<br>")
'Response.Write "<br>#MAIN GROUP: " + cstr(students3.count) + "#<br>"
'Response.Write Join(students4.Keys, "<br>")
'Response.Write "<br>#SUB GROUP: " + cstr(students4.count) + "#<br>"
'Response.Write "#####TOTAL: " + cstr(students3.count + students4.count) + "#<br>"

today_date = request.Form("selectedDate")
if mid(today_date, 4, 2) = "01" then today_date = mid(today_date, 1, 2) + " января " + mid(today_date, 7, 4)
if mid(today_date, 4, 2) = "02" then today_date = mid(today_date, 1, 2) + " февраля " + mid(today_date, 7, 4)
if mid(today_date, 4, 2) = "03" then today_date = mid(today_date, 1, 2) + " марта " + mid(today_date, 7, 4)
if mid(today_date, 4, 2) = "04" then today_date = mid(today_date, 1, 2) + " апреля " + mid(today_date, 7, 4)
if mid(today_date, 4, 2) = "05" then today_date = mid(today_date, 1, 2) + " мая " + mid(today_date, 7, 4)
if mid(today_date, 4, 2) = "06" then today_date = mid(today_date, 1, 2) + " июня " + mid(today_date, 7, 4)
if mid(today_date, 4, 2) = "07" then today_date = mid(today_date, 1, 2) + " июля " + mid(today_date, 7, 4)
if mid(today_date, 4, 2) = "08" then today_date = mid(today_date, 1, 2) + " августа " + mid(today_date, 7, 4)
if mid(today_date, 4, 2) = "09" then today_date = mid(today_date, 1, 2) + " сентября " + mid(today_date, 7, 4)
if mid(today_date, 4, 2) = "10" then today_date = mid(today_date, 1, 2) + " октября " + mid(today_date, 7, 4)
if mid(today_date, 4, 2) = "11" then today_date = mid(today_date, 1, 2) + " ноября " + mid(today_date, 7, 4)
if mid(today_date, 4, 2) = "12" then today_date = mid(today_date, 1, 2) + " декабря " + mid(today_date, 7, 4)
groupSearch = request.Form("groupSearch")
discSearch = request.Form("discSearch")
control = request.Form("control")
prepod = request.Form("prepod")
prepod2 = request.Form("prepod2")
prepod3 = request.Form("prepod3")
%>
<div style="width: 647px; margin: 0 auto;">
<p style="margin: 0; padding: 0; float: left;"><button class="button" onclick="history.back()"><span class="icon mif-undo"></span> Вернуться назад</button></p><p style="margin: 0; padding: 0; float: right;"><button class="button primary" onclick="window.print()"><span class="icon mif-printer"></span> Печать</button></p><br><br><br>
</div>
</div>
<div id=printableArea>
<% if students3.count > 0 then %>
<center>
<p style="font-size: 11pt;">КОТЛАССКИЙ ФИЛИАЛ ФГБОУ ВО «ГУМРФ имени адмирала С.О. Макарова»<br>КОТЛАССКОЕ РЕЧНОЕ УЧИЛИЩЕ</p>
                    
<p style="font-size: 12pt;"><b>ЗАЧЕТНО-ЭКЗАМЕНАЦИОННАЯ ВЕДОМОСТЬ № <u>            </u></b></p>

<table align="center" class="info-table">
    <tbody style="font-size: 12pt;">
        <tr>
            <td style="width: 3.69cm;">Дата</td>
            <td style="font-weight: bold; width: 14.5cm; border-bottom: 1px solid; padding-left: 0.19cm; padding-right: 0.19cm;"><%=today_date%> года</td>
        </tr>
        <tr>
            <td>Группа</td>
            <td style="font-weight: bold; border-bottom: 1px solid; padding-left: 0.19cm; padding-right: 0.19cm;"><% if request.Form("subgroup").count = 0 then response.Write(groupSearch) else if request.Form("subgroup").count > 0 then response.Write(mid(groupSearch, 1, 2) + right(groupSearch, 3))%></td>
        </tr>
        <tr>
            <td>Дисциплина</td>
            <td style="font-weight: bold; border-bottom: 1px solid; padding-left: 0.19cm; padding-right: 0.19cm;"><%=discSearch%></td>
        </tr>
        <tr>
            <td>Форма контроля</td>
            <td style="font-weight: bold; border-bottom: 1px solid; padding-left: 0.19cm; padding-right: 0.19cm;"><%=control%></td>
        </tr>
        <tr>
            <td>Преподаватель</td>
            <td style="font-weight: bold; border-bottom: 1px solid; padding-left: 0.19cm; padding-right: 0.19cm;"><%=prepod%></td>
        </tr>
        <% if prepod2 <> "" then %>
        <tr>
            <td></td>
            <td style="font-weight: bold; border-bottom: 1px solid; padding-left: 0.19cm; padding-right: 0.19cm;"><%=prepod2%></td>
        </tr>
        <% end if %>
        <% if prepod3 <> "" then %>
        <tr>
            <td></td>
            <td style="font-weight: bold; border-bottom: 1px solid; padding-left: 0.19cm; padding-right: 0.19cm;"><%=prepod3%></td>
        </tr>
        <% end if %>
    </tbody>
</table>
<br>
<table align="center" class="main-table" border="1px">
    <thead style="font-size: 11pt;">
        <tr>
            <td style="width: 1.11cm; height: 1.35cm; vertical-align: middle;"><p style="padding: 0;margin: 0;text-align: center;">№<br>п/п</p></td>
            <td style="width: 8.1cm; height: 1.35cm; vertical-align: middle; text-align: center;">Ф. И. О. экзаменующегося</td>
            <td style="width: 3.49cm; height: 1.35cm; vertical-align: middle; text-align: center; word-wrap: normal;">№ экзаменационного билета</td>
            <td style="width: 2.46cm; height: 1.35cm; vertical-align: middle; text-align: center;">Оценка</td>
            <td style="width: 3.4cm; height: 1.35cm; vertical-align: middle; text-align: center;">Подпись экзаменатора</td>
        </tr>
        <tr class="numbers" style="text-align: center;">
            <td>1</td>
            <td>2</td>
            <td>3</td>
            <td>4</td>
            <td>5</td>
        </tr>
    </thead>
    <tbody style="font-size: 12pt;">
        <%
        For i = 1 To students3.count
            Response.Write "<tr>"
            Response.Write "<td style='text-align: center;'>" + cstr(i) + "</td>"
            Response.Write "<td style='padding-left: 0.19cm; padding-right: 0.19cm; text-align: left;'>" + cstr(students5(i - 1)) + "</td>"
            Response.Write "<td></td>"
            Response.Write "<td></td>"
            Response.Write "<td></td>"
            Response.Write "</tr>"
        Next
        %>
    </tbody>
</table>
</center>
<% end if %>
<% if students4.count > 0 then %>
<center>
<p style="font-size: 11pt;">КОТЛАССКИЙ ФИЛИАЛ ФГБОУ ВО «ГУМРФ имени адмирала С.О. Макарова»<br>КОТЛАССКОЕ РЕЧНОЕ УЧИЛИЩЕ</p>
                    
<p style="font-size: 12pt;"><b>ЗАЧЕТНО-ЭКЗАМЕНАЦИОННАЯ ВЕДОМОСТЬ № <u>            </u></b></p>

<table align="center" class="info-table">
    <tbody style="font-size: 12pt;">
        <tr>
            <td style="width: 3.69cm;">Дата</td>
            <td style="font-weight: bold; width: 14.5cm; border-bottom: 1px solid; padding-left: 0.19cm; padding-right: 0.19cm;"><%=today_date%> года</td>
        </tr>
        <tr>
            <td>Группа</td>
            <td style="font-weight: bold; border-bottom: 1px solid; padding-left: 0.19cm; padding-right: 0.19cm;"><%=mid(groupSearch,4,5)%></td>
        </tr>
        <tr>
            <td>Дисциплина</td>
            <td style="font-weight: bold; border-bottom: 1px solid; padding-left: 0.19cm; padding-right: 0.19cm;"><%=discSearch%></td>
        </tr>
        <tr>
            <td>Форма контроля</td>
            <td style="font-weight: bold; border-bottom: 1px solid; padding-left: 0.19cm; padding-right: 0.19cm;"><%=control%></td>
        </tr>
        <tr>
            <td>Преподаватель</td>
            <td style="font-weight: bold; border-bottom: 1px solid; padding-left: 0.19cm; padding-right: 0.19cm;"><%=prepod%></td>
        </tr>
        <% if prepod2 <> "" then %>
        <tr>
            <td></td>
            <td style="font-weight: bold; border-bottom: 1px solid; padding-left: 0.19cm; padding-right: 0.19cm;"><%=prepod2%></td>
        </tr>
        <% end if %>
        <% if prepod3 <> "" then %>
        <tr>
            <td></td>
            <td style="font-weight: bold; border-bottom: 1px solid; padding-left: 0.19cm; padding-right: 0.19cm;"><%=prepod3%></td>
        </tr>
        <% end if %>
    </tbody>
</table>
<br>
<table align="center" class="main-table" border="1px">
    <thead style="font-size: 11pt;">
        <tr>
            <td style="width: 1.11cm; height: 1.35cm; vertical-align: middle;"><p style="padding: 0;margin: 0;text-align: center;">№<br>п/п</p></td>
            <td style="width: 8.1cm; height: 1.35cm; vertical-align: middle; text-align: center;">Ф. И. О. экзаменующегося</td>
            <td style="width: 3.49cm; height: 1.35cm; vertical-align: middle; text-align: center; word-wrap: normal;">№ экзаменационного билета</td>
            <td style="width: 2.46cm; height: 1.35cm; vertical-align: middle; text-align: center;">Оценка</td>
            <td style="width: 3.4cm; height: 1.35cm; vertical-align: middle; text-align: center;">Подпись экзаменатора</td>
        </tr>
        <tr class="numbers" style="text-align: center;">
            <td>1</td>
            <td>2</td>
            <td>3</td>
            <td>4</td>
            <td>5</td>
        </tr>
    </thead>
    <tbody style="font-size: 12pt;">
        <%
        For i = 1 To students4.count
            Response.Write "<tr>"
            Response.Write "<td style='text-align: center;'>" + cstr(i) + "</td>"
            Response.Write "<td style='padding-left: 0.19cm; padding-right: 0.19cm; text-align: left;'>" + cstr(students6(i - 1)) + "</td>"
            Response.Write "<td></td>"
            Response.Write "<td></td>"
            Response.Write "<td></td>"
            Response.Write "</tr>"
        Next
        %>
    </tbody>
</table>
</center>
</div>
<% elseif students3.count = 0 and students4.count = 0 then %>
<div id=notPrintableArea>
<center>
<h3 style="color: #ce352c;">Вы не выбрали ни одного студетна!</h3>
<h4>Вернитесь на предыдущую страницу и выберите хотя бы одного студента!</h4>
</center>
</div>
<% end if %>
</div>
<div id=notPrintableArea>
<% end if %>
<br>
<center>
<a href="group_change.asp?go=1"><button type="button" class="button subinfo">Вернуться к выбору группы</button><br>
<a href="help/02_13.asp" ><button class="button success"><span class="icon mif-info"></span> Помощь</button></a>
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