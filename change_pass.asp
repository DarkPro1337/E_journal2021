<%
'---------------------------------------------------
'Проверка на тип аутентификации
'---------------------------------------------------
select case request.form("pr")
	case "journal"
		if session("user") = "" then response.Redirect ("404.asp")
		Set Conn = Server.CreateObject("ADODB.Connection") 
		Set rs_prepod = Server.CreateObject("ADODB.Recordset")
		Set rs_group = Server.CreateObject("ADODB.Recordset")
		Set rs_zav_otdel = Server.CreateObject("ADODB.Recordset")
		strDBPath = Server.MapPath("base.mdb")
		Conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath
		strSQL_prepod = "SELECT tbl_disc.disc_name, tbl_plan.ID_plan, tbl_user.user_fio, tbl_plan.Semestr, tbl_group.group_name FROM tbl_user INNER JOIN (tbl_group INNER JOIN (tbl_disc INNER JOIN tbl_plan ON tbl_disc.ID_disc = tbl_plan.disc_name) ON tbl_group.id_group = tbl_plan.gr_name) ON tbl_user.id_user = tbl_plan.Prepod_name WHERE (((tbl_plan.ID_plan)=" & request.form("Select1") & ") AND ((tbl_plan.Semestr)='"& session ("sem") &"') AND ((tbl_group.group_name)='" & session ("gr") &"')) ORDER BY tbl_disc.disc_name DESC;"
		strSQL_group = "SELECT tbl_group.id_group, tbl_group.group_name, tbl_user.user_fio, tbl_spec.spec_name FROM tbl_spec INNER JOIN (tbl_user INNER JOIN tbl_group ON tbl_user.id_user = tbl_group.group_ruk) ON tbl_spec.id_spec = tbl_group.spec WHERE (((tbl_group.group_name)='" & session("gr") & "') AND ((tbl_user.user_fio)='" & session("user_fio") & "'));"
		strSQL_zav_otdel = "SELECT tbl_spec.id_spec, tbl_spec.spec_number, tbl_spec.spec_name, tbl_user.user_fio, tbl_group.group_name FROM (tbl_user INNER JOIN tbl_spec ON tbl_user.id_user = tbl_spec.zav_otdel) INNER JOIN tbl_group ON tbl_spec.id_spec = tbl_group.spec WHERE (((tbl_user.user_fio)='" & session("user_fio") & "') AND ((tbl_group.group_name)='" & session ("gr") &"'));"
		rs_prepod.Open strSQL_prepod, Conn, adOpenStatic
		rs_prepod.movefirst
		session("disc") = rs_prepod.fields(0)
		session("auth") = false
		if request.form("pr_journal_read") = "true" then response.redirect("journal_read.asp")
		if session("user_fio") = "Студент" then response.redirect("journal_read.asp")
		if session("type_root") = 1 or session("type_edit_journal") = 1 then
			session("auth") = true
			response.redirect("journal.asp")
		else
			rs_zav_otdel.Open strSQL_zav_otdel, Conn, adOpenStatic
			on error resume next
			rs_zav_otdel.movefirst
			if err.number = "0" then
				session("auth") = true
				response.redirect("journal.asp")
			else
				rs_group.Open strSQL_group, Conn, adOpenStatic
				on error resume next
				rs_group.movefirst
				if err.number = "0" then
					session("auth") = true
					response.redirect("journal.asp")
				else
					for i = 1 to len (rs_prepod.fields(2))
						on error resume next
						if mid(rs_prepod.fields(2), i, len(session("user_fio"))) = session("user_fio") then
							session("auth") = true
							response.redirect("journal.asp")
							exit for
						end if
					next
					response.redirect("journal_read.asp")
				end if
			end if
		end if
end select
'---------------------------------------------------
'
'---------------------------------------------------
go = request.QueryString ("go")
if go=1 then 
fio = session("fio")
pass = session("pass")
session("gr")= ""
session("sem")=""
else
fio = request ("Username")
pass = request("Password")
new_pass = request("New_Password")
session ("fio") = fio
session ("pass")= pass
end if
'---------------------------------------------------
'Защита от SQL-инъекции
'---------------------------------------------------
for i=1 to len(fio)
if len(fio)-6 > 0 then
if mid(fio,i,6)="SELECT" or mid(fio,i,6)="INSERT" or mid(fio,i,6)="DELETE" or mid(fio,i,5)="UNION" then 
response.redirect("pass_update.asp?err=1")
end if
end if
if len (fio)-3 > 0 then
if mid(fio,i,3)="AND" or mid(fio,i,3)="XOR" then
response.Redirect ("pass_update.asp?err=1")
end if
end if
if mid(fio,i,1)=";" then response.Redirect ("pass_update.asp?err=1")
next
'----------------------------------------------------
'Подключение БД
'----------------------------------------------------
Set Conn = Server.CreateObject("ADODB.Connection") 
Set RS = Server.CreateObject("ADODB.Recordset") 
strDBPath = Server.MapPath("base.mdb")
Conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
strDBPath
'----------------------------------------------------
'Авторизация
'----------------------------------------------------
'----------------------------------------------------
'Проверка остальных пользователей, проверка прав администратора
'----------------------------------------------------
strSQL = "SELECT tbl_user.id_user, tbl_user.user_fio, tbl_user.password, tbl_user.user_io, tbl_user.type_prepod, tbl_user.type_kl_rukovod, tbl_user.type_zav_otdel, tbl_user.type_administr, tbl_user.type_edit_journal, tbl_user.type_root, tbl_user.type_orsod, tbl_user.password_date FROM tbl_user WHERE (((tbl_user.user_fio)='" & fio &"'));"
on error resume next
RS.Open strSQL, Conn, adOpenStatic
if not err = 0 then response.Redirect ("bd_error.asp")
if rs.eof = true then response.Redirect ("pass_update.asp?err=1")
if rs.fields(2) = pass then
    session("user") = rs.Fields(0)
    session("user_io") = rs.Fields(3)
    session("user_fio") = rs.Fields(1)
    session("pass_date") = rs.Fields(11)
	session("type_prepod")= 0
	session("type_kl_rukovod")= 0
	session("type_zav_otdel")= 0
	session("type_administr")= 0
	session("type_edit_journal") = 0
	session("type_root")= 0
	session("type_orsod") = 0
	if rs.Fields(9)=true then
		session("type_root") = 1
		session("type_prepod") = 1
		session("type_kl_rukovod") = 1
		session("type_zav_otdel") = 1
		session("type_administr") = 1
		session("type_edit_journal") = 1
		session("type_orsod") = 1
	else
		if rs.Fields(4)=true then session("type_prepod") = 1
		if rs.Fields(5)=true then session("type_kl_rukovod") = 1
		if rs.Fields(6)=true then session("type_zav_otdel") = 1
		if rs.Fields(7)=true then session("type_administr") = 1
		if rs.Fields(8)=true then session("type_edit_journal") = 1
		if rs.Fields(10)=true then session("type_orsod") = 1
	end if
    rs.Close()
    today_short = mid(date(), 4, 2) + "/" + mid(date(), 1, 2) + "/" + mid(date(), 7, 4)

    strSQL = "UPDATE tbl_user SET tbl_user.[password] = '" + new_pass + "', tbl_user.password_date = #" + today_short + "# WHERE (((tbl_user.user_fio)='" + fio + "'));"
	Conn.Execute(strSQL)

    strSQL = "SELECT tbl_user.id_user, tbl_user.user_fio, tbl_user.password, tbl_user.user_io, tbl_user.type_prepod, tbl_user.type_kl_rukovod, tbl_user.type_zav_otdel, tbl_user.type_administr, tbl_user.type_edit_journal, tbl_user.type_root, tbl_user.type_orsod, tbl_user.password_date FROM tbl_user WHERE (((tbl_user.user_fio)='" & fio &"'));"
    rs.Open strSQL, Conn, adOpenStatic
    session("pass_date") = rs.Fields(11)
    rs.close()
    response.Redirect("group_change.asp")
    else 
    response.Redirect ("pass_update.asp?err=1")
end if
'-----------------------------------------------------
%>