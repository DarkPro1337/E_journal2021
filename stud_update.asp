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
	<title>�� "����������� ������". ������� ���������</title>
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
'������ �� ���������
if session("user") = "" or session("user") = "�������" or session("user") = 0 then response.Redirect ("404.asp")

confirm = request.querystring("confirm")

if confirm = 1 then
%>
<center>
<h4>�� ����������� ��������� ������� �������� ������ ������ �� ���� ����, �� �������?</h4>
<script>
    function ShowHideMiniLoader(operation){
        if (operation == "hide"){
            document.getElementById("mini_loader").style.display = "none";
        } else {
            document.getElementById("mini_loader").style.display = "block";    
        }
    }
</script>
<a href="stud_update.asp?confirm=0"><button class="button primary" onclick="ShowHideMiniLoader('show')">�����������<span class="mif-spinner3 mif-ani-spin" id="mini_loader" style="color: #fff; display: none; float: right; margin-left: 5px;"></span></button></a> <a href="group_change.asp"><button class="button danger">��������� �����</button></a>
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

'��������� ����������� � ��
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
        '������� ����������� ��� ��������
        strDbConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="& strdbpath & ";"
        Set objConn = Server.CreateObject("ADODB.Connection")
        objConn.Open(strDbConnection)
        strSQL = "INSERT INTO tbl_group (group_name, group_ruk, spec, nb_edit_data) VALUES ('" & group_name & "', " & group_ruk & ", " & spec & ", #" & nowDate & "#);" '�������������� ������
        '��������� ������
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

'��������� ������
dim students(3000,5) '������ �� 3000 ���������
if rs.RecordCount > 0 then
	ij = 1
	cnt = 1
	while rs.EOF <> true
		students(ij,1) = rs.Fields(0) 'ID
		students(ij,2) = rs.Fields(1) '���
		students(ij,3) = rs.Fields(2) '������� ������
        students(ij,4) = 0 '����� ������
        rs3.Open "SELECT tbl_student.id_student, tbl_student.student_fio, tbl_group.id_group, tbl_group.group_name FROM tbl_group INNER JOIN tbl_student ON tbl_group.id_group = tbl_student.group_name WHERE (((tbl_group.id_group)=" & students(ij,3) &"));", con, 3
            '������ ������������� ��
            groupFrom = "11-��" ' ������ �� ������� ���������
            groupTo   = "21-��" ' ������ � ������� ���������
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

            groupFrom = "21-��" ' ������ �� ������� ���������
            groupTo   = "31-��" ' ������ � ������� ���������
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

            groupFrom = "31-��" ' ������ �� ������� ���������
            groupTo   = "41-��" ' ������ � ������� ���������
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

            groupFrom = "41-��" ' ������ �� ������� ���������
            if rs3.Fields(3) = groupFrom then
                students(ij,4) = -1
            end if
            '/��
            
            ' ������ ������������� ��
            groupFrom = "11-��" ' ������ �� ������� ���������
            groupTo   = "21-��" ' ������ � ������� ���������
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

            groupFrom = "21-��" ' ������ �� ������� ���������
            groupTo   = "31-��" ' ������ � ������� ���������
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

            groupFrom = "31-��" ' ������ �� ������� ���������
            groupTo   = "41-��" ' ������ � ������� ���������
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

            groupFrom = "41-��" ' ������ �� ������� ���������
            if rs3.Fields(3) = groupFrom then
                students(ij,4) = -1
            end if
            '/��

            ' ������ ������������� ��
            groupFrom = "11-��" ' ������ �� ������� ���������
            groupTo   = "21/12-��" ' ������ � ������� ���������
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

            groupFrom = "21/12-��" ' ������ �� ������� ���������
            groupTo   = "31/22-��" ' ������ � ������� ���������
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

            groupFrom = "31/22-��" ' ������ �� ������� ���������
            groupTo   = "41/32-��" ' ������ � ������� ���������
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

            groupFrom = "41/32-��" ' ������ �� ������� ���������
            if rs3.Fields(3) = groupFrom then
                students(ij,4) = -1
            end if
            '/��

            ' ������ ������������� ��
            groupFrom = "11-��" ' ������ �� ������� ���������
            groupTo   = "21-��" ' ������ � ������� ���������
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

            groupFrom = "12-��" ' ������ �� ������� ���������
            groupTo   = "22-��" ' ������ � ������� ���������
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

            groupFrom = "21-��" ' ������ �� ������� ���������
            groupTo   = "31-��" ' ������ � ������� ���������
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

            groupFrom = "22-��" ' ������ �� ������� ���������
            groupTo   = "32-��" ' ������ � ������� ���������
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

            groupFrom = "31-��" ' ������ �� ������� ���������
            groupTo   = "41-��" ' ������ � ������� ���������
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

            groupFrom = "32-��" ' ������ �� ������� ���������
            groupTo   = "42-��" ' ������ � ������� ���������
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

            groupFrom = "41-��" ' ������ �� ������� ���������
            if rs3.Fields(3) = groupFrom then
                students(ij,4) = -1
            end if

            groupFrom = "42-��" ' ������ �� ������� ���������
            groupTo   = "52-��" ' ������ � ������� ���������
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

            groupFrom = "52-��" ' ������ �� ������� ���������
            if rs3.Fields(3) = groupFrom then
                students(ij,4) = -1
            end if
            '/��

            ' ������ ������������� ��
            groupFrom = "11-��" ' ������ �� ������� ���������
            groupTo   = "21-��" ' ������ � ������� ���������
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

            groupFrom = "21-��" ' ������ �� ������� ���������
            groupTo   = "31-��" ' ������ � ������� ���������
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

            groupFrom = "31-��" ' ������ �� ������� ���������
            groupTo   = "41-��" ' ������ � ������� ���������
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

            groupFrom = "41-��" ' ������ �� ������� ���������
            if rs3.Fields(3) = groupFrom then
                students(ij,4) = -1
            end if
            '/��
            
        rs3.Close
		ij = ij + 1
		cnt = cnt + 1
		rs.MoveNext
	wend
end if

if len(request.Form) > 0 then
    if request.Form("update") = "1" then
        '��������� ����������� ��, �������� � ����� �� �� �������
        dim fs
        set fs=Server.CreateObject("Scripting.FileSystemObject")
        dim CurrentDirectory
        CurrentDirectory = Server.MapPath(".")
        fs.CopyFile CurrentDirectory + "/base.mdb", CurrentDirectory + "/base_backup.mdb"
        set fs = nothing

        '������� ����������� ��� ��������
        strDbConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="& strdbpath & ";"
        Set objConn = Server.CreateObject("ADODB.Connection")
        objConn.Open(strDbConnection)

        strSQL = "DELETE FROM tbl_itog" '������ ������� ����
        objConn.Execute(strSQL)
        strSQL = "DELETE FROM tbl_journal" '������ ������� ������
        objConn.Execute(strSQL)
        strSQL = "DELETE FROM tbl_journal_del" '������ ������� ������
        objConn.Execute(strSQL)
        strSQL = "DELETE FROM tbl_zan" '������ ������� ������
        objConn.Execute(strSQL)

        for i = 1 to cnt '������ ����������� ��� ������� ��������
	        if students(i,1) > 0 then '���� ������������ � ������� ����� ���� ID
                if students(i, 4) > 0 then '���� ������ �������� �� ��������� � ����������, �.�. -1 �� ����� ��������� ��� �� ��������� ����
                    strSQL = "UPDATE tbl_group INNER JOIN tbl_student ON tbl_group.id_group = tbl_student.group_name SET tbl_student.group_name = " & students(i,4) & " WHERE (((tbl_student.id_student)=" & students(i,1) & "));" '�������������� ������
                    objConn.Execute(strSQL)
                elseif students(i,4) = -1 then
                    strSQL = "DELETE FROM tbl_student WHERE (((tbl_student.id_student)=" & students(i,1) & "));" '�������������� ������
                    objConn.Execute(strSQL)
                end if
            end if
        next

        response.Redirect("stud_update.asp?confirm=0") ' ��������� ��������� ��������� ��������� ��-�� ���������� �������� ����� POST �������, ������� �����, �� ��������� ��������.
    end if
end if
%>
<center>
<% if group_missmatch > 0 then %>
<h1 style="color: Red; font-weight: bold">��������!</h1>
<h4>��� ���������� ������� ������� ���� ������� ������ ������������� ��������� ����� ���������, ������� �������� ������������ ������ � ��!</h2>
<p><a href=#missmatch>���������� ������: <% response.Write(group_missmatch) %></a></p>


<details style="margin-bottom: 15px">
<summary style="background: #2086bf; color: #fff; border-color: #2086bf; box-shadow: rgb(0 0 0 / 20%) 0px 3px 5px;"><span class="icon mif-wrench"></span> �������� ������������ ������ � ��</summary>

<form name="addNew" action="stud_update.asp" method="post" style="margin-bottom: 0">
�������� ������
    <div class="input-control text" style="width: 75px">
    <input type=text name="group_name" placeholder="11-��" required pattern="[0-9]{2}-[�-ߨ]{2}">
    </div>
	<div class="input-control text" style="width: 150px">
	<select name="group_ruk"  required>
        <option disabled selected value>�������������</option>
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
        <option disabled selected value>��������</option>
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
	<button type="submit" class="button primary"><span class="icon mif-pencil"></span>���������� ������</button> <button type="reset" class="button danger"><span class="icon mif-undo"></span>�������</button>
</form>

</details>
<% elseif group_missmatch = 0 and groupSearch = "" then %>
<form name="studUpdate" action="stud_update.asp" method="post" style="margin-bottom: 0; width: 50vw">
<h4 style="color: #60a917;"><span class="icon mif-checkmark"></span> ������ �������������� � �� �� ���� ����������, �� ������ � �������� ���������, ���������?</h4>
<h6 style="color: Gray;">� ���������� ���������� ����� �������� ����� ������� ��������� ����� �� (base_backup.mdb � ����� �� �� �������), � ������ ������������� ������ � ������ ������������ ������� ����� ����� ��������� � ���������� ����� ��.</h6>
<input type="hidden" name="update" value=1 />
<button type="submit" class="button success"><span class="icon mif-magic-wand"></span>����������� ���������</button><br><br>
</form>
<%
elseif groupSearch <> "" then
%>
<h6 style="color: Gray;">�� ���������� � ������ ����������, ��� �������� � �������� ��������� �������� ������ �������� ������ "�����".</h6>
<%
end if 

'���������, ���������� �� ����� �� � ����� �� �� �������
dim f
set fs=Server.CreateObject("Scripting.FileSystemObject")
CurrentDirectory = Server.MapPath(".")
if fs.FileExists(CurrentDirectory + "/base_backup.mdb") then
    set f=fs.GetFile(CurrentDirectory + "/base_backup.mdb")
    %>
    <form name="backup" action="stud_update.asp" method="post" style="margin-bottom: 0; width: 50vw">
    <h4 style="color: #2086bf;"><span class="icon mif-notification"></span> ���� ���������� ��������� ����� �� �� <% Response.Write(f.DateLastModified) %></h4>
    <input type="hidden" name="useBackup" value=1 />
    <button type="submit" class="button primary"><span class="icon mif-undo"></span>����������� �� ����������� ����� ��</button><br><br>
    </form>
    <%
end if
set f=nothing
set fs = nothing
%>

<form name="filter" action="stud_update.asp" method="get" style="margin-bottom: 0">
���������� �� ������
	<div class="input-control text" style="width: 125px">
	<select name="groupSearch" required>
        <option disabled selected value>��������...</option>
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
    <button type="submit" class="button primary"><span class="icon mif-filter"></span>���������</button> <a href="?confirm=0"><button type="button" class="button danger"><span class="icon mif-undo"></span>�������</button></a><br><br>
</form>
<table class="table striped hovered cell-hovered border bordered" style="width: 50vw">
<thead align=center style="font-weight: bold">
<tr><th>ID</th><th>���</th><th>������� ������</th><th>����� ������</th></tr>
</thead>
<tbody align=center>
<%
for i = 1 to cnt '������ ����������� ��� ������� ��������
	if students(i,1) > 0 then '���� ������������ � ������� ����� ���� ID
		'�������������� ������
		response.Write("<tr>")
		
		'������ �������
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
            response.Write("<td style='color: green;'>��������</td>")
        elseif students(i,4) = -2 then
            response.Write("<td title='��� ������ ���������� � ��!'><a name='missmatch' style='color: red; font-weight: bold'>" + students(i,5) + "</a></td>")
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
<a href="group_change.asp?go=1"><button type="button" class="button subinfo">��������� � ������ ������</button><br>
<a href="help/03_2.asp" ><button class="button success"><span class="icon mif-info"></span> ������</button></a>
<a href="exit.asp"><button class="button danger" ><span class="icon mif-exit"></span> �����</button></a>
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