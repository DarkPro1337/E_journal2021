<meta content="text/html; charset=windows-1251" http-equiv="content-type" />
<link rel="shortcut icon" href="images/favicon.ico" />
<head>
	<meta content="text/html; charset=Windows-1251" http-equiv="content-type">
	<link rel="shortcut icon" href="images/favicon.ico"> 
	<link rel="stylesheet" href="css/metro.css">
	<link rel="stylesheet" href="css/metro-colors.css">
	<link rel="stylesheet" href="css/metro-icons.css">
	<link rel="stylesheet" href="css/metro-responsive.css">
	<link rel="stylesheet" href="css/metro-rtl.css">
	<link rel="stylesheet" href="css/metro-schemes.css">
	<link rel="stylesheet" href="css/metro-student.css">
	<script src="js/jquery-3.1.0.min.js"></script>
	<script src="js/metro.min.js"></script>
	<title>�� "����������� ������". ���������� ��������</title>
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
<h4>�� ����������� �������� �������, �� �������?</h4>
<script>
    function ShowHideMiniLoader(operation){
        if (operation == "hide"){
            document.getElementById("mini_loader").style.display = "none";
        } else {
            document.getElementById("mini_loader").style.display = "block";    
        }
    }
</script>
<a href="rate_update.asp?confirm=0"><button class="button primary" onclick="ShowHideMiniLoader('show')">�����������<span class="mif-spinner3 mif-ani-spin" id="mini_loader" style="color: #fff; display: none; float: right; margin-left: 5px;"></span></button></a> <a href="group_change.asp"><button class="button danger">��������� �����</button></a>
</center>
<%
elseif confirm = 0 then
'��������� ����������� � ��
Set con=Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.RecordSet")

strdbpath=server.mappath("base.mdb")
con.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strdbpath
rs.Open "RATING", con, 3, 3
'��������� ������
dim students_rate(3000,5) '������ �� 3000 �������������
if rs.RecordCount > 0 then
	ij = 1
	cnt = 1
	while rs.EOF <> true
		students_rate (ij,1) = rs.Fields(3) 'ID ������������
		students_rate (ij,2) = rs.Fields(1) '������� �������
		students_rate (ij,3) = rs.Fields(4)	'������ �������
		ij = ij + 1
		cnt = cnt + 1
		rs.MoveNext
	wend
end if
'������� ����������� ��� ��������
strdbpath=server.mappath("base.mdb")
strDbConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="& strdbpath & ";"
Set objConn = Server.CreateObject("ADODB.Connection")
objConn.Open(strDbConnection)
%>
<center>
<table class="table striped hovered cell-hovered border bordered">
<thead align=center style="font-weight: bold">
<tr><td>������� �������</td><td>������ �������</td><td>ID</td><td>��� �������</td></tr>
</thead>
<tbody align=center>
<%
'������� �������
for i = 1 to cnt '������ ����������� ��� ������� ��������
	if students_rate (i,1) > 0 then '���� ������������ � ������� ����� ���� ID
		'�������������� ������
		response.Write("<tr>")
		newrate = students_rate (i,2) + students_rate (i,3)
		id_st = students_rate (i,1)
		strSQL = "UPDATE tbl_student SET rating_old = '"& newrate &"' WHERE id_student = "& id_st '�������������� ������
		'������ �������
		response.Write("<td>" & round(students_rate (i,2), 2) & " </td><td> " & students_rate (i,3) & " </td><td> " & id_st & " </td><td> " & strSQL & "</td>")
		'��������� ������
		objConn.Execute(strSQL)
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
<a href="disc_change.asp?go=1"><button type="button" class="button subinfo">��������� � ������ ��������� � ����������</button><br>
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