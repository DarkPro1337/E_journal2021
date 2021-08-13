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
		<script src="js/jquery-3.1.0.min.js"></script>
		<script src="js/metro.min.js"></script>
		<title>�������������� ������� "����������� ������"</title>
	</head>
	<body>
	<div data-role="dialog" id="dialog_err" class="padding20 dialog info" style="display: none; text-align: left; width: auto; height: auto; width: 1000px; visibility: hidden; left: -700px; top: 349px;" data-close-button="true">
		<h1><font color=red>������ �����������. � ������� ��������!</font></h1>
		<p class="p">
			��������� ������ ����������� ������������ � �������������� �������, ��� ��� <span style="font-weight: bold;">��</span> �� ��������� ����� ���� <span style="font-weight: bold;">�����</span> - <span style="font-weight: bold;">������</span>.
		</p>
	</div>
		<table class="table border" style='width:90%; margin-top: 15px;' align=center> 
			<tr>
				<td>
					<!-- #include file="header.asp" -->
				</td>
			</tr>
				<%
					if session("user") = "" or session("user") = "�������" or session("user") = 0 then response.Redirect ("404.asp")
					login = request.Form("login")
					pass  = request.Form("pass")
                    session.Abandon()
				%>
			<tr>
				<td>
					<center>
						<b>����� ������</b><br><br>
						<% if err = 1 then %>
							<br><font color=red size=5>�������� ��������� ������ � ������, ���������� ��� ���!</font>
						<% end if %>
						<form name="frm_start" action="change_pass.asp" method="post" style="width: 350px;">
							<input type="hidden" autocomplete='off' name="Username" id="uname" value = "<%=login%>" required>
							<input type="hidden" autocomplete='on' name="Password" id="uname" value = "<%=pass%>" required>

                            <div class="panel">
                                <div class="heading">
                                    <span class="icon mif-warning"></span>
                                    <span class="title">���������� � ������</span>
                                </div>
                                <div class="content padding10" style="text-align: justify;">
                                    ����� ������ ������ ���� ������� �� ����� 5 � �� ����� 7 ��������, ��������� ��������� � �������� ��������� �����, � ��������� ���� ����������� �������, ���� �����.
                                </div>
                            </div>
                            
                            <div class="input-control modern text iconic" style='width: 300px'>
								<input type="password" autocomplete='on' name="New_Password" id="Password1" value = "" required pattern="(?=^.{5,7}$)((?=.*\d)|(?=.*\W+))(?![.\n])(?=.*[A-Z])(?=.*[a-z]).*$">
								<span class="label">����������, ������� ����� ������</span>
								<span class="informer"></span>
								<span class="icon mif-key fg-blue"></span>
								<button class="button helper-button clear" tabindex="-1" type="button" onClick='$("#new_password").val("");'><span class="mif-cross"></span></button>
							</div>
							<button id="change_pass" class='button primary' value="" ><span class="icon mif-key"></span>�������� ������</button>
						</form>
                        <a href="index.asp"><button id="back" class='button' value="" ><span class="icon mif-undo"></span>���������� � �����������</button></a><br>
						<a href="help/01_1.asp"><button class="button success"><span class="icon mif-info"></span> ������</button></a>
					</center>
				</td>
			</tr>
			<tr>
				<td>
				</td>
			</tr>
		</table>
	<script>
            $("img").mousedown(function(){
            return false;
            });
    </script>
	</body>
</html>
	