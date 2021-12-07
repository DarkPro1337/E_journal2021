<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01//EN" "http://www.w3.org/TR/html4/strict.dtd">
<html>
	<head>
		<meta content="text/html; charset=Windows-1251" http-equiv="content-type">
		<link rel="shortcut icon" href="images/favicon.ico" /> 
		<link rel="stylesheet" href="../css/metro.css">
		<link rel="stylesheet" href="../css/metro-colors.css">
		<link rel="stylesheet" href="../css/metro-icons.css">
		<link rel="stylesheet" href="../css/metro-responsive.css">
		<link rel="stylesheet" href="../css/metro-rtl.css">
		<link rel="stylesheet" href="../css/metro-schemes.css">
		<link rel="stylesheet" href="../css/style.css">
		<script src="../js/jquery-3.1.0.min.js"></script>
		<script src="../js/metro.min.js"></script>
		<title>ИС "Электронный журнал". Файл справки. Смена пароля пользователя.</title>
	</head>
	
	<body>
		<table class="table border" style='width:90%; margin-top: 15px;' align=center> 
			<tr>
				<td colspan=2>
					<!--#include file="header.asp"-->
				</td>
			</tr>
			<tr>
				<td width=30% valign=top>
					<!--#include file=help_menu.inc-->
				</td>
				<td style="vertical-align: top;">
					<h2>Смена пароля пользователя</h2>
            
					<p class="helptext">За каждой учётной записью пользователей закреплён пароль, изначально созданный администратором информационной системы, для обеспечения безопасности учётных записей пользователей. Для информационной системы <i>"Электронный журнал"</i> предусмотрена смена паролей пользователей раз в год, за месяц до окончания срока службы пароля пользователю будет сообщено о скором истечении срока пароля учётной записи, после истечения срока пароля учётной записи доступ к учётной записи будет заблокирован. Для того чтобы не потерять доступ к своей учётной записи пользователям рекомендуется не игнорировать окно уведомленя о истечении срока пароля, а поменять его сразу же! Это обеспечит безопасность вашей учётной записи и информационной системе в целом.</p>
					
					<p class="helptext">Рассмотрим уведолмения о истечении срока пароля подробнее.<br>За 30 дней до истечения срока пароля учётной записи вам будет отображено системное сообщение о необходимости смены пароля. В области (1) отображается текст уведомления, в нём содержится количество дней до потери доступа к учётной записи, в области (2) находится кнопка <i>Сменить пароль</i>, она перенаправит вас на страницу смены пароля, в области (3) располагается кнопка закрытия системного уведомления. Данное сообщение отображается на всех страницах электронного журнала, но его можно закрыть</p>

					<center><img src="../images/pass_update2.png" /><br /></center>

					<p class="helptext">По истечению срока службы пароля вам будет отображено системное уведомление о потери доступа к учётной записи и необходимости сменить пароль. В области (1) отображется текст уведолмения, в области (2) находится кнопка <i>Сменить пароль</i>, она перенаправит вас на страницу смены пароля, в области (3) находится кнопка <i>Выйти</i>, она завершит вашу сессию и перенаправит вас на страницу авторизации.</p>

					<center><img src="../images/pass_update1.png" /><br /></center>

					<p class="helptext">На странице смены пароля располагается лишь одно поле - для ввода нового пароля согласно требованиям к паролю. Он должен быть длиной не менее 5 и не более 7 символов, содержать заглавные и строчные латинские сиволы, и содержать либо специальные символы, либо числа. В области (1) располагается панель требований к паролю, в области (2) находится поле для ввода нового пароля, в области (3) находится кнопка <i>Сменить пароль</i>, она выполнит смену пароля, и перенаправит вас на главную страницу электронного журнла, в области (4) находится кнопка <i>Вернуться к авторизации</i>, на случай если вам нужно вернуться на страницу авторизации.</p>

					<center><img src="../images/pass_update3.png" /><br /></center>
				</td>
			</tr>
			<tr>
				<td colspan="2">
					<center>
					<script>
						(function($) {
							$(function() {
								$('#up').click(function() {
								$('body,html').animate({scrollTop:0},500);
							return false;
							})
						})
					})(jQuery)
					</script>
						<a href="#" href="#" onclick="javascript:history.back(); return false;"><button class="button primary"><span class="icon mif-undo"></span> Назад </button></a>
						<a href="#" ><button class="button success" id="up"><span class="icon mif-arrow-up"></span> Вверх </button></a>
						<a href="../exit.asp" ><button class="button danger"><span class="icon mif-exit"></span> Выход </button></a>
					</center>
				</td>
			</tr>
		</table>
	</body>
</html>