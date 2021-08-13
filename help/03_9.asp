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
		<title>ИС "Электронный журнал". Файл справки. Страница о студенте/курсанте.</title>
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
					<h2>Страница о студенте/курсанте</h2>
            
					<p class="helptext">Страница о студенте предназначена для просмотра оценок и отсутсвий по дисциплинам конкретного студента. Эта страница представляет собой сводную таблицу с перечнем дисциплин и оценок или пропусками по ней. ФИО студента и его группа отображается в области (1), в области (2) содержится перечень всех дисциплин выбранного студента, в области (3) содержится перечень всех оценок и пропусков.</p> <br />
					
					<center><img src="../images/stud1.png" /><br /></center>
					
					<p class="helptext">При наведении на оценку во всплывающей подсказе будет отображены детали интересующей оценки: дата получения оценки, вид работы и описание, если оно указано.</p> <br />
					
					<center><img src="../images/stud2.png" /><br /></center>
					
					<p class="helptext">Для оценок применяется следующее цветовое оформление:
					<ul class="simple-list small-bullet black-bullet">
					<li> Розовый цвет - оценка отлично;</li>
					<li> Зелёный цвет - оценка хорошо;</li>
					<li> Оранжевый цвет - оценка удовлетворительно;</li>
					<li> Синий цвет - оценка неудовлетворительно.</li>
					</ul>
					Для пропусков применяется следующее цветовое оформление:
					<ul class="simple-list small-bullet black-bullet">
					<li> Красный цвет - неуважительный пропуск;</li>
					<li> Бирюзовый цвет - уважительный пропуск;</li>
					<li> Серый цвет - необработанный пропуск.</li>
					</ul>
					</p>
					<br />
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