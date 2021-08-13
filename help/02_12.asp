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
		<title>ИС "Электронный журнал". Файл справки. Ведомость пропуска занятий.</title>
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
					<h2>Ведомость пропуска занятий</h2>
            
					<p class="helptext">Ведомость пропуска занятий предосмотрена для формирования сводной таблицы с количеством часов пропусков по каждому студенту группы. Прежде чем сформировать ведомость нужно выбрать группу в области (1), дату ведомости в области (2), включить или выключить отображение пустых колонок в области (3).
					
					<center><img src="../images/ved_propusk_zan1.png" /><br /></center>
					
					<p class="helptext">После нажатия кнопки <i>Подтвердить</i> ведомость в виде таблицы должна будет загрузиться. Предполагается печать данной таблицы, для этого нажмите на кнопку в области (1).</p> <br />
					
					<center><img src="../images/ved_propusk_zan2.png" /><br /></center>
					
					<p class="helptext">После нажатия на кнопку <i>Печать</i> вам будет отображено окно предпросмотра документа для печати. Внешний вид данного окна зависит от используемого вами браузера, в данном случае рассматривается окно печати в браузере Google Chrome, хоть и внешний вид может отличаться набор функций будет примерно одинаков. В этом окне вам доступен визуальный предпросмотр документа в области (1), в области (2) вы можете выбрать принтер на котором будете печатать, в области (3) вы можете выбрать количество копий ведомости, в области (4) вы можете выбрать раскладку документа, альбомную или книжную, в области (5) вам доступен раздел дополнительных настроек, по нажатии на кнопку <i>Печать</i> в области (6) в очередь принтера будет отправлена ведомость.</p><br />
					
					<center><img src="../images/ved_propusk_zan3.png" /><br /></center>
					
					<p class="helptext">При раскрытии раздела <i>Дополнительные настройки</i> вам будет отображён ряд доступных настроек. В области (1) вы можете изменить размер бумаги, если требуется, в области (2) вы можете выбрать отступы на странице, либо их отсутсвие, так же есть возможность индивидуально настроить их с каждой стороны документа, так же в области (3) вы можете включить двустороннюю печать если ваш принтер её поддерживает, для области (4) рекомендуется отключать все параметры, такие как фон и колонтитулы, в области (5) вам доступен переход в стандартное для вашей операционной системы системное диалоговое окно печати, переходите в него в случаях если у вас возникают проблемы с печатью через встроенное в браузер средство.</p><br /> 
					
					<center><img src="../images/ved_propusk_zan4.png" /><br /></center>
					
					<p class="helptext">Так же вы можете сохранить данную ведомость и в качестве PDF файла, рядом с пунктом <i>Принетр</i> выбрав из выпадающего списка пункт <i>Сохранить как PDF</i>.</p>
					
					<center><img src="../images/ved_propusk_zan5.png" /><br /></center>
					
					<p class="helptext">Если в этом списке отстуствует желаемый принтер вы можете его добавить выбрав пункт <i>Ещё...</i>. В открывшемся окне будет отображён полный перечень когда либо подключенных принтеров, в области (1) доступен поиск по названию принтера в системе, в области (2) доступен текущий результат поиска принтеров, кликните по желаемому чтобы добавить его в перечень принтеров, в области (3) располагается кнопки открытия системных настроек подключенных принтеров.</p><br /> 
					
					<center><img src="../images/ved_propusk_zan6.png" /><br /></center>
						
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