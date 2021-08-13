<% if session("user_fio") <> "Студент" then %>
<%
if DateDiff("d", session("pass_date"), date()) >= 335 And DateDiff("d", session("pass_date"), date()) < 365 then
%>
<div data-role="dialog" id="dialog" class="padding20 dialog" data-show=true data-type="info" data-windows-style=true data-close-button="true" data-overlay="true" data-overlay-color="op-dark" data-overlay-click-close="true">
    <center>
    <h1>Ваш пароль скоро истечёт!</h1>
    <p style="width: 75%">
        Ваш пароль не менялся уже <% response.Write( DateDiff("m", session("pass_date"), date()) ) %> месяцев. Требуется поменять пароль как можно скорее, в случае не смены пароля в срок вы потеряете доступ к своей учётной записи. <br><br> До потери доступа к учётной записи осталось <% response.Write( 365 - DateDiff("d", session("pass_date"), date()) ) %> дн.
    </p>
    <form action="pass_update.asp" method="post">
        <input type="hidden" name="login" value="<%=session("user_fio")%>">
        <input type="hidden" name="pass" value="<%=session("pass")%>">
        <button class="button block-shadow-info" type="submit">Сменить пароль</button>
    </form>
    </center>
</div>
<%
elseif DateDiff("d", session("pass_date"), date()) >= 365 then
%>
<div data-role="dialog" id="Div1" class="padding20 dialog" data-show=true data-type="alert" data-windows-style=true data-overlay="true" data-overlay-color="op-dark">
    <center>
    <h1>Ваш пароль истёк!</h1>
    <p style="width: 75%">
        Ваш пароль не менялся уже больше года. Доступ к учётной записи был утерян!<br><br>Пожалуйста, поменяйте пароль.
    </p>
    <form action="pass_update.asp" method="post">
        <input type="hidden" name="login" value="<%=session("user_fio")%>">
        <input type="hidden" name="pass" value="<%=session("pass")%>">
        <button class="button primary block-shadow-danger" type="submit">Сменить пароль</button>
        <a href="exit.asp"><button class="button block-shadow-info" type="button">Выйти</button></a>
    </form>
    </center>
</div>
<%
end if
%>
<% end if %>