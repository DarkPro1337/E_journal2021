<% if session("user_fio") <> "�������" then %>
<%
if DateDiff("d", session("pass_date"), date()) >= 335 And DateDiff("d", session("pass_date"), date()) < 365 then
%>
<div data-role="dialog" id="dialog" class="padding20 dialog" data-show=true data-type="info" data-windows-style=true data-close-button="true" data-overlay="true" data-overlay-color="op-dark" data-overlay-click-close="true">
    <center>
    <h1>��� ������ ����� �������!</h1>
    <p style="width: 75%">
        ��� ������ �� ������� ��� <% response.Write( DateDiff("m", session("pass_date"), date()) ) %> �������. ��������� �������� ������ ��� ����� ������, � ������ �� ����� ������ � ���� �� ��������� ������ � ����� ������� ������. <br><br> �� ������ ������� � ������� ������ �������� <% response.Write( 365 - DateDiff("d", session("pass_date"), date()) ) %> ��.
    </p>
    <form action="pass_update.asp" method="post">
        <input type="hidden" name="login" value="<%=session("user_fio")%>">
        <input type="hidden" name="pass" value="<%=session("pass")%>">
        <button class="button block-shadow-info" type="submit">������� ������</button>
    </form>
    </center>
</div>
<%
elseif DateDiff("d", session("pass_date"), date()) >= 365 then
%>
<div data-role="dialog" id="Div1" class="padding20 dialog" data-show=true data-type="alert" data-windows-style=true data-overlay="true" data-overlay-color="op-dark">
    <center>
    <h1>��� ������ ����!</h1>
    <p style="width: 75%">
        ��� ������ �� ������� ��� ������ ����. ������ � ������� ������ ��� ������!<br><br>����������, ��������� ������.
    </p>
    <form action="pass_update.asp" method="post">
        <input type="hidden" name="login" value="<%=session("user_fio")%>">
        <input type="hidden" name="pass" value="<%=session("pass")%>">
        <button class="button primary block-shadow-danger" type="submit">������� ������</button>
        <a href="exit.asp"><button class="button block-shadow-info" type="button">�����</button></a>
    </form>
    </center>
</div>
<%
end if
%>
<% end if %>