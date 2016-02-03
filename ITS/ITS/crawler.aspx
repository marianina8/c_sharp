<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="crawler.aspx.cs" Inherits="ITS.crawler" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
        Web Crawler<br />
        <br />
        Please enter the URL:<br />
        <asp:TextBox ID="txtURL" runat="server"></asp:TextBox>
&nbsp;<asp:Button ID="btnStart" runat="server" OnClick="btnStart_Click" Text="Start" />
        <br />
        <br />
        <asp:GridView ID="GridView1" runat="server">
        </asp:GridView>
    </form>
</body>
</html>
