<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Index.aspx.cs" Inherits="Word2Html.Index" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
    
        <asp:FileUpload ID="FileUpload1" runat="server" />
        <br />
        <br />
        <asp:Button ID="Button1" runat="server" OnClick="Button1_Click" Text="Upload" Width="83px" />
        <br />
        <br />
        <asp:HyperLink ID="HyperLink1" runat="server">HyperLink</asp:HyperLink>
        <br />
    
    </div>
    </form>
</body>
</html>
