<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="DocumentUrl.aspx.cs" Inherits="DocumentContainer.DocumentUrl" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <script src="Scirpts/Logic.js"></script>
</head>
<body>
    <form id="form1" runat="server">
        <div>
        <div>
            <h5>SearchDocument</h5>
            <span>
                <asp:TextBox runat="server" ID="txtDocumentName" /><asp:Button ID="btnGetDocument" Text="Get Document" runat="server" OnClick="btnGetDocument_Click" />
            </span>
        </div>
            <h5>Results</h5>
        <div runat="server" id="searchedDocuments"></div>
            </div>
        <div>
            <h5>All Documents</h5>
            <div runat="server" id="documents"></div>
        </div>
    </form>
</body>
</html>
