<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Task.aspx.cs" Title="Aspose Bulk Emailing using MailMerge Template" ViewStateMode="Disabled" Inherits="Gridview_Export_To_CSV.Gridview_To_CSV" %>

<link rel="stylesheet" href='<%= ResolveUrl("~/Content/bootstrap-theme.min.css") %>' />
<link rel="stylesheet" href='<%= ResolveUrl("~/Content/bootstrap-theme.css") %>' />
<link rel="stylesheet" href='<%= ResolveUrl("~/Content/Site.css") %>' />
<script src="https://code.jquery.com/jquery-2.1.4.js"></script>
<script src="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.5/js/bootstrap.min.js"></script>
<script type="text/javascript" src='<%=ResolveUrl("~/Script/jquery.SimpleMask.js") %>'></script>

<style type="text/css">
    .heading {
        color: #3eb0e4;
        font-size: 24px;
        font-weight: lighter;
        margin: 0 8px;
        padding: 0;
    }

    .btn {
        display: inline-block;
        color: #FFF !important;
        text-shadow: 0 -1px 0 rgba(0,0,0,0.25) !important;
        border: 5px solid #FFF;
        border-radius: 0;
        box-shadow: none !important;
        cursor: pointer;
        vertical-align: middle;
        position: relative;
        padding: 2px 12px;
        background-color: #87b87f !important;
        border-color: #87b87f;
    }

    .btn2 {
        display: inline-block;
        color: #FFF !important;
        text-shadow: 0 -1px 0 rgba(0,0,0,0.25) !important;
        border: 5px solid #FFF;
        border-radius: 0;
        box-shadow: none !important;
        cursor: pointer;
        vertical-align: middle;
        position: relative;
        /*padding: 2px 12px;*/
        background-color: rgb(235, 147, 22) !important;
        border-color: rgb(235, 147, 22) !important;
    }

    .btn3 {
        display: inline-block;
        color: #FFF !important;
        text-shadow: 0 -1px 0 rgba(0,0,0,0.25) !important;
        border: 5px solid #FFF;
        border-radius: 0;
        box-shadow: none !important;
        cursor: pointer;
        vertical-align: middle;
        position: relative;
        padding: 2px 12px;
        background-color: #428bca !important;
        border-color: #428bca;
    }



    .form-control {
        background-color: #fff;
        border: 1px solid #d5d5d5;
        box-shadow: none !important;
        color: #666;
        padding: 5px 4px;
        border-radius: 0 !important;
        line-height: 1.42857143;
        margin-top: 10px;
        margin-bottom: 5px;
    }
</style>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<script type="text/javascript">

     $(document).ready(function () {
    $('#txtPort').simpleMask({ 'mask': '###', 'nextInput': $('#txtPort') });
    $('#txtServer').keypress(function (key) {
        if ((key.charCode < 97 || key.charCode > 122) && (key.charCode < 65 || key.charCode > 90) && (key.charCode != 64) && (key.charCode != 46)) return false;
    });


});


             </script>

<body>


    <form id="form1" runat="server">

        <div id="logo" class="col-md-12" style="text-align: center;">
            <asp:Image ImageUrl="..\\uploads\\logo\\logo.png" runat="server" />
        </div>

        <div>
            <br />

        </div>


        <fieldset style="width: 835px; border-color: #FFFFFF; margin-left: 197px;">
            <legend style="margin-bottom: 7px; text-align: center;">
                <asp:Label class="heading" runat="server" Text="Aspose Bulk Emailing using MailMerge Template"></asp:Label></legend>

            <div style="margin-top: 10px;">
                <asp:Label class="col-md-2 control-label" runat="server" Text="SMTP Settings"></asp:Label>
                <br />
                <asp:TextBox ID="txtServer" runat="server" CssClass="form-control" ToolTip="Enter Gmail SeverName" />
                <asp:TextBox ID="txtUserName" runat="server" CssClass="form-control" ToolTip="Enter UserName (Gmail)" />
                <asp:TextBox ID="txtPassword" runat="server" CssClass="form-control" TextMode="Password" ToolTip=" Enter Password" />
                <asp:TextBox ID="txtPort" runat="server" CssClass="form-control" ToolTip="Enter Port Number (accepts only three digits)" />
                <asp:Button ID="btnTest" class="btn" Text="Test" runat="server" Style="margin-left: 91px;" ToolTip="Test Email Connection" OnClick="btnTest_Click" />
            </div>

            <div style="margin-top: 10px;">
                <asp:Label class="col-md-2 control-label" runat="server" Text="Load Contacts (CSV)"></asp:Label>
                <asp:FileUpload ID="contactsFileUpload" runat="server" Style="width: 565px; margin-left: 10px;" />
                <asp:Button ID="btnContactsUpload" Text="Upload CSV" ToolTip="Upload CSV" runat="server"  style=" margin-left: 3px; " class="btn2" OnClick="btnContactUpload_Click" />
            </div>

            <div style="margin-top: 10px;">
                <asp:Label class="col-md-2" runat="server" Text="Template (MailMerge)"></asp:Label>
                <asp:FileUpload ID="emailTemplateFileUpload" runat="server" Style="width: 536px; margin-left: 6px;" />
                <asp:Button ID="emailButtonUpload" Text="Upload MailMerge" ToolTip="Upload MailMerge"  runat="server" class="btn2" OnClick="emailButtonUpload_Click" />

                <br />
                <asp:Label  style=" margin-left: 60px; " Text="List of Users" runat="server" />
                <asp:Label  style=" margin-left: 114px; " Text="List of Emails" runat="server" />
                <br />
                <asp:ListBox ID="listUsers" runat="server" Width="200px" Height="100px" SelectionMode="Multiple"></asp:ListBox>
                <asp:ListBox ID="listEmails" runat="server" Width="200px" Height="100px" SelectionMode="Multiple"></asp:ListBox>
                <asp:TextBox runat="server" ID="mailMergeHtml" TextMode="MultiLine" Style="height: 97px; width: 49%;" />


            </div>

            <div style=" float: right; ">
                <asp:Button ID="btnSend" runat="server" ToolTip="Send Email" Text="Send" class="btn3" OnClick="btnSend_Click" ></asp:Button>
                <asp:Button ID="btnClear" style="margin-right: 16px;" runat="server" ToolTip="Clear" Text="Clear" class="btn3" OnClick="btnClear_Click" ></asp:Button>
            </div>

             <div>
                 <asp:Label id="message" runat="server" />
            </div>

        </fieldset>
    </form>
</body>
</html>
