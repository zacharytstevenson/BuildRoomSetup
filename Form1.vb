Imports System
Imports System.DirectoryServices
Imports System.DirectoryServices.ActiveDirectory
Imports System.Text
Imports System.IO

Public Class Form1

    Public firstName, lastName, username, password, serviceTag, lastFirst, practice, buildType, firstLast, user As String
    Public pageready

    Public Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        user = "invalid"
        Me.Text = "Build Installer App - z3 utilities"
        Me.ForeColor = Color.WhiteSmoke
        Me.BackColor = Color.FromArgb(32, 23, 71)
        DellSvcTag()
        Label8.Text = serviceTag
        WebBrowser1.Navigate("https://mail.leidoshealth.com")
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

            firstName = TextBox1.Text
            lastName = TextBox2.Text
            username = TextBox3.Text
            password = TextBox4.Text
            lastFirst = Chr(34) & lastName & ", " & firstName & Chr(34)
            practice = ComboBox1.Text
            firstLast = firstName & " " & lastName
            webLogin()

        If user = "valid" Then
            'MsgBox("Success")
            Me.Hide()
            Form2.Show()
        End If

        'MsgBox(firstName & vbCrLf & lastName & vbCrLf & username)
    End Sub

    Sub webLogin()

        'web element info

        '<input id="username" name="username" type="text" class="txt">
        '<input id="password" name="password" type="password" class="txt" onfocus="g_fFcs=0">
        '<input type="submit" class="btn" value="Sign in" onclick="clkLgn()"  _
        'onmouseover="this.className='btnOnMseOvr'" onmouseout="this.className='btn'" onmousedown="this.className='btnOnMseDwn'">


        Dim theElementCollection As HtmlElementCollection
        theElementCollection = WebBrowser1.Document.GetElementsByTagName("input")
        For Each curElement As HtmlElement In theElementCollection
            Dim controlName As String = curElement.GetAttribute("name").ToString
            If controlName = "username" Then
                curElement.SetAttribute("Value", Username)
            ElseIf controlName = "password" Then
                curElement.SetAttribute("Value", Password)

                'In addition,you can get element value like this:
                'MessageBox.Show(curElement.GetAttribute("Value"))
            End If

        Next

        ' Part 3: Automatically clck that Login button
        theElementCollection = WebBrowser1.Document.GetElementsByTagName("input")
        For Each curElement As HtmlElement In theElementCollection
            If curElement.GetAttribute("value").Equals("Sign in") Then
                curElement.InvokeMember("click")
                ' javascript has a click method for you need to invoke on button and hyperlink elements.
            End If
        Next
    End Sub

    Sub WebBrowser1DocumentCompleted(ByVal sender As System.Object, ByVal e As System.Windows.Forms.WebBrowserDocumentCompletedEventArgs) Handles WebBrowser1.DocumentCompleted
        'When webbrower finish opening the page, source page is diplayed in text box
        If WebBrowser1.DocumentText.Contains("The user name or password you entered isn't correct. Try entering it again.") = True Then
            MsgBox("Try Again")
        End If
        If WebBrowser1.DocumentTitle.Contains(firstLast & " - Outlook Web App") = True Then
            'MsgBox("Success")
            user = "valid"
        End If
        If WebBrowser1.DocumentTitle.Contains("Outlook Web App") = True Then
            'MsgBox("Success")
            user = "valid"
        End If
    End Sub

    Public Sub DellSvcTag()
        Dim strComputer = "."
        Dim objWMIService = GetObject("winmgmts:" _
        & "{impersonationLevel=impersonate}!\\" _
        & strComputer & "\root\cimv2")
        Dim colSMBIOS = objWMIService.ExecQuery _
        ("Select * from Win32_SystemEnclosure")
        For Each objSMBIOS In colSMBIOS
            '  MsgBox("Serial Number: " & objSMBIOS.SerialNumber)
            serviceTag = objSMBIOS.SerialNumber
        Next
        '        MsgBox(svcTag)
    End Sub

End Class
