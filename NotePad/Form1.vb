Imports System.ComponentModel
Imports System.IO

Public Class Form1

    Dim file_name As String

    Private Sub NewTabToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NewTabToolStripMenuItem.Click

        Dim newTabPage As TabPage = New System.Windows.Forms.TabPage()

        For Each ctrl As Control In TabControl1.TabPages(0).Controls
            Dim newCtrl As New Control()

            'newCtrl = Mem                      '''''''             <<---------------########  LEFT OFF

            newCtrl.Location = ctrl.Location
            newCtrl.Size = ctrl.Size
            newCtrl.Anchor = ctrl.Anchor
            newTabPage.Controls.Add(newCtrl)
        Next

        newTabPage.Text = "Untitled"

        TabControl1.TabPages.Add(newTabPage)

        TabControl1.SelectedTab = newTabPage

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'TabControl1.SelectedTab("Untitled")
    End Sub

    Private Sub CloseTabToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CloseTabToolStripMenuItem.Click

        If TabControl1.TabPages.Count <= 1 Then
            Me.Close()
        Else
            TabControl1.TabPages.Remove(TabControl1.SelectedTab)
        End If

    End Sub

    Private Sub NewWindowToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NewWindowToolStripMenuItem.Click
        MyForm()
    End Sub

    Private Sub SaveToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveToolStripMenuItem.Click
        Dim file_directory = Path.GetFileName(Application.ExecutablePath)

        If Path.GetFileNameWithoutExtension(file_directory) = "Untitled" Then
            SaveToFile()
        Else
            System.IO.File.WriteAllText(file_name, TextBox1.Text)
        End If

    End Sub

    Public Sub SaveToFile()
        'Adjusts Save File Dialog layout i.e. provides a dropdown list for the selection of the File Type 
        SaveFileDialog1.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
        'Opens Save File Dialog and writes file contents to selected path
        If SaveFileDialog1.ShowDialog = DialogResult.OK Then
            System.IO.File.WriteAllText(SaveFileDialog1.FileName, TextBox1.Text)
        End If
        'Change current Tab Page Title.
        TabControl1.SelectedTab.Text = Path.GetFileNameWithoutExtension(SaveFileDialog1.FileName)
    End Sub

    Public Sub OpenFile()
        'Adjusts Open File Dialog layout i.e. provides a dropdown list for the selection of the File Type 
        OpenFileDialog1.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
        'Opens Open File Dialog and reads file contents from selected path to TextBox
        If OpenFileDialog1.ShowDialog = DialogResult.OK Then
            file_name = OpenFileDialog1.FileName
            TextBox1.Text = System.IO.File.ReadAllText(file_name)
        End If
        'Change current Tab Page Title.
        TabControl1.SelectedTab.Text = Path.GetFileNameWithoutExtension(file_name)
    End Sub

    Private Sub SaveAsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveAsToolStripMenuItem.Click
        SaveToFile()
    End Sub

    Private Sub OpenToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenToolStripMenuItem.Click
        OpenFile()
    End Sub

    Private Sub PageSetupToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PageSetupToolStripMenuItem.Click
        PageSetupDialog1.Document = PrintDocument1
        If PageSetupDialog1.ShowDialog = DialogResult.OK Then
            PrintDocument1.DefaultPageSettings = PageSetupDialog1.PageSettings
        End If
    End Sub

    Private Sub PrintToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PrintToolStripMenuItem.Click
        PrintDialog1.Document = PageSetupDialog1.Document
        If PrintDialog1.ShowDialog = DialogResult.OK Then
            PrintDocument1.Print()
        End If
    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        Application.Exit()
    End Sub

    Private Sub CloseWindowToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CloseWindowToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub MyForm()
        'Create new Tab Window
        Dim newTabWindow As New Form1
        newTabWindow.Show()
    End Sub

End Class


