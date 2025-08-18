Imports System.IO
Imports System.Reflection
Imports AlibreAddOn
Imports AlibreX
Imports IronPython.Hosting
Imports Microsoft.Scripting.Hosting
Imports Microsoft.VisualBasic.Logging
Imports MessageBox = System.Windows.MessageBox

Namespace AlibreAddOnAssembly
    Public Module AlibreAddOn
        Private Property AlibreRoot As IADRoot
        Private Property TheAddOnInterface As AddOnRibbon
        Private Property PythonRunner As ScriptRunner

        Public Sub AddOnLoad(hwnd As IntPtr, pAutomationHook As IAutomationHook, unused As IntPtr)
            AlibreRoot = CType(pAutomationHook.Root, IADRoot)
            PythonRunner = New ScriptRunner(AlibreRoot)
            TheAddOnInterface = New AddOnRibbon(AlibreRoot)
        End Sub

        Public Sub AddOnUnload(hwnd As IntPtr, forceUnload As Boolean, ByRef cancel As Boolean, reserved1 As Integer, reserved2 As Integer)
            TheAddOnInterface = Nothing
            PythonRunner = Nothing
            AlibreRoot = Nothing
        End Sub

        Public Sub AddOnInvoke(pAutomationHook As IntPtr, sessionName As String, isLicensed As Boolean, reserved1 As Integer, reserved2 As Integer)
        End Sub

        Public Function GetAddOnInterface() As IAlibreAddOn
            Return TheAddOnInterface
        End Function

        Public Function GetScriptRunner() As ScriptRunner
            Return PythonRunner
        End Function
    End Module

    Public Class AddOnRibbon
        Implements IAlibreAddOn

        Private ReadOnly _menuManager As MenuManager
        Private ReadOnly _alibreRoot As IADRoot

        Public Sub New(alibreRoot As IADRoot)
            _alibreRoot = alibreRoot
            _menuManager = New MenuManager()
        End Sub

        Public ReadOnly Property RootMenuItem As Integer Implements IAlibreAddOn.RootMenuItem
            Get
                Return _menuManager.GetRootMenuItem().Id
            End Get
        End Property

        Public Function HasSubMenus(menuID As Integer) As Boolean Implements IAlibreAddOn.HasSubMenus
            Dim mi As MenuItem = _menuManager.GetMenuItemById(menuID)
            If mi Is Nothing Then
                Return False
            End If
            If mi.SubItems Is Nothing Then
                Return False
            End If
            Return mi.SubItems.Count > 0
        End Function

        Public Function SubMenuItems(menuID As Integer) As Array Implements IAlibreAddOn.SubMenuItems
            Dim mi As MenuItem = _menuManager.GetMenuItemById(menuID)
            If mi Is Nothing Then
                Return Nothing
            End If
            If mi.SubItems Is Nothing Then
                Return Nothing
            End If
            Dim ids(mi.SubItems.Count - 1) As Integer
            Dim i As Integer = 0
            For Each s As MenuItem In mi.SubItems
                ids(i) = s.Id
                i += 1
            Next
            Return ids
        End Function

        Public Function MenuItemText(menuID As Integer) As String Implements IAlibreAddOn.MenuItemText
            Dim mi As MenuItem = _menuManager.GetMenuItemById(menuID)
            If mi Is Nothing Then
                Return Nothing
            End If
            Return mi.Text
        End Function

        Public Function MenuItemToolTip(menuID As Integer) As String Implements IAlibreAddOn.MenuItemToolTip
            Dim mi As MenuItem = _menuManager.GetMenuItemById(menuID)
            If mi Is Nothing Then
                Return Nothing
            End If
            Return mi.ToolTip
        End Function

        ' Icon functionality disabled: always returns null
        Public Function MenuIcon(menuID As Integer) As String Implements IAlibreAddOn.MenuIcon
            Return Nothing
        End Function

        Public Function InvokeCommand(menuID As Integer, sessionIdentifier As String) As IAlibreAddOnCommand Implements IAlibreAddOn.InvokeCommand
            Dim session As IADSession = _alibreRoot.Sessions.Item(sessionIdentifier)
            Dim mi As MenuItem = _menuManager.GetMenuItemById(menuID)
            If mi Is Nothing Then
                Return Nothing
            End If
            If mi.Command Is Nothing Then
                Return Nothing
            End If
            Return mi.Command.Invoke(session)
        End Function

        Public Function MenuItemState(menuID As Integer, sessionIdentifier As String) As ADDONMenuStates Implements IAlibreAddOn.MenuItemState
            Return ADDONMenuStates.ADDON_MENU_ENABLED
        End Function

        Public Function PopupMenu(menuID As Integer) As Boolean Implements IAlibreAddOn.PopupMenu
            Return False
        End Function

        Public Function HasPersistentDataToSave(sessionIdentifier As String) As Boolean Implements IAlibreAddOn.HasPersistentDataToSave
            Return False
        End Function

        Public Function UseDedicatedRibbonTab() As Boolean Implements IAlibreAddOn.UseDedicatedRibbonTab
            Return True
        End Function

        Private Sub setIsAddOnLicensed(isLicensed As Boolean) Implements IAlibreAddOn.setIsAddOnLicensed
        End Sub

        Public Sub LoadData(pCustomData As Global.AlibreAddOn.IStream, sessionIdentifier As String)   
        End Sub

        Public Sub SaveData(pCustomData As Global.AlibreAddOn.IStream, sessionIdentifier As String)
        End Sub

        Private Sub IAlibreAddOn_LoadData(pCustomData As Global.AlibreAddOn.IStream, sessionIdentifier As String) Implements IAlibreAddOn.LoadData
            LoadData(pCustomData, sessionIdentifier)
        End Sub

        Private Sub IAlibreAddOn_SaveData(pCustomData As Global.AlibreAddOn.IStream, sessionIdentifier As String) Implements IAlibreAddOn.SaveData
            SaveData(pCustomData, sessionIdentifier)
        End Sub
    End Class

    Public Class MenuItem
        Public Property Id As Integer
        Public Property Text As String
        Public Property ToolTip As String
        Public Property Icon As String
        Public Property Command As Func(Of IADSession, IAlibreAddOnCommand)
        Public Property SubItems As List(Of MenuItem)

        Public Sub New(id As Integer, text As String, Optional toolTip As String = "", Optional icon As String = Nothing)
            Me.Id = id
            Me.Text = text
            Me.ToolTip = toolTip
            Me.Icon = Nothing
            Me.SubItems = New List(Of MenuItem)()
        End Sub

        Public Sub AddSubItem(subItem As MenuItem)
            Me.SubItems.Add(subItem)
        End Sub

        Public Function AboutCmd(session As IADSession) As IAlibreAddOnCommand
            MessageBox.Show("Pipe Tools add-on demo" & vbCrLf & vbCrLf)
            Return Nothing
        End Function

    End Class

    Public Class MenuManager
        Private ReadOnly _rootMenuItem As MenuItem
        Private ReadOnly _menuItems As Dictionary(Of Integer, MenuItem) = New Dictionary(Of Integer, MenuItem)()

        Private Class ScriptCommandRunner
            Public FileName As String

            Public Function Run(session As IADSession) As IAlibreAddOnCommand
                Dim r As ScriptRunner = AlibreAddOnAssembly.AlibreAddOn.GetScriptRunner()
                If r IsNot Nothing Then
                    r.ExecuteScript(session, Me.FileName)
                End If
                Return Nothing
            End Function
        End Class

        Public Sub New()
            _rootMenuItem = New MenuItem(401, "alibre-pipetools-addon", "alibre-pipetools-addon","logo.ico")
            BuildMenus()
            RegisterMenuItem(_rootMenuItem)
        End Sub

        Private Sub BuildMenus()
            Dim aboutItem As New MenuItem(9090, "About", "https://github.com/stephensmitchell/alibre-pipetools-addon")
            aboutItem.Command = New Func(Of IADSession, IAlibreAddOnCommand)(AddressOf aboutItem.AboutCmd)
            _rootMenuItem.AddSubItem(aboutItem)
            Try
                Dim addOnDirectory As String = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)
                Dim examplesPath As String = Path.Combine(addOnDirectory, "scripts")
                If Directory.Exists(examplesPath) Then
                    Dim currentMenuId As Integer = 10000
                    Dim scriptFiles As String() = Directory.GetFiles(examplesPath, "*.py")
                    Dim idx As Integer
                    For idx = 0 To scriptFiles.Length - 1
                        Dim scriptFile As String = scriptFiles(idx)
                        Dim fileName As String = Path.GetFileName(scriptFile)
                        If String.Equals(fileName, "alibre_setup.py", StringComparison.OrdinalIgnoreCase) Then
                            Continue For
                        End If
                        Dim baseName As String = Path.GetFileNameWithoutExtension(fileName)
                        Dim menuText As String = baseName.Replace("-", " ").Replace("_", " ")
                        Dim scriptMenuItem As New MenuItem(currentMenuId, menuText, "Run " & fileName,"logo.ico")
                        Dim runner As New ScriptCommandRunner()
                        runner.FileName = fileName
                        scriptMenuItem.Command = New Func(Of IADSession, IAlibreAddOnCommand)(AddressOf runner.Run)
                        _rootMenuItem.AddSubItem(scriptMenuItem)
                        currentMenuId += 1
                    Next
                End If
            Catch ex As Exception
                MessageBox.Show("Failed to load scripts dynamically: " & ex.Message, "Add-on Error")
            End Try
        End Sub

        Private Sub RegisterMenuItem(menuItem As MenuItem)
            _menuItems(menuItem.Id) = menuItem
            Dim s As MenuItem
            For Each s In menuItem.SubItems
                RegisterMenuItem(s)
            Next
        End Sub

        Public Function GetMenuItemById(id As Integer) As MenuItem
            If _menuItems.ContainsKey(id) Then
                Return _menuItems(id)
            End If
            Return Nothing
        End Function

        Public Function GetRootMenuItem() As MenuItem
            Return _rootMenuItem
        End Function
    End Class

    Public Class ScriptRunner
        Private ReadOnly _engine As ScriptEngine
        Private ReadOnly _alibreRoot As IADRoot

        Public Sub New(alibreRoot As IADRoot)
            _alibreRoot = alibreRoot
            _engine = Python.CreateEngine()
            Dim alibreInstallPath As String = System.Reflection.Assembly.GetAssembly(GetType(IADRoot)).Location.Replace("\Program\AlibreX.dll", "")
            Dim searchPaths = _engine.GetSearchPaths()
            searchPaths.Add(Path.Combine(alibreInstallPath, "Program"))
            searchPaths.Add(Path.Combine(alibreInstallPath, "Program", "Addons", "AlibreScript", "PythonLib"))
            searchPaths.Add(Path.Combine(alibreInstallPath, "Program", "Addons", "AlibreScript"))
            _engine.SetSearchPaths(searchPaths)
        End Sub

        Public Sub ExecuteScript(session As IADSession, mainScriptFileName As String)
            Try
                Dim addOnDirectory As String = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)
                Dim ScriptsPath As String = Path.Combine(addOnDirectory, "scripts")
                Dim setupScriptPath As String = Path.Combine(ScriptsPath, "alibre_setup.py")
                Dim mainScriptPath As String = Path.Combine(ScriptsPath, mainScriptFileName)
                If (Not File.Exists(setupScriptPath)) OrElse (Not File.Exists(mainScriptPath)) Then
                    MessageBox.Show("Error: Script not found." & vbLf & "Setup: " & setupScriptPath & vbLf & "Main: " & mainScriptPath, "Script Error")
                    Return
                End If
                Dim scope As ScriptScope = _engine.CreateScope()
                scope.SetVariable("ScriptFileName", mainScriptFileName)
                scope.SetVariable("ScriptFolder", ScriptsPath)
                scope.SetVariable("SessionIdentifier", session.Identifier)
                scope.SetVariable("Arguments", New List(Of String)())
                scope.SetVariable("AlibreRoot", _alibreRoot)
                scope.SetVariable("CurrentSession", session)
                _engine.ExecuteFile(setupScriptPath, scope)
                _engine.ExecuteFile(mainScriptPath, scope)
            Catch ex As Exception
                MessageBox.Show("An error occurred while running the script:" & vbLf & ex.ToString(), "Python Execution Error")
            End Try
        End Sub
    End Class
End Namespace
