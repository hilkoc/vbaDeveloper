Attribute VB_Name = "Menu"
Option Explicit

Private Const MENU_TITLE = "VbaDeveloper"

Public Sub createMenu()
    Dim rootMenu As CommandBarPopup

    'Add the top-level menu to the ribbon Add-ins section
    Set rootMenu = Application.CommandBars(1).Controls.Add(Type:=msoControlPopup, _
    Before:=10, _
    Temporary:=True)
    rootMenu.caption = MENU_TITLE


    Dim exSubMenu As CommandBarPopup
    Dim imSubMenu As CommandBarPopup
    Dim formatSubMenu As CommandBarPopup
    Set exSubMenu = addSubmenu(rootMenu, 1, "Export code for ...")
    Set imSubMenu = addSubmenu(rootMenu, 2, "Import code for ...")
    Set formatSubMenu = addSubmenu(rootMenu, 3, "Format code for ...")
    addMenuSeparator rootMenu
    Dim refreshItem As CommandBarButton
    Set refreshItem = addMenuItem(rootMenu, "Menu.refreshMenu", "Refresh this menu")
    refreshItem.FaceId = 37

    ' menuItem.FaceId = FaceId ' set a picture
    Dim vProject As Variant
    For Each vProject In Application.VBE.VBProjects
        ' We skip over unsaved projects where project.fileName throws error
        On Error GoTo nextProject
        Dim project As VBProject
        Set project = vProject
        Dim projectName As String
        projectName = project.name
        Dim caption As String
        caption = projectName & " (" & Dir(project.fileName) & ")" '<- this can throw error
        Dim exCommand As String
        Dim imCommand As String
        Dim formatCommand As String
        exCommand = "'Menu.exportVbProject """ & projectName & """'"
        imCommand = "'Menu.importVbProject """ & projectName & """'"
        formatCommand = "'Menu.formatVbProject """ & projectName & """'"
        addMenuItem exSubMenu, exCommand, caption
        addMenuItem imSubMenu, imCommand, caption
        addMenuItem formatSubMenu, formatCommand, caption
nextProject:
    Next vProject
    On Error GoTo 0 'reset the error handling
End Sub


Private Function addMenuItem(menu As CommandBarPopup, ByVal onAction As String, ByVal caption As String) As CommandBarButton
    Dim menuItem As CommandBarButton
    Set menuItem = menu.Controls.Add(Type:=msoControlButton)
    menuItem.onAction = onAction
    menuItem.caption = caption
    Set addMenuItem = menuItem
End Function


Private Function addSubmenu(menu As CommandBarPopup, ByVal position As Integer, ByVal caption As String) As CommandBarPopup
    Dim subMenu As CommandBarPopup
    Set subMenu = menu.Controls.Add(Type:=msoControlPopup)
    subMenu.onAction = position
    subMenu.caption = caption
    Set addSubmenu = subMenu
End Function


Private Sub addMenuSeparator(menuItem As CommandBarPopup)
    menuItem.BeginGroup = True
End Sub


'This sub should be executed when the workbook is closed
Public Sub deleteMenu()
    On Error Resume Next
    Application.CommandBars(1).Controls(MENU_TITLE).Delete
    On Error GoTo 0
End Sub

Public Sub refreshMenu()
    menu.deleteMenu
    menu.createMenu
End Sub

Public Sub exportVbProject(ByVal projectName As String)
    On Error GoTo exportVbProject_Error

    Dim project As VBProject
    Set project = Application.VBE.VBProjects(projectName)
    Build.exportVbaCode project
    Dim wb As Workbook
    Set wb = Build.openWorkbook(project.fileName)
    NamedRanges.exportNamedRanges wb
    MsgBox "Finished exporting code for: " & project.name

    On Error GoTo 0
    Exit Sub
exportVbProject_Error:
    ErrorHandling.handleError "Menu.exportVbProject"
End Sub


Public Sub importVbProject(ByVal projectName As String)
    On Error GoTo importVbProject_Error

    Dim project As VBProject
    Set project = Application.VBE.VBProjects(projectName)
    Build.importVbaCode project
    Dim wb As Workbook
    Set wb = Build.openWorkbook(project.fileName)
    NamedRanges.importNamedRanges wb
    MsgBox "Finished importing code for: " & project.name

    On Error GoTo 0
    Exit Sub
importVbProject_Error:
    ErrorHandling.handleError "Menu.importVbProject"
End Sub


Public Sub formatVbProject(ByVal projectName As String)
    On Error GoTo formatVbProject_Error

    Dim project As VBProject
    Set project = Application.VBE.VBProjects(projectName)
    Formatter.formatProject project
    MsgBox "Finished formatting code for: " & project.name & vbNewLine _
    & vbNewLine _
    & "Did you know you can also format your code, while writing it, by typing 'application.Run ""format""' in the immediate window?"

    On Error GoTo 0
    Exit Sub
formatVbProject_Error:
    ErrorHandling.handleError "Menu.formatVbProject"
End Sub
