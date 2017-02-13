Attribute VB_Name = "FormSerializer"
''
' FormSerializer v1.0.0
' (c) Georges Kuenzli - https://github.com/gkuenzli/vbaDeveloper
'
' `FormSerializer` produces a string JSON description of a MSForm.
'
' @module FormSerializer
' @author gkuenzli
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit


''
' Convert a VBComponent of type MSForm to a JSON descriptive data
'
' @method serializeMSForm
' @param {VBComponent} FormComponent
' @return {String} MSForm JSON descriptive data
''
Public Function SerializeMSForm(ByVal FormComponent As VBComponent) As String
    Dim dict As Dictionary
    Dim json As String
    Set dict = GetMSFormProperties(FormComponent)
    json = ConvertToJson(dict, vbTab)
    SerializeMSForm = json
End Function

Private Function GetMSFormProperties(ByVal FormComponent As VBComponent) As Dictionary
    Dim dict As New Dictionary
    Dim p As Property
    dict.Add "Name", FormComponent.name
    dict.Add "Designer", GetDesigner(FormComponent)
    dict.Add "Properties", GetProperties(FormComponent, FormComponent.Properties)
    Set GetMSFormProperties = dict
End Function

Private Function GetDesigner(ByVal FormComponent As VBComponent) As Dictionary
    Dim dict As New Dictionary
    dict.Add "Controls", GetControls(FormComponent.Designer.Controls)
    Set GetDesigner = dict
End Function

Private Function GetProperties(ByVal Context As Object, ByVal Properties As Properties) As Dictionary
    Dim dict As New Dictionary
    Dim props As New Collection
    Dim p As Property
    Dim i As Long
    For i = 1 To Properties.Count
        Set p = Properties(i)
        If IsSerializableProperty(Context, p) Then
            'props.Add GetProperty(Context, p)
            dict.Add p.name, GetValue(Context, p)
        End If
    Next i
    Set GetProperties = dict
End Function

Private Function IsSerializableProperty(ByVal Context As Object, ByVal Property As Property) As Boolean
    Dim tp As VbVarType
    On Error Resume Next
    tp = VarType(Property.Value)
    On Error GoTo 0
    IsSerializableProperty = _
        (tp <> vbEmpty) And (tp <> vbError) And _
        Left(Property.name, 1) <> "_" And _
        InStr("ActiveControls,Controls,Handle,MouseIcon,Picture,Selected,DesignMode,ShowToolbox,ShowGridDots,SnapToGrid,GridX,GridY,DrawBuffer,CanPaste", Property.name) = 0
        
    If TypeName(Context) = "VBComponent" Then
        ' We must ignore Top and Height MSForm properties since these seem to be related to the some settings in the Windows user profile.
        IsSerializableProperty = _
            IsSerializableProperty And _
            InStr("Top,Height", Property.name) = 0
    End If
End Function

Private Function GetProperty(ByVal Context As Object, ByVal Property As Property) As Dictionary
    Dim dict As New Dictionary
    dict.Add "Name", Property.name
    If Property.name = "Controls" Then
    Else
        dict.Add "Value", GetValue(Context, Property)
    End If
    Set GetProperty = dict
End Function

Private Function GetControls(ByVal Controls As Controls) As Collection
    Dim coll As New Collection
    Dim ctrl As Control
    For Each ctrl In Controls
        If Not ControlExistsInSubElements(Controls, ctrl.name, 0) Then
            coll.Add GetControl(ctrl)
        End If
    Next ctrl
    Set GetControls = coll
End Function

Private Function ControlExistsInSubElements(ByVal Controls As Controls, ByVal name As String, ByVal Depth As Long) As Boolean
    Dim ctrl As Control
    Dim o As Object
    For Each ctrl In Controls
        Set o = ctrl
        If Depth > 0 Then
            If name = ctrl.name Then
                ControlExistsInSubElements = True
                Exit Function
            End If
        End If
        On Error Resume Next
        ControlExistsInSubElements = ControlExistsInSubElements(o.Controls, name, Depth + 1)
        On Error GoTo 0
        If ControlExistsInSubElements Then
            Exit Function
        End If
    Next ctrl
End Function

Private Function GetControl(ByVal Control As Control) As Dictionary
    Dim dict As New Dictionary
    Dim o As Object
    Set o = Control
    On Error Resume Next
    dict.Add "Class", TypeName(o)
    dict.Add "Name", Control.name
    dict.Add "Cancel", Control.Cancel
    dict.Add "ControlSource", Control.ControlSource
    dict.Add "ControlTipText", Control.ControlTipText
    dict.Add "Default", Control.Default
    dict.Add "Height", Control.Height
    dict.Add "HelpContextID", Control.HelpContextID
    dict.Add "LayoutEffect", Control.LayoutEffect
    dict.Add "Left", Control.Left
    dict.Add "RowSource", Control.RowSource
    dict.Add "RowSourceType", Control.RowSourceType
    dict.Add "TabIndex", Control.TabIndex
    dict.Add "TabStop", Control.TabStop
    dict.Add "Tag", Control.Tag
    dict.Add "Top", Control.Top
    dict.Add "Visible", Control.Visible
    dict.Add "Width", Control.Width
    
    Select Case TypeName(o)
        Case "CheckBox"
            AddCheckBox dict, o
        Case "ComboBox"
            AddComboBox dict, o
        Case "CommandButton"
            AddCommandButton dict, o
        Case "Frame"
            AddFrame dict, o
        Case "Image"
            AddImage dict, o
        Case "Label"
            AddLabel dict, o
        Case "ListBox"
            AddListBox dict, o
        Case "MultiPage"
            AddMultiPage dict, o
        Case "OptionButton"
            AddOptionButton dict, o
        Case "Page"
            AddPage dict, o
        Case "ScrollBar"
            AddScrollBar dict, o
        Case "SpinButton"
            AddSpinButton dict, o
        Case "Tab"
            AddTab dict, o
        Case "TabStrip"
            AddTabStrip dict, o
        Case "TextBox"
            AddTextBox dict, o
        Case "ToggleButton"
            AddToggleButton dict, o
        Case "RefEdit"
            AddRefEdit dict, o
        Case Else
            Debug.Print "Unknown ActiveX Control Type Name : " & TypeName(o)
    End Select
    
    Set GetControl = dict
End Function

Private Sub AddCheckBox(ByVal dict As Dictionary, ByVal o As Object)
    On Error Resume Next
    dict.Add "Accelerator", o.Accelerator
    dict.Add "Alignment", o.Alignment
    dict.Add "AutoSize", o.AutoSize
    dict.Add "BackColor", o.BackColor
    dict.Add "BackStyle", o.BackStyle
    dict.Add "Caption", o.caption
    dict.Add "Enabled", o.Enabled
    dict.Add "Font", GetFont(o.Font)
    dict.Add "ForeColor", o.ForeColor
    dict.Add "GroupName", o.GroupName
    dict.Add "Locked", o.Locked
    dict.Add "MouseIcon", GetPicture(o.MouseIcon)
    dict.Add "MousePointer", o.MousePointer
    dict.Add "Picture", GetPicture(o.Picture)
    dict.Add "PicturePosition", o.PicturePosition
    dict.Add "SpecialEffect", o.SpecialEffect
    dict.Add "TextAlign", o.TextAlign
    dict.Add "TripleState", o.TripleState
    dict.Add "Value", o.Value
    dict.Add "WordWrap", o.WordWrap
End Sub

Private Sub AddComboBox(ByVal dict As Dictionary, ByVal o As Object)
    On Error Resume Next
    dict.Add "AutoSize", o.AutoSize
    dict.Add "AutoTab", o.AutoTab
    dict.Add "AutoWordSelect", o.AutoWordSelect
    dict.Add "BackColor", o.BackColor
    dict.Add "BackStyle", o.BackStyle
    dict.Add "BorderColor", o.BorderColor
    dict.Add "BorderStyle", o.BorderStyle
    dict.Add "BoundColumn", o.BoundColumn
'    dict.Add "CanPaste", o.CanPaste
    dict.Add "ColumnCount", o.ColumnCount
    dict.Add "ColumnHeads", o.ColumnHeads
    dict.Add "ColumnWidths", o.ColumnWidths
    dict.Add "DragBehavior", o.DragBehavior
    dict.Add "DropButtonStyle", o.DropButtonStyle
    dict.Add "Enabled", o.Enabled
    dict.Add "EnterFieldBehavior", o.EnterFieldBehavior
    dict.Add "Font", GetFont(o.Font)
    dict.Add "ForeColor", o.ForeColor
    dict.Add "HideSelection", o.HideSelection
    dict.Add "IMEMode", o.IMEMode
    dict.Add "ListRows", o.ListRows
    dict.Add "ListStyle", o.ListStyle
    dict.Add "ListWidth", o.ListWidth
    dict.Add "Locked", o.Locked
    dict.Add "MatchEntry", o.MatchEntry
    dict.Add "MatchRequired", o.MatchRequired
    dict.Add "MaxLength", o.MaxLength
    dict.Add "MouseIcon", GetPicture(o.MouseIcon)
    dict.Add "MousePointer", o.MousePointer
    dict.Add "SelectionMargin", o.SelectionMargin
    dict.Add "ShowDropButtonWhen", o.ShowDropButtonWhen
    dict.Add "SpecialEffect", o.SpecialEffect
    dict.Add "Style", o.Style
    dict.Add "Text", o.text
    dict.Add "TextAlign", o.TextAlign
    dict.Add "TextColumn", o.TextColumn
    dict.Add "TopIndex", o.TopIndex
    dict.Add "Value", o.Value
End Sub

Private Sub AddCommandButton(ByVal dict As Dictionary, ByVal o As Object)
    On Error Resume Next
    dict.Add "Accelerator", o.Accelerator
    dict.Add "AutoSize", o.AutoSize
    dict.Add "BackColor", o.BackColor
    dict.Add "BackStyle", o.BackStyle
    dict.Add "Caption", o.caption
    dict.Add "Enabled", o.Enabled
    dict.Add "Font", GetFont(o.Font)
    dict.Add "ForeColor", o.ForeColor
    dict.Add "Locked", o.Locked
    dict.Add "MouseIcon", GetPicture(o.MouseIcon)
    dict.Add "MousePointer", o.MousePointer
    dict.Add "Picture", GetPicture(o.Picture)
    dict.Add "PicturePosition", o.PicturePosition
    dict.Add "TakeFocusOnClick", o.TakeFocusOnClick
    dict.Add "WordWrap", o.WordWrap
End Sub

Private Sub AddFrame(ByVal dict As Dictionary, ByVal o As Object)
    On Error Resume Next
    dict.Add "BackColor", o.BackColor
    dict.Add "BorderColor", o.BorderColor
    dict.Add "BorderStyle", o.BorderStyle
    'dict.Add "CanPaste", o.CanPaste
    dict.Add "CanRedo", o.CanRedo
    dict.Add "CanUndo", o.CanUndo
    dict.Add "Caption", o.caption
    dict.Add "Controls", GetControls(o.Controls)
    dict.Add "Cycle", o.Cycle
    dict.Add "Enabled", o.Enabled
    dict.Add "Font", GetFont(o.Font)
    dict.Add "ForeColor", o.ForeColor
    dict.Add "InsideHeight", o.InsideHeight
    dict.Add "InsideWidth", o.InsideWidth
    dict.Add "KeepScrollBarsVisible", o.KeepScrollBarsVisible
    dict.Add "MouseIcon", GetPicture(o.MouseIcon)
    dict.Add "MousePointer", o.MousePointer
    dict.Add "Picture", GetPicture(o.Picture)
    dict.Add "PictureAlignment", o.PictureAlignment
    dict.Add "PictureSizeMode", o.PictureSizeMode
    dict.Add "PictureTiling", o.PictureTiling
    dict.Add "ScrollBars", o.ScrollBars
    dict.Add "ScrollHeight", o.ScrollHeight
    dict.Add "ScrollLeft", o.ScrollLeft
    dict.Add "ScrollTop", o.ScrollTop
    dict.Add "ScrollWidth", o.ScrollWidth
    dict.Add "SpecialEffect", o.SpecialEffect
    dict.Add "VerticalScrollBarSide", o.VerticalScrollBarSide
    dict.Add "Zoom", o.Zoom
End Sub

Private Sub AddImage(ByVal dict As Dictionary, ByVal o As Object)
    On Error Resume Next
    dict.Add "AutoSize", o.AutoSize
    dict.Add "BackColor", o.BackColor
    dict.Add "BackStyle", o.BackStyle
    dict.Add "BorderColor", o.BorderColor
    dict.Add "BorderStyle", o.BorderStyle
    dict.Add "Enabled", o.Enabled
    dict.Add "MouseIcon", GetPicture(o.MouseIcon)
    dict.Add "MousePointer", o.MousePointer
    dict.Add "Picture", GetPicture(o.Picture)
    dict.Add "PictureAlignment", o.PictureAlignment
    dict.Add "PictureSizeMode", o.PictureSizeMode
    dict.Add "PictureTiling", o.PictureTiling
    dict.Add "SpecialEffect", o.SpecialEffect
End Sub

Private Sub AddLabel(ByVal dict As Dictionary, ByVal o As Object)
    On Error Resume Next
    dict.Add "Accelerator", o.Accelerator
    dict.Add "AutoSize", o.AutoSize
    dict.Add "BackColor", o.BackColor
    dict.Add "BackStyle", o.BackStyle
    dict.Add "BorderColor", o.BorderColor
    dict.Add "BorderStyle", o.BorderStyle
    dict.Add "Caption", o.caption
    dict.Add "Enabled", o.Enabled
    dict.Add "Font", GetFont(o.Font)
    dict.Add "ForeColor", o.ForeColor
    dict.Add "MouseIcon", GetPicture(o.MouseIcon)
    dict.Add "MousePointer", o.MousePointer
    dict.Add "Picture", GetPicture(o.Picture)
    dict.Add "PicturePosition", o.PicturePosition
    dict.Add "SpecialEffect", o.SpecialEffect
    dict.Add "TextAlign", o.TextAlign
    dict.Add "WordWrap", o.WordWrap
End Sub

Private Sub AddListBox(ByVal dict As Dictionary, ByVal o As Object)
    On Error Resume Next
    dict.Add "BackColor", o.BackColor
    dict.Add "BorderColor", o.BorderColor
    dict.Add "BorderStyle", o.BorderStyle
    dict.Add "BoundColumn", o.BoundColumn
    dict.Add "ColumnHeads", o.ColumnHeads
    dict.Add "ColumnWidths", o.ColumnWidths
    dict.Add "Enabled", o.Enabled
    dict.Add "Font", GetFont(o.Font)
    dict.Add "ForeColor", o.ForeColor
    dict.Add "IMEMode", o.IMEMode
    dict.Add "IntegralHeight", o.IntegralHeight
    dict.Add "ListIndex", o.ListIndex
    dict.Add "ListStyle", o.ListStyle
    dict.Add "Locked", o.Locked
    dict.Add "MatchEntry", o.MatchEntry
    dict.Add "MouseIcon", GetPicture(o.MouseIcon)
    dict.Add "MousePointer", o.MousePointer
    dict.Add "MultiSelect", o.MultiSelect
    dict.Add "Selected", o.Selected
    dict.Add "SpecialEffect", o.SpecialEffect
    dict.Add "Text", o.text
    dict.Add "TextAlign", o.TextAlign
    dict.Add "TextColumn", o.TextColumn
    dict.Add "TopIndex", o.TopIndex
    dict.Add "Value", o.Value
End Sub

Private Sub AddMultiPage(ByVal dict As Dictionary, ByVal o As Object)
    On Error Resume Next
    dict.Add "BackColor", o.BackColor
    dict.Add "Enabled", o.Enabled
    dict.Add "Font", GetFont(o.Font)
    dict.Add "ForeColor", o.ForeColor
    dict.Add "MultiRow", o.MultiRow
    dict.Add "Pages", GetPages(o.Pages)
    dict.Add "Style", o.Style
    dict.Add "TabFixedHeight", o.TabFixedHeight
    dict.Add "TabFixedWidth", o.TabFixedWidth
    dict.Add "TabOrientation", o.TabOrientation
    dict.Add "Value", o.Value
End Sub

Private Sub AddOptionButton(ByVal dict As Dictionary, ByVal o As Object)
    On Error Resume Next
    dict.Add "Accelerator", o.Accelerator
    dict.Add "Alignment", o.Alignment
    dict.Add "AutoSize", o.AutoSize
    dict.Add "BackColor", o.BackColor
    dict.Add "BackStyle", o.BackStyle
    dict.Add "Caption", o.caption
    dict.Add "Enabled", o.Enabled
    dict.Add "Font", GetFont(o.Font)
    dict.Add "ForeColor", o.ForeColor
    dict.Add "GroupName", o.GroupName
    dict.Add "Locked", o.Locked
    dict.Add "MouseIcon", GetPicture(o.MouseIcon)
    dict.Add "MousePointer", o.MousePointer
    dict.Add "Picture", GetPicture(o.Picture)
    dict.Add "PicturePosition", o.PicturePosition
    dict.Add "SpecialEffect", o.SpecialEffect
    dict.Add "TextAlign", o.TextAlign
    dict.Add "TripleState", o.TripleState
    dict.Add "Value", o.Value
    dict.Add "WordWrap", o.WordWrap
End Sub

Private Sub AddPage(ByVal dict As Dictionary, ByVal o As Object)
    On Error Resume Next
    dict.Add "Accelerator", o.Accelerator
    'dict.Add "CanPaste", o.CanPaste
    dict.Add "CanRedo", o.CanRedo
    dict.Add "CanUndo", o.CanUndo
    dict.Add "Caption", o.caption
    dict.Add "Controls", GetControls(o.Controls)
    dict.Add "ControlTipText", o.ControlTipText
    dict.Add "Cycle", o.Cycle
    dict.Add "Enabled", o.Enabled
    dict.Add "Index", o.Index
    dict.Add "InsideHeight", o.InsideHeight
    dict.Add "InsideWidth", o.InsideWidth
    dict.Add "KeepScrollBarsVisible", o.KeepScrollBarsVisible
    dict.Add "Name", o.name
    dict.Add "Parent", o.Parent
    dict.Add "Picture", GetPicture(o.Picture)
    dict.Add "PictureAlignment", o.PictureAlignment
    dict.Add "PictureSizeMode", o.PictureSizeMode
    dict.Add "PictureTiling", o.PictureTiling
    dict.Add "ScrollBars", o.ScrollBars
    dict.Add "ScrollHeight", o.ScrollHeight
    dict.Add "ScrollLeft", o.ScrollLeft
    dict.Add "ScrollTop", o.ScrollTop
    dict.Add "ScrollWidth", o.ScrollWidth
    dict.Add "Tag", o.Tag
    dict.Add "TransitionEffect", o.TransitionEffect
    dict.Add "TransitionPeriod", o.TransitionPeriod
    dict.Add "VerticalScrollBarSide", o.VerticalScrollBarSide
    dict.Add "Visible", o.Visible
    dict.Add "Zoom", o.Zoom
End Sub

Private Sub AddScrollBar(ByVal dict As Dictionary, ByVal o As Object)
    On Error Resume Next
    dict.Add "BackColor", o.BackColor
    dict.Add "Delay", o.Delay
    dict.Add "Enabled", o.Enabled
    dict.Add "ForeColor", o.ForeColor
    dict.Add "LargeChange", o.LargeChange
    dict.Add "Max", o.Max
    dict.Add "Min", o.Min
    dict.Add "MouseIcon", GetPicture(o.MouseIcon)
    dict.Add "MousePointer", o.MousePointer
    dict.Add "Orientation", o.Orientation
    dict.Add "ProportionalThumb", o.ProportionalThumb
    dict.Add "SmallChange", o.SmallChange
    dict.Add "Value", o.Value
End Sub

Private Sub AddSpinButton(ByVal dict As Dictionary, ByVal o As Object)
    On Error Resume Next
    dict.Add "BackColor", o.BackColor
    dict.Add "Delay", o.Delay
    dict.Add "Enabled", o.Enabled
    dict.Add "ForeColor", o.ForeColor
    dict.Add "Max", o.Max
    dict.Add "Min", o.Min
    dict.Add "MouseIcon", GetPicture(o.MouseIcon)
    dict.Add "MousePointer", o.MousePointer
    dict.Add "Orientation", o.Orientation
    dict.Add "SmallChange", o.SmallChange
    dict.Add "Value", o.Value
End Sub

Private Sub AddTab(ByVal dict As Dictionary, ByVal o As Object)
    On Error Resume Next
    dict.Add "Accelerator", o.Accelerator
    dict.Add "Caption", o.caption
    dict.Add "ControlTipText", o.ControlTipText
    dict.Add "Enabled", o.Enabled
    dict.Add "Index", o.Index
    dict.Add "Name", o.name
    dict.Add "Tag", o.Tag
    dict.Add "Visible", o.Visible
End Sub

Private Sub AddTabStrip(ByVal dict As Dictionary, ByVal o As Object)
    On Error Resume Next
    dict.Add "BackColor", o.BackColor
    dict.Add "ClientHeight", o.ClientHeight
    dict.Add "ClientLeft", o.ClientLeft
    dict.Add "ClientTop", o.ClientTop
    dict.Add "ClientWidth", o.ClientWidth
    dict.Add "Enabled", o.Enabled
    dict.Add "Font", GetFont(o.Font)
    dict.Add "ForeColor", o.ForeColor
    dict.Add "MouseIcon", GetPicture(o.MouseIcon)
    dict.Add "MousePointer", o.MousePointer
    dict.Add "MultiRow", o.MultiRow
    dict.Add "SelectedItem", o.SelectedItem
    dict.Add "Style", o.Style
    dict.Add "TabFixedHeight", o.TabFixedHeight
    dict.Add "TabFixedWidth", o.TabFixedWidth
    dict.Add "TabOrientation", o.TabOrientation
    dict.Add "Tabs", GetTabs(o.Tabs)
    dict.Add "Value", o.Value
End Sub

Private Sub AddTextBox(ByVal dict As Dictionary, ByVal o As Object)
    On Error Resume Next
    dict.Add "AutoSize", o.AutoSize
    dict.Add "AutoTab", o.AutoTab
    dict.Add "AutoWordSelect", o.AutoWordSelect
    dict.Add "BackColor", o.BackColor
    dict.Add "BackStyle", o.BackStyle
    dict.Add "BorderColor", o.BorderColor
    dict.Add "BorderStyle", o.BorderStyle
    'dict.Add "CanPaste", o.CanPaste
    dict.Add "CurLine", o.CurLine
    dict.Add "DragBehavior", o.DragBehavior
    dict.Add "Enabled", o.Enabled
    dict.Add "EnterFieldBehavior", o.EnterFieldBehavior
    dict.Add "EnterKeyBehavior", o.EnterKeyBehavior
    dict.Add "Font", GetFont(o.Font)
    dict.Add "ForeColor", o.ForeColor
    dict.Add "HideSelection", o.HideSelection
    dict.Add "IMEMode", o.IMEMode
    dict.Add "IntegralHeight", o.IntegralHeight
    dict.Add "Locked", o.Locked
    dict.Add "MaxLength", o.MaxLength
    dict.Add "MouseIcon", GetPicture(o.MouseIcon)
    dict.Add "MousePointer", o.MousePointer
    dict.Add "MultiLine", o.MultiLine
    dict.Add "PasswordChar", o.PasswordChar
    dict.Add "ScrollBars", o.ScrollBars
    dict.Add "SelectionMargin", o.SelectionMargin
    dict.Add "SpecialEffect", o.SpecialEffect
    dict.Add "TabKeyBehavior", o.TabKeyBehavior
    dict.Add "Text", o.text
    dict.Add "TextAlign", o.TextAlign
    dict.Add "Value", o.Value
    dict.Add "WordWrap", o.WordWrap
End Sub

Private Sub AddToggleButton(ByVal dict As Dictionary, ByVal o As Object)
    On Error Resume Next
    dict.Add "Accelerator", o.Accelerator
    dict.Add "Alignment", o.Alignment
    dict.Add "AutoSize", o.AutoSize
    dict.Add "BackColor", o.BackColor
    dict.Add "BackStyle", o.BackStyle
    dict.Add "Caption", o.caption
    dict.Add "Enabled", o.Enabled
    dict.Add "ForeColor", o.ForeColor
    dict.Add "GroupName", o.GroupName
    dict.Add "Locked", o.Locked
    dict.Add "MouseIcon", GetPicture(o.MouseIcon)
    dict.Add "MousePointer", o.MousePointer
    dict.Add "Picture", GetPicture(o.Picture)
    dict.Add "PicturePosition", o.PicturePosition
    dict.Add "SpecialEffect", o.SpecialEffect
    dict.Add "TextAlign", o.TextAlign
    dict.Add "TripleState", o.TripleState
    dict.Add "Value", o.Value
    dict.Add "WordWrap", o.WordWrap
End Sub

Private Sub AddRefEdit(ByVal dict As Dictionary, ByVal o As Object)
    AddComboBox dict, o
    On Error Resume Next
End Sub

Private Function GetPages(ByVal Pages As MSForms.Pages) As Collection
    Dim coll As New Collection
    Dim i As Long
    Dim p As MSForms.Page
    For i = 0 To Pages.Count - 1
        Set p = Pages(i)
        coll.Add GetPage(p)
    Next i
    Set GetPages = coll
End Function

Private Function GetPage(ByVal Page As MSForms.Page) As Dictionary
    Dim dict As New Dictionary
    AddPage dict, Page
    Set GetPage = dict
End Function

Private Function GetTabs(ByVal Tabs As Tabs) As Collection
    Dim coll As New Collection
    Dim i As Long
    Dim p As MSForms.Tab
    For i = 0 To Tabs.Count - 1
        Set p = Tabs(i)
        coll.Add GetTab(p)
    Next i
    Set GetTabs = coll
End Function

Private Function GetTab(ByVal t As MSForms.Tab) As Dictionary
    Dim dict As New Dictionary
    AddTab dict, t
    Set GetTab = dict
End Function

Private Function GetFont(ByVal Font As NewFont) As Dictionary
    Dim dict As New Dictionary
    dict.Add "Bold", Font.Bold
    dict.Add "Charset", Font.Charset
    dict.Add "Italic", Font.Italic
    dict.Add "Name", Font.name
    dict.Add "Size", Font.size
    dict.Add "Strikethrough", Font.Strikethrough
    dict.Add "Underline", Font.Underline
    dict.Add "Weight", Font.Weight
    Set GetFont = dict
End Function

Private Function GetPicture(ByVal Picture As IPictureDisp) As String
    
    ' TODO: implement a Base64-encoding of the picture
    
End Function

Private Function GetValue(ByVal Context As Object, ByVal Property As Property) As Variant
    If VarType(Property.Value) = vbObject Then
        Select Case TypeName(Property.Value)
            Case "Properties"
                Set GetValue = GetProperties(Context, Property.Value)
            Case Else
                Set GetValue = Nothing
        End Select
    Else
        GetValue = Property.Value
    End If
End Function
