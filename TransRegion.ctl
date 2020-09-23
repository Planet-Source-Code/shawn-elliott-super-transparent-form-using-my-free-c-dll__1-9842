VERSION 5.00
Begin VB.UserControl TransRegion 
   AutoRedraw      =   -1  'True
   ClientHeight    =   450
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   525
   ControlContainer=   -1  'True
   ScaleHeight     =   30
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   35
   ToolboxBitmap   =   "TransRegion.ctx":0000
End
Attribute VB_Name = "TransRegion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Default Property Values:
Const m_def_MaskRed = 0
Const m_def_MaskGreen = 0
Const m_def_MaskBlue = 0
'Const m_def_MaskColor = 0
Const m_def_UseFormImage = 0
Const m_def_ParentHWND = 0

Private DiffX As Integer
Private DiffY As Integer

'Property Variables:
Dim m_MaskRed As Long
Dim m_MaskGreen As Long
Dim m_MaskBlue As Long
'Dim m_MaskColor As OLE_COLOR
Dim m_UseFormImage As Boolean
Dim m_ParentHWND As Long
'Event Declarations:
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 1 Then

        'Let's set the window position now
        
        Dim Flags As Long    'This will hold the Flags for the SetWindowPos call
        Dim NewX As Integer, NewY As Integer 'This will hold the New X & Y coords for the window's Upper left hand corner
        
        'Move the form according to where it was and the difference
        'between the old and new points
        Dim MsPos As POINTAPI
        
        Call GetCursorPos(MsPos)    'Get the Current Mouse Pos

        UserControl.ScaleMode = vbPixels
        
        NewX = MsPos.X - DiffX
        NewY = MsPos.Y - DiffY
               
        Flags = SWP_NOSIZE Or SWP_NOZORDER Or SWP_FRAMECHANGED
            'The SWP_NOACTIVATE will let the window do it's own zorder
            'The SWP_NOSIZE tells the window NOT to resize
            'and the SWP_NOZORDER tells the window to use it's current zorder
        
        RetVal = SetWindowPos(GetParent(UserControl.hwnd), HWND_TOP, NewX, NewY, 0, 0, Flags)
End If

End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    'Get the difference between the left and the mouse position
    Dim Rt As RECT, MsPos As POINTAPI
    
    Call GetCursorPos(MsPos)    'Get the Current Mouse Pos
    Call GetWindowRect(GetParent(UserControl.hwnd), Rt)
    DiffX = MsPos.X - Rt.Left
    DiffY = MsPos.Y - Rt.Top
End If

    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hDC
Public Property Get hdc() As Long
Attribute hdc.VB_Description = "Returns a handle (from Microsoft Windows) to the object's device context."
    hdc = UserControl.hdc
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hwnd() As Long
Attribute hwnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
    hwnd = UserControl.hwnd
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Picture
Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set Picture = UserControl.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set UserControl.Picture = New_Picture
    PropertyChanged "Picture"
       
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7
Public Function TransRegion(Optional Red As Long = 0, Optional Green As Long = 0, Optional Blue As Long = 0) As Integer
Attribute TransRegion.VB_Description = "This function creates a Transparent Region on the ParentHWND using the HDC of the picture property"
ParentHWND = GetParent(UserControl.hwnd)

'This will determine if we use the passed parameters or the colors specified by Control property

If Red = 0 And Green = 0 And Blue = 0 Then
    'Of the passed params show pure black then we use the specified color
    'This way if they do want black they can specify the MaskRed,Blue and Green to
    'be pure black
    Red = MaskRed
    Blue = MaskBlue
    Green = MaskGreen
End If

If ParentHWND = 0 Then
    'No handle to work with
    Call Err.Raise(1, "MakeTransparent", "Please specify the ParentHWND('The handle of the Form you want to make transparent') before calling TransRegion")
    TransRegion = 1
End If

    Dim CompDC As Long, hBmp As Long
    Dim SourceHDC As Long, SourceBMP As Long, Ret As Integer

    SourceHDC = UserControl.hdc
    
    SourceBMP = UserControl.Picture

    'Create a DC for this image
    CompDC = CreateCompatibleDC(SourceHDC)

    'Set the image
    hBmp = SelectObject(CompDC, SourceBMP)
        
    'Run the Transparent function in the TransRegion.dll

    Call MakeTransparent(ParentHWND, SourceHDC, Red, Blue, Green, UserControl.ScaleWidth, UserControl.ScaleHeight, 0)

End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get ParentHWND() As Long
Attribute ParentHWND.VB_Description = "Parent Form to create Transparent Regions on"
    ParentHWND = m_ParentHWND
End Property

Public Property Let ParentHWND(ByVal New_ParentHWND As Long)
    m_ParentHWND = New_ParentHWND
    PropertyChanged "ParentHWND"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_ParentHWND = m_def_ParentHWND

'    m_MaskColor = m_def_MaskColor
    m_UseFormImage = m_def_UseFormImage
    m_MaskRed = m_def_MaskRed
    m_MaskGreen = m_def_MaskGreen
    m_MaskBlue = m_def_MaskBlue
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    m_ParentHWND = PropBag.ReadProperty("ParentHWND", m_def_ParentHWND)
'    m_MaskColor = PropBag.ReadProperty("MaskColor", m_def_MaskColor)
    m_UseFormImage = PropBag.ReadProperty("UseFormImage", m_def_UseFormImage)
    m_MaskRed = PropBag.ReadProperty("MaskRed", m_def_MaskRed)
    m_MaskGreen = PropBag.ReadProperty("MaskGreen", m_def_MaskGreen)
    m_MaskBlue = PropBag.ReadProperty("MaskBlue", m_def_MaskBlue)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Picture", Picture, Nothing)
    Call PropBag.WriteProperty("ParentHWND", m_ParentHWND, m_def_ParentHWND)
'    Call PropBag.WriteProperty("MaskColor", m_MaskColor, m_def_MaskColor)
    Call PropBag.WriteProperty("UseFormImage", m_UseFormImage, m_def_UseFormImage)
    Call PropBag.WriteProperty("MaskRed", m_MaskRed, m_def_MaskRed)
    Call PropBag.WriteProperty("MaskGreen", m_MaskGreen, m_def_MaskGreen)
    Call PropBag.WriteProperty("MaskBlue", m_MaskBlue, m_def_MaskBlue)

End Sub
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=10,0,0,0
'Public Property Get MaskColor() As OLE_COLOR
'    MaskColor = m_MaskColor
'End Property
'
'Public Property Let MaskColor(ByVal New_MaskColor As OLE_COLOR)
'    m_MaskColor = New_MaskColor
'    PropertyChanged "MaskColor"

'End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get UseFormImage() As Boolean
Attribute UseFormImage.VB_Description = "Whether to use the Image specified in the TransRegion control (False) or the Image in the Form (True)"
    UseFormImage = m_UseFormImage
End Property

Public Property Let UseFormImage(ByVal New_UseFormImage As Boolean)
    m_UseFormImage = New_UseFormImage
    PropertyChanged "UseFormImage"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get MaskRed() As Long
Attribute MaskRed.VB_Description = "The Red part of the RGB value that will be Transparent.  eg 255,0,0 will make Pure red TransParent"
    MaskRed = m_MaskRed
End Property

Public Property Let MaskRed(ByVal New_MaskRed As Long)
    m_MaskRed = New_MaskRed
    PropertyChanged "MaskRed"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get MaskGreen() As Long
Attribute MaskGreen.VB_Description = "The Green part of the RGB value that will be Transparent.  eg 0,255,0 will make Pure green TransParent"
    MaskGreen = m_MaskGreen
End Property

Public Property Let MaskGreen(ByVal New_MaskGreen As Long)
    m_MaskGreen = New_MaskGreen
    PropertyChanged "MaskGreen"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get MaskBlue() As Long
Attribute MaskBlue.VB_Description = "The Blue part of the RGB value that will be Transparent.  eg 0,0,255 will make Pure blue TransParent"
    MaskBlue = m_MaskBlue
End Property

Public Property Let MaskBlue(ByVal New_MaskBlue As Long)
    m_MaskBlue = New_MaskBlue
    PropertyChanged "MaskBlue"
End Property

