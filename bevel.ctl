VERSION 5.00
Begin VB.UserControl Bevel 
   AutoRedraw      =   -1  'True
   ClientHeight    =   705
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3750
   ControlContainer=   -1  'True
   ScaleHeight     =   705
   ScaleWidth      =   3750
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   3
      X1              =   240
      X2              =   6240
      Y1              =   780
      Y2              =   780
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   2
      X1              =   1470
      X2              =   1470
      Y1              =   540
      Y2              =   3720
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   1
      X1              =   930
      X2              =   930
      Y1              =   1140
      Y2              =   4320
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   0
      X1              =   0
      X2              =   6000
      Y1              =   420
      Y2              =   420
   End
End
Attribute VB_Name = "Bevel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Enum TStyle
      VbLowered = 0
      VbRaised = 2
End Enum
Const m_def_BevelStyle = 0
Dim m_BevelStyle As TStyle

Private Sub StyleIT(Style As TStyle)
      Select Case Style
         Case VbLowered
            Line1(0).BorderColor = &H808080
            Line1(1).BorderColor = &H808080
            Line1(2).BorderColor = &HFFFFFF
            Line1(3).BorderColor = &HFFFFFF

         Case VbRaised
            Line1(0).BorderColor = &HFFFFFF
            Line1(1).BorderColor = &HFFFFFF
            Line1(2).BorderColor = &H808080
            Line1(3).BorderColor = &H808080
      End Select

End Sub

Private Sub UserControl_Resize()
      ' Left
      Line1(0).X1 = 0
      Line1(0).Y1 = UserControl.ScaleHeight
      Line1(0).X2 = 0
      Line1(0).Y2 = 0
      ' Top
      Line1(1).X1 = 0
      Line1(1).Y1 = 0
      Line1(1).X2 = UserControl.ScaleWidth
      Line1(1).Y2 = 0
      ' Right
      Line1(2).X1 = UserControl.ScaleWidth - 20
      Line1(2).Y1 = 0
      Line1(2).X2 = UserControl.ScaleWidth - 20
      Line1(2).Y2 = UserControl.ScaleHeight
      ' Bottom
      Line1(3).X1 = 0
      Line1(3).Y1 = UserControl.ScaleHeight - 20
      Line1(3).X2 = UserControl.ScaleWidth
      Line1(3).Y2 = UserControl.ScaleHeight - 20
End Sub
' MemberInfo=14,0,0,0
Public Property Get BevelStyle() As TStyle
      BevelStyle = m_BevelStyle
End Property

Public Property Let BevelStyle(ByVal New_BevelStyle As TStyle)
      m_BevelStyle = New_BevelStyle
      PropertyChanged "BevelStyle"
       StyleIT m_BevelStyle
End Property

' Initialize Properties for User Control
Private Sub UserControl_InitProperties()
      m_BevelStyle = m_def_BevelStyle
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
      m_BevelStyle = PropBag.ReadProperty("BevelStyle", m_def_BevelStyle)
      UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
      StyleIT m_BevelStyle
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
      Call PropBag.WriteProperty("BevelStyle", m_BevelStyle, m_def_BevelStyle)
      Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
End Sub
' MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
      BackColor = UserControl.BackColor
End Property
Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
      UserControl.BackColor() = New_BackColor
      PropertyChanged "BackColor"
End Property
