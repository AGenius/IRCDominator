VERSION 5.00
Begin VB.UserControl ProgYbar 
   Alignable       =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3465
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6495
   FillStyle       =   0  'Solid
   ScaleHeight     =   3465
   ScaleWidth      =   6495
   ToolboxBitmap   =   "Progbar.ctx":0000
   Begin VB.Timer timupdate 
      Enabled         =   0   'False
      Interval        =   800
      Left            =   1935
      Top             =   1980
   End
   Begin VB.TextBox txtoldpervalue 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3960
      TabIndex        =   0
      Text            =   "0"
      Top             =   3240
      Visible         =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "ProgYbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' *********************************************************************************************
' **********************************  General Declaration ****************************************
' *********************************************************************************************

Public Enum PgButton
   Pgnone
   Pgleft
   Pgright
End Enum

Public Enum EgMode
   PgHorizontalForward
   pgHorizontalBackward
   PgVerticalUpward
   PgVerticalDownward
End Enum

Public Enum EgBorder
   No_Border
   C3D_Border
End Enum

Private PgMin As Integer
Private PgMax As Double
Private PgBackcolor As OLE_COLOR
Private PgForeColor As OLE_COLOR

Private PgMarkColor As OLE_COLOR
Private PgMarkThick As Integer
Private PgMark As Boolean
Private PgBorder As EgBorder

Private PgMode As EgMode

Public Event MouseHover(Button As PgButton, X As Single, Y As Single, Value As Double)
Public Event ValueChange(Newval As Double, Oldval As Double)
Public Event ProgbarError(ErrVal As Integer, Error As String)
Public Event click(Value As Double)
' *********************************************************************************************
' *********************************************************************************************


Private Sub timupdate_Timer()
Dim X As Double

      X = Val(txtoldpervalue.Text) + 0.0000000000001
      txtoldpervalue.Text = "0"
      Updater X, PgMode, True
End Sub

' *********************************************************************************************
' ******************************* Control Initialization *******************************************
' *********************************************************************************************

Private Sub UserControl_Initialize()
      PgMin = 0
      PgMax = 100
      PgBackcolor = &H0&
      PgForeColor = &HFF&
      PgMarkColor = &HFFFF&
      PgMarkThick = 3
      PgMark = True
      PgBorder = C3D_Border
      PgMode = PgHorizontalForward
      UserControl.BackColor = PgBackcolor
End Sub
' *********************************************************************************************
' *********************************************************************************************


' *********************************************************************************************
' ***************************** Mind Bugging Imthiaz's Procedures ********************************
' *********************************************************************************************

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim b As PgButton
      If Button = 0 Then b = Pgnone
      If Button = 2 Then b = Pgright
      If Button = 1 Then
         b = Pgleft
         VIn X, Y
      End If
      Vout X, Y, b
End Sub
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim b As PgButton
      If Button = 0 Then b = Pgnone
      If Button = 2 Then b = Pgright
      If Button = 1 Then
         b = Pgleft
         VIn X, Y
      End If
      Vout X, Y, b
End Sub
Private Sub VIn(X As Single, Y As Single)
Dim per As Double
Dim x1 As Single
Dim y1 As Double
Dim Z As Double
      If PgMode = pgHorizontalBackward Then
         x1 = X
         per = (x1 / UserControl.ScaleWidth) * 100
         y1 = per * ((PgMax - PgMin) / 100)
         y1 = (PgMax - PgMin) - y1
         Updater y1, pgHorizontalBackward, False
         If PgMark = True Then
            Z = (((y1 / (Max - Min)) * 100) * UserControl.ScaleWidth) / 100
            UserControl.DrawWidth = PgMarkThick
            Line (UserControl.ScaleWidth - Z, 0)-(UserControl.ScaleWidth - Z, UserControl.ScaleHeight), PgMarkColor, BF
            UserControl.DrawWidth = 1
            txtoldpervalue.Text = 0
         End If
      End If
      If PgMode = PgHorizontalForward Then
         x1 = X
         per = (x1 / UserControl.ScaleWidth) * 100
         y1 = per * ((PgMax - PgMin) / 100)
         Updater y1, PgHorizontalForward, False
         If PgMark = True Then
            Z = (((y1 / (Max - Min)) * 100) * UserControl.ScaleWidth) / 100
            UserControl.DrawWidth = PgMarkThick
            Line (Z, 0)-(Z, UserControl.ScaleHeight), PgMarkColor, BF
            UserControl.DrawWidth = 1
         End If
      End If
      If PgMode = PgVerticalUpward Then
         x1 = Y
         per = (x1 / UserControl.ScaleHeight) * 100
         y1 = per * ((PgMax - PgMin) / 100)
         y1 = (PgMax - PgMin) - y1
         Updater y1, PgVerticalUpward, False
         If PgMark = True Then
            Z = (((y1 / (Max - Min)) * 100) * UserControl.ScaleHeight) / 100
            UserControl.DrawWidth = PgMarkThick
            Line (0, UserControl.ScaleHeight - Z)-(UserControl.ScaleWidth, UserControl.ScaleHeight - Z), PgMarkColor, BF
            UserControl.DrawWidth = 1
         End If
      End If
      If PgMode = PgVerticalDownward Then
         x1 = Y
         per = (x1 / UserControl.ScaleHeight) * 100
         y1 = per * ((PgMax - PgMin) / 100)
         Updater y1, PgVerticalDownward, False
         If PgMark = True Then
            Z = (((y1 / (Max - Min)) * 100) * UserControl.ScaleHeight) / 100
            UserControl.DrawWidth = PgMarkThick
            Line (0, Z)-(UserControl.ScaleWidth, Z), PgMarkColor, BF
            UserControl.DrawWidth = 1
            txtoldpervalue.Text = 0
         End If
      End If
      If y1 - 1 < PgMax - PgMin And y1 + 1 + PgMin > PgMin Then
         RaiseEvent click(PgMin + y1)
      Else
         RaiseEvent click(-1)
      End If
End Sub
Private Sub Vout(X As Single, Y As Single, Button As PgButton)
Dim per As Double
Dim x1 As Double
Dim y1 As Double
Dim Z As Double
      If PgMode = pgHorizontalBackward Then
         per = (X / UserControl.ScaleWidth) * 100
         y1 = per * ((PgMax - PgMin) / 100)
         y1 = (PgMax - PgMin) - y1
      End If
      If PgMode = PgHorizontalForward Then
         per = (X / UserControl.ScaleWidth) * 100
         y1 = per * ((PgMax - PgMin) / 100)
      End If
      If PgMode = PgVerticalUpward Then
         per = (X / UserControl.ScaleHeight) * 100
         y1 = per * ((PgMax - PgMin) / 100)
         y1 = (PgMax - PgMin) - y1
      End If
      If PgMode = PgVerticalDownward Then
         per = (Y / UserControl.ScaleHeight) * 100
         y1 = per * ((PgMax - PgMin) / 100)
      End If
      If y1 - 1 < PgMax - PgMin And y1 + 1 + PgMin > PgMin Then
         x1 = PgMin + y1
      Else
         x1 = -1
      End If
      RaiseEvent MouseHover(Button, X, Y, x1)
End Sub
Public Sub DrawBar(Value As Double)
      If Value > PgMax Then
         RaiseEvent ProgbarError(5, "Drawbar value cannot be more than maximum Value")
      End If
      If Value < PgMin Then
         RaiseEvent ProgbarError(6, "Drawbar value cannot be less than minumum Value")
      End If
      Updater Value, PgMode, True
End Sub
Public Sub Updater(Value As Double, Mode As EgMode, Ds As Boolean)
Dim X As Double
Dim cx As Double
Dim Max As Double
Dim Min As Double
Dim per As Double
Dim Oldval As Double
Dim Old As Double
    
      Max = PgMax
      Min = PgMin
      Old = Val(txtoldpervalue.Text)
      If Mode = PgHorizontalForward Then
    
         RaiseEvent ValueChange(Value, Old)
        
         Oldval = (((Old / (Max - Min)) * 100) * UserControl.ScaleWidth) / 100
         X = Max - Min
         per = (Value / X) * 100
         X = UserControl.ScaleWidth
         cx = (per * X) / 100
         If Ds = True Then
            If Value > Old Then
               Line (Oldval, 0)-(cx, UserControl.ScaleHeight), PgForeColor, BF
            Else
               Line (cx, 0)-(Oldval + (UserControl.ScaleWidth / 100), UserControl.ScaleHeight), UserControl.BackColor, BF
            End If
         Else
            Line (cx, 0)-(UserControl.ScaleWidth, UserControl.ScaleHeight), UserControl.BackColor, BF
            Line (0, 0)-(cx, UserControl.ScaleHeight), PgForeColor, BF
         End If
         If Value = 0 Then
            UserControl.Cls
         End If
      End If
    
      If Mode = pgHorizontalBackward Then
         RaiseEvent ValueChange(Value, Old)
         Oldval = (((Old / (Max - Min)) * 100) * UserControl.ScaleWidth) / 100
         X = Max - Min
         per = (Value / X) * 100
         X = UserControl.ScaleWidth
         cx = (per * X) / 100
         per = UserControl.ScaleHeight
         n = UserControl.ScaleWidth
         If Ds = True Then
            If Value > Val(txtoldpervalue.Text) Then
               Line (UserControl.ScaleWidth - Oldval, 0)-(UserControl.ScaleWidth - cx, UserControl.ScaleHeight), PgForeColor, BF
            Else
               Line (UserControl.ScaleWidth - cx, 0)-(UserControl.ScaleWidth - Oldval - (UserControl.ScaleWidth / 10), UserControl.ScaleHeight), UserControl.BackColor, BF
            End If
         Else
            Line (0, 0)-(UserControl.ScaleWidth - cx, UserControl.ScaleHeight), UserControl.BackColor, BF
            Line (UserControl.ScaleWidth, 0)-(UserControl.ScaleWidth - cx, UserControl.ScaleHeight), PgForeColor, BF
         End If
         If Value = 0 Then
            UserControl.Cls
         End If
      End If
    
      If Mode = PgVerticalDownward Then
         RaiseEvent ValueChange(Value, Old)
         Oldval = (((Old / (Max - Min)) * 100) * UserControl.ScaleHeight) / 100
         X = Max - Min
         per = (Value / X) * 100
         X = UserControl.ScaleHeight
         cx = (per * X) / 100
         If Ds = True Then
            If Value > Old Then
               Line (0, 0)-(UserControl.ScaleWidth, cx), PgForeColor, BF
            Else
               Line (0, Oldval + (UserControl.ScaleWidth / 10))-(UserControl.ScaleWidth, cx), UserControl.BackColor, BF
            End If
         Else
            Line (0, 0)-(UserControl.ScaleWidth, cx), PgForeColor, BF
            Line (0, UserControl.ScaleHeight)-(UserControl.ScaleWidth, cx), UserControl.BackColor, BF
         End If
         If Value = 0 Then
            UserControl.Cls
         End If
      End If
    
      If Mode = PgVerticalUpward Then
         RaiseEvent ValueChange(Value, Old)
         Oldval = (((Old / (Max - Min)) * 100) * UserControl.ScaleHeight) / 100
         X = Max - Min
         per = (Value / X) * 100
         X = UserControl.ScaleHeight
         cx = (per * X) / 100
         If Value > Old Then
            Line (0, X)-(UserControl.ScaleWidth, X - cx), PgForeColor, BF
         Else
            Line (UserControl.ScaleWidth, X - cx)-(0, X - Oldval - (UserControl.ScaleWidth / 10)), UserControl.BackColor, BF
         End If
         If Value = 0 Then
            UserControl.Cls
         End If
      End If
    
      txtoldpervalue.Text = Value - 0.0000000000001
End Sub
' *********************************************************************************************
' *********************************************************************************************


' *********************************************************************************************
' ******************************  Property Box Declaration ***************************************
' *********************************************************************************************

Public Property Get Max() As Double
Attribute Max.VB_Description = "Sets / Returns the Progbar Maximum Value"
      Max = PgMax
End Property
Public Property Let Max(Val As Double)
      If Val > PgMin Then
         PgMax = Val
         PropertyChanged "Max"
         UserControl.Cls
         txtoldpervalue.Text = 0
      Else
         RaiseEvent ProgbarError(8, "Invalid Max Value (it should be greater than Min i.e" + Str(PgMin) + ")")
         PropertyChanged "Max"
      End If
End Property
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Sets / Returns the Progbar Forecolor"
      ForeColor = PgForeColor
End Property
Public Property Let ForeColor(Val As OLE_COLOR)
      PgForeColor = Val
      PropertyChanged "ForeColor"
      txtoldpervalue.Text = 0
End Property
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Sets / Returns the Progbar BackColor"
      BackColor = PgBackcolor
End Property
Public Property Let BackColor(Val1 As OLE_COLOR)
Dim c As Double
      PgBackcolor = Val1
      PropertyChanged "BackColor"
      UserControl.BackColor = PgBackcolor
      c = Val(txtoldpervalue.Text)
      txtoldpervalue.Text = 0
      Updater c, PgMode, True
End Property
Public Property Get Mode() As EgMode
Attribute Mode.VB_Description = "Set / Returns the fill direction of the progbar"
      Mode = PgMode
End Property
Public Property Let Mode(Val As EgMode)
      If Val > -1 And Val < 4 Then
         PgMode = Val
         PropertyChanged "Mode"
         txtoldpervalue.Text = 0
         UserControl.Cls
      Else
         RaiseEvent ProgbarError(9, "Invalid Mode (it should be between 0 to 3)")
      End If
End Property
Public Property Get Border() As EgBorder
Attribute Border.VB_Description = "Sets / Return Border"
      Border = PgBorder
End Property
Public Property Let Border(Val As EgBorder)
      PgBorder = Val
      UserControl.BorderStyle = Val
      PropertyChanged "Border"
End Property

Public Property Get Mark() As Boolean
Attribute Mark.VB_Description = "Sets / Returns whether marking is on / Off"
      Mark = PgMark
End Property
Public Property Let Mark(Val As Boolean)
      PgMark = Val
      PropertyChanged "Mark"
End Property
Public Property Get MarkThickness() As Integer
Attribute MarkThickness.VB_Description = "Sets / Returns the Progbar markthickness"
      MarkThickness = PgMarkThick
End Property
Public Property Let MarkThickness(Val As Integer)
      If Val > 0 And Val < 8 Then
         PgMarkThick = Val
         PropertyChanged "MarkThicness"
      Else
         RaiseEvent ProgbarError(10, "Invalid Markthickness (it should be between 1 to 7)")
      End If
End Property
Public Property Get MarkColor() As OLE_COLOR
Attribute MarkColor.VB_Description = "Sets / Returns the Progbar markcolor"
      MarkColor = PgMarkColor
End Property
Public Property Let MarkColor(Val As OLE_COLOR)
      PgMarkColor = Val
      PropertyChanged "MarkColor"
End Property
Public Sub AboutBox()
      frmtest2.Show vbModal
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
      ' On Error Resume Next
      ForeColor = PropBag.ReadProperty("ForeColor", &HFF&)
      BackColor = PropBag.ReadProperty("BackColor", &H0&)
      Max = PropBag.ReadProperty("Max", 100)
      Mode = PropBag.ReadProperty("Mode", 1)
      Border = PropBag.ReadProperty("Border", 1)
      Mark = PropBag.ReadProperty("Mark", True)
      MarkThickness = PropBag.ReadProperty("MarkThicness", 3)
      MarkColor = PropBag.ReadProperty("MarkColor", &HFFFF&)
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
      ' On Error Resume Next
      PropBag.WriteProperty "ForeColor", ForeColor
      PropBag.WriteProperty "BackColor", BackColor
      PropBag.WriteProperty "Max", Max
      PropBag.WriteProperty "Mode", Mode
      PropBag.WriteProperty "Border", Border
      PropBag.WriteProperty "Mark", Mark
      PropBag.WriteProperty "MarkThicness", PgMarkThick
      PropBag.WriteProperty "MarkColor", MarkColor
End Sub
' *********************************************************************************************
' *********************************************************************************************
