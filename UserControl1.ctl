VERSION 5.00
Begin VB.UserControl SmoothBaar 
   ClientHeight    =   2715
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5820
   ScaleHeight     =   181
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   388
   ToolboxBitmap   =   "UserControl1.ctx":0000
   Begin VB.PictureBox buffer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   600
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   145
      TabIndex        =   3
      Top             =   1080
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.PictureBox display 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   480
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   89
      TabIndex        =   2
      Top             =   240
      Width           =   1335
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0 %"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   630
         TabIndex        =   4
         Top             =   0
         Width           =   315
      End
   End
   Begin VB.PictureBox PicEndColor 
      Height          =   255
      Left            =   1800
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   1
      Top             =   2160
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox PicStartColor 
      Height          =   255
      Left            =   1440
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   0
      Top             =   2160
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      X1              =   24
      X2              =   160
      Y1              =   8
      Y2              =   8
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      X1              =   16
      X2              =   16
      Y1              =   8
      Y2              =   32
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00404040&
      X1              =   32
      X2              =   168
      Y1              =   56
      Y2              =   56
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      X1              =   168
      X2              =   168
      Y1              =   32
      Y2              =   56
   End
End
Attribute VB_Name = "SmoothBaar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Enum AppearanceConst
    Raised = 0
    Sunken = 1
    Flat = 2
    Simple = 3
End Enum

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private mGradient As New clsGradient
Private MyStartColor As OLE_COLOR
Private DefStartColor As OLE_COLOR
Private MyEndColor As OLE_COLOR
Private DefEndColor As OLE_COLOR
Private MyFontColor As OLE_COLOR
Private DefFontColor As OLE_COLOR
Private MyFont As Font
Private DefBorder As Boolean
Private DefShowlabel As Boolean
Private myborder As Boolean
Private MyBorderColor As OLE_COLOR
Private DefBorderColor As OLE_COLOR
Private MyCaption As Boolean
Private myraised As Boolean
Private DefMax As Long
Private MyMax As Long
Private MyAppearance As AppearanceConst
Private Const MyDefAppearance = Flat
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)





Private Sub display_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub display_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub display_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Initialize()
    DefBorderColor = &H0
    MyBorderColor = DefBorderColor
    DefBorder = True
    Appearance = MyDefAppearance
    DefMax = 100
    Max = DefMax
    DefStartColor = &HFFFFFF
    ProgressStartColor = DefStartColor
    PicStartColor.BackColor = DefStartColor
    DefFontColor = &HFFFFFF
    FontColor = DefFontColor
    Label1.ForeColor = DefFontColor
    DefEndColor = &H0
    ProgressEndColor = DefEndColor
    PicEndColor.BackColor = DefEndColor
    DefShowlabel = True
    Caption = DefShowlabel
    value (Max)
    PropertyChanged ("ProgressStartColor")
    PropertyChanged ("ProgressEndColor")
    PropertyChanged ("FontColor")
    UserControl_Resize
    'PropertyChanged ("Font")
End Sub

Private Sub UserControl_InitProperties()
Set Font = Ambient.Font
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    'Define defaults
   
    'Read values
    Appearance = PropBag.ReadProperty("Appearance", MyDefAppearance)
    Max = PropBag.ReadProperty("Max", DefMax)
    FontColor = PropBag.ReadProperty("Fontcolor", DefFontColor)
    Set Font = PropBag.ReadProperty("Font", Ambient.Font)
    ProgressStartColor = PropBag.ReadProperty("ProgressStartColor", DefStartColor)
    ProgressEndColor = PropBag.ReadProperty("ProgressEndColor", DefEndColor)
    BorderColor = PropBag.ReadProperty("BorderColor", DefBorderColor)
    Caption = PropBag.ReadProperty("Caption", DefShowlabel)
    'Raised = PropBag.ReadProperty("Raised", DefBorder)
    UserControl_Resize
    PropertyChanged ("BerderWidth")
    PropertyChanged "Max"
    PropertyChanged "ProgressStartColor"
    PropertyChanged "ProgressEndColor"
    PropertyChanged "Caption"
    PropertyChanged "Appearance"
    PropertyChanged "Font"
End Sub

Private Sub UserControl_Resize()


If Appearance = Flat Then
    If Caption = True Then
        Label1.Visible = True
        Label1.Move 0, UserControl.ScaleHeight / 2 - (Label1.Height / 2), UserControl.ScaleWidth, Label1.Height
    Else
        Label1.Visible = False
    End If
    Line1.Visible = False
    Line2.Visible = False
    Line3.Visible = False
    Line4.Visible = False
    display.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
    buffer.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
End If
If Appearance = Raised Then

    If Caption = True Then
        Label1.Visible = True
        Label1.Move 0, UserControl.ScaleHeight / 2 - (Label1.Height / 2), UserControl.ScaleWidth, Label1.Height
    Else
        Label1.Visible = False
    End If
    Line1.Visible = True
    Line2.Visible = True
    Line3.Visible = True
    Line4.Visible = True
    Line1.BorderColor = &H404040
    Line2.BorderColor = &H404040
    Line3.BorderColor = &HFFFFFF
    Line4.BorderColor = &HFFFFFF
    
    Line3.X1 = 0
    Line3.X2 = 0
    Line3.Y1 = 0
    Line3.Y2 = UserControl.ScaleHeight - Line1.BorderWidth
    
    Line4.X1 = 0
    Line4.X2 = UserControl.ScaleWidth - Line1.BorderWidth
    Line4.Y1 = 0
    Line4.Y2 = 0
    
    Line2.X1 = 0
    Line2.X2 = UserControl.ScaleWidth - Line1.BorderWidth
    Line2.Y1 = UserControl.ScaleHeight - Line1.BorderWidth
    Line2.Y2 = UserControl.ScaleHeight - Line1.BorderWidth
    
    Line1.X1 = UserControl.ScaleWidth - Line1.BorderWidth
    Line1.X2 = UserControl.ScaleWidth - Line1.BorderWidth
    Line1.Y1 = 0
    Line1.Y2 = UserControl.ScaleHeight - Line1.BorderWidth
    display.Move Line1.BorderWidth, Line1.BorderWidth, UserControl.ScaleWidth - (Line1.BorderWidth * 2), UserControl.ScaleHeight - (Line1.BorderWidth * 2)
    buffer.Move Line1.BorderWidth, Line1.BorderWidth, UserControl.ScaleWidth - (Line1.BorderWidth * 2), UserControl.ScaleHeight - (Line1.BorderWidth * 2)
End If

If Appearance = Simple Then
    
    If Caption = True Then
        Label1.Visible = True
        Label1.Move 0, UserControl.ScaleHeight / 2 - (Label1.Height / 2), UserControl.ScaleWidth, Label1.Height
    Else
        Label1.Visible = False
    End If
    Line1.Visible = True
    Line2.Visible = True
    Line3.Visible = True
    Line4.Visible = True
    Line1.BorderColor = BorderColor
    Line2.BorderColor = BorderColor
    Line3.BorderColor = BorderColor
    Line4.BorderColor = BorderColor
    
    Line3.X1 = 0
    Line3.X2 = 0
    Line3.Y1 = 0
    Line3.Y2 = UserControl.ScaleHeight - Line1.BorderWidth
    
    Line4.X1 = 0
    Line4.X2 = UserControl.ScaleWidth - Line1.BorderWidth
    Line4.Y1 = 0
    Line4.Y2 = 0
    
    Line2.X1 = 0
    Line2.X2 = UserControl.ScaleWidth - Line1.BorderWidth
    Line2.Y1 = UserControl.ScaleHeight - Line1.BorderWidth
    Line2.Y2 = UserControl.ScaleHeight - Line1.BorderWidth
    
    Line1.X1 = UserControl.ScaleWidth - Line1.BorderWidth
    Line1.X2 = UserControl.ScaleWidth - Line1.BorderWidth
    Line1.Y1 = 0
    Line1.Y2 = UserControl.ScaleHeight - Line1.BorderWidth
    display.Move Line1.BorderWidth, Line1.BorderWidth, UserControl.ScaleWidth - (Line1.BorderWidth * 2), UserControl.ScaleHeight - (Line1.BorderWidth * 2)
    buffer.Move Line1.BorderWidth, Line1.BorderWidth, UserControl.ScaleWidth - (Line1.BorderWidth * 2), UserControl.ScaleHeight - (Line1.BorderWidth * 2)
End If

If Appearance = Sunken Then

    If Caption = True Then
        Label1.Visible = True
        Label1.Move 0, UserControl.ScaleHeight / 2 - (Label1.Height / 2), UserControl.ScaleWidth, Label1.Height
    Else
        Label1.Visible = False
    End If
    Line1.Visible = True
    Line2.Visible = True
    Line3.Visible = True
    Line4.Visible = True
    Line1.BorderColor = &H404040
    Line2.BorderColor = &H404040
    Line3.BorderColor = &HFFFFFF
    Line4.BorderColor = &HFFFFFF
    
    Line1.X1 = 0
    Line1.X2 = 0
    Line1.Y1 = 0
    Line1.Y2 = UserControl.ScaleHeight - Line1.BorderWidth
    
    Line2.X1 = 0
    Line2.X2 = UserControl.ScaleWidth - Line1.BorderWidth
    Line2.Y1 = 0
    Line2.Y2 = 0
    
    Line4.X1 = 0
    Line4.X2 = UserControl.ScaleWidth - Line1.BorderWidth
    Line4.Y1 = UserControl.ScaleHeight - Line1.BorderWidth
    Line4.Y2 = UserControl.ScaleHeight - Line1.BorderWidth
    
    Line3.X1 = UserControl.ScaleWidth - Line1.BorderWidth
    Line3.X2 = UserControl.ScaleWidth - Line1.BorderWidth
    Line3.Y1 = 0
    Line3.Y2 = UserControl.ScaleHeight - Line1.BorderWidth
    display.Move Line1.BorderWidth, Line1.BorderWidth, UserControl.ScaleWidth - (Line1.BorderWidth * 2), UserControl.ScaleHeight - (Line1.BorderWidth * 2)
    buffer.Move Line1.BorderWidth, Line1.BorderWidth, UserControl.ScaleWidth - (Line1.BorderWidth * 2), UserControl.ScaleHeight - (Line1.BorderWidth * 2)
End If


value (Max)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Call PropBag.WriteProperty("Max", MyMax, DefMax)
Call PropBag.WriteProperty("Raised", Raised, DefBorder)
Call PropBag.WriteProperty("ProgressStartColor", MyStartColor, DefStartColor)
Call PropBag.WriteProperty("ProgressEndColor", MyEndColor, DefEndColor)
Call PropBag.WriteProperty("BorderColor", MyBorderColor, DefBorderColor)
Call PropBag.WriteProperty("Fontcolor", MyFontColor, DefFontColor)
Call PropBag.WriteProperty("Font", MyFont, Ambient.Font)
Call PropBag.WriteProperty("Caption", MyCaption, DefShowlabel)
Call PropBag.WriteProperty("Appearance", MyAppearance, MyDefAppearance)
End Sub

Public Property Get Caption() As Boolean
    Caption = MyCaption
End Property
Public Property Let Caption(ByVal Bdata As Boolean)
    MyCaption = Bdata
    PropertyChanged "Caption"
    UserControl_Resize
End Property


Public Property Get Max() As Long
    Max = MyMax
End Property

Public Property Let Max(ByVal Lmax As Long)
    MyMax = Lmax
    PropertyChanged "Max"
End Property


Public Property Get ProgressStartColor() As OLE_COLOR
    ProgressStartColor = MyStartColor
End Property
Public Property Let ProgressStartColor(ByVal vData As OLE_COLOR)
    MyStartColor = vData
    PicStartColor.BackColor = MyStartColor
    PropertyChanged "ProgressStartColor"
    UserControl_Resize
End Property

Public Property Get BorderColor() As OLE_COLOR
    BorderColor = MyBorderColor
End Property
Public Property Let BorderColor(ByVal vData As OLE_COLOR)
    MyBorderColor = vData
    PropertyChanged "ProgressStartColor"
    UserControl_Resize
End Property

Public Property Get FontColor() As OLE_COLOR
    FontColor = MyFontColor
End Property
Public Property Let FontColor(ByVal vData As OLE_COLOR)
    MyFontColor = vData
    Label1.ForeColor = MyFontColor
PropertyChanged "FontColor"
UserControl_Resize
End Property

Public Property Get Font() As Font
    Set Font = MyFont
End Property
Public Property Set Font(ByVal vData As Font)
    Set MyFont = vData
    Set Label1.Font = MyFont
    Call UserControl_Resize
PropertyChanged "Font"
End Property

Public Property Get ProgressEndColor() As OLE_COLOR
    ProgressEndColor = MyEndColor
End Property
Public Property Let ProgressEndColor(ByVal vData As OLE_COLOR)
    MyEndColor = vData
    PicEndColor.BackColor = MyEndColor
    buffer.BackColor = MyEndColor
    PropertyChanged "ProgressEndColor"
    UserControl_Resize
End Property


Public Property Get Appearance() As AppearanceConst
    Appearance = MyAppearance
End Property
Public Property Let Appearance(ByVal vData As AppearanceConst)
    MyAppearance = vData
    PropertyChanged "Appearance"
    UserControl_Resize
End Property



Public Sub value(value As Long)
On Error Resume Next
Dim ColorArray() As Long
If value > Max Then Exit Sub
Label1.Caption = Int((value / Max) * 100) & " %"
Call mGradient.ColorArray(PicStartColor.BackColor, PicEndColor.BackColor, display.ScaleWidth * (value / Max), ColorArray)
buffer.Cls
For X = 0 To display.ScaleWidth * (value / Max)
    buffer.ForeColor = ColorArray(X)
    buffer.Line (X, 0)-(X, display.ScaleHeight)
Next X
Call BitBlt(display.hDC, 0, 0, display.Width, display.Height, buffer.hDC, 0, 0, vbSrcCopy)
display.Refresh
End Sub

