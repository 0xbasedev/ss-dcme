VERSION 5.00
Begin VB.UserControl cUpDown 
   ClientHeight    =   300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   750
   ScaleHeight     =   20
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   50
   ToolboxBitmap   =   "cUpDown.ctx":0000
   Begin VB.VScrollBar scroll 
      Height          =   285
      Left            =   480
      Max             =   1
      Min             =   2
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Value           =   1
      Width           =   255
   End
   Begin VB.TextBox txtValue 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Text            =   "0"
      Top             =   0
      Width           =   495
   End
End
Attribute VB_Name = "cUpDown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Valeurs de propriétés par défaut:
Const m_def_BackStyle = 0
'Variables de propriétés:
Dim m_BackStyle As Integer
'Déclarations d'événements:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Se produit lorsque l'utilisateur appuie sur un bouton de la souris puis le relâche au-dessus d'un objet."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=scroll,scroll,-1,KeyDown
Attribute KeyDown.VB_Description = "Se produit lorsque l'utilisateur appuie sur une touche alors qu'un objet a le focus."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=scroll,scroll,-1,KeyPress
Attribute KeyPress.VB_Description = "Se produit lorsque l'utilisateur appuie sur une touche ANSI puis la relâche ."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=scroll,scroll,-1,KeyUp
Attribute KeyUp.VB_Description = "Se produit lorsque l'utilisateur relâche une touche alors qu'un objet a le focus."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Se produit lorsque l'utilisateur appuie sur le bouton de la souris alors qu'un objet a le focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Se produit lorsque l'utilisateur déplace la souris."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Se produit lorsque l'utilisateur relâche le bouton de la souris alors qu'un objet a le focus."
Event change() 'MappingInfo=scroll,scroll,-1,Change
Attribute change.VB_Description = "Se produit lorsque le contenu d'un contrôle a été modifié."


'ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
'MappingInfo=txtValue,txtValue,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Renvoie ou définit la couleur d'arrière-plan utilisée pour afficher le texte et les graphiques d'un objet."
10        BackColor = txtValue.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
10        txtValue.BackColor() = New_BackColor
20        PropertyChanged "BackColor"
End Property

'ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
'MappingInfo=txtValue,txtValue,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Renvoie ou définit la couleur de premier plan utilisée pour afficher le texte et les graphiques d'un objet."
10        ForeColor = txtValue.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
10        txtValue.ForeColor() = New_ForeColor
20        PropertyChanged "ForeColor"
End Property

'ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Renvoie ou définit une valeur qui détermine si un objet peut répondre à des événements générés par l'utilisateur."
10        Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
10        UserControl.Enabled() = New_Enabled
20        PropertyChanged "Enabled"
End Property

'ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
'MappingInfo=txtValue,txtValue,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Renvoie un objet Font."
Attribute Font.VB_UserMemId = -512
10        Set Font = txtValue.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
10        Set txtValue.Font = New_Font
20        PropertyChanged "Font"
End Property

'ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
'MemberInfo=7,0,0,0
Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "Indique si un contrôle Label ou l'arrière-plan d'un contrôle Shape sont transparent ou opaque."
10        BackStyle = m_BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
10        m_BackStyle = New_BackStyle
20        PropertyChanged "BackStyle"
End Property

'ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
'MappingInfo=UserControl,UserControl,-1,BorderStyle
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Renvoie ou définit le style de la bordure d'un objet."
10        BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
10        UserControl.BorderStyle() = New_BorderStyle
20        PropertyChanged "BorderStyle"
End Property

'ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
'MemberInfo=5
Public Sub Refresh()
Attribute Refresh.VB_Description = "Force un nouvel affichage complet d'un objet."
     
End Sub

Private Sub txtValue_Change()
10        Call removeDisallowedCharacters(txtValue, scroll.Max, scroll.Min, False)
          
20        scroll.value = val(txtValue.Text)

End Sub

Private Sub UserControl_Click()
10        RaiseEvent Click
End Sub

Private Sub scroll_KeyDown(KeyCode As Integer, Shift As Integer)
10        RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub scroll_KeyPress(KeyAscii As Integer)
10        RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub scroll_KeyUp(KeyCode As Integer, Shift As Integer)
10        RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
10        RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
10        RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
10        RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub scroll_Change()
10        txtValue.Text = scroll.value
          
20        RaiseEvent change
End Sub

'ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
'MappingInfo=scroll,scroll,-1,Min
Public Property Get Min() As Integer
Attribute Min.VB_Description = "Renvoie ou définit un paramètre maximal de propriété Value relatif à la position de la barre de défilement."
10        Min = scroll.Max
End Property

Public Property Let Min(ByVal New_Min As Integer)
10        scroll.Max() = New_Min
20        PropertyChanged "Min"
End Property

'ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
'MappingInfo=scroll,scroll,-1,Max
Public Property Get Max() As Integer
Attribute Max.VB_Description = "Renvoie ou définit un paramètre maximal de propriété Value relatif à la position de la barre de défilement."
10        Max = scroll.Min
End Property

Public Property Let Max(ByVal New_Max As Integer)
10        scroll.Min = New_Max
20        PropertyChanged "Max"
End Property

'ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
'MappingInfo=scroll,scroll,-1,Value
Public Property Get value() As Integer
Attribute value.VB_Description = "Renvoie ou définit la valeur d'un objet."
10        value = scroll.value
End Property

Public Property Let value(ByVal New_Value As Integer)
10        scroll.value = New_Value
20        txtValue.Text = scroll.value
30        PropertyChanged "Value"
End Property

'Initialiser les propriétés pour le contrôle utilisateur
Private Sub UserControl_InitProperties()
10        m_BackStyle = m_def_BackStyle
End Sub

'Charger les valeurs des propriétés à partir du stockage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

10        txtValue.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
20        txtValue.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
30        UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
40        Set txtValue.Font = PropBag.ReadProperty("Font", Ambient.Font)
50        m_BackStyle = PropBag.ReadProperty("BackStyle", m_def_BackStyle)
60        UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
70        scroll.Min = PropBag.ReadProperty("Max", 32767)
80        scroll.Max = PropBag.ReadProperty("Min", 0)
90        scroll.value = PropBag.ReadProperty("Value", 0)
End Sub

'Écrire les valeurs des propriétés dans le stockage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

10        Call PropBag.WriteProperty("BackColor", txtValue.BackColor, &H80000005)
20        Call PropBag.WriteProperty("ForeColor", txtValue.ForeColor, &H80000008)
30        Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
40        Call PropBag.WriteProperty("Font", txtValue.Font, Ambient.Font)
50        Call PropBag.WriteProperty("BackStyle", m_BackStyle, m_def_BackStyle)
60        Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
70        Call PropBag.WriteProperty("Min", scroll.Max, 0)
80        Call PropBag.WriteProperty("Max", scroll.Min, 32767)
90        Call PropBag.WriteProperty("Value", scroll.value, 0)
End Sub



Private Sub removeDisallowedCharacters(ByRef txtbox As TextBox, lowerBound As Single, upperBound As Single, Optional dec As Boolean = False)
10        If lowerBound > upperBound Then
20            txtbox.Text = lowerBound
30            Exit Sub
40        End If
          
50        If (Not IsNumeric(txtbox.Text) And (txtbox.Text <> "-" Or lowerBound >= 0)) _
              Or InStr(txtbox.Text, "e") > 0 Or InStr(txtbox.Text, "E") > 0 _
              Or (Not dec And (InStr(txtbox.Text, ".") > 0 Or InStr(txtbox.Text, ",") > 0)) _
              Or (lowerBound < 0 And InStr(2, txtbox.Text, "-") > 1) _
              Or (lowerBound >= 0 And InStr(txtbox.Text, "-") > 0) Then
              
              Dim oldselstart As Integer
60            oldselstart = txtbox.selstart - 1    'char  typed so always one more
70            If oldselstart < 0 Then oldselstart = 0

              'remove all characters aside from nrs
              Dim i As Integer
              Dim finalresult As String
80            For i = 1 To Len(txtbox.Text)
90                If (Asc(Mid$(txtbox.Text, i, 1)) < Asc("0") Or _
                      Asc(Mid$(txtbox.Text, i, 1)) > Asc("9")) Then
                      Dim result As String
100                   If i - 1 >= 1 Then result = Mid$(txtbox.Text, 1, i - 1)
110                   If i + 1 <= Len(txtbox.Text) Then result = result + Mid$(txtbox.Text, i + 1, Len(txtbox.Text) - (i))
120                   finalresult = result
130               End If
140           Next
150           txtbox.Text = finalresult
160           If oldselstart > Len(txtbox.Text) Then
170               txtbox.selstart = Len(txtbox.Text)
180           Else
190               txtbox.selstart = oldselstart
200           End If
210       End If

220       If val(txtbox.Text) < lowerBound Then
230           txtbox.Text = lowerBound
240       End If

250       If val(txtbox.Text) > upperBound Then
260           txtbox.Text = upperBound
270       End If

End Sub

