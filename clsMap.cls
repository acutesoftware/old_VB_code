VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarGridWidth As Long 'local copy
Private mvarGridHeight As Long 'local copy
Private mvarx_current As Long 'local copy
Private mvary_current As Long 'local copy
Private mvarCellValue(16, 16) As Integer 'local copy
'local variable(s) to hold property value(s)
Private mvarCellPositionX(16, 16) As Integer 'local copy
Private mvarCellPositionY(16, 16) As Integer 'local copy
Private mvarCellDetails(16, 16) As Variant 'local copy
Public Property Let CellDetails(X As Long, y As Long, ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CellDetails = 5
    mvarCellDetails(X, y) = vData
End Property


Public Property Set CellDetails(X As Long, y As Long, ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.CellDetails = Form1
    Set mvarCellDetails(X, y) = vData
End Property


Public Property Get CellDetails(X As Long, y As Long) As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CellDetails
    If IsObject(mvarCellDetails) Then
        Set CellDetails = mvarCellDetails(X, y)
    Else
        CellDetails = mvarCellDetails(X, y)
    End If
End Property



Public Property Let CellPositionY(X As Long, y As Long, ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CellPositionY = 5
    mvarCellPositionY(X, y) = vData
End Property


Public Property Get CellPositionY(X As Long, y As Long) As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CellPositionY
    CellPositionY = mvarCellPositionY(X, y)
End Property



Public Property Let CellPositionX(X As Long, y As Long, ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CellPositionX = 5
    mvarCellPositionX(X, y) = vData
End Property


Public Property Get CellPositionX(X As Long, y As Long) As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CellPositionX
    CellPositionX = mvarCellPositionX(X, y)
End Property



Public Property Let CellValue(X As Long, y As Long, ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CellValue = 5
    mvarCellValue(X, y) = vData
End Property


Public Property Get CellValue(X As Long, y As Long) As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CellValue
    CellValue = mvarCellValue(X, y)
End Property



Public Function IsEmpty(X As Long, y As Long) As Boolean
End Function

Public Property Let y_current(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.y_current = 5
    mvary_current = vData
End Property

Public Function CalculateCurrentCoordinates(cur As Integer) As Integer
    'this function gets the current tile clicked 'cur'
    'and finds its co-ordinates in the grid and then
    'assigns the current x,y co-ordinates to Xcur and Ycur
    Dim countX As Long
    Dim countY As Long
    Debug.Print "LOOKING FOR Cell " & Str$(cur)
    For countX = 1 To GridWidth
        For countY = 1 To GridHeight
            If CellDetails(countX, countY) = cur Then   'found it!!
                y_current = countY
                x_current = countX
                Debug.Print "Found Cell " & Str$(cur) & " at grid pos" & Str$(countX) & Str$(countY)
                Exit Function
            End If
        Next countY
    Next countX
                
                
    
End Function

Public Property Get y_current() As Long
Attribute y_current.VB_UserMemId = 0
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.y_current
    y_current = mvary_current
End Property



Public Property Let x_current(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.x_current = 5
    mvarx_current = vData
End Property


Public Property Get x_current() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.x_current
    x_current = mvarx_current
End Property



Public Property Let GridHeight(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.GridHeight = 5
    mvarGridHeight = vData
End Property


Public Property Get GridHeight() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.GridHeight
    GridHeight = mvarGridHeight
End Property



Public Property Let GridWidth(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.GridWidth = 5
    mvarGridWidth = vData
End Property


Public Property Get GridWidth() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.GridWidth
    GridWidth = mvarGridWidth
End Property


Public Function Randomise()
    'this function shuffles the grid CellDetails
    Dim num As Integer
    Dim startx, starty As Integer
    Dim spacing As Integer
    Dim found As Boolean
    spacing = 735
    Dim i As Long, j As Long, ii As Long, jj As Long
    For i = 1 To GridWidth
        For j = 1 To GridHeight
            CellDetails(i, j) = ""  'clear grid
        Next j
    Next i
    
    For i = 1 To GridWidth
        For j = 1 To GridHeight
            found = False
            While Not found
                num = Int(Rnd(1) * 16) + 1
                'is it unique? - loop through grid
                 found = True
                 For ii = 1 To GridWidth
                    For jj = 1 To GridHeight
                        If CellDetails(ii, jj) = Str$(num) Then
                            found = False
                        End If
                    Next jj
                 Next ii
            Wend
            CellDetails(i, j) = Str$(num)
            'CellPositionX(i, j) = CInt(startx + (spacing * (i)) - spacing / 2)
            'CellPositionY(i, j) = CInt(starty + (spacing * (i)) - spacing / 2)
        Next j
    Next i
                 
                
End Function
