VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Converter"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5295
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   5295
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   3975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   7011
      _Version        =   393216
      Tabs            =   6
      TabHeight       =   520
      TabCaption(0)   =   "Length"
      TabPicture(0)   =   "frmMain.frx":1CCA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblAnswer(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cboLengthTo"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cboLengthFrom"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdConvert(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtConvert(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Area"
      TabPicture(1)   =   "frmMain.frx":1CE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cboAreaTo"
      Tab(1).Control(1)=   "cboAreaFrom"
      Tab(1).Control(2)=   "cmdConvert(1)"
      Tab(1).Control(3)=   "txtConvert(1)"
      Tab(1).Control(4)=   "lblAnswer(1)"
      Tab(1).Control(5)=   "Label1(3)"
      Tab(1).Control(6)=   "Label1(2)"
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "Weight"
      TabPicture(2)   =   "frmMain.frx":1D02
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cboWeightTo"
      Tab(2).Control(1)=   "cboWeightFrom"
      Tab(2).Control(2)=   "cmdConvert(2)"
      Tab(2).Control(3)=   "txtConvert(2)"
      Tab(2).Control(4)=   "lblAnswer(2)"
      Tab(2).Control(5)=   "Label1(5)"
      Tab(2).Control(6)=   "Label1(4)"
      Tab(2).ControlCount=   7
      TabCaption(3)   =   "Volume"
      TabPicture(3)   =   "frmMain.frx":1D1E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "cboVolumeTo"
      Tab(3).Control(1)=   "cboVolumeFrom"
      Tab(3).Control(2)=   "cmdConvert(4)"
      Tab(3).Control(3)=   "txtConvert(4)"
      Tab(3).Control(4)=   "lblAnswer(4)"
      Tab(3).Control(5)=   "Label1(9)"
      Tab(3).Control(6)=   "Label1(8)"
      Tab(3).ControlCount=   7
      TabCaption(4)   =   "Capacity"
      TabPicture(4)   =   "frmMain.frx":1D3A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "cboCapacityTo"
      Tab(4).Control(1)=   "cboCapacityFrom"
      Tab(4).Control(2)=   "cmdConvert(3)"
      Tab(4).Control(3)=   "txtConvert(3)"
      Tab(4).Control(4)=   "lblAnswer(3)"
      Tab(4).Control(5)=   "Label1(7)"
      Tab(4).Control(6)=   "Label1(6)"
      Tab(4).ControlCount=   7
      TabCaption(5)   =   "Temperature"
      TabPicture(5)   =   "frmMain.frx":1D56
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "optConvertToF"
      Tab(5).Control(1)=   "optConvertToC"
      Tab(5).Control(2)=   "txtConvertTemp"
      Tab(5).Control(3)=   "cmdConvertTemp"
      Tab(5).Control(4)=   "Label2"
      Tab(5).Control(5)=   "lblTempAnswer"
      Tab(5).ControlCount=   6
      Begin VB.ComboBox cboVolumeTo 
         Height          =   315
         Left            =   -71880
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   1080
         Width           =   1815
      End
      Begin VB.ComboBox cboVolumeFrom 
         Height          =   315
         Left            =   -74760
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   1080
         Width           =   1815
      End
      Begin VB.CommandButton cmdConvert 
         Caption         =   "&Convert"
         Height          =   495
         Index           =   4
         Left            =   -73200
         TabIndex        =   36
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox txtConvert 
         Height          =   285
         Index           =   4
         Left            =   -72840
         TabIndex        =   35
         Top             =   1440
         Width           =   855
      End
      Begin VB.ComboBox cboCapacityTo 
         Height          =   315
         Left            =   -71880
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   1080
         Width           =   1815
      End
      Begin VB.ComboBox cboCapacityFrom 
         Height          =   315
         Left            =   -74760
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   1080
         Width           =   1815
      End
      Begin VB.CommandButton cmdConvert 
         Caption         =   "&Convert"
         Height          =   495
         Index           =   3
         Left            =   -73200
         TabIndex        =   29
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox txtConvert 
         Height          =   285
         Index           =   3
         Left            =   -72840
         TabIndex        =   28
         Top             =   1440
         Width           =   855
      End
      Begin VB.ComboBox cboWeightTo 
         Height          =   315
         Left            =   -71880
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   1080
         Width           =   1815
      End
      Begin VB.ComboBox cboWeightFrom 
         Height          =   315
         Left            =   -74760
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   1080
         Width           =   1815
      End
      Begin VB.CommandButton cmdConvert 
         Caption         =   "&Convert"
         Height          =   495
         Index           =   2
         Left            =   -73200
         TabIndex        =   22
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox txtConvert 
         Height          =   285
         Index           =   2
         Left            =   -72840
         TabIndex        =   21
         Top             =   1440
         Width           =   855
      End
      Begin VB.ComboBox cboAreaTo 
         Height          =   315
         Left            =   -71880
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   1080
         Width           =   1815
      End
      Begin VB.ComboBox cboAreaFrom 
         Height          =   315
         Left            =   -74760
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   1080
         Width           =   1815
      End
      Begin VB.CommandButton cmdConvert 
         Caption         =   "&Convert"
         Height          =   495
         Index           =   1
         Left            =   -73200
         TabIndex        =   15
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox txtConvert 
         Height          =   285
         Index           =   1
         Left            =   -72840
         TabIndex        =   14
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox txtConvert 
         Height          =   285
         Index           =   0
         Left            =   2160
         TabIndex        =   8
         Top             =   1440
         Width           =   855
      End
      Begin VB.CommandButton cmdConvert 
         Caption         =   "&Convert"
         Height          =   495
         Index           =   0
         Left            =   1800
         TabIndex        =   7
         Top             =   1800
         Width           =   1575
      End
      Begin VB.OptionButton optConvertToF 
         Caption         =   "Convert to Fahrenheit"
         Height          =   375
         Left            =   -74760
         TabIndex        =   6
         Top             =   840
         Width           =   1935
      End
      Begin VB.OptionButton optConvertToC 
         Caption         =   "Convert to Celsius"
         Height          =   375
         Left            =   -74760
         TabIndex        =   5
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox txtConvertTemp 
         Height          =   285
         Left            =   -74040
         TabIndex        =   4
         Top             =   1665
         Width           =   855
      End
      Begin VB.CommandButton cmdConvertTemp 
         Caption         =   "Convert Temperature"
         Height          =   495
         Left            =   -72600
         TabIndex        =   3
         Top             =   1200
         Width           =   1815
      End
      Begin VB.ComboBox cboLengthFrom 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1080
         Width           =   1815
      End
      Begin VB.ComboBox cboLengthTo 
         Height          =   315
         Left            =   3120
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label lblAnswer 
         BorderStyle     =   1  'Fixed Single
         Height          =   975
         Index           =   4
         Left            =   -74880
         TabIndex        =   41
         Top             =   2880
         Width           =   4935
      End
      Begin VB.Label Label1 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   -71880
         TabIndex        =   40
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "From"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   -74760
         TabIndex        =   39
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblAnswer 
         BorderStyle     =   1  'Fixed Single
         Height          =   975
         Index           =   3
         Left            =   -74880
         TabIndex        =   34
         Top             =   2880
         Width           =   4935
      End
      Begin VB.Label Label1 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   -71880
         TabIndex        =   33
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "From"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   -74760
         TabIndex        =   32
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblAnswer 
         BorderStyle     =   1  'Fixed Single
         Height          =   975
         Index           =   2
         Left            =   -74880
         TabIndex        =   27
         Top             =   2880
         Width           =   4935
      End
      Begin VB.Label Label1 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   -71880
         TabIndex        =   26
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "From"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   -74760
         TabIndex        =   25
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblAnswer 
         BorderStyle     =   1  'Fixed Single
         Height          =   975
         Index           =   1
         Left            =   -74880
         TabIndex        =   20
         Top             =   2880
         Width           =   4935
      End
      Begin VB.Label Label1 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   -71880
         TabIndex        =   19
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "From"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   -74760
         TabIndex        =   18
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "From"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   13
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   3120
         TabIndex        =   12
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblAnswer 
         BorderStyle     =   1  'Fixed Single
         Height          =   975
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   2880
         Width           =   4935
      End
      Begin VB.Label Label2 
         Caption         =   "Convert:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   10
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label lblTempAnswer 
         BorderStyle     =   1  'Fixed Single
         Height          =   975
         Left            =   -74880
         TabIndex        =   9
         Top             =   2880
         Width           =   4935
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdConvert_Click(Index As Integer)
    'Which Convert button was clicked? Sub calls are self-explanatory
    Select Case Index
        Case 0
            Call convertLength
        Case 1
            Call convertArea
        Case 2
            Call convertWeight
        Case 3
            Call convertCapacity
        Case 4
            Call convertVolume
    End Select
End Sub

Private Sub cmdConvertTemp_Click()
    
    Dim sngToConvert, sngAnswer As Single
    Dim sConvertFrom, sConvertTo As String
    
    On Error GoTo ErrorHandle

    If optConvertToC.Value = True Then 'from fahrenheit to celsius
        sConvertFrom = "Fahrenheit"
        sConvertTo = "Celsius"
        sngAnswer = (txtConvertTemp.Text - 32) * 5 / 9
    Else
        sConvertFrom = "Celsius" 'vice-versa
        sConvertTo = "Fahrenheit"
        sngAnswer = (txtConvertTemp.Text * 9 / 5) + 32
    End If
    
    'Display
    lblTempAnswer.Caption = txtConvertTemp.Text & " degrees " & _
        sConvertFrom & " is equal to " & _
        Round(sngAnswer, 4) & " degrees " & _
        sConvertTo & "."
    
    Exit Sub

ErrorHandle:
    MsgBox "Invalid entry - try again or quit.", vbOKOnly + vbCritical, "Error"
    Exit Sub

End Sub

Private Sub Form_Load()
    'Add items to comboboxes
    With cboLengthFrom
        .AddItem "Inches"
        .AddItem "Feet"
        .AddItem "Yards"
        .AddItem "Miles"
        .AddItem "Millimeters"
        .AddItem "Centimeters"
        .AddItem "Meters"
        .AddItem "Kilometers"
        .ListIndex = 0 'Inches is default for this combobox
    End With
    With cboLengthTo
        .AddItem "Inches"
        .AddItem "Feet"
        .AddItem "Yards"
        .AddItem "Miles"
        .AddItem "Millimeters"
        .AddItem "Centimeters"
        .AddItem "Meters"
        .AddItem "Kilometers"
        .ListIndex = 1 'Feet default - rest of comboboxes are similar
    End With
    
    With cboAreaFrom
        .AddItem "Sq Inches"
        .AddItem "Sq Feet"
        .AddItem "Sq Yards"
        .AddItem "Sq Miles"
        .AddItem "Sq Centimetres"
        .AddItem "Sq Meters"
        .AddItem "Sq Kilometres"
        .AddItem "Acres"
        .AddItem "Hectares"
        .ListIndex = 0
    End With
    With cboAreaTo
        .AddItem "Sq Inches"
        .AddItem "Sq Feet"
        .AddItem "Sq Yards"
        .AddItem "Sq Miles"
        .AddItem "Sq Centimetres"
        .AddItem "Sq Meters"
        .AddItem "Sq Kilometres"
        .AddItem "Acres"
        .AddItem "Hectares"
        .ListIndex = 1
    End With
    
    With cboWeightFrom
        .AddItem "Ounces"
        .AddItem "Pounds"
        .AddItem "Stones"
        .AddItem "Tons"
        .AddItem "Grams"
        .AddItem "Kilgrams"
        .AddItem "Ton (Metric)"
        .ListIndex = 0
    End With
    With cboWeightTo
        .AddItem "Ounces"
        .AddItem "Pounds"
        .AddItem "Stones"
        .AddItem "Tons"
        .AddItem "Grams"
        .AddItem "Kilgrams"
        .AddItem "Ton (Metric)"
        .ListIndex = 1
    End With
    
    With cboCapacityFrom
        .AddItem "Pints"
        .AddItem "Gallons"
        .AddItem "Millilitres"
        .AddItem "Litres"
        .ListIndex = 0
    End With
    With cboCapacityTo
        .AddItem "Pints"
        .AddItem "Gallons"
        .AddItem "Millilitres"
        .AddItem "Litres"
        .ListIndex = 1
    End With
    
    With cboVolumeFrom
        .AddItem "Cu Inches"
        .AddItem "Cu Feet"
        .AddItem "Cu Yards"
        .AddItem "Cu Centimetres"
        .AddItem "Cu Metres"
        .ListIndex = 0
    End With
    With cboVolumeTo
        .AddItem "Cu Inches"
        .AddItem "Cu Feet"
        .AddItem "Cu Yards"
        .AddItem "Cu Centimetres"
        .AddItem "Cu Metres"
        .ListIndex = 1
    End With
    
    'Default termperature convert is Celsius to Fahrenheit
    optConvertToF.Value = True
    
End Sub

Sub convertLength()
    Dim sngToConvert As Single
    Dim sIndexSelect As String
    Dim sngMult As Single
    
    On Error GoTo ErrorHandle
    
    sngToConvert = txtConvert(0).Text
    
    'Concatenate combobox selections to form single string
    sIndexSelect = Trim(Str(cboLengthFrom.ListIndex)) + Trim(Str(cboLengthTo.ListIndex))
    
    'Select case of concatenated strings
    Select Case sIndexSelect
        Case "00", "11", "22", "33", "44", "55", "66", "77"
            sngMult = 1         'if both options are the same, then answer is equal
        Case "01"               'i.e. 1 foot equals one foot.
            sngMult = 1 / 12    'All other cases must apply formulae that convert
        Case "02"               'one option to another.
            sngMult = 1 / 36    'Damned if I remember where I found them all...!
        Case "03"
            sngMult = 1 / 63360
        Case "04"
            sngMult = 25.4
        Case "05"
            sngMult = 2.54
        Case "06"
            sngMult = 0.0254
        Case "07"
            sngMult = 0.0000254
        Case "10"
            sngMult = 12
        Case "12"
            sngMult = 1 / 3
        Case "13"
            sngMult = 1 / 5280
        Case "14"
            sngMult = 304.8
        Case "15"
            sngMult = 30.48
        Case "16"
            sngMult = 0.3048
        Case "17"
            sngMult = 0.0003048
        Case "20"
            sngMult = 36
        Case "21"
            sngMult = 3
        Case "23"
            sngMult = 1 / 1760
        Case "24"
            sngMult = 914.4
        Case "25"
            sngMult = 91.44
        Case "26"
            sngMult = 0.9144
        Case "27"
            sngMult = 0.0009144
        Case "30"
            sngMult = 63360
        Case "31"
            sngMult = 5280
        Case "32"
            sngMult = 1760
        Case "34"
            sngMult = 1609344
        Case "35"
            sngMult = 160934.4
        Case "36"
            sngMult = 1609.344
        Case "37"
            sngMult = 1.609344
        Case "40"
            sngMult = 0.03937
        Case "41"
            sngMult = 0.003280833
        Case "42"
            sngMult = 0.001093611
        Case "43"
            sngMult = 0.00000062137
        Case "45"
            sngMult = 0.1
        Case "46"
            sngMult = 0.001
        Case "47"
            sngMult = 0.000001
        Case "50"
            sngMult = 0.3937
        Case "51"
            sngMult = 0.03280833
        Case "52"
            sngMult = 0.01093611
        Case "53"
            sngMult = 0.0000062137
        Case "54"
            sngMult = 10
        Case "56"
            sngMult = 0.01
        Case "57"
            sngMult = 0.00001
        Case "60"
            sngMult = 39.37
        Case "61"
            sngMult = 3.280833
        Case "62"
            sngMult = 1.093611
        Case "63"
            sngMult = 0.00062137
        Case "64"
            sngMult = 1000
        Case "65"
            sngMult = 100
        Case "67"
            sngMult = 0.001
        Case "70"
            sngMult = 39370
        Case "71"
            sngMult = 3280.833
        Case "72"
            sngMult = 1093.611
        Case "73"
            sngMult = 0.62137
        Case "74"
            sngMult = 1000000
        Case "75"
            sngMult = 100000
        Case "76"
            sngMult = 1000
    End Select
    
    'Concatenated a string to show in the answer labels.
    'Called calcConvert function (see botton) to universally
    'convert one item to another.
    lblAnswer(0).Caption = txtConvert(0).Text & " " & _
        cboLengthFrom.Text & " is equal to " & _
        calcConvert(sngToConvert, sngMult) & " " & _
        cboLengthTo.Text & "."
        
    Exit Sub

ErrorHandle:
    MsgBox "Invalid entry - try again or quit.", vbOKOnly + vbCritical, "Error"
    Exit Sub
    'All other convert subs are more or less the same, except for the formulae.
End Sub

Sub convertArea()
    Dim sngToConvert As Single
    Dim sIndexSelect As String
    Dim sngMult As Single
    
    On Error GoTo ErrorHandle
    
    sngToConvert = txtConvert(1).Text
    
    sIndexSelect = Trim(Str(cboAreaFrom.ListIndex)) + Trim(Str(cboAreaTo.ListIndex))
    
    Select Case sIndexSelect
        Case "00", "11", "22", "33", "44", "55", "66", "77", "88"
            sngMult = 1
        Case "01"
            sngMult = 0.0069445161
        Case "02"
            sngMult = 0.0007716064
        Case "03"
            sngMult = 0.0000000002491
        Case "04"
            sngMult = 6.4516129032
        Case "05"
            sngMult = 0.0006451612
        Case "06"
            sngMult = 0.0000000006452
        Case "07"
            sngMult = 0.0000001594
        Case "08"
            sngMult = 0.0000000645
        Case "10"
            sngMult = 144
        Case "12"
            sngMult = 0.1111111111
        Case "13"
            sngMult = 0.00000003587006
        Case "14"
            sngMult = 929.03129906
        Case "15"
            sngMult = 0.0929031299
        Case "16"
            sngMult = 0.00000009290304
        Case "17"
            sngMult = 0.0000229567
        Case "18"
            sngMult = 0.0000092903
        Case "20"
            sngMult = 1296
        Case "21"
            sngMult = 9
        Case "23"
            sngMult = 0.0000003228305
        Case "24"
            sngMult = 8361.2739236
        Case "25"
            sngMult = 0.8361273923
        Case "26"
            sngMult = 0.0000008361273
        Case "27"
            sngMult = 0.0002066104
        Case "28"
            sngMult = 0.0000836127
        Case "30"
            sngMult = 4014459600#
        Case "31"
            sngMult = 27878400
        Case "32"
            sngMult = 3097600
        Case "34"
            sngMult = 25899881103.36
        Case "35"
            sngMult = 2589988.110336
        Case "36"
            sngMult = 2.589988110336
        Case "37"
            sngMult = 640
        Case "38"
            sngMult = 259.004451639
        Case "40"
            sngMult = 0.155
        Case "41"
            sngMult = 0.00107639
        Case "42"
            sngMult = 0.000119599
        Case "43"
            sngMult = 0.00000000003861
        Case "45"
            sngMult = 0.0001
        Case "46"
            sngMult = 0.0000000001
        Case "47"
            sngMult = 0.0000000247104
        Case "48"
            sngMult = 0.00000001
        Case "50"
            sngMult = 1550
        Case "51"
            sngMult = 10.7639
        Case "52"
            sngMult = 1.19599
        Case "53"
            sngMult = 0.00000038610215
        Case "54"
            sngMult = 10000
        Case "56"
            sngMult = 0.000001
        Case "57"
            sngMult = 0.000247104
        Case "58"
            sngMult = 0.0001
        Case "60"
            sngMult = 1550003100.006
        Case "61"
            sngMult = 10763910.4167
        Case "62"
            sngMult = 1195990.046301
        Case "63"
            sngMult = 0.386102158542
        Case "64"
            sngMult = 10000000000#
        Case "65"
            sngMult = 1000000
        Case "67"
            sngMult = 247.1053814671
        Case "68"
            sngMult = 100
        Case "70"
            sngMult = 6272662.52262
        Case "71"
            sngMult = 43560
        Case "72"
            sngMult = 4840.02687127
        Case "73"
            sngMult = 0.0015625
        Case "74"
            sngMult = 40468790.4687
        Case "75"
            sngMult = 4046.87904687
        Case "76"
            sngMult = 0.004046856422
        Case "78"
            sngMult = 0.40468790468
        Case "80"
            sngMult = 15500000
        Case "81"
            sngMult = 107639
        Case "82"
            sngMult = 11960
        Case "83"
            sngMult = 0.0038609375
        Case "84"
            sngMult = 100000000
        Case "85"
            sngMult = 10000
        Case "86"
            sngMult = 0.01
        Case "87"
            sngMult = 2.47104
    End Select
    
    lblAnswer(1).Caption = txtConvert(1).Text & " " & _
        cboAreaFrom.Text & " is equal to " & _
        calcConvert(sngToConvert, sngMult) & " " & _
        cboAreaTo.Text & "."
        
    Exit Sub

ErrorHandle:
    MsgBox "Invalid entry - try again or quit.", vbOKOnly + vbCritical, "Error"
    Exit Sub

End Sub

Sub convertWeight()
    Dim sngToConvert As Single
    Dim sIndexSelect As String
    Dim sngMult As Single
    
    On Error GoTo ErrorHandle
    
    sngToConvert = txtConvert(2).Text
    
    sIndexSelect = Trim(Str(cboWeightFrom.ListIndex)) + Trim(Str(cboWeightTo.ListIndex))
    
    Select Case sIndexSelect
        Case "00", "11", "22", "33", "44", "55", "66"
            sngMult = 1
        Case "01"
            sngMult = 0.0625
        Case "02"
            sngMult = 0.004464286
        Case "03"
            sngMult = 0.000027901785
        Case "04"
            sngMult = 28.349523125
        Case "05"
            sngMult = 0.028349523125
        Case "06"
            sngMult = 0.000028349518
        Case "10"
            sngMult = 16
        Case "12"
            sngMult = 1 / 14
        Case "13"
            sngMult = 0.000446428571
        Case "14"
            sngMult = 453.59237
        Case "15"
            sngMult = 0.45359237
        Case "16"
            sngMult = 0.00045359237
        Case "20"
            sngMult = 224
        Case "21"
            sngMult = 14
        Case "23"
            sngMult = 0.00625
        Case "24"
            sngMult = 6350.29318
        Case "25"
            sngMult = 6.35029318
        Case "26"
            sngMult = 0.00635029138
        Case "30"
            sngMult = 35840
        Case "31"
            sngMult = 2240
        Case "32"
            sngMult = 160
        Case "34"
            sngMult = 1016046.9088
        Case "35"
            sngMult = 1016.0469088
        Case "36"
            sngMult = 1.01604673452
        Case "40"
            sngMult = 0.035273961949
        Case "41"
            sngMult = 0.002204622621
        Case "42"
            sngMult = 0.000157473044
        Case "43"
            sngMult = 0.000000984206
        Case "45"
            sngMult = 0.001
        Case "46"
            sngMult = 0.000001
        Case "50"
            sngMult = 35.27396194958
        Case "51"
            sngMult = 2.204622621848
        Case "52"
            sngMult = 0.157473044417
        Case "53"
            sngMult = 0.000984206527
        Case "54"
            sngMult = 1000
        Case "56"
            sngMult = 0.001
        Case "60"
            sngMult = 35273.968
        Case "61"
            sngMult = 2204.623
        Case "62"
            sngMult = 157.4730714286
        Case "63"
            sngMult = 0.984206696429
        Case "64"
            sngMult = 1000000
        Case "65"
            sngMult = 1000
    End Select
    
    lblAnswer(2).Caption = txtConvert(2).Text & " " & _
        cboWeightFrom.Text & " is equal to " & _
        calcConvert(sngToConvert, sngMult) & " " & _
        cboWeightTo.Text & "."
    
    Exit Sub

ErrorHandle:
    MsgBox "Invalid entry - try again or quit.", vbOKOnly + vbCritical, "Error"
    Exit Sub
        
End Sub

Sub convertCapacity()
    Dim sngToConvert As Single
    Dim sIndexSelect As String
    Dim sngMult As Single
    
    On Error GoTo ErrorHandle
    
    sngToConvert = txtConvert(3).Text
    
    sIndexSelect = Trim(Str(cboCapacityFrom.ListIndex)) + Trim(Str(cboCapacityTo.ListIndex))
    
    Select Case sIndexSelect
        Case "00", "11", "22", "33"
            sngMult = 1
        Case "01"
            sngMult = 0.125
        Case "02"
            sngMult = 568.26125
        Case "03"
            sngMult = 0.56826125
        Case "10"
            sngMult = 8
        Case "12"
            sngMult = 4546
        Case "13"
            sngMult = 4.546
        Case "20"
            sngMult = 0.001759
        Case "21"
            sngMult = 0.0002199
        Case "23"
            sngMult = 0.001
        Case "30"
            sngMult = 1 / 0.56826125
        Case "31"
            sngMult = 0.2199
        Case "32"
            sngMult = 1000
    End Select
    
    lblAnswer(3).Caption = txtConvert(3).Text & " " & _
        cboCapacityFrom.Text & " is equal to " & _
        calcConvert(sngToConvert, sngMult) & " " & _
        cboCapacityTo.Text & "."
    
    Exit Sub

ErrorHandle:
    MsgBox "Invalid entry - try again or quit.", vbOKOnly + vbCritical, "Error"
    Exit Sub

End Sub

Sub convertVolume()
    Dim sngToConvert As Single
    Dim sIndexSelect As String
    Dim sngMult As Single
    
    On Error GoTo ErrorHandle
    
    sngToConvert = txtConvert(4).Text
    
    sIndexSelect = Trim(Str(cboVolumeFrom.ListIndex)) + Trim(Str(cboVolumeTo.ListIndex))
    
    Select Case sIndexSelect
        Case "00", "11", "22", "33", "44"
            sngMult = 1
        Case "01"
            sngMult = 0.000578703703
        Case "02"
            sngMult = 0.00002143347
        Case "03"
            sngMult = 16.387064
        Case "04"
            sngMult = 0.000016387064
        Case "10"
            sngMult = 1728
        Case "12"
            sngMult = 0.037037037037
        Case "13"
            sngMult = 28316.846592
        Case "14"
            sngMult = 0.028316846592
        Case "20"
            sngMult = 46656
        Case "21"
            sngMult = 27
        Case "23"
            sngMult = 764554.857984
        Case "24"
            sngMult = 0.764554857984
        Case "30"
            sngMult = 0.061023744094
        Case "31"
            sngMult = 0.000035314666
        Case "32"
            sngMult = 0.00000130795
        Case "34"
            sngMult = 0.000001
        Case "40"
            sngMult = 61023.74409473
        Case "41"
            sngMult = 35.31466672148
        Case "42"
            sngMult = 1.307950619314
        Case "43"
            sngMult = 1000000
    End Select
    
    lblAnswer(4).Caption = txtConvert(4).Text & " " & _
        cboVolumeFrom.Text & " is equal to " & _
        calcConvert(sngToConvert, sngMult) & " " & _
        cboVolumeTo.Text & "."
        
    Exit Sub

ErrorHandle:
    MsgBox "Invalid entry - try again or quit.", vbOKOnly + vbCritical, "Error"
    Exit Sub
        
End Sub

Function calcConvert(sngToConvertMult As Single, sngMultSelect As Single) As Single
    'Not exactly rocket science!
    calcConvert = Round((sngToConvertMult * sngMultSelect), 4)

End Function
