VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmAcronym 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Acronym Help"
   ClientHeight    =   4500
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   5670
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAcronym.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   5670
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSave 
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   5025
      Picture         =   "frmAcronym.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Save List"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.OptionButton optSpc 
      DownPicture     =   "frmAcronym.frx":0B9C
      Enabled         =   0   'False
      Height          =   400
      Index           =   4
      Left            =   4785
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   0
      Width           =   260
   End
   Begin VB.Timer tmrInitialise 
      Interval        =   1
      Left            =   5160
      Top             =   2160
   End
   Begin VB.OptionButton optSpc 
      DownPicture     =   "frmAcronym.frx":0C9E
      Enabled         =   0   'False
      Height          =   400
      Index           =   0
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   0
      Width           =   260
   End
   Begin VB.CommandButton cmdSort 
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Sort"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.OptionButton optSpc 
      DownPicture     =   "frmAcronym.frx":0DA0
      Enabled         =   0   'False
      Height          =   400
      Index           =   3
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   0
      Width           =   260
   End
   Begin VB.Frame fraAdd 
      Height          =   1935
      Left            =   0
      TabIndex        =   9
      Top             =   2520
      Visible         =   0   'False
      Width           =   5655
      Begin VB.CommandButton cmdUpDate 
         Caption         =   "Update"
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   1440
         Width           =   1455
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   3960
         TabIndex        =   17
         Top             =   1440
         Width           =   1455
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         Height          =   375
         Left            =   2160
         TabIndex        =   16
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox txtAddInfo 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3880
         TabIndex        =   12
         Top             =   500
         Width           =   1695
      End
      Begin VB.TextBox txtAddMean 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1720
         TabIndex        =   11
         Top             =   500
         Width           =   2175
      End
      Begin VB.TextBox txtAddAcro 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   40
         TabIndex        =   10
         Top             =   500
         Width           =   1695
      End
      Begin VB.Label lblLastAcro 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   40
         TabIndex        =   21
         Top             =   880
         Width           =   1695
      End
      Begin VB.Label lblLastMean 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1720
         TabIndex        =   20
         Top             =   880
         Width           =   2175
      End
      Begin VB.Label lblLastInfo 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   3880
         TabIndex        =   19
         Top             =   880
         Width           =   1695
      End
      Begin VB.Label lblAddInfo 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Information"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3880
         TabIndex        =   15
         Top             =   200
         Width           =   1695
      End
      Begin VB.Label lblAddMean 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Meaning"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1720
         TabIndex        =   14
         Top             =   200
         Width           =   2175
      End
      Begin VB.Label lblAddAcro 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Acronym"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   40
         TabIndex        =   13
         Top             =   200
         Width           =   1695
      End
   End
   Begin MSFlexGridLib.MSFlexGrid gridAcro 
      Height          =   1935
      Left            =   0
      TabIndex        =   22
      Top             =   420
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   3413
      _Version        =   327680
      Cols            =   3
      FixedCols       =   0
      WordWrap        =   -1  'True
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
   End
   Begin VB.CommandButton cmdDel 
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   4200
      Picture         =   "frmAcronym.frx":10E2
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Deleted Highlighted Item"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.OptionButton optSpc 
      DownPicture     =   "frmAcronym.frx":1854
      Enabled         =   0   'False
      Height          =   400
      Index           =   2
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   0
      Width           =   260
   End
   Begin VB.CommandButton cmdEdit 
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   3360
      Picture         =   "frmAcronym.frx":1B96
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Edit Highlighted Item"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.OptionButton optSpc 
      DownPicture     =   "frmAcronym.frx":2308
      Enabled         =   0   'False
      Height          =   400
      Index           =   1
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   0
      Width           =   260
   End
   Begin VB.CommandButton cmdNew 
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   2520
      Picture         =   "frmAcronym.frx":240A
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "New Item"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmdFind 
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Search"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.ComboBox cboFind 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      ItemData        =   "frmAcronym.frx":2B7C
      Left            =   1305
      List            =   "frmAcronym.frx":2B7E
      TabIndex        =   1
      Top             =   0
      Width           =   390
   End
   Begin VB.Shape shpPgrs 
      BackColor       =   &H8000000F&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   0
      Left            =   0
      Top             =   0
      Width           =   100
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ACRONYMS"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   450
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File          "
      Begin VB.Menu mnuFileOpts 
         Caption         =   "&New"
         Index           =   0
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpts 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuFileOpts 
         Caption         =   "&Edit"
         Index           =   2
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuFileOpts 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuFileOpts 
         Caption         =   "&Save"
         Index           =   4
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileOpts 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuFileOpts 
         Caption         =   "&Delete"
         Index           =   6
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuFileOpts 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnuFileOpts 
         Caption         =   "S&ort"
         Index           =   8
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileOpts 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu mnuFileOpts 
         Caption         =   "E&xit"
         Index           =   10
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About          "
      Begin VB.Menu mnuAboutApp 
         Caption         =   "   Acronym Help Application"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuAboutKH 
         Caption         =   "  Written by - Kev Heywood"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuAboutCopyRight 
         Caption         =   "    Copyright (c) May 2002"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuAboutEmail 
         Caption         =   "kevan@000h.freeserve.co.uk"
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "frmAcronym"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************************************************
'The MsFlexGrid isn't as straight forward as most of the VB controls
'and there are a few Idiosyncrasies to take into account.
'
'This project not ony demonstrates many of the features of the MsFlexGrid,
'but is purposely designed to make the MsFlexGrid behave like a ListBox
'to the extent that when a grid cell is clicked the entire row is highlighted
'and the previous row is restored to a non-highlighted condition.  In addition
'the whole project is a complete application with a data file of over 1500
'computer related Acronyms and their meanings.  Functionality includes -
'Add new entries, Edit existing entries, Deleting, Saving, Sorting and Searching.
'
'The MsFlexGrid is a great control for displaying information in a neat and
'attractive manner and certainly complements data when it needs to be displayed
'in a similar fashion to that of database and spreadsheet design.  You can modify
'this project to suit your own purpose and even add further functionality.  However,
'how the control manages it's internal array is dubious to say the least and losing
'track of the array position of an item can be problematic in it's cause.  Therefore
'you should pay special attention to any Sort, Add, Edit and Delete procedures.
'
'All the best - Kev Heywood - Email kevan@000h.freeserve.co.uk
'***********************************************************************************
Option Explicit
'Integer Declarations
Dim FileNum1 As Integer 'Assign a file number that is not already in use.
Dim FileNum2 As Integer 'Assign a file number that is not already in use.
Dim gSort As Integer    'Assign a value that sorts rows
'String Declarations
Dim acroPath As String  'Assign the Application Path
Dim acroFile As String  'Assign the Acronym Data FileName
Dim Acro As String      'Assign the Acronym Field
Dim Mean As String      'Assign the Meaning Field
Dim Info As String      'Assign the Information Field
'Array Declaratios
Dim Tbox(2) As String   'Assign lblTitle Captions
'Boolean Declarations
Dim itemlistChanged As Boolean  'Assign True when the contents of the MsFlexGrid changes

'******************************************************************************************
                '*********************************************************
                '*                  EVENT DRIVEN PROCEDURES              *
                '*********************************************************
'******************************************************************************************

Private Sub Form_Load()
    'Centre the form and position controls-------------------------------
    Me.Height = 3100
    Me.Width = Screen.Width - 400
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Height = 1100
    
    cboFind.Width = 2500
    
    lblTitle.Width = (Me.Width - cmdSort.Width - optSpc(0).Width - cboFind.Width - _
    cmdFind.Width - cmdNew.Width - optSpc(1).Width - cmdEdit.Width - _
    optSpc(2).Width - optSpc(3).Width - cmdDel.Width - 130) - optSpc(4).Width - cmdSave.Width
    
    cmdSort.Left = lblTitle.Width
    optSpc(0).Left = lblTitle.Width + cmdSort.Width
    cboFind.Left = lblTitle.Width + cmdSort.Width + optSpc(0).Width
    cmdFind.Left = cboFind.Left + cboFind.Width
    optSpc(1).Left = cmdFind.Left + cmdFind.Width
    cmdNew.Left = optSpc(1).Left + optSpc(1).Width
    optSpc(2).Left = cmdNew.Left + cmdNew.Width
    cmdEdit.Left = optSpc(2).Left + optSpc(2).Width
    optSpc(3).Left = cmdEdit.Left + cmdEdit.Width
    cmdDel.Left = optSpc(3).Left + optSpc(3).Width
    optSpc(4).Left = cmdDel.Left + cmdDel.Width
    cmdSave.Left = optSpc(4).Left + optSpc(4).Width
    gridAcro.Width = Me.Width - 130
    '--------------------------------------------------------------------
    'Position Add-Frame and controls-------------------------------
    fraAdd.Width = gridAcro.Width
    fraAdd.Top = gridAcro.Top
    fraAdd.Left = gridAcro.Left
    lblAddAcro.Width = (fraAdd.Width / 100) * 10
    lblAddMean.Width = (fraAdd.Width / 100) * 65
    lblAddInfo.Width = (fraAdd.Width / 100) * 24
    txtAddAcro.Width = lblAddAcro.Width
    txtAddMean.Width = lblAddMean.Width
    txtAddInfo.Width = lblAddInfo.Width
    lblLastAcro.Width = lblAddAcro.Width
    lblLastMean.Width = lblAddMean.Width
    lblLastInfo.Width = lblAddInfo.Width
    lblAddAcro.Left = 40
    lblAddMean.Left = lblAddAcro.Left + lblAddAcro.Width
    lblAddInfo.Left = lblAddMean.Left + lblAddMean.Width
    txtAddAcro.Left = 40
    txtAddMean.Left = txtAddAcro.Left + txtAddAcro.Width
    txtAddInfo.Left = txtAddMean.Left + txtAddMean.Width
    lblLastAcro.Left = 40
    lblLastMean.Left = lblLastAcro.Left + lblLastAcro.Width
    lblLastInfo.Left = lblLastMean.Left + lblLastMean.Width
    cmdUpDate.Left = 200
    cmdClose.Left = fraAdd.Width / 2 - cmdClose.Width / 2
    cmdCancel.Left = fraAdd.Width - cmdCancel.Width - 200
    '--------------------------------------------------------------------
    Tbox(0) = "ACRONYMS": Tbox(1) = "Loading..": Tbox(2) = "Saving..." 'lblTitle Captions
    acroPath = App.Path + "\": 'Acronym File Location"
    acroFile = "MainAcro.Dat" 'Acronym FileName"
    gSort = 1 'Initialse the Sort direction
    cmdFind.Tag = 1 'Initialise to FirstSearch status
    cmdFind.Picture = optSpc(2).DownPicture 'Initialise picture property to FirstSearch
    cmdSort.Picture = optSpc(0).DownPicture 'Initialise picture property to Sort Ascending
    'The width of the each column is a percentage of the width of the Grid control
    gridAcro.ColWidth(0) = (gridAcro.Width / 100) * 12
    gridAcro.ColWidth(1) = (gridAcro.Width / 100) * 50
    gridAcro.ColWidth(2) = (gridAcro.Width / 100) * 35
    
    'Load and Position Progress Array Shapes
    With shpPgrs(0)
        .Width = 100: .Height = 275: .Top = 50: .Left = 80
    End With
    Dim shploop As Integer  'Simple Loop variable
    For shploop = 1 To 29
        Load shpPgrs(shploop)   'Create a new instance of the shape
        shpPgrs(shploop).Left = shpPgrs(shploop).Left _
        + (shpPgrs(shploop - 1).Width - 10) * shploop   'Position shapes in a neat row
        shpPgrs(shploop).ZOrder 0 'Bring To Front
    Next shploop
    
    'Parameters = (Show/Hide, Caption String Array Index, Caption alignment, FontSize)
    Call HideShowProgShp(True, 1, 1, 10)
    
    'Check to see if AcroMain file is present in App directory
    Dim rFileSize As Long   'Return value specifying the length of a file in bytes.
    Dim response As Integer 'Return value indicating which button the user clicked.
    On Error GoTo ERROR_FILE_MISSING    'Error Handler
    rFileSize = FileLen(acroPath + acroFile)      'This forces an error if file missing
    If rFileSize = 0 Then GoTo ERROR_FILE_MISSING 'Treat an empty file as a missing file
    
    Exit Sub 'No errors so exit the sub
    
'Deal with 'Acronym data file missing' Scenario
ERROR_FILE_MISSING:
    response = MsgBox("File " + UCase(acroFile) + " missing from location" + vbCrLf + _
    UCase(acroPath) + vbCrLf + vbCrLf + _
    "To create an empty File Click 'Yes', else if you want to try" + vbCrLf + _
    "and locate this file yourself Click 'No' to exit the program." _
    , vbYesNo + vbCritical, "(ACRONYM HELP) -  WARNING - " + acroFile + " - MISSING")
    If response = vbYes Then
        FileNum1 = FreeFile 'Get next File number
        Open acroPath + acroFile For Output As FileNum1
            'Add three row records (3 because it will make the FlexGrid scrollbar appear)
            Write #FileNum1, "FGMWYDT", "Files Go Missing When You Delete Them", "Check your Recycle Bin"
            Write #FileNum1, "INGAM", "It's No Good Asking Me", "It must be your fault"
            Write #FileNum1, "MSDOS", "MicroSoft Disk Operating System", "CPM reversed engineered"
        Close FileNum1
    Else
        Unload Me 'The Query_Unload event procedure will automatically be called
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    Dim response As Integer 'Return value indicating which button the user clicked.
    Dim uloop As Integer    'Simple Loop variable
    
    If itemlistChanged Then 'Check to see if MsFlexGrid contents have changed
        response = MsgBox("Do you want to save the changes you made?" + vbCrLf + vbCrLf _
        , vbYesNoCancel + vbInformation, "SAVE CHANGES")
        If response = vbYes Then
            Call cmdSave_Click  'Save Changes
        ElseIf response = vbCancel Then
            Cancel = 1: Exit Sub    'Stop the form and application from closing.
        End If
    End If
    
    'Tidy-up by releasing any loaded resources, unload the form and close the application
    For uloop = 1 To 29: Unload shpPgrs(uloop): Next uloop 'Destroy instances of shapes
    Unload Me
    End
End Sub

Private Sub tmrInitialise_Timer()
    'The use of a timer is really just to allow the form to fully load and display
    'before loading in the Acronym data file.
    
    tmrInitialise.Enabled = False    'Disabled the Timer - won't use it again
    gridAcro.Visible = False         'Hide the MsFlexGrid while the data loads
    Call AddEditRecord(False, False) 'Hide and disabled controls while the data loads
    Call PopListView                 'Populate the MsFlexGrid from the Data File
    Me.Height = 3100                 'Adjust Height of form to show MsFlexGrid
    Call HideShowProgShp(False, 0, 2, 16) 'Hide the Progress Bar and set Caption and Font
    Call AddEditRecord(False, True)  'Show and Enabled controls
End Sub

Private Sub gridAcro_Click()
    'Highlight the entire row in a fashion identical to a ListBox
    Call SelectGridCells(gridAcro.Row, False)
End Sub

Private Sub cmdNew_Click()
    fraAdd.Visible = True            'Show the Add/Edit Frame controls
    Call AddEditRecord(True, False)  'Show Frame Add controls and disable form controls
    txtAddAcro.SetFocus
End Sub

Private Sub cmdEdit_Click()
    fraAdd.Visible = True            'Show the Add/Edit Frame controls
    Call AddEditRecord(False, False) 'Hide Frame Add controls and disable form controls
    txtAddAcro = gridAcro.TextMatrix(gridAcro.Row, 0) 'Copy from selected row cell 0
    txtAddMean = gridAcro.TextMatrix(gridAcro.Row, 1) 'Copy from selected row cell 1
    txtAddInfo = gridAcro.TextMatrix(gridAcro.Row, 2) 'Copy from selected row cell 2
    txtAddAcro.SetFocus
End Sub

Private Sub cmdUpDate_Click()
    'Take no action if Acronym text box is empty - notify user
    If txtAddAcro = "" Then MsgBox "Acronym Text Empty.", vbCritical, "Empty Text": Exit Sub
    
    Dim indexLoop As Long   'MsFlexGrid Index loop counter
    Dim newEntry As String  'Concatenation of New Row Data
    
    txtAddAcro = UCase(txtAddAcro)                        'Convert Acronym to upper case
    If lblLastAcro.Visible Then                           'Add a new item
        newEntry = txtAddAcro + vbTab + txtAddMean + vbTab + txtAddInfo
        gridAcro.AddItem newEntry                         'Add to MsFlexGrid
        gridAcro.RowHeight(gridAcro.Rows - 1) = 780       'Set the height of new row
        lblLastAcro.Caption = txtAddAcro                  'Display last added item cell 0
        lblLastMean.Caption = txtAddMean                  'Display last added item cell 1
        lblLastInfo.Caption = txtAddInfo                  'Display last added item cell 2
        Call SelectGridCells(gridAcro.Rows - 1, False)    'Record position of new row
    Else                                                  'Edit highlighted item
        gridAcro.TextMatrix(gridAcro.Row, 0) = txtAddAcro 'Copy to highlighted row cell 0
        gridAcro.TextMatrix(gridAcro.Row, 1) = txtAddMean 'Copy to highlighted row cell 1
        gridAcro.TextMatrix(gridAcro.Row, 2) = txtAddInfo 'Copy to highlighted row cell 2
        Call SelectGridCells(gridAcro.Row, False)         'Record position of editted row
    End If
    
    'We now need to either position the new row item or reposition an editted row item
    'alphabetically in the MsFlexGrid in whatever direction selected by the Sort property.
    
    gridAcro.Col = 2        '***' Setting these properties in this order is vital in
    gridAcro.ColSel = 0     '***' maintaining the integrity of the true location
    gridAcro.Sort = gSort   '***' of the highlighted item.
    
    For indexLoop = 1 To gridAcro.Rows - 1  'Loop through the item collection
        gridAcro.Row = indexLoop    'Set each row and check if highlighted
        If gridAcro.CellBackColor = vbHighlight Then gridAcro.TopRow = indexLoop: Exit For
    Next indexLoop
    
    'Restore the old position to a non-highlighted condition and hightlight the new position
    Call SelectGridCells(indexLoop, False)
    
    txtAddAcro = "": txtAddMean = "": txtAddInfo = ""   'Reset the Text Boxes
    itemlistChanged = True  'Set to record that a change has taken place
    If Not cmdClose.Visible Then
        Call cmdCancel_Click   'If Edit then return to MsFlexGrid
    Else
        txtAddAcro.SetFocus    'If Add new item then move the focus to the Acronym TextBox
    End If
End Sub

Private Sub cmdCancel_Click()
    fraAdd.Visible = False 'Hide the Add/Edit Frame controls
    'Reset Text Boxes
    txtAddAcro = "": txtAddMean = "": txtAddInfo = ""   'Reset the Text Boxes
    Call AddEditRecord(False, True) 'Hide Frame Add controls and Enable form controls
End Sub

Private Sub cmdClose_Click()
    Call cmdCancel_Click    'Return to MsFlexGrid and controls
End Sub

Private Sub cmdDel_Click()
    'Attempting to remove the last non-fixed row causes runtime error 30015,
    'therefore only allow a delete operation when the Rows count is greater than 1
    If gridAcro.Rows > 2 Then
        gridAcro.RemoveItem gridAcro.Row
        Call SelectGridCells(gridAcro.Row, True)
        itemlistChanged = True
    End If
End Sub

Private Sub cmdSave_Click()
    Call AddEditRecord(False, False)        'Disable controls
    Call HideShowProgShp(True, 2, 1, 10)    'Show Progress Shapes
    Dim indexLoop As Long                   'MsFlexGrid Index loop counter
    FileNum2 = FreeFile                     'Get next File number
    'Save the contents of the MsFlexGrid
    Open acroPath + acroFile For Output As FileNum1
        For indexLoop = 1 To gridAcro.Rows - 1
            Write #FileNum2, gridAcro.TextMatrix(indexLoop, 0), gridAcro.TextMatrix(indexLoop, 1), gridAcro.TextMatrix(indexLoop, 2)
            If indexLoop Mod 10 = 0 Then Call ProgressBar   'Progress ratio 10% of file
        Next indexLoop
    Close FileNum2  'Close File
    Call HideShowProgShp(False, 0, 2, 16)   'Hide Progress Shapes
    Call AddEditRecord(False, True)         'Enable controls
    itemlistChanged = False
End Sub

Private Sub mnuFileOpts_Click(Index As Integer)
    'Menu Selections
    Select Case Index
        Case 0
            Call cmdNew_Click
        Case 2
            Call cmdEdit_Click
        Case 4
            Call cmdSave_Click
        Case 6
            Call cmdDel_Click
        Case 8
            Call cmdSort_Click
        Case 10
            Unload Me
    End Select
End Sub

Private Sub cboFind_KeyPress(KeyAscii As Integer)
    'When Return/Enter Key pressed equivalent to clicking the Find command button
    If KeyAscii = 13 Then Call cmdFind_Click
End Sub

Private Sub cboFind_GotFocus()
    cmdFind.Picture = optSpc(2).DownPicture 'Change picture on control to FirstSearch
    cmdFind.Tag = 1                         'Reset NextSearch to FirstSearch status
End Sub

Private Sub cmdFind_Click()
    If cboFind.Text = "" Then
        'Exit the procedure when search string is empty
        Exit Sub
    Else
        Dim itmFound As Boolean 'Test for when search is positive
        Dim idxFindNext As Long 'Position in list to start search
        Dim loopg As Long 'Simple For/Next variable
        idxFindNext = Val(cmdFind.Tag) 'Assign value from Tag property
        If idxFindNext <= gridAcro.Rows - 1 Then 'Search only within the bounds
            Me.MousePointer = 11 'Hourglass
            For loopg = idxFindNext To gridAcro.Rows - 1 'From last position+1 to end
                'Case insensitive comparison
                If UCase(cboFind.Text) = UCase(gridAcro.TextMatrix(loopg, 0)) Then
                    itmFound = True 'Match found so set test and exit loop
                    Exit For
                End If
            Next loopg
        End If
        If Not itmFound Then   ' If no match, inform user and exit.
            Me.MousePointer = 0 'Pointer
            MsgBox "No match found", vbInformation, "Find" 'Inform User
            cmdFind.Tag = 1 'Reset NextSearch to FirstSearch status
            If idxFindNext > 1 Then cmdFind.Picture = optSpc(2).DownPicture 'Reset picture
            Exit Sub
        Else
            Dim Dupf As Boolean 'Test for duplicate entries in the combo list
            Dim loopf As Integer 'Simple For/Next variable
            For loopf = 0 To cboFind.ListCount - 1 'Loop through the Combo List
                'Compare Combo Text property against Combo list entries
                If UCase(cboFind.Text) = UCase(cboFind.List(loopf)) Then Dupf = True: Exit For
            Next loopf
            If Not Dupf Then cboFind.AddItem cboFind.Text 'If not in list then Add
            
            gridAcro.TopRow = loopg     'Scroll FlexGrid to show found Item.
            Call SelectGridCells(loopg, False) 'Highlight found selection
            If idxFindNext = 1 Then cmdFind.Picture = optSpc(3).DownPicture 'Change picture
            idxFindNext = loopg + 1 'Set the position for NextSearch
            cmdFind.Tag = idxFindNext 'Save the NextSearch Position in the Tag property
        End If
    End If
    Me.MousePointer = 0 'Pointer
End Sub

Private Sub cmdSort_Click()
    Dim newPosRow As Long                   'MsFlexGrid index variable
    cmdFind.Tag = 1                         'Reset the Search/SearchNext value
    cmdFind.Picture = optSpc(2).DownPicture 'Swap picture from between controls
    
    'Alternate between Ascending/Descending Sort Order
    If gSort = 1 Then gSort = 2 Else gSort = 1
    
    'Calculate the new position of the highlighted item
    newPosRow = (gridAcro.Rows - gridAcro.Row)
    
    'The Sort property always sorts entire rows.  The keys used for sorting are
    'determined by the Col and ColSel properties, always from the left to the right.
    'For example, if Col = 2 and ColSel = 0, the sort would be done according
    'to the contents of columns 0, then 1, then 2.
    'IMPORTANT - for this to work as intended you must assign a value to the
    'Col property first followed by the ColSel, as below.
    gridAcro.Col = 2
    gridAcro.ColSel = 0
    
    gridAcro.Sort = gSort                   'Activate the new sort order
    gridAcro.TopRow = newPosRow             'Scroll to show the highlighted item
    Call SelectGridCells(newPosRow, False)  'Highlight the new position
    
    'You could use an ImageList or PictureClip control when changing pictures on an object.
    'However, I've used 5 disabled option buttons as spacers so why not use their picture
    'properties. (For Commercial or Professional projects use the appropriate objects,
    'Toolbar, ImageList/PictureClip - if you don't; you will get in a mess when developing
    'further.)
    cmdSort.Picture = optSpc(gSort - 1).DownPicture 'Swap picture from between controls
End Sub

Private Sub txtAddAcro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtAddMean.SetFocus 'On Return Key move focus to next TextBox
End Sub

Private Sub txtAddMean_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtAddInfo.SetFocus 'On Return Key move focus to next TextBox
End Sub

Private Sub txtAddInfo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call cmdUpDate_Click 'On Return Key Activate Update procedure
End Sub

'******************************************************************************************
                '*********************************************************
                '*                  NON-EVENT PROCEDURES                 *
                '*********************************************************
'******************************************************************************************

Private Sub PopListView()
    Dim indexCount As Long  'MsFlexGrid Index array counter
    gridAcro.Clear          'Clear the Grid before re-populating
    'Set column headings
    gridAcro.TextMatrix(0, 0) = "Acronym"
    gridAcro.TextMatrix(0, 1) = "Meaning"
    gridAcro.TextMatrix(0, 2) = "Information"
    FileNum1 = FreeFile     'Get next File number
    Open acroPath + acroFile For Input As FileNum1
    While Not EOF(FileNum1) 'Loop until the end of the file is reached
        Input #FileNum1, Acro, Mean, Info
        indexCount = indexCount + 1 'Increment counter
        If indexCount Mod 10 = 0 Then Call ProgressBar  'Progress ratio 10% of file
        gridAcro.AddItem "" 'Create empty row
        gridAcro.RowHeight(indexCount) = 780 'Set the height of every row
        gridAcro.TextMatrix(indexCount, 0) = Acro 'Assign Acronym to Column 0
        gridAcro.TextMatrix(indexCount, 1) = Mean 'Assign Meaning to Column 1
        gridAcro.TextMatrix(indexCount, 2) = Info 'Assign Information to Column 2
        'Align the text in the columns
        gridAcro.ColAlignment(0) = 1
        gridAcro.ColAlignment(1) = 1
        gridAcro.ColAlignment(2) = 1
    Wend
    Close FileNum1  'Close the File
    
    'Strange one this! But the OCX seems to add an unwanted entry at the end, nothing
    'to do with the code - still never mind just remove it. (Thanks Bill - another problem)
    gridAcro.RemoveItem indexCount + 1
    
    gridAcro.Visible = True         'Show the MsFlexGrid
    gridAcro.Sort = gSort           'Set Ascending Sort
    gridAcro.TopRow = 1             'Initialise at the first row
    gridAcro.Col = 0                'Initialise at the first column
    Call SelectGridCells(1, False)  'Highlight the first row
End Sub

Private Sub SelectGridCells(GridR As Long, DeletedEntrie As Boolean)
    
    Dim loopn As Integer        'Simple Loop variable
    Static prevGridR As Long    'Save previous highlighted selection
    
    'Set the below property at design time in the property box
    'to ensure you have control over highlighted selections.
    '--------------------------------------------------------'
    'gridAcro.HighLight = flexHighlightNever
    
    
    'IMPORTANT - Determine which column should be left active before
    'leaving this Procedure.  This will affect the sort property.  In
    'this app I leave the first column active by implementing reverse
    'loops - the first column is the 'Acronym' which we sort and search by.
    'Note: If forward loops were used then the third column would be the
    'active one and the Generic Ascending/Descending Sort would use the
    'Information column which is contrary to the design intentions.
    
    'REMOVE PREVIOUSLY HIGHLIGHTED ROW/COL'S
    If prevGridR <> 0 And Not DeletedEntrie Then
        gridAcro.Row = prevGridR
        For loopn = gridAcro.Cols - 1 To 0 Step -1
            gridAcro.Col = loopn
            gridAcro.CellBackColor = vbWindowBackground
            gridAcro.CellForeColor = vbWindowText
        Next loopn
    End If
    'HIGHLIGHTED SELECTED ROW/COL'S
    gridAcro.Row = GridR
    For loopn = gridAcro.Cols - 1 To 0 Step -1
        gridAcro.Col = loopn
        gridAcro.CellBackColor = vbHighlight
        gridAcro.CellForeColor = vbHighlightText
    Next loopn
    prevGridR = GridR
    'Note: msflexgrid.row = GridR.  Column will be the first column.
    Me.Caption = "Acronym Help  -  Record Number " + Format(GridR, "0000#")
End Sub

Private Sub AddEditRecord(vCntrls As Boolean, eCntrls As Boolean)
    'One procedure to Hide/Show and Enable/Disable controls to prevent the user
    'from taking actions inappropiate to the requested task.  Another way would be to
    'create other forms and call these modally, but for a small application it is
    'preferable to keep code together.
    lblLastAcro.Visible = vCntrls
    lblLastMean.Visible = vCntrls
    lblLastInfo.Visible = vCntrls
    cmdClose.Visible = vCntrls
    mnuFile.Enabled = eCntrls
    mnuAbout.Enabled = eCntrls
    cmdSort.Enabled = eCntrls
    cmdFind.Enabled = eCntrls
    cboFind.Enabled = eCntrls
    cmdNew.Enabled = eCntrls
    cmdEdit.Enabled = eCntrls
    cmdDel.Enabled = eCntrls
    cmdSave.Enabled = eCntrls
End Sub

Private Sub HideShowProgShp(ShowShp As Boolean, TboxCaption As Integer, CapAlign As Integer, TitleFontSize As Integer)
    Dim hsloop As Integer   'Simple Loop variable
    For hsloop = 0 To 29
        shpPgrs(hsloop).BackColor = &H8000000F  'Reset the BackColor of the shapes
        shpPgrs(hsloop).Visible = ShowShp       'Hide or Show the shapes
    Next hsloop
    'Set the Title caption, Alignment and Font size
    lblTitle.Caption = Tbox(TboxCaption): lblTitle.Alignment = CapAlign: lblTitle.FontSize = TitleFontSize
End Sub

Private Sub ProgressBar()
    Static progCount As Integer                             'Counter variable
    Static colCount As Integer                              'Counter variable
    Dim chgcol(2) As Long                                   'Backcolor array
    chgcol(0) = 255: chgcol(1) = 65280: chgcol(2) = 65535   'Colors - Red, Green, Yellow
    shpPgrs(progCount).BackColor = chgcol(colCount)         'Assign new colour
    DoEvents                                                'Process other events
    progCount = progCount + 1                               'Increment counter
    If progCount > 29 Then                                  'Keep within array bounds
        progCount = 0                                       'Reset counter
        colCount = colCount + 1                             'Increment counter
        If colCount > 2 Then colCount = 0                   'Keep within array bounds
    End If
End Sub
