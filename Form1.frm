VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   Caption         =   "Demo membuat menu dengan objek TreeView"
   ClientHeight    =   5295
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7365
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5295
   ScaleWidth      =   7365
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.TreeView mnuTree 
      Height          =   4935
      Left            =   195
      TabIndex        =   1
      Top             =   195
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   8705
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Enabled         =   0   'False
      Height          =   5085
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3420
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'***************************************************************************
' MMMM  MMMMM  OMMM   MMMO    OMMM    OMMM    OMMMMO     OMMMMO    OMMMMO  '
'  MM    MM   MM MM    MMMO  OMMM    MM MM    MM   MO   OM    MO  OM    MO '
'  MM  MM    MM  MM    MM  OO  MM   MM  MM    MM   MO   OM    MO       OMO '
'  MMMM     MMMMMMMM   MM  MM  MM  MMMMMMMM   MMMMMO     OMMMMO      OMO   '
'  MM  MM        MM    MM      MM       MM    MM   MO   OM    MO   OMO     '
'  MM    MM      MM    MM      MM       MM    MM    MO  OM    MO  OM   MM  '
' MMMM  MMMM    MMMM  MMMM    MMMM     MMMM  MMMM  MMMM  OMMMMO   MMMMMMM  '
'                                                                          '
' K4m4r82's Laboratory                                                     '
' http://coding4ever.wordpress.com                                         '
'***************************************************************************

Private Sub addMenu()
    Dim rsMenuInduk As ADODB.Recordset
    Dim rsMenuAnak  As ADODB.Recordset
    
    Dim root        As Node
    Dim i           As Long
    Dim x           As Long
    Dim rowCount(1) As Long
    
    mnuTree.Nodes.Clear
    With mnuTree.Nodes
        'menampilkan menu induk
        strSql = "SELECT id, menu_name, menu_caption " & _
                 "FROM menu_induk " & _
                 "ORDER BY id"
        Set rsMenuInduk = openRecordset(strSql)
        If Not rsMenuInduk.EOF Then
            rowCount(0) = getRecordCount(rsMenuInduk)
            
            For i = 1 To rowCount(0)
                Set root = .Add(, , rsMenuInduk("menu_name").Value, rsMenuInduk("menu_caption").Value)
                root.Bold = True
    
                'menampilkan menu anak
                strSql = "SELECT menu_name, menu_caption " & _
                         "FROM menu_anak " & _
                         "WHERE menu_induk_id = " & rsMenuInduk("id").Value & " " & _
                         "ORDER BY id"
                Set rsMenuAnak = openRecordset(strSql)
                If Not rsMenuAnak.EOF Then
                    rowCount(1) = getRecordCount(rsMenuAnak)
                    
                    For x = 1 To rowCount(1)
                        .Add root, tvwChild, rsMenuAnak("menu_name").Value, rsMenuAnak("menu_caption").Value
                        
                        rsMenuAnak.MoveNext
                    Next x
                End If
                Call closeRecordset(rsMenuAnak)
                
                rsMenuInduk.MoveNext
            Next i
        End If
        Call closeRecordset(rsMenuInduk)
    End With
    
    For i = 1 To mnuTree.Nodes.Count
        mnuTree.Nodes(i).Expanded = True
    Next
End Sub

Private Sub Form_Load()
    Dim ret As Boolean
    
    ret = KonekToServer
    
    'inisialisasi treeview
    With mnuTree
        .Style = tvwTreelinesPlusMinusText
        .LineStyle = tvwRootLines
        .Indentation = 300.47
    End With
    
    Call addMenu
End Sub

Private Sub mnuTree_DblClick()
    If mnuTree.Nodes(mnuTree.SelectedItem.Index).Children = 0 Then 'menu anak
        Select Case mnuTree.SelectedItem.Key
            Case "mnuBarang": 'TODO : tampilkan frmBarang disini
            Case "mnuCustomer"
            Case "mnuSupplier"
            Case "mnuPembelian"
            Case "mnuReturPembelian"
            Case "mnuPenjualan"
            Case "mnuBiayaOperasional"
            Case "mnuGajiKaryawan"
            Case "mnuLapPembelian"
            Case "mnuLapJthTempo"
            Case "mnuLapPenjualan"
        End Select
    End If
End Sub
