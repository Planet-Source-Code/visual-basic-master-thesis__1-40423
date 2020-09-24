Attribute VB_Name = "MODMAIN"

Option Explicit

Public OrdAmount As Double
Public Qty As Long
Public OcashStat As String
Public Status As String
Public ord As String
Public DBMain As DAO.Database
Public WSMain As DAO.Workspace
Public AddStocks As String

Public Const Alphatext = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ "
Public Const variable$ = "0123456789"
Public Const num = "0123456789."
Dim r As String

Public Function AutoOrder() As String
    Dim TQR As DAO.QueryDef
    Dim TRS As DAO.Recordset
    Dim Start As String
    Start = "0000"
    Set TQR = DBMain.CreateQueryDef("", "SELECT ordno FROM Sales")
    Set TRS = TQR.OpenRecordset()
    Do While Not TRS.EOF
        TRS.FindFirst "Ordno='O-" + Start & "'"
        If Not TRS.NoMatch Then
            Start = Format(Str(Val(Mid$(Start, 3)) + 1), "0000")
        Else
            AutoOrder = "O-" + Start
            Exit Function
        End If
    Loop
    AutoOrder = "O-" + Start
End Function

Public Function AutoDel() As String
    Dim TQR As DAO.QueryDef
    Dim TRS As DAO.Recordset
    Dim Start As String
    Start = "0000"
    Set TQR = DBMain.CreateQueryDef("", "SELECT delno FROM Deliveries")
    Set TRS = TQR.OpenRecordset()
    Do While Not TRS.EOF
        TRS.FindFirst "Delno='D-" + Start + "'"
        If Not TRS.NoMatch Then
            Start = Format(Str(Val(Mid$(Start, 3)) + 1), "0000")
        Else
            AutoDel = "D-" + Start
            Exit Function
        End If
    Loop
    AutoDel = "D-" + Start
End Function

Public Function AutoSupplier() As String
    Dim TQR As DAO.QueryDef
    Dim TRS As DAO.Recordset
    Dim Start As String
    Start = "0000"
    Set TQR = DBMain.CreateQueryDef("", "SELECT suppno FROM suppliers")
    Set TRS = TQR.OpenRecordset()
    Do While Not TRS.EOF
        TRS.FindFirst "SUPPNO='" + Start + "'"
        If Not TRS.NoMatch Then
            Start = Format(Str(Val(Mid$(Start, 3)) + 1), "0000")
        Else
            AutoSupplier = Start
            Exit Function
        End If
    Loop
    AutoSupplier = Start
End Function

Public Function AutoStock() As String
    Dim TQR As DAO.QueryDef
    Dim TRS As DAO.Recordset
    Dim Start As String
    Start = "0000"
    Set TQR = DBMain.CreateQueryDef("", "SELECT stockno FROM stocks")
    Set TRS = TQR.OpenRecordset()
    Do While Not TRS.EOF
        TRS.FindFirst "stockno='I-" + Start + "'"
        If Not TRS.NoMatch Then
            Start = Format(Str(Val(Mid$(Start, 3)) + 1), "0000")
        Else
            AutoStock = "I-" + Start
            Exit Function
        End If
    Loop
    AutoStock = "I-" + Start
End Function





Public Sub Main()
    'AnalyzeRegistry
 
    ' Create ODBC workspace
     Set WSMain = DBEngine.Workspaces(0)
    ' Open Main Database
    Set DBMain = WSMain.OpenDatabase(App.Path + "\Database\MasterDB.mdb")
    MDIMAIN.Show
End Sub

' Enables / disables textbox in fields of customers
Public Sub SetText(ByVal Status As Boolean)
    Dim x As Long
    For x = 1 To 6
        FRMCUSTOMERS.TXTF(x).Enabled = Status
    Next x
End Sub

Public Sub FillDatas()
    On Error GoTo trapper
    Dim e As Long
    Dim x As Long
    Dim TQR As DAO.QueryDef
    Dim TRS As DAO.Recordset
    Dim LISTCUST  As ListItem
    Set TQR = DBMain.QueryDefs("FILLDATAS")
    Set TRS = TQR.OpenRecordset()
    With FRMCUSTOMERS
    Do While Not TRS.EOF
        Set LISTCUST = .LSTCUSTOMERS.ListItems.Add(, , TRS.Fields(0), , 1)
        LISTCUST.SubItems(1) = TRS.Fields(1)
        LISTCUST.SubItems(2) = TRS.Fields(2)
        LISTCUST.SubItems(3) = TRS.Fields(3)
        LISTCUST.SubItems(4) = TRS.Fields(4)
        LISTCUST.SubItems(5) = TRS.Fields(5)
        LISTCUST.SubItems(6) = TRS.Fields(6)
        TRS.MoveNext
    Loop
    End With
    With FRMCUSTOMERS
        .TXTF(0).Text = .LSTCUSTOMERS.ListItems.Item(1).Text
        For x = 1 To 6
            .TXTF(x).Text = .LSTCUSTOMERS.ListItems.Item(1).SubItems(x)
        Next x
    End With
    Exit Sub
trapper:
End Sub
Public Sub ClearTextE()
    Dim x As Long
    For x = 1 To 4
        FRMSTOCKS.TXTF(x).Text = ""
    Next x
End Sub


Public Sub FillStocks()
    On Error GoTo trapper
    Dim TQR As DAO.QueryDef
    Dim TRS As DAO.Recordset
    Dim List As ListItem
    
    Set TQR = DBMain.QueryDefs("AllStocks")
    Set TRS = TQR.OpenRecordset()
    Do While Not TRS.EOF
        Set List = FRMSTOCKS.LSTSTOCKS.ListItems.Add(, , TRS.Fields(0), , 1)
        With List
            .SubItems(1) = TRS.Fields(1)
            .SubItems(2) = Format(TRS.Fields(2), "###,###,##0.00")
            .SubItems(3) = TRS.Fields(3)
            .SubItems(4) = TRS.Fields(4)
        End With
        TRS.MoveNext
    Loop
    With FRMSTOCKS.LSTSTOCKS
        FRMSTOCKS.TXTF(0).Text = .ListItems.Item(1).Text
        FRMSTOCKS.TXTF(1).Text = .ListItems.Item(1).SubItems(1)
        FRMSTOCKS.TXTF(2).Text = .ListItems.Item(1).SubItems(2)
        FRMSTOCKS.TXTF(3).Text = .ListItems.Item(1).SubItems(3)
        FRMSTOCKS.TXTF(4).Text = .ListItems.Item(1).SubItems(4)
        '.SetFocus
    End With
trapper:
End Sub
Public Sub ShowStocks()
    On Error GoTo trapper
    Dim TQR As DAO.QueryDef
    Dim TRS As DAO.Recordset
    Dim List As ListItem
    
    Set TQR = DBMain.QueryDefs("ShowStocks")
    Set TRS = TQR.OpenRecordset()
    Do While Not TRS.EOF
        Set List = FRMDAMAGED.LSTASTOCKS.ListItems.Add(, , TRS.Fields(0), , 0)
        With List
            .SubItems(1) = TRS.Fields(1)
            .SubItems(2) = TRS.Fields(2)
            .SubItems(3) = TRS.Fields(3)
             End With
        TRS.MoveNext
    Loop
FRMDAMAGED.Label1.Caption = "Number of Records:" & " " & TRS.RecordCount
trapper:
End Sub
Public Sub DamagedStocks()
    'On Error GoTo trapper
    Dim TQR As DAO.QueryDef
    Dim TRS As DAO.Recordset
    Dim List As ListItem
    Dim QryStr As String
    QryStr = "SELECT Damaged.stockno,damaged.desc,damaged.qty,damaged.[date_ret] FROM damaged WHERE damaged.replaced=False;"
    Set TQR = DBMain.CreateQueryDef("", QryStr)
    Set TRS = TQR.OpenRecordset()
    Do While Not TRS.EOF
        Set List = FRMLDAMAGED.LSTDAMAGED.ListItems.Add(, , TRS.Fields(0), , 0)
        With List
            .SubItems(1) = TRS.Fields(1)
            .SubItems(2) = TRS.Fields(2)
            .SubItems(3) = TRS.Fields(3)
             End With
        TRS.MoveNext
    Loop
 End Sub

Public Function GetQuantity() As Long
    FRMDLG.Show vbModal
    GetQuantity = Qty
End Function
Public Sub UpdateStockQuantity(ByVal Quantity As Long, ByVal StkNo As String)
    DBMain.Execute "UPDATE stocks SET quantity = quantity + " & Str(Quantity) & " WHERE stockno ='" & StkNo & "';"
End Sub

Public Sub UpdateStocksQuantity(ByVal Qty As Long, ByVal Stk As String)
   DBMain.Execute "UPDATE stocks SET quantity =  " & Str(Qty) & " WHERE stockno ='" & Stk & "';"
End Sub
Public Sub Shutdown()
    If MsgBox("Are you sure you want to Shutdown ? ", vbExclamation + vbYesNo, "Shutdown") = vbYes Then
        Unload FRMMAINMENU
        Unload MDIMAIN
    
    End If
End Sub
'*******************************************************


Public Sub DisplayRecord(ByVal List As ListView, ByVal Record As DAO.Recordset, ByVal NFields As Long, ByVal Inum As Long)
    Dim LST As ListItem
    Dim x As Long
    List.ListItems.Clear
    Do While Not Record.EOF
        Set LST = List.ListItems.Add(, , Record.Fields(0), , Inum)
        With LST
            For x = 1 To NFields
                .SubItems(x) = Record.Fields(x)
            Next x
        End With
        Record.MoveNext
    Loop
End Sub




Public Sub RecordDelivery(ByVal stockno As String, ByVal suppno As String, ByVal Quantity As String, ByVal PDate As String, ByVal Delno As String)
    Dim Query As String
    Query = "INSERT INTO deliveries (stockno,suppno,quantity,[date],delno) VALUES ('" & stockno & "','" & suppno & "'," & Quantity & ",#" & PDate & "#,'" & Delno & "');"
    DBMain.Execute Query
    Query = "UPDATE stocks SET quantity = quantity + " & Quantity & " WHERE stockno ='" & stockno & "';"
    DBMain.Execute Query
End Sub




Public Sub AllowOnlyIntegers(KeyAscii As Integer)

    If KeyAscii <> 8 Then
       If InStr(variable, Chr(KeyAscii)) = 0 Then
          KeyAscii = 0
            Exit Sub
        End If
    End If
End Sub
Public Sub Integers1(KeyAscii As Integer)

    If KeyAscii <> 8 Then
       If InStr(num, Chr(KeyAscii)) = 0 Then
          KeyAscii = 0
            Exit Sub
        End If
    End If
End Sub
Public Sub Alpha(KeyAscii As Integer)

    If KeyAscii <> 8 Then
       If InStr(Alphatext, Chr(KeyAscii)) = 0 Then
          KeyAscii = 0
            Exit Sub
        End If
    End If
End Sub

