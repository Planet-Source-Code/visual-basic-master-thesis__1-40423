Attribute VB_Name = "ModPrinter"
Option Explicit

Global MAX_Listbox, MIN_Listbox As Integer
Public HorizontalMargin, VerticalMargin As Single

'----------------------------------------------------------------
'Easily reset font types and sizes from other function/procedure
'----------------------------------------------------------------

Public Sub SetFont(size As Integer, b, i, u, s As Boolean)
'Set the fonts
'The user decides to use a Black ink printer
Printer.ForeColor = RGB(0, 0, 0) 'Black color

'Making Arial font type the default font
Printer.FontName = "Arial"

'These are all variables
Printer.FontSize = size
Printer.FontBold = b
Printer.FontItalic = i
Printer.FontUnderline = u
Printer.FontStrikethru = s
End Sub

'------------------------
'Center Justify
'-------------------------
Public Sub pCenter(ByVal strText As String)
      Printer.CurrentX = ((Printer.ScaleWidth - Printer.TextWidth(strText)) / 2)
End Sub

'------------------------
'Print line
'-------------------------

Public Sub pLineHere(Optional LeftPos As Single = 0)
    Printer.Line (LeftPos, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY)
End Sub

'-------------------
'Check Page length
'-------------------

Public Sub CheckPageLen()
    If pEndOfPage Then
        Printer.NewPage
    End If
End Sub

'------------------------
'Check for End-of-Page
'-------------------------

Public Function pEndOfPage() As Boolean
Dim n As Single
    n = Printer.ScaleHeight - 2
    If Printer.CurrentY = n Then pEndOfPage = True
End Function


'-------------------------------------------------------------
'Print header/footer.
'User suggested using the printer to print the header because
'according to the user, its much more cheaper than using
'normal "letter-heads" A4 paper.
'-------------------------------------------------------------

Public Sub PrintHeader(ByVal Printhead As Object)
Printhead.CurrentY = VerticalMargin - 1
Printhead.CurrentX = HorizontalMargin

Printhead.Print "";

SetFont 36, True, True, True, False
pCenter "Mj Merchandizing, Inc."
Printhead.Print "Mj Merchandizing Inc."

SetFont 10, False, False, False, False
pCenter "(Inayawan Laray Cebu City.)"
Printhead.Print "(Inayawan Laray Cebu City.)"

SetFont 10, False, False, False, False
pCenter "Tel No: 222222222"
Printhead.Print "Tel No: 2222222"
End Sub

'Public Sub PrintFooter(ByVal PrintFoot As Object)
'PrintFoot.CurrentY = Printer.ScaleHeight - 1.5
'PrintFoot.CurrentX = HorizontalMargin

'PrintFoot.Print "";
'SetFont 10, False, False, False, False
'pCenter "No. 37, Jalan Petaling, 50000 Kuala Lumpuer, Malaysia. Tel:603-2011819(6 lines),2389504 Fax: 603-2305388"
'PrintFoot.Print "No. 37, Jalan Petaling, 50000 Kuala Lumpuer, Malaysia. Tel:603-2011819(6 lines),2389504 Fax: 603-2305388"
'End Sub



