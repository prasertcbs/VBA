' Create QR code from selected text in Word Document
' Author: Prasert Kanawattanachai
' prasert@cbs.chula.ac.th

Option Explicit

Public Sub insert_qrcode_selected_text()
' create QR Code from selected text and insert QR code right after selected text
    Dim data
    data = Selection.Text
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.InlineShapes.AddPicture FileName:= _
        qrcode(data), LinkToFile:=False, _
        SaveWithDocument:=True
End Sub

Public Sub insert_qrcode_dialog()
    Dim data
    data = InputBox$("enter data")
    Selection.InlineShapes.AddPicture FileName:= _
        qrcode(data), LinkToFile:=False, _
        SaveWithDocument:=True
End Sub

Private Function qrcode(data)
' API: https://barcode.tec-it.com/
    qrcode = "https://barcode.tec-it.com/barcode.ashx?code=MobileQRUrl&multiplebarcodes=false&translate-esc=false&unit=Fit&dpi=96&imagetype=png&eclevel=M&dmsize=Default&download=true&data=" & data
End Function


