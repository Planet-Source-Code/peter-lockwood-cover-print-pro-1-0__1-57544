Attribute VB_Name = "ModPrint"
'---Print Module By i8cRoWts@Hotmail.co.uk---'

Public Sub PrintFront()
Printer.Orientation = 1 'Prints in protrait
 Printer.ScaleMode = vbMillimeters  'Size In Millimeters
  Printer.PaintPicture Main.front.Picture, 10, 10, 121.5, 121.45
   Printer.EndDoc
End Sub

Public Sub PrintBack()
Printer.Orientation = 1 'Prints in protrait
 Printer.ScaleMode = vbMillimeters  'Size In Millimeters
  Printer.PaintPicture Main.back.Picture, 10, 122.45, 149.7, 117.6
   Printer.EndDoc
End Sub

Public Sub PrintBoth()
Printer.Orientation = 1 'Prints in protrait
 Printer.ScaleMode = vbMillimeters  'Size In Millimeters
  Printer.PaintPicture Main.front.Picture, 7, 7, 121.5, 121.45
   Printer.PaintPicture Main.back.Picture, 7, 140.3, 149.7, 117.6
    Printer.EndDoc
End Sub
