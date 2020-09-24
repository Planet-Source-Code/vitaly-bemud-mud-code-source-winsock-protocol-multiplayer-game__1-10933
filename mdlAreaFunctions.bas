Attribute VB_Name = "mdlAreaFunctions"
Option Explicit
'Gets the description of the current room
Function CurrentDesc(Index As Integer) As String
    CurrentDesc = GetIni("Room " & Char(Index).locX & "," & Char(Index).locY, "Desc", "Void")
End Function
'Gets the aviable exits of the current room
Function CurrentExits(Index As Integer) As String
    CurrentExits = GetIni("Room " & Char(Index).locX & "," & Char(Index).locY, "Exits", "Void")
End Function
'Gets the location to where the exit will lead
Function FutureExit(Index As Integer, Direction As String) As String
    FutureExit = GetIni("Room " & Char(Index).locX & "," & Char(Index).locY, Direction, "Void")
End Function
'Gets Y of the room Vnum
Function GetExitY(RoomLocation As String)
GetExitY = Right(RoomLocation, Len(RoomLocation) - InStr(RoomLocation, ","))
End Function
'Gets X of the room Vnum
Function GetExitX(RoomLocation As String)
GetExitX = Left(RoomLocation, Len(RoomLocation) - InStr(RoomLocation, ","))
End Function
'Adds the PC to the room records
Sub AddPC(Index)
PutIni "Room " & Char(Index).locX & "," & Char(Index).locY, "PCs", GetIni("Room " & Char(Index).locX & "," & Char(Index).locY, "PCs", "Void") & "[" & Index & "]", "Void"
End Sub
'Removes the PC to the room records
Sub RemovePC(Index)
PutIni "Room " & Char(Index).locX & "," & Char(Index).locY, "PCs", Replace(GetIni("Room " & Char(Index).locX & "," & Char(Index).locY, "PCs", "Void"), "[" & Index & "]", ""), "Void"
End Sub
Function PCs(Index As Integer) As Variant
PCs = GetIni("Room " & Char(Index).locX & "," & Char(Index).locY, "PCs", "Void")
End Function
