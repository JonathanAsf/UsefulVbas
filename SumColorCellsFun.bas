'This code calculate the quantity of cells that has the interior color that you pick
'A cell need to be the color that we want to sum the values that have it
Function SumColor(MatchColor As Range, SumRange As Range) As Double
 Dim cell As Range 'The range the the Sum will be done
 Dim myColor As Long 'The color that we will pick
 myColor = MatchColor.Cells(1, 1).Interior.Color 

 
 For Each cell In SumRange
  If cell.Interior.Color = myColor Then
    SumColor = SumColor + cell.Value
  End If
 Next cell
End Function
