Attribute VB_Name = "Module1"
Option Explicit

Function CalculateAdCost(numAds As Integer) As Double
    Dim costPerAd As Double
    
    ' Look up the cost per ad based on the number of ads
    If numAds <= 5 Then
        costPerAd = Application.WorksheetFunction.VLookup(numAds, Worksheets("Ad Cost").Range("A2:B5"), 2, True)
    ElseIf numAds <= 10 Then
        costPerAd = Application.WorksheetFunction.VLookup(numAds, Worksheets("Ad Cost").Range("A6:B9"), 2, True)
    ElseIf numAds <= 20 Then
        costPerAd = Application.WorksheetFunction.VLookup(numAds, Worksheets("Ad Cost").Range("A10:B13"), 2, True)
    Else
        costPerAd = Application.WorksheetFunction.VLookup(numAds, Worksheets("Ad Cost").Range("A14:B17"), 2, True)
    End If
    
    ' Calculate the total cost
    CalculateAdCost = numAds * costPerAd
End Function


