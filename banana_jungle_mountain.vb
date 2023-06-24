Option Explicit

Public Function WaterFiltration(ByVal waterSource As Integer, ByVal waterAmount As Integer, ByVal filterType As Integer) As Integer

'Declare Variables
Dim gallonsPerHour As Integer
Dim totalCost As Integer
Dim totalGallons As Integer
Dim filterCapacity As Integer
Dim gallonsRemaining As Integer

'Set initial values of variables
gallonsPerHour = 0
totalCost = 0
totalGallons = 0
filterCapacity = 0

'Check the source of water to determine gallons per hour
Select Case waterSource
    Case 0 'Tap Water
        gallonsPerHour = 10
    Case 1 'Well Water
        gallonsPerHour = 5
    Case 2 'River Water
        gallonsPerHour = 3
End Select

'Check the filter type to determine filter capacity
Select Case filterType
    Case 0 'Sediment Filter
        filterCapacity = 1000
    Case 1 'Carbon Filter
        filterCapacity = 1500
    Case 2 'Reverse Osmosis Filter
        filterCapacity = 2000
End Select

'Calculate total gallons needed
totalGallons = waterAmount * gallonsPerHour

'Calculate total filter capacity needed
totalCost = totalGallons / filterCapacity

'Calculate gallons remaining after filtration
gallonsRemaining = waterAmount - totalCost

'Return the total cost of the system
WaterFiltration = totalCost

End Function

Public Sub Main()

'Declare Variables
Dim gallons As Integer
Dim waterSource As Integer
Dim filterType As Integer

'Get the amount of water in gallons from user
gallons = InputBox("Please enter the amount of water you would like to filter (in gallons):")

'Check for valid input
If (gallons <= 0) Then
    MsgBox "Invalid input. Please enter a valid number."
    Exit Sub
End If

'Get the source of water from user
waterSource = InputBox("Please enter the source of the water to be filtered: 0 for tap, 1 for well, 2 for river.")

'Check for valid input
If (waterSource < 0) Or (waterSource > 2) Then
    MsgBox "Invalid input. Please enter a valid number."
    Exit Sub
End If

'Get the type of filter from user
filterType = InputBox("Please enter the type of filter to be used: 0 for sediment, 1 for carbon, 2 for reverse osmosis.")

'Check for valid input
If (filterType < 0) Or (filterType > 2) Then
    MsgBox "Invalid input. Please enter a valid number."
    Exit Sub
End If

'Calculate the total cost of the water filtration system
Dim cost As Integer
cost = WaterFiltration(waterSource, gallons, filterType)

'Display the total cost of the system
MsgBox "The total cost of the water filtration system is " & CStr(cost) & "."

End Sub