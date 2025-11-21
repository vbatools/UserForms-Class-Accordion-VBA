# Implementation Examples for VBA Acordion Class

## Introduction

This document contains various implementation examples of the `clsAcardion` class in VBA. Examples cover basic and advanced usage scenarios, demonstrating the flexibility and functionality of the class.

## Example 1: Simple accordion on a form

### Description
Creating a simple accordion with three items on a custom form.

### Code
```vba
Sub CreateSimpleAccordion()
    Dim acardion As New clsAcardion
    acardion.SetParentForm Me
    
    ' Add accordion items
    acardion.AddItem "Introduction", "This is introductory information about the project"
    acardion.AddItem "Features", "List of main application features"
    acardion.AddItem "Contacts", "Information for contacting the developer"
    
    ' Create controls on the form
    acardion.CreateControls
End Sub
```

## Example 2: Styled accordion

### Description
Creating an accordion with customized color scheme and header height.

### Code
```vba
Sub CreateStyledAccordion()
    Dim acardion As New clsAcardion
    acardion.SetParentForm Me
    
    ' Appearance settings
    acardion.HeaderHeight = 35
    acardion.SetStyle RGB(65, 105, 225), RGB(248, 248, 255) ' Blue headers, almost white content
    acardion.SetAnimation 6 ' Medium animation speed
    
    ' Add items
    acardion.AddItem "Interface Settings", "Application appearance parameters"
    acardion.AddItem "Security Settings", "Data protection parameters"
    acardion.AddItem "Performance Settings", "Work optimization parameters"
    
    acardion.CreateControls
End Sub
```

## Example 3: Dynamic item addition

### Description
Example of adding accordion items during program execution based on data from an array.

### Code
```vba
Sub CreateDynamicAccordion()
    Dim acardion As New clsAcardion
    acardion.SetParentForm Me
    
    ' Array with data for the accordion
    Dim headers(1 To 3) As String
    Dim contents(1 To 3) As String
    
    headers(1) = "Item 1"
    contents(1) = "Content of first item"
    headers(2) = "Item 2"
    contents(2) = "Content of second item"
    headers(3) = "Item 3"
    contents(3) = "Content of third item"
    
    ' Add items from array
    Dim i As Integer
    For i = 1 To 3
        acardion.AddItem headers(i), contents(i)
    Next i
    
    acardion.CreateControls
End Sub
```

## Example 4: Accordion with event handling

### Description
Creating an accordion with handling events for item opening and closing.

### Code
```vba
Sub CreateEventHandlingAccordion()
    Dim acardion As New clsAcardion
    acardion.SetParentForm Me
    
    ' Add items
    acardion.AddItem "User Data", "Information about current user"
    acardion.AddItem "Profile Settings", "Profile configuration parameters"
    acardion.AddItem "Action History", "User action log"
    
    ' Create controls
    acardion.CreateControls
    
    ' Event handling (pseudocode - requires additional setup in actual implementation)
    ' When item expands
    ' Call acardion.OnItemExpanded(AddressOf HandleItemExpanded)
    ' When item collapses
    ' Call acardion.OnItemCollapsed(AddressOf HandleItemCollapsed)
End Sub

' Event handling subroutines
Sub HandleItemExpanded(itemIndex As Integer)
    Debug.Print "Item " & itemIndex & " expanded"
End Sub

Sub HandleItemCollapsed(itemIndex As Integer)
    Debug.Print "Item " & itemIndex & " collapsed"
End Sub
```

## Example 5: Accordion with nested elements

### Description
Creating a multi-level accordion with the ability to nest elements within each other.

### Code
```vba
Sub CreateNestedAccordion()
    Dim mainAcardion As New clsAcardion
    mainAcardion.SetParentForm Me
    
    ' Create main accordion
    mainAcardion.AddItem "Category 1", ""
    mainAcardion.AddItem "Category 2", ""
    mainAcardion.AddItem "Category 3", ""
    
    ' Create nested accordion for the second category
    Dim nestedAcardion As New clsAcardion
    nestedAcardion.SetParentForm Me
    nestedAcardion.HeaderHeight = 25 ' Smaller height for nested items
    
    nestedAcardion.AddItem "Subcategory 2.1", "Detailed information about subcategory 2.1"
    nestedAcardion.AddItem "Subcategory 2.2", "Detailed information about subcategory 2.2"
    nestedAcardion.AddItem "Subcategory 2.3", "Detailed information about subcategory 2.3"
    
    ' Insert nested accordion into the content of the second category
    ' This requires additional implementation in the class
    ' mainAcardion.Items(1).SetContentControl nestedAcardion
    
    mainAcardion.CreateControls
End Sub
```

## Example 6: Accordion with action buttons

### Description
Creating an accordion where each item contains buttons for performing actions.

### Code
```vba
Sub CreateActionAccordion()
    Dim acardion As New clsAcardion
    acardion.SetParentForm Me
    
    ' Add items with content including buttons
    acardion.AddItem "Backup", "Create data backup" & vbCrLf & "Button: [Create Backup]"
    acardion.AddItem "Data Cleanup", "Clean temporary files" & vbCrLf & "Button: [Clean]"
    acardion.AddItem "Export Report", "Export report to Excel" & vbCrLf & "Button: [Export]"
    
    acardion.CreateControls
    
    ' Add handlers for buttons (requires additional implementation)
    ' This may require class modification to support button embedding
End Sub
```

## Example 7: Accordion with state preservation

### Description
Creating an accordion that saves and restores state (which items were open/closed) between sessions.

### Code
```vba
Sub CreateStatePreservingAccordion()
    Dim acardion As New clsAcardion
    acardion.SetParentForm Me
    
    ' Add items
    acardion.AddItem "Connection Settings", "Database connection parameters"
    acardion.AddItem "Report Settings", "Report generation parameters"
    acardion.AddItem "Interface Settings", "Application appearance parameters"
    
    ' Restore state from saved data (pseudocode)
    ' Dim savedStates As Variant
    ' savedStates = GetSavedStates() ' Function to get saved state
    
    ' Apply saved state
    ' If IsArray(savedStates) Then
    '     Dim i As Integer
    '     For i = 0 To acardion.Items.Count - 1
    '         If i < UBound(savedStates) + 1 Then
    '             If savedStates(i) = True Then
    '                 acardion.Items(i).Expand
    '             Else
    '                 acardion.Items(i).Collapse
    '             End If
    '         End If
    '     Next i
    ' End If
    
    acardion.CreateControls
End Sub

' Subroutine to save state
Sub SaveAccordionState(acardion As clsAcardion)
    ' Save state of each item
    Dim states() As Boolean
    ReDim states(0 To acardion.Items.Count - 1)
    
    Dim i As Integer
    For i = 0 To acardion.Items.Count - 1
        states(i) = acardion.Items(i).Expanded
    Next i
    
    ' Save array to persistent storage (e.g., application settings)
    ' SaveStatesToStorage states
End Sub
```

## Conclusion

These examples demonstrate various ways to use the `clsAcardion` class in VBA projects. You can adapt and combine these examples depending on your specific requirements and use cases.