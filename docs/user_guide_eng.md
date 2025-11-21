# User Guide for VBA Acordion Class

## Introduction

The `clsAcardion` class provides a convenient way to create accordions (collapsible panels) in VBA applications. This component allows users to control the display of information, enabling them to expand and collapse content as needed.

## Installation and Setup

1. Import the class files `clsAcardion.cls` and `clsAcardionItem.cls` into your VBA project
2. If needed, add the module `modShowForms.bas` and form `frmTestClass.frm` for testing

## Basic Usage

### Creating an Accordion

To create a basic accordion, follow these steps:

```vba
Dim acardion As New clsAcardion
acardion.SetParentForm Me ' Set the parent form
acardion.AddItem "Header 1", "Content of first item"
acardion.AddItem "Header 2", "Content of second item"
acardion.CreateControls ' Create the controls
```

### Adding Items

Use the `AddItem` method to add new items:

```vba
acardion.AddItem "New Header", "New Content"
```

### Managing State

You can control the state of items:

```vba
acardion.ExpandAll ' Expand all items
acardion.CollapseAll ' Collapse all items
```

## Advanced Features

### Styling

The class allows customization of appearance:

```vba
acardion.SetStyle RGB(20, 200, 200), RGB(255, 25, 255) ' Header and content colors
```

### Animation Settings

Control animation speed:

```vba
acardion.SetAnimation 5 ' Set animation speed (1-10)
```

### Working with Individual Items

Access individual items:

```vba
Dim item As clsAcardionItem
Set item = acardion.Items(0) ' Get the first item
item.Expand ' Expand a specific item
```

## Practical Examples

### Example 1: Simple Accordion

```vba
Sub CreateSimpleAccordion()
    Dim acardion As New clsAcardion
    acardion.SetParentForm Me
    
    acardion.AddItem "Description", "This is the description section"
    acardion.AddItem "Settings", "This is the settings section"
    acardion.AddItem "Help", "This is the help information section"
    
    acardion.CreateControls
End Sub
```

### Example 2: Styled Accordion

```vba
Sub CreateStyledAccordion()
    Dim acardion As New clsAcardion
    acardion.SetParentForm Me
    
    ' Add items
    acardion.AddItem "Item 1", "Content of first item"
    acardion.AddItem "Item 2", "Content of second item"
    
    ' Style settings
    acardion.SetStyle RGB(70, 130, 180), RGB(240, 248, 255) ' Header and content styles
    
    ' Animation settings
    acardion.SetAnimation 7
    
    acardion.CreateControls
End Sub
```

## Frequently Asked Questions

### How do I change header height?

You can change the header height by setting the `HeaderHeight` property:

```vba
acardion.HeaderHeight = 30 ' Set header height to 30 pixels
```

### Can I add items after creating controls?

Yes, you can add items at any time, but after adding new items you need to call `CreateControls` again or use methods to update the layout.

### How do I handle events when items open/close?

The class supports events that can be used to handle user actions. Implement the appropriate event handlers in your code.