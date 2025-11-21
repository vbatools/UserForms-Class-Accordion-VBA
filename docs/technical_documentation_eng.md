# Technical Documentation for VBA Accordion Class

## Overview

The `clsAccordion` class implements accordion functionality in VBA. It allows creating interactive controls that can expand and collapse. The class consists of the main `clsAccordion` class and an auxiliary `clsAccordionItem` class representing a single accordion item.

## Class Structure

### clsAccordion
- `ParentForm` - reference to the parent form
- `Items` - collection of accordion items
- `HeaderHeight` - header height
- `AnimationSpeed` - animation speed
- `HeaderColor` - header color
- `ContentColor` - content color

### clsAccordionItem
- `Header` - header text
- `Content` - item content
- `Expanded` - state (expanded/collapsed)
- `HeaderControl` - header control element
- `ContentControl` - content control element

## Methods

### clsAccordion
- `AddItem(Header As String, Content As String)` - adds a new item
- `RemoveItem(Index As Integer)` - removes item by index
- `ClearItems()` - clears all items
- `CreateControls()` - creates controls on the form
- `ExpandAll()` - expands all items
- `CollapseAll()` - collapses all items
- `SetStyle(HeaderColor As Long, ContentColor As Long)` - sets styles
- `SetAnimation(Speed As Integer)` - sets animation speed

### clsAccordionItem
- `SetHeader(Header As String)` - sets header text
- `SetContent(Content As String)` - sets content
- `Expand()` - expands the item
- `Collapse()` - collapses the item
- `Toggle()` - toggles item state
- `UpdateLayout()` - updates item layout

## Usage

1. Create an instance of `clsAccordion` class
2. Set parent form using `SetParentForm`
3. Add items using `AddItem`
4. Call `CreateControls` to create controls on the form
5. Use `ExpandAll`, `CollapseAll` methods to control state

## Code Example

```vba
Dim accordion As clsAccordion
Set accordion = New clsAccordion
accordion.SetParentForm Me
accordion.AddItem "Item 1", "Content of first item"
accordion.AddItem "Item 2", "Content of second item"
accordion.CreateControls
```

## Events

The class supports events:
- `OnItemExpanded` - when an item is expanded
- `OnItemCollapsed` - when an item is collapsed
- `OnAllExpanded` - when all items are expanded
- `OnAllCollapsed` - when all items are collapsed