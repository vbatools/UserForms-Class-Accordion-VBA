# Technical Documentation for VBA Acordion Class

## Overview

The `clsAcardion` class implements accordion functionality in VBA. It allows creating interactive controls that can expand and collapse. The class consists of the main `clsAcardion` class and an auxiliary `clsAcardionItem` class representing a single accordion item.

## Class Structure

### clsAcardion
- `ParentForm` - reference to the parent form
- `Items` - collection of accordion items
- `HeaderHeight` - header height
- `AnimationSpeed` - animation speed
- `HeaderColor` - header color
- `ContentColor` - content color

### clsAcardionItem
- `Header` - header text
- `Content` - item content
- `Expanded` - state (expanded/collapsed)
- `HeaderControl` - header control element
- `ContentControl` - content control element

## Methods

### clsAcardion
- `AddItem(Header As String, Content As String)` - adds a new item
- `RemoveItem(Index As Integer)` - removes item by index
- `ClearItems()` - clears all items
- `CreateControls()` - creates controls on the form
- `ExpandAll()` - expands all items
- `CollapseAll()` - collapses all items
- `SetStyle(HeaderColor As Long, ContentColor As Long)` - sets styles
- `SetAnimation(Speed As Integer)` - sets animation speed

### clsAcardionItem
- `SetHeader(Header As String)` - sets header text
- `SetContent(Content As String)` - sets content
- `Expand()` - expands the item
- `Collapse()` - collapses the item
- `Toggle()` - toggles item state
- `UpdateLayout()` - updates item layout

## Usage

1. Create an instance of `clsAcardion` class
2. Set parent form using `SetParentForm`
3. Add items using `AddItem`
4. Call `CreateControls` to create controls on the form
5. Use `ExpandAll`, `CollapseAll` methods to control state

## Code Example

```vba
Dim acardion As New clsAcardion
acardion.SetParentForm Me
acardion.AddItem "Item 1", "Content of first item"
acardion.AddItem "Item 2", "Content of second item"
acardion.CreateControls
```

## Events

The class supports events:
- `OnItemExpanded` - when an item is expanded
- `OnItemCollapsed` - when an item is collapsed
- `OnAllExpanded` - when all items are expanded
- `OnAllCollapsed` - when all items are collapsed