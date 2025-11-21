# VBA Accordion Class

![Project Demo](Project_Image.gif)

This repository contains a VBA implementation of the Accordion class. The class provides creation of interactive accordions with customizable appearance and behavior.

## Contents
1. [Features](#features)
2. [Components](#components)
3. [Installation](#installation)
4. [Quick Start](#quick-start)
5. [Main Functions](#main-functions)
6. [Working with Controls](#working-with-controls)
7. [Style Configuration](#style-configuration)
8. [Troubleshooting](#troubleshooting)

## Features

- Creation of multi-level accordions
- Customization of header and content appearance
- Support for open/close animation
- Ability to programmatically control elements
- Integration with existing forms
- Support for events when opening/closing elements
- Flexible styling system

## Components

- [`vba-files/Class/clsAccordion.cls`](vba-files/Class/clsAccordion.cls): Main class implementation
- [`vba-files/Class/clsAccordionItem.cls`](vba-files/Class/clsAccordionItem.cls): Accordion item implementation
- [`vba-files/Form/frmTestClass.frm`](vba-files/Form/frmTestClass.frm): Test form demonstrating usage
- [`vba-files/Module/modShowForms.bas`](vba-files/Module/modShowForms.bas): Module containing form display functions
- Documentation in the `docs/` folder:
  - [`docs/technical_documentation_rus.md`](docs/technical_documentation_rus.md) - Technical documentation in Russian
  - [`docs/technical_documentation_eng.md`](docs/technical_documentation_eng.md) - Technical documentation in English
  - [`docs/user_guide_rus.md`](docs/user_guide_rus.md) - User guide in Russian
  - [`docs/user_guide_eng.md`](docs/user_guide_eng.md) - User guide in English
  - [`docs/implementation_examples_rus.md`](docs/implementation_examples_rus.md) - Implementation examples in Russian
  - [`docs/implementation_examples_eng.md`](docs/implementation_examples_eng.md) - Implementation examples in English

## Installation

1. Download the `acardion_v3.xlsm` file
2. Open it in Excel
3. In the VBA editor, import the `clsAccordion.cls` and `clsAccordionItem.cls` classes into your project
4. If needed, add the `modShowForms.bas` module and `frmTestClass.frm` form for testing

## Quick Start

To start working with the accordion class, create an instance of the class and configure its parameters:

### Simple Usage Example
```vba
Dim accordion As clsAccordion
Set accordion = New clsAccordion
accordion.SetParentForm Me ' Specify the parent form
accordion.AddItem "Header 1", "Content of first item"
accordion.AddItem "Header 2", "Content of second item"
accordion.CreateControls ' Create the controls
```

## Main Functions

The `clsAccordion` class provides the following main functions:
- `AddItem` - adding a new accordion item
- `RemoveItem` - removing an accordion item
- `ClearItems` - clearing all items
- `ExpandAll` - expand all items
- `CollapseAll` - collapse all items
- `SetStyle` - configure display style
- `SetAnimation` - configure animation parameters

## Working with Controls

Each accordion item is represented by the `clsAccordionItem` class, which allows:
- Control item visibility
- Change header and content
- Appearance customization
- Open/close event handling

## Style Configuration

The class supports style configuration through methods:
- `SetHeaderStyle` - header style
- `SetContentStyle` - content style
- `SetColors` - color scheme
- `SetFont` - font configuration

## Troubleshooting

If you encounter problems when using the class, check the following:
- Make sure all dependencies are imported correctly
- Verify that the parent form is set using `SetParentForm`
- Ensure that controls are created after adding all items

## License

This project is licensed under the Apache License 2.0 - see the [LICENSE](LICENSE) file for details.