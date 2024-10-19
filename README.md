# XLOM

ABAP library like the Excel Object Model, which can be used to evaluate Excel formulas in ABAP. The formulas must be expressed in English.

This first version supports very few items (objects, methods, properties, functions and operators).

Very simple example to demonstrate the formula to calculate `25 + 10 = 35`:
```
DATA(xlom_application) = zcl_xlom_application=>create( ).

DATA(xlom_workbook) = xlom_application->workbooks->add( ).

DATA(xlom_worksheet) = xlom_workbook->worksheets->add( 'Sheet1' ).

xlom_worksheet->range( cell1_string = 'A1' )->set_value( zcl_xlom__va_number=>get( 25 ) ).

xlom_worksheet->range( cell1_string = 'B1' )->set_formula2( 'A1+10' ).

ASSERT 35 = CAST zcl_xlom__va_number( xlom_worksheet->range( cell1_string = 'B1' )->value( ) )->get_number( ).
```

# Currently-supported items

## Objects and methods
- Application
  - Methods
    - Calculate
    - Intersect
  - Properties
    - ActiveSheet
    - Calculation
    - Workbooks
- Columns
  - Properties
    - Count
- Range
  - Methods
    - Address
    - Calculate
    - Offset
    - Resize
  - Properties
    - Application
    - Cells
    - Columns
    - Count
    - Formula2
    - Parent
    - Row
    - Rows
    - Value
- Rows
  - Properties
    - Count
- Sheet
- Workbook
  - Methods
    - SaveAs
  - Properties
    - Application
    - Name
    - Path
    - Worksheets
- Workbooks
  - Methods
    - Add
    - Item
  - Properties
    - Application
    - Count
- Worksheet
  - Methods
    - Calculate
    - Cells
    - Range
  - Properties
    - Application
    - Name
    - Parent
    - UsedRange
- Worksheets
  - Methods
    - Add
    - Item
  - Properties
    - Application
    - Count

## Functions
- ADDRESS
- CELL
- COUNTIF
- FIND
- IF
- IFERROR
- INDEX
- INDIRECT
- LEN
- MATCH
- OFFSET
- RIGHT
- ROW
- T

## Operators
- `&`
- `:`
- `=`
- `-`
- `-` (unary operator like in OFFSET(B2,-1))
- `*`
- `+`
