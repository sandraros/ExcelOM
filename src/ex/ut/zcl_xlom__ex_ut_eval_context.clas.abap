CLASS zcl_xlom__ex_ut_eval_context DEFINITION
  PUBLIC FINAL
  CREATE PRIVATE.

  PUBLIC SECTION.
    TYPES:
      BEGIN OF ts_containing_cell,
        row    TYPE i,
        column TYPE i,
      END OF ts_containing_cell.

    DATA worksheet       TYPE REF TO zcl_xlom_worksheet READ-ONLY.
    DATA containing_cell TYPE ts_containing_cell        READ-ONLY.

    CLASS-METHODS create
      IMPORTING worksheet       TYPE REF TO zcl_xlom_worksheet
                containing_cell TYPE ts_containing_cell
      RETURNING VALUE(result)   TYPE REF TO zcl_xlom__ex_ut_eval_context.

    METHODS set_containing_cell
      IMPORTING !value TYPE ts_containing_cell.
ENDCLASS.


CLASS zcl_xlom__ex_ut_eval_context IMPLEMENTATION.
  METHOD create.
    result = NEW zcl_xlom__ex_ut_eval_context( ).
    result->worksheet       = worksheet.
    result->containing_cell = containing_cell.
  ENDMETHOD.

  METHOD set_containing_cell.
    containing_cell = value.
  ENDMETHOD.
ENDCLASS.
