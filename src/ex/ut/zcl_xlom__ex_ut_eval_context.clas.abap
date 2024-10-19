class ZCL_XLOM__EX_UT_EVAL_CONTEXT definition
  public
  final
  create private .

public section.

  types:
    BEGIN OF ts_containing_cell,
        row    TYPE i,
        column TYPE i,
      END OF ts_containing_cell .

  data WORKSHEET type ref to ZCL_XLOM_WORKSHEET read-only .
  data CONTAINING_CELL type TS_CONTAINING_CELL read-only .

  class-methods CREATE
    importing
      !WORKSHEET type ref to ZCL_XLOM_WORKSHEET
      !CONTAINING_CELL type TS_CONTAINING_CELL
    returning
      value(RESULT) type ref to ZCL_XLOM__EX_UT_EVAL_CONTEXT .
  methods SET_CONTAINING_CELL
    importing
      !VALUE type TS_CONTAINING_CELL .
protected section.
private section.
ENDCLASS.



CLASS ZCL_XLOM__EX_UT_EVAL_CONTEXT IMPLEMENTATION.


  method CREATE.

    result = NEW zcl_xlom__ex_ut_eval_context( ).
    result->worksheet       = worksheet.
    result->containing_cell = containing_cell.

  endmethod.


  method SET_CONTAINING_CELL.

    me->containing_cell = value.

  endmethod.
ENDCLASS.
