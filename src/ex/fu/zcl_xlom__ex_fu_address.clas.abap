"! ADDRESS(row_num, column_num, [abs_num], [a1], [sheet_text])
"! https://support.microsoft.com/en-us/office/address-function-d0c26c0d-3991-446b-8de4-ab46431d4f89
CLASS zcl_xlom__ex_fu_address DEFINITION
  PUBLIC FINAL
  CREATE PRIVATE.

  PUBLIC SECTION.
    INTERFACES zif_xlom__ut_all_friends.
    INTERFACES zif_xlom__ex.

    CLASS-METHODS create
      IMPORTING row_num       TYPE REF TO zif_xlom__ex
                column_num    TYPE REF TO zif_xlom__ex
                abs_num       TYPE REF TO zif_xlom__ex OPTIONAL
                a1            TYPE REF TO zif_xlom__ex OPTIONAL
                sheet_text    TYPE REF TO zif_xlom__ex OPTIONAL
      RETURNING VALUE(result) TYPE REF TO zcl_xlom__ex_fu_address.

  PRIVATE SECTION.
    DATA row_num    TYPE REF TO zif_xlom__ex.
    DATA column_num TYPE REF TO zif_xlom__ex.
    DATA abs_num    TYPE REF TO zif_xlom__ex.
    DATA a1         TYPE REF TO zif_xlom__ex.
    DATA sheet_text TYPE REF TO zif_xlom__ex.
ENDCLASS.


CLASS zcl_xlom__ex_fu_address IMPLEMENTATION.
  METHOD create.
    result = NEW zcl_xlom__ex_fu_address( ).
    result->zif_xlom__ex~type = zif_xlom__ex=>c_type-function-address.
    result->row_num           = row_num.
    result->column_num        = column_num.
    result->abs_num           = abs_num.
    result->a1                = a1.
    result->sheet_text        = sheet_text.
  ENDMETHOD.

  METHOD zif_xlom__ex~evaluate.
    RAISE EXCEPTION TYPE zcx_xlom_todo.
  ENDMETHOD.

  METHOD zif_xlom__ex~evaluate_single.
    RAISE EXCEPTION TYPE zcx_xlom_todo.
  ENDMETHOD.

  METHOD zif_xlom__ex~is_equal.
    RAISE EXCEPTION TYPE zcx_xlom_todo.
  ENDMETHOD.

  METHOD zif_xlom__ex~set_result.
    RAISE EXCEPTION TYPE zcx_xlom_todo.
  ENDMETHOD.
ENDCLASS.
