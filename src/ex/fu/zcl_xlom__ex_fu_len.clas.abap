"! LEN(text)
"! https://support.microsoft.com/en-us/office/len-lenb-functions-29236f94-cedc-429d-affd-b5e33d2c67cb
CLASS zcl_xlom__ex_fu_len DEFINITION
  PUBLIC FINAL
  CREATE PRIVATE.

  PUBLIC SECTION.
    INTERFACES zif_xlom__ex.

    CLASS-METHODS create
      IMPORTING !text         TYPE REF TO zif_xlom__ex
      RETURNING VALUE(result) TYPE REF TO zcl_xlom__ex_fu_len.

  PRIVATE SECTION.
    DATA text TYPE REF TO zif_xlom__ex.
ENDCLASS.


CLASS zcl_xlom__ex_fu_len IMPLEMENTATION.
  METHOD create.
    result = NEW zcl_xlom__ex_fu_len( ).
    result->zif_xlom__ex~type = zif_xlom__ex=>c_type-function-len.
    result->text              = text.
  ENDMETHOD.

  METHOD zif_xlom__ex~evaluate.
    DATA(array_evaluation) = zcl_xlom__ex_ut_eval=>evaluate_array_operands(
                                 expression = me
                                 context    = context
                                 operands   = VALUE #( ( name = 'TEXT' object = text ) ) ).
    IF array_evaluation-result IS BOUND.
      result = array_evaluation-result.
    ELSE.
      result = zif_xlom__ex~evaluate_single( arguments = array_evaluation-operand_results
                                             context   = context ).
    ENDIF.
  ENDMETHOD.

  METHOD zif_xlom__ex~evaluate_single.
    TRY.
        result = zif_xlom__ex~set_result(
                     zcl_xlom__va_number=>create(
                         strlen( zcl_xlom__va=>to_string( arguments[ name = 'TEXT' ]-object )->get_string( ) ) ) ).
      CATCH zcx_xlom__va INTO DATA(error).
        result = error->result_error.
    ENDTRY.
  ENDMETHOD.

  METHOD zif_xlom__ex~is_equal.
    IF expression->type <> zif_xlom__ex=>c_type-function-len.
      RETURN.
    ENDIF.
    DATA(len) = CAST zcl_xlom__ex_fu_len( expression ).
    result = xsdbool( text->is_equal( len->text ) ).
  ENDMETHOD.

  METHOD zif_xlom__ex~set_result.
    zif_xlom__ex~result_of_evaluation = value.
    result = value.
  ENDMETHOD.
ENDCLASS.
