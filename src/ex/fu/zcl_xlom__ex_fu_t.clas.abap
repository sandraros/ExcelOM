"! T(value)
"! If value is or refers to text, T returns value. If value does not refer to text, T returns "" (empty text).
"! Examples: T("text") = "text", T(1) = "", T({1}) = "", T(FALSE) = "", T(XFD1024000) = "" (empty). But T(#N/A) = #N/A.
"! https://support.microsoft.com/en-us/office/t-function-fb83aeec-45e7-4924-af95-53e073541228
CLASS zcl_xlom__ex_fu_t DEFINITION
  PUBLIC FINAL
  CREATE PRIVATE.

  PUBLIC SECTION.
    INTERFACES zif_xlom__ex.

    CLASS-METHODS create
      IMPORTING !value        TYPE REF TO zif_xlom__ex
      RETURNING VALUE(result) TYPE REF TO zcl_xlom__ex_fu_t.

  PRIVATE SECTION.
    DATA value TYPE REF TO zif_xlom__ex.
ENDCLASS.


CLASS zcl_xlom__ex_fu_t IMPLEMENTATION.
  METHOD create.
    result = NEW zcl_xlom__ex_fu_t( ).
    result->zif_xlom__ex~type = zif_xlom__ex=>c_type-function-t.
    result->value             = value.
  ENDMETHOD.

  METHOD zif_xlom__ex~evaluate.
    DATA(array_evaluation) = zcl_xlom__ex_ut_eval=>evaluate_array_operands(
                                 expression = me
                                 context    = context
                                 operands   = VALUE #( ( name = 'VALUE' object = value ) ) ).
    IF array_evaluation-result IS BOUND.
      result = array_evaluation-result.
    ELSE.
      result = zif_xlom__ex~evaluate_single( arguments = array_evaluation-operand_results
                                             context   = context ).
    ENDIF.
  ENDMETHOD.

  METHOD zif_xlom__ex~evaluate_single.
    TRY.
        DATA(value_result) = arguments[ name = 'VALUE' ]-object.
        result = zif_xlom__ex~set_result(
            zcl_xlom__va_string=>create( COND #( WHEN value_result->type = value_result->c_type-string
                                                 THEN zcl_xlom__va=>to_string( value_result )->get_string( )
                                                 ELSE '' ) ) ).
      CATCH zcx_xlom__va INTO DATA(error).
        result = error->result_error.
    ENDTRY.
  ENDMETHOD.

  METHOD zif_xlom__ex~is_equal.
    RAISE EXCEPTION TYPE zcx_xlom_todo.
  ENDMETHOD.

  METHOD zif_xlom__ex~set_result.
    zif_xlom__ex~result_of_evaluation = value.
    result = value.
  ENDMETHOD.
ENDCLASS.
