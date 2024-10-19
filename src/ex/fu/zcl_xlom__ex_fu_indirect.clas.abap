"! INDIRECT(ref_text, [a1])
"! https://support.microsoft.com/en-us/office/indirect-function-474b3a3a-8a26-4f44-b491-92b6306fa261
CLASS zcl_xlom__ex_fu_indirect DEFINITION
  PUBLIC FINAL
  CREATE PRIVATE.

  PUBLIC SECTION.
    INTERFACES zif_xlom__ex.

    "! @parameter ref_text | Range address
    "! @parameter a1 | Optional. A logical value that specifies what type of reference is contained in the cell ref_text.
    "!                 <ul>
    "!                 <li>If a1 is TRUE or omitted, ref_text is interpreted as an A1-style reference.</li>
    "!                 <li>If a1 is FALSE, ref_text is interpreted as an R1C1-style reference.</li>
    "!                 </ul>
    CLASS-METHODS create
      IMPORTING ref_text      TYPE REF TO zif_xlom__ex
                a1            TYPE REF TO zif_xlom__ex OPTIONAL
      RETURNING VALUE(result) TYPE REF TO zcl_xlom__ex_fu_indirect.

  PRIVATE SECTION.
    DATA ref_text TYPE REF TO zif_xlom__ex.
    DATA a1       TYPE REF TO zif_xlom__ex.
ENDCLASS.


CLASS zcl_xlom__ex_fu_indirect IMPLEMENTATION.
  METHOD create.
    result = NEW zcl_xlom__ex_fu_indirect( ).
    result->zif_xlom__ex~type = zif_xlom__ex=>c_type-function-indirect.
    result->ref_text          = ref_text.
    result->a1                = a1.
  ENDMETHOD.

  METHOD zif_xlom__ex~evaluate.
    TRY.
        " INDIRECT("A1:D4")
        DATA(ref_text_result) = zcl_xlom__va=>to_string( ref_text->evaluate( context ) )->get_string( ).
        result = zif_xlom__ex~set_result( zcl_xlom_range=>create_from_address_or_name(
                                              address     = ref_text_result
                                              relative_to = context->worksheet ) ).
      CATCH zcx_xlom__va INTO DATA(error).
        result = error->result_error.
    ENDTRY.
  ENDMETHOD.

  METHOD zif_xlom__ex~evaluate_single.
    RAISE EXCEPTION TYPE zcx_xlom_unexpected.
  ENDMETHOD.

  METHOD zif_xlom__ex~is_equal.
    RAISE EXCEPTION TYPE zcx_xlom_todo.
  ENDMETHOD.

  METHOD zif_xlom__ex~set_result.
    zif_xlom__ex~result_of_evaluation = value.
    result = value.
  ENDMETHOD.
ENDCLASS.
