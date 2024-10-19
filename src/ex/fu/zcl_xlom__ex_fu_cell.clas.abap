"! CELL(info_type, [reference])
"! CELL("filename",A1) => https://myoffice.accenture.com/personal/xxxxxxxxxxxxxxxxxxxxx/Documents/[activity log.xlsx]Log
"! In cell B1, formula =CELL("address",A1:A6) is the same result as =CELL("address",A1), which is $A$1 in cell B1;
"! the cells B2 to B6 are not initialized with $A$2, $A$3, etc.
"! https://support.microsoft.com/en-us/office/cell-function-51bd39a5-f338-4dbe-a33f-955d67c2b2cf
CLASS zcl_xlom__ex_fu_cell DEFINITION
  PUBLIC FINAL
  CREATE PRIVATE.

  PUBLIC SECTION.
    INTERFACES zif_xlom__ut_all_friends.
    INTERFACES zif_xlom__ex.

    "! @parameter reference | The parameter is technically OPTIONAL but should be passed, as explained at
    "!                        https://support.microsoft.com/en-us/office/cell-function-51bd39a5-f338-4dbe-a33f-955d67c2b2cf:
    "!                        "<em>Important: Although technically reference is <strong>optional</strong>, including it in your formula is encouraged,
    "!                        unless you understand the effect its absence has on your formula result and want that effect in place.
    "!                        Omitting the reference argument does not reliably produce information about a specific cell, for the following reasons:</em>"
    "!                        <ul>
    "!                        <li>"<em>In automatic calculation mode, when a cell is modified by a user the calculation may be triggered
    "!                            before or after the selection has progressed, depending on the platform you're using for Excel.
    "!                            For example, Excel for Windows currently triggers calculation before selection changes, but Excel
    "!                            for the web triggers it afterward.</em></li>
    "!                        <li>"<em>When Co-Authoring with another user who makes an edit, this function will report your active cell rather than the editor's.</em>"</li>
    "!                        <li>"<em>Any recalculation, for instance pressing F9, will cause the function to return a new result even though no cell edit has occurred.</em>"</li>
    "!                        </ul>
    CLASS-METHODS create
      IMPORTING info_type     TYPE REF TO zcl_xlom__ex_el_string
                !reference    TYPE REF TO zcl_xlom__ex_el_range OPTIONAL
      RETURNING VALUE(result) TYPE REF TO zcl_xlom__ex_fu_cell.

  PRIVATE SECTION.
    TYPES ty_info_type TYPE i.

    CONSTANTS:
      BEGIN OF c_info_type,
        filename TYPE string VALUE 'filename',
      END OF c_info_type.

    DATA info_type TYPE REF TO zcl_xlom__ex_el_string.
    DATA reference TYPE REF TO zcl_xlom__ex_el_range.
ENDCLASS.


CLASS zcl_xlom__ex_fu_cell IMPLEMENTATION.
  METHOD create.
    result = NEW zcl_xlom__ex_fu_cell( ).
    result->zif_xlom__ex~type = zif_xlom__ex=>c_type-function-cell.
    result->info_type         = info_type.
    result->reference         = reference.
    IF reference IS NOT BOUND.
      RAISE EXCEPTION TYPE zcx_xlom_todo.
    ENDIF.
  ENDMETHOD.

  METHOD zif_xlom__ex~evaluate.
    DATA temp_result TYPE REF TO zcl_xlom__va_string.

    TRY.
        " In cell B1, formula =CELL("address",A1:A6) is the same result as =CELL("address",A1), which is $A$1 in cell B1;
        " the cells B2 to B6 are not initialized with $A$2, $A$3, etc.
        DATA(info_type_result) = zcl_xlom__va=>to_string( info_type->zif_xlom__ex~evaluate( context ) )->get_string( ).
        CASE info_type_result.
          WHEN c_info_type-filename.
            " Retourne par exemple "C:\temp\[Book1.xlsx]Sheet1"
            temp_result = zcl_xlom__va_string=>create( context->worksheet->parent->path
                                                        && |\\[{ context->worksheet->parent->name }]|
                                                        && context->worksheet->name ).
          WHEN OTHERS.
            RAISE EXCEPTION TYPE zcx_xlom_todo.
        ENDCASE.
        result = zif_xlom__ex~set_result( temp_result ).
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
