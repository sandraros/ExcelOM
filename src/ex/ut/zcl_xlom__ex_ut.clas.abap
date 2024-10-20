CLASS zcl_xlom__ex_ut DEFINITION
  PUBLIC
  CREATE PUBLIC.

  PUBLIC SECTION.
    CLASS-METHODS are_equal
      IMPORTING expression_1  TYPE REF TO zif_xlom__ex
                expression_2  TYPE REF TO zif_xlom__ex
      RETURNING VALUE(result) TYPE abap_bool.

    CLASS-METHODS check_arguments_or_operands
      IMPORTING expression            TYPE REF TO zif_xlom__ex
      CHANGING  arguments_or_operands TYPE zif_xlom__ex=>tt_argument_or_operand.

  PROTECTED SECTION.

  PRIVATE SECTION.
ENDCLASS.


CLASS zcl_xlom__ex_ut IMPLEMENTATION.
  METHOD are_equal.
    IF    (     expression_1 IS NOT BOUND
            AND expression_2 IS BOUND )
       OR (     expression_1 IS BOUND
            AND expression_2 IS NOT BOUND ).
      RETURN.
    ELSEIF     expression_1 IS BOUND
           AND expression_2 IS BOUND.
      IF expression_1->type <> expression_2->type.
        RETURN.
      ENDIF.
      IF lines( expression_1->arguments_or_operands ) <> lines( expression_2->arguments_or_operands ).
        RETURN.
      ENDIF.
      LOOP AT expression_1->arguments_or_operands INTO DATA(expression_1_argument_or_opera).
        DATA(expression_2_argument_or_opera) = VALUE #( expression_2->arguments_or_operands[ sy-tabix ] ).
        IF expression_1_argument_or_opera->type <> expression_2_argument_or_opera->type.
          RETURN.
        ENDIF.
        IF NOT are_equal( expression_1 = expression_1_argument_or_opera
                          expression_2 = expression_2_argument_or_opera ).
          RETURN.
        ENDIF.
      ENDLOOP.
    ENDIF.
    result = abap_true.
  ENDMETHOD.

  METHOD check_arguments_or_operands.
    DATA(parameter_number) = 0.
    LOOP AT expression->parameters REFERENCE INTO DATA(parameter).
      parameter_number = parameter_number + 1.
      DATA(argument_or_operand) = REF #( arguments_or_operands[ parameter_number ] OPTIONAL ).
      IF     argument_or_operand IS NOT BOUND
         AND parameter->default  IS NOT BOUND.
        " Mandatory parameter not passed
        RAISE EXCEPTION TYPE zcx_xlom_todo.
      ENDIF.
      IF parameter->default IS BOUND.
        IF argument_or_operand IS NOT BOUND.
          INSERT parameter->default INTO arguments_or_operands INDEX parameter_number.
        ELSEIF argument_or_operand->* IS NOT BOUND.
          argument_or_operand->* = parameter->default.
        ENDIF.
      ENDIF.
    ENDLOOP.
  ENDMETHOD.
ENDCLASS.
