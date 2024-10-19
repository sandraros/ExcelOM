class ZCL_XLOM__EX_UT_EVAL definition
  public
  create public .

public section.

  class-methods EVALUATE_ARRAY_OPERANDS
    importing
      !EXPRESSION type ref to ZIF_XLOM__EX
      !CONTEXT type ref to ZCL_XLOM__EX_UT_EVAL_CONTEXT
      !OPERANDS type ZIF_XLOM__EX=>TT_OPERAND_EXPR
    returning
      value(RESULT) type ZIF_XLOM__EX=>TS_EVALUATE_ARRAY_OPERANDS .
  PROTECTED SECTION.

  PRIVATE SECTION.
ENDCLASS.



CLASS ZCL_XLOM__EX_UT_EVAL IMPLEMENTATION.


  METHOD evaluate_array_operands.
    " One simple example to understand:
    "
    " In E1, INDEX(A1:D4,{1,2;3,4},{1,2;3,4}) will correspond to four values in E1:F2:
    " E1: INDEX(A1:D4,1,1) (i.e. A1)
    " F1: INDEX(A1:D4,2,2) (i.e. B2)
    " E2: INDEX(A1:D4,3,3) (i.e. C3)
    " F2: INDEX(A1:D4,4,4) (i.e. D4)
    "
    " Below is the generalization.
    "
    "  formula with operation on arrays        result
    "
    "  -> if the left operand is one line high, the line is replicated till max lines of the right operand.
    "  -> if the left operand is one column wide, the column is replicated till max columns of the right operand.
    "  -> if the right operand is one line high, the line is replicated till max lines of the left operand.
    "  -> if the right operand is one column wide, the column is replicated till max columns of the left operand.
    "
    "  -> if the left operand has less lines than the right operand, additional lines are added with #N/A.
    "  -> if the left operand has less columns than the right operand, additional columns are added with #N/A.
    "  -> if the right operand has less lines than the left operand, additional lines are added with #N/A.
    "  -> if the right operand has less columns than the left operand, additional columns are added with #N/A.
    "
    "  -> target array size = max lines of both operands + max columns of both operands.
    "  -> each target cell of the target array is calculated like this:
    "     T(1,1) = L(1,1) op R(1,1)
    "     T(2,1) = L(2,1) op R(2,1)
    "     etc.
    "     If the left cell or right cell is #N/A, the target cell is also #N/A.
    "
    "  Examples where one of the two operands has 1 cell, 1 line or 1 column
    "
    "  a | b | c   op   k | l | m | n          a op k | b op l | c op m | #N/A
    "
    "  a | b | c   op   k                      a op k | b op k | c op k
    "  d | e | f                               d op k | e op k | f op k
    "  g | h | i                               g op k | h op k | i op k
    "
    "  a | b | c   op   k | l | m | n          a op k | b op l | c op m | #N/A
    "  d | e | f                               d op k | e op l | f op m | #N/A
    "  g | h | i                               g op k | h op l | i op m | #N/A
    "
    "  a | b | c   op   k                      a op k | b op k | c op k
    "  d | e | f        l                      d op l | e op l | f op l
    "  g | h | i        m                      g op m | h op m | i op m
    "                   n                      #N/A   | #N/A   | #N/A
    "
    "  a | b | c   op   k                      a op k | b op k | c op k
    "  d | e | f        l                      d op l | e op l | f op l
    "  g | h | i                               #N/A   | #N/A   | #N/A
    "
    "  a | b | c   op   k                      a op k | b op k | c op k
    "                   l                      a op l | b op l | c op l
    "                   m                      a op m | b op m | c op m
    "
    "  Both operands have more than 1 line and more than 1 column
    "
    "  a | b | c   op   k | n                  a op k | b op n | #N/A
    "  d | e | f        l | o                  d op l | e op o | #N/A
    "  g | h | i                               #N/A   | #N/A   | #N/A
    "
    "  a | b | c   op   k | n                  a op k | b op n | #N/A
    "  d | e | f        l | o                  d op l | e op o | #N/A
    "                   m | p                  #N/A   | #N/A   | #N/A
    "
    "  a | b       op   k | n | q              a op k | b op n | #N/A
    "  d | e            l | o | r              d op l | e op o | #N/A
    "  g | h                                   #N/A   | #N/A   | #N/A
    DATA(at_least_one_array_or_range) = abap_false.
    LOOP AT operands REFERENCE INTO DATA(operand).
      IF operand->object IS NOT BOUND.
        " e.g. NUM_CHARS not passed to function RIGHT
        INSERT VALUE #( name                     = operand->name
                        object                   = VALUE #( )
                        not_part_of_result_array = operand->not_part_of_result_array )
               INTO TABLE result-operand_results.
      ELSE.
        " EVALUATE THE OPERAND
        INSERT VALUE #( name                     = operand->name
                        object                   = operand->object->evaluate( context )
                        not_part_of_result_array = operand->not_part_of_result_array )
               INTO TABLE result-operand_results
               REFERENCE INTO DATA(operand_result).
        " Should we perform array evaluation on more than 1 cell?
        IF     operand_result->not_part_of_result_array = abap_false
           AND (    operand_result->object->type = operand_result->object->c_type-array
                 OR operand_result->object->type = operand_result->object->c_type-range )
           AND (    CAST zif_xlom__va_array( operand_result->object )->row_count    > 1
                 OR CAST zif_xlom__va_array( operand_result->object )->column_count > 1 ).
          at_least_one_array_or_range = abap_true.
        ENDIF.
      ENDIF.
    ENDLOOP.

    " Continue only if there's at least one array or one range with more than one cell.
    IF at_least_one_array_or_range = abap_false.
      RETURN.
    ENDIF.

    DATA(max_row_count) = 1.
    DATA(max_column_count) = 1.
    LOOP AT result-operand_results REFERENCE INTO operand_result
         WHERE     not_part_of_result_array  = abap_false
               AND object                   IS BOUND.
      CASE operand_result->object->type.
        WHEN operand_result->object->c_type-array
          OR operand_result->object->c_type-range.
          max_row_count = nmax( val1 = max_row_count
                                val2 = CAST zif_xlom__va_array( operand_result->object )->row_count ).
          max_column_count = nmax( val1 = max_column_count
                                   val2 = CAST zif_xlom__va_array( operand_result->object )->column_count ).
      ENDCASE.
    ENDLOOP.

    DATA(target_array) = zcl_xlom__va_array=>create_initial( row_count    = max_row_count
                                                             column_count = max_column_count ).
    DATA(row) = 1.
    DO max_row_count TIMES.

      DATA(column) = 1.
      DO max_column_count TIMES.

        DATA(single_cell_operands) = VALUE zif_xlom__ex=>tt_operand_result( ).
        LOOP AT result-operand_results REFERENCE INTO operand_result.
          IF     operand_result->not_part_of_result_array  = abap_false
             AND operand_result->object                   IS BOUND.
            IF operand_result->object->type = operand_result->object->c_type-array.
              DATA(operand_result_array) = CAST zcl_xlom__va_array( operand_result->object ).
              DATA(cell) = operand_result_array->zif_xlom__va_array~get_cell_value( column = column
                                                                                    row    = row ).
            ELSEIF operand_result->object->type = operand_result->object->c_type-range.
              DATA(operand_result_range) = CAST zcl_xlom_range( operand_result->object ).
              cell = operand_result_range->cells( row    = row
                                                  column = column ).
            ELSE.
              cell = operand_result->object.
            ENDIF.
          ELSE.
            cell = operand_result->object.
          ENDIF.
          INSERT VALUE #( name   = operand_result->name
                          object = cell )
                 INTO TABLE single_cell_operands.
        ENDLOOP.
        DATA(single_cell_result) = expression->evaluate_single( arguments = single_cell_operands
                                                                context   = context ).
        target_array->zif_xlom__va_array~set_cell_value( row    = row
                                                         column = column
                                                         value  = single_cell_result ).

        column = column + 1.
      ENDDO.

      row = row + 1.
    ENDDO.
    result-result = target_array.
  ENDMETHOD.
ENDCLASS.
