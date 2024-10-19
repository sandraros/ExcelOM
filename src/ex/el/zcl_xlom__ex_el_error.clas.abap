CLASS zcl_xlom__ex_el_error DEFINITION
  PUBLIC FINAL
  CREATE PRIVATE.

  PUBLIC SECTION.
    INTERFACES zif_xlom__ex.

    TYPES ty_error_number TYPE i.

    CLASS-DATA blocked                    TYPE REF TO zcl_xlom__ex_el_error READ-ONLY.
    CLASS-DATA calc                       TYPE REF TO zcl_xlom__ex_el_error READ-ONLY.
    CLASS-DATA connect                    TYPE REF TO zcl_xlom__ex_el_error READ-ONLY.
    "! #DIV/0! Is produced by =1/0
    CLASS-DATA division_by_zero           TYPE REF TO zcl_xlom__ex_el_error READ-ONLY.
    CLASS-DATA field                      TYPE REF TO zcl_xlom__ex_el_error READ-ONLY.
    CLASS-DATA getting_data               TYPE REF TO zcl_xlom__ex_el_error READ-ONLY.
    "! #N/A. Is produced by =ERROR.TYPE(1) or if C1 contains =A1:A2+B1:B3 -> C3=#N/A.
    CLASS-DATA na_not_applicable          TYPE REF TO zcl_xlom__ex_el_error READ-ONLY.
    "! #NAME! Is produced by =XXXX if XXXX is not an existing range name.
    CLASS-DATA name                       TYPE REF TO zcl_xlom__ex_el_error READ-ONLY.
    CLASS-DATA null                       TYPE REF TO zcl_xlom__ex_el_error READ-ONLY.
    "! #NUM! Is produced by =1E+240*1E+240
    CLASS-DATA num                        TYPE REF TO zcl_xlom__ex_el_error READ-ONLY.
    "! TODO #PYTHON! internal error number is not 2222, what is it?
    CLASS-DATA python                     TYPE REF TO zcl_xlom__ex_el_error READ-ONLY.
    "! #REF! Is produced by =INDEX(A1,2,1)
    CLASS-DATA ref                        TYPE REF TO zcl_xlom__ex_el_error READ-ONLY.
    "! #SPILL! Is produced by A1 containing ={1,2} and B1 containing a value -> A1=#SPILL!
    CLASS-DATA spill                      TYPE REF TO zcl_xlom__ex_el_error READ-ONLY.
    CLASS-DATA unknown                    TYPE REF TO zcl_xlom__ex_el_error READ-ONLY.
    "! #VALUE! Is produced by =1+"a". #VALUE! in English, #VALEUR! in French.
    CLASS-DATA value_cannot_be_calculated TYPE REF TO zcl_xlom__ex_el_error READ-ONLY.

    "! English error name (different in other languages e.g. #VALUE! is #VALEUR! in French)
    DATA english_error_name    TYPE string          READ-ONLY.
    "! Value of enumeration xlCVError e.g. xlErrValue = 2015 (#VALUE!)
    DATA internal_error_number TYPE ty_error_number READ-ONLY.

    CLASS-METHODS class_constructor.

    CLASS-METHODS get_from_error_name
      IMPORTING error_name    TYPE csequence
      RETURNING VALUE(result) TYPE REF TO zif_xlom__ex.

  PRIVATE SECTION.
    TYPES:
      BEGIN OF ts_error,
        "! English error name (different in other languages e.g. #VALUE! is #VALEUR! in French)
        english_error_name    TYPE string,
        "! Value of enumeration xlCVError e.g. xlErrValue = 2015 (#VALUE!)
        internal_error_number TYPE ty_error_number,
        object                TYPE REF TO zcl_xlom__ex_el_error,
      END OF ts_error.
    TYPES tt_error TYPE STANDARD TABLE OF ts_error WITH EMPTY KEY.

    CLASS-DATA errors TYPE tt_error.

    CLASS-METHODS create
      IMPORTING english_error_name    TYPE ts_error-english_error_name
                internal_error_number TYPE ts_error-internal_error_number
      RETURNING VALUE(result)         TYPE REF TO zcl_xlom__ex_el_error.
ENDCLASS.


CLASS zcl_xlom__ex_el_error IMPLEMENTATION.
  METHOD class_constructor.
    blocked                    = zcl_xlom__ex_el_error=>create( english_error_name    = '#BLOCKED!     '
                                                                internal_error_number = 2047 ).
    calc                       = zcl_xlom__ex_el_error=>create( english_error_name    = '#CALC!        '
                                                                internal_error_number = 2050 ).
    connect                    = zcl_xlom__ex_el_error=>create( english_error_name    = '#CONNECT!     '
                                                                internal_error_number = 2046 ).
    division_by_zero           = zcl_xlom__ex_el_error=>create( english_error_name    = '#DIV/0!       '
                                                                internal_error_number = 2007 ).
    field                      = zcl_xlom__ex_el_error=>create( english_error_name    = '#FIELD!       '
                                                                internal_error_number = 2049 ).
    getting_data               = zcl_xlom__ex_el_error=>create( english_error_name    = '#GETTING_DATA!'
                                                                internal_error_number = 2043 ).
    na_not_applicable          = zcl_xlom__ex_el_error=>create( english_error_name    = '#N/A          '
                                                                internal_error_number = 2042 ).
    name                       = zcl_xlom__ex_el_error=>create( english_error_name    = '#NAME?        '
                                                                internal_error_number = 2029 ).
    null                       = zcl_xlom__ex_el_error=>create( english_error_name    = '#NULL!        '
                                                                internal_error_number = 2000 ).
    num                        = zcl_xlom__ex_el_error=>create( english_error_name    = '#NUM!         '
                                                                internal_error_number = 2036 ).
    python                     = zcl_xlom__ex_el_error=>create( english_error_name    = '#PYTHON!      '
                                                                internal_error_number = 2222 ).
    ref                        = zcl_xlom__ex_el_error=>create( english_error_name    = '#REF!         '
                                                                internal_error_number = 2023 ).
    spill                      = zcl_xlom__ex_el_error=>create( english_error_name    = '#SPILL!       '
                                                                internal_error_number = 2045 ).
    unknown                    = zcl_xlom__ex_el_error=>create( english_error_name    = '#UNKNOWN!     '
                                                                internal_error_number = 2048 ).
    value_cannot_be_calculated = zcl_xlom__ex_el_error=>create( english_error_name    = '#VALUE!       '
                                                                internal_error_number = 2015 ).
  ENDMETHOD.

  METHOD create.
    result = NEW zcl_xlom__ex_el_error( ).
    result->zif_xlom__ex~type     = zif_xlom__ex=>c_type-error.
    result->english_error_name    = english_error_name.
    result->internal_error_number = internal_error_number.
    INSERT VALUE #( english_error_name    = english_error_name
                    internal_error_number = internal_error_number
                    object                = result )
           INTO TABLE errors.
  ENDMETHOD.

  METHOD get_from_error_name.
    result = errors[ english_error_name = error_name ]-object.
  ENDMETHOD.

  METHOD zif_xlom__ex~evaluate.
    result = zif_xlom__ex~set_result( zcl_xlom__va_error=>get_by_error_number( internal_error_number ) ).
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
