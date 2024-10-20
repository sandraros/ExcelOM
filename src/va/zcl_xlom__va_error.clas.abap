"! https://learn.microsoft.com/en-us/office/vba/excel/concepts/cells-and-ranges/cell-error-values
"! NB: many errors are missing, the list of the other errors can be found in xlCVError enumeration.
"! #VALUE! in English, #VALEUR! in French, etc.
"!
"! You can insert a cell error value into a cell or test the value of a cell for an error value by
"! using the CVErr function. The cell error values can be one of the following xlCVError constants.
"! <ul>
"! <li>Constant . .Error number . .Cell error value</li>
"! <li>xlErrDiv0 . 2007 . . . . . .#DIV/0!         </li>
"! <li>xlErrNA . . 2042 . . . . . .#N/A            </li>
"! <li>xlErrName . 2029 . . . . . .#NAME?          </li>
"! <li>xlErrNull . 2000 . . . . . .#NULL!          </li>
"! <li>xlErrNum . .2036 . . . . . .#NUM!           </li>
"! <li>xlErrRef . .2023 . . . . . .#REF!           </li>
"! <li>xlErrValue .2015 . . . . . .#VALUE!         </li>
"! </ul>
"! VB example:
"! <ul>
"! <li>If IsError(ActiveCell.Value) Then            </li>
"! <li>. If ActiveCell.Value = CVErr(xlErrDiv0) Then</li>
"! <li>. End If                                     </li>
"! <li>End If                                       </li>
"! </ul>
"! NB:
"! <ul>
"! <li>CVErr(xlErrDiv0) is of type Variant/Error and Locals/Watches shows: Error 2007</li>
"! <li>There is no Error data type, only Variant can be used.                        </li>
"! </ul>
CLASS zcl_xlom__va_error DEFINITION
  PUBLIC FINAL
  CREATE PRIVATE.

  PUBLIC SECTION.
    INTERFACES zif_xlom__va.

    TYPES ty_error_number TYPE i.

    CLASS-DATA blocked                    TYPE REF TO zcl_xlom__va_error READ-ONLY.
    CLASS-DATA calc                       TYPE REF TO zcl_xlom__va_error READ-ONLY.
    CLASS-DATA connect                    TYPE REF TO zcl_xlom__va_error READ-ONLY.
    "! #DIV/0! Is produced by =1/0
    CLASS-DATA division_by_zero           TYPE REF TO zcl_xlom__va_error READ-ONLY.
    CLASS-DATA field                      TYPE REF TO zcl_xlom__va_error READ-ONLY.
    CLASS-DATA getting_data               TYPE REF TO zcl_xlom__va_error READ-ONLY.
    "! #N/A. Is produced by =ERROR.TYPE(1) or if C1 contains =A1:A2+B1:B3 -> C3=#N/A.
    CLASS-DATA na_not_applicable          TYPE REF TO zcl_xlom__va_error READ-ONLY.
    "! #NAME! Is produced by =XXXX if XXXX is not an existing range name.
    CLASS-DATA name                       TYPE REF TO zcl_xlom__va_error READ-ONLY.
    CLASS-DATA null                       TYPE REF TO zcl_xlom__va_error READ-ONLY.
    "! #NUM! Is produced by =1E+240*1E+240
    CLASS-DATA num                        TYPE REF TO zcl_xlom__va_error READ-ONLY.
    "! TODO #PYTHON! internal error number is not 2222, what is it?
    CLASS-DATA python                     TYPE REF TO zcl_xlom__va_error READ-ONLY.
    "! #REF! Is produced by =INDEX(A1,2,1)
    CLASS-DATA ref                        TYPE REF TO zcl_xlom__va_error READ-ONLY.
    "! #SPILL! Is produced by A1 containing ={1,2} and B1 containing a value -> A1=#SPILL!
    CLASS-DATA spill                      TYPE REF TO zcl_xlom__va_error READ-ONLY.
    CLASS-DATA unknown                    TYPE REF TO zcl_xlom__va_error READ-ONLY.
    "! #VALUE! Is produced by =1+"a". #VALUE! in English, #VALEUR! in French.
    CLASS-DATA value_cannot_be_calculated TYPE REF TO zcl_xlom__va_error READ-ONLY.

    "! English error name (different in other languages e.g. #VALUE! is #VALEUR! in French)
    DATA english_error_name    TYPE string          READ-ONLY.
    "! Example how the error is obtained
    DATA description           TYPE string          READ-ONLY.
    "! Value of enumeration xlCVError e.g. xlErrValue = 2015 (#VALUE!)
    DATA internal_error_number TYPE ty_error_number READ-ONLY.
    "! Result of formula function ERROR.TYPE e.g. 3 for =ERROR.TYPE(#VALUE!)
    DATA formula_error_number  TYPE ty_error_number READ-ONLY.

    CLASS-METHODS class_constructor.

    CLASS-METHODS get_by_error_number
      IMPORTING !type         TYPE ty_error_number
      RETURNING VALUE(result) TYPE REF TO zcl_xlom__va_error.

  PRIVATE SECTION.
    TYPES:
      BEGIN OF ts_error,
        "! English error name (different in other languages e.g. #VALUE! is #VALEUR! in French)
        english_error_name    TYPE string,
        "! Value of enumeration xlCVError e.g. xlErrValue = 2015 (#VALUE!)
        internal_error_number TYPE ty_error_number,
        "! Result of formula function ERROR.TYPE e.g. 3 for =ERROR.TYPE(#VALUE!)
        formula_error_number  TYPE ty_error_number,
        object                TYPE REF TO zcl_xlom__va_error,
      END OF ts_error.
    TYPES tt_error TYPE STANDARD TABLE OF ts_error WITH EMPTY KEY.

    CLASS-DATA errors TYPE tt_error.

    "! @parameter english_error_name | English error name (different in other languages e.g. #VALUE! is #VALEUR! in French)
    "! @parameter internal_error_number | Value of enumeration xlCVError e.g. xlErrValue = 2015 (#VALUE!)
    "! @parameter formula_error_number | Result of formula function ERROR.TYPE e.g. 3 for =ERROR.TYPE(#VALUE!)
    "! @parameter description | Example how the error is obtained
    CLASS-METHODS create
      IMPORTING english_error_name    TYPE string
                internal_error_number TYPE ty_error_number
                formula_error_number  TYPE ty_error_number
                !description          TYPE string OPTIONAL
      RETURNING VALUE(result)         TYPE REF TO zcl_xlom__va_error.
ENDCLASS.


CLASS zcl_xlom__va_error IMPLEMENTATION.
  METHOD class_constructor.
    blocked                    = zcl_xlom__va_error=>create( english_error_name    = '#BLOCKED!     '
                                                             internal_error_number = 2047
                                                             formula_error_number  = 11 ).
    calc                       = zcl_xlom__va_error=>create( english_error_name    = '#CALC!        '
                                                             internal_error_number = 2050
                                                             formula_error_number  = 14 ).
    connect                    = zcl_xlom__va_error=>create( english_error_name    = '#CONNECT!     '
                                                             internal_error_number = 2046
                                                             formula_error_number  = 10 ).
    division_by_zero           = zcl_xlom__va_error=>create( english_error_name    = '#DIV/0!       '
                                                             internal_error_number = 2007
                                                             formula_error_number  = 2
                                                             description           = 'Is produced by =1/0' ).
    field                      = zcl_xlom__va_error=>create( english_error_name    = '#FIELD!       '
                                                             internal_error_number = 2049
                                                             formula_error_number  = 13 ).
    getting_data               = zcl_xlom__va_error=>create( english_error_name    = '#GETTING_DATA!'
                                                             internal_error_number = 2043
                                                             formula_error_number  = 8 ).
    na_not_applicable          = zcl_xlom__va_error=>create(
        english_error_name    = '#N/A          '
        internal_error_number = 2042
        formula_error_number  = 7
        description           = 'Is produced by =ERROR.TYPE(1) or if C1 contains =A1:A2+B1:B3 -> C3=#N/A' ).
    name                       = zcl_xlom__va_error=>create(
        english_error_name    = '#NAME?        '
        internal_error_number = 2029
        formula_error_number  = 5
        description           = 'Is produced by =XXXX if XXXX is not an existing range name' ).
    null                       = zcl_xlom__va_error=>create( english_error_name    = '#NULL!        '
                                                             internal_error_number = 2000
                                                             formula_error_number  = 1 ).
    num                        = zcl_xlom__va_error=>create( english_error_name    = '#NUM!         '
                                                             internal_error_number = 2036
                                                             formula_error_number  = 6
                                                             description           = 'Is produced by =1E+240*1E+240' ).
    python                     = zcl_xlom__va_error=>create( english_error_name    = '#PYTHON!      '
                                                             internal_error_number = 2222
                                                             formula_error_number  = 19 ).
    ref                        = zcl_xlom__va_error=>create( english_error_name    = '#REF!         '
                                                             internal_error_number = 2023
                                                             formula_error_number  = 4
                                                             description           = 'Is produced by =INDEX(A1,2,1)' ).
    spill                      = zcl_xlom__va_error=>create(
        english_error_name    = '#SPILL!       '
        internal_error_number = 2045
        formula_error_number  = 9
        description           = 'Is produced by A1 containing ={1,2} and B1 containing a value -> A1=#SPILL!' ).
    unknown                    = zcl_xlom__va_error=>create( english_error_name    = '#UNKNOWN!     '
                                                             internal_error_number = 2048
                                                             formula_error_number  = 12 ).
    value_cannot_be_calculated = zcl_xlom__va_error=>create( english_error_name    = '#VALUE!       '
                                                             internal_error_number = 2015
                                                             formula_error_number  = 3
                                                             description           = 'Is produced by =1+"a"' ).
  ENDMETHOD.

  METHOD create.
    result = NEW zcl_xlom__va_error( ).
    result->zif_xlom__va~type     = zif_xlom__va=>c_type-error.
    result->english_error_name    = english_error_name.
    result->internal_error_number = internal_error_number.
    result->formula_error_number  = formula_error_number.
    result->description           = description.
    INSERT VALUE #( english_error_name    = english_error_name
                    internal_error_number = internal_error_number
                    formula_error_number  = formula_error_number
                    object                = result )
           INTO TABLE errors.
  ENDMETHOD.

  METHOD get_by_error_number.
    result = errors[ internal_error_number = type ]-object.
  ENDMETHOD.

  METHOD zif_xlom__va~get_value.
    RAISE EXCEPTION TYPE zcx_xlom_todo.
  ENDMETHOD.

  METHOD zif_xlom__va~is_array.
    result = abap_false.
  ENDMETHOD.

  METHOD zif_xlom__va~is_boolean.
    result = abap_false.
  ENDMETHOD.

  METHOD zif_xlom__va~is_equal.
    RAISE EXCEPTION TYPE zcx_xlom_todo.
  ENDMETHOD.

  METHOD zif_xlom__va~is_error.
    result = abap_true.
  ENDMETHOD.

  METHOD zif_xlom__va~is_number.
    result = abap_false.
  ENDMETHOD.

  METHOD zif_xlom__va~is_string.
    result = abap_false.
  ENDMETHOD.
ENDCLASS.
