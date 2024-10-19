class ZCX_XLOM__VA definition
  public
  inheriting from CX_STATIC_CHECK
  create public .

public section.

  data RESULT_ERROR type ref to ZCL_XLOM__VA_ERROR read-only .

  methods CONSTRUCTOR
    importing
      !RESULT_ERROR type ref to ZCL_XLOM__VA_ERROR .
protected section.
private section.
ENDCLASS.



CLASS ZCX_XLOM__VA IMPLEMENTATION.


  method CONSTRUCTOR.

    super->constructor( textid = textid previous = previous ).
    me->result_error = result_error.

  endmethod.
ENDCLASS.
