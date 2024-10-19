class ZCX_XLOM__EX_UT_PARSER definition
  public
  inheriting from CX_STATIC_CHECK
  create public .

public section.

  methods CONSTRUCTOR
    importing
      !TEXT type CSEQUENCE optional
      !MSGV1 type CSEQUENCE optional
      !MSGV2 type CSEQUENCE optional
      !MSGV3 type CSEQUENCE optional
      !MSGV4 type CSEQUENCE optional
      !TEXTID like TEXTID optional
      !PREVIOUS like PREVIOUS optional .

  methods GET_TEXT
    redefinition .
  methods GET_LONGTEXT
    redefinition .
protected section.
private section.

  data TEXT type STRING .
  data MSGV1 type STRING .
  data MSGV2 type STRING .
  data MSGV3 type STRING .
  data MSGV4 type STRING .
ENDCLASS.



CLASS ZCX_XLOM__EX_UT_PARSER IMPLEMENTATION.


  method CONSTRUCTOR.

    super->constructor( previous = previous
                        textid   = textid ).
    me->text  = text.
    me->msgv1 = msgv1.
    me->msgv2 = msgv2.
    me->msgv3 = msgv3.
    me->msgv4 = msgv4.

  endmethod.


  method GET_LONGTEXT.

    IF text IS NOT INITIAL.
      result = get_text( ).
    ELSE.
      result = super->get_longtext( ).
    ENDIF.

  endmethod.


  method GET_TEXT.

    IF text IS NOT INITIAL.
      result = text.
      REPLACE '&1' IN result WITH msgv1.
      REPLACE '&2' IN result WITH msgv2.
      REPLACE '&3' IN result WITH msgv3.
      REPLACE '&4' IN result WITH msgv4.
    ELSE.
      result = super->get_text( ).
    ENDIF.

  endmethod.
ENDCLASS.
