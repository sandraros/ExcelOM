*&---------------------------------------------------------------------*
*& Report z_xlom_viewer
*&---------------------------------------------------------------------*
*&
*&---------------------------------------------------------------------*
REPORT z_xlom_viewer.
DATA(xlom_application) = zcl_xlom_application=>create( ).

DATA(xlom_workbook) = xlom_application->workbooks->add( ).

DATA(xlom_worksheet) = xlom_workbook->worksheets->add( 'Sheet1' ).

xlom_worksheet->range( cell1_string = 'A1' )->set_value( zcl_xlom__va_number=>get( 25 ) ).

xlom_worksheet->range( cell1_string = 'B1' )->set_formula2( 'A1+10' ).

*ASSERT 35 = CAST zcl_xlom__va_number( xlom_worksheet->range( cell1_string = 'B1' )->value( ) )->get_number( ).
zcl_xlom__vw=>view( application = xlom_application ).

ASSERT 1 = 1. " Debug helper to set a break-point
