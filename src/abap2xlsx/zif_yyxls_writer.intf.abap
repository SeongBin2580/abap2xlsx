INTERFACE zif_yyxls_writer
  PUBLIC .


  METHODS write_file
    IMPORTING
      !io_excel      TYPE REF TO zcl_yyxls
    RETURNING
      VALUE(ep_file) TYPE xstring
    RAISING
      zcx_yyxls.
ENDINTERFACE.
