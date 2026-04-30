CLASS zcl_yy_xls_theme_objectdefault DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

  PUBLIC SECTION.

    METHODS load
      IMPORTING
        !io_object_def TYPE REF TO if_ixml_element .
    METHODS build_xml
      IMPORTING
        !io_document TYPE REF TO if_ixml_document .
  PROTECTED SECTION.
  PRIVATE SECTION.

    DATA objectdefaults TYPE REF TO if_ixml_element .
ENDCLASS.



CLASS zcl_yy_xls_theme_objectdefault IMPLEMENTATION.


  METHOD build_xml.
    DATA: lo_theme TYPE REF TO if_ixml_element.
    DATA: lo_theme_objdef TYPE REF TO if_ixml_element.
    CHECK io_document IS BOUND.
    lo_theme ?= io_document->get_root_element( ).
    CHECK lo_theme IS BOUND.
    IF objectdefaults IS INITIAL.
      lo_theme_objdef ?= io_document->create_simple_element_ns( prefix = zcl_yy_xls_theme=>c_theme_prefix
                                                                name   = zcl_yy_xls_theme=>c_theme_object_def
                                                                parent = lo_theme ).
    ELSE.
      lo_theme->append_child( new_child = objectdefaults ).
    ENDIF.
  ENDMETHOD.                    "build_xml


  METHOD load.
    objectdefaults = zcl_yy_xls_common=>clone_ixml_with_namespaces( io_object_def ).
  ENDMETHOD.                    "load
ENDCLASS.
