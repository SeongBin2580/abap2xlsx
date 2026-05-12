CLASS zcl_pp_xls_style_border DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

  PUBLIC SECTION.

*"* public components of class ZCL_EXCEL_STYLE_BORDER
*"* do not include other source files here!!!
    DATA border_style TYPE zppe_xls_border .
    DATA border_color TYPE zpps_xls_style_color .
    CONSTANTS c_border_none TYPE zppe_xls_border VALUE 'none'. "#EC NOTEXT
    CONSTANTS c_border_dashdot TYPE zppe_xls_border VALUE 'dashDot'. "#EC NOTEXT
    CONSTANTS c_border_dashdotdot TYPE zppe_xls_border VALUE 'dashDotDot'. "#EC NOTEXT
    CONSTANTS c_border_dashed TYPE zppe_xls_border VALUE 'dashed'. "#EC NOTEXT
    CONSTANTS c_border_dotted TYPE zppe_xls_border VALUE 'dotted'. "#EC NOTEXT
    CONSTANTS c_border_double TYPE zppe_xls_border VALUE 'double'. "#EC NOTEXT
    CONSTANTS c_border_hair TYPE zppe_xls_border VALUE 'hair'. "#EC NOTEXT
    CONSTANTS c_border_medium TYPE zppe_xls_border VALUE 'medium'. "#EC NOTEXT
    CONSTANTS c_border_mediumdashdot TYPE zppe_xls_border VALUE 'mediumDashDot'. "#EC NOTEXT
    CONSTANTS c_border_mediumdashdotdot TYPE zppe_xls_border VALUE 'mediumDashDotDot'. "#EC NOTEXT
    CONSTANTS c_border_mediumdashed TYPE zppe_xls_border VALUE 'mediumDashed'. "#EC NOTEXT
    CONSTANTS c_border_slantdashdot TYPE zppe_xls_border VALUE 'slantDashDot'. "#EC NOTEXT
    CONSTANTS c_border_thick TYPE zppe_xls_border VALUE 'thick'. "#EC NOTEXT
    CONSTANTS c_border_thin TYPE zppe_xls_border VALUE 'thin'. "#EC NOTEXT

    METHODS constructor .
*"* protected components of class ZABAP_EXCEL_STYLE_FONT
*"* do not include other source files here!!!
*"* protected components of class ZABAP_EXCEL_STYLE_FONT
*"* do not include other source files here!!!
  PROTECTED SECTION.
*"* private components of class ZCL_EXCEL_STYLE_BORDER
*"* do not include other source files here!!!
  PRIVATE SECTION.
ENDCLASS.



CLASS zcl_pp_xls_style_border IMPLEMENTATION.


  METHOD constructor.
    border_style = zcl_pp_xls_style_border=>c_border_none.
    border_color-theme     = zcl_pp_xls_style_color=>c_theme_not_set.
    border_color-indexed   = zcl_pp_xls_style_color=>c_indexed_not_set.
  ENDMETHOD.
ENDCLASS.
