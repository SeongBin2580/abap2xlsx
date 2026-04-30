INTERFACE zif_yy_xls_sheet_properties
  PUBLIC .


  CONSTANTS c_hidden TYPE zyye_xls_sheet_hidden VALUE 'X'.    "#EC NOTEXT
  CONSTANTS c_veryhidden TYPE zyye_xls_sheet_hidden VALUE '2'. "#EC NOTEXT
  CONSTANTS c_hidezero TYPE zyye_xls_sheet_showzeros VALUE ''. "#EC NOTEXT
  CONSTANTS c_showzero TYPE zyye_xls_sheet_showzeros VALUE 'X'. "#EC NOTEXT
  CONSTANTS c_visible TYPE zyye_xls_sheet_hidden VALUE ''.    "#EC NOTEXT
  DATA hidden TYPE zyye_xls_sheet_hidden .
  DATA show_zeros TYPE zyye_xls_sheet_showzeros .
  DATA style TYPE zyye_xls_cell_style .
  DATA zoomscale TYPE zyye_xls_sheet_zoomscale .
  DATA zoomscale_normal TYPE zyye_xls_sheet_zoomscale .
  DATA zoomscale_pagelayoutview TYPE zyye_xls_sheet_zoomscale .
  DATA zoomscale_sheetlayoutview TYPE zyye_xls_sheet_zoomscale .
  DATA summarybelow TYPE zyye_xls_sheet_summary .
  CONSTANTS c_below_on TYPE zyye_xls_sheet_summary VALUE 1.   "#EC NOTEXT
  CONSTANTS c_below_off TYPE zyye_xls_sheet_summary VALUE 0.  "#EC NOTEXT
  DATA summaryright TYPE zyye_xls_sheet_summary .
  CONSTANTS c_right_on TYPE zyye_xls_sheet_summary VALUE 1.   "#EC NOTEXT
  CONSTANTS c_right_off TYPE zyye_xls_sheet_summary VALUE 0.  "#EC NOTEXT
  DATA selected TYPE zyye_xls_sheet_selected .
  CONSTANTS c_selected TYPE zyye_xls_sheet_selected VALUE 'X'. "#EC NOTEXT
  DATA hide_columns_from TYPE zyye_xls_cell_column_alpha .

  METHODS initialize .
  METHODS get_right_to_left
    RETURNING
      VALUE(result) TYPE abap_bool.
  METHODS get_style
    RETURNING
      VALUE(ep_style) TYPE zyye_xls_cell_style .
  METHODS set_right_to_left
    IMPORTING
      !right_to_left TYPE abap_bool .
  METHODS set_style
    IMPORTING
      !ip_style TYPE zyye_xls_cell_style .
ENDINTERFACE.
