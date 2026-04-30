INTERFACE zif_yy_xls_style_changer
  PUBLIC .

  METHODS apply
    IMPORTING
      ip_worksheet   TYPE REF TO zcl_yy_xls_worksheet
      ip_column      TYPE simple
      ip_row         TYPE zyye_xls_cell_row
    RETURNING
      VALUE(ep_guid) TYPE zyye_xls_cell_style
    RAISING
      zcx_yy_xls.
  METHODS get_guid
    RETURNING
      VALUE(result) TYPE zyye_xls_cell_style.
  METHODS set_complete
    IMPORTING
      ip_complete   TYPE zyys_xls_cstyle_complete
      ip_xcomplete  TYPE zyys_xls_cstylex_complete
    RETURNING
      VALUE(result) TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_complete_font
    IMPORTING
      ip_font       TYPE zyys_xls_cstyle_font
      ip_xfont      TYPE zyys_xls_cstylex_font OPTIONAL
    RETURNING
      VALUE(result) TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_complete_fill
    IMPORTING
      ip_fill       TYPE zyys_xls_cstyle_fill
      ip_xfill      TYPE zyys_xls_cstylex_fill OPTIONAL
    RETURNING
      VALUE(result) TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_complete_borders
    IMPORTING
      ip_borders    TYPE zyys_xls_cstyle_borders
      ip_xborders   TYPE zyys_xls_cstylex_borders OPTIONAL
    RETURNING
      VALUE(result) TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_complete_alignment
    IMPORTING
      ip_alignment  TYPE zyys_xls_cstyle_alignment
      ip_xalignment TYPE zyys_xls_cstylex_alignment OPTIONAL
    RETURNING
      VALUE(result) TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_complete_protection
    IMPORTING
      ip_protection  TYPE zyys_xls_cstyle_protection
      ip_xprotection TYPE zyys_xls_cstylex_protection OPTIONAL
    RETURNING
      VALUE(result)  TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_complete_borders_all
    IMPORTING
      ip_borders_allborders  TYPE zyys_xls_cstyle_border
      ip_xborders_allborders TYPE zyys_xls_cstylex_border OPTIONAL
    RETURNING
      VALUE(result)          TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_complete_borders_diagonal
    IMPORTING
      ip_borders_diagonal  TYPE zyys_xls_cstyle_border
      ip_xborders_diagonal TYPE zyys_xls_cstylex_border OPTIONAL
    RETURNING
      VALUE(result)        TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_complete_borders_down
    IMPORTING
      ip_borders_down  TYPE zyys_xls_cstyle_border
      ip_xborders_down TYPE zyys_xls_cstylex_border OPTIONAL
    RETURNING
      VALUE(result)    TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_complete_borders_left
    IMPORTING
      ip_borders_left  TYPE zyys_xls_cstyle_border
      ip_xborders_left TYPE zyys_xls_cstylex_border OPTIONAL
    RETURNING
      VALUE(result)    TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_complete_borders_right
    IMPORTING
      ip_borders_right  TYPE zyys_xls_cstyle_border
      ip_xborders_right TYPE zyys_xls_cstylex_border OPTIONAL
    RETURNING
      VALUE(result)     TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_complete_borders_top
    IMPORTING
      ip_borders_top  TYPE zyys_xls_cstyle_border
      ip_xborders_top TYPE zyys_xls_cstylex_border OPTIONAL
    RETURNING
      VALUE(result)   TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_number_format
    IMPORTING
      value         TYPE zyye_xls_number_format
    RETURNING
      VALUE(result) TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_font_bold
    IMPORTING
      value         TYPE flag
    RETURNING
      VALUE(result) TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_font_color
    IMPORTING
      value         TYPE zyys_xls_style_color
    RETURNING
      VALUE(result) TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_font_color_rgb
    IMPORTING
      value         TYPE zyye_xls_style_color_argb
    RETURNING
      VALUE(result) TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_font_color_indexed
    IMPORTING
      value         TYPE zyye_xls_style_color_indexed
    RETURNING
      VALUE(result) TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_font_color_theme
    IMPORTING
      value         TYPE zyye_xls_style_color_theme
    RETURNING
      VALUE(result) TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_font_color_tint
    IMPORTING
      value         TYPE zyye_xls_style_color_tint
    RETURNING
      VALUE(result) TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_font_family
    IMPORTING
      value         TYPE zyye_xls_style_font_family
    RETURNING
      VALUE(result) TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_font_italic
    IMPORTING
      value         TYPE flag
    RETURNING
      VALUE(result) TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_font_name
    IMPORTING
      value         TYPE zyye_xls_style_font_name
    RETURNING
      VALUE(result) TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_font_scheme
    IMPORTING
      value         TYPE zyye_xls_style_font_scheme
    RETURNING
      VALUE(result) TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_font_size
    IMPORTING
      value         TYPE numeric
    RETURNING
      VALUE(result) TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_font_strikethrough
    IMPORTING
      value         TYPE flag
    RETURNING
      VALUE(result) TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_font_underline
    IMPORTING
      value         TYPE flag
    RETURNING
      VALUE(result) TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_font_underline_mode
    IMPORTING
      value         TYPE zyye_xls_style_font_underline
    RETURNING
      VALUE(result) TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_fill_filltype
    IMPORTING
      value         TYPE zyye_xls_fill_type
    RETURNING
      VALUE(result) TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_fill_rotation
    IMPORTING
      value         TYPE zyye_xls_rotation
    RETURNING
      VALUE(result) TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_fill_fgcolor
    IMPORTING
      value         TYPE zyys_xls_style_color
    RETURNING
      VALUE(result) TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_fill_fgcolor_rgb
    IMPORTING
      value         TYPE zyye_xls_style_color_argb
    RETURNING
      VALUE(result) TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_fill_fgcolor_indexed
    IMPORTING
      value         TYPE zyye_xls_style_color_indexed
    RETURNING
      VALUE(result) TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_fill_fgcolor_theme
    IMPORTING
      value         TYPE zyye_xls_style_color_theme
    RETURNING
      VALUE(result) TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_fill_fgcolor_tint
    IMPORTING
      value         TYPE zyye_xls_style_color_tint
    RETURNING
      VALUE(result) TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_fill_bgcolor
    IMPORTING
      value         TYPE zyys_xls_style_color
    RETURNING
      VALUE(result) TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_fill_bgcolor_rgb
    IMPORTING
      value         TYPE zyye_xls_style_color_argb
    RETURNING
      VALUE(result) TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_fill_bgcolor_indexed
    IMPORTING
      value         TYPE zyye_xls_style_color_indexed
    RETURNING
      VALUE(result) TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_fill_bgcolor_theme
    IMPORTING
      value         TYPE zyye_xls_style_color_theme
    RETURNING
      VALUE(result) TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_fill_bgcolor_tint
    IMPORTING
      value         TYPE zyye_xls_style_color_tint
    RETURNING
      VALUE(result) TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_fill_gradtype_type
    IMPORTING
      value         TYPE zyys_xls_gradient_type-type
    RETURNING
      VALUE(result) TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_fill_gradtype_degree
    IMPORTING
      value         TYPE zyys_xls_gradient_type-degree
    RETURNING
      VALUE(result) TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_fill_gradtype_bottom
    IMPORTING
      value         TYPE zyys_xls_gradient_type-bottom
    RETURNING
      VALUE(result) TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_fill_gradtype_left
    IMPORTING
      value         TYPE zyys_xls_gradient_type-left
    RETURNING
      VALUE(result) TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_fill_gradtype_top
    IMPORTING
      value         TYPE zyys_xls_gradient_type-top
    RETURNING
      VALUE(result) TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_fill_gradtype_right
    IMPORTING
      value         TYPE zyys_xls_gradient_type-right
    RETURNING
      VALUE(result) TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_fill_gradtype_position1
    IMPORTING
      value         TYPE zyys_xls_gradient_type-position1
    RETURNING
      VALUE(result) TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_fill_gradtype_position2
    IMPORTING
      value         TYPE zyys_xls_gradient_type-position2
    RETURNING
      VALUE(result) TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_fill_gradtype_position3
    IMPORTING
      value         TYPE zyys_xls_gradient_type-position3
    RETURNING
      VALUE(result) TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_borders_diagonal_mode
    IMPORTING
      value         TYPE zyye_xls_diagonal
    RETURNING
      VALUE(result) TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_alignment_horizontal
    IMPORTING
      value         TYPE zyye_xls_alignment
    RETURNING
      VALUE(result) TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_alignment_vertical
    IMPORTING
      value         TYPE zyye_xls_alignment
    RETURNING
      VALUE(result) TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_alignment_textrotation
    IMPORTING
      value         TYPE zyye_xls_text_rotation
    RETURNING
      VALUE(result) TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_alignment_wraptext
    IMPORTING
      value         TYPE flag
    RETURNING
      VALUE(result) TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_alignment_shrinktofit
    IMPORTING
      value         TYPE flag
    RETURNING
      VALUE(result) TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_alignment_indent
    IMPORTING
      value         TYPE zyye_xls_indent
    RETURNING
      VALUE(result) TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_protection_hidden
    IMPORTING
      value         TYPE zyye_xls_cell_protection
    RETURNING
      VALUE(result) TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_protection_locked
    IMPORTING
      value         TYPE zyye_xls_cell_protection
    RETURNING
      VALUE(result) TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_borders_allborders_style
    IMPORTING
      value         TYPE zyye_xls_border
    RETURNING
      VALUE(result) TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_borders_allborders_color
    IMPORTING
      value         TYPE zyys_xls_style_color
    RETURNING
      VALUE(result) TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_borders_allbo_color_rgb
    IMPORTING
      value         TYPE zyye_xls_style_color_argb
    RETURNING
      VALUE(result) TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_borders_allbo_color_indexe
    IMPORTING
      value         TYPE zyye_xls_style_color_indexed
    RETURNING
      VALUE(result) TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_borders_allbo_color_theme
    IMPORTING
      value         TYPE zyye_xls_style_color_theme
    RETURNING
      VALUE(result) TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_borders_allbo_color_tint
    IMPORTING
      value         TYPE zyye_xls_style_color_tint
    RETURNING
      VALUE(result) TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_borders_diagonal_style
    IMPORTING
      value         TYPE zyye_xls_border
    RETURNING
      VALUE(result) TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_borders_diagonal_color
    IMPORTING
      value         TYPE zyys_xls_style_color
    RETURNING
      VALUE(result) TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_borders_diagonal_color_rgb
    IMPORTING
      value         TYPE zyye_xls_style_color_argb
    RETURNING
      VALUE(result) TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_borders_diagonal_color_ind
    IMPORTING
      value         TYPE zyye_xls_style_color_indexed
    RETURNING
      VALUE(result) TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_borders_diagonal_color_the
    IMPORTING
      value         TYPE zyye_xls_style_color_theme
    RETURNING
      VALUE(result) TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_borders_diagonal_color_tin
    IMPORTING
      value         TYPE zyye_xls_style_color_tint
    RETURNING
      VALUE(result) TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_borders_down_style
    IMPORTING
      value         TYPE zyye_xls_border
    RETURNING
      VALUE(result) TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_borders_down_color
    IMPORTING
      value         TYPE zyys_xls_style_color
    RETURNING
      VALUE(result) TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_borders_down_color_rgb
    IMPORTING
      value         TYPE zyye_xls_style_color_argb
    RETURNING
      VALUE(result) TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_borders_down_color_indexed
    IMPORTING
      value         TYPE zyye_xls_style_color_indexed
    RETURNING
      VALUE(result) TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_borders_down_color_theme
    IMPORTING
      value         TYPE zyye_xls_style_color_theme
    RETURNING
      VALUE(result) TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_borders_down_color_tint
    IMPORTING
      value         TYPE zyye_xls_style_color_tint
    RETURNING
      VALUE(result) TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_borders_left_style
    IMPORTING
      value         TYPE zyye_xls_border
    RETURNING
      VALUE(result) TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_borders_left_color
    IMPORTING
      value         TYPE zyys_xls_style_color
    RETURNING
      VALUE(result) TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_borders_left_color_rgb
    IMPORTING
      value         TYPE zyye_xls_style_color_argb
    RETURNING
      VALUE(result) TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_borders_left_color_indexed
    IMPORTING
      value         TYPE zyye_xls_style_color_indexed
    RETURNING
      VALUE(result) TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_borders_left_color_theme
    IMPORTING
      value         TYPE zyye_xls_style_color_theme
    RETURNING
      VALUE(result) TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_borders_left_color_tint
    IMPORTING
      value         TYPE zyye_xls_style_color_tint
    RETURNING
      VALUE(result) TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_borders_right_style
    IMPORTING
      value         TYPE zyye_xls_border
    RETURNING
      VALUE(result) TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_borders_right_color
    IMPORTING
      value         TYPE zyys_xls_style_color
    RETURNING
      VALUE(result) TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_borders_right_color_rgb
    IMPORTING
      value         TYPE zyye_xls_style_color_argb
    RETURNING
      VALUE(result) TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_borders_right_color_indexe
    IMPORTING
      value         TYPE zyye_xls_style_color_indexed
    RETURNING
      VALUE(result) TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_borders_right_color_theme
    IMPORTING
      value         TYPE zyye_xls_style_color_theme
    RETURNING
      VALUE(result) TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_borders_right_color_tint
    IMPORTING
      value         TYPE zyye_xls_style_color_tint
    RETURNING
      VALUE(result) TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_borders_top_style
    IMPORTING
      value         TYPE zyye_xls_border
    RETURNING
      VALUE(result) TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_borders_top_color
    IMPORTING
      value         TYPE zyys_xls_style_color
    RETURNING
      VALUE(result) TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_borders_top_color_rgb
    IMPORTING
      value         TYPE zyye_xls_style_color_argb
    RETURNING
      VALUE(result) TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_borders_top_color_indexed
    IMPORTING
      value         TYPE zyye_xls_style_color_indexed
    RETURNING
      VALUE(result) TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_borders_top_color_theme
    IMPORTING
      value         TYPE zyye_xls_style_color_theme
    RETURNING
      VALUE(result) TYPE REF TO zif_yy_xls_style_changer.
  METHODS set_borders_top_color_tint
    IMPORTING
      value         TYPE zyye_xls_style_color_tint
    RETURNING
      VALUE(result) TYPE REF TO zif_yy_xls_style_changer.
  DATA: complete_style  TYPE zyys_xls_cstyle_complete READ-ONLY,
        complete_stylex TYPE zyys_xls_cstylex_complete READ-ONLY.
ENDINTERFACE.
