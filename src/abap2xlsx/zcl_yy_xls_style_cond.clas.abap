CLASS zcl_yy_xls_style_cond DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

  PUBLIC SECTION.
    CLASS zcl_excel_style_conditional DEFINITION LOAD .

*"* public components of class ZCL_EXCEL_STYLE_COND
*"* do not include other source files here!!!
    TYPES tv_conditional_show_value TYPE c LENGTH 1.
    TYPES tv_textfunction TYPE string.
    TYPES: BEGIN OF ts_conditional_textfunction,
             text         TYPE string,
             textfunction TYPE tv_textfunction,
             cell_style   TYPE zyye_xls_cell_style,
           END OF ts_conditional_textfunction.
    CONSTANTS c_cfvo_type_formula TYPE zyye_xls_conditional_type VALUE 'formula'. "#EC NOTEXT
    CONSTANTS c_cfvo_type_max TYPE zyye_xls_conditional_type VALUE 'max'. "#EC NOTEXT
    CONSTANTS c_cfvo_type_min TYPE zyye_xls_conditional_type VALUE 'min'. "#EC NOTEXT
    CONSTANTS c_cfvo_type_number TYPE zyye_xls_conditional_type VALUE 'num'. "#EC NOTEXT
    CONSTANTS c_cfvo_type_percent TYPE zyye_xls_conditional_type VALUE 'percent'. "#EC NOTEXT
    CONSTANTS c_cfvo_type_percentile TYPE zyye_xls_conditional_type VALUE 'percentile'. "#EC NOTEXT
    CONSTANTS c_iconset_3arrows TYPE zyye_xls_conditi_rule_iconset VALUE '3Arrows'. "#EC NOTEXT
    CONSTANTS c_iconset_3arrowsgray TYPE zyye_xls_conditi_rule_iconset VALUE '3ArrowsGray'. "#EC NOTEXT
    CONSTANTS c_iconset_3flags TYPE zyye_xls_conditi_rule_iconset VALUE '3Flags'. "#EC NOTEXT
    CONSTANTS c_iconset_3signs TYPE zyye_xls_conditi_rule_iconset VALUE '3Signs'. "#EC NOTEXT
    CONSTANTS c_iconset_3symbols TYPE zyye_xls_conditi_rule_iconset VALUE '3Symbols'. "#EC NOTEXT
    CONSTANTS c_iconset_3symbols2 TYPE zyye_xls_conditi_rule_iconset VALUE '3Symbols2'. "#EC NOTEXT
    CONSTANTS c_iconset_3trafficlights TYPE zyye_xls_conditi_rule_iconset VALUE ''. "#EC NOTEXT
    CONSTANTS c_iconset_3trafficlights2 TYPE zyye_xls_conditi_rule_iconset VALUE '3TrafficLights2'. "#EC NOTEXT
    CONSTANTS c_iconset_4arrows TYPE zyye_xls_conditi_rule_iconset VALUE '4Arrows'. "#EC NOTEXT
    CONSTANTS c_iconset_4arrowsgray TYPE zyye_xls_conditi_rule_iconset VALUE '4ArrowsGray'. "#EC NOTEXT
    CONSTANTS c_iconset_4rating TYPE zyye_xls_conditi_rule_iconset VALUE '4Rating'. "#EC NOTEXT
    CONSTANTS c_iconset_4redtoblack TYPE zyye_xls_conditi_rule_iconset VALUE '4RedToBlack'. "#EC NOTEXT
    CONSTANTS c_iconset_4trafficlights TYPE zyye_xls_conditi_rule_iconset VALUE '4TrafficLights'. "#EC NOTEXT
    CONSTANTS c_iconset_5arrows TYPE zyye_xls_conditi_rule_iconset VALUE '5Arrows'. "#EC NOTEXT
    CONSTANTS c_iconset_5arrowsgray TYPE zyye_xls_conditi_rule_iconset VALUE '5ArrowsGray'. "#EC NOTEXT
    CONSTANTS c_iconset_5quarters TYPE zyye_xls_conditi_rule_iconset VALUE '5Quarters'. "#EC NOTEXT
    CONSTANTS c_iconset_5rating TYPE zyye_xls_conditi_rule_iconset VALUE '5Rating'. "#EC NOTEXT
    CONSTANTS c_operator_beginswith TYPE zyye_xls_conditi_operator VALUE 'beginsWith'. "#EC NOTEXT
    CONSTANTS c_operator_between TYPE zyye_xls_conditi_operator VALUE 'between'. "#EC NOTEXT
    CONSTANTS c_operator_containstext TYPE zyye_xls_conditi_operator VALUE 'containsText'. "#EC NOTEXT
    CONSTANTS c_operator_endswith TYPE zyye_xls_conditi_operator VALUE 'endsWith'. "#EC NOTEXT
    CONSTANTS c_operator_equal TYPE zyye_xls_conditi_operator VALUE 'equal'. "#EC NOTEXT
    CONSTANTS c_operator_greaterthan TYPE zyye_xls_conditi_operator VALUE 'greaterThan'. "#EC NOTEXT
    CONSTANTS c_operator_greaterthanorequal TYPE zyye_xls_conditi_operator VALUE 'greaterThanOrEqual'. "#EC NOTEXT
    CONSTANTS c_operator_lessthan TYPE zyye_xls_conditi_operator VALUE 'lessThan'. "#EC NOTEXT
    CONSTANTS c_operator_lessthanorequal TYPE zyye_xls_conditi_operator VALUE 'lessThanOrEqual'. "#EC NOTEXT
    CONSTANTS c_operator_none TYPE zyye_xls_conditi_operator VALUE ''. "#EC NOTEXT
    CONSTANTS c_operator_notcontains TYPE zyye_xls_conditi_operator VALUE 'notContains'. "#EC NOTEXT
    CONSTANTS c_operator_notequal TYPE zyye_xls_conditi_operator VALUE 'notEqual'. "#EC NOTEXT
    CONSTANTS c_textfunction_beginswith TYPE tv_textfunction VALUE 'beginsWith'. "#EC NOTEXT
    CONSTANTS c_textfunction_containstext TYPE tv_textfunction VALUE 'containsText'. "#EC NOTEXT
    CONSTANTS c_textfunction_endswith TYPE tv_textfunction VALUE 'endsWith'. "#EC NOTEXT
    CONSTANTS c_textfunction_notcontains TYPE tv_textfunction VALUE 'notContains'. "#EC NOTEXT
    CONSTANTS c_rule_cellis TYPE zyye_xls_conditi_rule VALUE 'cellIs'. "#EC NOTEXT
    CONSTANTS c_rule_containstext TYPE zyye_xls_conditi_rule VALUE 'containsText'. "#EC NOTEXT
    CONSTANTS c_rule_databar TYPE zyye_xls_conditi_rule VALUE 'dataBar'. "#EC NOTEXT
    CONSTANTS c_rule_expression TYPE zyye_xls_conditi_rule VALUE 'expression'. "#EC NOTEXT
    CONSTANTS c_rule_iconset TYPE zyye_xls_conditi_rule VALUE 'iconSet'. "#EC NOTEXT
    CONSTANTS c_rule_colorscale TYPE zyye_xls_conditi_rule VALUE 'colorScale'. "#EC NOTEXT
    CONSTANTS c_rule_none TYPE zyye_xls_conditi_rule VALUE 'none'. "#EC NOTEXT
    CONSTANTS c_rule_textfunction TYPE zyye_xls_conditi_rule VALUE '<textfunction>'. "#EC NOTEXT
    CONSTANTS c_rule_top10 TYPE zyye_xls_conditi_rule VALUE 'top10'. "#EC NOTEXT
    CONSTANTS c_rule_above_average TYPE zyye_xls_conditi_rule VALUE 'aboveAverage'. "#EC NOTEXT
    CONSTANTS c_showvalue_false TYPE tv_conditional_show_value VALUE 0. "#EC NOTEXT
    CONSTANTS c_showvalue_true TYPE tv_conditional_show_value VALUE 1. "#EC NOTEXT
    DATA mode_cellis TYPE zexcel_conditional_cellis .
    DATA mode_textfunction TYPE ts_conditional_textfunction .
    DATA mode_colorscale TYPE zexcel_conditional_colorscale .
    DATA mode_databar TYPE zexcel_conditional_databar .
    DATA mode_expression TYPE zexcel_conditional_expression .
    DATA mode_iconset TYPE zexcel_conditional_iconset .
    DATA mode_top10 TYPE zexcel_conditional_top10 .
    DATA mode_above_average TYPE zexcel_conditional_above_avg .
    DATA priority TYPE zyye_xls_style_priority VALUE 1. "#EC NOTEXT .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  . " .
    DATA rule TYPE zyye_xls_conditi_rule .

    METHODS constructor
      IMPORTING
        !ip_guid            TYPE zyye_xls_cell_style OPTIONAL
        !ip_dimension_range TYPE string.
    METHODS get_dimension_range
      RETURNING
        VALUE(ep_dimension_range) TYPE string .
    METHODS set_range
      IMPORTING
        !ip_start_row    TYPE zyye_xls_cell_row
        !ip_start_column TYPE zyye_xls_cell_column_alpha
        !ip_stop_row     TYPE zyye_xls_cell_row
        !ip_stop_column  TYPE zyye_xls_cell_column_alpha
      RAISING
        zcx_yy_xls .
    METHODS add_range
      IMPORTING
        !ip_start_row    TYPE zyye_xls_cell_row
        !ip_start_column TYPE zyye_xls_cell_column_alpha
        !ip_stop_row     TYPE zyye_xls_cell_row
        !ip_stop_column  TYPE zyye_xls_cell_column_alpha
      RAISING
        zcx_yy_xls .
    CLASS-METHODS factory_cond_style_iconset
      IMPORTING
        !io_worksheet        TYPE REF TO zcl_yy_xls_worksheet
        !iv_icon_type        TYPE zyye_xls_conditi_rule_iconset DEFAULT c_iconset_3trafficlights2
        !iv_cfvo1_type       TYPE zyye_xls_conditional_type DEFAULT c_cfvo_type_percent
        !iv_cfvo1_value      TYPE zyye_xls_conditional_value OPTIONAL
        !iv_cfvo2_type       TYPE zyye_xls_conditional_type DEFAULT c_cfvo_type_percent
        !iv_cfvo2_value      TYPE zyye_xls_conditional_value OPTIONAL
        !iv_cfvo3_type       TYPE zyye_xls_conditional_type DEFAULT c_cfvo_type_percent
        !iv_cfvo3_value      TYPE zyye_xls_conditional_value OPTIONAL
        !iv_cfvo4_type       TYPE zyye_xls_conditional_type DEFAULT c_cfvo_type_percent
        !iv_cfvo4_value      TYPE zyye_xls_conditional_value OPTIONAL
        !iv_cfvo5_type       TYPE zyye_xls_conditional_type DEFAULT c_cfvo_type_percent
        !iv_cfvo5_value      TYPE zyye_xls_conditional_value OPTIONAL
        !iv_showvalue        TYPE tv_conditional_show_value DEFAULT zcl_yy_xls_style_cond=>c_showvalue_true
      RETURNING
        VALUE(eo_style_cond) TYPE REF TO zcl_yy_xls_style_cond .
    METHODS get_guid
      RETURNING
        VALUE(ep_guid) TYPE zyye_xls_cell_style .
*"* protected components of class ZABAP_EXCEL_STYLE_FONT
*"* do not include other source files here!!!
*"* protected components of class ZABAP_EXCEL_STYLE_FONT
*"* do not include other source files here!!!
  PROTECTED SECTION.
  PRIVATE SECTION.

    DATA mv_rule_range TYPE string .
    DATA guid TYPE zyye_xls_cell_style .
ENDCLASS.



CLASS zcl_yy_xls_style_cond IMPLEMENTATION.


  METHOD add_range.
    DATA: lv_column    TYPE zyye_xls_cell_column,
          lv_row_alpha TYPE string,
          lv_col_alpha TYPE string,
          lv_coords1   TYPE string,
          lv_coords2   TYPE string.


    lv_column = zcl_yy_xls_common=>convert_column2int( ip_start_column ).

    lv_col_alpha = ip_start_column.
    lv_row_alpha = ip_start_row.
    SHIFT lv_row_alpha RIGHT DELETING TRAILING space.
    SHIFT lv_row_alpha LEFT DELETING LEADING space.
    CONCATENATE lv_col_alpha lv_row_alpha INTO lv_coords1.

    IF ip_stop_column IS NOT INITIAL.
      lv_column = zcl_yy_xls_common=>convert_column2int( ip_stop_column ).
    ELSE.
      lv_column = zcl_yy_xls_common=>convert_column2int( ip_start_column ).
    ENDIF.

    IF ip_stop_row IS NOT INITIAL. " If we don't get explicitly a stop column use start column
      lv_row_alpha = ip_stop_row.
    ELSE.
      lv_row_alpha = ip_start_row.
    ENDIF.
    IF ip_stop_column IS NOT INITIAL. " If we don't get explicitly a stop column use start column
      lv_col_alpha = ip_stop_column.
    ELSE.
      lv_col_alpha = ip_start_column.
    ENDIF.
    SHIFT lv_row_alpha RIGHT DELETING TRAILING space.
    SHIFT lv_row_alpha LEFT DELETING LEADING space.
    CONCATENATE lv_col_alpha lv_row_alpha INTO lv_coords2.
    IF lv_coords2 IS NOT INITIAL AND lv_coords2 <> lv_coords1.
      CONCATENATE me->mv_rule_range ` ` lv_coords1 ':' lv_coords2 INTO me->mv_rule_range.
    ELSE.
      CONCATENATE me->mv_rule_range ` ` lv_coords1  INTO me->mv_rule_range.
    ENDIF.
    SHIFT me->mv_rule_range LEFT DELETING LEADING space.

  ENDMETHOD.


  METHOD constructor.

    DATA: ls_iconset TYPE zexcel_conditional_iconset.
    ls_iconset-iconset     = zcl_yy_xls_style_cond=>c_iconset_3trafficlights.
    ls_iconset-cfvo1_type  = zcl_yy_xls_style_cond=>c_cfvo_type_percent.
    ls_iconset-cfvo1_value = '0'.
    ls_iconset-cfvo2_type  = zcl_yy_xls_style_cond=>c_cfvo_type_percent.
    ls_iconset-cfvo2_value = '20'.
    ls_iconset-cfvo3_type  = zcl_yy_xls_style_cond=>c_cfvo_type_percent.
    ls_iconset-cfvo3_value = '40'.
    ls_iconset-cfvo4_type  = zcl_yy_xls_style_cond=>c_cfvo_type_percent.
    ls_iconset-cfvo4_value = '60'.
    ls_iconset-cfvo5_type  = zcl_yy_xls_style_cond=>c_cfvo_type_percent.
    ls_iconset-cfvo5_value = '80'.


    me->rule          = zcl_yy_xls_style_cond=>c_rule_none.
    me->mode_iconset  = ls_iconset.
    me->priority      = 1.

* inizialize dimension range
    me->mv_rule_range     = ip_dimension_range.

    IF ip_guid IS NOT INITIAL.
      me->guid = ip_guid.
    ELSE.
      me->guid = zcl_yy_xls_obsolete_func_wrap=>guid_create( ).
    ENDIF.

  ENDMETHOD.


  METHOD factory_cond_style_iconset.

  ENDMETHOD.


  METHOD get_dimension_range.

    ep_dimension_range = me->mv_rule_range.

  ENDMETHOD.


  METHOD get_guid.
    ep_guid = me->guid.
  ENDMETHOD.


  METHOD set_range.

    CLEAR: me->mv_rule_range.

    me->add_range( ip_start_row    = ip_start_row
                   ip_start_column = ip_start_column
                   ip_stop_row     = ip_stop_row
                   ip_stop_column  = ip_stop_column ).

  ENDMETHOD.
ENDCLASS.
