********************************************************************************
*&---------------------------------------------------------------------*
*&--> Author  :  ZONBAS--(ZAFER ONBAŞ / Itelligence TR)----------------*
*&--> Date    :  09.11.2015 10:14:38-----------------------------------*
*&--> Title   :  OLE2_OBJECT Utility Class-----------------------------*
*&---------------------------------------------------------------------*
********************************************************************************
* The MIT License (MIT)
*
* Copyright (c) 2017 Zafer Onbaş
*
* Permission is hereby granted, free of charge, to any person obtaining a copy
* of this software and associated documentation files (the "Software"), to deal
* in the Software without restriction, including without limitation the rights
* to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
* copies of the Software, and to permit persons to whom the Software is
* furnished to do so, subject to the following conditions:
*
* The above copyright notice and this permission notice shall be included in all
* copies or substantial portions of the Software.
*
* THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
* IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
* FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
* AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
* LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
* OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
* SOFTWARE.
********************************************************************************

CLASS zcl_ole2_object DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

  PUBLIC SECTION.

    TYPES:
      BEGIN OF gty_st_string,
        val TYPE string,
      END OF   gty_st_string .
    TYPES:
      BEGIN OF gty_st_char255,
        val TYPE char255,
      END OF   gty_st_char255 .
    TYPES:
      gty_tt_char255 TYPE STANDARD TABLE OF gty_st_char255 .
    TYPES:
      gty_tt_string TYPE STANDARD TABLE OF gty_st_string .
    TYPES:
      gty_tt_char1500 TYPE STANDARD TABLE OF char1500 .

    CONSTANTS c_col_black TYPE i VALUE 0 ##NO_TEXT.
    CONSTANTS c_col_white TYPE i VALUE 16777215 ##NO_TEXT.
    CONSTANTS c_col_red TYPE i VALUE 255 ##NO_TEXT.
    CONSTANTS c_col_green TYPE i VALUE 5287936 ##NO_TEXT.
    CONSTANTS c_col_blue TYPE i VALUE 12611584 ##NO_TEXT.
    CONSTANTS c_col_yellow TYPE i VALUE 65535 ##NO_TEXT.
    CONSTANTS c_col_orange TYPE i VALUE 49407 ##NO_TEXT.
    CONSTANTS c_ha_left TYPE i VALUE -4131 ##NO_TEXT.
    CONSTANTS c_a_center TYPE i VALUE -4108 ##NO_TEXT.
    CONSTANTS c_ha_right TYPE i VALUE -4152 ##NO_TEXT.
    CONSTANTS c_va_top TYPE i VALUE -4160 ##NO_TEXT.
    CONSTANTS c_va_bottom TYPE i VALUE -4107 ##NO_TEXT.
    CONSTANTS c_ori_portrait TYPE i VALUE 1 ##NO_TEXT.
    CONSTANTS c_ori_landscape TYPE i VALUE 2 ##NO_TEXT.
    CONSTANTS c_ul_none TYPE i VALUE -4142 ##NO_TEXT.
    CONSTANTS c_ul_single TYPE i VALUE 2 ##NO_TEXT.
    CONSTANTS c_ul_double TYPE i VALUE -4119 ##NO_TEXT.

    METHODS constructor
      IMPORTING
        !i_visible TYPE i DEFAULT 1
        !i_wsname  TYPE char30 DEFAULT 'Sheet1'
        !i_wbname  TYPE char30 DEFAULT 'Book1'
        !i_path    TYPE string OPTIONAL
        !i_mode    TYPE char1 DEFAULT 'C'
      EXCEPTIONS
        empty_path
        missing_filename .
    METHODS add_header_via_fcat
      IMPORTING
        !it_fcat TYPE lvc_t_fcat .
    METHODS add_line
      IMPORTING
        !is_line TYPE any
        !i_fcat  TYPE char1 OPTIONAL .
    METHODS add_line_via_tab
      IMPORTING
        !it_table TYPE gty_tt_string .
    METHODS add_lines_via_tab
      IMPORTING
        !it_table TYPE ANY TABLE
        !i_fcat   TYPE char1 OPTIONAL .
    METHODS set_format
      IMPORTING
        !i_bgnrow TYPE i
        !i_bgncol TYPE i
        !i_endrow TYPE i
        !i_endcol TYPE i
        !i_fname  TYPE string OPTIONAL
        !i_fcolor TYPE i OPTIONAL
        !i_fbold  TYPE i OPTIONAL
        !i_fsize  TYPE dec5_2 OPTIONAL
        !i_fitlc  TYPE i OPTIONAL
        !i_fuline TYPE i OPTIONAL
        !i_bcolor TYPE i OPTIONAL
        !i_merge  TYPE char1 OPTIONAL
        !i_nfdec  TYPE char255 OPTIONAL .
    METHODS set_border
      IMPORTING
        !i_bgnrow TYPE i
        !i_bgncol TYPE i
        !i_endrow TYPE i
        !i_endcol TYPE i
        !i_trbli  TYPE char5 .
    METHODS set_alignment
      IMPORTING
        !i_bgnrow TYPE i
        !i_bgncol TYPE i
        !i_endrow TYPE i
        !i_endcol TYPE i
        !i_halign TYPE i OPTIONAL
        !i_valign TYPE i OPTIONAL
        !i_wraptx TYPE char1 OPTIONAL
        !i_orient TYPE i OPTIONAL .
    METHODS set_autofill
      IMPORTING
        !i_src_bgnrow TYPE i
        !i_src_bgncol TYPE i
        !i_src_endrow TYPE i
        !i_src_endcol TYPE i
        !i_des_bgnrow TYPE i
        !i_des_bgncol TYPE i
        !i_des_endrow TYPE i
        !i_des_endcol TYPE i .
    METHODS set_autofit
      IMPORTING
        !i_col TYPE char10 OPTIONAL
        !i_row TYPE char10 OPTIONAL .
    METHODS set_col_width
      IMPORTING
        !i_col   TYPE char10 OPTIONAL
        !i_width TYPE dec5_2 .
    METHODS set_row_height
      IMPORTING
        !i_row    TYPE char10 OPTIONAL
        !i_height TYPE dec5_2 .
    METHODS set_pagesetup
      IMPORTING
        !i_fittopageswide TYPE i DEFAULT 1
        !i_fittopagestall TYPE i DEFAULT 1
        !i_topmargin      TYPE dec5_2 OPTIONAL
        !i_rightmargin    TYPE dec5_2 OPTIONAL
        !i_bottommargin   TYPE dec5_2 OPTIONAL
        !i_leftmargin     TYPE dec5_2 OPTIONAL
        !i_headermargin   TYPE dec5_2 OPTIONAL
        !i_footermargin   TYPE dec5_2 OPTIONAL
        !i_orientation    TYPE i DEFAULT c_ori_portrait .
    METHODS set_image
      IMPORTING
        !i_path   TYPE string
        !i_left   TYPE i
        !i_top    TYPE i
        !i_width  TYPE i
        !i_height TYPE i .
    METHODS set_value
      IMPORTING
        !i_row   TYPE i
        !i_col   TYPE i
        !i_value TYPE simple .
    METHODS get_value
      IMPORTING
        !i_row         TYPE i
        !i_col         TYPE i
      RETURNING
        VALUE(r_value) TYPE char1500 .
    METHODS get_row_count
      RETURNING
        VALUE(e_count) TYPE i .
    METHODS get_app_object
      RETURNING
        VALUE(r_application) TYPE ole2_object .
    METHODS clipboard_export
      IMPORTING
        !i_row TYPE i
        !i_col TYPE i .
    METHODS save_document
      IMPORTING
        !i_visible     TYPE i DEFAULT 1
        !i_compcheck   TYPE i DEFAULT 0
        !i_replacefile TYPE i DEFAULT 1 .
    METHODS clear_clipboard .
    METHODS free
      IMPORTING
        !i_force TYPE char1 DEFAULT 'X' .
    METHODS print .
  PROTECTED SECTION.
  PRIVATE SECTION.

    TYPES:
      gty_dec_16_13 TYPE p LENGTH 16 DECIMALS 13 .

    CONSTANTS c_rs_r1c1 TYPE i VALUE -4150 ##NO_TEXT.
    CONSTANTS c_rs_a1 TYPE i VALUE 1 ##NO_TEXT.
    CONSTANTS c_ws_minimized TYPE i VALUE -4140 ##NO_TEXT.
    CONSTANTS c_ws_normal TYPE i VALUE -4143 ##NO_TEXT.
    CONSTANTS c_ws_maximized TYPE i VALUE -4137 ##NO_TEXT.
    DATA application TYPE ole2_object .
    DATA workbook TYPE ole2_object .
    DATA workbooks TYPE ole2_object .
    DATA worksheet TYPE ole2_object .
    DATA worksheets TYPE ole2_object .
    DATA path TYPE string .
    DATA gt_data TYPE gty_tt_char1500 .
    DATA gs_data TYPE char1500 .
    DATA t_fcat TYPE lvc_t_fcat .

    METHODS create_object
      IMPORTING
        !i_visible TYPE i
        !i_wsname  TYPE char30
        !i_mode    TYPE char1 DEFAULT 'C' .
    METHODS conv_cm2inch
      IMPORTING
        !i_input        TYPE dec5_2
      RETURNING
        VALUE(r_output) TYPE gty_dec_16_13 .
    METHODS writeline
      IMPORTING
        !io_abap_typedescr TYPE REF TO cl_abap_typedescr
        !iv_val            TYPE any .
ENDCLASS.



CLASS ZCL_OLE2_OBJECT IMPLEMENTATION.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Public Method ZCL_OLE2_OBJECT->ADD_HEADER_VIA_FCAT
* +-------------------------------------------------------------------------------------------------+
* | [--->] IT_FCAT                        TYPE        LVC_T_FCAT
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD add_header_via_fcat.
************************************************************************
* @Importing@  IT_FCAT  -> Liste görüntüleyici kontrolü için alan kataloğu
************************************************************************
    DATA: ls_fcat  TYPE lvc_s_fcat,
          lt_lines TYPE gty_tt_string,
          ls_lines TYPE gty_st_string.
    " fcat
    t_fcat = it_fcat.
    "*-
    LOOP AT it_fcat INTO ls_fcat.
      IF ls_fcat-reptext IS NOT INITIAL.
        ls_lines-val = ls_fcat-reptext.
        APPEND ls_lines TO lt_lines.
      ELSE.
        ls_lines-val = ls_fcat-coltext.
        APPEND ls_lines TO lt_lines.
      ENDIF.
      CLEAR: ls_lines.
    ENDLOOP.
    " Add line
    me->add_line_via_tab( it_table = lt_lines ).
  ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Public Method ZCL_OLE2_OBJECT->ADD_LINE
* +-------------------------------------------------------------------------------------------------+
* | [--->] IS_LINE                        TYPE        ANY
* | [--->] I_FCAT                         TYPE        CHAR1(optional)
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD add_line.
************************************************************************
* @Importing@	 IS_LINE  -> Line
* @Importing@	 I_FCAT   -> Using field catalog match?
************************************************************************
    DATA: lo_abap_typedescr TYPE REF TO cl_abap_typedescr,
          lv_counter        TYPE i,
          ls_fcat           TYPE lvc_s_fcat.
    FIELD-SYMBOLS: <lv_val> TYPE any.

    CLEAR gs_data.
    " fcat
    IF i_fcat EQ 'X'.
      LOOP AT t_fcat INTO ls_fcat.
        ASSIGN COMPONENT ls_fcat-fieldname OF STRUCTURE is_line TO <lv_val>.
        IF sy-subrc IS NOT INITIAL.
          CONTINUE.
        ENDIF.
        " getType
        lo_abap_typedescr = cl_abap_typedescr=>describe_by_data( p_data = <lv_val> ).
        " Write line
        me->writeline(
              EXPORTING
                io_abap_typedescr = lo_abap_typedescr " Runtime Type Services
                iv_val            = <lv_val> " Value
            ).
      ENDLOOP.
    ELSE.
      DO.
        " check?
        ADD 1 TO lv_counter.
        ASSIGN COMPONENT lv_counter OF STRUCTURE is_line TO <lv_val>.
        IF sy-subrc NE 0.
          EXIT.
        ENDIF.
        " getType
        lo_abap_typedescr = cl_abap_typedescr=>describe_by_data( p_data = <lv_val> ).
        " Write line
        me->writeline(
              EXPORTING
                io_abap_typedescr = lo_abap_typedescr " Runtime Type Services
                iv_val            = <lv_val> " Value
            ).
      ENDDO.
    ENDIF.
    "*-
    SHIFT gs_data BY 1 PLACES LEFT.
    APPEND gs_data TO gt_data.
    CLEAR  gs_data.
  ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Public Method ZCL_OLE2_OBJECT->ADD_LINES_VIA_TAB
* +-------------------------------------------------------------------------------------------------+
* | [--->] IT_TABLE                       TYPE        ANY TABLE
* | [--->] I_FCAT                         TYPE        CHAR1(optional)
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD add_lines_via_tab.
************************************************************************
* @Importing@  IT_TABLE  -> Table
* @Importing@  I_FCAT    -> Using field catalog match?
************************************************************************
    FIELD-SYMBOLS: <ls_line> TYPE any.
    "*-
    LOOP AT it_table ASSIGNING <ls_line>.
      me->add_line(
        EXPORTING
          is_line = <ls_line>
          i_fcat  = i_fcat " Using field catalog match?
      ).
    ENDLOOP.
  ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Public Method ZCL_OLE2_OBJECT->ADD_LINE_VIA_TAB
* +-------------------------------------------------------------------------------------------------+
* | [--->] IT_TABLE                       TYPE        GTY_TT_STRING
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD add_line_via_tab.
************************************************************************
* @Importing@  IT_TABLE  -> Table with Strings
************************************************************************
    DATA: ls_string TYPE gty_st_string.
    "*-
    CLEAR gs_data.
    LOOP AT it_table INTO ls_string.
      gs_data = gs_data && cl_abap_char_utilities=>horizontal_tab && ls_string-val.
    ENDLOOP.
    "*-
    SHIFT gs_data BY 1 PLACES LEFT.
    APPEND gs_data TO gt_data.
    CLEAR  gs_data.
  ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Public Method ZCL_OLE2_OBJECT->CLEAR_CLIPBOARD
* +-------------------------------------------------------------------------------------------------+
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD clear_clipboard.
    CLEAR gt_data.
  ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Public Method ZCL_OLE2_OBJECT->CLIPBOARD_EXPORT
* +-------------------------------------------------------------------------------------------------+
* | [--->] I_ROW                          TYPE        I
* | [--->] I_COL                          TYPE        I
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD clipboard_export.
************************************************************************
* @Importing@  I_ROW  -> Row
* @Importing@  I_COL  -> Column
************************************************************************
    DATA: lv_rc    TYPE i,
          ls_cells TYPE ole2_object.
    " Flush
    cl_gui_cfw=>flush(
      EXCEPTIONS
        cntl_system_error = 1
        cntl_error        = 2
        OTHERS            = 3
    ).
    IF sy-subrc IS NOT INITIAL.
      MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
                 WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
    ENDIF.
    " Export
    cl_gui_frontend_services=>clipboard_export(
      IMPORTING
        data                 = gt_data
      CHANGING
        rc                   = lv_rc
      EXCEPTIONS
        cntl_error           = 1
        error_no_gui         = 2
        not_supported_by_gui = 3
        no_authority         = 4
        OTHERS               = 5
    ).
    IF sy-subrc IS NOT INITIAL.
      MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
                 WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
    ENDIF.
    " Paste
    CALL METHOD OF worksheet 'Cells' = ls_cells
      EXPORTING
        #1 = i_row
        #2 = i_col.
    CALL METHOD OF ls_cells 'Select'.
    CALL METHOD OF worksheet 'Paste'.
  ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Public Method ZCL_OLE2_OBJECT->CONSTRUCTOR
* +-------------------------------------------------------------------------------------------------+
* | [--->] I_VISIBLE                      TYPE        I (default =1)
* | [--->] I_WSNAME                       TYPE        CHAR30 (default ='Sheet1')
* | [--->] I_WBNAME                       TYPE        CHAR30 (default ='Book1')
* | [--->] I_PATH                         TYPE        STRING(optional)
* | [--->] I_MODE                         TYPE        CHAR1 (default ='C')
* | [EXC!] EMPTY_PATH
* | [EXC!] MISSING_FILENAME
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD constructor.
************************************************************************
* @Importing@  I_VISIBLE         -> Visible? (1/0)
* @Importing@  I_WSNAME          -> Worksheet name
* @Importing@  I_WBNAME          -> Workbook name (Filename)
* @Importing@  I_PATH            -> Path for Open/Create file
* @Importing@  I_MODE            -> O:Open / C:Create
* @Exception@  EMPTY_PATH        -> Dosya/Klasör yolu seçilmedi
* @Exception@  MISSING_FILENAME  -> Açılacak dosyanın tam yolu seçilmedi
************************************************************************
    DATA: lv_visible TYPE i,
          lv_wsname  TYPE text30,
          lv_wbname  TYPE text30,
          lv_title   TYPE string,
          lt_filetab TYPE filetable,
          lv_rc      TYPE i.
    " Path
    IF i_path IS INITIAL.
      CASE i_mode.
        WHEN 'O'. " Open
          cl_gui_frontend_services=>file_open_dialog(
            EXPORTING
              default_extension       = 'xls'            " Default Extension
              file_filter             = '(*.xls)|*.xls|' " File Extension Filter String
              initial_directory       = 'C:\'            " Initial Directory
              multiselection          =  ''              " Multiple selections poss.
            CHANGING
              file_table              = lt_filetab       " Table Holding Selected Files
              rc                      = lv_rc            " Return Code, Number of Files or -1 If Error Occurred
            EXCEPTIONS
              file_open_dialog_failed = 1
              cntl_error              = 2
              error_no_gui            = 3
              not_supported_by_gui    = 4
              OTHERS                  = 5
          ).
          IF sy-subrc <> 0.
            MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
                       WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
          ELSEIF lv_rc EQ 1.
            path = lt_filetab[ 1 ]-filename.
          ENDIF.
        WHEN 'C'. " Create
          cl_gui_frontend_services=>directory_browse(
            EXPORTING
              initial_folder       = 'C:\'     " Start Browsing Here
            CHANGING
              selected_folder      = path      " Folder Path Selected By User
            EXCEPTIONS
              cntl_error           = 1
              error_no_gui         = 2
              not_supported_by_gui = 3
              OTHERS               = 4
          ).
          IF sy-subrc <> 0.
            MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
                       WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
          ENDIF.
      ENDCASE.
    ELSE.
      path = i_path.
    ENDIF.
    " check path?
    IF path IS INITIAL.
      MESSAGE 'File/Folder path is not selected'(001) TYPE 'E' RAISING empty_path.
    ENDIF.
    " check if there is a filename in path or not?
    FIND ALL OCCURRENCES OF
      REGEX '[\w,\s-]+\.[A-Za-z]{3,4}$' IN path
      MATCH COUNT  sy-tabix.
    IF sy-subrc IS NOT INITIAL.
      IF i_mode EQ 'O'.
        MESSAGE 'Filename is missing in path'(002) TYPE 'E' RAISING missing_filename.
      ENDIF.
      " Workbook name
      IF i_wbname IS INITIAL.
        lv_wbname = 'Book1'.
      ELSE.
        lv_wbname = i_wbname.
      ENDIF.
      path = path && '\' && lv_wbname && '.xls'.
      REPLACE ALL OCCURRENCES OF '\\' IN path WITH '\'.
    ENDIF.
    " visible?
    IF i_visible NE 0 AND
       i_visible NE 1.
      lv_visible = 0.
    ELSE.
      lv_visible = i_visible.
    ENDIF.
    " Worksheet name
    IF i_wsname IS INITIAL.
      lv_wsname = 'Sheet1'.
    ELSE.
      lv_wsname = i_wsname.
    ENDIF.
    " Create object
    me->create_object(
      EXPORTING
        i_visible = lv_visible  " Visible?
        i_wsname  = lv_wsname   " Worksheet Name
        i_mode    = i_mode      " O:Open / C:Create
    ).
  ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Private Method ZCL_OLE2_OBJECT->CONV_CM2INCH
* +-------------------------------------------------------------------------------------------------+
* | [--->] I_INPUT                        TYPE        DEC5_2
* | [<-()] R_OUTPUT                       TYPE        GTY_DEC_16_13
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD conv_cm2inch.
************************************************************************
* @Importing@  I_INPUT            -> Input value (cm)
* @Returning@  value( R_OUTPUT )  -> Output value (inch)
************************************************************************
    " Measurement unit conversion
    CALL FUNCTION 'UNIT_CONVERSION_SIMPLE'
      EXPORTING
        input                = i_input  " Input Value
        unit_in              = 'CM'     " Unit of input value
        unit_out             = 'IN'     " Unit of output value
      IMPORTING
        output               = r_output " Output value
      EXCEPTIONS
        conversion_not_found = 1
        division_by_zero     = 2
        input_invalid        = 3
        output_invalid       = 4
        overflow             = 5
        type_invalid         = 6
        units_missing        = 7
        unit_in_not_found    = 8
        unit_out_not_found   = 9
        OTHERS               = 10.
    IF sy-subrc IS NOT INITIAL.
      r_output = 0.
    ENDIF.
  ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Private Method ZCL_OLE2_OBJECT->CREATE_OBJECT
* +-------------------------------------------------------------------------------------------------+
* | [--->] I_VISIBLE                      TYPE        I
* | [--->] I_WSNAME                       TYPE        CHAR30
* | [--->] I_MODE                         TYPE        CHAR1 (default ='C')
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD create_object.
************************************************************************
* @Importing@  I_VISIBLE  -> Visible?
* @Importing@  I_WSNAME   -> Worksheet name
* @Importing@  I_MODE     -> O:Open / C:Create
************************************************************************
    " Application
    CREATE OBJECT application 'Excel.Application'.
    " Reference Style
    SET PROPERTY OF application 'ReferenceStyle' = c_rs_r1c1.
    " Window State
    SET PROPERTY OF application 'WindowState' = c_ws_maximized.
    " Workbooks
    CALL METHOD OF application 'Workbooks' = workbooks.
    IF i_mode EQ 'O'.
      " Open workbook
      CALL METHOD OF workbooks 'Open' = workbook
        EXPORTING
          #1 = path.
      " Save
      me->save_document(
        EXPORTING
          i_visible     = i_visible
          i_compcheck   = 0
          i_replacefile = 1
      ).
    ELSE.
      " Create workbook
      CALL METHOD OF workbooks 'Add' = workbook.
    ENDIF.
    " Visible?
    SET PROPERTY OF application 'Visible' = i_visible.
    " Worksheet
    GET PROPERTY OF application 'ActiveSheet' = worksheet.
    SET PROPERTY OF worksheet 'Name' = i_wsname.
  ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Public Method ZCL_OLE2_OBJECT->FREE
* +-------------------------------------------------------------------------------------------------+
* | [--->] I_FORCE                        TYPE        CHAR1 (default ='X')
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD free.
************************************************************************
* @Importing@  I_FORCE  -> Force close
************************************************************************
    " Release & Free object
    IF i_force EQ 'X'.
      SET PROPERTY OF application 'DisplayAlerts ' = 0.
    ENDIF.
    "*-
    CALL METHOD OF application 'Quit'.
    FREE OBJECT application.
  ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Public Method ZCL_OLE2_OBJECT->GET_APP_OBJECT
* +-------------------------------------------------------------------------------------------------+
* | [<-()] R_APPLICATION                  TYPE        OLE2_OBJECT
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD get_app_object.
************************************************************************
* @Returning@  value( R_APPLICATION )  -> Application object
************************************************************************
    r_application = application.
  ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Public Method ZCL_OLE2_OBJECT->GET_ROW_COUNT
* +-------------------------------------------------------------------------------------------------+
* | [<-()] E_COUNT                        TYPE        I
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD get_row_count.
************************************************************************
* @Returning@  value( E_COUNT )  -> Row Count
************************************************************************
    " Clipboard table row count
    e_count = lines( gt_data ).
  ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Public Method ZCL_OLE2_OBJECT->GET_VALUE
* +-------------------------------------------------------------------------------------------------+
* | [--->] I_ROW                          TYPE        I
* | [--->] I_COL                          TYPE        I
* | [<-()] R_VALUE                        TYPE        CHAR1500
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD get_value.
************************************************************************
* @Importing@  I_ROW    -> Row
* @Importing@  I_COL    -> Column
* @Returning@  R_VALUE  -> Value
************************************************************************
    DATA: ls_cells TYPE ole2_object.
    "*-
    CALL METHOD OF worksheet 'Cells' = ls_cells
      EXPORTING
        #1 = i_row
        #2 = i_col.
    "*-
    GET PROPERTY OF ls_cells 'Value' = r_value.
  ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Public Method ZCL_OLE2_OBJECT->PRINT
* +-------------------------------------------------------------------------------------------------+
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD print.
    " Print
    CALL METHOD OF worksheet 'PrintOut'.
  ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Public Method ZCL_OLE2_OBJECT->SAVE_DOCUMENT
* +-------------------------------------------------------------------------------------------------+
* | [--->] I_VISIBLE                      TYPE        I (default =1)
* | [--->] I_COMPCHECK                    TYPE        I (default =0)
* | [--->] I_REPLACEFILE                  TYPE        I (default =1)
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD save_document.
************************************************************************
* @Importing@  I_VISIBLE      -> Show document after save (if created invisible)
* @Importing@  I_COMPCHECK    -> Compability check?
* @Importing@  I_REPLACEFILE  -> Replace file directly if found (no popup)
************************************************************************
    DATA: ls_cells    TYPE ole2_object,
          ls_activewb TYPE ole2_object.
    " Force window state to get maximized ( min->max )
    SET PROPERTY OF application 'WindowState' = c_ws_minimized.
    SET PROPERTY OF application 'WindowState' = c_ws_maximized.
    " Alert Display
    IF i_replacefile EQ 0.
      SET PROPERTY OF application 'DisplayAlerts ' = 1.
    ELSE.
      SET PROPERTY OF application 'DisplayAlerts ' = 0.
    ENDIF.
    " Compability check?
    CALL METHOD OF application 'ActiveWorkbook' = ls_activewb.
    SET PROPERTY OF ls_activewb 'CheckCompatibility' = i_compcheck.
    " Set cursor to top
    CALL METHOD OF worksheet 'Cells' = ls_cells
      EXPORTING
        #1 = 1
        #2 = 1.
    CALL METHOD OF ls_cells 'Select'.
    " Visible?
    SET PROPERTY OF application 'Visible' = i_visible.
    " SaveAs
    CALL METHOD OF workbook 'SaveAs'
      EXPORTING
        #1 = path.
    " Alert Display (back to default setting)
    IF i_replacefile NE 0.
      SET PROPERTY OF application 'DisplayAlerts ' = 1.
    ENDIF.
  ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Public Method ZCL_OLE2_OBJECT->SET_ALIGNMENT
* +-------------------------------------------------------------------------------------------------+
* | [--->] I_BGNROW                       TYPE        I
* | [--->] I_BGNCOL                       TYPE        I
* | [--->] I_ENDROW                       TYPE        I
* | [--->] I_ENDCOL                       TYPE        I
* | [--->] I_HALIGN                       TYPE        I(optional)
* | [--->] I_VALIGN                       TYPE        I(optional)
* | [--->] I_WRAPTX                       TYPE        CHAR1(optional)
* | [--->] I_ORIENT                       TYPE        I(optional)
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD set_alignment.
************************************************************************
* @Importing@  I_BGNROW -> Begin row
* @Importing@  I_BGNCOL -> Begin column
* @Importing@  I_ENDROW -> End row
* @Importing@  I_ENDCOL -> End column
* @Importing@  I_HALIGN -> Horizontal align
* @Importing@  I_VALIGN -> Vertical align
* @Importing@  I_WRAPTX -> Wrap-text
* @Importing@  I_ORIENT -> Text-Orientation (-90<=..<=90)
************************************************************************
    DATA: ls_cellbgn TYPE ole2_object,
          ls_cellend TYPE ole2_object,
          ls_range   TYPE ole2_object.
    " Select All
    IF i_bgnrow EQ 0 AND
       i_bgncol EQ 0 AND
       i_endrow EQ 0 AND
       i_endcol EQ 0.
      CALL METHOD OF worksheet 'Cells' = ls_range.
    ELSE.
      " Select Range
      CALL METHOD OF worksheet 'Cells' = ls_cellbgn
        EXPORTING
          #1 = i_bgnrow
          #2 = i_bgncol.
      CALL METHOD OF worksheet 'Cells' = ls_cellend
        EXPORTING
          #1 = i_endrow
          #2 = i_endcol.
      CALL METHOD OF worksheet 'Range' = ls_range
        EXPORTING
          #1 = ls_cellbgn
          #2 = ls_cellend.
    ENDIF.
    CALL METHOD OF ls_range 'Select'.
    " Horizontal
    IF i_halign IS SUPPLIED.
      SET PROPERTY OF ls_range 'HorizontalAlignment' = i_halign.
    ENDIF.
    " Vertical
    IF i_valign IS SUPPLIED.
      SET PROPERTY OF ls_range 'VerticalAlignment'   = i_valign.
    ENDIF.
    " Wrap text
    IF i_wraptx EQ 'X'.
      SET PROPERTY OF ls_range 'WrapText' = 1.
    ENDIF.
    " Text-Orientation
    IF i_orient IS SUPPLIED AND
       i_orient GE -90 AND
       i_orient LE 90.
      SET PROPERTY OF ls_range 'Orientation' = i_orient.
    ENDIF.
  ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Public Method ZCL_OLE2_OBJECT->SET_AUTOFILL
* +-------------------------------------------------------------------------------------------------+
* | [--->] I_SRC_BGNROW                   TYPE        I
* | [--->] I_SRC_BGNCOL                   TYPE        I
* | [--->] I_SRC_ENDROW                   TYPE        I
* | [--->] I_SRC_ENDCOL                   TYPE        I
* | [--->] I_DES_BGNROW                   TYPE        I
* | [--->] I_DES_BGNCOL                   TYPE        I
* | [--->] I_DES_ENDROW                   TYPE        I
* | [--->] I_DES_ENDCOL                   TYPE        I
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD set_autofill.
************************************************************************
* @Importing@  I_SRC_BGNROW  -> Source Begin row
* @Importing@  I_SRC_BGNCOL  -> Source Begin column
* @Importing@  I_SRC_ENDROW  -> Source End row
* @Importing@  I_SRC_ENDCOL  -> Source End column
* @Importing@  I_DES_BGNROW  -> Destination Begin row
* @Importing@  I_DES_BGNCOL  -> Destination Begin column
* @Importing@  I_DES_ENDROW  -> Destination End row
* @Importing@  I_DES_ENDCOL  -> Destination End column
************************************************************************
    DATA: ls_cellbgn   TYPE ole2_object,
          ls_cellend   TYPE ole2_object,
          ls_src_range TYPE ole2_object,
          ls_des_range TYPE ole2_object.
    " Select Source Range
    CALL METHOD OF worksheet 'Cells' = ls_cellbgn
      EXPORTING
        #1 = i_src_bgnrow
        #2 = i_src_bgncol.
    CALL METHOD OF worksheet 'Cells' = ls_cellend
      EXPORTING
        #1 = i_src_endrow
        #2 = i_src_endcol.
    CALL METHOD OF worksheet 'Range' = ls_src_range
      EXPORTING
        #1 = ls_cellbgn
        #2 = ls_cellend.
    " Select Destination Range
    CALL METHOD OF worksheet 'Cells' = ls_cellbgn
      EXPORTING
        #1 = i_des_bgnrow
        #2 = i_des_bgncol.
    CALL METHOD OF worksheet 'Cells' = ls_cellend
      EXPORTING
        #1 = i_des_endrow
        #2 = i_des_endcol.
    CALL METHOD OF worksheet 'Range' = ls_des_range
      EXPORTING
        #1 = ls_cellbgn
        #2 = ls_cellend.
    " AutoFill
    CALL METHOD OF ls_src_range 'AutoFill'
      EXPORTING
        #1 = ls_des_range
        #2 = 0. " xlFillDefault
  ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Public Method ZCL_OLE2_OBJECT->SET_AUTOFIT
* +-------------------------------------------------------------------------------------------------+
* | [--->] I_COL                          TYPE        CHAR10(optional)
* | [--->] I_ROW                          TYPE        CHAR10(optional)
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD set_autofit.
************************************************************************
* @Importing@  I_COL  -> Column (ex:'2','B','A:DZ',''=All)
* @Importing@  I_ROW  -> Row (ex:'2','1:5',''=All)
************************************************************************
    DATA: ls_columns TYPE ole2_object,
          ls_rows    TYPE ole2_object.
    " Column
    IF i_col IS NOT INITIAL.
      CALL METHOD OF worksheet 'Columns' = ls_columns
          EXPORTING
            #1 = i_col.
    ELSE.
      CALL METHOD OF worksheet 'Columns' = ls_columns.
    ENDIF.
    IF ls_columns IS NOT INITIAL.
      CALL METHOD OF ls_columns 'Autofit'.
    ENDIF.
    " Row
    IF i_row IS NOT INITIAL.
      CALL METHOD OF worksheet 'Rows' = ls_rows
          EXPORTING
            #1 = i_row.
    ELSE.
      CALL METHOD OF worksheet 'Rows' = ls_rows.
    ENDIF.
    IF ls_rows IS NOT INITIAL.
      CALL METHOD OF ls_rows 'Autofit'.
    ENDIF.
  ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Public Method ZCL_OLE2_OBJECT->SET_BORDER
* +-------------------------------------------------------------------------------------------------+
* | [--->] I_BGNROW                       TYPE        I
* | [--->] I_BGNCOL                       TYPE        I
* | [--->] I_ENDROW                       TYPE        I
* | [--->] I_ENDCOL                       TYPE        I
* | [--->] I_TRBLI                        TYPE        CHAR5
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD set_border.
************************************************************************
* @Importing@  I_BGNROW  -> Begin row
* @Importing@  I_BGNCOL  -> Begin column
* @Importing@  I_ENDROW  -> End row
* @Importing@  I_ENDCOL  -> End column
* @Importing@  I_TRBLI   -> Top-Right-Bottom-Left-Inside (ex:01010)
************************************************************************
    DATA: ls_cellbgn TYPE ole2_object,
          ls_cellend TYPE ole2_object,
          ls_range   TYPE ole2_object,
          ls_borders TYPE ole2_object.
    " Select All
    IF i_bgnrow EQ 0 AND
       i_bgncol EQ 0 AND
       i_endrow EQ 0 AND
       i_endcol EQ 0.
      CALL METHOD OF worksheet 'Cells' = ls_range.
    ELSE.
      " Select Range
      CALL METHOD OF worksheet 'Cells' = ls_cellbgn
        EXPORTING
          #1 = i_bgnrow
          #2 = i_bgncol.
      CALL METHOD OF worksheet 'Cells' = ls_cellend
        EXPORTING
          #1 = i_endrow
          #2 = i_endcol.
      CALL METHOD OF worksheet 'Range' = ls_range
        EXPORTING
          #1 = ls_cellbgn
          #2 = ls_cellend.
    ENDIF.
    CALL METHOD OF ls_range 'Select'.
    " Top
    IF i_trbli(1) NE ''.
      CALL METHOD OF ls_range 'Borders' = ls_borders
        EXPORTING
          #1 = '8'. " xlEdgeTop
      SET PROPERTY OF ls_borders 'LineStyle' = i_trbli(1).
    ENDIF.
    " Right
    IF i_trbli+1(1) NE ''.
      CALL METHOD OF ls_range 'Borders' = ls_borders
        EXPORTING
          #1 = '10'. "xlEdgeRight
      SET PROPERTY OF ls_borders 'LineStyle' = i_trbli+1(1).
    ENDIF.
    " Bottom
    IF i_trbli+2(1) NE ''.
      CALL METHOD OF ls_range 'Borders' = ls_borders
        EXPORTING
          #1 = '9'. "xlEdgeBottom
      SET PROPERTY OF ls_borders 'LineStyle' = i_trbli+2(1).
    ENDIF.
    " Left
    IF i_trbli+3(1) NE ''.
      CALL METHOD OF ls_range 'Borders' = ls_borders
        EXPORTING
          #1 = '7'. "xlEdgeLeft
      SET PROPERTY OF ls_borders 'LineStyle' = i_trbli+3(1).
    ENDIF.
    " Inside
    IF i_trbli+4(1) NE ''.
      " Vertical
      CALL METHOD OF ls_range 'Borders' = ls_borders
        EXPORTING
          #1 = '11'. "xlInsideVertical
      SET PROPERTY OF ls_borders 'LineStyle' = i_trbli+4(1).
      " Horizontal
      CALL METHOD OF ls_range 'Borders' = ls_borders
        EXPORTING
          #1 = '12'. "xlInsideHorizontal
      SET PROPERTY OF ls_borders 'LineStyle' = i_trbli+4(1).
    ENDIF.
  ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Public Method ZCL_OLE2_OBJECT->SET_COL_WIDTH
* +-------------------------------------------------------------------------------------------------+
* | [--->] I_COL                          TYPE        CHAR10(optional)
* | [--->] I_WIDTH                        TYPE        DEC5_2
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD set_col_width.
************************************************************************
* @Importing@  I_COL    -> Column (ex:'2','B','A:DZ',''=All)
* @Importing@  I_WIDTH  -> Width
************************************************************************
    DATA: ls_columns TYPE ole2_object,
          lv_i       TYPE i.
    FIELD-SYMBOLS <lv_v> TYPE any.
    "*-
    IF i_col IS NOT INITIAL.
      TRY.
          " Numeric conv. for single column
          lv_i = i_col.
          ASSIGN lv_i TO <lv_v>.
        CATCH cx_sy_conversion_no_number.
          ASSIGN i_col TO <lv_v>.
      ENDTRY.
      CALL METHOD OF worksheet 'Columns' = ls_columns
          EXPORTING
            #1 = <lv_v>.
    ELSE.
      CALL METHOD OF worksheet 'Columns' = ls_columns.
    ENDIF.
    "*-
    SET PROPERTY OF ls_columns 'ColumnWidth' = i_width.
  ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Public Method ZCL_OLE2_OBJECT->SET_FORMAT
* +-------------------------------------------------------------------------------------------------+
* | [--->] I_BGNROW                       TYPE        I
* | [--->] I_BGNCOL                       TYPE        I
* | [--->] I_ENDROW                       TYPE        I
* | [--->] I_ENDCOL                       TYPE        I
* | [--->] I_FNAME                        TYPE        STRING(optional)
* | [--->] I_FCOLOR                       TYPE        I(optional)
* | [--->] I_FBOLD                        TYPE        I(optional)
* | [--->] I_FSIZE                        TYPE        DEC5_2(optional)
* | [--->] I_FITLC                        TYPE        I(optional)
* | [--->] I_FULINE                       TYPE        I(optional)
* | [--->] I_BCOLOR                       TYPE        I(optional)
* | [--->] I_MERGE                        TYPE        CHAR1(optional)
* | [--->] I_NFDEC                        TYPE        CHAR255(optional)
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD set_format.
************************************************************************
* @Importing@  I_BGNROW  -> Begin row
* @Importing@  I_BGNCOL  -> Begin column
* @Importing@  I_ENDROW  -> End row
* @Importing@  I_ENDCOL  -> End column
* @Importing@  I_FNAME   -> Font name (family)
* @Importing@  I_FCOLOR  -> Font color
* @Importing@  I_FBOLD   -> Font bold
* @Importing@  I_FSIZE   -> Font size
* @Importing@  I_FITLC   -> Font italic
* @Importing@  I_FULINE  -> Font underline
* @Importing@  I_BCOLOR  -> Background color
* @Importing@  I_MERGE   -> Merge Cells
* @Importing@  I_NFDEC   -> Number format (ex:'0.00')
************************************************************************
    DATA: ls_cellbgn  TYPE ole2_object,
          ls_cellend  TYPE ole2_object,
          ls_range    TYPE ole2_object,
          ls_font     TYPE ole2_object,
          ls_interior TYPE ole2_object.
    " Select All
    IF i_bgnrow EQ 0 AND
       i_bgncol EQ 0 AND
       i_endrow EQ 0 AND
       i_endcol EQ 0.
      CALL METHOD OF worksheet 'Cells' = ls_range.
    ELSE.
      " Select Range
      CALL METHOD OF worksheet 'Cells' = ls_cellbgn
        EXPORTING
          #1 = i_bgnrow
          #2 = i_bgncol.
      CALL METHOD OF worksheet 'Cells' = ls_cellend
        EXPORTING
          #1 = i_endrow
          #2 = i_endcol.
      CALL METHOD OF worksheet 'Range' = ls_range
        EXPORTING
          #1 = ls_cellbgn
          #2 = ls_cellend.
    ENDIF.
    CALL METHOD OF ls_range 'Select'.
    GET PROPERTY OF ls_range 'Font' = ls_font.
    " Font Name
    IF i_fname IS SUPPLIED.
      SET PROPERTY OF ls_font 'Name' = i_fname.
    ENDIF.
    " Font Color
    IF i_fcolor IS SUPPLIED.
      SET PROPERTY OF ls_font 'Color' = i_fcolor.
    ENDIF.
    " Font Bold
    IF i_fbold IS SUPPLIED.
      SET PROPERTY OF ls_font 'Bold' = i_fbold.
    ENDIF.
    " Font Size
    IF i_fsize IS SUPPLIED.
      SET PROPERTY OF ls_font 'Size' = i_fsize.
    ENDIF.
    " Font Italic
    IF i_fitlc IS SUPPLIED.
      SET PROPERTY OF ls_font 'Italic' = i_fitlc.
    ENDIF.
    " Font Underline
    IF i_fuline IS SUPPLIED.
      SET PROPERTY OF ls_font 'Underline' = i_fuline.
    ENDIF.
    " Background Color
    IF i_bcolor IS SUPPLIED.
      GET PROPERTY OF ls_range 'Interior' = ls_interior.
      SET PROPERTY OF ls_interior 'Color' = i_bcolor.
    ENDIF.
    " Merge
    IF i_merge IS SUPPLIED.
      CALL METHOD OF ls_range 'Merge'.
    ENDIF.
    " Number Format
    IF i_nfdec IS SUPPLIED.
*    " e.g.
      " NumberFormat = "General"
      " NumberFormat = "0.00","0.000","0" (Decimal)
      " NumberFormat = "#,##0.00 $","#,##0 $" (Currency)
      " NumberFormat = "dd/mm/yyyy;@","dd/mm/yy;@","[$-41F]d mmmm yyyy;@" (Date)
      " NumberFormat = "hh:mm;@","hh:mm:ss;@" (Time)
      " NumberFormat = "0.00%" (Percentage)
      SET PROPERTY OF ls_range 'NumberFormat' = i_nfdec.
    ENDIF.
  ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Public Method ZCL_OLE2_OBJECT->SET_IMAGE
* +-------------------------------------------------------------------------------------------------+
* | [--->] I_PATH                         TYPE        STRING
* | [--->] I_LEFT                         TYPE        I
* | [--->] I_TOP                          TYPE        I
* | [--->] I_WIDTH                        TYPE        I
* | [--->] I_HEIGHT                       TYPE        I
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD set_image.
************************************************************************
* @Importing@  I_PATH    -> Path
* @Importing@  I_LEFT    -> Left
* @Importing@  I_TOP     -> Top
* @Importing@  I_WIDTH   -> Width
* @Importing@  I_HEIGHT  -> Height
************************************************************************
    DATA: ls_shapes TYPE ole2_object.
    "*-
    GET PROPERTY OF worksheet 'Shapes' = ls_shapes.
    CALL METHOD OF ls_shapes 'AddPicture'
      EXPORTING
        #1 = i_path    " Filename
        #2 = '1'       " LinkToFile
        #3 = '1'       " SaveWithDocument
        #4 = i_left    " Left
        #5 = i_top     " Top
        #6 = i_width   " Width
        #7 = i_height. " Height
  ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Public Method ZCL_OLE2_OBJECT->SET_PAGESETUP
* +-------------------------------------------------------------------------------------------------+
* | [--->] I_FITTOPAGESWIDE               TYPE        I (default =1)
* | [--->] I_FITTOPAGESTALL               TYPE        I (default =1)
* | [--->] I_TOPMARGIN                    TYPE        DEC5_2(optional)
* | [--->] I_RIGHTMARGIN                  TYPE        DEC5_2(optional)
* | [--->] I_BOTTOMMARGIN                 TYPE        DEC5_2(optional)
* | [--->] I_LEFTMARGIN                   TYPE        DEC5_2(optional)
* | [--->] I_HEADERMARGIN                 TYPE        DEC5_2(optional)
* | [--->] I_FOOTERMARGIN                 TYPE        DEC5_2(optional)
* | [--->] I_ORIENTATION                  TYPE        I (default =C_ORI_PORTRAIT)
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD set_pagesetup.
************************************************************************
* @Importing@  I_FITTOPAGESWIDE  -> Fit to pages wide
* @Importing@  I_FITTOPAGESTALL  -> Fit to pages tall
* @Importing@  I_TOPMARGIN       -> Top margin (cm)
* @Importing@  I_RIGHTMARGIN     -> Right margin (cm)
* @Importing@  I_BOTTOMMARGIN    -> Bottom margin (cm)
* @Importing@  I_LEFTMARGIN      -> Left margin (cm)
* @Importing@  I_HEADERMARGIN    -> Header margin (cm)
* @Importing@  I_FOOTERMARGIN    -> Footer margin (cm)
* @Importing@  I_ORIENTATION     -> Orientation
************************************************************************
    DATA: lv_p_16_13        TYPE gty_dec_16_13,
          ls_inchestopoints TYPE ole2_object,
          ls_pagesetup      TYPE ole2_object.
    "*-
    SET PROPERTY OF application 'PrintCommunication' = 1.
    GET PROPERTY OF worksheet 'PageSetup' = ls_pagesetup.
    SET PROPERTY OF ls_pagesetup 'PrintArea' = ''.
    SET PROPERTY OF application 'PrintCommunication' = 0.
    " fitToWide
    SET PROPERTY OF ls_pagesetup 'FitToPagesWide' = i_fittopageswide.
    " fitToTall
    SET PROPERTY OF ls_pagesetup 'FitToPagesTall' = i_fittopagestall.
    " Orientation
    SET PROPERTY OF ls_pagesetup 'Orientation' = i_orientation.
    " topMargin
    IF i_topmargin IS SUPPLIED.
      lv_p_16_13 = conv_cm2inch( i_input = i_topmargin ).
      CALL METHOD OF application 'InchesToPoints' = ls_inchestopoints
        EXPORTING
          #1 = lv_p_16_13.
      SET PROPERTY OF ls_pagesetup 'TopMargin' = ls_inchestopoints.
    ENDIF.
    " rightMargin
    IF i_rightmargin IS SUPPLIED.
      lv_p_16_13 = conv_cm2inch( i_input = i_rightmargin ).
      CALL METHOD OF application 'InchesToPoints' = ls_inchestopoints
        EXPORTING
          #1 = lv_p_16_13.
      SET PROPERTY OF ls_pagesetup 'RightMargin' = ls_inchestopoints.
    ENDIF.
    " bottomMargin
    IF i_bottommargin IS SUPPLIED.
      lv_p_16_13 = conv_cm2inch( i_input = i_bottommargin ).
      CALL METHOD OF application 'InchesToPoints' = ls_inchestopoints
        EXPORTING
          #1 = lv_p_16_13.
      SET PROPERTY OF ls_pagesetup 'BottomMargin' = ls_inchestopoints.
    ENDIF.
    " leftMargin
    IF i_leftmargin IS SUPPLIED.
      lv_p_16_13 = conv_cm2inch( i_input = i_leftmargin ).
      CALL METHOD OF application 'InchesToPoints' = ls_inchestopoints
        EXPORTING
          #1 = lv_p_16_13.
      SET PROPERTY OF ls_pagesetup 'LeftMargin' = ls_inchestopoints.
    ENDIF.
    " headerMargin
    IF i_headermargin IS SUPPLIED.
      lv_p_16_13 = conv_cm2inch( i_input = i_headermargin ).
      CALL METHOD OF application 'InchesToPoints' = ls_inchestopoints
        EXPORTING
          #1 = lv_p_16_13.
      SET PROPERTY OF ls_pagesetup 'HeaderMargin' = ls_inchestopoints.
    ENDIF.
    " footerMargin
    IF i_footermargin IS SUPPLIED.
      lv_p_16_13 = conv_cm2inch( i_input = i_footermargin ).
      CALL METHOD OF application 'InchesToPoints' = ls_inchestopoints
        EXPORTING
          #1 = lv_p_16_13.
      SET PROPERTY OF ls_pagesetup 'FooterMargin' = ls_inchestopoints.
    ENDIF.
    "*-
    SET PROPERTY OF application 'PrintCommunication' = 1.
  ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Public Method ZCL_OLE2_OBJECT->SET_ROW_HEIGHT
* +-------------------------------------------------------------------------------------------------+
* | [--->] I_ROW                          TYPE        CHAR10(optional)
* | [--->] I_HEIGHT                       TYPE        DEC5_2
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD set_row_height.
************************************************************************
* @Importing@  I_ROW     -> Row (ex:'2','1:5',''=All)
* @Importing@  I_HEIGHT  -> Height
************************************************************************
    DATA: ls_rows TYPE ole2_object.
    "*-
    IF i_row IS NOT INITIAL.
      CALL METHOD OF worksheet 'Rows' = ls_rows
          EXPORTING
            #1 = i_row.
    ELSE.
      CALL METHOD OF worksheet 'Rows' = ls_rows.
    ENDIF.
    "*-
    SET PROPERTY OF ls_rows 'RowHeight' = i_height.
  ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Public Method ZCL_OLE2_OBJECT->SET_VALUE
* +-------------------------------------------------------------------------------------------------+
* | [--->] I_ROW                          TYPE        I
* | [--->] I_COL                          TYPE        I
* | [--->] I_VALUE                        TYPE        SIMPLE
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD set_value.
************************************************************************
* @Importing@  I_ROW    -> Row
* @Importing@  I_COL    -> Column
* @Importing@  I_VALUE  -> Value
************************************************************************
    DATA: ls_cells TYPE ole2_object.
    "*-
    CALL METHOD OF worksheet 'Cells' = ls_cells
      EXPORTING
        #1 = i_row
        #2 = i_col.
    "*-
    SET PROPERTY OF ls_cells 'Value' = i_value.
  ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Private Method ZCL_OLE2_OBJECT->WRITELINE
* +-------------------------------------------------------------------------------------------------+
* | [--->] IO_ABAP_TYPEDESCR              TYPE REF TO CL_ABAP_TYPEDESCR
* | [--->] IV_VAL                         TYPE        ANY
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD writeline.
************************************************************************
* @Importing@  IO_ABAP_TYPEDESCR  -> Runtime Type Services
* @Importing@  IV_VAL             -> Value
************************************************************************
    DATA: lv_char TYPE char1500.
    " Write
    CASE io_abap_typedescr->type_kind.
      WHEN io_abap_typedescr->typekind_char.   " C
        WRITE iv_val TO lv_char.
      WHEN io_abap_typedescr->typekind_num OR  " N
           io_abap_typedescr->typekind_int OR  " I
           io_abap_typedescr->typekind_packed. " P.
        IF iv_val IS NOT INITIAL.
          lv_char = CONV char1500( iv_val ).
        ENDIF.
      WHEN io_abap_typedescr->typekind_date.   " D
        IF iv_val IS NOT INITIAL.
          IF iv_val NA sy-abcde.
            WRITE iv_val TO lv_char DD/MM/YYYY.
          ELSE.
            lv_char = CONV char1500( iv_val ).
          ENDIF.
        ENDIF.
      WHEN io_abap_typedescr->typekind_time.   " T
        IF iv_val IS NOT INITIAL.
          IF iv_val NA sy-abcde.
            WRITE iv_val TO lv_char ENVIRONMENT TIME FORMAT.
          ELSE.
            lv_char = CONV char1500( iv_val ).
          ENDIF.
        ENDIF.
      WHEN OTHERS. " Others
        WRITE iv_val TO lv_char.
    ENDCASE.
    gs_data = gs_data && cl_abap_char_utilities=>horizontal_tab && lv_char.
  ENDMETHOD.
ENDCLASS.
