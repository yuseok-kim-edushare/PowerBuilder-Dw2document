$PBExportHeader$dw_exporter.sru

forward
global type dw_exporter from nonvisualobject
end type
end forward

global type dw_exporter from nonvisualobject
event ue_error ( )
end type
global dw_exporter dw_exporter

type variables
// Use the nvo_datawindowexporter instead of direct DotNetObject
nvo_datawindowexporter inv_exporter
end variables

forward prototypes
public function boolean of_initialize ()
public function string of_export_to_excel (datawindow adw_source, string as_output_path, string as_sheet_name)
public function string of_export_to_word (datawindow adw_source, string as_output_path)
private function string of_datawindow_to_json (datawindow adw_source)
private function string of_datawindow_structure_to_json (datawindow adw_source)
private function string of_datawindow_rows_to_json (datawindow adw_source)
private function string of_datawindow_to_virtualgrid_json (datawindow adw_source)
private function string json_escape(string value)
private function string json_null_if_special(string value)
public subroutine of_seterrorhandler (powerobject apo_newhandler, string as_newevent)
end prototypes

public function boolean of_initialize ();
// Initialize the .NET exporter using nvo_datawindowexporter
try
    if IsNull(inv_exporter) or Not IsValid(inv_exporter) then
        inv_exporter = create nvo_datawindowexporter
    end if
    
    return inv_exporter.of_createondemand()
catch (RuntimeError lre_error)
    MessageBox("Initialization Error", "Failed to create exporter: " + lre_error.Text)
    return false
end try
end function

public function string of_export_to_excel (datawindow adw_source, string as_output_path, string as_sheet_name);

// Export DataWindow to Excel
try
    if IsNull(inv_exporter) or Not IsValid(inv_exporter) then
        if not of_initialize() then 
            return "Error: Failed to initialize exporter"
        end if
    end if
    
    // Convert DataWindow to virtual grid JSON data
    string ls_json_data
    ls_json_data = of_datawindow_to_virtualgrid_json(adw_source)
	 
	// export json into text file for debugging
	integer li_file
	li_file = FileOpen("C:\temp\dw2docs-json.txt",LineMode!, Write!)
	FileWrite(li_file,ls_json_data)
	FileClose(li_file)
    
    // Try first deleting any existing file
    FileDelete(as_output_path)
    
    // Export to Excel using the nvo_datawindowexporter
    string ls_result
    ls_result = inv_exporter.of_exporttoexcel(ls_json_data, as_output_path, as_sheet_name)
    
    MessageBox("Debug", "Result: " + ls_result)
    
    return ls_result
catch (RuntimeError lre_error)
    return "Error: " + lre_error.Text
end try
end function

public function string of_export_to_word (datawindow adw_source, string as_output_path);
// Export DataWindow to Word
try
    if IsNull(inv_exporter) or Not IsValid(inv_exporter) then
        if not of_initialize() then 
            return "Error: Failed to initialize exporter"
        end if
    end if
    
    // Convert DataWindow to virtual grid JSON data
    string ls_json_data
    ls_json_data = of_datawindow_to_virtualgrid_json(adw_source)
    
    // Export to Word using the nvo_datawindowexporter
    string ls_result
    ls_result = inv_exporter.of_exporttoword(ls_json_data, as_output_path)
    
    return ls_result
catch (RuntimeError lre_error)
    return "Error: " + lre_error.Text
end try
end function

public subroutine of_seterrorhandler (powerobject apo_newhandler, string as_newevent);
// Forward the error handler setting to nvo_datawindowexporter
if not IsNull(inv_exporter) and IsValid(inv_exporter) then
    inv_exporter.of_seterrorhandler(apo_newhandler, as_newevent)
end if
end subroutine

private function string of_datawindow_to_json (datawindow adw_source);
// Export DataWindow to JSON with structure and data
string ls_json

// Add metadata
ls_json = "{"
ls_json += '"metadata":{'
ls_json += '"name":"' + json_escape(adw_source.Describe("DataWindow.Name")) + '",'
ls_json += '"title":"' + json_escape(adw_source.Describe("DataWindow.Title")) + '"'
ls_json += '},'

// Add structure
ls_json += '"structure":' + of_datawindow_structure_to_json(adw_source) + ','

// Add data rows
ls_json += '"rows":' + of_datawindow_rows_to_json(adw_source)

ls_json += "}"

return ls_json
end function

private function string of_datawindow_structure_to_json (datawindow adw_source);
// Export DataWindow structure (columns, objects, bands) to JSON
string ls_json, ls_objects, ls_objname, ls_type, ls_coltype, ls_expr, ls_band
string ls_columns, ls_objs, ls_bands
boolean lb_first

// --- Columns ---
ls_columns = "["
ls_objects = adw_source.Describe("DataWindow.Objects")
string ls_objlist[]
integer li_pos_start = 1, li_pos_end, li_colcount = 0

// Split objects by tab
do while li_pos_start > 0
    li_pos_end = Pos(ls_objects, "~t", li_pos_start)
    if li_pos_end = 0 then
        ls_objname = Mid(ls_objects, li_pos_start)
        li_pos_start = 0
    else
        ls_objname = Mid(ls_objects, li_pos_start, li_pos_end - li_pos_start)
        li_pos_start = li_pos_end + 1
    end if

    ls_type = adw_source.Describe(ls_objname + ".Type")
    if ls_type = "column" or ls_type = "compute" then
        if li_colcount > 0 then ls_columns += ","
        ls_columns += "{"
        ls_columns += '"name":"' + json_escape(ls_objname) + '",'
        ls_columns += '"type":"' + json_escape(adw_source.Describe(ls_objname + ".ColType")) + '"'
        if ls_type = "compute" then
            ls_expr = adw_source.Describe(ls_objname + ".Expression")
            ls_columns += ',"expression":"' + json_escape(ls_expr) + '"'
        end if
        ls_columns += ',"format":' + json_null_if_special(adw_source.Describe(ls_objname + ".Format")) + '"'
        ls_columns += "}"
        li_colcount++
    end if
loop
ls_columns += "]"

// --- Objects ---
ls_objs = "["
li_pos_start = 1
li_pos_end = 0
integer li_objcount = 0
do while li_pos_start > 0
    li_pos_end = Pos(ls_objects, "~t", li_pos_start)
    if li_pos_end = 0 then
        ls_objname = Mid(ls_objects, li_pos_start)
        li_pos_start = 0
    else
        ls_objname = Mid(ls_objects, li_pos_start, li_pos_end - li_pos_start)
        li_pos_start = li_pos_end + 1
    end if

    ls_type = adw_source.Describe(ls_objname + ".Type")
    if li_objcount > 0 then ls_objs += ","
    ls_objs += "{"
    ls_objs += '"name":"' + json_escape(ls_objname) + '",'
    ls_objs += '"type":"' + json_escape(ls_type) + '"'
    
    // Add position/size if available
    string ls_x, ls_y, ls_w, ls_h, ls_band
    ls_x = adw_source.Describe(ls_objname + ".X")
    ls_y = adw_source.Describe(ls_objname + ".Y")
    ls_w = adw_source.Describe(ls_objname + ".Width")
    ls_h = adw_source.Describe(ls_objname + ".Height")
    ls_band = adw_source.Describe(ls_objname + ".Band")
    
    ls_objs += ',"x":' + json_null_if_special(ls_x)
    ls_objs += ',"y":' + json_null_if_special(ls_y)
    ls_objs += ',"width":' + json_null_if_special(ls_w)
    ls_objs += ',"height":' + json_null_if_special(ls_h)
    ls_objs += ',"band":"' + json_escape(ls_band) + '"'
    
    // Type-specific attributes
    choose case ls_type
        case "text"
            // Text object attributes
            ls_objs += ',"text":"' + json_escape(adw_source.Describe(ls_objname + ".Text")) + '"'
            ls_objs += ',"alignment":"' + json_escape(adw_source.Describe(ls_objname + ".Alignment")) + '"'
            ls_objs += ',"font_face":"' + json_escape(adw_source.Describe(ls_objname + ".Font.Face")) + '"'
            ls_objs += ',"font_height":' + json_null_if_special(adw_source.Describe(ls_objname + ".Font.Height"))
            ls_objs += ',"font_weight":' + json_null_if_special(adw_source.Describe(ls_objname + ".Font.Weight"))
            ls_objs += ',"underline":"' + json_escape(adw_source.Describe(ls_objname + ".Font.Underline")) + '"'
            ls_objs += ',"italic":"' + json_escape(adw_source.Describe(ls_objname + ".Font.Italic")) + '"'
            ls_objs += ',"strikethrough":"' + json_escape(adw_source.Describe(ls_objname + ".Font.Strikethrough")) + '"'
            ls_objs += ',"color":' + json_null_if_special(adw_source.Describe(ls_objname + ".Color"))
            ls_objs += ',"background_color":' + json_null_if_special(adw_source.Describe(ls_objname + ".Background.Color"))
            
        case "line"
            // Line object attributes
            ls_objs += ',"x1":' + json_null_if_special(adw_source.Describe(ls_objname + ".X1"))
            ls_objs += ',"y1":' + json_null_if_special(adw_source.Describe(ls_objname + ".Y1"))
            ls_objs += ',"x2":' + json_null_if_special(adw_source.Describe(ls_objname + ".X2"))
            ls_objs += ',"y2":' + json_null_if_special(adw_source.Describe(ls_objname + ".Y2"))
            ls_objs += ',"pen_width":' + json_null_if_special(adw_source.Describe(ls_objname + ".Pen.Width"))
            ls_objs += ',"pen_color":' + json_null_if_special(adw_source.Describe(ls_objname + ".Pen.Color"))
            ls_objs += ',"pen_style":"' + json_escape(adw_source.Describe(ls_objname + ".Pen.Style")) + '"'
            
        case "compute"
            // Compute object attributes
            ls_objs += ',"expression":"' + json_escape(adw_source.Describe(ls_objname + ".Expression")) + '"'
            ls_objs += ',"format":"' + json_escape(adw_source.Describe(ls_objname + ".Format")) + '"'
            ls_objs += ',"alignment":"' + json_escape(adw_source.Describe(ls_objname + ".Alignment")) + '"'
            ls_objs += ',"font_face":"' + json_escape(adw_source.Describe(ls_objname + ".Font.Face")) + '"'
            ls_objs += ',"font_height":' + json_null_if_special(adw_source.Describe(ls_objname + ".Font.Height"))
            ls_objs += ',"font_weight":' + json_null_if_special(adw_source.Describe(ls_objname + ".Font.Weight"))
            ls_objs += ',"color":' + json_null_if_special(adw_source.Describe(ls_objname + ".Color"))
            ls_objs += ',"background_color":' + json_null_if_special(adw_source.Describe(ls_objname + ".Background.Color"))
            
        case "column"
            // Column object attributes  
            ls_objs += ',"col_type":"' + json_escape(adw_source.Describe(ls_objname + ".ColType")) + '"'
            ls_objs += ',"format":"' + json_escape(adw_source.Describe(ls_objname + ".Format")) + '"'
            ls_objs += ',"alignment":"' + json_escape(adw_source.Describe(ls_objname + ".Alignment")) + '"'
            ls_objs += ',"font_face":"' + json_escape(adw_source.Describe(ls_objname + ".Font.Face")) + '"'
            ls_objs += ',"font_height":' + json_null_if_special(adw_source.Describe(ls_objname + ".Font.Height"))
            ls_objs += ',"font_weight":' + json_null_if_special(adw_source.Describe(ls_objname + ".Font.Weight"))
            ls_objs += ',"color":' + json_null_if_special(adw_source.Describe(ls_objname + ".Color"))
            ls_objs += ',"background_color":' + json_null_if_special(adw_source.Describe(ls_objname + ".Background.Color"))
            
        case else
            // Basic attributes for other object types
            ls_objs += ',"font_face":"' + json_escape(adw_source.Describe(ls_objname + ".Font.Face")) + '"'
            ls_objs += ',"font_height":' + json_null_if_special(adw_source.Describe(ls_objname + ".Font.Height"))
            ls_objs += ',"color":' + json_null_if_special(adw_source.Describe(ls_objname + ".Color"))
            ls_objs += ',"background_color":' + json_null_if_special(adw_source.Describe(ls_objname + ".Background.Color"))
    end choose
    
    ls_objs += "}"
    li_objcount++
loop
ls_objs += "]"

// --- Bands ---
ls_bands = "["
string ls_bandnames[]
ls_bandnames[1] = "header"
ls_bandnames[2] = "detail"
ls_bandnames[3] = "footer"
ls_bandnames[4] = "summary"
integer i
for i = 1 to UpperBound(ls_bandnames)
    if i > 1 then ls_bands += ","
    ls_band = ls_bandnames[i]
    ls_bands += "{"
    ls_bands += '"name":"' + json_escape(ls_band) + '"'
    // You can add more band properties if needed
    ls_bands += "}"
next
ls_bands += "]"

// --- Combine all ---
ls_json = "{"
ls_json += '"columns":' + ls_columns + ','
ls_json += '"objects":' + ls_objs + ','
ls_json += '"bands":' + ls_bands
ls_json += "}"

return ls_json
end function

private function string of_datawindow_rows_to_json (datawindow adw_source);
// Export DataWindow rows (data only) to JSON
string ls_json, ls_column, ls_value, ls_visualobjects, ls_dwtype
long ll_rows, ll_row
boolean lb_first_column

ll_rows = adw_source.RowCount()
ls_json = "["

for ll_row = 1 to ll_rows
    if ll_row > 1 then ls_json += ","
    ls_json += "{"
    lb_first_column = true

    // Get column/object names dynamically
    ls_visualobjects = adw_source.Describe("DataWindow.VisualObjects")
    integer li_pos_start, li_pos_end
    li_pos_start = 1
    li_pos_end = 0

    do while li_pos_start > 0
        li_pos_end = Pos(ls_visualobjects, "~t", li_pos_start)
        if li_pos_end = 0 then
            ls_column = Mid(ls_visualobjects, li_pos_start)
            li_pos_start = 0
        else
            ls_column = Mid(ls_visualobjects, li_pos_start, li_pos_end - li_pos_start)
            li_pos_start = li_pos_end + 1
        end if

        ls_dwtype = adw_source.Describe(ls_column + ".Type")
        if ls_dwtype = "column" or ls_dwtype = "compute" then
            string ls_coltype
            ls_coltype = adw_source.Describe(ls_column + ".ColType")
            ls_value = ""
            if ll_row <= ll_rows then
                choose case Left(ls_coltype, 5)
                    case "char("
                        ls_value = adw_source.GetItemString(ll_row, ls_column)
                        if IsNull(ls_value) then ls_value = ""
                    case "date"
                        date ld_value
                        ld_value = adw_source.GetItemDate(ll_row, ls_column)
                        if not IsNull(ld_value) then ls_value = String(ld_value)
                    case "datet"
                        datetime ldt_value
                        ldt_value = adw_source.GetItemDateTime(ll_row, ls_column)
                        if not IsNull(ldt_value) then ls_value = String(ldt_value)
                    case "decim", "numbe", "long", "ulong", "int", "uint", "real", "doubl"
                        double ldb_value
                        ldb_value = adw_source.GetItemNumber(ll_row, ls_column)
                        if not IsNull(ldb_value) then ls_value = String(ldb_value)
                    case else
                        ls_value = String(adw_source.GetItemString(ll_row, ls_column))
                        if IsNull(ls_value) then ls_value = ""
                end choose
            end if

            if not lb_first_column then ls_json += ","
            lb_first_column = false
            ls_json += '"' + json_escape(ls_column) + '":"' + json_escape(ls_value) + '"'
        end if
    loop

    ls_json += "}"
next

ls_json += "]"
return ls_json
end function

private function string json_escape(string value);
// Properly escape strings for JSON according to the spec
// Only escape the characters that MUST be escaped in JSON:
// Backslash, double quote, and control chars (< ASCII 32)

if IsNull(value) then return ""

string ls_result
integer i, len
len = Len(value)
ls_result = ""

// Process the string one character at a time for proper JSON escaping
for i = 1 to len
    string ch
    ch = Mid(value, i, 1)
    integer ascii_val
    ascii_val = Asc(ch)
    
    choose case ascii_val
        // Escape the backslash
        case 92 // '\'
            ls_result += "\\"
        // Escape the double quote    
        case 34 // '"'
            ls_result += "\\"+"~""
        // Handle ASCII control characters
        case 0 to 31
            choose case ascii_val
                case 8  // Backspace
                    ls_result += "\\b"
                case 9  // Tab
                    ls_result += "\\t"
                case 10 // Newline
                    ls_result += "\\n"
                case 12 // Form feed
                    ls_result += "\\f"
                case 13 // Carriage return
                    ls_result += "\\r"
                case else
                    // Other control chars use Unicode escape \u00XX
                    string hex
                    hex = Right("0" + String(ascii_val, "~~h"), 2)
                    ls_result += "\\u00" + hex
            end choose
        case else
            // Normal characters are added unchanged
            ls_result += ch
    end choose
next

return ls_result
end function

private function string json_null_if_special(string value);
if IsNull(value) or value = "?" or value = "!" then
    return "null"
else
    return '"' + json_escape(value) + '"'
end if
end function

private function string of_datawindow_to_virtualgrid_json (datawindow adw_source);
// Build a JSON string representing a virtual grid for your C# DLL
string ls_json, ls_columns, ls_rows, ls_bands, ls_cellattrs
integer i, j, col_count, row_count
string col_name, col_type, col_width, col_format
string row_json, cell_value, band_name
string cell_attr_json, cell_name
string ls_font, ls_fontsize, ls_fontweight, ls_underline, ls_italics, ls_strikethrough, ls_align, ls_fontcolor, ls_bgcolor

// --- Columns ---
ls_columns = "["
col_count = Integer(adw_source.Describe("DataWindow.Column.Count"))
for i = 1 to col_count
    if i > 1 then ls_columns += ","
    col_name = adw_source.Describe("#" + String(i) + ".Name")
    col_type = adw_source.Describe(col_name + ".ColType")
    col_width = adw_source.Describe(col_name + ".Width")
    col_format = adw_source.Describe(col_name + ".Format")
    ls_columns += '{'
    ls_columns += '"name":"' + json_escape(col_name) + '",'
    ls_columns += '"type":"' + json_escape(col_type) + '",'
    ls_columns += '"width":' + json_null_if_special(col_width) + ','
    ls_columns += '"format":' + json_null_if_special(col_format) + '"'
    ls_columns += '}'
next
ls_columns += "]"

// --- Bands ---
ls_bands = '[{"name":"header"},{"name":"detail"},{"name":"footer"},{"name":"summary"}]'

// --- Rows and Cell Attributes ---
ls_rows = "["
ls_cellattrs = "{"
row_count = adw_source.RowCount()
for i = 1 to row_count
    if i > 1 then ls_rows += ","
    row_json = "{"
    for j = 1 to col_count
        col_name = adw_source.Describe("#" + String(j) + ".Name")
        if j > 1 then row_json += ","
        // Get value as string, handle nulls
        cell_value = adw_source.GetItemString(i, col_name)
        if IsNull(cell_value) then cell_value = ""
        row_json += '"' + json_escape(col_name) + '":"' + json_escape(cell_value) + '"'

        // --- Cell Attributes ---
        cell_name = "cell_" + String(i - 1) + "_" + String(j - 1)
        if not (i = 1 and j = 1) then ls_cellattrs += ","
        cell_attr_json = '{'
        cell_attr_json += '"text":"' + json_escape(cell_value) + '",'
        cell_attr_json += '"is_visible":true'
        // Font
        ls_font = adw_source.Describe(col_name + ".Font.Face")
        if ls_font <> "" then cell_attr_json += ',"font":' + json_null_if_special(ls_font)
        // Font size
        ls_fontsize = adw_source.Describe(col_name + ".Font.Height")
        if ls_fontsize <> "" then cell_attr_json += ',"font_size":' + json_null_if_special(ls_fontsize)
        // Font weight
        ls_fontweight = adw_source.Describe(col_name + ".Font.Weight")
        if ls_fontweight <> "" then cell_attr_json += ',"font_weight":' + json_null_if_special(ls_fontweight)
        // Underline
        ls_underline = adw_source.Describe(col_name + ".Font.Underline")
        if ls_underline = "1" then cell_attr_json += ',"underline":true' else cell_attr_json += ',"underline":false'
        // Italics
        ls_italics = adw_source.Describe(col_name + ".Font.Italic")
        if ls_italics = "1" then cell_attr_json += ',"italics":true' else cell_attr_json += ',"italics":false'
        // Strikethrough
        ls_strikethrough = adw_source.Describe(col_name + ".Font.Strikethrough")
        if ls_strikethrough = "1" then cell_attr_json += ',"strikethrough":true' else cell_attr_json += ',"strikethrough":false'
        // Alignment
        ls_align = adw_source.Describe(col_name + ".Alignment")
        if ls_align <> "" then cell_attr_json += ',"alignment":"' + Lower(ls_align) + '"'
        // Font color
        ls_fontcolor = adw_source.Describe(col_name + ".Color")
        if ls_fontcolor <> "" then cell_attr_json += ',"font_color":' + json_null_if_special(ls_fontcolor)
        // Background color
        ls_bgcolor = adw_source.Describe(col_name + ".Background.Color")
        if ls_bgcolor <> "" then cell_attr_json += ',"background_color":' + json_null_if_special(ls_bgcolor)
        cell_attr_json += '}'
        ls_cellattrs += '"' + cell_name + '":' + cell_attr_json
    next
    row_json += '}'
    ls_rows += row_json
next
ls_rows += "]"
ls_cellattrs += "}"

// --- Combine all ---
ls_json = "{"
ls_json += '"columns":' + ls_columns + ','
ls_json += '"bands":' + ls_bands + ','
ls_json += '"rows":' + ls_rows + ','
ls_json += '"cell_attributes":' + ls_cellattrs
ls_json += "}"

return ls_json
end function

event ue_error();
// Event handler for errors - can be customized
end event

on dw_exporter.create
call super::create
end on

on dw_exporter.destroy
if IsValid(inv_exporter) then destroy inv_exporter
call super::destroy
end on