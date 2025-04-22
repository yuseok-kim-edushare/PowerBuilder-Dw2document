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
private function string json_escape(string value)
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
    
    // Convert DataWindow to JSON data
    string ls_json_data
    ls_json_data = of_datawindow_to_json(adw_source)
	 
	// export json into text file for debugging
	integer li_file
	li_file = FileOpen("C:\temp\dw2docs-json.txt",LineMode!, Write!)
	FileWrite(li_file,ls_json_data)
	FileClose(li_file)
    
    // Try first deleting any existing file
    FileDelete(as_output_path)
    
    // Export to Excel using the nvo_datawindowexporter
    string ls_result
    string ls_path = "C:\Temp\test.xlsx"
    ls_result = inv_exporter.of_exporttoexcel(ls_json_data, ls_path, as_sheet_name)
    
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
    
    // Convert DataWindow to JSON data
    string ls_json_data
    ls_json_data = of_datawindow_to_json(adw_source)
    
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
// Convert DataWindow to JSON with cell properties, skipping empty rows
string ls_json, ls_error
long ll_rows, ll_row
string ls_colname, ls_value

try
    // Start JSON object
    ls_json = "{";
    
    // Add metadata
    ls_json += '"metadata":{';
    ls_json += '"name":"' + adw_source.Describe("DataWindow.Name") + '",';
    ls_json += '"title":"' + adw_source.Describe("DataWindow.Title") + '"';
    ls_json += '},';
    
    // Get row count
    ll_rows = adw_source.RowCount()
    
    // Add data rows
    ls_json += '"rows":[';
    
    boolean lb_first_row = true
    for ll_row = 1 to ll_rows
        // Prepare to build the row JSON
        string ls_row_json
        boolean lb_first_column = true
        ls_row_json = "{";
        
        // Get column names dynamically
        string ls_column_list, ls_column
        integer li_pos_start = 1, li_pos_end
        ls_column_list = adw_source.Describe("DataWindow.Objects")
        
        do while li_pos_start > 0
            li_pos_end = Pos(ls_column_list, "~t", li_pos_start)
            if li_pos_end = 0 then
                ls_column = Mid(ls_column_list, li_pos_start)
                li_pos_start = 0
            else
                ls_column = Mid(ls_column_list, li_pos_start, li_pos_end - li_pos_start)
                li_pos_start = li_pos_end + 1
            end if
            
            // Check if it's a column
            string ls_type
            ls_type = adw_source.Describe(ls_column + ".Type")
            if ls_type = "column" then
                // Get column value
                ls_value = String(adw_source.GetItemString(ll_row, ls_column))
                if IsNull(ls_value) then ls_value = ""
                
                // Get cell properties
                string ls_font_face, ls_font_height, ls_font_weight, ls_font_italic, ls_font_underline, ls_font_strikethrough
                string ls_color, ls_bgcolor, ls_alignment, ls_format
                
                ls_font_face = adw_source.Describe(ls_column + ".Font.Face")
                ls_font_height = adw_source.Describe(ls_column + ".Font.Height")
                ls_font_weight = adw_source.Describe(ls_column + ".Font.Weight")
                ls_font_italic = adw_source.Describe(ls_column + ".Font.Italic")
                ls_font_underline = adw_source.Describe(ls_column + ".Font.Underline")
                ls_font_strikethrough = adw_source.Describe(ls_column + ".Font.Strikethrough")
                ls_color = adw_source.Describe(ls_column + ".Color")
                ls_bgcolor = adw_source.Describe(ls_column + ".Background.Color")
                ls_alignment = adw_source.Describe(ls_column + ".Alignment")
                ls_format = adw_source.Describe(ls_column + ".Format")
                
                // Only add the column if it has a value or at least one property (not all are empty)
                boolean lb_has_data = ls_value <> "" or ls_font_face <> "" or ls_font_height <> "" or ls_font_weight <> "" or ls_font_italic <> "" or ls_font_underline <> "" or ls_font_strikethrough <> "" or ls_color <> "" or ls_bgcolor <> "" or ls_alignment <> "" or ls_format <> ""
                if lb_has_data then
                    if not lb_first_column then ls_row_json += ","
                    lb_first_column = false
                    // Add to JSON (nested object)
                    ls_row_json += '"' + ls_column + '":{';
                    ls_row_json += '"value":"' + json_escape(ls_value) + '",';
                    ls_row_json += '"font_face":"' + json_escape(ls_font_face) + '",';
                    ls_row_json += '"font_height":"' + json_escape(ls_font_height) + '",';
                    ls_row_json += '"font_weight":"' + json_escape(ls_font_weight) + '",';
                    ls_row_json += '"font_italic":"' + json_escape(ls_font_italic) + '",';
                    ls_row_json += '"font_underline":"' + json_escape(ls_font_underline) + '",';
                    ls_row_json += '"font_strikethrough":"' + json_escape(ls_font_strikethrough) + '",';
                    ls_row_json += '"color":"' + json_escape(ls_color) + '",';
                    ls_row_json += '"background_color":"' + json_escape(ls_bgcolor) + '",';
                    ls_row_json += '"alignment":"' + json_escape(ls_alignment) + '",';
                    ls_row_json += '"format":"' + json_escape(ls_format) + '"';
                    ls_row_json += "}"
                end if
            end if
        loop
        
        ls_row_json += "}"
        // Only add the row if it has at least one column
        if not lb_first_column then
            if not lb_first_row then ls_json += ","
            lb_first_row = false
            ls_json += ls_row_json
        end if
    next
    
    ls_json += "]";
    
    // Close the main JSON object
    ls_json += "}";
    
    return ls_json
catch (Exception le_ex)
    MessageBox("JSON Conversion Error", le_ex.getMessage())
    return "{}"
end try
end function

private function string json_escape(string value);
string ls_result
integer li_pos

ls_result = value

// Escape backslash
li_pos = Pos(ls_result, "~\\")
do while li_pos > 0
    ls_result = Left(ls_result, li_pos - 1) + "~\\~\\" + Mid(ls_result, li_pos + 1)
    li_pos = Pos(ls_result, "~\\", li_pos + 2)
loop

// Escape double quote
li_pos = Pos(ls_result, '~"')
do while li_pos > 0
    ls_result = Left(ls_result, li_pos - 1) + "~\\~"" + Mid(ls_result, li_pos + 1)
    li_pos = Pos(ls_result, '~"', li_pos + 3)
loop

// Escape newline
li_pos = Pos(ls_result, Char(10))
do while li_pos > 0
    ls_result = Left(ls_result, li_pos - 1) + "\\n" + Mid(ls_result, li_pos + 1)
    li_pos = Pos(ls_result, Char(10), li_pos + 2)
loop

// Escape carriage return
li_pos = Pos(ls_result, Char(13))
do while li_pos > 0
    ls_result = Left(ls_result, li_pos - 1) + "\\r" + Mid(ls_result, li_pos + 1)
    li_pos = Pos(ls_result, Char(13), li_pos + 2)
loop

// Escape tab
li_pos = Pos(ls_result, Char(9))
do while li_pos > 0
    ls_result = Left(ls_result, li_pos - 1) + "\\t" + Mid(ls_result, li_pos + 1)
    li_pos = Pos(ls_result, Char(9), li_pos + 2)
loop

// Escape other control characters (ASCII < 32 except for \n, \r, \t)
integer i
for i = 1 to Len(ls_result)
    integer ascii_val
    ascii_val = Asc(Mid(ls_result, i, 1))
    if ascii_val < 32 and ascii_val <> 10 and ascii_val <> 13 and ascii_val <> 9 then
        string hex = String(ascii_val, "~x04")
        ls_result = Left(ls_result, i - 1) + "\\u00" + Right("0" + String(ascii_val, "~x2"), 2) + Mid(ls_result, i + 1)
        // After replacement, move i forward by 5 (\u00XX is 6 chars, but we replaced 1)
        i += 5
    end if
next

return ls_result
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