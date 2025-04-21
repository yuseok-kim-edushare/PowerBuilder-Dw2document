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
// Convert DataWindow to JSON
// Keep existing implementation as it's not in nvo_datawindowexporter
string ls_json, ls_error
long ll_rows, ll_columns, ll_row, ll_col
string ls_colname, ls_value

try
    // Start JSON object
    ls_json = "{"
    
    // Add metadata
    ls_json += '"metadata":{'
    ls_json += '"name":"' + adw_source.Describe("DataWindow.Name") + '",'
    ls_json += '"title":"' + adw_source.Describe("DataWindow.Title") + '"'
    ls_json += '},'
    
    // Get row and column counts
    ll_rows = adw_source.RowCount()
    
    // Add data rows
    ls_json += '"rows":['
    
    for ll_row = 1 to ll_rows
        if ll_row > 1 then ls_json += ","
        
        ls_json += "{"
        
        // Get column names dynamically
        string ls_column_list, ls_column
        integer li_pos_start = 1, li_pos_end
        ls_column_list = adw_source.Describe("DataWindow.Objects")
        
        boolean lb_first_column = true
        
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
                if not lb_first_column then ls_json += ","
                lb_first_column = false
                
                // Get column value
                ls_value = String(adw_source.GetItemString(ll_row, ls_column))
                if IsNull(ls_value) then ls_value = ""
                
                // Add to JSON
                ls_json += '"' + ls_column + '":"' + json_escape(ls_value) + '"'
            end if
        loop
        
        ls_json += "}"
    next
    
    ls_json += "]"
    
    // Close the main JSON object
    ls_json += "}"
    
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
li_pos = Pos(ls_result, "~"")
do while li_pos > 0
    ls_result = Left(ls_result, li_pos - 1) + "~\\~"" + Mid(ls_result, li_pos + 1)
    li_pos = Pos(ls_result, "~"", li_pos + 3)
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