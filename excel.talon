app: /Excel/
-

tag(): user.find_and_replace
tag(): user.tabs

settings():
    # Pasting into a column heading in a PivotTable doesn't work,
    # but typing does
    user.paste_to_insert_threshold = -1

# Test if Talon recognizes Excel
test command: insert("It works!")

# ----------------------------------------
# Unsure about these commands

# Toggle AutoFilter on/off
filter: key(ctrl-shift-L)

# Open Sort dialog
sort: key(ctrl-shift-r)

# Convert range to table or create new table
table: key(ctrl-t)

# Open Insert Function dialog
formula: key(shift-f3)

# Show AutoComplete dropdown for current cell
complete: key(alt-down)

# Copy formula from cell above (without formatting)
ditto: key(ctrl-shift-')
# ----------------------------------------
# Housekeeping

# Save workbook as .xlsx format
save as excel: user.excel_save_as_format("Excel Workbook (.xlsx)")

# Rename current worksheet
rename sheet: key(alt-h o r)

# Create PivotTable from selected range
pivot that: key(alt-n v t)

# ----------------------------------------
# Formatting

# Copy cell content down to selected cells
fill down: key(ctrl-d)

# Copy cell content right to selected cells
fill right: key(ctrl-r)

# Insert cells, rows, or columns
insert that: key(ctrl-shift-=)

# Delete cells, rows, or columns
delete that: key(ctrl--)

# Open Paste Special dialog
paste special: key(ctrl-alt-v)

# Clear all content, formats, and comments
clear all: key(alt-h e a)

# Clear only cell formatting
clear format: key(alt-h e f)

# Clear only cell contents
clear content: key(alt-h e c)

# Clear comments and notes from cells
clear (comment | note) : key(alt-h e m)

# Remove hyperlinks from cells
clear hyperlinks: key(alt-h e l)

# Left align cell content
align left: key(alt-h a l)

# Center align cell content
align center: key(alt-h a c)

# Right align cell content
align right: key(alt-h a r)

# Freeze top row
freeze top row: key(alt-w f r)

# Freeze / unfreeze panes
(freeze | unfreeze) that: key(alt-w f f)

# Apply / remove bold formatting
bold: key(ctrl-b)

# Apply / remove italic formatting
italic: key(ctrl-i)

# Apply / remove underline formatting
underline: key(ctrl-u)

# Apply / remove strikethrough formatting
strike through: key(ctrl-5)

# Apply General number format
format general: key(ctrl-~)

# Apply Currency format
format currency: key(ctrl-$)

# Apply Percentage format
format (percent | percentage): key(ctrl-%)

# Apply Number format with thousand separators
format number: key(ctrl-!)

# Apply Scientific notation format
format scientific: key(ctrl-^)

# Apply Date format
format date: key(ctrl-#)

# Apply Time format
format time: key(ctrl-@)

# Access border dialogue box
cell border: key(alt-h b)

# Add or edit cell note
(add | edit) note: key(shift-f2)

# Add or edit threaded comment
(add | edit) comment: key(ctrl-shift-f2)

# Open context menu for selected cell
cell menu: key(shift-f10)

# Hide selected columns
hide column: key(ctrl-0)

# Unhide columns in selection
unhide column: key(alt-h o u l)

# Select entire column(s)
select column: key(ctrl-space)

# Insert new column(s)
insert column: key(ctrl-space ctrl-shift-=)

# Delete selected column(s)
delete column: key(ctrl-space ctrl--)

# Auto-fit column width to content
fit column: key(alt-h o i)

# Apply | remove column filters
(apply | remove) filter: key(alt-h s f)

# Open Column Width dialog
column width: key(alt-h o w)

# Hide selected rows
hide row: key(ctrl-9)

# Unhide rows in selection
unhide row: key(ctrl-shift-9)

# Select entire row(s)
row select: key(shift-space)

# Insert new row(s) above selection
row insert [up]: key(shift-space ctrl-shift-=)

# Delete selected row(s)
delete row: key(shift-space ctrl--)

# Auto-fit row height to content
fit row: key(alt-h o a)

# Format cells dialogue box
format cell: key(alt-h o)

# Insert new worksheet
sheet new: key(shift-f11)

# ----------------------------------------
# Navigation

# Select only the active cell (deselect range)
select cell: key(shift-backspace)

# Navigate to specific cell or range (e.g., "go to B2")
(select | go [to]) <user.excel_reference>:
    key(ctrl-g)
    sleep(100ms)
    insert(excel_reference)
    key(return)

# Move to previous cell (left or up in selection)
go previous cell: key(shift-tab)

# Move one cell up
go up: key(up)

# Move one cell down
go down: key(down)

# Move one cell left
go left: key(left)

# Move one cell right
go right: key(right)

# Select current data region
select region: key(ctrl-shift-space)

# Select entire worksheet
select all: key(ctrl-a ctrl-a ctrl-a)

# Move to top of current data region in column
column top: key(ctrl-up)

# Move to bottom of current data region in column
column bottom: key(ctrl-down)

# Move to start of current data region in row
row start: key(ctrl-left)

# Move to end of current data region in row
row end: key(ctrl-right)

# Jump to left edge of data region
jump left: key(ctrl-left)

# Jump to right edge of data region
jump right: key(ctrl-right)

# Jump to top edge of data region
jump up: key(ctrl-up)

# Jump to bottom edge of data region
jump down: key(ctrl-down)

# Jump to next nonblank cell to the left
end left: key(end left)

# Jump to next nonblank cell to the right
end right: key(end right)

# Jump to next nonblank cell above
end up: key(end up)

# Jump to next nonblank cell below
end down: key(end down)

# Select from current cell to last used cell
select to last: key(ctrl-shift-end)

# Move to top-left visible cell (requires Scroll Lock on)
go top left: key(home scroll_lock)

# Scroll down one screen
page down: key(pagedown)

# Scroll up one screen
page up: key(pageup)

# Scroll right one screen
page right: key(alt-pagedown)

# Scroll left one screen
page left: key(alt-pageup)

# Zoom in on worksheet
zoom in: key(ctrl-alt-=)

# Zoom out on worksheet
zoom out: key(ctrl-alt--)

# Move to previous worksheet tab
sheet previous: key(ctrl-pageup)

# Move to next worksheet tab
sheet next: key(ctrl-pagedown)

