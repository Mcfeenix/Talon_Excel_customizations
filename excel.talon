app: /Excel/
-
tag(): user.find_and_replace
tag(): user.tabs

settings():
    # Pasting into a column heading in a PivotTable doesn't work,
    # but typing does
    user.paste_to_insert_threshold = -1

test command: insert("It works!")


save as excel: user.excel_save_as_format("Excel Workbook (.xlsx)")

fill down: key(ctrl-d)
fill right: key(ctrl-r)
insert that: key(ctrl-shift-=)
delete that: key(ctrl--)

paste special: key(ctrl-alt-v)

clear all: user.menu_select("Edit|Clear|All")
clear (style | formats): user.menu_select("Edit|Clear|Formats")
clear that: user.menu_select("Edit|Clear|Contents")
clear comments: user.menu_select("Edit|Clear|Clear Comments and Notes")
clear hyperlinks: user.menu_select("Edit|Clear|Hyperlinks")
clear series: user.menu_select("Edit|Clear|Series")

align left: key(ctrl-l)
align center: key(ctrl-e)
align right: key(ctrl-r)

filter: key(ctrl-shift-L)
sort: key(ctrl-shift-r)
table: key(ctrl-t)

formula: key(shift-f3)
#reference: key(ctrl-t)

#edit: key(ctrl-u)
complete: key(alt-down)
ditto: key(ctrl-shift-')

bold: key(ctrl-b)
italic: key(ctrl-i)
underline: key(ctrl-u)
strike through: key(ctrl-shift-x)

format general: key(ctrl-~)
format currency: key(ctrl-$)
format (percent | percentage): key(ctrl-%)
format (decimal | number): key(ctrl-!)
format exponential: key(ctrl-^)
format date: key(ctrl-#)
format time: key(ctrl-@)

cell border: key(ctrl-alt-0)
cell border left: key(ctrl-alt-left)
cell border right: key(ctrl-alt-right)
cell border top: key(ctrl-alt-up)
cell border bottom: key(ctrl-alt-down)
clear cell border: key(ctrl-alt--)

cell select: key(shift-backspace)
cell note: key(shift-f2)
cell comment: key(ctrl-shift-f2)
cell name: key(ctrl-f3)
cell menu: key(shift-f10)

array select: key(ctrl-/)

column hide: key(ctrl-0)
column unhide: key(ctrl-shift-0)
# XXX Sometimes ctrl-space selects more than a single column despite the documentation
column select: key(ctrl-space)
column insert: key(ctrl-space ctrl-shift-=)
column delete: key(ctrl-space ctrl--)
column top: key(ctrl-up)
column bottom: key(ctrl-down)
column fit: user.menu_select("Format|Column|AutoFit Selection")
column filter: key(ctrl-down ctrl-up alt-down)
column width: user.menu_select("Format|Column|Width...")

row hide: key(ctrl-9)
row unhide: key(ctrl-shift-9)
row select: key(shift-space)
row insert [up]: key(shift-space ctrl-shift-=)
row delete: key(shift-space ctrl--)
row start: key(ctrl-left)
row end: key(ctrl-right)
row fit: user.menu_select("Format|Row|AutoFit")
row height: user.menu_select("Format|Row|Height...")

table select: key(ctrl-a)
select all: key(ctrl-a ctrl-a ctrl-a)

sheet new: key(shift-f11)
sheet previous: key(alt-left)
sheet next: key(alt-right)
sheet rename:
    key(esc)
    user.menu_select("Format|Sheet|Rename")

pivot that: user.menu_select("Data|Summarize with PivotTable")
mail this: user.menu_select("File|Share|Send Workbook")

#ribbon: key(ctrl-alt-r)

(select | take | go [to]) <user.excel_reference>:
    key(ctrl-g)
	sleep(100ms)
    insert(excel_reference)
    key(return)
