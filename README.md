# diff_excel_files

Create an action-items list to make one spreadsheet look like another.

Before making changes to a spreadsheet, save a copy of the original. Update values in the original. Do not add or remove rows. Do not add or remove columns. Do not add or remove worksheets. Do not change the names of any rows or columns.

If you do all of that, you can run this scripts to create step-by-step instructions for updating the old values to the new values. This isn't very robust, so don't expect it  to fail gracefully if you break any of the rules. In summary, you can change values anywhere *except* the first row and the first column.

## Example:

### start with this

`spreadsheet 1`

|Name|Position|Salary|
| -- | -- | -- |
|Joan|Assistant|60k|
|Jim |Assistant|60k|

### update to this

`spreadsheet 2`

|Name|Position|Salary|
| -- | -- | -- |
|Joan|Manager|80k|
|Jim |Assistant|60k|

### run the script

`python .\src\diff_excel_files\main.py spreadsheet_1 spreadsheet_2`

You will get a report:

```
INFO:root:Starting
* In Joan, update 'Position' from Assistant to Manager
* In Joan, update 'Salary' from 60k to 80k
INFO:root:Done
```

The use case I imagine is

1. export data as a spreadsheet from some system
2. make updates on that spreadsheet
3. create a worksheet to update the data on the system
4. delegate to several people

The best way to do this would be to work with the system owner to create an import of the updated spreadsheet data, but that isn't always safe, allowed, or cost effective.
