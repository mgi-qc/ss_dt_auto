import sys
import os
import glob
import smartsheet

"""
Once QC has completed and files are staged for transfer and user has checked pass/fail status: 
Run this script to update DT Sheet

1.  Make sure expected files exist to transfer - get correct file names to attach to row
2.  Get the DT directory from the cwl.report file for row comment
3.  Get the work order name to match in the DT sheet column Work Order ID
4.  Update the DT sheet with the attached files, row comment, DT transfer stage and change DT Assigned to column to blank 

"""  
sheet_id = 4855853141518212

# Make sure expected files exist to transfer - get correct file names to attach to row

#check dir name to ensure correct work order

def get_files():

    files = glob.glob('*cwl.report.*.txt') + glob.glob('*cwl.results.*.tsv')

    if len(files) == 0:
        print('Exiting: No matching files were not found.')
        sys.exit(1)
    
    if len(files) == 2:
        # print('Found 2 files to attach to row.')
        return files

# Get DT directory for row comment from the report file 
def get_dt_dir():
    with open(get_files()[0], 'r') as f:
        for line in f:
            if line.startswith('Data Transfer Directory'):
                dt_dir = line.split('=')[1].strip()
                return dt_dir
            
# Need the work order name to match in the DT sheet column Work Order ID
def get_work_order():
    with open(get_files()[0], 'r') as f:
        for line in f:
            if line.startswith('Data quality report for work order'):
                work_order = line.split(':')[1].strip()
                return work_order

def get_column_ids(id_):
    column_id_dict = {}
    for col in ss.Sheets.get_columns(id_).data:
        column_id_dict[col.title] = col.id
        column_id_dict[col.id] = col.title
    return column_id_dict

# Using the SS API SDK - you will need to find the row and the the row ID.  To do this:
# Pull the DT sheet using the sheet ID: 4855853141518212

# Set the API access token
key = os.environ.get('SMRT_API')

if key is None:
    print('Environment variable SMRT_API not found.')
    sys.exit(1)

# Initialize client
ss = smartsheet.Smartsheet(key)

#get column ids
sheet_columns_dict = get_column_ids(sheet_id)

#get only the column of interest from the sheet 
column = ss.Sheets.get_sheet(4855853141518212, column_ids = [sheet_columns_dict.get('Primary Column')])

attachments = get_files()
comment = '< email alert > : QC has been completed and files are staged for transfer at: ' + get_dt_dir()

# Iterate over that sheet object - rows

# row_found = False

for row in column.rows:
    
    col = row.get_column(sheet_columns_dict['Primary Column'])
    
    if col.value == get_work_order() :

        print('Found the row to update.')
        
        row_id = row.id
        
        # Update the DT sheet with the attached files, row comment, DT transfer stage and change DT Assigned to column to blank
        
        #use a dict to update values
        
        # Build new cell value for Data Transfer Stage 
        new_cell_transfer = smartsheet.models.Cell()
        new_cell_transfer.column_id = sheet_columns_dict['Data Transfer Stage']
        new_cell_transfer.value = 'QC@MGI Complete'
        new_cell_transfer.strict = False
        # Build the row to update
        new_row1 = smartsheet.models.Row()
        new_row1.id = row_id
        new_row1.cells.append(new_cell_transfer)
        # Update row
        updated_row_transfer = ss.Sheets.update_rows(sheet_id, [new_row1])

        # Build new cell value for DT Assigned To
        new_cell_assigned = smartsheet.models.Cell()
        new_cell_assigned.column_id = sheet_columns_dict['DT Assigned To']
        new_cell_assigned.value = ' '
        new_cell_assigned.strict = False
        # Build the row to update
        new_row2 = smartsheet.models.Row()
        new_row2.id = row_id
        new_row2.cells.append(new_cell_assigned)
        # Update rows
        updated_row_assigned = ss.Sheets.update_rows(sheet_id, [new_row2])

        # Add attachments
        attachment_txt = ss.Attachments.attach_file_to_row(sheet_id, row_id, (attachments[0], open(os.path.abspath(attachments[0]), 'rb')))
        attachment_tsv = ss.Attachments.attach_file_to_row(sheet_id, row_id, (attachments[1], open(os.path.abspath(attachments[1]), 'rb')))
    
        # Add comment to row
        add_comment = ss.Discussions.create_discussion_on_row(
            sheet_id,
            row_id, 
            smartsheet.models.Discussion({
                'comment': smartsheet.models.Comment({
                'text': comment,
                    })
                })
            )

        break

print('Smartsheet updated.')    
