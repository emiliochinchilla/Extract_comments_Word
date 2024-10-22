import subprocess
import pandas as pd
import win32com.client as win32
import xlsxwriter

def get_comments(filepath):
    doc = word.Documents.Open(filepath)
    doc.Activate()
    activeDoc = word.ActiveDocument

    comment_data = []
    for c in activeDoc.Comments:
        if c.Ancestor is None:  # checking if this is a top-level comment
            comment_text = c.Range.Text
            paragraph = c.Scope.Paragraphs(1)
            paragraph_text = paragraph.Range.Text.strip()
            referenced_start = c.Scope.Start
            referenced_end = c.Scope.End
            referenced_text = activeDoc.Range(referenced_start, referenced_end).Text.strip()

            comment_data.append({
                'Author': c.Author,
                'Comment': comment_text,
                'Highlighted': referenced_text,
                'Regarding': paragraph_text,
                'Replies': len(c.Replies)
            })
            if len(c.Replies) > 0:  # if the comment has replies
                for r in range(1, len(c.Replies) + 1):
                    reply_text = c.Replies(r).Range.Text
                    reply_paragraph = c.Replies(r).Scope.Paragraphs(1)
                    reply_paragraph_text = reply_paragraph.Range.Text.strip()
                    reply_referenced_start = c.Replies(r).Scope.Start
                    reply_referenced_end = c.Replies(r).Scope.End
                    reply_referenced_text = activeDoc.Range(reply_referenced_start, reply_referenced_end).Text.strip()
                    comment_data.append({
                        'Author': c.Replies(r).Author,
                        'Comment': reply_text,
                        'Highlighted': reply_referenced_text,
                        'Regarding': reply_paragraph_text,
                        'Replies': 0
                    })
    doc.Close()

    return pd.DataFrame(comment_data, columns=['Author', 'Comment', 'Highlighted', 'Regarding', 'Replies'])

# Get the currently active Word document
word = win32.gencache.EnsureDispatch('Word.Application')
active_doc = word.ActiveDocument
filepath = active_doc.FullName

# Call the function and export to Excel
df = get_comments(filepath)
print(df.to_dict(orient="records"))
dictionary = df.to_dict(orient="records")

# Create an Excel file using xlsxwriter
excel_file = 'comment_data.xlsx'
workbook = xlsxwriter.Workbook(excel_file)
worksheet = workbook.add_worksheet()

# Write the headers to the Excel file
worksheet.write_row(0, 0, df.columns)

# Set the width and wrap text for the columns
worksheet.set_column('A:A', 22, workbook.add_format({'text_wrap': True}))
worksheet.set_column('B:B', 65, workbook.add_format({'text_wrap': True}))
worksheet.set_column('C:C', 65, workbook.add_format({'text_wrap': True}))
worksheet.set_column('D:D', 65, workbook.add_format({'text_wrap': True}))

# Write the data to the Excel file
for i, row in enumerate(df.iterrows(), start=1):
    worksheet.write_row(i, 0, row[1].tolist())

workbook.close()

# Open the Excel file
subprocess.Popen([excel_file], shell=True)