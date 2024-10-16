import win32com.client as win32
import pandas as pd
import xlsxwriter
import subprocess
import sys
import PyInstaller.__main__

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

            comment_data.append({
                'Author': c.Author,
                'Comment': comment_text,
                'Regarding': paragraph_text,
                'Replies': len(c.Replies)
            })
            if len(c.Replies) > 0:  # if the comment has replies
                for r in range(1, len(c.Replies) + 1):
                    reply_text = c.Replies(r).Range.Text
                    reply_paragraph = c.Replies(r).Scope.Paragraphs(1)
                    reply_paragraph_text = reply_paragraph.Range.Text.strip()
                    comment_data.append({
                        'Author': c.Replies(r).Author,
                        'Comment': reply_text,
                        'Regarding': reply_paragraph_text,
                        'Replies': 0
                    })
    doc.Close()

    return pd.DataFrame(comment_data)

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

# Write the data to the Excel file
for i, row in enumerate(df.iterrows(), start=1):
    worksheet.write_row(i, 0, row[1].tolist())

workbook.close()

# Open the Excel file
subprocess.Popen([excel_file], shell=True)

if __name__ == '__main__':
    PyInstaller.__main__.run([
        '--onefile',
        '--windowed',
        sys.argv[0]
    ])