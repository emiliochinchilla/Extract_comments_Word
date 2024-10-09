import win32com.client as win32
from win32com.client import constants
import pandas as pd
import xlsxwriter

word = win32.gencache.EnsureDispatch('Word.Application')
word.Visible = False
filepath = "C:\\Users\\emichin\\Desktop\\Py\\CommentExtraction\\pythonProject\\test1.docx"


def get_comments(filepath):
    doc = word.Documents.Open(filepath)
    doc.Activate()
    activeDoc = word.ActiveDocument

    comment_data = []
    for c in activeDoc.Comments:
        if c.Ancestor is None:  # checking if this is a top-level comment
            comment_text = c.Range.Text
            sentence = str(c.Scope.Sentences(1)).strip()

            comment_data.append({
                'Author': c.Author,
                'Comment': comment_text,
                'Regarding': sentence,
                'Replies': len(c.Replies)
            })
            if len(c.Replies) > 0:  # if the comment has replies
                for r in range(1, len(c.Replies) + 1):
                    reply_text = c.Replies(r).Range.Text
                    comment_data.append({
                        'Author': c.Replies(r).Author,
                        'Comment': reply_text,
                        'Regarding': sentence,
                        'Replies': 0
                    })
    doc.Close()

    return pd.DataFrame(comment_data)


# Call the function and export to Excel
df = get_comments(filepath)
print(df.to_dict(orient="records"))
dictionary = df.to_dict(orient="records")

# Create an Excel file using xlsxwriter
workbook = xlsxwriter.Workbook('comment_data.xlsx')
worksheet = workbook.add_worksheet()

# Write the headers to the Excel file
worksheet.write_row(0, 0, df.columns)

# Write the data to the Excel file
for i, row in enumerate(df.iterrows(), start=1):
    worksheet.write_row(i, 0, row[1].tolist())

workbook.close()