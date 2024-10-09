import win32com.client as win32
from win32com.client import constants

word = win32.gencache.EnsureDispatch('Word.Application')
word.Visible = False
filepath = "C:\\Users\\emichin\\Desktop\\Py\\CommentExtraction\\pythonProject\\test.docx"


def get_comments(filepath):
    doc = word.Documents.Open(filepath)
    doc.Activate()
    activeDoc = word.ActiveDocument
    for c in activeDoc.Comments:
        if c.Ancestor is None:  # checking if this is a top-level comment
            print("Comment by: " + c.Author)
            print("Comment text: " + c.Range.Text)  # text of the comment
            print("Regarding: " + c.Scope.Sentences(1).Text)  # get the entire sentence the comment is referencing
            if len(c.Replies) > 0:  # if the comment has replies
                print("Number of replies: " + str(len(c.Replies)))
                for r in range(1, len(c.Replies) + 1):
                    print("Reply by: " + c.Replies(r).Author)
                    print("Reply text: " + c.Replies(r).Range.Text)  # text of the reply
    doc.Close()

get_comments(filepath)