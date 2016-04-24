Word = actxserver('Word.application'); 
Word.Visible = 1; 
set(Word,'DisplayAlerts',0); 
Docs = Word.Documents; 
Doc = Docs.Open('C:\Users\FENG0\Google Drive\Report_Generator\Report Generator Input Form 2.docx'); 