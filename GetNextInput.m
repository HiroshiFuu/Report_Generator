function input = GetNextInput(Doc)
    global row_index
    
    input = Doc.Tables.Item(1).Cell(row_index, 3).Range.Text;
    row_index = row_index + 1;
    input = strrep(input, char(13), '');
    input = strrep(input, char(7), '');
end 