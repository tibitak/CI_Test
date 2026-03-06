function rowIndex = findValueInTable(tbl, calendarWeek, colleague)
    % Find the row index where the variable contains the specified value
    rowIndex = find(and(tbl.CalendarWeek == calendarWeek, tbl.Name == colleague));
    
    % If no rows are found, return an empty array
    if isempty(rowIndex)
        rowIndex = [];
    end
end