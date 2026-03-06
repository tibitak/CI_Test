function results = ConvertWeekString(week_string)
         % To handle different types of DateTime format,
         % check the index of '.' and '-'
         indexDot = strfind(week_string,'.');
         indexMinus = strfind(week_string,'-');

         % Convert week_string to datetime
         % In case it uses '.'
         if contains(week_string, '.')
             % If it has a proper date format
             if length(split(week_string, '.')) == 3
                 % Check format and convert accordingly
                 if strlength(week_string) == 10 && indexDot(1) == 5
                     % Format is yyyy.MM.dd
                     results.DateTime = datetime(week_string, 'InputFormat', 'yyyy.MM.dd');
                 else
                     % Format is dd.MM.yyyy
                     results.DateTime = datetime(week_string, 'InputFormat', 'dd.MM.yyyy');
                 end

                 % Indicate the a proper date is found
                 results.IsFound = true;
             end
         % Convert weekString to datetime in case '-'
         % is used
         elseif contains(week_string, '-')
             % Check if it a correct format
             if length(split(week_string, '-')) == 3 && ((contains(week_string, 'Jul') || contains(week_string, 'Jan') || contains(week_string, 'Feb') || ...
                        contains(week_string, 'Mar') || contains(week_string, 'Apr') || contains(week_string, 'May') || ...
                        contains(week_string, 'Jun') || contains(week_string, 'Aug') || contains(week_string, 'Sep') || ...
                        contains(week_string, 'Oct') || contains(week_string, 'Nov') || contains(week_string, 'Dec')))
                 % Check format and convert accordingly
                 if indexMinus(1) == 5
                     % Format is dd-MMM-yyyy
                     results.DateTime = datetime(week_string, 'InputFormat', 'yyyy-MMM-dd');
                 else
                     % Format is dd.MM.yyyy
                     results.DateTime = datetime(week_string, 'InputFormat', 'dd-MMM-yyyy');
                 end
                 
                 % Indicate the a proper date is found
                 results.IsFound = true;
             end
         end            
    end