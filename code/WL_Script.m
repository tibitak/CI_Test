result = true;

makePPTCompilable(); 
% Import the PPT API
import mlreportgen.ppt.*;

% Store the current parameters to a structure
data.WorkLoad_FilePath_ToStore = "C:\Users\tibor.takacs\SciEngineer Kft\Engineering - Engineering-Leadership - Dokumentumok\Engineering-Leadership\Reporting\Workload";
data.WorkinHours_Path_ToStore = "C:\Users\tibor.takacs\SciEngineer Kft\Engineering - Dokumentumok\General\03 Engineering Documents\02 Organization Documents\Reporting\AE_Dashboards_sources\AE_WorkingHours.xlsx";

data.StartWeek_ToStore = datetime("2026-01-06");

% Get the start week and year from the StartWeek parameter
StartWeek = week(data.StartWeek_ToStore);
StartYear = year(data.StartWeek_ToStore);

% Perfom the analyzis only if the input parameters exis
if isfile(data.WorkinHours_Path_ToStore) && exist(data.WorkLoad_FilePath_ToStore, 'dir') == 7
    % Read Names and Working Hours from Working Hours excel
    WorkingHours_Names = readtable(data.WorkinHours_Path_ToStore, 'Sheet', 'Names');
    WorkingHours_List = readtable(data.WorkinHours_Path_ToStore, 'Sheet', 'WorkingHours');
    
    % Arrange the working hours for each week for each
    % colleague
    WorkingHours_PerColleage_PerWeek = groupsummary(WorkingHours_List,{'AEName', 'Year', 'Week_ID'}, {'sum', 'sum', 'sum'}, {'WorkingHours', 'WorkingHours', 'WorkingHours'});
    
    % Map colleagues from WorkingHours_Names to WorkingHours_PerColleage_PerWeek
    WorkingHours_PerColleage_PerWeek = innerjoin(WorkingHours_PerColleage_PerWeek, WorkingHours_Names, 'LeftKeys', 'AEName', 'RightKeys', 'Id', 'RightVariables',{'Name', 'FileName'});
    
    % Load excels with "WorkLoad_Plan_..."
    files = dir(fullfile(data.WorkLoad_FilePath_ToStore, 'Workload_Plan_*.xlsx'));
    
    % Initialize a cell array to store the data from each file
    WorkLoads = cell(length(files), 2);
    
    % Loop through each file and read the data
    for i = 1:length(files)
        filePath = fullfile(data.WorkLoad_FilePath_ToStore, files(i).name);
        WorkLoads{i,2} = readtable(filePath, 'Sheet', 'WeeklyTasks');
        WorkLoads{i,1} = files(i).name;
    end
    
    % Initialize the weekly load table
    WeekLoad = table('Size', [0, 4], 'VariableTypes', {'string', 'string', 'double','double'}, 'VariableNames', {'Name', 'CalendarWeek', 'Load', 'Target'});

    % Initialize the weekly load table
    WeelyMainTasks = table('Size', [0, 4], 'VariableTypes', {'string', 'string', 'string','double'}, 'VariableNames', {'Name', 'CalendarWeek', 'Topic', 'Load'});
    
    % Go through all workload table
    for i = 1:height(WorkLoads)
        % Get the table for WorkLoad
        currentTable = WorkLoads{i, 2};
        % Get the filename
        [~, filename, ~] = fileparts(WorkLoads{i, 1});
        % Get the name of the collegaue
        nameOfColleague = string(extractAfter(filename, 14));
        
        % Go through the WorkLoad table
        for j = 1:height(currentTable)
             % Get the week and duration information from each
             % line
             weekString = string(currentTable{j,1});
             weekTopic = string(currentTable{j,2});

             try
                d = datetime(weekString);      % attempt parse
                isDate = ~isnat(d);   % true if parsed
             catch
                isDate = false;
             end

             % Get effor
             effortPlanned = double(currentTable{j,3});
    
             % If the current line has a date, and has entries in each cell
             if isDate && 0 == strcmp(currentTable{j,1},"") && 0 == strcmp(currentTable{j,2},"") && 0 == strcmp(currentTable{j,3},"")
                 % Convert the date to a uniform date value
                 dateValueConverted = ConvertWeekString(weekString);
                 dateValue = dateValueConverted.DateTime;
                 fDateFound = dateValueConverted.IsFound;

                 % Get the calendar week from the date
                 calendarWeek = week(dateValue, 'iso-weekofyear');
                 calendarYear = year(dateValue);

                 if (fDateFound) && ((calendarWeek >= StartWeek && calendarYear >= StartYear) || (calendarYear > StartYear))
                     % Generate ID
                     IDofWeek = string(calendarYear) + "/" + sprintf('%02d', calendarWeek);
                     weekIndex = findValueInTable(WeekLoad, IDofWeek, nameOfColleague);

                     % Get the target load from
                     % WorkingHours
                     targetLoadRow = table2cell(WorkingHours_PerColleage_PerWeek(WorkingHours_PerColleage_PerWeek.Year == calendarYear & WorkingHours_PerColleage_PerWeek.Week_ID == calendarWeek & WorkingHours_PerColleage_PerWeek.FileName == nameOfColleague, 'sum_WorkingHours'));

                     % If there is no target load (e.g. for
                     % a trainee, set it to a default value)
                     if height(targetLoadRow) == 0
                         targetLoadRow = {40};
                     end    
                     
                     if (~isempty(weekIndex))
                         WeekLoad(weekIndex(1), 'Load') = WeekLoad(weekIndex(1), 'Load') + effortPlanned;
                     else
                         WeekLoad(end + 1, :) = {nameOfColleague, string(calendarYear) + "/" + sprintf('%02d', calendarWeek), effortPlanned, double(targetLoadRow{1}*0.8)};     
                     end

                     % Get the matching rows
                     weekIndex = findValueInTable(WeelyMainTasks, IDofWeek, nameOfColleague);

                     % Check if two main topics are already
                     % gathered for the colleague for the
                     % current calendarweek
                     if (length(weekIndex) == 2)
                        a = WeelyMainTasks(weekIndex(1), :).Load;
                        b = WeelyMainTasks(weekIndex(2), :).Load;
                        
                        if (effortPlanned > b) && (effortPlanned > a)
                            if (a > b)
                                % Replace b with the new
                                % effortPlanned
                                WeelyMainTasks(weekIndex(2), :) = {nameOfColleague, string(calendarYear) + "/" + sprintf('%02d', calendarWeek), weekTopic, effortPlanned}; 
                            else
                                % Replace a with the new
                                % effortPlanned
                                WeelyMainTasks(weekIndex(1), :) = {nameOfColleague, string(calendarYear) + "/" + sprintf('%02d', calendarWeek), weekTopic, effortPlanned}; 
                            end
                        elseif (effortPlanned > b)
                                % Replace b with the new
                                % effortPlanned
                                WeelyMainTasks(weekIndex(2), :) = {nameOfColleague, string(calendarYear) + "/" + sprintf('%02d', calendarWeek), weekTopic, effortPlanned}; 
                        elseif (effortPlanned > a)
                                % Replace a with the new
                                % effortPlanned
                                WeelyMainTasks(weekIndex(1), :) = {nameOfColleague, string(calendarYear) + "/" + sprintf('%02d', calendarWeek), weekTopic, effortPlanned}; 
                        end                      
                    else
                        WeelyMainTasks(end + 1, :) = {nameOfColleague, string(calendarYear) + "/" + sprintf('%02d', calendarWeek), weekTopic, effortPlanned};     
                    end
                 end
             end
        end
    end

    % Create a new presentation
    ppt = Presentation(data.WorkLoad_FilePath_ToStore + "\WorkLoad_Summary.pptx"); % Specify the output file name
    
    % Get unique colleagues and calendar weeks
    uniqueColleagues = unique(WeekLoad.Name);

    % Plot each colleague's effort per calendar week
    for k = 1:length(uniqueColleagues)    
        colleague = uniqueColleagues(k);                    

        % Show figures only if the corresponding checkbox is
        % selected
        showFigure = 'off';

        % Create a figure for the plot
        figure('Visible', showFigure,'Name','WorkLoad for ' + colleague);
        hold on
        % Extract efforts for the current colleague
        efforts = WeekLoad.Load(WeekLoad.Name == colleague);
        targets = WeekLoad.Target(WeekLoad.Name == colleague);
        weeks = WeekLoad.CalendarWeek(WeekLoad.Name == colleague);
        
        % It is needed to be able to use bar, plot and xline on
        % the same figure
        x = unique(categorical(weeks));

        if length(x) == numel(efforts)
            % Plot the data
            b = bar(x, efforts);
            legend('Work Load [h]')
        
            % Get the x-coordinates and y-coordinates for the text placement
            xtips = b.XEndPoints; % X coordinates of the bar centers
            ytips = b.YEndPoints; % Y coordinates of the bar heights
            labels = string(b.YData); % Convert Y data to string for labeling
        
            % Display the values on top of the bars
            text(xtips, ytips, labels, 'HorizontalAlignment', 'center', ...
                 'VerticalAlignment', 'bottom');
        
            % Plot the target
            plot(x, targets, 'or', 'LineWidth', 2, 'DisplayName', 'Target Load [h]');
            ylim([0 60]);
                            
            % Add a vertical line to the figure at the actual week
            currentWeek = week(datetime('today'));
            currentYear = year(datetime('today'));
            currentCalendarWeek = categorical(string(currentYear) + "/" + sprintf('%02d', currentWeek));
            xline(currentCalendarWeek, '--k', 'Current Week', 'LabelHorizontalAlignment', 'left', 'LabelVerticalAlignment', 'bottom', 'DisplayName', 'Current Week');
        
            % Add labels and title
            xlabel('Calendar Week');
            ylabel('Effort');
            title('Planned WorkLoad for ' + colleague);
            legend('show');
            grid on;
        
            figPicName = data.WorkLoad_FilePath_ToStore + "\" + colleague + '_plot.png';
        
            % Export the current figure to a PNG file
            exportgraphics(gcf, figPicName);
        
            % Add a title slide
            slideColl = add(ppt, 'Blank');
        
            % Create a Picture object for the image you want to add
            plane = Picture(figPicName); % Specify the image file
            plane.X = '0.1in'; % X position
            plane.Y = '0.1in'; % Y position
            plane.Width = '12in'; % Width of the image
            plane.Height = '7in'; % Height of the image
        
            % Add the Picture object to the slide
            add(slideColl, plane);
        
            hold off
        end
    end

    % Calculate the load in percentage compared to the target
    WeekLoad.LoadPer = WeekLoad.Load ./ WeekLoad.Target;

    % Create Pivot for WeekLoad
    pivot_WeekLoads = pivot(WeekLoad, Rows="CalendarWeek",Method="mean", DataVariable="LoadPer");

    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    % Create a figure for the plot
    figure('Visible', showFigure,'Name','Average Load per Calendar Week [%]');
    hold on
    % Extract efforts for the current colleague
    load = pivot_WeekLoads.mean_LoadPer * 100;
    weeks = pivot_WeekLoads.CalendarWeek;
    
    % Plot the data
    b = bar(weeks, load);
    legend('Average Work Load [%]')

    % Get the x-coordinates and y-coordinates for the text placement
    xtips = b.XEndPoints; % X coordinates of the bar centers
    ytips = b.YEndPoints; % Y coordinates of the bar heights
    labels = compose('%.2f%%', b.YData); % Convert Y data to string for labeling

    % Display the values on top of the bars
    text(xtips, ytips, labels, 'HorizontalAlignment', 'center', ...
         'VerticalAlignment', 'bottom');

                   
    % Add a vertical line to the figure at the actual week
    currentWeek = week(datetime('today'));
    currentYear = year(datetime('today'));
    currentCalendarWeek = categorical(string(currentYear) + "/" + sprintf('%02d', currentWeek));
    xline(currentCalendarWeek, '--k', 'Current Week', 'LabelHorizontalAlignment', 'left', 'LabelVerticalAlignment', 'bottom', 'DisplayName', 'Current Week');

    % Add labels and title
    xlabel('Calendar Week');
    ylabel('Average Work Load [%]');
    title('Averge Planned WorkLoad for each Calendar Week');
    legend('show');
    grid on;

    figPicName = data.WorkLoad_FilePath_ToStore + "\AverageLoadPerCalWeek" + '_plot.png';

    % Export the current figure to a PNG file
    exportgraphics(gcf, figPicName);

    % Add a title slide
    slideColl = add(ppt, 'Blank');

    % Create a Picture object for the image you want to add
    plane = Picture(figPicName); % Specify the image file
    plane.X = '0.1in'; % X position
    plane.Y = '0.1in'; % Y position
    plane.Width = '12in'; % Width of the image
    plane.Height = '7in'; % Height of the image

    % Add the Picture object to the slide
    add(slideColl, plane);

    hold off
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%

    % Add a slide (layout name depends on your template)
    slide = add(ppt, 'Title and Content');

    currentLoad = WeelyMainTasks(WeelyMainTasks.CalendarWeek == string(currentCalendarWeek),:);
    currentLoad.Load = string(currentLoad.Load); % Needed to be able to set the font size for this column as well
    
    % Create a Table object and add it to the slide
    tbl = Table(currentLoad);
    tbl.Width = '900pt';
    tbl.X = '0.1in';
    tbl.Y = '0.1in';
    tbl.StyleName = "Medium Style 2 - Accent 1"; 
    
    % Create the format and set the properties
    tblStyle = TableStyleOptions(); 
    tblStyle.FirstRow = true; 
    tblStyle.LastRow = false; 
    tblStyle.FirstColumn = true; 
    tblStyle.LastColumn = false; 
    tblStyle.BandedRows = true; 
    tblStyle.BandedColumns = false; 

    tbl.entry(1,1).Style = {ColWidth("100pt")};
    tbl.entry(1,2).Style = {ColWidth("150pt")};
    tbl.entry(1,3).Style = {ColWidth("550pt")};
    tbl.entry(1,4).Style = {ColWidth("100pt")};
    
    % Apply the formatting to the table
    tbl.Style = [tblStyle, FontSize("16pt")];
    
    add(slide, tbl);

    % Close and save the presentation
    close(ppt);                

    % Get data and time for the filename
    currentDate = datetime('today');
    currentDate.Format = 'yyyy_MM_dd';
    currentTime = datetime('now');
    currentTime.Format = 'HH:mm';
    
    % Create FileName with unique name using date and time
    File_Name = "WorkLoad_Stat" + "_" + string(currentDate) +"_" + string(hour(currentTime)) + "_" + string(minute(currentTime)) + ".xlsx";
    
    % Save the tables to the ResultFile
    writetable(WeekLoad, data.WorkLoad_FilePath_ToStore + "\" + File_Name, 'Sheet', 'WeekLoads');
    writetable(pivot_WeekLoads, data.WorkLoad_FilePath_ToStore + "\" + File_Name, 'Sheet', 'Mean_WeekLoad');
    writetable(WeelyMainTasks, data.WorkLoad_FilePath_ToStore + "\" + File_Name, 'Sheet', 'WeeklyMainTasks');                
    writetable(WorkingHours_PerColleage_PerWeek, data.WorkLoad_FilePath_ToStore + "\" + File_Name, 'Sheet', 'WorkingHours');                
end