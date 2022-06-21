function [values] = askforyearmonth(values,condition)

load Temp\list.mat list

if strcmp(condition,'SO') == 1
    message_y = 'Sales order xmls: Select a year';
    message_m = 'Sales order xmls: Select a month';
elseif strcmp(condition,'SS') == 1    
    message_y = 'Sales shipment xmls: Select a year';
    message_m = 'Sales shipment xmls: Select a month';
elseif strcmp(condition,'ShippedCases') == 1    
    message_y = 'Shipped cases to production sites: Select a year';
    message_m = 'Shipped cases to production sites: Select a month';  
elseif strcmp(condition,'CheckOMSStatus') == 1
    message_y = 'Checking OMS statusses: Select a year';
    message_m = 'Checking OMS statusses: Select a month';
elseif strcmp(condition,'GenerateRSLabels') == 1
    message_y = 'Generation RS-labels: Select a year';
    message_m = 'Generation RS-labels: Select a month';
elseif strcmp(condition,'StockCorrections') == 1
    message_y = 'Stock corrections: Select a year';
    message_m = 'Stock corrections: Select a month';
end

% Get the year index
[indx,tf] = listdlg('PromptString',message_y,'SelectionMode','single','ListString',list.years,'ListSize',[500 300]);
values.RequestedYear = char(list.years(indx));

% Get the month index
[indx,tf] = listdlg('PromptString',message_m,'SelectionMode','single','ListString',list.months,'ListSize',[500 300]);
% Compensate for the first line in the list.
indx = indx - 1;
if indx < 10
    values.RequestedMonth = ['0' num2str(indx)];
else
    values.RequestedMonth = num2str(indx);
end

% IF generation of RS labels, ask for the initial and final numbers
goon = 0;
if strcmp(condition,'GenerateRSLabels') == 1
    while goon == 0
        prompt = {'Initial number','Final number'};
        title = 'Enter the desired number range for the RS-labels';
        dims = [1 70];
        definput = {'',''};
        answer = inputdlg(prompt,title,dims,definput);
        values.InitialRSlabelnr = str2double(char(answer(1,1)));
        values.FinalRSlabelnr = str2double(char(answer(2,1)));
        if values.InitialRSlabelnr <= values.FinalRSlabelnr
            goon = 1;
        else
            % Check is the data is correct.
            message = 'The initial number may not be higher than the final number! Please provide correct input.';
            choice = questdlg(message,'Caution','Ok','Also ok','Ok');
            switch choice
                case 'Ok'
                    % Donothing = 1;
                case 'Also ok'
                    % Donothing = 1;
            end
        end
    end
end


end