function [values] = warningmessages(values)

% See who's rocking the show. If no-one, abort everything.
% --------------------------------------------------------
if values.Username_nr  == 1
    errordlg('Sorry, no can do! Please select your name first!');  
    clear all
else
    disp(['Great work, ' values.Username_str '!']);
end

if values.NavisionFunction_nr > 1 && values.DesiredFunction_nr > 1
    errordlg('Sorry, no can do! Select maximum one function to perform! If this error keeps occuring, close and restart the interface.');
    clear all
end

if strcmp(values.Username_str,'testversion') == 1 && strcmp(values.Username_str,'Lode') == 0
    errordlg('This is the Matlab Test Environment. Please close this version and open the regular Matlab version.');
    clear all
end

% % Check the destination country
% % -----------------------------
% if strcmp(values.Username_str,'Lode') ...
%    || strcmp(values.Username_str,'Pieter-Jan') ...
%    || strcmp(values.Username_str,'Pieter') ...
%    || strcmp(values.Username_str,'Natascha')
%    
%     choice = questdlg('Country of production and shipment is BE?','Caution','Yes','No, switch to US','Yes');
%     switch choice
%         case 'Yes'
%             values.CountryOfShipment = 'BE';
%         case 'No, switch to US'
%             values.CountryOfShipment = 'US';
%     end
% elseif strcmp(values.Username_str,'Trevor') || strcmp(values.Username_str,'Spencer') || strcmp(values.Username_str,'JAred')
%     values.CountryOfShipment = 'US';
% else 
%     values.CountryOfShipment = 'BE';
% end
values.CountryOfShipment = 'BE';

% Generate some warning messages if necessary
% -------------------------------------------------------------------------
if values.SortOutDelNoteWSxml == 1
    choice = questdlg('Is the UPS excel closed and are you logged on to the network?','Caution','Yes','No','Yes');
    switch choice
        case 'Yes'
            % Donothing = 1;
        case 'No'
            disp('Mayday, Mayday! Aborting mission!')
            clear all
    end
end
if strcmp(values.version,'testversion') == 1
    choice = questdlg('Did you copy the SOC, agreements and UPS file to the testenvironment?','Caution','Yes','No','Yes');
    switch choice
        case 'Yes'
            % Donothing = 1;
        case 'No'
            disp('Mayday, Mayday! Aborting mission!')
            clear all
    end
end

if values.OpenLogFile == 1 && values.NewCaseIDsPresent == 0
    errordlg('You want a logfile and you do not provide case IDs? Not a very smart move :)');
end

end