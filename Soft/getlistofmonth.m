function [invoice] = getlistofmonth(onlinereport,stringstofetch_invoice,year,month,condition,customer)

% year and month => must be integers
% condition => 'created' or 'shipped'
% customer => if you want to filter out the data of one specific customer. If not, leave a blank string ('')

months = {'January','February','March','April','May','June','July','August','September','October','November','December'};
month_str = char(months(1,month));
year_str = num2str(year);
disp(['Creating invoices - please wait - gathering data for cases ' condition ' in ' month_str ' ' year_str]);
nrofcases_OR = size(onlinereport,1);
casecounter = 0;

if strcmp(customer,'') == 0
    specificcustomer = 1;
else 
    specificcustomer = 0;
end
customer = [customer ' Company'];

% Get the raw data for the cases shipped in the desired month
if strcmp(condition,'shipped') == 1
    for currentcase = 2:nrofcases_OR
        if (specificcustomer == 1 && strcmp(cell2mat(onlinereport(currentcase,cell2mat(stringstofetch_invoice(2,9)))),customer) == 1) || specificcustomer == 0 % Check if for specific customer or not
            if strcmp(char(onlinereport(currentcase,cell2mat(stringstofetch_invoice(2,4)))),'//') == 0 % first check on case cancelled
                if cell2mat(onlinereport(currentcase,cell2mat(stringstofetch_invoice(2,14)))) == year
                    if cell2mat(onlinereport(currentcase,cell2mat(stringstofetch_invoice(2,15)))) == month
                        if isnan(cell2mat(onlinereport(currentcase,cell2mat(stringstofetch_invoice(2,12))))) % double check on case cancelled
                            casecounter = casecounter + 1;
                            invoice.raw(casecounter,:) = onlinereport(currentcase,:);
                        end
                    end
                end
            end
        end
    end
end
% Get the raw data for the cases created in the desired month
if strcmp(condition,'created') == 1
    for currentcase = 2:nrofcases_OR
        if (specificcustomer == 1 && strcmp(cell2mat(onlinereport(currentcase,cell2mat(stringstofetch_invoice(2,9)))),customer) == 1) || specificcustomer == 0 % Check if for specific customer or not
            if cell2mat(onlinereport(currentcase,cell2mat(stringstofetch_invoice(2,19)))) == year
                if cell2mat(onlinereport(currentcase,cell2mat(stringstofetch_invoice(2,20)))) == month
                    if isnan(cell2mat(onlinereport(currentcase,cell2mat(stringstofetch_invoice(2,12))))) % double check on case cancelled
                        casecounter = casecounter + 1;
                        invoice.raw(casecounter,:) = onlinereport(currentcase,:);
                    end
                end
            end
        end
    end
end

if exist('invoice')
    donothing = 1;
else 
    invoice = 0;
end

end