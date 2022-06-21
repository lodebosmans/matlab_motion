function [agreements] = addagreement(agreements,cafull,heardersagr,cnr,agreementnr)

% agreements = struct with agreements (already present)
% cafull = current agreement full line
% heardersagr = the headers of the agreement data
% cnr = current customer number (with underscores if necessary)
% agreementnr = number of the agreement that is being processed. Starts with 1 for the first one

nrofcols = length(heardersagr);
% Write the data to the struct
for x = 1:nrofcols
    cpar = char(heardersagr(1,x));
    cval = cafull(1,x);
    if iscellstr(cval) == 0
        eval(['agreements.' cnr '.PriceAgreements.Agreement' num2str(agreementnr) '.' cpar ' = cell2mat(cval);']);
    else
        eval(['agreements.' cnr '.PriceAgreements.Agreement' num2str(agreementnr) '.' cpar ' = char(cval);']);
    end
end

if strcmp(cnr,'C1019_002') == 1
    sqfsd = 1;
end

% ShipTo
% Check if a ShipTo is defined.
if strcmp('-',eval(['agreements.' cnr '.PriceAgreements.Agreement' num2str(agreementnr) '.ShipTo'])) == 1
    % If not, define it.    
    eval(['agreements.' cnr '.PriceAgreements.Agreement' num2str(agreementnr) '.ShipTo = cnr;']);
    eval(['agreements.' cnr '.PriceAgreements.Agreement' num2str(agreementnr) '.ShipTo = strrep(agreements.' cnr '.PriceAgreements.Agreement' num2str(agreementnr) '.ShipTo,''_'',''-'');']);
    eval(['agreements.' cnr '.PriceAgreements.Agreement' num2str(agreementnr) '.ShipTo(1) = ''D'' ;']);
else
    % If yes, do nothing. The ShipTo is already defined in the Excel file.    
end

% Delivery only + maybe change customernumber
DeliveryOnly = eval(['agreements.' cnr '.PriceAgreements.Agreement' num2str(agreementnr) '.DeliveryOnly']);
if strcmp('-',DeliveryOnly) == 1
    % If no delivery only, do nothhing.
elseif strcmp('Yes',DeliveryOnly) == 1
    % Change the customer number to the main account.
    eval(['agreements.' cnr '.PriceAgreements.Agreement' num2str(agreementnr) '.CustomerNumber = cnr(1:5);' ])

    % InvoiceTo
    % Check if a InvoiceTo is defined.
    if strcmp('-',eval(['agreements.' cnr '.PriceAgreements.Agreement' num2str(agreementnr) '.InvoiceTo'])) == 1
        % If not, define it.
        eval(['agreements.' cnr '.PriceAgreements.Agreement' num2str(agreementnr) '.InvoiceTo = cnr(1:5);']);
    else
        % If yes, do nothing. The InvoiceTo is already defined in the Excel file.
    end
    
else
    disp('DeliveryOnly input is not valid.');
    clear all
end





end