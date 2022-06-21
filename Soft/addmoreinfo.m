function [info] = addmoreinfo(info,condition)

load Temp\values.mat values

if strcmp(condition,'sortshipmentexcels') == 1
    info.newfile = [values.shipmentexcelfolder 'New\' info.filename];
    info.destinationfolder_sorted = [values.shipmentexcelfolder 'Sorted\' info.destination '\' info.year '\' info.month];
    info.destinationfolder_all = [values.shipmentexcelfolder 'All'];
 
    if isfile([info.destinationfolder_sorted '\' info.filename]) == 1
        info.alreadypresent = 1;
    else
        info.alreadypresent = 0;
    end
elseif strcmp(condition,'sortshipmentexcels_FB') == 1
    info.newfile = [values.shipmentexcelfolder 'New\' info.filename];
    info.destinationfolder_sorted = [values.shipmentexcelfolder 'Sorted\' info.destination '\' info.year '\' info.month];
    info.destinationfolder_all = [values.shipmentexcelfolder 'All'];
 
    if isfile([info.destinationfolder_sorted '\' info.filename]) == 1
        info.alreadypresent = 1;
    else
        info.alreadypresent = 0;
    end
elseif strcmp(condition,'linkinvoiceitocasesshipments') == 1
    info.newfile = [values.invoicesfolder 'New\' info.filename];
    info.destinationfolder_sorted = [values.invoicesfolder 'Sorted\' info.year];
    
    if isfile([info.destinationfolder_sorted '\' info.filename]) == 1
        info.alreadypresent = 1;
    else
        info.alreadypresent = 0;
    end
elseif strcmp(condition,'redundant') == 1
    info.newfile = [values.invoicesfolder 'New\OLD doc\redundant\' info.filename];
    info.destinationfolder_sorted = [values.invoicesfolder 'Sorted\' info.year];
    
    if isfile([info.destinationfolder_sorted '\' info.filename]) == 1
        info.alreadypresent = 1;
    else
        info.alreadypresent = 0;
    end
elseif strcmp(condition,'prenavisionera') == 1
    info.year = '2018';
    info.month = '00';
    info.newfile = [values.invoicesfolder 'New\OLD doc\' info.filename];
    info.destinationfolder_sorted = [values.invoicesfolder 'Sorted\' info.year];

    if isfile([info.destinationfolder_sorted '\' info.filename]) == 1
        info.alreadypresent = 1;
    else
        info.alreadypresent = 0;
    end
end

end