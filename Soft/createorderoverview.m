function createorderoverview(values) %#ok<INUSD>

load Temp\customers customers  
load Temp\onlinereport.mat onlinereport 

disp('Mapping all cases to the customers - please wait')

% purpose: list all cases for the customers and add them to the customer variable

% Go over all cases in OR
% Select all cases from the last two months
% Link the every case to the correct customer number (based on invoice number)
% Add case to customer
% Count everything

[y,mo,d,h,mi,s] = getdate(); %#ok<ASGLU>

yc = str2num(['20' y]);
moc = str2num(mo);

mop = moc - 1;
yp = yc;
if mop == 0
    mop = 12;
    yp = yc - 1;
end

nrofrowsOR = size(onlinereport,1);
headersOR = onlinereport(1,:);
counter = 0;
for cr = 2:nrofrowsOR
    lineOR = onlinereport(cr,:);
    temp = catchcolumnindex({'CreatedYear'},headersOR,1);
    ccy = cell2mat(lineOR(1,cell2mat(temp(2,1))));    
    temp = catchcolumnindex({'CreatedMonth'},headersOR,1);
    ccmo = cell2mat(lineOR(1,cell2mat(temp(2,1))));    
%     temp = catchcolumnindex({'CreatedDay'},headersOR,1);
%     ccd = cell2mat(onlinereport(cr,cell2mat(temp(2,1))));
        
    if (ccy == yc || ccy == yp) && (ccmo == moc || ccmo ==  mop) 
        counter = counter + 1;
        % Get caseID
        temp = catchcolumnindex({'CaseCode'},headersOR,1);
        cccaseID = cell2mat(lineOR(1,cell2mat(temp(2,1))));
        salesorderinput.overview()
        
        % Get company name
        temp = catchcolumnindex({'HospitalName'},headersOR,1);
        cccompany = cell2mat(lineOR(1,cell2mat(temp(2,1))));
        
        % Get delivery facility
        temp = catchcolumnindex({'DeliveryOfficeName'},headersOR,1);
        ccdelivery = cell2mat(lineOR(1,cell2mat(temp(2,1))));
        
        % Get customernumber
        
        % Get itemnumber
        [itemnr] = getitemnumber(values,caseid,headersOR,lineOR);
        orders.overview()
        

        
        
    end
end



names = fieldnames(customers);


s.a.f1= {'Sunday'  'Monday'}
s.f2= {'Sunday'  'Monday'  'Tuesday'}
s.f3= 'Tuesday'
s.f4= 'Wednesday'
s.f5= 'Thursday'
s.f6= 'Friday'
s.f7= 'Saturday'
index2 = structfun(@(x) any(strcmp(x, 'Monday')),s)




end
