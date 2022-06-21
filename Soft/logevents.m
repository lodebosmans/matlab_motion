function logevents(phase)

load Temp\values.mat values

disp('Logging events - please wait');

% Check if Lode is messing around or not
if values.TestDebug == 1
    testdebug = 'TEST ';
else
    testdebug = '';
end

timestring = ['20' values.y '/' values.mo '/' values.d ' ' values.h 'h' values.mi 'm' values.s   's: '];

% Write all cases into one batch.


% Go over all individual cases
if values.LogEvents == 1
    % Get the event that is occuring at this moment.
    if values.FinRepLabels == 1
        procedure = 'Finishing report + label created';
    elseif values.ManDelNote == 1
        procedure = 'Manual delivery noted created';
    elseif values.SortOutDelNoteWSxml == 1
        procedure = 'Shipment sort out + delivery note + WorldShip xmls + FDA label created';
    elseif values.FDAlabels == 1
        procedure = 'FDA label created';    
    end
    
    load Temp\caseids_production.mat caseids_production
    nrofcases = size(caseids_production,1);
    for cr = 1:nrofcases
        ccid = char(caseids_production(cr,1));        
        writelogtofile(ccid,timestring,phase,testdebug,values,procedure)
    end
end
if values.DailySnapshot == 1 
    nrofprocedures = size(values.snapshotprocedures,2);
    for cp = 1:nrofprocedures
        save Temp\cp.mat cp
        procedure = char(values.snapshotprocedures_text(1,cp));
        readcaseids()
        load Temp\caseids_production.mat caseids_production
        if strcmp(phase,'Initiation') == 1
            % Write entire snapshot to file
            writelogtofile_snapshot(caseids_production,timestring,testdebug,values,procedure);
        end
        % Write individual cases to file
        nrofcaseids = size(caseids_production,1);        
        for x = 1:nrofcaseids
            ccid = char(caseids_production(x,1));            
            writelogtofile(ccid,timestring,phase,testdebug,values,procedure);
        end        
    end    
end

end