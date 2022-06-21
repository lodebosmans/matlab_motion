function createlabels()

load Temp\caseids_production_tocreate.mat caseids_production_tocreate

if size(caseids_production_tocreate,1) > 0
    
    load Temp\values.mat values
    load Temp\db_production.mat db_production
    %load Temp\caseids_production.mat caseids_production
    load Temp\customers.mat customers
    
    %nrofcases = size(caseids_production,1);
    nrofcases = size(db_production.overview,1);
    
    % % Create text file for printing label code
    % fid = fopen('Output\Labels\results_labels.txt','wt');
    % fprintf(fid, '${');
    % fprintf(fid, '\n');
    
    % Create for loop to iterate over all case IDs
    for caseindex = 1:nrofcases
        donotprintaddress = 0;
        displaytext = ['Creating labels - please wait - processing ' num2str(caseindex) ' of ' num2str(nrofcases)];
        disp(displaytext);
        % Get the caseID from the list of caseIDs
        clear currentcaseid
        %currentcaseid_long = char(caseids_production(caseindex,1));
        currentcaseid_long = char(db_production.overview(caseindex,1));
        currentcaseid = [currentcaseid_long(1:4) currentcaseid_long(6:8) currentcaseid_long(10:12)];
        % Get the customer number to define if the address should be displayed or not.
        cnr = db_production.raw.(currentcaseid).customernumber;
        cnr = strrep(cnr,'-','_');
        if strcmp(upper(customers.(cnr).NoAddressLabel),'YES') == 1 %#ok<STCI>
            donotprintaddress = 1;
        end
        
        % Orientation for 180° rotated (^POI) print (alphabetical order)
        %     if strcmp(values.CountryOfShipment,"BE") == 1 % dimensions 102x38 mm (203 dpi)
        code.part1 = ['^XA^CFT^POI^FO60,20,0^FD' db_production.raw.(currentcaseid).caseid.full '^FS^FWZ,1^CFQ^FO560,78,1^FD' db_production.raw.(currentcaseid).company '^FS^CFS^FO560,106,1^FD' db_production.raw.(currentcaseid).referenceid '^FS^CFS^FO560,20^FD' strrep(db_production.raw.(currentcaseid).customernumber,'_','-') '^FS^FWZ,0'];
        code.part2 = '^FO60,148^GB10,1,2^FS^FO80,148^GB10,1,2^FS^FO100,148^GB10,1,2^FS^FO120,148^GB10,1,2^FS^FO140,148^GB10,1,2^FS^FO160,148^GB10,1,2^FS^FO180,148^GB10,1,2^FS^FO200,148^GB10,1,2^FS^FO220,148^GB10,1,2^FS^FO240,148^GB10,1,2^FS^FO260,148^GB10,1,2^FS^FO280,148^GB10,1,2^FS^FO300,148^GB10,1,2^FS^FO320,148^GB10,1,2^FS^FO340,148^GB10,1,2^FS^FO360,148^GB10,1,2^FS^FO380,148^GB10,1,2^FS^FO400,148^GB10,1,2^FS^FO420,148^GB10,1,2^FS^FO440,148^GB10,1,2^FS^FO460,148^GB10,1,2^FS^FO480,148^GB10,1,2^FS^FO500,148^GB10,1,2^FS^FO520,148^GB10,1,2^FS^FO540,148^GB10,1,2^FS';
        if donotprintaddress == 0
            code.part3 = ['^CFS^FO60,155,0^FD' db_production.raw.(currentcaseid).delivery_company '^FS^CFQ^FO60,195,0^FD' db_production.raw.(currentcaseid).delivery_street '^FS^FO60,225,0^FD' db_production.raw.(currentcaseid).delivery_zip_city '^FS^FO60,255,0^FD' db_production.raw.(currentcaseid).delivery_country '^FS'];
        else
            code.part3 = '';
        end
        code.part4 = ['^CFQ^FO620,260^FDMade in Belgium^FS^CFS^FWZ,1^FO560,250,1^FD' db_production.raw.(currentcaseid).estimatedshippingdate '^FS^FWZ,0^FO595,20^BQN,2,10^FDQA,' db_production.raw.(currentcaseid).caseid.full];
        %         if caseindex < nrofcases
        %             code.part5 = '^FS^XB^XZ';
        %         end
        %         if caseindex == nrofcases
        %             code.part5 = '^FS^XZ';
        %         end
        %     elseif strcmp(values.CountryOfShipment,"US") == 1 % dimensions 143x53 mm (300 dpi)
        %         code.part1 = ['^XA^CFV^POI^FO60,25,0^FD' db_production.raw.(currentcaseid).caseid.full '^FS^FWZ,1^CFS^FO840,110,1^FD' db_production.raw.(currentcaseid).company '^FS^CFS^FO840,152,1^FD' db_production.raw.(currentcaseid).referenceid '^FS^CFT^FO840,35^FD' strrep(db_production.raw.(currentcaseid).customernumber,'_','-') '^FS^FWZ,0'];
        %         code.part2 = '^FO60,200^GB10,1,2^FS^FO80,200^GB10,1,2^FS^FO100,200^GB10,1,2^FS^FO120,200^GB10,1,2^FS^FO140,200^GB10,1,2^FS^FO160,200^GB10,1,2^FS^FO180,200^GB10,1,2^FS^FO200,200^GB10,1,2^FS^FO220,200^GB10,1,2^FS^FO240,200^GB10,1,2^FS^FO260,200^GB10,1,2^FS^FO280,200^GB10,1,2^FS^FO300,200^GB10,1,2^FS^FO320,200^GB10,1,2^FS^FO340,200^GB10,1,2^FS^FO360,200^GB10,1,2^FS^FO380,200^GB10,1,2^FS^FO400,200^GB10,1,2^FS^FO420,200^GB10,1,2^FS^FO440,200^GB10,1,2^FS^FO460,200^GB10,1,2^FS^FO480,200^GB10,1,2^FS^FO500,200^GB10,1,2^FS^FO520,200^GB10,1,2^FS^FO540,200^GB10,1,2^FS^FO560,200^GB10,1,2^FS^FO580,200^GB10,1,2^FS^FO600,200^GB10,1,2^FS^FO620,200^GB10,1,2^FS^FO640,200^GB10,1,2^FS^FO660,200^GB10,1,2^FS^FO680,200^GB10,1,2^FS^FO700,200^GB10,1,2^FS^FO720,200^GB10,1,2^FS^FO740,200^GB10,1,2^FS^FO760,200^GB10,1,2^FS^FO780,200^GB10,1,2^FS^FO800,200^GB10,1,2^FS^FO820,200^GB10,1,2^FS';
        %         if donotprintaddress == 0
        %             code.part3 = ['^CFT^FO60,220,0^FD' db_production.raw.(currentcaseid).delivery_company '^FS^CFS^FO60,270,0^FD' db_production.raw.(currentcaseid).delivery_street '^FS^FO60,310,0^FD' db_production.raw.(currentcaseid).delivery_zip_city '^FS^FO60,350,0^FD' db_production.raw.(currentcaseid).delivery_country '^FS'];
        %         else
        %             code.part3 = '';
        %         end
        %         code.part4 = ['^FO960,320^FDMade in^FS^FO925,360^FDUnited States^FS^CFT^FWZ,1^FO1100,35,1^FD' db_production.raw.(currentcaseid).estimatedshippingdate '^FS^FWZ,0^FO910,88^BQN,2,10^FDQA,' db_production.raw.(currentcaseid).caseid.full];
        %         if caseindex < nrofcases
        %             code.part5 = '^FS^XB^XZ';
        %         end
        %         if caseindex == nrofcases
        %             code.part5 = '^FS^XZ';
        %         end
        %     end
        
        %     fprintf(fid, code.part1);
        %     fprintf(fid, code.part2);
        %     fprintf(fid, code.part3);
        %     fprintf(fid, code.part4);
        %     fprintf(fid, code.part5);
        %     fprintf(fid, '\n');
        
        % Print the label also as a separate file
        fid2 = fopen([values.finrepfolder  currentcaseid '.txt'],'wt');
        fprintf(fid2, '${');
        fprintf(fid2, code.part1);
        fprintf(fid2, code.part2);
        fprintf(fid2, code.part3);
        fprintf(fid2, code.part4);
        fprintf(fid2, '^FS^XZ}$');
        fclose(fid2);
        
        clear currentcaseid
        clear code
    end
    % fprintf(fid, '}$');
    % fclose(fid);
    
end

end