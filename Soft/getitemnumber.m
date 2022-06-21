function [itemnr,cnr] = getitemnumber(caseid,headersOR,customers,UPSfile_customer,lineOR)

% lineOR (content of a specific line of the online report - not a linenr) is optional.

% load Temp\values.mat values

% If four inputs provided, no need to read the online report
if nargin == 5
    donothing = 1;
    % If only three inputs provided, read the online report
elseif nargin == 4
    load Temp\onlinereport.mat onlinereport
    temp = catchrowindex({caseid},onlinereport(:,2),1);
    rowccOR = cell2mat(temp(2,1));
    lineOR = onlinereport(rowccOR,:);
end



% Get the customernumber
col_hospitalname = catchcolumnindex({'HospitalName'},headersOR,1);
col_hospitalname = cell2mat(col_hospitalname(2,1));
col_deliveryofficename = catchcolumnindex({'DeliveryOfficeName'},headersOR,1);
col_deliveryofficename = cell2mat(col_deliveryofficename(2,1));
col_DeliveryOfiiceAccountNumber = catchcolumnindex({'DeliveryOfiiceAccountNumber'},headersOR,1);
col_DeliveryOfiiceAccountNumber = cell2mat(col_DeliveryOfiiceAccountNumber(2,1));
col_CustomerSurname = catchcolumnindex({'CustomerSurname'},headersOR,1);
col_CustomerSurname = cell2mat(col_CustomerSurname(2,1));
cusdata.company = char(lineOR(1,col_hospitalname));
cusdata.delivery = char(lineOR(1,col_deliveryofficename));
% if strcmp(cusdata.company,'Materialise HQ Company') == 1
%     cnr = 'C0202';
% elseif strcmp(cusdata.company,'SJV Medical Products Company') == 1
%     cnr = 'C0015';
% elseif strcmp(cusdata.company,'Munkebo Fodterapi Bente Frederiksen Company') == 1
%     cnr = 'C1157';
% elseif strcmp(cusdata.company,'OS Funesco GCV Company') == 1
%     cnr = 'C1031';
% else
%     if length(cusdata.delivery) > 9
%         if strcmp(cusdata.delivery(1:9),'OBSOLETE_') == 1
%             if strcmp(cusdata.delivery(10:end),'Be Balanced Be Fit') == 1
%                 cusdata.delivery = 'Be Balanced Be Fit BVBA';
%             else
%                 cusdata.delivery = cusdata.delivery(10:end);
%             end
%         end
%     end
%     [cnr] = getcustomernumber(cusdata);
%     cnr = strrep(cnr,'-','_');
% end
if strcmp(cusdata.company,'Podologie Rios Company') == 1 && strcmp(cusdata.delivery,'Otkas BVBA') == 1
    cusdata.delivery = 'Podologie Rios';
end
if strcmp(cusdata.company,'SW Podiatry LTD Company') == 1 && strcmp(cusdata.delivery,'SW Podiatry LTD Company') == 1
    cusdata.delivery = 'SW Podiatry LTD';
end
if strcmp(cusdata.company,'Cheshire Movement Clinic Company') == 1 && strcmp(cusdata.delivery,'Cheshire Movement Clinic Company') == 1
    cusdata.delivery = 'Cheshire Movement Clinic';
end
if strcmp(cusdata.company,'Orthotic Centre Active Company') == 1 && strcmp(cusdata.delivery,'Orthotic Centre Active Company') == 1
    cusdata.delivery = 'Orthotic Centre Active';
end

if cell2mat(lineOR(1,1)) >= 13416
    [cnr] = getcustomernumber(cusdata,UPSfile_customer);
    cnr = strrep(cnr,'-','_');
else
    cnr = char(lineOR(1,col_DeliveryOfiiceAccountNumber));
    cnr = strrep(cnr,' ','_');
    if strcmp(cnr,'-') == 1 || strcmp(cnr,'') == 1 || strcmp(cnr,'-') == 1 || isempty(cnr) == 1
        cnr = 'CXXXX';
    end
end





if cell2mat(lineOR(1,1)) >= 13416
    if cell2mat(lineOR(1,1)) == 105707
        test = 1;
    end
    % Get the surname and check the content
    temp_surname = lineOR(1,col_CustomerSurname);
    if isa(temp_surname{1},'double') == 1
        temp_surname = num2str(cell2mat(temp_surname));
    elseif temp_surname{1} == 1
        temp_surname = 'True';
    else
        temp_surname = char(temp_surname);
    end
    % Insert here the cases that have issues for some reason
    if sum(contains({'RS18-GEN-QEF','RS18-KOR-QEZ','RS19-NOK-SIK','RS19-EXU-QEN','RS20-SAT-IQI', 'RS20-QOT-VAN', 'RS20-DOH-TAN', ...
            'RS20-KUX-EHA','RS20-EBU-SOV','RS20-ELI-NAG','RS20-ZUP-UXA','RS20-ONO-AVA','RS20-ITE-ABI','RS20-VUB-POC','RS21-KIT-DAZ', ...
            'RS21-NIV-DUD','RS21-RAP-IVO','RS21-ROZ-ZOL','RS21-LEQ-KOB','RS21-PEJ-GAG','RS21-PAJ-FOZ','RS21-ADA-QUK','RS21-HAP-LOM', ...
            'RS21-QEF-QUD','RS21-MET-AHE','RS22-UXE-KUJ','RS22-VUF-LIZ','RS22-HUD-NIS'},caseid)) == 1 ...
            || strcmp(caseid(1:4),'FAKE') == 1
        
        itemnr = ['S' '10' '00' '00' '00000' '00' 'N' '00' '000'];
    elseif strcmp(temp_surname,'CleanData') == 1
        itemnr = ['S' '10' '00' '00' '00000' '00' 'N' '00' '000'];
    else
        
        temp = catchcolumnindex({'CaseId'},headersOR,1);
        caseidmtls = cell2mat(lineOR(1,cell2mat(temp(2,1))));
        
        % AA - Get the info of the printed part
        temp = catchcolumnindex({'BuiltActor'},headersOR,1);
        builtactor = char(lineOR(1,cell2mat(temp(2,1))));
        
        temp = catchcolumnindex({'InsoleType'},headersOR,1);
        insoletype = char(lineOR(1,cell2mat(temp(2,1))));
        
        temp = catchcolumnindex({'Base_type'},headersOR,1);
        basetype = char(lineOR(1,cell2mat(temp(2,1))));
        
        temp = catchcolumnindex({'HospitalRegion'},headersOR,1);
        hopsitalregion = char(lineOR(1,cell2mat(temp(2,1))));
        
        % Define the insole type
        % Location is defined here, but also in createnavisionsalesorderinput line 88 => double check this!
        [A2] = getinsoletype(insoletype,basetype);
        if strcmp('MPROD US Materialise',builtactor) == 1 ...
                || (strcmp('Flowbuilt Production',builtactor) == 1 && caseidmtls > 47780) ...
                || (strcmp('Flowbuilt Production_temp',builtactor) == 1 && caseidmtls > 47780) ...
                || (strcmp('USA',hopsitalregion) == 1 && caseidmtls > 47780) ...
                || (strcmp('CAN',hopsitalregion) == 1 && caseidmtls > 47780)
            location = 'USPROD';
            AA = ['2' A2];
        elseif strcmp('-',builtactor) == 1
            location = 'Not known yet.';
            AA = ['X' A2];
        elseif strcmp('EU/AP/LAT',hopsitalregion) == 1
            location = 'BEPROD';
            AA = ['1' A2];
        else
            location = 'error';
            AA = ['X' A2];
        end
        
        % Check if serial number needs to be defined or overruled
        if strcmp(upper(customers.(cnr).NoTopLayerSN),'YES') == 1 % If something goes wrong on this line, the combination of OMS company and/or office is not recognized. Please check the UPS Excel.
            itemnr = ['S' AA '00' '00' '00000' '00' 'N' '00' '000'];
        else
            % Get the assembly parameters
            temp = catchcolumnindex({'ServiceType'},headersOR,1);
            assembly = char(lineOR(1,cell2mat(temp(2,1))));
            if strcmp(assembly,'No top layer') == 0
                
                % BB - Get the top layer information
                
                BB_str = {'10','11','20','21','30'}; % 22 (smooth) is currently not an option, as it cannot be chosen in the DW. See exceptions below.
                BB_label = {'EVA','EVA Synthetic Leather','PU Soft','PU Soft Synthetic Leather','EVA Carbon'};
                temp = catchcolumnindex({'Top_material'},headersOR,1);
                material = char(lineOR(1,cell2mat(temp(2,1))));
                temp = catchcolumnindex({material},BB_label,1);
                BB = char(BB_str(1,cell2mat(temp(2,1))));
                
                CC_str = {'20','30','40','60','B1'};
                CC_label = {'2 mm','3 mm','4 mm','6 mm','22 mm block'};
                temp = catchcolumnindex({'Top_thickness'},headersOR,1);
                topthickness = char(lineOR(1,cell2mat(temp(2,1))));
                temp = catchcolumnindex({topthickness},CC_label,1);
                CC = char(CC_str(1,cell2mat(temp(2,1))));
                
                DDDDD_str = {'00080','00090','00100','00110','00120','00130','10010','10020','10030','10040','10050','10060','10070','10080','10090','10100','10110','10120','10130','10140','10150','10160','10170','10180','10190','10200','10210','10220','99999'};
                DDDDD_label = {'8.0 kids UK','9.0 kids UK','10.0 kids UK','11.0 kids UK','12.0 kids UK','13.0 kids UK','1.0 UK','2.0 UK','3.0 UK','4.0 UK','5.0 UK','6.0 UK','7.0 UK','8.0 UK','9.0 UK','10.0 UK','11.0 UK','12.0 UK','13.0 UK','14.0 UK','15.0 UK','16.0 UK','17.0 UK','18.0 UK','19.0 UK','20.0 UK','21.0 UK','22.0 UK','sheet'};
                temp = catchcolumnindex({'Top_size'},headersOR,1);
                topsize = char(lineOR(1,cell2mat(temp(2,1))));
                temp = catchcolumnindex({topsize},DDDDD_label,1);
                if strcmp(insoletype,'Wide') == 1
                    % Add one size if it is a wide insole type
                    DDDDD = char(DDDDD_str(1,cell2mat(temp(2,1))+1));
                else
                    % Just take the regular size
                    DDDDD = char(DDDDD_str(1,cell2mat(temp(2,1))));
                end
                
                EE_str = {'20','30','35','40','50'};
                EE_label = {'20 Shore','30 Shore','35 Shore','40 Shore','50 Shore'};
                temp = catchcolumnindex({'Top_hardness'},headersOR,1);
                tophardness = char(lineOR(1,cell2mat(temp(2,1))));
                temp = catchcolumnindex({tophardness},EE_label,1);
                EE = char(EE_str(1,cell2mat(temp(2,1))));
                
                %     F_str = {'L','R','P','S','N','D','B'};
                %     F_label = {'left','right','pair','sheet','none','demo','big sheet'};
                %     temp = catchcolumnindex({'Top hardness'},onlinereport,1);
                %     config = char(onlinereport(rowccOR,cell2mat(temp(2,1))));
                %     temp = catchcolumnindex({config},F_label,1);
                %     F = char(F_str(1,cell2mat(temp(2,1))));
                F = 'P';
                
                % define extra toppings
                temp = catchcolumnindex({'Heel_pad___left'},headersOR,1);
                hpl = char(lineOR(1,cell2mat(temp(2,1))));
                temp = catchcolumnindex({'Heel_pad___right'},headersOR,1);
                hpr = char(lineOR(1,cell2mat(temp(2,1))));
                
                % if a heel pad is present
                if strcmp(hpl,'Fat Pad') == 1 || strcmp(hpl,'Plantar Fasciitis') == 1 || strcmp(hpl,'Heel Spur') == 1  || ...
                        strcmp(hpr,'Fat Pad') == 1 || strcmp(hpr,'Plantar Fasciitis') == 1 || strcmp(hpr,'Heel Spur') == 1
                    % if heel pad with EVA finishing, take EVA extra top
                    if strcmp(BB,'10') == 1
                        GG = '01';
                        % if heel pad with synthetich leather finishing, take synthetic leather extra top
                    elseif strcmp(BB,'11') == 1
                        GG = '02';
                        % Reset top layer to regular EVA (we can't use preformed EVA + leather for heel pad finishing)
                        BB = '10';
                        % if heel pad with EVA carbon finishing, take anti-static extra top
                    elseif strcmp(BB,'30') == 1
                        GG = '03';
                    else
                        disp('Fatal error')
                        clear all
                    end
                    
                    % if EVA carbon is chosen
                elseif strcmp(BB,'30') == 1
                    % for now at least, we add a anti-static layer for safety insoles,
                    % but we don't track it..
                    GG = '00';
                    % if a full length insole is ordered
                elseif strcmp(AA,'11') == 1 || strcmp(AA,'21') == 1
                    % if full length insole, take correct top material
                    % CC remains the same
                    % DDDDD remains the same
                    % EE remains the same
                    % F remains the same
                    GG = '00';
                    % if EVA
                    if strcmp(BB,'10') == 1
                        GG = '00';
                        % if EVA synthetic leather
                    elseif strcmp(BB,'11') == 1
                        BB = '11';
                        GG = '00';
                        % if PU soft black fabric
                    elseif strcmp(BB,'20') == 1
                        GG = '00';
                        % if PU soft synthetic leather
                    elseif strcmp(BB,'21') == 1
                        BB = '21';
                        GG = '00';
                    end
                    % HH remains the same
                    
                else
                    GG = '00';
                end
                
                % HHH
                HHH = '000';
                
            else
                % in case of no top layer, everything on zero.
                BB = '00';
                CC = '00';
                DDDDD = '00000';
                EE = '00';
                F = 'N';
                GG = '00';
                HHH = '000';
            end
            
            % Temporary solution for FB assembly (SLS printed part + assembly done by FB)
            if strcmp('Flowbuilt Production',builtactor) == 1 && caseidmtls < 47781
                itemnr = ['S1' AA(2) BB CC DDDDD EE F GG HHH(1:2) 'A'];
            else
                itemnr = ['S' AA BB CC DDDDD EE F GG HHH];
            end
        end
    end
else
    itemnr = 'Not available';
end


end