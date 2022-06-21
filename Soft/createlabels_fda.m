function createlabels_fda()

load Temp\values.mat values
load Temp\customers.mat customers
load Temp\db_production.mat db_production
load Temp\caseids_production.mat caseids_production
copyfile([values.salesordercheck_folder 'salesordercheck_v3.mat'],'Temp\salesordercheck_v3.mat'); % always keep full pathway
load Temp\salesordercheck_v3.mat salesordercheck_v3

temp = catchcolumnindex({'CaseID'},salesordercheck_v3.headers,1);
col_caseid = cell2mat(temp(2,1));
temp3 = catchcolumnindex({'ItemNumber'},salesordercheck_v3.headers,1);
col_itemnr = cell2mat(temp3(2,1));

UScases = cell(0);
UScases_counter = 0;

nrofcases = size(db_production.overview,1);
% Create for loop to iterate over all case IDs and filter out the ones with destination US
for caseindex = 1:nrofcases
    % Get the caseID from the list of caseIDs
    clear currentcaseid
    currentcaseid_long = char(db_production.overview(caseindex,1));
    if length(currentcaseid_long) == 12
        currentcaseid = [currentcaseid_long(1:4) currentcaseid_long(6:8) currentcaseid_long(10:12)];
        % Get destination country
        destinationcountry = customers.(db_production.raw.(currentcaseid).customernumber).DeliveryAddress.DeliveryCountryCode;
        % Get itemnumber
        temp2 = catchrowindex({currentcaseid_long},salesordercheck_v3.data,col_caseid);
        rowcc = cell2mat(temp2(2,1));
        itemnr = char(salesordercheck_v3.data(rowcc,col_itemnr));
        if strcmp(destinationcountry,"US") == 1
            UScases_counter = UScases_counter + 1;
            UScases(UScases_counter,1) = {currentcaseid_long};
            UScases(UScases_counter,2) = {itemnr};
        end
    end
end
clear caseindex

% Create text file for printing label code
fid = fopen('Output\Labels\results_labels_fda.txt','wt');
fprintf(fid, '${');
fprintf(fid, '\n');

if UScases_counter > 0
    % Create for loop to iterate over all case IDs and filter out the ones with destination US
    for caseindex = 1:UScases_counter
        displaytext = ['Creating FDA labels - please wait - processing ' num2str(caseindex) ' of ' num2str(UScases_counter)];
        disp(displaytext);
        currentcaseid_long = char(UScases(caseindex,1));
        itemnr = char(UScases(caseindex,2));
        if strcmp(itemnr(2),'X') == 1
            % If the case still has a temporary itemnumber, the FDA label can't be printed.
            code.part1 = ['^XA^CFO,30^FO70,30^FD' currentcaseid_long '^FS^FO70,70^FD' itemnr '^FS'];
            code.part2 = '^FO70,110^FDFDA label cannot be created.^FS';
            code.part3 = '';
        else
            if strcmp(itemnr(2),'1') == 1
                % If produced in BE
                code.part1 = '^XA^CFO,30^FO71,30^FDphits^FS^FO70,30^FDphits^FS^FO71,29^FDphits^FS^FO71,70^FDcustom orthotics^FS^FO70,70^FDcustom orthotics^FS^FO71,69^FDcustom orthotics^FS^CFO,20^FX Lot and ref boxes^FO70,115^GB70,40,3^FS^FO85,127^FDREF^FS^FO70,165^GB70,40,3^FS^FO85,177^FDLOT^FS^FO150,227^FDQTY-1^FS';
                code.part2 = ['^FX Itemnumber and case ID^FO150,127^FD' itemnr '^FS^FO150,177^FD' currentcaseid_long '^FS'];
                code.part3 = '^FX MTLS address section^FO550,142^FDMaterialise NV^FS^FO550,172^FDTechnologielaan 15^FS^FO550,202^FD3001 Leuven^FS^FO550,232^FDBelgium^FS^FO551,270^FDMADE IN BELGIUM^FS^FO550,270^FDMADE IN BELGIUM^FS^FO551,271^FDMADE IN BELGIUM^FS^FO550,30^FDMANUFACTURED BY^FS^FO150,270^FDL-100352-01^FS^FX MTLS logo^FO550,50^GFA,1775,1775,25,,:hM01,hM0F,hL0FF,hK07FF,hJ03IF,hI03JF,hH01KF,hH0LF,hG07LF,h07MF,gY03NF,gX03OF,gW01PF,gW0QF,gV0RF,gU07RF,gT03SF,gS01TF,gR01UF,gR0VF,gQ07VF,gP07WF,gO03XF,gN03YF,gN07YF,V03PF,V03OF8,V03NF8,V03MFC,V03LFE,V03KFE,V03KF,V03JF8,V03IF8,V03FFC,V03FE,V03E,V03,,::gK06N06,gK0FL0F0F8,:S078P0FL0F0F8,S078P0FL0F07,S078W0F,:00FE1FC007F807FE01FC007F8F01FF00F0700FE003F8,07FF7FF01FFE07FF07FF00FF8F07FF80F0F03FF80FFE,0LF81IF07FF0IF81FF8F07FFC0F0F07FFC1IF,0LF81F3F07FE1F8FC1FF8F03C7E0F0F07C781F1F8,0F83F07CI0F87801E03C3E00FI01E0F0F078103C078,0F01E07CI0787801E01C3E00FI01E0F0F0FI03C038,0F01E03CI0787803C01E3C00FI01E0F0F078007803C,0F01E03C03E787803C01E3C00F0079E0F0F07F007803C,0F01E03C0IF87803IFE3C00F03FFE0F0F03FE07IFC,0F01E03C3IF87803IFE3C00F07FFE0F0F01FF87IFC,0F01E03C3E0F87803IFC3C00F0F83E0F0F007FC7IF8,0F01E03C780787803CI03C00F0F01E0F0FI07C78,0F01E03C780787803CI03C00F0E01E0F0FI03E38,0F01E03C780787801EI03C00F0F01E0F0FI01E3C,0F01E03C7C0787C01F0183C00F0F01E0F0F0703C3E03,0F01E03C3F7F83FF0IF83C00F0FEFE0F0F0IFC1IF,0F01E03C3IF03FF07FFC3C00F07FFE0F0F0IF80IF8,0F01E03C0FFE01FF03FF83C00F03FF80F0F03FF007FF,O03FI07C00FCN0FCM07C001F8,,:^FS^FX Fabriek logo^FO450,140^GFA,585,585,9,,R07IFC,R0JFC,::::::::::::::::::::::00CI0CI0C0JFC,00EI0EI0E0JFC,01EI0E001E0JFC,01E001E001F0JFC,01F003F801F0JFC,03F807F803F8JFC,07F807F807F8JFC,07FC07FC07FCJFC,0FFC07FC0FFEJFC,1FFE0FFC0FFEJFC,1IF1FFE1FFEJFC,1IF1IF1MFC,3IF3IF3MFC,3IF3QFC,3UFC,VFC,::::::::::::::::::::::::^FS^FX Attention logo^FO450,30^GFA,710,710,10,N01C,N03E,N07F,:N0FF8,:M01F7C,M01F3E,M03E3E,M03E1F,M07C1F,M07C0F8,M0F80F8,L01F807C,L01F007C,L03E003E,:L07C001F,:L0F8I0F8,L0F8080F8,K01F03E07C,K01F07F07C,K03E07F03E,:K07C07F01F,:K0F807F00F8,:J01F007F007C,:J03E007F003E,J03E007F003F,J07C007F001F,J07C007FI0F8,J0F8007FI0F8,I01F8007FI07C,I01FI07FI07C,I03FI03FI03E,I03EI03FI03E,I07EI03EI01F,I07CI03EI01F,I0F8I03EJ0F8,:001FJ03EJ07C,:003EJ03EJ03E,:007CJ03EJ01F,007CJ01CJ01F,00F8K08K0F8,00F8Q0F8,01FR07C,01FK01CK07C,03EK03EK03E,07EK03FK03E,07CK07FK01F,0FCK03EK01F,0F8K03EL0F8,1F8L08L0F8,1FT07C,3FT07C,3ET03E,7CT03F,7CT01F,FCT01F8,XF8,::7WF,3VFE,^FS';
            elseif strcmp(itemnr(2),'2') == 1
                % If produced in US
                code.part1 = '^XA^CFO,45^FO56,45^FDphits^FS^FO55,45^FDphits^FS^FO56,44^FDphits^FS^FO56,105^FDcustom orthotics^FS^FO55,105^FDcustom orthotics^FS^FO56,104^FDcustom orthotics^FS^CFO,30^FX Lot and ref boxes^FO55,173^GB105,60,4^FS^FO78,191^FDREF^FS^FO55,248^GB105,60,4^FS^FO78,266^FDLOT^FS^FO175,340^FDQTY-1^FS';
                code.part2 = ['^FX Itemnumber and case ID^FO175,191^FD' itemnr '^FS^FO175,266^FD' currentcaseid_long '^FS'];
                code.part3 = '^FX MTLS address section^FO700,213^FDMaterialise USA LLC^FS^FO700,258^FD44650 Helm Court^FS^FO700,303^FDPlymouth, MI 48170, US^FS^FO701,405^FDMADE IN THE UNITED STATES^FS^FO700,405^FDMADE IN THE UNITED STATES^FS^FO700,406^FDMADE IN THE UNITED STATES^FS^FO700,45^FDMANUFACTURED BY^FS^FO175,405^FDL-100352-01^FS^FX MTLS logo^FO700,100^GFA,1775,1775,25,,:hM01,hM0F,hL0FF,hK07FF,hJ03IF,hI03JF,hH01KF,hH0LF,hG07LF,h07MF,gY03NF,gX03OF,gW01PF,gW0QF,gV0RF,gU07RF,gT03SF,gS01TF,gR01UF,gR0VF,gQ07VF,gP07WF,gO03XF,gN03YF,gN07YF,V03PF,V03OF8,V03NF8,V03MFC,V03LFE,V03KFE,V03KF,V03JF8,V03IF8,V03FFC,V03FE,V03E,V03,,::gK06N06,gK0FL0F0F8,:S078P0FL0F0F8,S078P0FL0F07,S078W0F,:00FE1FC007F807FE01FC007F8F01FF00F0700FE003F8,07FF7FF01FFE07FF07FF00FF8F07FF80F0F03FF80FFE,0LF81IF07FF0IF81FF8F07FFC0F0F07FFC1IF,0LF81F3F07FE1F8FC1FF8F03C7E0F0F07C781F1F8,0F83F07CI0F87801E03C3E00FI01E0F0F078103C078,0F01E07CI0787801E01C3E00FI01E0F0F0FI03C038,0F01E03CI0787803C01E3C00FI01E0F0F078007803C,0F01E03C03E787803C01E3C00F0079E0F0F07F007803C,0F01E03C0IF87803IFE3C00F03FFE0F0F03FE07IFC,0F01E03C3IF87803IFE3C00F07FFE0F0F01FF87IFC,0F01E03C3E0F87803IFC3C00F0F83E0F0F007FC7IF8,0F01E03C780787803CI03C00F0F01E0F0FI07C78,0F01E03C780787803CI03C00F0E01E0F0FI03E38,0F01E03C780787801EI03C00F0F01E0F0FI01E3C,0F01E03C7C0787C01F0183C00F0F01E0F0F0703C3E03,0F01E03C3F7F83FF0IF83C00F0FEFE0F0F0IFC1IF,0F01E03C3IF03FF07FFC3C00F07FFE0F0F0IF80IF8,0F01E03C0FFE01FF03FF83C00F03FF80F0F03FF007FF,O03FI07C00FCN0FCM07C001F8,,:^FS^FX Fabriek logo^FO600,220^GFA,585,585,9,,R07IFC,R0JFC,::::::::::::::::::::::00CI0CI0C0JFC,00EI0EI0E0JFC,01EI0E001E0JFC,01E001E001F0JFC,01F003F801F0JFC,03F807F803F8JFC,07F807F807F8JFC,07FC07FC07FCJFC,0FFC07FC0FFEJFC,1FFE0FFC0FFEJFC,1IF1FFE1FFEJFC,1IF1IF1MFC,3IF3IF3MFC,3IF3QFC,3UFC,VFC,::::::::::::::::::::::::^FS^FX Attention logo^FO600,45^GFA,710,710,10,N01C,N03E,N07F,:N0FF8,:M01F7C,M01F3E,M03E3E,M03E1F,M07C1F,M07C0F8,M0F80F8,L01F807C,L01F007C,L03E003E,:L07C001F,:L0F8I0F8,L0F8080F8,K01F03E07C,K01F07F07C,K03E07F03E,:K07C07F01F,:K0F807F00F8,:J01F007F007C,:J03E007F003E,J03E007F003F,J07C007F001F,J07C007FI0F8,J0F8007FI0F8,I01F8007FI07C,I01FI07FI07C,I03FI03FI03E,I03EI03FI03E,I07EI03EI01F,I07CI03EI01F,I0F8I03EJ0F8,:001FJ03EJ07C,:003EJ03EJ03E,:007CJ03EJ01F,007CJ01CJ01F,00F8K08K0F8,00F8Q0F8,01FR07C,01FK01CK07C,03EK03EK03E,07EK03FK03E,07CK07FK01F,0FCK03EK01F,0F8K03EL0F8,1F8L08L0F8,1FT07C,3FT07C,3ET03E,7CT03F,7CT01F,FCT01F8,XF8,::7WF,3VFE,^FS';
            else
                error(['The location code' itemnr(2) '(in itemnumber ' itemnr ') is not known for ' currentcaseid_long '.']);
            end
        end
        if caseindex < UScases_counter
            code.part4 = '^XB^XZ';
        elseif caseindex == UScases_counter
            code.part4 = '^XZ';
        end
        fprintf(fid, code.part1);
        fprintf(fid, code.part2);
        fprintf(fid, code.part3);
        fprintf(fid, code.part4);
        fprintf(fid, '\n');
        fprintf(fid, '\n');
        clear currentcaseid
        clear code
    end
else
    fprintf(fid, 'There are no FDA labels to print.');
end

fprintf(fid, '}$');
fclose(fid);

end