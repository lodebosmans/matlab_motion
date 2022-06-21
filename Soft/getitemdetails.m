function [answer] = getitemdetails(harmcodemessage,Shipment,itemnumber,camount,cdescription,cunitprice,charmcode)

if strcmp('TLharm',harmcodemessage) == 1
    harmcodemessage2 = 'Harmonized code (standard: top layer material - adjust if necessary)';
elseif strcmp('noTLharm',harmcodemessage) == 1
    harmcodemessage2 = 'Harmonized code injected from input (should be correct)';
end

goon = 0;
while goon == 0
    prompt = {'Itemnumber','Amount','Unit price',harmcodemessage2,'Description'};
    title = ['Provide input for the content of the shipment to ' char(Shipment.ShipTo.CompanyOrName) ' before proceding'];
    dims = [1 150];
    definput = {itemnumber,camount,cunitprice,charmcode,cdescription};
    answer = inputdlg(prompt,title,dims,definput);
    nrrequests = size(prompt,2);
    goon = 1;
    for crq = 1:nrrequests
        if size(char(answer(crq,1)),2) == 0
            goon = 0;
        end
    end
end
end