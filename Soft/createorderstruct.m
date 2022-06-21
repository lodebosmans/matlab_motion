function [salesordercheck_v3] = createorderstruct(salesordercheck_v3,cusnr,cy)

% salesordercheck_v3.orders.(cusnr).(['y' num2str(cy)]).overview = cell(0);
% for mo = 1:12
%     salesordercheck_v3.orders.(cusnr).(['y' num2str(cy)]).(['m' num2str(mo)]) = cell(0);
% end
% salesordercheck_v3.orders.(cusnr).(['y' num2str(cy)]).Type_RegSemi = 0;
% salesordercheck_v3.orders.(cusnr).(['y' num2str(cy)]).Type_RegFull = 0;
% salesordercheck_v3.orders.(cusnr).(['y' num2str(cy)]).Type_SlimSemi = 0;
% salesordercheck_v3.orders.(cusnr).(['y' num2str(cy)]).Type_SlimFull = 0;
% salesordercheck_v3.orders.(cusnr).(['y' num2str(cy)]).Type_TL_RegSemi_None = 0;
% salesordercheck_v3.orders.(cusnr).(['y' num2str(cy)]).Type_TL_RegSemi_Eva_NotAs = 0;
% salesordercheck_v3.orders.(cusnr).(['y' num2str(cy)]).Type_TL_RegSemi_Eva_As = 0;
% salesordercheck_v3.orders.(cusnr).(['y' num2str(cy)]).Type_TL_Pusoft_NotAs = 0;
% salesordercheck_v3.orders.(cusnr).(['y' num2str(cy)]).Type_TL_PuSoft_As = 0;
% salesordercheck_v3.orders.(cusnr).(['y' num2str(cy)]).Type_TL_EvaCarbon_NotAs = 0;
% salesordercheck_v3.orders.(cusnr).(['y' num2str(cy)]).Type_TL_EvaCarbon_As = 0;

levels = {'','Peasant_'};
types = {'RegSemi','RegFull','SlimSemi','SlimFull'};
tls = {'None','Eva_NotAs','Eva_As','Pusoft_NotAs','PuSoft_As','EvaCarbon_NotAs','EvaCarbon_As'};
nrlevels = size(levels,2);
nrtypes = size(types,2);
nrtls = size(tls,2);

for l = 1:nrlevels
    clevel = char(levels(1,l));
    salesordercheck_v3.orders.(cusnr).(['y' num2str(cy)]).([clevel 'overview']) = cell(0);
    for mo = 1:12
        salesordercheck_v3.orders.(cusnr).(['y' num2str(cy)]).([clevel 'm' num2str(mo)]) = cell(0);
    end
    for t = 1:nrtypes
        ctype = char(types(1,t));
        salesordercheck_v3.orders.(cusnr).(['y' num2str(cy)]).([clevel 'Type_' ctype]) = 0;
    end
    for t = 1:nrtypes
        ctype = char(types(1,t));
        for tl = 1:nrtls
            ctl = char(tls(1,tl));
            salesordercheck_v3.orders.(cusnr).(['y' num2str(cy)]).([clevel 'Type_TL_' ctype '_' ctl ]) = 0;
        end
    end
end

end