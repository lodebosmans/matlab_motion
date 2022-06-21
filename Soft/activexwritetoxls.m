function activexwritetoxls(range,data,Activesheet)

ActivesheetRange = get(Activesheet,'Range',range);
set(ActivesheetRange, 'Value', data);

end