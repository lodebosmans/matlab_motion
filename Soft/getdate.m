function [y,mo,d,h,mi,s] = getdate()

date2day = date;
y = date2day(10:11);
mo = date2day(4:6);
if mo == 'Jan'
    mo = '01';
elseif mo == 'Feb'
    mo = '02';
elseif mo == 'Mar'
    mo = '03';
elseif mo == 'Apr'
    mo = '04';
elseif mo == 'May'
    mo = '05';
elseif mo == 'Jun'
    mo = '06';
elseif mo == 'Jul'
    mo = '07';
elseif mo == 'Aug'
    mo = '08';
elseif mo == 'Sep'
    mo = '09';
elseif mo == 'Oct'
    mo = '10';
elseif mo == 'Nov'
    mo = '11';
elseif mo == 'Dec'
    mo = '12';
end
d = date2day(1:2);

t = datetime('now','Format','HH:mm:ss');
[h,mi,s] = hms(t);
s = ceil(s);

if h < 10
    h = ['0' num2str(h)];
elseif h >= 10
    h = num2str(h);
end
if mi < 10
    mi = ['0' num2str(mi)];
elseif  mi >= 10
    mi =num2str(mi);
end
if s < 10
    s = ['0' num2str(s)];
elseif s >= 10
    s = num2str(s);
end

end