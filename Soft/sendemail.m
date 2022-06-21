
%receipients = {'tom.peeters@rsprint.be','lode.bosmans@rsprint.be','support@rsprint.be','info@rsprint.be'};
recipients = {'lode.bosmans@rsprint.be','bosmanslode@hotmail.com'};
subject = 'Cases in Phits production';
message = ['Dear customer,' 10 ...
            10 ...
            'In attachment you can find an overiew of your Phits cases in production' 10 ...
            10 ...
            'Best regards,' 10 ...
            'The Phits team'];
attachments = {'C:\Users\mathlab\Documents\MATLAB\Interface\Input\Templates\Grumpy.jpg'};

setpref('Internet','SMTP_Server','smtp.office365.com');
setpref('Internet','SMTP_Username','lode.bosmans@rsprint.be');
setpref('Internet','SMTP_Password','Kriekiekriek99w');
setpref('Internet','E_mail','production@rsprint.be');

sendmail(recipients,subject,message,attachments);


mail = 'lode.bosmans@rsprint.be';    % Replace with your email address
password = 'Kriekiekriek99w';          % Replace with your email password
server = 'smtp.office365.com';     % Replace with your SMTP server
props = java.lang.System.getProperties;
props.setProperty('mail.smtp.port','25');
%props.setProperty('mail.smtp.port','587');
% Apply prefs and props
setpref('Internet','E_mail',mail);
setpref('Internet','SMTP_Server',server);
setpref('Internet','SMTP_Username',mail);
setpref('Internet','SMTP_Password',password);
% Send the email
sendmail(recipients,subject,message,attachments);



%%

%https://groups.google.com/forum/#!msg/comp.soft-sys.matlab/AyOiqw1PcVI/Z2Cc1db-gOwJ

NET.addAssembly('System.Net')
import System.Net.Mail.*;
mySmtpClient = SmtpClient('smtp.office365.com');
mySmtpClient.UseDefaultCredentials = false;
mySmtpClient.Credentials = System.Net.NetworkCredential('lode.bosmans@rsprint.be', 'Kriekiekriek99w');
from = MailAddress('lode.bosmans@rsprint.be');
to = MailAddress('bosmanslode@hotmail.com');
myMail = MailMessage(from, to);
myMail.Subject = 'Test email matlab';
myMail.SubjectEncoding = System.Text.Encoding.UTF8;
myMail.Body = message;
myMail.BodyEncoding = System.Text.Encoding.UTF8;
myMail.IsBodyHtml = true;
mySmtpClient.Send(myMail)
%mySmtpClient.Credentials = New System.Net.NetworkCredential("email@domain.com", "password", "domain.com")

%%

% https://nl.mathworks.com/matlabcentral/answers/94446-can-i-send-e-mail-through-matlab-using-microsoft-outlook

to = {'lode.bosmans@rsprint.be','bosmanslode@hotmail.com'};
subject = 'Cases in Phits production';
body = ['Dear customer,' 10 ...
            10 ...
            'In attachment you can find an overiew of your Phits cases in production' 10 ...
            10 ...
            'Best regards,' 10 ...
            'The Phits team'];
attachments = {'C:\Users\mathlab\Documents\MATLAB\Interface\Input\Templates\Grumpy.jpg'};

%function sendolmail(to,subject,body,attachments)
%Sends email using MS Outlook. The format of the function is 
%Similar to the SENDMAIL command.
% Create object and set parameters.
h = actxserver('outlook.Application');
mail = h.CreateItem('olMail');
mail.Subject = subject;
mail.To = to;
mail.BodyFormat = 'olFormatHTML';
mail.HTMLBody = body;
% Add attachments, if specified.
if nargin == 4
    for i = 1:length(attachments)
        mail.attachments.Add(attachments{i});
    end
end
% Send message and release object.
mail.Send;
h.release;

%% 

% https://groups.google.com/forum/#!msg/comp.soft-sys.matlab/AyOiqw1PcVI/Z2Cc1db-gOwJ

recipients = {'lode.bosmans@rsprint.be','bosmanslode@hotmail.com'};
subject = 'Cases in Phits production';
message = ['Dear customer,' 10 ...
            10 ...
            'In attachment you can find an overiew of your Phits cases in production' 10 ...
            10 ...
            'Best regards,' 10 ...
            'The Phits team'];
attachments = {'C:\Users\mathlab\Documents\MATLAB\Interface\Input\Templates\Grumpy.jpg'};

mail='lode.bosmans@rsprint.be';
password='Kriekiekriek99w';

setpref('Internet','E_mail',mail);
setpref('Internet','SMTP_Server','smtp.office365.com');
setpref('Internet','SMTP_Username',mail);
setpref('Internet','SMTP_Password',password);
props=java.lang.System.getProperties;
pp=props.setProperty('mail.smtp.auth','true');
pp=props.setProperty('mail.smtp.socketFactory.class','javax.net.ssl.SSLSocketFactory');
%pp=props.setProperty('mail.smtp.socketFactory.port','465');
%pp=props.setProperty('mail.smtp.socketFactory.port','25');
pp=props.setProperty('mail.smtp.socketFactory.port','587');
sendmail(recipients,subject,message,attachments);


%%

dos('start chrome www.google.com')

%%


