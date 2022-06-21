function varargout = InterfaceTest3(varargin)
% InterfaceTest3 MATLAB code for InterfaceTest3.fig
%      InterfaceTest3, by itself, creates a new InterfaceTest3 or raises the existing
%      singleton*.
%
%      H = InterfaceTest3 returns the handle to a new InterfaceTest3 or the handle to
%      the existing singleton*.
%H 
%      InterfaceTest3('CALLBACK',hObject,eventData,handles,...) calls the local
%      function named CALLBACK in InterfaceTest3.M with the given input arguments.
%
%      InterfaceTest3('Property','Value',...) creates a new InterfaceTest3 or raises the
%      existing singleton*.  Starting from the left, property value pairs are
%      applied to the GUI before InterfaceTest3_OpeningFcn gets called.  An
%      unrecognized property name or invalid value makes property application
%      stop.  All inputs are passed to InterfaceTest3_OpeningFcn via varargin.
%
%      *See GUI Options on GUIDE's Tools menu.  Choose "GUI allows only one
%      instance to run (singleton)".
%
% See also: GUIDE, GUIDATA, GUIHANDLES

% Edit the above text to modify the response to help InterfaceTest3

% Last Modified by GUIDE v2.5 13-Jul-2017 10:24:37

% Begin initialization code - DO NOT EDIT
gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @InterfaceTest3_OpeningFcn, ...
                   'gui_OutputFcn',  @InterfaceTest3_OutputFcn, ...
                   'gui_LayoutFcn',  [] , ...
                   'gui_Callback',   []);
if nargin && ischar(varargin{1})
    gui_State.gui_Callback = str2func(varargin{1});
end

if nargout
    [varargout{1:nargout}] = gui_mainfcn(gui_State, varargin{:});
else
    gui_mainfcn(gui_State, varargin{:});
end
% End initialization code - DO NOT EDIT


% --- Executes just before InterfaceTest3 is made visible.
function InterfaceTest3_OpeningFcn(hObject, eventdata, handles, varargin)
% This function has no output args, see OutputFcn.
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
% varargin   command line arguments to InterfaceTest3 (see VARARGIN)

% Choose default command line output for InterfaceTest3
handles.output = hObject;

% Update handles structure
guidata(hObject, handles);

% UIWAIT makes InterfaceTest3 wait for user response (see UIRESUME)
% uiwait(handles.figure1);


% --- Outputs from this function are returned to the command line.
function varargout = InterfaceTest3_OutputFcn(hObject, eventdata, handles) 
% varargout  cell array for returning output args (see VARARGOUT);
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Get default command line output from handles structure
varargout{1} = handles.output;

% --- Executes on button press in pushbutton1.
function pushbutton1_Callback(hObject, eventdata, handles)
% hObject    handle to pushbutton1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
disp('test');

value_box1 = get(handles.AutoLabels,'Value');
value_box2 = get(handles.AutoFinishingReports,'Value');
value_box3 = get(handles.ManualLabels,'Value');
value_box4 = get(handles.ManualFinishingReports,'Value');
value_edit4 = get(handles.edit4, 'string');

test = 1;
if test == 1
   disp('Will be printing this.. and that...'); 
   disp(value_box1); 
   disp(value_box2); 
   disp(value_box3); 
   disp(value_box4); 
   disp(value_edit4); 
end



% --- Executes on button press in AutoLabels.
function AutoLabels_Callback(hObject, eventdata, handles)
% hObject    handle to AutoLabels (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hint: get(hObject,'Value') returns toggle state of AutoLabels

% --- Executes on button press in AutoFinishingReports.
function AutoFinishingReports_Callback(hObject, eventdata, handles)
% hObject    handle to AutoFinishingReports (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hint: get(hObject,'Value') returns toggle state of AutoFinishingReports

% --- Executes on button press in ManualLabels.
function ManualLabels_Callback(hObject, eventdata, handles)
% hObject    handle to ManualLabels (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hint: get(hObject,'Value') returns toggle state of ManualLabels

% --- Executes on button press in ManualFinishingReports.
function ManualFinishingReports_Callback(hObject, eventdata, handles)
% hObject    handle to ManualFinishingReports (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hint: get(hObject,'Value') returns toggle state of ManualFinishingReports



function edit4_Callback(hObject, eventdata, handles)
% hObject    handle to edit4 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit4 as text
%        str2double(get(hObject,'String')) returns contents of edit4 as a double


% --- Executes during object creation, after setting all properties.
function edit4_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit4 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes during object creation, after setting all properties.
function figure1_CreateFcn(hObject, eventdata, handles)
% hObject    handle to figure1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called
