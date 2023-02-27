%%%%%%%%%%Start%%%%%%%%%%
clear;clc;

%%%%%%%%%%READ%%%%%%%%%%
[file,path] = uigetfile('*.xlsx');
file_path=[path file];

%%%%%%%%%%%输出文件命名%%%%%%%%%%%%%%%
file_out(:,length(file)+4) = 'x';
file_out(:,length(file)+3) = 's';
file_out(:,length(file)+2) = 'l';
file_out(:,length(file)+1) = 'x';
file_out(:,length(file)) = '.';
file_out(:,length(file)-1) = 't';
file_out(:,length(file)-2) = 'u';
file_out(:,length(file)-3) = 'o';
file_out(:,length(file)-4) = '_';

for j =1:length(file)-5
    file_out(:,j)=file(:,j);end

%%%%%%%%%%%清除重名文件%%%%%%%%%%%
if exist(file_out,"file")
    delete(file_out);
end

%%%%%%%%%%%%预处理%%%%%%%%%%%%
data = readmatrix(file_path);
data1 = [data(:,2),data(:,6),data(:,8),data(:,10),data(:,12),data(:,14),data(:,16),data(:,18),data(:,20),data(:,22),data(:,24),data(:,26),data(:,28),data(:,30),data(:,32),data(:,34),data(:,36),data(:,38),data(:,40),data(:,42),data(:,44),data(:,46),data(:,48),data(:,50),data(:,52),data(:,54),data(:,56),data(:,58),data(:,60),data(:,62),data(:,64),data(:,66),data(:,68),data(:,70),data(:,72),data(:,74),data(:,76),data(:,78),data(:,80),data(:,82),data(:,84),data(:,86),data(:,88),data(:,90),data(:,92),data(:,94),data(:,96),data(:,98),data(:,100),data(:,102),data(:,104),data(:,106),data(:,108),data(:,110),data(:,112),data(:,114),data(:,116),data(:,118),data(:,120),data(:,122),data(:,124),data(:,126),data(:,128),data(:,130)];

%%%%%%%%%%处理相位差（平移函数）%%%%%%%%%%
data1=circshift(data1,[8 0]);

%%%%%%%%%%处理幅值、次数%%%%%%%%%%
Maximum_PD = max(max(data1));
N_PD = sum(data1(:)>10);

%%%%%%%%%%处理相位%%%%%%%%%%
Phase_Position_Positive = 0;
Phase_Position_Negative = 0;

for x = 1:64
    for y = 1:64
        if data1(x,y)>10
            Phase_Position_Positive = [Phase_Position_Positive x];
        else
        end
    end
end
Phase_Position_Positive(:,1) = [];

for x = 65:124
    for y = 1:64
        if data1(x,y)>10
            Phase_Position_Negative = [Phase_Position_Negative x];
        else
        end
    end
end
Phase_Position_Negative(:,1) = [];

%%%%%%%%%%S、K特征量%%%%%%%%%%
S_Positive = skewness(Phase_Position_Positive);
S_Negative = skewness(Phase_Position_Negative);
K_Positive = kurtosis(Phase_Position_Positive)-3;
K_Negative = kurtosis(Phase_Position_Negative)-3;

%%%%%%%%%%PRPD_OUT%%%%%%%%%%
temp_write = [0 0];
for i = 1:124
    for k = 1:64
        if data1(i,k)>8
        temp_write = [temp_write ; 360*i/124 data1(i,k)];
        else
        end
    end
end
temp_write(1,:)=[]
writematrix(temp_write,file_out,'Sheet',1,'WriteMode','append');
%%%%%%%%%%特征量Out%%%%%%%%%%
Statistics = {  'Max_Amplitude',Maximum_PD;
    'Count_PD',N_PD;
    'Count_PD_per_second',N_PD/1.28;
    'S_Positive',S_Positive;
    'S_Negative',S_Negative;
    'K_Positive',K_Positive;
    'K_Negative',K_Negative};

writecell(Statistics,file_out,'Sheet',2);
