clc;
clear;
%row_size=length(arr(:,1));
%col_size=length(arr(1,:));

minmax_all('C:\Users\55378\Desktop\data\C.xlsx','C:\Users\55378\Desktop\data\g归一化.xlsx',"col");

%Excel自动归一化
%file_address:文件地址
%file_out:输出地址
function[]=minmax_all(file_address,file_out,style)
    [sheet_len,Sheet]=get_sheet(file_address);
    for i=1:sheet_len
        try
            disp("starting at "+Sheet(i))
            x1=xlsread(file_address,""+Sheet(i));
            x1=minmax_file(x1,style);
            xlswrite(file_out,x1,""+Sheet(i));
        catch exception
            disp(exception)
        end
    end
end

%获取excel的sheet数量
%file_address:文件地址
function[num,Sheep]=get_sheet(file_address)
    [~, Sheep, ~]=xlsfinfo(file_address);
    num=length(Sheep);
end

%Sheet自动归一化
%input为excel表格输入
%style为归一化模式，col为按列归一化，row为按行归一化
function[result]=minmax_file(input,style)
    col_size=length(input(1,:));
    for i=1:col_size
        result(:,i)=minmax(input(:,i),style);
    end
end

%数据归一化
%参数说明
%arr为所选矩阵
%style为归一化模式，col为按列归一化，row为按行归一化
function[min_max]=minmax(arr,style)
    if style=="col"
        minn=min(arr(:,1));
        maxx=max(arr(:,1));
        min_max=(arr(:,1)-minn)/(maxx-minn);
    else 
    if style=="row"
        minn=min(arr(1,:));
        maxx=max(arr(1,:));
        min_max=(arr(1,:)-minn)/(maxx-minn); 
    end
    end
end
