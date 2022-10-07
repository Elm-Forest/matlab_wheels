clc;
clear;

%===-------------参数设置--------------===%
%因变量数据地址，应当是一个excel文件
X_address='C:\Users\55378\Desktop\X.xlsx';

%选择X的sheet索引编号，应当是目标sheet在excel中的序号
X_sheet_index=1;

%自变量数据地址
Y_address='C:\Users\55378\Desktop\data\X_C5归一化.xlsx';

%选择Y的sheet的索引编号
Y_sheet_index=1;

%Pearson系数矩阵输出地址，此处为默认输出，可以更改
pearson_output="C:\Users\55378\Desktop\X"+X_sheet_index+"_Y"+Y_sheet_index+"_pearson3";
%===-------------参数设置---------------===%

%获取pearson分析矩阵
coeff=get_peasson_byIndex(X_address,X_sheet_index,Y_address,Y_sheet_index);
%保存结果
save_output(pearson_output,coeff);
%得到最大关系
get_max_pearson(coeff);
%plot_coeff(abs(coeff));

X=xlsread(X_address);
Y=xlsread(Y_address);

%bar(minmax(X(:,1)),[1:length(X(:,1))]);
%bar_coeff(coeff);
%plot_coeff(coeff);
%stem_coeff(coeff);
draw_coffplot(Y);


function[]=draw_coffplot(coeff)
    figure
    index_name={'Y1','Y2','Y3','Y4','Y5','X1','X2','X3','X4'};
    corrplot(coeff);
    disp("drawing~");
end

%coeff的条形图
function[]=stem_coeff(coeff)
    figure(3);
    for i=1:length(coeff(1,:))
        stem([1:length(coeff(:,i))],coeff(:,i))
        hold on;
    end
    hold off;
end

%coeff的条形图
function[]=bar_coeff(coeff)
    figure(1);
    for i=1:length(coeff(1,:))     
        bar([1:length(coeff(:,i))],coeff(:,i))
        hold on;
    end
    hold off;
end

%针头图
function[]=draw_stem(X,Y)
    stem(minmax(X(:,1)),[1:length(X(:,1))])
    hold on;
    stem(minmax(Y(:,1)),[1:length(Y(:,1))])
end

function[]=draw_scatter(X,Y)
    scatter(minmax(X(:,1)),[1:length(X(:,1))],'*');
    hold on;
    scatter(minmax(Y(:,1)),[1:length(Y(:,1))],'.');
end

%归一化
function[min_max]=minmax(arr)
    minn=min(arr(:,1));
    maxx=max(arr(:,1));
    min_max=(arr(:,1)-minn)/(maxx-minn);
end

%coeff折线图
function[]=plot_coeff(coeff)
    figure(2);
    coeff=abs(coeff);
    for i=1:length(coeff(1,:))
        plot(coeff(:,i))
        grid on;
        hold on;
    end
end

%得到最大关系的索引
function[]=get_max_pearson(coeff)
    [val,index]=max(abs(coeff),[],2);
    disp("index="+index+",Max Abs pearson="+val)
end

%保存结果
function[]=save_output(output_address,coeff)
    xlswrite(output_address,coeff);
    disp("结果成功保存于："+output_address+'.xls');
end

%通过sheet下标index来获取对应的两表以及各变量之间的pearson系数
function[coeff]=get_peasson_byIndex(file_X,X_index,file_Y,Y_index)
    [~,sheetX,~]=xlsfinfo(file_X);
    [~,sheetY,~]=xlsfinfo(file_Y);
    X=xlsread(file_X,""+sheetX(X_index));
    Y=xlsread(file_Y,""+sheetY(Y_index));
    coeff=auto_pearson(X,Y);
end

%自动获取X,Y之间的pearson系数
%X是因变量矩阵，Y是自变量矩阵，每列为一组变量
%X与Y的维度应当小于等于2维，即二维matrix
function[coeff]=auto_pearson(X,Y)
    X_col_size=length(X(1,:));
    Y_col_size=length(Y(1,:));
    coeff=zeros(X_col_size,Y_col_size);
    for i=1:X_col_size
        X_row_size=length(find(~isnan(X(:,i))));
        for j=1:Y_col_size
            Y_row_size=length(find(~isnan(Y(:,j))));
            row_min=min(X_row_size,Y_row_size);
            coeff(i,j)= corr(X(1:row_min,i) , Y(1:row_min,j));
            disp("X的第"+i+'列与Y的第'+j+'列之间的Pearson相关系数为：'+coeff(i,j));
        end
    end
end
