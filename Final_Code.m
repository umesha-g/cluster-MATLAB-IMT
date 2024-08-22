
clear;
clc;

% Input for linkage selection --------------------------------------------
fprintf('Hello there , Please choose your Clustering method.\n\n')

fprintf('Your Options are\n\n')

fprintf('  1 - Single Linkage method\n')
fprintf('  2 - Complete Linkage method\n')
fprintf('  3 - Average Linkage method\n')
fprintf('  4 - Ward’s Method\n')
fprintf('  5 - Centroid method\n')
fprintf('  6 - Median method\n\n')

Sel = input('Please Enter your choice  -  ');
%-------------------------------------------------------------------------

% Check for input validity
% -----------------------------------------------
 
while isempty(Sel) == 1 ||  (Sel <1 || Sel >6 || (floor(Sel) ~= ceil(Sel)))
    clear Sel
    Sel = input('\n\nPlease enter a valid number  -  ');
end


%-------------------------------------------------------------------------

%Check selection and start -----------------------------------------------
if Sel == 1
    fprintf('\n\nYou Have chosen Single Linkage Method\n');
    input('\n\nIs your Input File Ready ? If YES Press enter\n\n','s');
    [A,C,D,m,n] = Input();
    SingleLinkage(A,D,C,n,m);

elseif Sel == 2
    fprintf('\n\nYou Have chosen Complete Linkage Method\n');
    input('\n\nIs your Input File Ready ? If YES Press enter\n\n','s');
    [A,C,D,m,n] = Input();
     CompleteLinkage(A,D,C,n,m);

elseif Sel == 3
    fprintf('\n\nYou Have chosen Average Linkage Method\n');
    input('\n\nIs your Input File Ready ? If YES Press enter\n\n','s');
    [A,C,D,m,n] = Input();
     AverageLinkage(A,D,C,n,m);

elseif Sel == 4
    fprintf('\n\nYou Have chosen Ward s Method\n');
    input('\n\nIs your Input File Ready ? If YES Press enter\n\n','s');
    [A,C,D,m,n] = Input();
     Wards(A,D,C,n,m);

elseif Sel == 5
    fprintf('\n\nYou Have chosen Centroid Method\n');
    input('\n\nIs your Input File Ready ? If YES Press enter\n\n','s');
    [A,C,D,m,n] = Input();
     centroid(A,D,C,n,m);

elseif Sel == 6
    fprintf('\n\nYou Have chosen Median Method\n');
    input('\n\nIs your Input File Ready ? If YES Press enter\n\n','s');
    [A,C,D,m,n] = Input();
     median(A,D,C,n,m);
  
end


Res=input('Do you want to Restart the Programme ? If Yes Input Y. Any other Input will consider as No\n\n','s');
if Res=='Y'
    run('Final.m');
else
    fprintf('                                                <strong>------------PROGRAMME ENDED | THANK YOU-----------</strong>\n\n\n')
end
%------------------------------------------------------------------------

% Input taking Function--------------------------------------------------
function [A,C,D,m,n] = Input()

Inp_Tbl = readtable('input.xlsx','ReadRowNames',true);
row_num = (height(Inp_Tbl)+1);
H= strcat('A1:A',num2str(row_num));
Row_names = readtable('input.xlsx','Range',H);
NTbl2mat=table2array(Inp_Tbl);
CTbl2mat=table2array(Row_names);

    n=(row_num-1);
    
    m=n;
    
    C=transpose(CTbl2mat);
    
    A = NTbl2mat;

    B= pdist(A);

    D =squareform(B);

end
%------------------------------------------------------------------------  

% Single Linkage Function-------------------------------------------------
function SingleLinkage(A,D,C,n,m)

    fprintf('                                                <strong>Single Linkage method(Nearest Neighbor)</strong>\n\n\n')
    CC=C;
    DT=array2table(D,'RowNames',C);
    disp(DT)

    % generating file name -----------------------------------------------
    Datetime = datestr(now,'mmmm_dd_yyyy_HH_MM_SS_PM');
    DateName1= strcat('output_Table(Single_Linkage)_',Datetime,'.xlsx');
    DateName2= strcat('output_Dendrogram(Single_Linkage)_',Datetime,'.png');
    %---------------------------------------------------------------------

    %File Creating -------------------------------------------------------
    writetable(DT,DateName1,'Sheet',1,'WriteRowNames',true,'WriteMode','overwritesheet','AutoFitWidth',true,'PreserveFormat',true,'WriteVariableNames',false);
    %---------------------------------------------------------------------

    %Loop for cluster process --------------------------------------------
    for k=1:m-2
        [numRows,numCols] = size(D);
        U=zeros(numRows,numCols);
        V=zeros(numRows,numCols);

        % Finding minimum number -----------------------------------------
        min=100000000000000000000000000000;
        for i=1:n
            for j=1:n
                if D(i,j) >0
                    if D(i,j) < min
                        min=D(i,j);
                        min_i = i;
                        min_j= j;
                    end
                end
            end
        end
        %-----------------------------------------------------------------
       
        % loop For Horizontal minimum chekup -----------------------------
        for i=min_j
            for j=1:numCols
                if D(i,j) >0
                    if D(i,j) < D(min_i,j)
                        V(i,j)= D(i,j);
                    else
                       V(i,j)= D(min_i,j);
                    end
                end
            end
        end
        [numRows,numCols] = size(D);
        %-----------------------------------------------------------------
    
        % loop For VERTICAL minimum chekup -------------------------------
        for j=min_j
            for i=1:numRows
                if D(i,j) >0
                    if D(i,j) < D(i,min_i)
                       U(i,j)= D(i,j);
                    else
                       U(i,j)= D(i,min_i);
                    end
                end
            end
        end
        %-----------------------------------------------------------------
    
        %reshaping distance matrix ---------------------------------------
        for i= min_j
            for j=1:numCols
                D(i,j) = V(i,j);
            end
        end
        
        for j= min_j
            for i=1:numRows
                D(i,j) = U(i,j);
            end
        end
        
        D(min_i,:)=[];
        D(:,min_i)=[];
        %-----------------------------------------------------------------
        
        % Auto clustre name generator-------------------------------------
        [numRowsCity,numColsCity] = size(C);
        CN = num2str(abs(m-n+1));
        CNN = ['C' CN];
        %-----------------------------------------------------------------
    
        % adding genereted name to city matrix----------------------------
        C2 = [C(1 : min_j-1),CNN,C(min_j+1:numColsCity)];
        C2(:,min_i)=[];
        C=C2;
        %-----------------------------------------------------------------
    
        %Creating display table-------------------------------------------
        [numRows,numCols] = size(D);
        DT=array2table(D,'RowNames',C2);
        disp(DT)

        sheet = k+1;
        writetable(DT,DateName1,'Sheet',sheet,'WriteRowNames',true,'WriteMode','overwritesheet','AutoFitWidth',true,'PreserveFormat',true,'WriteVariableNames',false);
        %-----------------------------------------------------------------
        
        n= n-1;
    end
    %---------------------------------------------------------------------

    %Creating dedrogram---------------------------------------------------
    BB= pdist(A);
    tree = linkage(A,'single');
    leafOrder = optimalleaforder(tree,BB);
    %create cell of labels 
    labels = cellstr(CC);
    %plot dendogram with custom labels
    dendrogram(tree, 0, 'Labels', labels, 'orientation', 'left')
    saveas(gcf,DateName2)
    %---------------------------------------------------------------------
    
end
%-------------------------------------------------------------------------

% complete Linkage Function-----------------------------------------------
function CompleteLinkage(A,D,C,n,m)

    fprintf('                                                <strong>Complete Linkage method(Farthest Neighbor)</strong>\n\n\n')

    CC=C;
    DT=array2table(D,'RowNames',C);
    disp(DT)

    % generating file name -----------------------------------------------
    Datetime = datestr(now,'mmmm_dd_yyyy_HH_MM_SS_PM');
    DateName1= strcat('output_Table(complete_Linkage)_',Datetime,'.xlsx');
    DateName2= strcat('output_Dendrogram(complete_Linkage)_',Datetime,'.png');
    %---------------------------------------------------------------------

    %File Creating -------------------------------------------------------
    writetable(DT,DateName1,'Sheet',1,'WriteRowNames',true,'WriteMode','overwritesheet','AutoFitWidth',true,'PreserveFormat',true,'WriteVariableNames',false);
    %---------------------------------------------------------------------

    

    %Loop for cluster process --------------------------------------------
    for k=1:m-2
        [numRows,numCols] = size(D);
        U=zeros(numRows,numCols);
        V=zeros(numRows,numCols);

        % Finding maximum number -----------------------------------------
        min=1000000000000000000000;
        for i=1:n
            for j=1:n
                if D(i,j) >0
                    if D(i,j) < min
                        min=D(i,j);
                        max_i = i;
                        max_j= j;
                    end
                end
            end
        end
        %-----------------------------------------------------------------
       
        % For Horizontal maximum chekup ----------------------------------
        for i=max_j
            for j=1:numCols
                if D(i,j) >0
                    if D(i,j) > D(max_i,j)
                        V(i,j)= D(i,j);
                    else
                       V(i,j)= D(max_i,j);
                    end
                end
            end
        end
        [numRows,numCols] = size(D);
        %---------------------------------------------------------------------
    
        % For VERTICAL maximum chekup ----------------------------------------
        for j=max_j
            for i=1:numRows
                if D(i,j) >0
                    if D(i,j) > D(i,max_i)
                       U(i,j)= D(i,j);
                    else
                       U(i,j)= D(i,max_i);
                    end
                end
            end
        end
        %---------------------------------------------------------------------
    
        %reshaping distance matrix -------------------------------------------
        for i= max_j
            for j=1:numCols
                D(i,j) = V(i,j);
            end
        end
        
        for j= max_j
            for i=1:numRows
                D(i,j) = U(i,j);
            end
        end
        
        D(max_i,:)=[];
        D(:,max_i)=[];
        %---------------------------------------------------------------------
        
        % Auto clustre name generator-----------------------------------------
        [numRowsCity,numColsCity] = size(C);
        CN = num2str(abs(m-n+1));
        CNN = ['C' CN];
        %---------------------------------------------------------------------
    
        % adding genereted name to city matrix--------------------------------
        C2 = [C(1 : max_j-1),CNN,C(max_j+1:numColsCity)];
        C2(:,max_i)=[];
        C=C2;
        %---------------------------------------------------------------------
    
        %Creating display table-----------------------------------------------
        [numRows,numCols] = size(D);
        DT=array2table(D,'RowNames',C2);
        disp(DT)
        sheet = k+1;
        
        writetable(DT,DateName1,'Sheet',sheet,'WriteRowNames',true,'WriteMode','overwritesheet','AutoFitWidth',true,'PreserveFormat',true,'WriteVariableNames',false);
        %---------------------------------------------------------------------
    
        n= n-1;
    end
    %---------------------------------------------------------------------

    %Creating dedrogram---------------------------------------------------
    BB= pdist(A);
    tree = linkage(A,'complete');
    leafOrder = optimalleaforder(tree,BB);
    %create cell of labels 
    labels = cellstr(CC);
    %plot dendogram with custom labels
    dendrogram(tree, 0, 'Labels', labels, 'orientation', 'left')
    saveas(gcf,DateName2)
    %---------------------------------------------------------------------
end
%-------------------------------------------------------------------------

% Average Linkage Function------------------------------------------------
function AverageLinkage(A,D,C,n,m)
    
    na= n;
    nb= n;
    x= A;
    y= A;

    fprintf('                                                <strong>Average Linkage method</strong>\n\n\n')
    
    CC=C;
    D =pdist(A);
    D2=squareform(D);

    DT=array2table(D2,'RowNames',CC); 

    % generating file name ----------------------------------------------
    Datetime = datestr(now,'mmmm_dd_yyyy_HH_MM_SS_PM');
    DateName1= strcat('output_Table(Average_Linkage)_',Datetime,'.xlsx');
    DateName2= strcat('output_Dendrogram(Average_Linkage)_',Datetime,'.png');
    %--------------------------------------------------------------------

    %File Creating -------------------------------------------------------
    writetable(DT,DateName1,'Sheet',1,'WriteRowNames',true,'WriteMode','overwritesheet','AutoFitWidth',true,'PreserveFormat',true,'WriteVariableNames',false);
    %---------------------------------------------------------------------

    %Creating Dendrigram--------------------------------------------------
    disp(DT)
    %generating tree
    tree = linkage(D2,'average');
    %create cell of labels 
    labels = cellstr(CC);
    %plot dendogram with custom labels
    dendrogram(tree, 0, 'Labels', labels, 'orientation', 'left')
    saveas(gcf,DateName2)
    %---------------------------------------------------------------------
end
%-------------------------------------------------------------------------

% wards Function---------------------------------------------------------
function Wards(A,D,C,n,m)

    na=n;
    nb=n;
    x=A;
    y=A;
    
    CC=C;
    DT=array2table(D,'RowNames',C);
    disp(DT)
    
    % generating file name ----------------------------------------------
    Datetime = datestr(now,'mmmm_dd_yyyy_HH_MM_SS_PM');
    DateName1= strcat('output_Table(Wards)_',Datetime,'.xlsx');
    DateName2= strcat('output_Dendrogram(Wards)_',Datetime,'.png');
    %-------------------------------------------------------------------- 

    %creating file-------------------------------------------------------
    writetable(DT,DateName1,'Sheet',1,'WriteRowNames',true,'WriteMode','overwritesheet','AutoFitWidth',true,'PreserveFormat',true,'WriteVariableNames',false);
    %--------------------------------------------------------------------

    %Creating dedrogram---------------------------------------------------
    BB= pdist(A);
    tree = linkage(A,'ward');
    leafOrder = optimalleaforder(tree,BB);
    %create cell of labels 
    labels = cellstr(CC);
    %plot dendogram with custom labels
    dendrogram(tree, 0, 'Labels', labels, 'orientation', 'left')
    saveas(gcf,DateName2)
    %---------------------------------------------------------------------
    
end
%-------------------------------------------------------------------------

% centroid Function---------------------------------------------------------
function centroid(A,D,C,n,m)

    na=n;
    nb=n;
    x=A;
    y=A;
    
    CC=C;
    DT=array2table(D,'RowNames',C);
    disp(DT)
    
    % generating file name ----------------------------------------------
    Datetime = datestr(now,'mmmm_dd_yyyy_HH_MM_SS_PM');
    DateName1= strcat('output_Table(Centroid)_',Datetime,'.xlsx');
    DateName2= strcat('output_Dendrogram(Centroid)_',Datetime,'.png');
    %-------------------------------------------------------------------- 

    %creating file-------------------------------------------------------
    writetable(DT,DateName1,'Sheet',1,'WriteRowNames',true,'WriteMode','overwritesheet','AutoFitWidth',true,'PreserveFormat',true,'WriteVariableNames',false);
    %--------------------------------------------------------------------

    %Creating dedrogram---------------------------------------------------
    BB= pdist(A);
    tree = linkage(A,'centroid');
    leafOrder = optimalleaforder(tree,BB);
    %create cell of labels 
    labels = cellstr(CC);
    %plot dendogram with custom labels
    dendrogram(tree, 0, 'Labels', labels, 'orientation', 'left')
    saveas(gcf,DateName2)
    %---------------------------------------------------------------------
end
%-------------------------------------------------------------------------

% median Function---------------------------------------------------------
function median(A,D,C,n,m)

    na=n;
    nb=n;
    x=A;
    y=A;
    
    CC=C;
    DT=array2table(D,'RowNames',C);
    disp(DT)
    
    % generating file name ----------------------------------------------
    Datetime = datestr(now,'mmmm_dd_yyyy_HH_MM_SS_PM');
    DateName1= strcat('output_Table(Median)_',Datetime,'.xlsx');
    DateName2= strcat('output_Dendrogram(Median)_',Datetime,'.png');
    %-------------------------------------------------------------------- 

    %creating file-------------------------------------------------------
    writetable(DT,DateName1,'Sheet',1,'WriteRowNames',true,'WriteMode','overwritesheet','AutoFitWidth',true,'PreserveFormat',true,'WriteVariableNames',false);
    %--------------------------------------------------------------------

    %Creating dedrogram---------------------------------------------------
    BB= pdist(A);
    tree = linkage(A,'median');
    leafOrder = optimalleaforder(tree,BB);
    %create cell of labels 
    labels = cellstr(CC);
    %plot dendogram with custom labels
    dendrogram(tree, 0, 'Labels', labels, 'orientation', 'left')
    saveas(gcf,DateName2)
    %---------------------------------------------------------------------
end
%------------------------------------------------------------------------- 