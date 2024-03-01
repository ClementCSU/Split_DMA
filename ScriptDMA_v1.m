% This program reads an xlsx file and writes in a separate folder the chunk of
% data based on the first column values.
% Author:   Clement Brousse
%           clbrous colostate.edu
% Last modification: 02/29/2024

%% Clear the workspace

clear all
close all
clc

%% Get all the variables

% Get the path to the input file 
[file,path,indx] = uigetfile('*.xlsx','Please select the input file','MultiSelect', 'off');
Path2File = fullfile(path, file);
if ~isfile(Path2File)
    error("The input file is incorrect.")
end
% Get the path to the folder where the files are saved.
OutputFolder = uigetdir('', 'Please select the output folder.')
if OutputFolder == 0
    error("The output folder specified is incorrect.")
end
% Ask # of row and the addition of the header
NumberRow2Keep =  input("How many row should be kept per file? ")
if ~isa(NumberRow2Keep,'double') & NumberRow2Keep>0
    error("The answer should be a boolan. 0 or 1.")
end
AddHeader =  input("Does the header need to be added in each files? 1 for yes or 0 for no");
AddHeader = logical(AddHeader);
if ~isa(AddHeader,'logical')
    error("The answer should be a boolan. 0 or 1.")
end
%% Read the data - Find end of the file

opts = spreadsheetImportOptions("NumVariables", 8);
% Specify sheet and range
opts.Sheet = "ID 12-29-2023 (100) (001) multi";
opts.DataRange = "A34:H717";
% Specify column names and types
opts.VariableNames = ["Module", "DMA", "VarName3", "VarName4", "VarName5", "VarName6", "VarName7", "VarName8"];
opts.VariableTypes = ["double", "double", "double", "double", "double", "double", "double", "double"];
% Import the data
Data = readtable(Path2File, opts, "UseExcel", false)
clear opts

if mod(abs(size(Data, 1)/NumberRow2Keep),1)~=0
    error("Number of row and total length of the document inconsistant")
end
%% Read the header

opts = spreadsheetImportOptions("NumVariables", 8);
% Specify sheet and range
opts.Sheet = "ID 12-29-2023 (100) (001) multi";
opts.DataRange = "A32:H33";
% Specify column names and types
opts.VariableNames = ["Module", "DMA", "VarName3", "VarName4", "VarName5", "VarName6", "VarName7", "VarName8"];
opts.VariableTypes = ["string", "string", "string", "string", "string", "string", "string", "string"];
% Import the data
% Import the data
Header = readtable(Path2File, opts, "UseExcel", false)
clear opts

%% Format and export the data

for ii = 1:size(Data, 1)/NumberRow2Keep
    FileName = [file(1:end-5), ' - ', num2str(ii), '.csv']; 
    
    data = Data(1+(ii-1)*NumberRow2Keep:ii*NumberRow2Keep, :);
    data = sortrows(data, 1);
    if AddHeader == 1
        data = [Header;data];
    end

    writetable(data,FileName)
end