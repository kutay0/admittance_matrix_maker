% Clear previous variables, close figures, and clear the command window.
clc
clear
close all

%% Choosing file
% Get the list of all txt files in the folder
files = dir('*.txt');
% Check if there are any txt files available, if there arent terminate thescript
if isempty(files)
    disp('No txt files found in the folder.');
    return
end
% Display the list of available files to the user
disp('Files in folder:');
for i = 1:length(files)
    if i == 3
        % Display the third file with the special message using fprintf
        fprintf('%d. %s *Unfortunately this file cannot work due to problems with bus numbering.\n', i, files(i).name);
    else
        % Display all other files normally
        fprintf('%d. %s\n', i, files(i).name); %s string %d decimal (d. integer)
    end
end
% Ask the user to select a file by entering the corresponding number
while true
    file_index = input('Enter the number of the file you want to use (or 0 to quit): ');  
    % Exit the program if the user chooses 0
    if isnumeric(file_index) && file_index == 0
        disp('Exiting program...');
        return;
    end
    % Exit the program if the user chooses 3
    if isnumeric(file_index) && file_index == 3
        disp('Exiting program...');
        return;
    end
    % Validate the user's choice
    if isnumeric(file_index) && file_index >= 1 && file_index <= length(files)
        % If the choice is valid, break out of the loop
        break;
    else
        disp('Invalid choice. Please select a valid file number.');
    end
end   

%% Creating the Admittance Matrix
% Select a variable for the selected filename
filename = files(file_index).name;
% Open the selected file
fid = fopen(filename, 'r');
% Loop to read the file
while true
    tline = fgetl(fid);
    % Check for end of file
    if ~ischar(tline)
        break;
    end
    % Extract the bus count when 'BUS DATA FOLLOWS' is found
    if contains(tline, 'BUS DATA FOLLOWS')
        bus_count = 0;  % Initialize bus_count
        while true
            tline = fgetl(fid);  % Read the next line after 'BUS DATA FOLLOWS'
            if contains(tline, '-999')  % Look for the end marker
                break;
            end
            % Increment bus_count for each bus entry (assuming one bus per line)
            bus_count = bus_count + 1;
        end
        break;  % Exit the loop after processing bus data
    end
end
% Construct an n-by-n matrix of zeros
Ybus_matrix = zeros(bus_count);

%% Filling the matrix
% Process the branch data to fill the admittance matrix
while true
    tline = fgetl(fid);
    % Check for end of file
    if ~ischar(tline)
        break;
    end
    % Look for the line with branch data and extract the branch count if the line contains 'BRANCH DATA FOLLOWS'
    if contains(tline, 'BRANCH DATA FOLLOWS')
        branch_count = sscanf(tline, 'BRANCH DATA FOLLOWS %d ITEMS');  % Directly extract the number
        % Start reading the branch data lines
        for n = 1:branch_count
            tline = fgetl(fid);
            % Read branch data
            branch_data = sscanf(tline, '%d %d %d %d %d %d %f %f %f');
            % Extract values: From Bus, To Bus, Resistance (R), Reactancen (X), Susceptance (B)
            from_bus = branch_data(1);
            to_bus = branch_data(2);
            R = branch_data(7);  % Resistance
            X = branch_data(8);  % Reactance
            B = branch_data(9);  % Line charging susceptance
            % Calculating the admittance
            Y = 1 / (R + 1j*X);
            % Put the admittances into the matrix
            % Non-diagonal elements
            Ybus_matrix(from_bus, to_bus) = round(Ybus_matrix(from_bus, to_bus) - Y,2); %-Yij
            Ybus_matrix(to_bus, from_bus) = round(Ybus_matrix(to_bus, from_bus) - Y,2); %-Yji
            % Diagonal elements
            Ybus_matrix(from_bus, from_bus) = round(Ybus_matrix(from_bus, from_bus) + Y + 1j*B,2); %Yii + jB
            Ybus_matrix(to_bus, to_bus) = round(Ybus_matrix(to_bus, to_bus) + Y + 1j*B,2); %Yjj + jB
        end
        break;
    end
end
% Close the file after reading
fclose(fid);
% Display the admittance matrix
disp("Admittance Matrix for '"+bus_count+"' bus system:"+newline);
disp(num2str(Ybus_matrix));

%% Writing the matrix to excel file
% Convert the matrix to cell
Ybus_cell = num2cell(Ybus_matrix);
excel_cell = cellfun(@num2str , Ybus_cell, 'UniformOutput', false);
% Assign the excel file to be written on
excel_filename = 'AdmittanceMatrix.xlsx';
% Export the cell to the excel file
writecell(excel_cell, excel_filename,"WriteMode","replacefile");
disp(newline+"Admittance Matrix has been exported to: "+ excel_filename);