function [mpc] = Gopt2Mat(case_name)
% This function transmits the GOPT file into the Matpower case file
%% define and load the basic parameters
Q_per = 0.2; % the percentage of the Q, compared with P
define_constants;
mpc = struct();
mpc.version = '2';
mpc.baseMVA = 100;

%% transform the bus data
[~, ~, bus_raw] = xlsread([case_name '.csv'],'节点');
bus_raw(1, :) = [];
num_bus = size(bus_raw, 1);
bus = ones(num_bus, VMIN);
% the bus name
bus(:, BUS_I) = 1:num_bus;
% the type of the bus is remained to be determined after the generator
% information

% the active and reactive load level
bus(:, PD) = cell2mat(bus_raw(:, 4));
bus(:, QD) = bus(:, PD) * Q_per;
% the voltage level
bus(:, BASE_KV) = cell2mat(bus_raw(:, 8));
% the values of Vmax and Vmin
bus(:, VMAX) = 1.1;
bus(:, VMIN) = 0.9;

%% transform the generator data
[~, ~, gen_raw] = xlsread([case_name '.csv'],'机组');
gen_raw(1, :) = [];
num_gen = size(gen_raw, 1);
gen = zeros(num_gen, APF);
% the bus that the generator connected to
for i = 1:num_gen
    gen(i, GEN_BUS) = ...
        find(strcmp(bus_raw(:,1), cellstr(gen_raw(i, 2))));
end
% the capacity of the generator
gen(:, PG) = cell2mat(gen_raw(:, 8));
gen(:, PMAX) = cell2mat(gen_raw(:, 8));
gen(:, PMIN) = cell2mat(gen_raw(:, 9));
% the Qmax and Qmin of the generator
gen(:, QMAX) = gen(:, PG) * Q_per;
gen(:, QMIN) = -gen(:, QMAX);
% the voltage magnitude set point
gen(:, VG) = ones(num_gen, 1);
% MVA base of generator
gen(:, MBASE) = ones(num_gen, 1) * mpc.baseMVA;
% status of generator
gen(:, GEN_STATUS) = ones(num_gen, 1);

%% define the types of buses
bus_gen = zeros(num_bus, 1);% the generation capacity of each bus
for i = 1:num_gen
    bus_gen(gen(i, GEN_BUS), 1) = ...
        bus_gen(gen(i, GEN_BUS), 1) + gen(i, PG);
end
bus(bus_gen>0.1, BUS_TYPE) = 2;% PV bus if connected with generators
[~, ref_id] = max(bus_gen);
bus(ref_id, BUS_TYPE) = 3;% ref bus if connected with the largest capacity
%% transform the branch data
[~, ~, branch_raw] = xlsread([case_name '.csv'],'线路');
branch_raw(1, :) = [];
num_branch = size(branch_raw, 1);
branch = zeros(num_branch, ANGMAX);
% the from bus and the to bus of the branch
for i = 1:num_branch
    branch(i, F_BUS) = ...
        find(strcmp(bus_raw(:, 1), cellstr(branch_raw(i, 4))));
    branch(i, T_BUS) = ...
        find(strcmp(bus_raw(:, 1), cellstr(branch_raw(i, 5))));
end
% resistance, reactance, and line charging susceptance
branch(:, BR_R) = cell2mat(branch_raw(:, 8));
branch(:, BR_X) = cell2mat(branch_raw(:, 7));
w = 50 * 2 * pi;
branch(:, BR_B) = cell2mat(branch_raw(:, 9)) * 2 * w;
% the rate A B C
branch(:, RATE_A) = cell2mat(branch_raw(:, 10));
branch(:, RATE_B) = cell2mat(branch_raw(:, 11));
branch(:, RATE_C) = cell2mat(branch_raw(:, 21));
% the status of branch
branch(:, BR_STATUS) = ones(num_branch, 1);
% angle min and angle max
branch(:, ANGMAX) = 360 * ones(num_branch, 1);
branch(:, ANGMIN) = -branch(:, ANGMAX);

%% formulate mpc
mpc.bus = bus;
mpc.gen = gen;
mpc.branch = branch;
end