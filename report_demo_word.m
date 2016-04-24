function [] = report_demo_word(varargin)
    if isempty(varargin)
        return;
    end
    format short
    infilename =  sprintf('%s', varargin{1});
    
    Word = actxserver('Word.application'); 
    Word.Visible = 0; 
    set(Word,'DisplayAlerts',0); 
    Docs = Word.Documents; 
    Doc = Docs.Open(infilename); 

    global row_index
    row_index = 5;

    Quotation_No = GetNextNumberInput(Doc);
    Customer = GetNextInput(Doc);
    Customer_project_no = GetNextInput(Doc);
    Shipyard = GetNextInput(Doc);
    Yard_No = GetNextInput(Doc);
    Classification_Society = GetNextInput(Doc);
    DP_Class_Selected = GetNextNumberInput(Doc);
    Additional_Class_Battery_Notation = GetNextInput(Doc);
    Revision = GetNextNumberInput(Doc);
    Revision_Description = GetNextInput(Doc);
    Revision_Date = GetNextInput(Doc);
    Author = GetNextInput(Doc);

    row_index = row_index + 1;
    row_index = row_index + 1;

    Total_Power_Required = GetNextNumberInput(Doc);
    Total_Capacity_Required = GetNextNumberInput(Doc);
    Required_Battery_lifetime = GetNextNumberInput(Doc);

    row_index = row_index + 1;

    Description_Case_1 = GetNextInput(Doc);
    Input_Required_Power_in_kW_Case_1 = GetNextNumberInput(Doc);
    Input_Time_Case_1 = GetNextNumberInput(Doc);
    Description_Case_2 = GetNextInput(Doc);
    Input_Required_Power_in_kW_Case_2 = GetNextNumberInput(Doc);
    Input_Time_Case_2 = GetNextNumberInput(Doc);
    Description_Case_3 = GetNextInput(Doc);
    Input_Required_Power_in_kW_Case_3 = GetNextNumberInput(Doc);
    Input_Time_Case_3 = GetNextNumberInput(Doc);
    Description_Case_4 = GetNextInput(Doc);
    Input_Required_Power_in_kW_Case_4 = GetNextNumberInput(Doc);
    Input_Time_Case_4 = GetNextNumberInput(Doc);
    Description_Case_5 = GetNextInput(Doc);
    Input_Required_Power_in_kW_Case_5 = GetNextNumberInput(Doc);
    Input_Time_Case_5 = GetNextNumberInput(Doc);
    Description_Case_6 = GetNextInput(Doc);
    Input_Required_Power_in_kW_Case_6 = GetNextNumberInput(Doc);
    Input_Time_Case_6 = GetNextNumberInput(Doc);

    row_index = row_index + 1;
    row_index = row_index + 1;
    row_index = row_index + 1;
    row_index = row_index + 1;
    row_index = row_index + 1;

    Input_No_of_ESU_units = GetNextNumberInput(Doc);
    Voltage_Level_Switchboard = GetNextNumberInput(Doc);
    Frequency_Level_Switchboard = GetNextNumberInput(Doc);
    Voltage_Level_ESU = GetNextNumberInput(Doc);
    Input_Application_Name = GetNextInput(Doc);
    Converter_Control_Method = GetNextInput(Doc);
    Motor_Starting_Current = GetNextNumberInput(Doc);
    No_of_Shore_Connections = GetNextNumberInput(Doc);
    row_index = row_index + 1;
    Shore_Connection_1_Level_1_High = GetNextNumberInput(Doc);
    Shore_Connection_1_Level_1_Low = GetNextNumberInput(Doc);
    Shore_Connection_1_Level_2_High = GetNextNumberInput(Doc);
    Shore_Connection_1_Level_2_Low = GetNextNumberInput(Doc);
    Shore_Connection_1_Level_3_High = GetNextNumberInput(Doc);
    Shore_Connection_1_Level_3_Low = GetNextNumberInput(Doc);
    Shore_Connection_1_Current = GetNextNumberInput(Doc);
    row_index = row_index + 1;
    Shore_Connection_2_Level_1_High = GetNextNumberInput(Doc);
    Shore_Connection_2_Level_1_Low = GetNextNumberInput(Doc);
    Shore_Connection_2_Level_2_High = GetNextNumberInput(Doc);
    Shore_Connection_2_Level_2_Low = GetNextNumberInput(Doc);
    Shore_Connection_2_Level_3_High = GetNextNumberInput(Doc);
    Shore_Connection_2_Level_3_Low = GetNextNumberInput(Doc);
    Shore_Connection_2_Current = GetNextNumberInput(Doc);
    Hours_for_Engineering_ESU = GetNextNumberInput(Doc);
    Commissioning_Days = GetNextNumberInput(Doc);

    % Common Fixed Data
    Maximum_voltage = 1000;
    ESU_Dimensions_D = 800;
    ESU_Dimensions_H = 2040;
    Price_in_Euro_Per_Unit = 58611;
    Labour_Cost_Project_Management_Engineering_in_Norway = 1300;
    Labour_Cost_Service_Commissioning_Test = 1050;
    Cost_Commissioning_Travel = 25000;
    Cost_Hotel = 2500;
    Cost_Transport = 660;

    if strcmp(Input_Application_Name, 'High energy batteries')
        disp('High energy batteries');
        Nominal_Energy_in_kWh_High_Power_Batteries = 49.16;
        Nominal_Energy_in_kWh_High_Energy_Batteries = 1;
        Maximum_charge_A_1C_One_String_High_Energy_Batteries = 80;
        Maximum_charge_A_3_3C_One_String_High_Power_Batteries = 1;
        Minimum_voltage_1C_discharge_at_20_SOC = 830;
        Minimum_voltage_2C_discharge_at_20_SOC = 1;
        Maximum_discharge_A_2C_One_String_High_Energy_Batteries	= 160;
        Maximum_discharge_A_3_3C_One_String_High_Power_Batteries = 1;
        Continuous_rms_charge_discharge_A_T_25C = 40;

        Required_power_in_kWh = (Input_Required_Power_in_kW_Case_1*Input_Time_Case_1) + (Input_Required_Power_in_kW_Case_2*Input_Time_Case_2) + ...
            (Input_Required_Power_in_kW_Case_3*Input_Time_Case_3) + (Input_Required_Power_in_kW_Case_4*Input_Time_Case_4) + ...
            (Input_Required_Power_in_kW_Case_5*Input_Time_Case_5) + (Input_Required_Power_in_kW_Case_6*Input_Time_Case_6);
        if strcmp(Input_Application_Name, 'High power batteries')
            ESU_Application_Type = 'VL30PFe';
        end
        Energy_Start_of_Life_in_kWh	= Nominal_Energy_in_kWh_High_Power_Batteries * 0.8;
        Energy_End_of_Life_in_kWh = Energy_Start_of_Life_in_kWh * 0.8;
        Maximum_charge_kW_1C_One_String_High_Energy_Batteries = (Maximum_charge_A_1C_One_String_High_Energy_Batteries * Minimum_voltage_1C_discharge_at_20_SOC) / 1000;
        Maximum_discharge_kW_2C_One_String_High_Energy_Batteries = (Maximum_discharge_A_2C_One_String_High_Energy_Batteries * Minimum_voltage_1C_discharge_at_20_SOC)/1000;
        Continuous_rms_charge_discharge_kW_T_25C = (Continuous_rms_charge_discharge_A_T_25C * Minimum_voltage_2C_discharge_at_20_SOC)/1000;
        Total_no_of_battery_strings_for_each_ESU_unit = Required_power_in_kWh / Energy_End_of_Life_in_kWh;
        ESU_Total_Continuous_Maximum_charge_Amp	= Maximum_charge_A_1C_One_String_High_Energy_Batteries * Total_no_of_battery_strings_for_each_ESU_unit;
        ESU_Total_Continuous_Maximum_charge_kW = Maximum_charge_kW_1C_One_String_High_Energy_Batteries * Total_no_of_battery_strings_for_each_ESU_unit;
        ESU_Total_Continuous_Maximum_discharge_Amp = Maximum_discharge_A_2C_One_String_High_Energy_Batteries * Total_no_of_battery_strings_for_each_ESU_unit;
        ESU_Total_Continuous_Maximum_discharge_kW = Maximum_discharge_kW_2C_One_String_High_Energy_Batteries * Total_no_of_battery_strings_for_each_ESU_unit;
    end
        Required_power_in_kW = max([Input_Required_Power_in_kW_Case_1 Input_Required_Power_in_kW_Case_2 Input_Required_Power_in_kW_Case_3, ...
            Input_Required_Power_in_kW_Case_4 Input_Required_Power_in_kW_Case_5 Input_Required_Power_in_kW_Case_6, ...
            (Voltage_Level_Switchboard * 3^0.5 * Motor_Starting_Current)]);
    if strcmp(Input_Application_Name, 'High power batteries')
        disp('High power batteries');
        Nominal_Energy_in_kWh_High_Power_Batteries = 1;
        Nominal_Energy_in_kWh_High_Energy_Batteries = 68.47;
        Maximum_charge_A_1C_One_String_High_Energy_Batteries = 1;
        Maximum_charge_A_3_3C_One_String_High_Power_Batteries = 200;
        Minimum_voltage_1C_discharge_at_20_SOC = 1;
        Minimum_voltage_2C_discharge_at_20_SOC = 830;
        Maximum_discharge_A_2C_One_String_High_Energy_Batteries	= 1;
        Maximum_discharge_A_3_3C_One_String_High_Power_Batteries = 200;
        Continuous_rms_charge_discharge_A_T_25C = 120;

        if strcmp(Input_Application_Name, 'High energy batteries')
            ESU_Application_Type = 'VL41MFe';
        end
        Energy_Start_of_Life_in_kWh	= Nominal_Energy_in_kWh_High_Energy_Batteries * 0.8;
        Energy_End_of_Life_in_kWh = Energy_Start_of_Life_in_kWh * 0.8;
        Maximum_charge_kW_3_3C_One_String_High_Power_Batteries = (Maximum_charge_A_3_3C_One_String_High_Power_Batteries * Minimum_voltage_2C_discharge_at_20_SOC) / 1000;
        Maximum_discharge_kW_3_3C_One_String_High_Power_Batteries = (Maximum_discharge_A_3_3C_One_String_High_Power_Batteries * Minimum_voltage_2C_discharge_at_20_SOC)/1000;
        Continuous_rms_charge_discharge_kW_T_25C = (Continuous_rms_charge_discharge_A_T_25C * Minimum_voltage_1C_discharge_at_20_SOC)/1000;
        Total_no_of_battery_strings_for_each_ESU_unit = Required_power_in_kW / Maximum_charge_kW_3_3C_One_String_High_Power_Batteries;
        ESU_Total_Continuous_Maximum_charge_Amp	= Maximum_charge_A_3_3C_One_String_High_Power_Batteries * Total_no_of_battery_strings_for_each_ESU_unit;
        ESU_Total_Continuous_Maximum_charge_kW = Maximum_charge_kW_3_3C_One_String_High_Power_Batteries * Total_no_of_battery_strings_for_each_ESU_unit;
        ESU_Total_Continuous_Maximum_discharge_Amp = Maximum_discharge_A_3_3C_One_String_High_Power_Batteries * Total_no_of_battery_strings_for_each_ESU_unit;
        ESU_Total_Continuous_Maximum_discharge_kW = Maximum_discharge_kW_3_3C_One_String_High_Power_Batteries * Total_no_of_battery_strings_for_each_ESU_unit;
    end
        Energy_Storage_Converter_kW	= Required_power_in_kW;
        Energy_Storage_Converter_Current = Energy_Storage_Converter_kW / Voltage_Level_ESU;
        Energy_Storage_Transformer_kW = Required_power_in_kW;

        [Energy_Storage_Converter_kVA, Energy_Storage_Converter_Cost] = FindCloestMatch(1000, 10000, 1000, Energy_Storage_Converter_kW);
        [Energy_Storage_Transformer_kVA, Energy_Storage_Transformer_Cost] = FindCloestMatch(1000, 10000, 1000, Energy_Storage_Transformer_kW);
        Energy_Storage_Switchgear_Current = Energy_Storage_Transformer_kVA / (Voltage_Level_Switchboard * 3^0.5);

    % Common Calculateion
    No_of_ESU_units	= Input_No_of_ESU_units;
    Application_name = Input_Application_Name;
    Total_Installed_Energy_Start_of_Life = Nominal_Energy_in_kWh_High_Power_Batteries * Total_no_of_battery_strings_for_each_ESU_unit;
    Total_Usable_Energy_Start_of_Life = Energy_Start_of_Life_in_kWh * Total_no_of_battery_strings_for_each_ESU_unit;
    Total_Usable_Energy_End_of_Life_20_Reduction = Energy_End_of_Life_in_kWh * Total_no_of_battery_strings_for_each_ESU_unit;
    ESU_Total_Continuous_rms_charge_discharge_kW_T_25C = Continuous_rms_charge_discharge_kW_T_25C * Total_no_of_battery_strings_for_each_ESU_unit;
    ESU_Total_Continuous_rms_charge_discharge_Amp_T_25C	= Continuous_rms_charge_discharge_A_T_25C * Total_no_of_battery_strings_for_each_ESU_unit;
    Price_in_Euro_Total	= Price_in_Euro_Per_Unit * Total_no_of_battery_strings_for_each_ESU_unit;
    Price_in_Euro_Per_Installed_kWh	= Price_in_Euro_Total / Total_Installed_Energy_Start_of_Life;
    if mod(Total_no_of_battery_strings_for_each_ESU_unit, 2)
        ESU_Dimensions_W = (Total_no_of_battery_strings_for_each_ESU_unit - 1) * 900;
    else
        ESU_Dimensions_W = (Total_no_of_battery_strings_for_each_ESU_unit * 900) + 1200;
    end
    ESU_Dimensions_m2 = (ESU_Dimensions_D * ESU_Dimensions_W) / 1000000;
    ESU_Dimensions_m3 = (ESU_Dimensions_H * ESU_Dimensions_D * ESU_Dimensions_W) / 1000000000;
    ESU_Total_Units_Weight_in_kg = Total_no_of_battery_strings_for_each_ESU_unit * 1000;
    Total_Price_Engineering_ESU	= Hours_for_Engineering_ESU * Labour_Cost_Project_Management_Engineering_in_Norway;
    Total_Price_Included_Engineering_in_NOK_Without_Commisioning = Total_Price_Engineering_ESU;
    Commissioning_Days_10_Hours = Commissioning_Days * 10 * Labour_Cost_Service_Commissioning_Test;
    Hotel_and_Diet = Commissioning_Days * Cost_Hotel;
    Transport = Commissioning_Days * Cost_Transport;
    Total_Price_Commisioning_Energy_Storage_in_NOK  = Commissioning_Days_10_Hours + Hotel_and_Diet + Transport+Cost_Commissioning_Travel;
    Total_Price_in_NOK_Included_Commisioning = Total_Price_Included_Engineering_in_NOK_Without_Commisioning + Total_Price_Commisioning_Energy_Storage_in_NOK;
    EST_Secondary_Voltage = Voltage_Level_ESU / 2^0.5;
    Shore_Connection_1_Max_Voltage = max([Shore_Connection_1_Level_1_High Shore_Connection_1_Level_2_High Shore_Connection_1_Level_3_High]);
    Shore_Connection_2_Max_Voltage = max([Shore_Connection_2_Level_1_High Shore_Connection_2_Level_2_High Shore_Connection_2_Level_3_High]);
    Shore_Connection_1_kVA = Shore_Connection_1_Max_Voltage * 3^0.5 * Shore_Connection_1_Current;
    Shore_Connection_2_kVA = Shore_Connection_2_Max_Voltage * 3^0.5 * Shore_Connection_2_Current;
    [Shore_Connection_1_Transformer_kVA, Shore_Connection_1_Transformer_Cost] = FindCloestMatch(50, 500, 50, Shore_Connection_1_kVA);
    [Shore_Connection_2_Transformer_kVA, Shore_Connection_2_Transformer_Cost] = FindCloestMatch(50, 500, 50, Shore_Connection_2_kVA);
    [Shore_Connection_1_Converter_kVA, Shore_Connection_1_Converter_Cost] = FindCloestMatch(50, 500, 50, Shore_Connection_1_kVA);
    [Shore_Connection_2_Converter_kVA, Shore_Connection_2_Converter_Cost] = FindCloestMatch(50, 500, 50, Shore_Connection_2_kVA);
    
    Docs.Close; 
    invoke(Word,'Quit'); 
    delete(Word);

    assignin('base', 'outfilename', datestr(datetime('now','TimeZone','local'), 'yyyy-mm-ddTHH-MM-SS-FFF'));
    Dimensions = strcat(num2str(ESU_Dimensions_H), '*', num2str(ESU_Dimensions_W), '*', num2str(ESU_Dimensions_D));
    outputs = {'Quotation no.' 'Revision' 'Date' 'Author' 'Application Type' 'Customer' 'Customer project' 'Shipyard' ... 
        'Yard No.' 'Class' 'Total Price of ESU' 'Total Price Commissioning' 'Total Price Including Commissioning' 'No. of ESU' 'ESU Application Type' ... 
        'Total no. of battery strings for each ESU unit' 'Maximum voltage' 'Minimum voltage (1C discharge at 20% SOC)' 'Maximum charge current (1C)' ...
        'Maximum charge kW (1C)' 'Maximum discharge current (2C)' 'Maximum discharge kW (2C)' 'Continuous (rms) charge/discharge current ¶§T 25°„C' ...
        'Continuous (rms) charge/discharge kW ¶§T 25°„C' 'Total installed start of life energy' 'Total usable energy at end of life (100 - 20%)' ...
        'Total usable energy end of life (20% reduction)' 'Weight approx.' 'Dimensions (H*W*D)' 'Switchboard AC voltage (V)' 'Required power (kW)' ...
        'Required capacity (kWh)' 'ESU DC link voltage (V)' 'Number of ESU' 'Switchboard AC frequency (Hz)' 'Converter control method' ...
        'Number of shore connections' 'Energy storage converter rating (kVA)' 'Energy storage converter cost (NOK)' 'Energy storage transformer rating (kVA)' ...
        'Energy storage transformer cost (NOK)' 'Shore connection 1 transformer rating (kVA)' 'Shore connection 1 transformer cost (NOK)' ...
        'Shore connection 2 transformer rating (kVA)' 'Shore connection 2 transformer cost (NOK)' 'Shore connection 1 converter rating (kVA)' ...
        'Shore connection 1 converter cost (NOK)' 'Shore connection 2 converter rating (kVA)' 'Shore connection 2 converter cost (NOK)'; 
        Quotation_No Revision Revision_Date Author Application_name Customer Customer_project_no Shipyard Yard_No Classification_Society ...
        Total_Price_Engineering_ESU Total_Price_Commisioning_Energy_Storage_in_NOK Total_Price_in_NOK_Included_Commisioning No_of_ESU_units Application_name ...
        Total_no_of_battery_strings_for_each_ESU_unit Maximum_voltage Minimum_voltage_1C_discharge_at_20_SOC ESU_Total_Continuous_Maximum_charge_Amp ...
        ESU_Total_Continuous_Maximum_charge_kW ESU_Total_Continuous_Maximum_discharge_Amp ESU_Total_Continuous_Maximum_discharge_kW ...
        ESU_Total_Continuous_rms_charge_discharge_Amp_T_25C ESU_Total_Continuous_rms_charge_discharge_kW_T_25C Total_Installed_Energy_Start_of_Life ...
        Total_Usable_Energy_Start_of_Life Total_Usable_Energy_End_of_Life_20_Reduction ESU_Total_Units_Weight_in_kg Dimensions Voltage_Level_Switchboard ...
        Total_Power_Required Total_Capacity_Required Voltage_Level_ESU Input_No_of_ESU_units Frequency_Level_Switchboard Converter_Control_Method ...
        No_of_Shore_Connections Energy_Storage_Converter_kVA Energy_Storage_Converter_Cost Energy_Storage_Transformer_kVA Energy_Storage_Transformer_Cost ...
        Shore_Connection_1_Transformer_kVA Shore_Connection_1_Transformer_Cost Shore_Connection_2_Transformer_kVA Shore_Connection_2_Transformer_Cost ...
        Shore_Connection_1_Converter_kVA Shore_Connection_1_Converter_Cost Shore_Connection_2_Converter_kVA Shore_Connection_2_Converter_Cost
    };
    assignin('base', 'LENGTH', size(outputs, 2));
    assignin('base', 'outputs', outputs);
    
    report('report_demo_word2');
    return;
end