"""
This code is meant to run the multiple analyses that Cook County is looking for in their Internal
    project.

THey have 3 buildings to analyze:
    1. Cermak
    2. Markham
    3. Provident
Cook County has clean energy goals, thus, they are hoping to reach carbon neutral.
Additionally, they have a 40% federal incentive that they hope to get at with Cermak and Markham versus 30% with Provident.

    A) The analysis will follow the breakdown below:I'll run a multi-scenario analysis with increasing emissions reduction 
    (10%, 20%, ... , 100%) for all 3 buildings (Provident, Markham, Cermak).
    B) I'll also run 2 different sets of these scenarios for each building.
        a. PV+Battery
        b. PV+Battery+ASHP ... ASHP = air source heat pumps.
        c. + CHP for Provident
"""

"""
====================================================================================================================================================================
Cermak part a. 
PV+Battery
"""

using REopt
using HiGHS
using JSON
using JuMP
using CSV 
using DataFrames #to construct comparison
using XLSX 
using DelimitedFiles

# Setup inputs Cermak part a
data_file = "CermakA.JSON" 
input_data = JSON.parsefile("scenarios/$data_file")

cermak_rates = "C:/Users/dbernal/OneDrive - NREL/General - Cook County C2C/Internal - REopt Analysis/Custom Rates/Internal Sites/Cook County Internal Cermak New.json"
cermak_rates_1 = JSON.parsefile(cermak_rates)

function read_csv_without_bom(filepath::String)
    # Read the file content as a string
    file_content = read(filepath, String)
    
    # Remove BOM if present
    if startswith(file_content, "\ufeff")
        file_content = file_content[4:end]
    end
    
    # Parse the CSV content as Float64, assuming no header
    data = readdlm(IOBuffer(file_content), ',', Float64)
    return data
end

# Define the file path
cermak_electric_load = "C:/Users/dbernal/OneDrive - NREL/Non-shared files/REopt/C2C/Cook County/Internal/Load_profile_electric_DOC_Cermak.csv"

# Read the CSV file
cermak_loads_kw = read_csv_without_bom(cermak_electric_load)

# Convert matrix to a one-dimensional array 
cermak_loads_kw = reshape(cermak_loads_kw, :)  # This flattens the matrix into a one-dimensional array
cermak_loads_kw = cermak_loads_kw[8761:17520] #take off the hours and leave the loads
println("Correctly obtained data_file")

#the lat/long will be representative of the regions (MW, NE, S, W)
#cities chosen are Chicago, Boston, Houston, San Francisco
cities = ["Chicago", "Chicago", "Chicago", "Chicago", "Chicago", "Chicago", "Chicago", "Chicago", "Chicago", "Chicago"]
lat = [41.834, 41.834, 41.834, 41.834, 41.834, 41.834, 41.834, 41.834, 41.834, 41.834]
long = [-88.044, -88.044, -88.044, -88.044, -88.044, -88.044, -88.044, -88.044, -88.044, -88.044]

#emissions reduction fraction every scenario by 10%
emissions_reduction_min = [0.10, 0.20, 0.30, 0.40, 0.50, 0.60, 0.70, 0.80, 0.90, 1.00]

Site = [raw"Cermak-10%", raw"Cermak-20%", raw"Cermak-30%", raw"Cermak-40%", raw"Cermak-50%", raw"Cermak-60%", raw"Cermak-70%", raw"Cermak-80%", raw"Cermak-90%", raw"Cermak-100%"]

site_analysis = []

sites_iter = eachindex(lat)
for i in sites_iter
    input_data_site = copy(input_data)
    # Site Specific
    input_data_site["Site"]["latitude"] = lat[i]
    input_data_site["Site"]["longitude"] = long[i]
    input_data_site["ElectricLoad"]["loads_kw"] = cermak_loads_kw
    input_data_site["ElectricTariff"]["urdb_response"] = cermak_rates_1

    #location of PV being mounted, both, roof, or ground
    input_data_site["PV"]["location"] = "roof"

    #emissions reduction min 
    input_data_site["Site"]["CO2_emissions_reduction_min_fraction"] = emissions_reduction_min[i]
                
    s = Scenario(input_data_site)
    inputs = REoptInputs(s)

     # HiGHS solver
     m1 = Model(optimizer_with_attributes(HiGHS.Optimizer, 
     "time_limit" => 600.0,
     "mip_rel_gap" => 0.01,
     "output_flag" => false, 
     "log_to_console" => false)
     )

    m2 = Model(optimizer_with_attributes(HiGHS.Optimizer, 
     "time_limit" => 600.0,
     "mip_rel_gap" => 0.01,
     "output_flag" => false, 
     "log_to_console" => false)
     )            

    results = run_reopt([m1,m2], inputs)
    append!(site_analysis, [(input_data_site, results)])

    sleep(60)
end
println("Completed optimization")

#write onto JSON file
write.("./results/cook_county_cermakA.json", JSON.json(site_analysis))
println("Successfully printed results on JSON file")

# Populate the DataFrame with the results produced and inputs
df = DataFrame(
    City = cities,
    PV_size = [round(site_analysis[i][2]["PV"]["size_kw"], digits=0) for i in sites_iter],
    PV_year1_production = [round(site_analysis[i][2]["PV"]["year_one_energy_produced_kwh"], digits=0) for i in sites_iter],
    PV_annual_energy_production_avg = [round(site_analysis[i][2]["PV"]["annual_energy_produced_kwh"], digits=0) for i in sites_iter],
    PV_energy_lcoe = [round(site_analysis[i][2]["PV"]["lcoe_per_kwh"], digits=0) for i in sites_iter],
    PV_energy_exported = [round(site_analysis[i][2]["PV"]["annual_energy_exported_kwh"], digits=0) for i in sites_iter],
    PV_energy_curtailed = [sum(site_analysis[i][2]["PV"]["electric_curtailed_series_kw"]) for i in sites_iter],
    PV_energy_to_Battery_year1 = [sum(site_analysis[i][2]["PV"]["electric_to_storage_series_kw"]) for i in sites_iter],
    Battery_size_kW = [round(site_analysis[i][2]["ElectricStorage"]["size_kw"], digits=0) for i in sites_iter], 
    Battery_size_kWh = [round(site_analysis[i][2]["ElectricStorage"]["size_kwh"], digits=0) for i in sites_iter], 
    Battery_serve_electric_load = [sum(site_analysis[i][2]["ElectricStorage"]["storage_to_load_series_kw"], digits=0) for i in sites_iter], 
    Grid_Electricity_Supplied_kWh_annual = [round(site_analysis[i][2]["ElectricUtility"]["annual_energy_supplied_kwh"], digits=0) for i in sites_iter],
    BAU_Existing_Boiler_Fuel_Consump_MMBtu = [round(site_analysis[i][2]["ExistingBoiler"]["annual_fuel_consumption_mmbtu_bau"], digits=0) for i in sites_iter],
    BAU_Existing_Boiler_Thermal_Prod_MMBtu = [round(site_analysis[i][2]["ExistingBoiler"]["annual_thermal_production_mmbtu_bau"], digits=0) for i in sites_iter],
    NG_Annual_Consumption_MMBtu = [round(site_analysis[i][2]["ExistingBoiler"]["annual_fuel_consumption_mmbtu"], digits=0) for i in sites_iter],
    Total_Annual_Emissions_CO2 = [round(site_analysis[i][2]["Site"]["annual_emissions_tonnes_CO2"], digits=4) for i in sites_iter],
    ElecUtility_Annual_Emissions_CO2 = [round(site_analysis[i][2]["ElectricUtility"]["annual_emissions_tonnes_CO2"], digits=4) for i in sites_iter],
    BAU_Total_Annual_Emissions_CO2 = [round(site_analysis[i][2]["Site"]["annual_emissions_tonnes_CO2_bau"], digits=4) for i in sites_iter],
    LifeCycle_Emissions_CO2 = [round(site_analysis[i][2]["Site"]["lifecycle_emissions_tonnes_CO2"], digits=2) for i in sites_iter],
    BAU_LifeCycle_Emissions_CO2 = [round(site_analysis[i][2]["Site"]["lifecycle_emissions_tonnes_CO2_bau"], digits=2) for i in sites_iter],
    NG_LifeCycle_Emissions_CO2 = [round(site_analysis[i][2]["Site"]["lifecycle_emissions_from_fuelburn_tonnes_CO2"], digits=2) for i in sites_iter],
    Emissions_from_NG = [round(site_analysis[i][2]["Site"]["annual_emissions_from_fuelburn_tonnes_CO2"], digits=0) for i in sites_iter],
    LifeCycle_Emission_Reduction_Fraction = [round(site_analysis[i][2]["Site"]["lifecycle_emissions_reduction_CO2_fraction"], digits=2) for i in sites_iter],
    npv = [round(site_analysis[i][2]["Financial"]["npv"], digits=2) for i in sites_iter],
    lcc = [round(site_analysis[i][2]["Financial"]["lcc"], digits=2) for i in sites_iter]
    )
println(df)

# Define path to xlsx file
file_storage_location = "results/Cook_County_results.xlsx"

# Check if the Excel file already exists
if isfile(file_storage_location)
    # Open the Excel file in read-write mode
    XLSX.openxlsx(file_storage_location, mode="rw") do xf
        counter = 0
        while true
            sheet_name = "CermakA_" * string(counter)
            try
                sheet = xf[sheet_name]
                counter += 1
            catch
                break
            end
        end
        sheet_name = "CermakA_" * string(counter)
        # Add new sheet
        XLSX.addsheet!(xf, sheet_name)
        # Write DataFrame to the new sheet
        XLSX.writetable!(xf[sheet_name], df)
    end
else
    # Write DataFrame to a new Excel file
    XLSX.writetable!(file_storage_location, df)
end

println("Successful write into XLSX file: $file_storage_location")

"""
=======================================================================================================================================================================
Cermak part b.
PV+Battery+ASHP
"""

# Setup inputs Cermak Part b
data_file = "CermakB.JSON" 
input_data = JSON.parsefile("scenarios/$data_file")

cermak_rates = "C:/Users/dbernal/OneDrive - NREL/General - Cook County C2C/Internal - REopt Analysis/Custom Rates/Internal Sites/Cook County Internal Cermak New.json"
cermak_rates_1 = JSON.parsefile(cermak_rates)

cermak_electric_load = "C:/Users/dbernal/OneDrive - NREL/General - Cook County C2C/Internal - REopt Analysis/REopt Loads/Load_profile_electric_DOC_Cermak.csv"

println("Correctly obtained data_file")

#the lat/long will be representative of the regions (MW, NE, S, W)
#cities chosen are Chicago, Boston, Houston, San Francisco
cities = ["Chicago", "Chicago", "Chicago", "Chicago", "Chicago", "Chicago", "Chicago", "Chicago", "Chicago", "Chicago"]
lat = [41.834, 41.834, 41.834, 41.834, 41.834, 41.834, 41.834, 41.834, 41.834, 41.834]
long = [-88.044, -88.044, -88.044, -88.044, -88.044, -88.044, -88.044, -88.044, -88.044, -88.044]

#emissions reduction fraction every scenario by 10%
emissions_reduction_min = [0.10, 0.20, 0.30, 0.40, 0.50, 0.60, 0.70, 0.80, 0.90, 1.00]

Site = [raw"Cermak-10%", raw"Cermak-20%", raw"Cermak-30%", raw"Cermak-40%", raw"Cermak-50%", raw"Cermak-60%", raw"Cermak-70%", raw"Cermak-80%", raw"Cermak-90%", raw"Cermak-100%"]

site_analysis = []

sites_iter = eachindex(lat)
for i in sites_iter
    input_data_site = copy(input_data)
    # Site Specific
    input_data_site["Site"]["latitude"] = lat[i]
    input_data_site["Site"]["longitude"] = long[i]
    input_data_site["ElectricLoad"]["path_to_csv"] = cermak_electric_load
    input_data_site["ElectricTariff"]["urdb_response"] = cermak_rates_1
    
    #location of PV being mounted, both, roof, or ground
    input_data_site["PV"]["location"] = "roof"

    #emissions reduction min 
    input_data_site["Site"]["CO2_emissions_reduction_min_fraction"] = emissions_reduction_min[i]
                
    s = Scenario(input_data_site)
    inputs = REoptInputs(s)

     # HiGHS solver
     m1 = Model(optimizer_with_attributes(HiGHS.Optimizer, 
     "time_limit" => 450.0,
     "mip_rel_gap" => 0.01,
     "output_flag" => false, 
     "log_to_console" => false)
     )

    m2 = Model(optimizer_with_attributes(HiGHS.Optimizer, 
     "time_limit" => 450.0,
     "mip_rel_gap" => 0.01,
     "output_flag" => false, 
     "log_to_console" => false)
     )            

    results = run_reopt([m1,m2], inputs)
    append!(site_analysis, [(input_data_site, results)])

    sleep(180)
end
println("Completed optimization")

#write onto JSON file
write.("./results/cook_county_cermakA.json", JSON.json(site_analysis))
println("Successfully printed results on JSON file")

# Populate the DataFrame with the results produced and inputs
df = DataFrame(
    City = cities,
    PV_size = [round(site_analysis[i][2]["PV"]["size_kw"], digits=0) for i in sites_iter],
    PV_year1_production = [round(site_analysis[i][2]["PV"]["year_one_energy_produced_kwh"], digits=0) for i in sites_iter],
    PV_annual_energy_production_avg = [round(site_analysis[i][2]["PV"]["annual_energy_produced_kwh"], digits=0) for i in sites_iter],
    PV_energy_lcoe = [round(site_analysis[i][2]["PV"]["lcoe_per_kwh"], digits=0) for i in sites_iter],
    PV_energy_exported = [round(site_analysis[i][2]["PV"]["annual_energy_exported_kwh"], digits=0) for i in sites_iter],
    PV_energy_curtailed = [sum(site_analysis[i][2]["PV"]["electric_curtailed_series_kw"]) for i in sites_iter],
    PV_energy_to_Battery_year1 = [sum(site_analysis[i][2]["PV"]["electric_to_storage_series_kw"]) for i in sites_iter],
    Battery_size_kW = [round(site_analysis[i][2]["ElectricStorage"]["size_kw"], digits=0) for i in sites_iter], 
    Battery_size_kWh = [round(site_analysis[i][2]["ElectricStorage"]["size_kwh"], digits=0) for i in sites_iter], 
    Battery_serve_electric_load = [sum(site_analysis[i][2]["ElectricStorage"]["storage_to_load_series_kw"], digits=0) for i in sites_iter], 
    ASHP_size_tonhour = [round(site_analysis[i][2]["ASHP"]["size_ton"], digits=0) for i in sites_iter],
    ASHP_annual_electric_consumption_kwh = [round(site_analysis[i][2]["ASHP"]["annual_electric_consumption_kwh"], digits=0) for i in sites_iter],
    ASHP_annual_thermal_production_mmbtu = [round(site_analysis[i][2]["ASHP"]["annual_thermal_production_mmbtu"], digits=0) for i in sites_iter],
    ASHP_annual_cooling_tonhour = [round(site_analysis[i][2]["ASHP"]["annual_thermal_production_tonhour"], digits=0) for i in sites_iter],
    Grid_Electricity_Supplied_kWh_annual = [round(site_analysis[i][2]["ElectricUtility"]["annual_energy_supplied_kwh"], digits=0) for i in sites_iter],
    Annual_Total_HeatingLoad_MMBtu = [round(site_analysis[i][2]["HeatingLoad"]["annual_calculated_total_heating_thermal_load_mmbtu"], digits=0) for i in sites_iter],
    BAU_Existing_Boiler_Fuel_Consump_MMBtu = [round(site_analysis[i][2]["ExistingBoiler"]["annual_fuel_consumption_mmbtu_bau"], digits=0) for i in sites_iter],
    BAU_Existing_Boiler_Thermal_Prod_MMBtu = [round(site_analysis[i][2]["ExistingBoiler"]["annual_thermal_production_mmbtu_bau"], digits=0) for i in sites_iter],
    NG_Annual_Consumption_MMBtu = [round(site_analysis[i][2]["ExistingBoiler"]["annual_fuel_consumption_mmbtu"], digits=0) for i in sites_iter],
    Total_Annual_Emissions_CO2 = [round(site_analysis[i][2]["Site"]["annual_emissions_tonnes_CO2"], digits=4) for i in sites_iter],
    ElecUtility_Annual_Emissions_CO2 = [round(site_analysis[i][2]["ElectricUtility"]["annual_emissions_tonnes_CO2"], digits=4) for i in sites_iter],
    BAU_Total_Annual_Emissions_CO2 = [round(site_analysis[i][2]["Site"]["annual_emissions_tonnes_CO2_bau"], digits=4) for i in sites_iter],
    LifeCycle_Emissions_CO2 = [round(site_analysis[i][2]["Site"]["lifecycle_emissions_tonnes_CO2"], digits=2) for i in sites_iter],
    BAU_LifeCycle_Emissions_CO2 = [round(site_analysis[i][2]["Site"]["lifecycle_emissions_tonnes_CO2_bau"], digits=2) for i in sites_iter],
    NG_LifeCycle_Emissions_CO2 = [round(site_analysis[i][2]["Site"]["lifecycle_emissions_from_fuelburn_tonnes_CO2"], digits=2) for i in sites_iter],
    Emissions_from_NG = [round(site_analysis[i][2]["Site"]["annual_emissions_from_fuelburn_tonnes_CO2"], digits=0) for i in sites_iter],
    LifeCycle_Emission_Reduction_Fraction = [round(site_analysis[i][2]["Site"]["lifecycle_emissions_reduction_CO2_fraction"], digits=2) for i in sites_iter],
    npv = [round(site_analysis[i][2]["Financial"]["npv"], digits=2) for i in sites_iter],
    lcc = [round(site_analysis[i][2]["Financial"]["lcc"], digits=2) for i in sites_iter]
    )
println(df)

# Define path to xlsx file
file_storage_location = "results/Cook_County_results.xlsx"

# Check if the Excel file already exists
if isfile(file_storage_location)
    # Open the Excel file in read-write mode
    XLSX.openxlsx(file_storage_location, mode="rw") do xf
        counter = 0
        while true
            sheet_name = "CermakB_" * string(counter)
            try
                sheet = xf[sheet_name]
                counter += 1
            catch
                break
            end
        end
        sheet_name = "CermakB_" * string(counter)
        # Add new sheet
        XLSX.addsheet!(xf, sheet_name)
        # Write DataFrame to the new sheet
        XLSX.writetable!(xf[sheet_name], df)
    end
else
    # Write DataFrame to a new Excel file
    XLSX.writetable!(file_storage_location, df)
end

println("Successful write into XLSX file: $file_storage_location")

"""
=======================================================================================================================================================================
Markham Part A
PV+Battery
"""
"""
# Setup inputs Markham part a
data_file = "MarkhamA.JSON" 
input_data = JSON.parsefile("scenarios/$data_file")

markham_rates = "C:/Users/dbernal/OneDrive - NREL/General - Cook County C2C/Internal - REopt Analysis/Custom Rates/Internal Sites/Cook County Internal Markham New.json"
markham_rates_1 = JSON.parsefile(markham_rates)

cermak_electric_load = "C:/Users/dbernal/OneDrive - NREL/General - Cook County C2C/Internal - REopt Analysis/REopt Loads/Load_profile_electric_Markham.csv"

println("Correctly obtained data_file")

#the lat/long will be representative of the regions (MW, NE, S, W)
#cities chosen are Chicago, Boston, Houston, San Francisco
cities = ["Chicago", "Chicago", "Chicago", "Chicago", "Chicago", "Chicago", "Chicago", "Chicago", "Chicago", "Chicago"]
lat = [41.834, 41.834, 41.834, 41.834, 41.834, 41.834, 41.834, 41.834, 41.834, 41.834]
long = [-88.044, -88.044, -88.044, -88.044, -88.044, -88.044, -88.044, -88.044, -88.044, -88.044]

#emissions reduction fraction every scenario by 10%
emissions_reduction_min = [0.10, 0.20, 0.30, 0.40, 0.50, 0.60, 0.70, 0.80, 0.90, 1.00]

#Existing PV at Markham
markham_existing_pv = [743, 743, 743, 743, 743, 743, 743, 743, 743, 743]

Site = [raw"Markham-10%", raw"Markham-20%", raw"Markham-30%", raw"Markham-40%", raw"Markham-50%", raw"Markham-60%", raw"Markham-70%", raw"Markham-80%", raw"Markham-90%", raw"Markham-100%"]

site_analysis = []

sites_iter = eachindex(lat)
for i in sites_iter
    input_data_site = copy(input_data)
    # Site Specific
    input_data_site["Site"]["latitude"] = lat[i]
    input_data_site["Site"]["longitude"] = long[i]
    input_data_site["ElectricLoad"]["path_to_csv"] = cermak_electric_load
    input_data_site["ElectricTariff"]["urdb_response"] = cermak_rates_1
    
    #existing PV on Markham
    input_data_site["PV"]["existing_kw"] = markham_existing_pv[i]

    #emissions reduction min 
    input_data_site["Site"]["CO2_emissions_reduction_min_fraction"] = emissions_reduction_min[i]
                
    s = Scenario(input_data_site)
    inputs = REoptInputs(s)

     # HiGHS solver
     m1 = Model(optimizer_with_attributes(HiGHS.Optimizer, 
     "time_limit" => 450.0,
     "mip_rel_gap" => 0.01,
     "output_flag" => false, 
     "log_to_console" => false)
     )

    m2 = Model(optimizer_with_attributes(HiGHS.Optimizer, 
     "time_limit" => 450.0,
     "mip_rel_gap" => 0.01,
     "output_flag" => false, 
     "log_to_console" => false)
     )            

    results = run_reopt([m1,m2], inputs)
    append!(site_analysis, [(input_data_site, results)])

    sleep(180)
end
println("Completed optimization")

#write onto JSON file
write.("./results/cook_county_markhamA.json", JSON.json(site_analysis))
println("Successfully printed results on JSON file")

# Populate the DataFrame with the results produced and inputs
df = DataFrame(
    City = cities,
    PV_size = [round(site_analysis[i][2]["PV"]["size_kw"], digits=0) for i in sites_iter],
    PV_year1_production = [round(site_analysis[i][2]["PV"]["year_one_energy_produced_kwh"], digits=0) for i in sites_iter],
    PV_annual_energy_production_avg = [round(site_analysis[i][2]["PV"]["annual_energy_produced_kwh"], digits=0) for i in sites_iter],
    PV_energy_lcoe = [round(site_analysis[i][2]["PV"]["lcoe_per_kwh"], digits=0) for i in sites_iter],
    PV_energy_exported = [round(site_analysis[i][2]["PV"]["annual_energy_exported_kwh"], digits=0) for i in sites_iter],
    PV_energy_curtailed = [sum(site_analysis[i][2]["PV"]["electric_curtailed_series_kw"]) for i in sites_iter],
    PV_energy_to_Battery_year1 = [sum(site_analysis[i][2]["PV"]["electric_to_storage_series_kw"]) for i in sites_iter],
    Battery_size_kW = [round(site_analysis[i][2]["ElectricStorage"]["size_kw"], digits=0) for i in sites_iter], 
    Battery_size_kWh = [round(site_analysis[i][2]["ElectricStorage"]["size_kwh"], digits=0) for i in sites_iter], 
    Battery_serve_electric_load = [sum(site_analysis[i][2]["ElectricStorage"]["storage_to_load_series_kw"], digits=0) for i in sites_iter], 
    Grid_Electricity_Supplied_kWh_annual = [round(site_analysis[i][2]["ElectricUtility"]["annual_energy_supplied_kwh"], digits=0) for i in sites_iter],
    BAU_Existing_Boiler_Fuel_Consump_MMBtu = [round(site_analysis[i][2]["ExistingBoiler"]["annual_fuel_consumption_mmbtu_bau"], digits=0) for i in sites_iter],
    BAU_Existing_Boiler_Thermal_Prod_MMBtu = [round(site_analysis[i][2]["ExistingBoiler"]["annual_thermal_production_mmbtu_bau"], digits=0) for i in sites_iter],
    NG_Annual_Consumption_MMBtu = [round(site_analysis[i][2]["ExistingBoiler"]["annual_fuel_consumption_mmbtu"], digits=0) for i in sites_iter],
    Total_Annual_Emissions_CO2 = [round(site_analysis[i][2]["Site"]["annual_emissions_tonnes_CO2"], digits=4) for i in sites_iter],
    ElecUtility_Annual_Emissions_CO2 = [round(site_analysis[i][2]["ElectricUtility"]["annual_emissions_tonnes_CO2"], digits=4) for i in sites_iter],
    BAU_Total_Annual_Emissions_CO2 = [round(site_analysis[i][2]["Site"]["annual_emissions_tonnes_CO2_bau"], digits=4) for i in sites_iter],
    LifeCycle_Emissions_CO2 = [round(site_analysis[i][2]["Site"]["lifecycle_emissions_tonnes_CO2"], digits=2) for i in sites_iter],
    BAU_LifeCycle_Emissions_CO2 = [round(site_analysis[i][2]["Site"]["lifecycle_emissions_tonnes_CO2_bau"], digits=2) for i in sites_iter],
    NG_LifeCycle_Emissions_CO2 = [round(site_analysis[i][2]["Site"]["lifecycle_emissions_from_fuelburn_tonnes_CO2"], digits=2) for i in sites_iter],
    Emissions_from_NG = [round(site_analysis[i][2]["Site"]["annual_emissions_from_fuelburn_tonnes_CO2"], digits=0) for i in sites_iter],
    LifeCycle_Emission_Reduction_Fraction = [round(site_analysis[i][2]["Site"]["lifecycle_emissions_reduction_CO2_fraction"], digits=2) for i in sites_iter],
    npv = [round(site_analysis[i][2]["Financial"]["npv"], digits=2) for i in sites_iter],
    lcc = [round(site_analysis[i][2]["Financial"]["lcc"], digits=2) for i in sites_iter]
    )
println(df)

# Define path to xlsx file
file_storage_location = "results/Cook_County_results.xlsx"

# Check if the Excel file already exists
if isfile(file_storage_location)
    # Open the Excel file in read-write mode
    XLSX.openxlsx(file_storage_location, mode="rw") do xf
        counter = 0
        while true
            sheet_name = "MarkhamA_" * string(counter)
            try
                sheet = xf[sheet_name]
                counter += 1
            catch
                break
            end
        end
        sheet_name = "MarkhamA_" * string(counter)
        # Add new sheet
        XLSX.addsheet!(xf, sheet_name)
        # Write DataFrame to the new sheet
        XLSX.writetable!(xf[sheet_name], df)
    end
else
    # Write DataFrame to a new Excel file
    XLSX.writetable!(file_storage_location, df)
end

println("Successful write into XLSX file: $file_storage_location")

"""
=======================================================================================================================================================================
Markham Part B
PV+Battery+ASHP
"""
# Setup inputs Markham part a
data_file = "MarkhamB.JSON" 
input_data = JSON.parsefile("scenarios/$data_file")

markham_rates = "C:/Users/dbernal/OneDrive - NREL/General - Cook County C2C/Internal - REopt Analysis/Custom Rates/Internal Sites/Cook County Internal Markham New.json"
markham_rates_1 = JSON.parsefile(markham_rates)

cermak_electric_load = "C:/Users/dbernal/OneDrive - NREL/General - Cook County C2C/Internal - REopt Analysis/REopt Loads/Load_profile_electric_Markham.csv"

println("Correctly obtained data_file")

#the lat/long will be representative of the regions (MW, NE, S, W)
#cities chosen are Chicago, Boston, Houston, San Francisco
cities = ["Chicago", "Chicago", "Chicago", "Chicago", "Chicago", "Chicago", "Chicago", "Chicago", "Chicago", "Chicago"]
lat = [41.834, 41.834, 41.834, 41.834, 41.834, 41.834, 41.834, 41.834, 41.834, 41.834]
long = [-88.044, -88.044, -88.044, -88.044, -88.044, -88.044, -88.044, -88.044, -88.044, -88.044]

#emissions reduction fraction every scenario by 10%
emissions_reduction_min = [0.10, 0.20, 0.30, 0.40, 0.50, 0.60, 0.70, 0.80, 0.90, 1.00]

#Existing PV at Markham
markham_existing_pv = [743, 743, 743, 743, 743, 743, 743, 743, 743, 743]

Site = [raw"Markham-10%", raw"Markham-20%", raw"Markham-30%", raw"Markham-40%", raw"Markham-50%", raw"Markham-60%", raw"Markham-70%", raw"Markham-80%", raw"Markham-90%", raw"Markham-100%"]

site_analysis = []

sites_iter = eachindex(lat)
for i in sites_iter
    input_data_site = copy(input_data)
    # Site Specific
    input_data_site["Site"]["latitude"] = lat[i]
    input_data_site["Site"]["longitude"] = long[i]
    input_data_site["ElectricLoad"]["path_to_csv"] = cermak_electric_load
    input_data_site["ElectricTariff"]["urdb_response"] = cermak_rates_1

        #existing PV on Markham
    input_data_site["PV"]["existing_kw"] = markham_existing_pv[i]

    #emissions reduction min 
    input_data_site["Site"]["CO2_emissions_reduction_min_fraction"] = emissions_reduction_min[i]
                
    s = Scenario(input_data_site)
    inputs = REoptInputs(s)

     # HiGHS solver
     m1 = Model(optimizer_with_attributes(HiGHS.Optimizer, 
     "time_limit" => 450.0,
     "mip_rel_gap" => 0.01,
     "output_flag" => false, 
     "log_to_console" => false)
     )

    m2 = Model(optimizer_with_attributes(HiGHS.Optimizer, 
     "time_limit" => 450.0,
     "mip_rel_gap" => 0.01,
     "output_flag" => false, 
     "log_to_console" => false)
     )            

    results = run_reopt([m1,m2], inputs)
    append!(site_analysis, [(input_data_site, results)])

    sleep(180)
end
println("Completed optimization")

#write onto JSON file
write.("./results/cook_county_markhamB.json", JSON.json(site_analysis))
println("Successfully printed results on JSON file")

# Populate the DataFrame with the results produced and inputs
df = DataFrame(
    City = cities,
    PV_size = [round(site_analysis[i][2]["PV"]["size_kw"], digits=0) for i in sites_iter],
    PV_year1_production = [round(site_analysis[i][2]["PV"]["year_one_energy_produced_kwh"], digits=0) for i in sites_iter],
    PV_annual_energy_production_avg = [round(site_analysis[i][2]["PV"]["annual_energy_produced_kwh"], digits=0) for i in sites_iter],
    PV_energy_lcoe = [round(site_analysis[i][2]["PV"]["lcoe_per_kwh"], digits=0) for i in sites_iter],
    PV_energy_exported = [round(site_analysis[i][2]["PV"]["annual_energy_exported_kwh"], digits=0) for i in sites_iter],
    PV_energy_curtailed = [sum(site_analysis[i][2]["PV"]["electric_curtailed_series_kw"]) for i in sites_iter],
    PV_energy_to_Battery_year1 = [sum(site_analysis[i][2]["PV"]["electric_to_storage_series_kw"]) for i in sites_iter],
    Battery_size_kW = [round(site_analysis[i][2]["ElectricStorage"]["size_kw"], digits=0) for i in sites_iter], 
    Battery_size_kWh = [round(site_analysis[i][2]["ElectricStorage"]["size_kwh"], digits=0) for i in sites_iter], 
    Battery_serve_electric_load = [sum(site_analysis[i][2]["ElectricStorage"]["storage_to_load_series_kw"], digits=0) for i in sites_iter], 
    ASHP_size_tonhour = [round(site_analysis[i][2]["ASHP"]["size_ton"], digits=0) for i in sites_iter],
    ASHP_annual_electric_consumption_kwh = [round(site_analysis[i][2]["ASHP"]["annual_electric_consumption_kwh"], digits=0) for i in sites_iter],
    ASHP_annual_thermal_production_mmbtu = [round(site_analysis[i][2]["ASHP"]["annual_thermal_production_mmbtu"], digits=0) for i in sites_iter],
    ASHP_annual_cooling_tonhour = [round(site_analysis[i][2]["ASHP"]["annual_thermal_production_tonhour"], digits=0) for i in sites_iter],
    Grid_Electricity_Supplied_kWh_annual = [round(site_analysis[i][2]["ElectricUtility"]["annual_energy_supplied_kwh"], digits=0) for i in sites_iter],
    Annual_Total_HeatingLoad_MMBtu = [round(site_analysis[i][2]["HeatingLoad"]["annual_calculated_total_heating_thermal_load_mmbtu"], digits=0) for i in sites_iter],
    BAU_Existing_Boiler_Fuel_Consump_MMBtu = [round(site_analysis[i][2]["ExistingBoiler"]["annual_fuel_consumption_mmbtu_bau"], digits=0) for i in sites_iter],
    BAU_Existing_Boiler_Thermal_Prod_MMBtu = [round(site_analysis[i][2]["ExistingBoiler"]["annual_thermal_production_mmbtu_bau"], digits=0) for i in sites_iter],
    NG_Annual_Consumption_MMBtu = [round(site_analysis[i][2]["ExistingBoiler"]["annual_fuel_consumption_mmbtu"], digits=0) for i in sites_iter],
    Total_Annual_Emissions_CO2 = [round(site_analysis[i][2]["Site"]["annual_emissions_tonnes_CO2"], digits=4) for i in sites_iter],
    ElecUtility_Annual_Emissions_CO2 = [round(site_analysis[i][2]["ElectricUtility"]["annual_emissions_tonnes_CO2"], digits=4) for i in sites_iter],
    BAU_Total_Annual_Emissions_CO2 = [round(site_analysis[i][2]["Site"]["annual_emissions_tonnes_CO2_bau"], digits=4) for i in sites_iter],
    LifeCycle_Emissions_CO2 = [round(site_analysis[i][2]["Site"]["lifecycle_emissions_tonnes_CO2"], digits=2) for i in sites_iter],
    BAU_LifeCycle_Emissions_CO2 = [round(site_analysis[i][2]["Site"]["lifecycle_emissions_tonnes_CO2_bau"], digits=2) for i in sites_iter],
    NG_LifeCycle_Emissions_CO2 = [round(site_analysis[i][2]["Site"]["lifecycle_emissions_from_fuelburn_tonnes_CO2"], digits=2) for i in sites_iter],
    Emissions_from_NG = [round(site_analysis[i][2]["Site"]["annual_emissions_from_fuelburn_tonnes_CO2"], digits=0) for i in sites_iter],
    LifeCycle_Emission_Reduction_Fraction = [round(site_analysis[i][2]["Site"]["lifecycle_emissions_reduction_CO2_fraction"], digits=2) for i in sites_iter],
    npv = [round(site_analysis[i][2]["Financial"]["npv"], digits=2) for i in sites_iter],
    lcc = [round(site_analysis[i][2]["Financial"]["lcc"], digits=2) for i in sites_iter]
    )
println(df)

# Define path to xlsx file
file_storage_location = "results/Cook_County_results.xlsx"

# Check if the Excel file already exists
if isfile(file_storage_location)
    # Open the Excel file in read-write mode
    XLSX.openxlsx(file_storage_location, mode="rw") do xf
        counter = 0
        while true
            sheet_name = "MarkhamB_" * string(counter)
            try
                sheet = xf[sheet_name]
                counter += 1
            catch
                break
            end
        end
        sheet_name = "MarkhamB_" * string(counter)
        # Add new sheet
        XLSX.addsheet!(xf, sheet_name)
        # Write DataFrame to the new sheet
        XLSX.writetable!(xf[sheet_name], df)
    end
else
    # Write DataFrame to a new Excel file
    XLSX.writetable!(file_storage_location, df)
end

println("Successful write into XLSX file: $file_storage_location")

"""
=======================================================================================================================================================================
Provident Part A
PV+Battery
"""

# Setup inputs Provident part a
data_file = "ProvidentA.JSON" 
input_data = JSON.parsefile("scenarios/$data_file")

markham_rates = "C:/Users/dbernal/OneDrive - NREL/General - Cook County C2C/Internal - REopt Analysis/Custom Rates/Internal Sites/Cook County Internal Provident New.json"
markham_rates_1 = JSON.parsefile(markham_rates)

cermak_electric_load = "C:/Users/dbernal/OneDrive - NREL/General - Cook County C2C/Internal - REopt Analysis/REopt Loads/Load_profile_electric_Provident_hourly.csv"

println("Correctly obtained data_file")

#the lat/long will be representative of the regions (MW, NE, S, W)
#cities chosen are Chicago, Boston, Houston, San Francisco
cities = ["Chicago", "Chicago", "Chicago", "Chicago", "Chicago", "Chicago", "Chicago", "Chicago", "Chicago", "Chicago"]
lat = [41.834, 41.834, 41.834, 41.834, 41.834, 41.834, 41.834, 41.834, 41.834, 41.834]
long = [-88.044, -88.044, -88.044, -88.044, -88.044, -88.044, -88.044, -88.044, -88.044, -88.044]

#emissions reduction fraction every scenario by 10%
emissions_reduction_min = [0.10, 0.20, 0.30, 0.40, 0.50, 0.60, 0.70, 0.80, 0.90, 1.00]

Site = [raw"Markham-10%", raw"Markham-20%", raw"Markham-30%", raw"Markham-40%", raw"Markham-50%", raw"Markham-60%", raw"Markham-70%", raw"Markham-80%", raw"Markham-90%", raw"Markham-100%"]

site_analysis = []

sites_iter = eachindex(lat)
for i in sites_iter
    input_data_site = copy(input_data)
    # Site Specific
    input_data_site["Site"]["latitude"] = lat[i]
    input_data_site["Site"]["longitude"] = long[i]
    input_data_site["ElectricLoad"]["path_to_csv"] = cermak_electric_load
    input_data_site["ElectricTariff"]["urdb_response"] = cermak_rates_1
    input_data_site["DomesticHotWaterLoad"]["annual_mmbtu"] = avg_ng_load[i] * 8760

    #emissions reduction min 
    input_data_site["Site"]["CO2_emissions_reduction_min_fraction"] = emissions_reduction_min[i]
                
    s = Scenario(input_data_site)
    inputs = REoptInputs(s)

     # HiGHS solver
     m1 = Model(optimizer_with_attributes(HiGHS.Optimizer, 
     "time_limit" => 450.0,
     "mip_rel_gap" => 0.01,
     "output_flag" => false, 
     "log_to_console" => false)
     )

    m2 = Model(optimizer_with_attributes(HiGHS.Optimizer, 
     "time_limit" => 450.0,
     "mip_rel_gap" => 0.01,
     "output_flag" => false, 
     "log_to_console" => false)
     )            

    results = run_reopt([m1,m2], inputs)
    append!(site_analysis, [(input_data_site, results)])

    sleep(180)
end
println("Completed optimization")

#write onto JSON file
write.("./results/cook_county_providentA.json", JSON.json(site_analysis))
println("Successfully printed results on JSON file")

# Populate the DataFrame with the results produced and inputs
df = DataFrame(
    City = cities,
    PV_size = [round(site_analysis[i][2]["PV"]["size_kw"], digits=0) for i in sites_iter],
    PV_year1_production = [round(site_analysis[i][2]["PV"]["year_one_energy_produced_kwh"], digits=0) for i in sites_iter],
    PV_annual_energy_production_avg = [round(site_analysis[i][2]["PV"]["annual_energy_produced_kwh"], digits=0) for i in sites_iter],
    PV_energy_lcoe = [round(site_analysis[i][2]["PV"]["lcoe_per_kwh"], digits=0) for i in sites_iter],
    PV_energy_exported = [round(site_analysis[i][2]["PV"]["annual_energy_exported_kwh"], digits=0) for i in sites_iter],
    PV_energy_curtailed = [sum(site_analysis[i][2]["PV"]["electric_curtailed_series_kw"]) for i in sites_iter],
    PV_energy_to_Battery_year1 = [sum(site_analysis[i][2]["PV"]["electric_to_storage_series_kw"]) for i in sites_iter],
    Battery_size_kW = [round(site_analysis[i][2]["ElectricStorage"]["size_kw"], digits=0) for i in sites_iter], 
    Battery_size_kWh = [round(site_analysis[i][2]["ElectricStorage"]["size_kwh"], digits=0) for i in sites_iter], 
    Battery_serve_electric_load = [sum(site_analysis[i][2]["ElectricStorage"]["storage_to_load_series_kw"], digits=0) for i in sites_iter], 
    Grid_Electricity_Supplied_kWh_annual = [round(site_analysis[i][2]["ElectricUtility"]["annual_energy_supplied_kwh"], digits=0) for i in sites_iter],
    BAU_Existing_Boiler_Fuel_Consump_MMBtu = [round(site_analysis[i][2]["ExistingBoiler"]["annual_fuel_consumption_mmbtu_bau"], digits=0) for i in sites_iter],
    BAU_Existing_Boiler_Thermal_Prod_MMBtu = [round(site_analysis[i][2]["ExistingBoiler"]["annual_thermal_production_mmbtu_bau"], digits=0) for i in sites_iter],
    NG_Annual_Consumption_MMBtu = [round(site_analysis[i][2]["ExistingBoiler"]["annual_fuel_consumption_mmbtu"], digits=0) for i in sites_iter],
    Total_Annual_Emissions_CO2 = [round(site_analysis[i][2]["Site"]["annual_emissions_tonnes_CO2"], digits=4) for i in sites_iter],
    ElecUtility_Annual_Emissions_CO2 = [round(site_analysis[i][2]["ElectricUtility"]["annual_emissions_tonnes_CO2"], digits=4) for i in sites_iter],
    BAU_Total_Annual_Emissions_CO2 = [round(site_analysis[i][2]["Site"]["annual_emissions_tonnes_CO2_bau"], digits=4) for i in sites_iter],
    LifeCycle_Emissions_CO2 = [round(site_analysis[i][2]["Site"]["lifecycle_emissions_tonnes_CO2"], digits=2) for i in sites_iter],
    BAU_LifeCycle_Emissions_CO2 = [round(site_analysis[i][2]["Site"]["lifecycle_emissions_tonnes_CO2_bau"], digits=2) for i in sites_iter],
    NG_LifeCycle_Emissions_CO2 = [round(site_analysis[i][2]["Site"]["lifecycle_emissions_from_fuelburn_tonnes_CO2"], digits=2) for i in sites_iter],
    Emissions_from_NG = [round(site_analysis[i][2]["Site"]["annual_emissions_from_fuelburn_tonnes_CO2"], digits=0) for i in sites_iter],
    LifeCycle_Emission_Reduction_Fraction = [round(site_analysis[i][2]["Site"]["lifecycle_emissions_reduction_CO2_fraction"], digits=2) for i in sites_iter],
    npv = [round(site_analysis[i][2]["Financial"]["npv"], digits=2) for i in sites_iter],
    lcc = [round(site_analysis[i][2]["Financial"]["lcc"], digits=2) for i in sites_iter]
    )
println(df)

# Define path to xlsx file
file_storage_location = "results/Cook_County_results.xlsx"

# Check if the Excel file already exists
if isfile(file_storage_location)
    # Open the Excel file in read-write mode
    XLSX.openxlsx(file_storage_location, mode="rw") do xf
        counter = 0
        while true
            sheet_name = "ProvidentA_" * string(counter)
            try
                sheet = xf[sheet_name]
                counter += 1
            catch
                break
            end
        end
        sheet_name = "ProvdientA_" * string(counter)
        # Add new sheet
        XLSX.addsheet!(xf, sheet_name)
        # Write DataFrame to the new sheet
        XLSX.writetable!(xf[sheet_name], df)
    end
else
    # Write DataFrame to a new Excel file
    XLSX.writetable!(file_storage_location, df)
end

println("Successful write into XLSX file: $file_storage_location")

"""
=======================================================================================================================================================================
Provident Part B
PV+Battery+ASHP
"""
# Setup inputs Provident part a
data_file = "ProvidentB.JSON" 
input_data = JSON.parsefile("scenarios/$data_file")

markham_rates = "C:/Users/dbernal/OneDrive - NREL/General - Cook County C2C/Internal - REopt Analysis/Custom Rates/Internal Sites/Cook County Internal Provident New.json"
markham_rates_1 = JSON.parsefile(markham_rates)

cermak_electric_load = "C:/Users/dbernal/OneDrive - NREL/General - Cook County C2C/Internal - REopt Analysis/REopt Loads/Load_profile_electric_Provident_hourly.csv"


#the lat/long will be representative of the regions (MW, NE, S, W)
#cities chosen are Chicago, Boston, Houston, San Francisco
cities = ["Chicago", "Chicago", "Chicago", "Chicago", "Chicago", "Chicago", "Chicago", "Chicago", "Chicago", "Chicago"]
lat = [41.834, 41.834, 41.834, 41.834, 41.834, 41.834, 41.834, 41.834, 41.834, 41.834]
long = [-88.044, -88.044, -88.044, -88.044, -88.044, -88.044, -88.044, -88.044, -88.044, -88.044]

#emissions reduction fraction every scenario by 10%
emissions_reduction_min = [0.10, 0.20, 0.30, 0.40, 0.50, 0.60, 0.70, 0.80, 0.90, 1.00]

Site = [raw"Markham-10%", raw"Markham-20%", raw"Markham-30%", raw"Markham-40%", raw"Markham-50%", raw"Markham-60%", raw"Markham-70%", raw"Markham-80%", raw"Markham-90%", raw"Markham-100%"]

site_analysis = []

sites_iter = eachindex(lat)
for i in sites_iter
    input_data_site = copy(input_data)
    # Site Specific
    input_data_site["Site"]["latitude"] = lat[i]
    input_data_site["Site"]["longitude"] = long[i]
    input_data_site["ElectricLoad"]["path_to_csv"] = cermak_electric_load
    input_data_site["ElectricTariff"]["urdb_response"] = cermak_rates_1
    
    #emissions reduction min 
    input_data_site["Site"]["CO2_emissions_reduction_min_fraction"] = emissions_reduction_min[i]
                
    s = Scenario(input_data_site)
    inputs = REoptInputs(s)

     # HiGHS solver
     m1 = Model(optimizer_with_attributes(HiGHS.Optimizer, 
     "time_limit" => 450.0,
     "mip_rel_gap" => 0.01,
     "output_flag" => false, 
     "log_to_console" => false)
     )

    m2 = Model(optimizer_with_attributes(HiGHS.Optimizer, 
     "time_limit" => 450.0,
     "mip_rel_gap" => 0.01,
     "output_flag" => false, 
     "log_to_console" => false)
     )            

    results = run_reopt([m1,m2], inputs)
    append!(site_analysis, [(input_data_site, results)])

    sleep(180)
end
println("Completed optimization")

#write onto JSON file
write.("./results/cook_county_provientB.json", JSON.json(site_analysis))
println("Successfully printed results on JSON file")

# Populate the DataFrame with the results produced and inputs
df = DataFrame(
    City = cities,
    PV_size = [round(site_analysis[i][2]["PV"]["size_kw"], digits=0) for i in sites_iter],
    PV_year1_production = [round(site_analysis[i][2]["PV"]["year_one_energy_produced_kwh"], digits=0) for i in sites_iter],
    PV_annual_energy_production_avg = [round(site_analysis[i][2]["PV"]["annual_energy_produced_kwh"], digits=0) for i in sites_iter],
    PV_energy_lcoe = [round(site_analysis[i][2]["PV"]["lcoe_per_kwh"], digits=0) for i in sites_iter],
    PV_energy_exported = [round(site_analysis[i][2]["PV"]["annual_energy_exported_kwh"], digits=0) for i in sites_iter],
    PV_energy_curtailed = [sum(site_analysis[i][2]["PV"]["electric_curtailed_series_kw"]) for i in sites_iter],
    PV_energy_to_Battery_year1 = [sum(site_analysis[i][2]["PV"]["electric_to_storage_series_kw"]) for i in sites_iter],
    Battery_size_kW = [round(site_analysis[i][2]["ElectricStorage"]["size_kw"], digits=0) for i in sites_iter], 
    Battery_size_kWh = [round(site_analysis[i][2]["ElectricStorage"]["size_kwh"], digits=0) for i in sites_iter], 
    Battery_serve_electric_load = [sum(site_analysis[i][2]["ElectricStorage"]["storage_to_load_series_kw"], digits=0) for i in sites_iter], 
    ASHP_size_tonhour = [round(site_analysis[i][2]["ASHP"]["size_ton"], digits=0) for i in sites_iter],
    ASHP_annual_electric_consumption_kwh = [round(site_analysis[i][2]["ASHP"]["annual_electric_consumption_kwh"], digits=0) for i in sites_iter],
    ASHP_annual_thermal_production_mmbtu = [round(site_analysis[i][2]["ASHP"]["annual_thermal_production_mmbtu"], digits=0) for i in sites_iter],
    ASHP_annual_cooling_tonhour = [round(site_analysis[i][2]["ASHP"]["annual_thermal_production_tonhour"], digits=0) for i in sites_iter],
    Grid_Electricity_Supplied_kWh_annual = [round(site_analysis[i][2]["ElectricUtility"]["annual_energy_supplied_kwh"], digits=0) for i in sites_iter],
    Annual_Total_HeatingLoad_MMBtu = [round(site_analysis[i][2]["HeatingLoad"]["annual_calculated_total_heating_thermal_load_mmbtu"], digits=0) for i in sites_iter],
    BAU_Existing_Boiler_Fuel_Consump_MMBtu = [round(site_analysis[i][2]["ExistingBoiler"]["annual_fuel_consumption_mmbtu_bau"], digits=0) for i in sites_iter],
    BAU_Existing_Boiler_Thermal_Prod_MMBtu = [round(site_analysis[i][2]["ExistingBoiler"]["annual_thermal_production_mmbtu_bau"], digits=0) for i in sites_iter],
    NG_Annual_Consumption_MMBtu = [round(site_analysis[i][2]["ExistingBoiler"]["annual_fuel_consumption_mmbtu"], digits=0) for i in sites_iter],
    Total_Annual_Emissions_CO2 = [round(site_analysis[i][2]["Site"]["annual_emissions_tonnes_CO2"], digits=4) for i in sites_iter],
    ElecUtility_Annual_Emissions_CO2 = [round(site_analysis[i][2]["ElectricUtility"]["annual_emissions_tonnes_CO2"], digits=4) for i in sites_iter],
    BAU_Total_Annual_Emissions_CO2 = [round(site_analysis[i][2]["Site"]["annual_emissions_tonnes_CO2_bau"], digits=4) for i in sites_iter],
    LifeCycle_Emissions_CO2 = [round(site_analysis[i][2]["Site"]["lifecycle_emissions_tonnes_CO2"], digits=2) for i in sites_iter],
    BAU_LifeCycle_Emissions_CO2 = [round(site_analysis[i][2]["Site"]["lifecycle_emissions_tonnes_CO2_bau"], digits=2) for i in sites_iter],
    NG_LifeCycle_Emissions_CO2 = [round(site_analysis[i][2]["Site"]["lifecycle_emissions_from_fuelburn_tonnes_CO2"], digits=2) for i in sites_iter],
    Emissions_from_NG = [round(site_analysis[i][2]["Site"]["annual_emissions_from_fuelburn_tonnes_CO2"], digits=0) for i in sites_iter],
    LifeCycle_Emission_Reduction_Fraction = [round(site_analysis[i][2]["Site"]["lifecycle_emissions_reduction_CO2_fraction"], digits=2) for i in sites_iter],
    npv = [round(site_analysis[i][2]["Financial"]["npv"], digits=2) for i in sites_iter],
    lcc = [round(site_analysis[i][2]["Financial"]["lcc"], digits=2) for i in sites_iter]
    )
println(df)

# Define path to xlsx file
file_storage_location = "results/Cook_County_results.xlsx"

# Check if the Excel file already exists
if isfile(file_storage_location)
    # Open the Excel file in read-write mode
    XLSX.openxlsx(file_storage_location, mode="rw") do xf
        counter = 0
        while true
            sheet_name = "ProvidentB_" * string(counter)
            try
                sheet = xf[sheet_name]
                counter += 1
            catch
                break
            end
        end
        sheet_name = "ProvidentB_" * string(counter)
        # Add new sheet
        XLSX.addsheet!(xf, sheet_name)
        # Write DataFrame to the new sheet
        XLSX.writetable!(xf[sheet_name], df)
    end
else
    # Write DataFrame to a new Excel file
    XLSX.writetable!(file_storage_location, df)
end

println("Successful write into XLSX file: $file_storage_location")