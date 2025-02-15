
using REopt
using HiGHS
using JSON
using JuMP
using CSV 
using DataFrames #to construct comparison
using XLSX 
using DelimitedFiles
using Plots
using Dates

#ENV["NREL_DEVELOPER_API_KEY"]="gAXbkyLjfTFEFfiO3YhkxxJ6rkufRaSktk40ho4x"
#defining functions necessary for code to run

# Function to safely extract values from JSON with default value if key is missing
function safe_get(data::Dict{String, Any}, keys::Vector{String}, default=0)
    try
        for k in keys
            data = data[k]
        end
        return data
    catch e
        if e isa KeyError
            return default
        else
            rethrow(e)
        end
    end
end

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

# In the current directory, create a folder called InputData
    # If using a custom load, critical load, and/or utility tariff, place the following .csv files in the InputData folder: load data, critical load data, utility_tariff

# Note: The ERP_run function likely does not need to be modified, so you can minimize the "ERP_run" function (lines 27 - 182) by clicking the arrow to the left of the word "function"
# Use this function for running the ERP analysis 
function ERP_run(; REopt_results = "", REopt_post_inputs = "", post = "", maximumoutageduration = "") 
    if "Generator" in keys(REopt_results)
        num_gen = 1
        each_gen_kw = round(REopt_results["Generator"]["size_kw"] / num_gen; digits=0)
        FuelAvailable = REopt_post_inputs["Generator"]["fuel_avail_gal"]
    else
        each_gen_kw = 0
        num_gen = 0
        FuelAvailable = 0
    end 

    if "ElectricStorage" in keys(REopt_results)
        if REopt_results["ElectricStorage"]["size_kw"] > 0
            Batterykw = REopt_results["ElectricStorage"]["size_kw"]
        else
            Batterykw = 0
        end 

        if REopt_results["ElectricStorage"]["size_kwh"] > 0
            Batterykwh = REopt_results["ElectricStorage"]["size_kwh"]
            BatterySOC = REopt_results["ElectricStorage"]["soc_series_fraction"]
        else
            Batterykwh = 0
            BatterySOC = zeros(35040)
        end 
    else 
        BatterySOC = zeros(35040)
        Batterykw = 0
        Batterykwh = 0
    end

    #print("\n the battery kw input into the ERP is: "*string(Batterykw))
    #print("\n the battery kwh input into the ERP is: "*string(Batterykwh))
    #print("\n the average battery soc fraction input into the ERP is: "*string(mean(BatterySOC)))

    # Aggregate the PV generation, if there are multiple PVs used
    #=
    total_pv_kw = 0
    total_pv_kwh = [0 for i in range(1,length(REopt_results["ElectricLoad"]["load_series_kw"]))]
    
    # For a multiple PV system:
    
    # Create an array with the PV names from the post file
    PV_names_in_the_postfiles = ["1_AngledRoof", "3_FlatRoof","5_Carport","6_Carport","7_ParkRoof"] # ["Flatroof1","Flatroof2","Carport1","Carport2","BasketballCourtRoof"]
    
    print("\n  Using ERP with these PV system options:")
    print(PV_names_in_the_postfiles)
    print("\n")
    
    for pv in PV_names_in_the_postfiles
            total_pv_kw += REopt_results[pv]["size_kw"]
            total_pv_kwh += 0.25 * (REopt_results[pv]["electric_to_storage_series_kw"] + REopt_results[pv]["electric_curtailed_series_kw"] + 
                REopt_results[pv]["electric_to_load_series_kw"] + REopt_results[pv]["electric_to_grid_series_kw"]
            )
    end 
    pv_production_factor_series = total_pv_kwh / total_pv_kw
    print("\n   The total PV kW used in the ERP analysis is (kW): "*string(total_pv_kw))
    print("\n   The maximum kWh output, over the time step length, from the combined PV system is (kWh): "*string(maximum(total_pv_kwh)))
    =#

    # Treat it as a single PV system because all of the PV results are combined into REopt_results["PV"] earlier in the code
    # For a single PV system
    total_pv_kw = REopt_results["PV"]["size_kw"]
    total_pv_kwh = (REopt_results["PV"]["electric_to_storage_series_kw"] + REopt_results["PV"]["electric_curtailed_series_kw"] + 
                    REopt_results["PV"]["electric_to_load_series_kw"] + REopt_results["PV"]["electric_to_grid_series_kw"]
                    )
    pv_production_factor_series = total_pv_kwh / total_pv_kw
    

    reliability_inputs = Dict(
        "critical_loads_kw" => REopt_results["ElectricLoad"]["critical_load_series_kw"],
        "pv_size_kw" => total_pv_kw,
        "pv_production_factor_series"=> pv_production_factor_series,
        "battery_size_kw" => Batterykw,
        "battery_size_kwh" => Batterykwh,
        "battery_starting_soc_series_fraction" => BatterySOC,
        "generator_size_kw" => each_gen_kw,  # the size of each generator
        "fuel_limit" => FuelAvailable, # Can specify per generator. # Amount of fuel available, either by generator type or per generator, depending on fuel_limit_is_per_generator. Change generator_fuel_burn_rate_per_kwh for different fuel efficiencies. Fuel units should be consistent with generator_fuel_intercept_per_hr and generator_fuel_burn_rate_per_kwh
        "fuel_limit_is_per_generator" => false,
    
        "max_outage_duration" => maximumoutageduration, # maximum outage duration in timesteps
        "num_generators" => num_gen, # number of generators
        "generator_operational_availability" => 0.995,  # Fraction of year generators not down for maintenance
        "generator_failure_to_start" => 0.0094,  # Chance of generator starting given outage
        "generator_mean_time_to_failure" => 1100, # Average number of time steps between a generator's failures. 1/(failure to run probability) 
        "battery_operational_availability" => 0.97, # 97% is the default 
        "num_battery_bins" => 101, #  Number of bins for discretely modeling battery state of charge
        "battery_charge_efficiency" => 0.948, # default is 0.948
        "battery_discharge_efficiency" => 0.948, # default is 0.948
        "battery_minimum_soc_fraction" => 0, # The minimum battery state of charge (represented as a fraction) allowed during outages
        "microgrid_only" => false,  # Boolean to specify if only microgrid upgraded technologies run during grid outage
    
        "pv_operational_availability" => 0.98, #0.98 is the default
        "wind_operational_availability" => 0.97, # 0.97 is the default
    
        "fuel_limit_is_per_generator" => false,  # Boolean to determine whether fuel limit is given per generator or per generator type
        "generator_fuel_intercept_per_hr" => 0.0, # Amount of fuel burned each time step while idling. Fuel units should be consistent with fuel_limit and generator_fuel_burn_rate_per_kwh
        "generator_fuel_burn_rate_per_kwh" => 0.076, # Amount of fuel used per kWh generated. Fuel units should be consistent with fuel_limit and generator_fuel_intercept_per_hr
                
        )

        # Notes
            # backup_reliability can be called in several different ways - including just running it with reliability_inputs and no REopt results
            # Input num_generators must be the same length as generator_size_kw or a scalar if generator_size_kw not provided. 
            # Backup_reliability will take the total size of the generator from the REopt results, and divide that by the number of generators in the r dictionary, to get the size of each generator

        #reliability_results = backup_reliability(REopt_results, REopt_post_inputs, reliability_inputs) 

        # Note: run backup reliability with just the "reliability_inputs" dictionary in order to be able to define the inputs more specifically (like a PV generation profile, if need to combine all of them in the case of multiple PVs)
        reliability_results = backup_reliability(reliability_inputs) 


        TimeStepsPerHour = post["Settings"]["time_steps_per_hour"]
        Hours = collect(1:reliability_inputs["max_outage_duration"]) / TimeStepsPerHour
        
        if TimeStepsPerHour == 1
            HoursInYear = collect(1:8760)
            TimeStepsPerDay = 24
        elseif TimeStepsPerHour == 4
            HoursInYear = collect(1:35040)
            TimeStepsPerDay = 24*4            
        end
        """
        # Compare resilience predictions with and without reliability considered
        Plots.plot(Hours, 100*reliability_results["mean_cumulative_survival_by_duration"], label = "ERP results")
        Plots.plot!(Hours, 100* reliability_results["mean_fuel_survival_by_duration"], label = "Reliability not considered")

        Plots.xlabel!("Length of Outage (Hours)")
        Plots.ylabel!("Probability of Survival (Percent)")
        display(Plots.title!("Outage Survival Probability by Outage Duration"))

        # Plot the probability of surviving the maximum outage length (defined in the reliability_inputs dictionary) at each time of the year
        DayInYear = HoursInYear/TimeStepsPerDay
        Plots.plot(DayInYear, 100*reliability_results["fuel_survival_final_time_step"], label = "Without Reliability (fuel_outage_survival_final_time_step)")
        Plots.plot!(DayInYear, 100*reliability_results["cumulative_survival_final_time_step"], label = "With reliability (cumulative_outage_survival_final_time_step)")
        Plots.xlims!(180,187) 
        Plots.xlabel!("Day of the year")
        Plots.ylabel!("Probability (Percent)")
        display(Plots.title!("Probability of Surviving the maximum outage length, at each time step"))

        # Generate a plot with the box plot data
        Plots.plot(reliability_results["monthly_min_cumulative_survival_final_time_step"], label = "Minimum")
        Plots.plot!(reliability_results["monthly_lower_quartile_cumulative_survival_final_time_step"], label = "Lower Quartile")
        Plots.plot!(reliability_results["monthly_median_cumulative_survival_final_time_step"], label = "Median")
        Plots.plot!(reliability_results["monthly_upper_quartile_cumulative_survival_final_time_step"], label = "Upper Quartile")
        Plots.plot!(reliability_results["monthly_max_cumulative_survival_final_time_step"], label = "Maximum")
        Plots.xlabel!("Month Number")
        Plots.ylabel!("Probability")
        display(Plots.title!("Data for survival of the maximum outage length, by month")) """
        fig_saving = [1, 2, 3]
        for i in eachindex(fig_saving)
            # Generate plots
            p1 = plot(Hours, 100 * reliability_results["mean_cumulative_survival_by_duration"], label="ERP results")
            plot!(p1, Hours, 100 * reliability_results["mean_fuel_survival_by_duration"], label="Reliability not considered")
            xlabel!(p1, "Length of Outage (Hours)")
            ylabel!(p1, "Probability of Survival (Percent)")
            title!(p1, "Outage Survival Probability by Outage Duration")
            savefig(p1, "./results/markham_outage_survival_probability_by_duration_$i.png")
    
            DayInYear = HoursInYear / TimeStepsPerDay
            p2 = plot(DayInYear, 100 * reliability_results["fuel_survival_final_time_step"], label="Without Reliability (fuel_outage_survival_final_time_step)")
            plot!(p2, DayInYear, 100 * reliability_results["cumulative_survival_final_time_step"], label="With reliability (cumulative_outage_survival_final_time_step)")
            xlims!(p2, 180, 187)
            xlabel!(p2, "Day of the year")
            ylabel!(p2, "Probability (Percent)")
            title!(p2, "Probability of Surviving the maximum outage length, at each time step")
            savefig(p2, "./results/markham_probability_of_surviving_max_outage_by_time_step_$i.png")
    
            p3 = plot(reliability_results["monthly_min_cumulative_survival_final_time_step"], label="Minimum")
            plot!(p3, reliability_results["monthly_lower_quartile_cumulative_survival_final_time_step"], label="Lower Quartile")
            plot!(p3, reliability_results["monthly_median_cumulative_survival_final_time_step"], label="Median")
            plot!(p3, reliability_results["monthly_upper_quartile_cumulative_survival_final_time_step"], label="Upper Quartile")
            plot!(p3, reliability_results["monthly_max_cumulative_survival_final_time_step"], label="Maximum")
            xlabel!(p3, "Month Number")
            ylabel!(p3, "Probability")
            title!(p3, "Data for survival of the maximum outage length, by month")
            savefig(p3, "./results/markham_survival_of_max_outage_by_month_$i.png")
        end
        ProbabilityOfSurvivingMaximumOutageLength_percent =  100*reliability_results["mean_cumulative_survival_final_time_step"]
        #print("\n The probability of survival, considering component reliability, for the maximum outage length is: "*string(ProbabilityOfSurvivingMaximumOutageLength_percent)*"%")
        #print("\n ")

        return reliability_results, ProbabilityOfSurvivingMaximumOutageLength_percent
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
cities = ["Chicago", "Chicago", "Chicago"]
lat = [ 41.834, 41.834, 41.834]
long = [-88.044, -88.044, -88.044]

#Column for inputs
column_inputs = ["inputs", "inputs", "inputs"]

#emissions reduction fraction every scenario by 10%
emissions_reduction_min = [0.10, 0.20, 0.30, 0.40, 0.50, 0.60, 0.70, 0.80, 0.90, 1.00]
#hours of outage to sustain
outage_minimum_sustain = [8, 16, 24] #input_data_site["Site"]["min_resil_time_steps"] = outage_minimum_sustain[i]
outage_durations = [8, 16, 24] #"ElectricUtility""outage_duration"

site_analysis = []
ERP_results = [] #to store resilience results 

sites_iter = eachindex(lat)
for i in sites_iter
    input_data_site = copy(input_data)
    # Site Specific
    input_data_site["Site"]["latitude"] = lat[i]
    input_data_site["Site"]["longitude"] = long[i]
    input_data_site["Site"]["min_resil_time_steps"] = outage_minimum_sustain[i]
    input_data_site["ElectricLoad"]["loads_kw"] = cermak_loads_kw
    input_data_site["ElectricTariff"]["urdb_response"] = cermak_rates_1
    input_data_site["ElectricUtility"]["outage_durations"] = [outage_durations[i]]

    #emissions reduction min 
    #input_data_site["Site"]["CO2_emissions_reduction_min_fraction"] = emissions_reduction_min[i]
                
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

    #sleep(60)
    OutageSurvival = zeros(sites_iter)
    AllResilienceResults, OutageSurvival[i] = ERP_run(REopt_results = results, REopt_post_inputs = inputs, post = input_data_site, maximumoutageduration = outage_durations[i])
    append!(ERP_results, AllResilienceResults) # Append the resilience results
    
end
println("Completed optimization")


#write onto JSON file
write.("./results/cook_county_cermakA.json", JSON.json(site_analysis))
println("Successfully printed results on JSON file")
write.("./results/cook_county_cermakA_ERP.json", JSON.json(ERP_results))
println("Successfuly printed ERP results onto JSON file")

# Populate the DataFrame with the results produced and inputs
df = DataFrame(
    City = cities,
    PV_size = [round(safe_get(site_analysis[i][2], ["PV", "size_kw"]), digits=0) for i in sites_iter],
    PV_year1_production = [round(safe_get(site_analysis[i][2], ["PV", "year_one_energy_produced_kwh"]), digits=0) for i in sites_iter],
    PV_annual_energy_production_avg = [round(safe_get(site_analysis[i][2], ["PV", "annual_energy_produced_kwh"]), digits=0) for i in sites_iter],
    PV_energy_lcoe = [round(safe_get(site_analysis[i][2], ["PV", "lcoe_per_kwh"]), digits=0) for i in sites_iter],
    PV_energy_exported = [round(safe_get(site_analysis[i][2], ["PV", "annual_energy_exported_kwh"]), digits=0) for i in sites_iter],
    PV_energy_curtailed = [sum(safe_get(site_analysis[i][2], ["PV", "electric_curtailed_series_kw"], 0)) for i in sites_iter],
    PV_energy_to_Battery_year1 = [sum(safe_get(site_analysis[i][2], ["PV", "electric_to_storage_series_kw"], 0)) for i in sites_iter],
    Battery_size_kW = [round(safe_get(site_analysis[i][2], ["ElectricStorage", "size_kw"]), digits=0) for i in sites_iter], 
    Battery_size_kWh = [round(safe_get(site_analysis[i][2], ["ElectricStorage", "size_kwh"]), digits=0) for i in sites_iter], 
    Battery_serve_electric_load = [sum(safe_get(site_analysis[i][2], ["ElectricStorage", "storage_to_load_series_kw"], 0)) for i in sites_iter], 
    Grid_Electricity_Supplied_kWh_annual = [round(safe_get(site_analysis[i][2], ["ElectricUtility", "annual_energy_supplied_kwh"]), digits=0) for i in sites_iter],
    Total_Annual_Emissions_CO2 = [round(safe_get(site_analysis[i][2], ["Site", "annual_emissions_tonnes_CO2"]), digits=4) for i in sites_iter],
    ElecUtility_Annual_Emissions_CO2 = [round(safe_get(site_analysis[i][2], ["ElectricUtility", "annual_emissions_tonnes_CO2"]), digits=4) for i in sites_iter],
    BAU_Total_Annual_Emissions_CO2 = [round(safe_get(site_analysis[i][2], ["Site", "annual_emissions_tonnes_CO2_bau"]), digits=4) for i in sites_iter],
    LifeCycle_Emissions_CO2 = [round(safe_get(site_analysis[i][2], ["Site", "lifecycle_emissions_tonnes_CO2"]), digits=2) for i in sites_iter],
    BAU_LifeCycle_Emissions_CO2 = [round(safe_get(site_analysis[i][2], ["Site", "lifecycle_emissions_tonnes_CO2_bau"]), digits=2) for i in sites_iter],
    LifeCycle_Emission_Reduction_Fraction = [round(safe_get(site_analysis[i][2], ["Site", "lifecycle_emissions_reduction_CO2_fraction"]), digits=2) for i in sites_iter],
    LifeCycle_capex_costs_for_generation_techs = [round(safe_get(site_analysis[i][2], ["Financial", "lifecycle_generation_tech_capital_costs"]), digits=2) for i in sites_iter],
    LifeCycle_capex_costs_for_battery = [round(safe_get(site_analysis[i][2], ["Financial", "lifecycle_storage_tech_capital_costs"]), digits=2) for i in sites_iter],
    Initial_upfront_capex_wo_incentives = [round(safe_get(site_analysis[i][2], ["Financial", "initial_capital_costs"]), digits=2) for i in sites_iter],
    Initial_upfront_capex_w_incentives = [round(safe_get(site_analysis[i][2], ["Financial", "initial_capital_costs_after_incentives"]), digits=2) for i in sites_iter],
    Yr1_energy_cost_after_tax = [round(safe_get(site_analysis[i][2], ["ElectricTariff", "year_one_energy_cost_before_tax"]), digits=2) for i in sites_iter],
    Yr1_demand_cost_after_tax = [round(safe_get(site_analysis[i][2], ["ElectricTariff", "year_one_demand_cost_before_tax"]), digits=2) for i in sites_iter],
    Yr1_total_energy_bill_before_tax = [round(safe_get(site_analysis[i][2], ["ElectricTariff", "year_one_bill_before_tax"]), digits=2) for i in sites_iter],
    Yr1_export_benefit_before_tax = [round(safe_get(site_analysis[i][2], ["ElectricTariff", "year_one_export_benefit_before_tax"]), digits=2) for i in sites_iter],
    Annual_renewable_electricity_kwh = [round(safe_get(site_analysis[i][2], ["Site", "annual_renewable_electricity_kwh"]), digits=2) for i in sites_iter],
    Annual_renewable_electricity_kwh_fraction = [round(safe_get(site_analysis[i][2], ["Site", "renewable_electricity_fraction"]), digits=2) for i in sites_iter],
    npv = [round(safe_get(site_analysis[i][2], ["Financial", "npv"]), digits=2) for i in sites_iter],
    lcc = [round(safe_get(site_analysis[i][2], ["Financial", "lcc"]), digits=2) for i in sites_iter]
    )
println(df)

# Define path to xlsx file
file_storage_location = "C:/Users/dbernal/Documents/GitHub/REopt_related_code/Cook County Project-C2C/results/Cook_County_results.xlsx"

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
Markham Part A
PV+Battery
"""

# Setup inputs Markham part a
data_file = "MarkhamA.JSON" 
input_data = JSON.parsefile("scenarios/$data_file")

markham_rates = "C:/Users/dbernal/OneDrive - NREL/General - Cook County C2C/Internal - REopt Analysis/Custom Rates/Internal Sites/Cook County Internal Markham New.json"
markham_rates_1 = JSON.parsefile(markham_rates)

markham_electric_load = "C:/Users/dbernal/OneDrive - NREL/General - Cook County C2C/Internal - REopt Analysis/REopt Loads/Load_profile_electric_Markham.csv"

# Read the CSV file
markham_loads_kw = read_csv_without_bom(markham_electric_load)

# Convert matrix to a one-dimensional array 
markham_loads_kw = reshape(markham_loads_kw, :)  # This flattens the matrix into a one-dimensional array
markham_loads_kw = markham_loads_kw[8761:17520] #take off the hours and leave the loads

println("Correctly obtained data_file")

#the lat/long will be representative of the city, Chicago
cities = ["Chicago", "Chicago", "Chicago"]
lat = [ 41.834, 41.834, 41.834]
long = [-88.044, -88.044, -88.044]

#emissions reduction fraction every scenario by 10%
#emissions_reduction_min = [0.10, 0.20, 0.30, 0.40, 0.50, 0.60, 0.70, 0.80, 0.90, 1.00]
#hours of outage to sustain
outage_minimum_sustain = [8, 16, 24] #input_data_site["Site"]["min_resil_time_steps"] = outage_minimum_sustain[i]
outage_durations = [8, 16, 24] #"ElectricUtility""outage_duration"

#Existing PV at Markham
markham_existing_pv = [743, 743, 743]

site_analysis = []
ERP_results = [] #to store resilience results 

sites_iter = eachindex(lat)
for i in sites_iter
    input_data_site = copy(input_data)
    # Site Specific
    input_data_site["Site"]["latitude"] = lat[i]
    input_data_site["Site"]["longitude"] = long[i]
    input_data_site["Site"]["min_resil_time_steps"] = outage_minimum_sustain[i]
    input_data_site["ElectricLoad"]["loads_kw"] = markham_loads_kw
    input_data_site["ElectricTariff"]["urdb_response"] = markham_rates_1
    input_data_site["ElectricUtility"]["outage_durations"] = [outage_durations[i]]

    
    #existing PV on Markham
    input_data_site["PV"]["existing_kw"] = markham_existing_pv[i]

    #emissions reduction min 
    input_data_site["Site"]["CO2_emissions_reduction_min_fraction"] = emissions_reduction_min[i]
                
    s = Scenario(input_data_site)
    inputs = REoptInputs(s)

     # HiGHS solver
     m1 = Model(optimizer_with_attributes(HiGHS.Optimizer, 
     "time_limit" => 600.0,
     "mip_rel_gap" => 0.02,
     "output_flag" => false, 
     "log_to_console" => false)
     )

    m2 = Model(optimizer_with_attributes(HiGHS.Optimizer, 
     "time_limit" => 600.0,
     "mip_rel_gap" => 0.02,
     "output_flag" => false, 
     "log_to_console" => false)
     )            

    results = run_reopt([m1,m2], inputs)
    append!(site_analysis, [(input_data_site, results)])

    #sleep(60)
    OutageSurvival = zeros(sites_iter)
    AllResilienceResults, OutageSurvival[i] = ERP_run(REopt_results = results, REopt_post_inputs = inputs, post = input_data_site, maximumoutageduration = outage_durations[i])
    append!(ERP_results, AllResilienceResults)

end
println("Completed optimization")

#write onto JSON file
write.("./results/cook_county_markhamA.json", JSON.json(site_analysis))
println("Successfully printed results on JSON file")
write.("./results/cook_county_markhamA_ERP.json", JSON.json(ERP_results))
println("Successfuly printed ERP results onto JSON file")

# Populate the DataFrame with the results produced and inputs
df = DataFrame(
    City = cities,
    PV_size = [round(safe_get(site_analysis[i][2], ["PV", "size_kw"]), digits=0) for i in sites_iter],
    PV_year1_production = [round(safe_get(site_analysis[i][2], ["PV", "year_one_energy_produced_kwh"]), digits=0) for i in sites_iter],
    PV_annual_energy_production_avg = [round(safe_get(site_analysis[i][2], ["PV", "annual_energy_produced_kwh"]), digits=0) for i in sites_iter],
    PV_energy_lcoe = [round(safe_get(site_analysis[i][2], ["PV", "lcoe_per_kwh"]), digits=0) for i in sites_iter],
    PV_energy_exported = [round(safe_get(site_analysis[i][2], ["PV", "annual_energy_exported_kwh"]), digits=0) for i in sites_iter],
    PV_energy_curtailed = [sum(safe_get(site_analysis[i][2], ["PV", "electric_curtailed_series_kw"], 0)) for i in sites_iter],
    PV_energy_to_Battery_year1 = [sum(safe_get(site_analysis[i][2], ["PV", "electric_to_storage_series_kw"], 0)) for i in sites_iter],
    Battery_size_kW = [round(safe_get(site_analysis[i][2], ["ElectricStorage", "size_kw"]), digits=0) for i in sites_iter], 
    Battery_size_kWh = [round(safe_get(site_analysis[i][2], ["ElectricStorage", "size_kwh"]), digits=0) for i in sites_iter], 
    Battery_serve_electric_load = [sum(safe_get(site_analysis[i][2], ["ElectricStorage", "storage_to_load_series_kw"], 0)) for i in sites_iter], 
    Grid_Electricity_Supplied_kWh_annual = [round(safe_get(site_analysis[i][2], ["ElectricUtility", "annual_energy_supplied_kwh"]), digits=0) for i in sites_iter],
    Total_Annual_Emissions_CO2 = [round(safe_get(site_analysis[i][2], ["Site", "annual_emissions_tonnes_CO2"]), digits=4) for i in sites_iter],
    ElecUtility_Annual_Emissions_CO2 = [round(safe_get(site_analysis[i][2], ["ElectricUtility", "annual_emissions_tonnes_CO2"]), digits=4) for i in sites_iter],
    BAU_Total_Annual_Emissions_CO2 = [round(safe_get(site_analysis[i][2], ["Site", "annual_emissions_tonnes_CO2_bau"]), digits=4) for i in sites_iter],
    LifeCycle_Emissions_CO2 = [round(safe_get(site_analysis[i][2], ["Site", "lifecycle_emissions_tonnes_CO2"]), digits=2) for i in sites_iter],
    BAU_LifeCycle_Emissions_CO2 = [round(safe_get(site_analysis[i][2], ["Site", "lifecycle_emissions_tonnes_CO2_bau"]), digits=2) for i in sites_iter],
    LifeCycle_Emission_Reduction_Fraction = [round(safe_get(site_analysis[i][2], ["Site", "lifecycle_emissions_reduction_CO2_fraction"]), digits=2) for i in sites_iter],
    LifeCycle_capex_costs_for_generation_techs = [round(safe_get(site_analysis[i][2], ["Financial", "lifecycle_generation_tech_capital_costs"]), digits=2) for i in sites_iter],
    LifeCycle_capex_costs_for_battery = [round(safe_get(site_analysis[i][2], ["Financial", "lifecycle_storage_tech_capital_costs"]), digits=2) for i in sites_iter],
    Initial_upfront_capex_wo_incentives = [round(safe_get(site_analysis[i][2], ["Financial", "initial_capital_costs"]), digits=2) for i in sites_iter],
    Initial_upfront_capex_w_incentives = [round(safe_get(site_analysis[i][2], ["Financial", "initial_capital_costs_after_incentives"]), digits=2) for i in sites_iter],
    Yr1_energy_cost_after_tax = [round(safe_get(site_analysis[i][2], ["ElectricTariff", "year_one_energy_cost_before_tax"]), digits=2) for i in sites_iter],
    Yr1_demand_cost_after_tax = [round(safe_get(site_analysis[i][2], ["ElectricTariff", "year_one_demand_cost_before_tax"]), digits=2) for i in sites_iter],
    Yr1_total_energy_bill_before_tax = [round(safe_get(site_analysis[i][2], ["ElectricTariff", "year_one_bill_before_tax"]), digits=2) for i in sites_iter],
    Yr1_export_benefit_before_tax = [round(safe_get(site_analysis[i][2], ["ElectricTariff", "year_one_export_benefit_before_tax"]), digits=2) for i in sites_iter],
    Annual_renewable_electricity_kwh = [round(safe_get(site_analysis[i][2], ["Site", "annual_renewable_electricity_kwh"]), digits=2) for i in sites_iter],
    Annual_renewable_electricity_kwh_fraction = [round(safe_get(site_analysis[i][2], ["Site", "renewable_electricity_fraction"]), digits=2) for i in sites_iter],
    npv = [round(safe_get(site_analysis[i][2], ["Financial", "npv"]), digits=2) for i in sites_iter],
    lcc = [round(safe_get(site_analysis[i][2], ["Financial", "lcc"]), digits=2) for i in sites_iter]
    )
println(df)

# Define path to xlsx file
file_storage_location = "./results/Cook_County_results.xlsx"

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
Provident Part A
PV+Battery
"""

# Setup inputs Provident part a
data_file = "ProvidentA.JSON" 
input_data = JSON.parsefile("scenarios/$data_file")

provident_rates = "C:/Users/dbernal/OneDrive - NREL/General - Cook County C2C/Internal - REopt Analysis/Custom Rates/Internal Sites/Cook County Internal Provident New.json"
provident_rates_1 = JSON.parsefile(markham_rates)

provident_electric_load = "C:/Users/dbernal/OneDrive - NREL/General - Cook County C2C/Internal - REopt Analysis/REopt Loads/Load_profile_electric_Provident_hourly.csv"

# Read the CSV file
provident_loads_kw = read_csv_without_bom(provident_electric_load)

# Convert matrix to a one-dimensional array 
provident_loads_kw = reshape(provident_loads_kw, :)  # This flattens the matrix into a one-dimensional array
provident_loads_kw = provident_loads_kw[8761:17520] #take off the hours and leave the loads

println("Correctly obtained data_file")

#the lat/long will be representative of the regions (MW, NE, S, W)
#cities chosen are Chicago, Boston, Houston, San Francisco
cities = ["Chicago", "Chicago", "Chicago"]
lat = [ 41.834, 41.834, 41.834]
long = [ -88.044, -88.044, -88.044]

#emissions reduction fraction every scenario by 10%
#emissions_reduction_min = [0.10, 0.20, 0.30, 0.40, 0.50, 0.60, 0.70, 0.80, 0.90, 1.00]
#hours of outage to sustain
outage_minimum_sustain = [8, 16, 24] #input_data_site["Site"]["min_resil_time_steps"] = outage_minimum_sustain[i]
outage_durations = [8, 16, 24] #"ElectricUtility""outage_duration"

site_analysis = []
ERP_results = [] #to store resilience results 

sites_iter = eachindex(lat)
for i in sites_iter
    input_data_site = copy(input_data)
    # Site Specific
    input_data_site["Site"]["latitude"] = lat[i]
    input_data_site["Site"]["longitude"] = long[i]
    input_data_site["Site"]["min_resil_time_steps"] = outage_minimum_sustain[i]
    input_data_site["ElectricLoad"]["loads_kw"] = provident_loads_kw
    input_data_site["ElectricTariff"]["urdb_response"] = provident_rates_1
    input_data_site["ElectricUtility"]["outage_durations"] = [outage_durations[i]]
    
    #emissions reduction min 
    #input_data_site["Site"]["CO2_emissions_reduction_min_fraction"] = emissions_reduction_min[i]
                
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

    #sleep(60)
    OutageSurvival = zeros(sites_iter)
    AllResilienceResults, OutageSurvival[i] = ERP_run(REopt_results = results, REopt_post_inputs = inputs, post = input_data_site, maximumoutageduration = outage_durations[i])
    append!(ERP_results, AllResilienceResults)
end
println("Completed optimization")

#write onto JSON file
write.("./results/cook_county_providentA.json", JSON.json(site_analysis))
println("Successfully printed results on JSON file")
write.("./results/cook_county_providentA_ERP.json", JSON.json(ERP_results))
println("Successfully printed ProvidentA ERP results on JSON file")

# Populate the DataFrame with the results produced and inputs
df = DataFrame(
    City = cities,
    PV_size = [round(safe_get(site_analysis[i][2], ["PV", "size_kw"]), digits=0) for i in sites_iter],
    PV_year1_production = [round(safe_get(site_analysis[i][2], ["PV", "year_one_energy_produced_kwh"]), digits=0) for i in sites_iter],
    PV_annual_energy_production_avg = [round(safe_get(site_analysis[i][2], ["PV", "annual_energy_produced_kwh"]), digits=0) for i in sites_iter],
    PV_energy_lcoe = [round(safe_get(site_analysis[i][2], ["PV", "lcoe_per_kwh"]), digits=0) for i in sites_iter],
    PV_energy_exported = [round(safe_get(site_analysis[i][2], ["PV", "annual_energy_exported_kwh"]), digits=0) for i in sites_iter],
    PV_energy_curtailed = [sum(safe_get(site_analysis[i][2], ["PV", "electric_curtailed_series_kw"], 0)) for i in sites_iter],
    PV_energy_to_Battery_year1 = [sum(safe_get(site_analysis[i][2], ["PV", "electric_to_storage_series_kw"], 0)) for i in sites_iter],
    Battery_size_kW = [round(safe_get(site_analysis[i][2], ["ElectricStorage", "size_kw"]), digits=0) for i in sites_iter], 
    Battery_size_kWh = [round(safe_get(site_analysis[i][2], ["ElectricStorage", "size_kwh"]), digits=0) for i in sites_iter], 
    Battery_serve_electric_load = [sum(safe_get(site_analysis[i][2], ["ElectricStorage", "storage_to_load_series_kw"], 0)) for i in sites_iter], 
    Grid_Electricity_Supplied_kWh_annual = [round(safe_get(site_analysis[i][2], ["ElectricUtility", "annual_energy_supplied_kwh"]), digits=0) for i in sites_iter],
    Total_Annual_Emissions_CO2 = [round(safe_get(site_analysis[i][2], ["Site", "annual_emissions_tonnes_CO2"]), digits=4) for i in sites_iter],
    ElecUtility_Annual_Emissions_CO2 = [round(safe_get(site_analysis[i][2], ["ElectricUtility", "annual_emissions_tonnes_CO2"]), digits=4) for i in sites_iter],
    BAU_Total_Annual_Emissions_CO2 = [round(safe_get(site_analysis[i][2], ["Site", "annual_emissions_tonnes_CO2_bau"]), digits=4) for i in sites_iter],
    LifeCycle_Emissions_CO2 = [round(safe_get(site_analysis[i][2], ["Site", "lifecycle_emissions_tonnes_CO2"]), digits=2) for i in sites_iter],
    BAU_LifeCycle_Emissions_CO2 = [round(safe_get(site_analysis[i][2], ["Site", "lifecycle_emissions_tonnes_CO2_bau"]), digits=2) for i in sites_iter],
    LifeCycle_Emission_Reduction_Fraction = [round(safe_get(site_analysis[i][2], ["Site", "lifecycle_emissions_reduction_CO2_fraction"]), digits=2) for i in sites_iter],
    LifeCycle_capex_costs_for_generation_techs = [round(safe_get(site_analysis[i][2], ["Financial", "lifecycle_generation_tech_capital_costs"]), digits=2) for i in sites_iter],
    LifeCycle_capex_costs_for_battery = [round(safe_get(site_analysis[i][2], ["Financial", "lifecycle_storage_tech_capital_costs"]), digits=2) for i in sites_iter],
    Initial_upfront_capex_wo_incentives = [round(safe_get(site_analysis[i][2], ["Financial", "initial_capital_costs"]), digits=2) for i in sites_iter],
    Initial_upfront_capex_w_incentives = [round(safe_get(site_analysis[i][2], ["Financial", "initial_capital_costs_after_incentives"]), digits=2) for i in sites_iter],
    Yr1_energy_cost_after_tax = [round(safe_get(site_analysis[i][2], ["ElectricTariff", "year_one_energy_cost_before_tax"]), digits=2) for i in sites_iter],
    Yr1_demand_cost_after_tax = [round(safe_get(site_analysis[i][2], ["ElectricTariff", "year_one_demand_cost_before_tax"]), digits=2) for i in sites_iter],
    Yr1_total_energy_bill_before_tax = [round(safe_get(site_analysis[i][2], ["ElectricTariff", "year_one_bill_before_tax"]), digits=2) for i in sites_iter],
    Yr1_export_benefit_before_tax = [round(safe_get(site_analysis[i][2], ["ElectricTariff", "year_one_export_benefit_before_tax"]), digits=2) for i in sites_iter],
    Annual_renewable_electricity_kwh = [round(safe_get(site_analysis[i][2], ["Site", "annual_renewable_electricity_kwh"]), digits=2) for i in sites_iter],
    Annual_renewable_electricity_kwh_fraction = [round(safe_get(site_analysis[i][2], ["Site", "renewable_electricity_fraction"]), digits=2) for i in sites_iter],
    npv = [round(safe_get(site_analysis[i][2], ["Financial", "npv"]), digits=2) for i in sites_iter],
    lcc = [round(safe_get(site_analysis[i][2], ["Financial", "lcc"]), digits=2) for i in sites_iter]
    )
println(df)

# Define path to xlsx file
file_storage_location = "./results/Cook_County_results.xlsx"

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

