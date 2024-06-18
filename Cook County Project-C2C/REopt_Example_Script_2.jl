#=

Example Script #2
This is example code for running REopt.jl, which shows more functionality than example script #1

Great resource for additional information: https://nrel.github.io/REopt.jl/dev/ 

=#

using REopt, Pkg, JuMP, CSV, DataFrames, Plots, Statistics, Cbc, PlotlyJS, Formatting, Dates, JSON
Pkg.status()

# Enter the location of the folder 
CurrentFolder = ""
# Enter your NREL developer key here
ENV["NREL_DEVELOPER_API_KEY"] = "" 


include(CurrentFolder*"/Dispatch_Plotting_Code.jl") # include the filepath for the Dispatch_Plotting_Code.jl file
cd(CurrentFolder) # this sets your current directory's filepath

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
        display(Plots.title!("Data for survival of the maximum outage length, by month"))

        ProbabilityOfSurvivingMaximumOutageLength_percent =  100*reliability_results["mean_cumulative_survival_final_time_step"]
        #print("\n The probability of survival, considering component reliability, for the maximum outage length is: "*string(ProbabilityOfSurvivingMaximumOutageLength_percent)*"%")
        #print("\n ")

        return reliability_results, ProbabilityOfSurvivingMaximumOutageLength_percent
end 

# Load an electric load file - should be a single column without any headers of kW's
load_filename =   CurrentFolder*"/InputData/loads_example.csv" 
    loads = CSV.read(load_filename, DataFrame, header=false)[!,"Column1"]

# load a critical load file
critical_load_filename = CurrentFolder*"/InputData/critical_loads_example.csv"
    critical_loads = CSV.read(critical_load_filename, DataFrame, header=false)[!,"Column1"]

# Load the electricity tariff structure 
load_electric_rate = CurrentFolder*"/InputData/utility_rate_example.json"
    utility_rate = JSON.parsefile(load_electric_rate)

# Plot the full set of load data, with time and date labels 
dates = DateTime("2017-01-01T00:00:00.000"):Dates.Hour(1):DateTime("2017-12-31T23:00:00.000")
display(Plots.plot(dates, loads, label = "Electric Load"))

# Plot part of the load data 
Hours = collect(1:8760)/24
Plots.plot!(Hours,loads, label = "Electric Load")
Plots.xlims!(30,37) 
Plots.xlabel!("Hours")
Plots.ylabel!("kW")
display(Plots.title!("Load Input to REopt"))


inputs = Dict([
    ("Settings",Dict([
        ("time_steps_per_hour", 1),
        ("include_climate_in_objective", false), # default is false; this only includes the costs from CO2 (use the health cost field to include costs from other emissions)
        ("off_grid_flag", false),
        ("add_soc_incentive", false)
        ])),
    ("Site",Dict([
        ("longitude", -90.0715), # New Orleans
        ("latitude", 29.9511), # New Orleans
        ("land_acres", 100),
        #("roof_squarefeet", 250), 
        ("renewable_electricity_min_fraction", 0.0 ), #minRE)
        ("renewable_electricity_max_fraction", 1.0) # maxRE)
        
        ])),
    ("ElectricLoad",Dict([
        #("doe_reference_name", "MidriseApartment"),
        ("year", 2017),
        ("loads_kw", loads), # takes the loads vector from above
        ("critical_loads_kw", critical_loads), #input an array with the critical loads defined
        #("critical_load_fraction", 0.5)
        ])),
    ("ElectricUtility",Dict([
        # note: outage indexing begins with 1 (not zero), so the index for Jan 1 6am - 7am is: 7 to 7. midnight to 3am is 0 to 3. The outage is inclusive of the outage time step.
        ("outage_start_time_step", 300*24), # add for resilience - but should use the outage simulator instead (set these to zero and then run the outage simulator)
        ("outage_end_time_step", (300*24)+28),   # add for resilience
        ("net_metering_limit_kw", 0),
        ])),
    ("ElectricTariff",Dict([ 
        #("blended_annual_energy_rate", 0.12),
        #("blended_annual_demand_rate", 6),
        ("urdb_response", utility_rate),
        #("wholesale_rate", 0.04) # can be a single value or a list of hourly, 30 min, or 15 min interval year-long data
        ])),
    ("Financial",Dict([ 
        #("om_cost_escalation_rate_fraction", 0.0),
        ("elec_cost_escalation_rate_fraction", 0.019),
        #("generator_fuel_cost_escalation_rate_fraction", 0.027),
        ("third_party_ownership", false),
        ("owner_tax_rate_fraction", 0.26),
        ("owner_discount_rate_fraction", 0.0564),
        ("offtaker_tax_rate_fraction", 0.26), # without third party ownership, offtaker is used instead of owner
        ("offtaker_discount_rate_fraction", 0.0564),
        ("analysis_years", 25),
        #("CO2_cost_per_tonne", 185),  
        #("CO2_cost_escalation_rate_fraction", 0.042173)
        ])),
    ("PV",
        #[  # uncomment this line to generate an array of dictionaries for multiple PV array inputs 
        Dict([
        ("name", "PV"),  
        ("location", "ground"), # list either "roof", "ground", or "both"
        ("min_kw", 0), 
        ("max_kw", 1000), 
        ("array_type", 0),  # 0 for ground fixed mount, 4 other design options
        ("tilt", 30),  # is by default the abs(latitude)
        ("azimuth", 180), # set to zero for southern hemisphere
        ("installed_cost_per_kw", 1800), # should be 1330 for the Boise baseline cost 
        ("om_cost_per_kw", 17),
        ("degradation_fraction",0.005),
        ("macrs_option_years", 7),
        ("macrs_bonus_fraction", 0.4),
        ("macrs_itc_reduction", 0.5),
        ("kw_per_square_foot", 0.001),
        ("acres_per_kw", 0.003),
        #("inv_eff", 0.96),
        ("dc_ac_ratio",1.2),
        #("production_factor_series", pf_data), # use this if uploading own solar data 
        ("federal_itc_fraction", 0.3),
        ("federal_rebate_per_kw",0),
        ("state_ibi_fraction",0),
        ("state_ibi_max",0),
        ("state_rebate_per_kw",0),
        ("state_rebate_max",0),
        ("utility_ibi_fraction",0),
        ("utility_ibi_max",0),
        ("utility_rebate_per_kw",0),
        ("utility_rebate_max",0),
        ("production_incentive_per_kwh",0),
        ("production_incentive_max_benefit",0),
        ("production_incentive_years",0),
        ("production_incentive_max_kw",0), 
        ("can_net_meter", true),
        ("can_curtail", true),
        ("can_wholesale", false)
        ])),
    ("ElectricStorage",Dict([
        ("min_kw", 0),
        ("max_kw", 1000),
        ("min_kwh", 0), 
        ("max_kwh", 1000), 
        ("internal_efficiency_fraction", 0.975),
        ("inverter_efficiency_fraction", 0.96),
        ("rectifier_efficiency_fraction", 0.96),
        ("soc_min_fraction", 0.2), 
        ("soc_init_fraction", 0.5),  
        ("can_grid_charge", true),
        ("installed_cost_per_kw", 775), 
        ("installed_cost_per_kwh", 388), 
        ("replace_cost_per_kw", 440), 
        ("replace_cost_per_kwh", 220),
        ("inverter_replacement_year", 10),
        ("battery_replacement_year",10),
        ("macrs_option_years", 5),
        ("macrs_bonus_fraction", 0.4),
        ("macrs_itc_reduction", 0.5),
        ("total_itc_fraction",0.3),
        ("total_rebate_per_kw",0),
        ("total_rebate_per_kwh",0)
        ])),
    #=
    ("Generator",Dict([
            ("existing_kw", 0),
            ("min_kw",0),
            ("max_kw", 250),
            ("installed_cost_per_kw", 0.0),
            ("om_cost_per_kw", 0.0),
            ("om_cost_per_kwh",0),
            ("fuel_cost_per_gallon", 3.0),
            ("fuel_avail_gal", 250),  
            ("min_turn_down_fraction", 0),
            ("only_runs_during_grid_outage", true),
            ("sells_energy_back_to_grid", false),
            ("can_net_meter", false),
            ("can_wholesale", false),
            ("can_export_beyond_nem_limit", false),
            ("can_curtail", false),
            ("macrs_option_years",0),
            ("macrs_bonus_fraction", 0.0),
            ("macrs_itc_reduction", 0),
            ("federal_itc_fraction", 0.0),
            ("federal_rebate_per_kw",0),
            ("state_ibi_fraction",0),
            ("state_ibi_max", 0.0),
            ("state_rebate_per_kw",0),
            ("state_rebate_max",0),
            ("utility_ibi_fraction",0),
            ("utility_ibi_max", 0.0),
            ("utility_rebate_per_kw",0),
            ("utility_rebate_max",0.0),
            ("production_incentive_per_kwh",0),
            ("production_incentive_max_benefit",0.0),
            ("production_incentive_years",0),
            ("production_incentive_max_kw",0.0),
            ("fuel_renewable_energy_fraction",0)
            # emissions factor data not included
            #("replacement_year",10)
            #("replace_cost_per_kw") # this is defined by REopt as the installed_cost_per_kw             
        ]))
        =#
])

# Convert the REopt inputs into the proper format
inputs_scenario = Scenario(inputs)
REopt_Inputs = REoptInputs(inputs_scenario)

# create two models so that REopt computes the Business As Usual (BAU) case
m1 = Model(Cbc.Optimizer)
m2 = Model(Cbc.Optimizer)

# run the models
results = run_reopt([m1,m2], REopt_Inputs)

# Make a dispatch plot
display(DispatchPlot(results, inputs["ElectricLoad"]["year"]))

#Print some results

PVSize = results["PV"]["size_kw"]
Battery_kw = results["ElectricStorage"]["size_kw"]
Battery_kwh = results["ElectricStorage"]["size_kwh"]
NPV = results["Financial"]["npv"]

print("\n The status of the solver is: "*results["status"])
print("\n The optimal PV size is: "*format(round(PVSize, digits = 1), commas = true)*" kW")
print("\n The optimal battery size is: "*format(round(Battery_kw, digits = 1), commas=true)*" kW  and "*format(round(Battery_kwh, digits = 1), commas=true)*" kWh")
print("\n The net present value is: "*format(round(NPV, digits=0), commas = true))



# Run the ERP multiple times for 1 or more outage durations
outagelengths =  [4, 12, 24] # in time steps - for instance, with interval electric load data with 1-hr intervals, 4 is 4 hours
outagelength_survivalpercent = zeros(length(outagelengths))
for i in 1:length(outagelengths)
    print("\n Starting the ERP run "*string(i)*" of "*string(length(outagelengths)))
    outagelength = outagelengths[i]
    ERPResults, PercentSurvival = ERP_run(REopt_results = results, REopt_post_inputs = REopt_Inputs, post = inputs, maximumoutageduration = outagelength)
    print("\n The ERP is completed")
    outagelength_survivalpercent[i] = PercentSurvival
    print("\n  The percent survival for "*string(outagelength)*" time steps is: "*string(PercentSurvival)*" %")
end



