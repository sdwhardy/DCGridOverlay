# Script to create the DC grid overlay project's grid
# Refer to the excel file in the package
# 7th August 2023
using XLSX
using PowerModels; const _PM = PowerModels
using PowerModelsACDC; const _PMACDC = PowerModelsACDC

function create_grid(;output_filename::String = "./test_cases/DC_overlay_grid.json")

    # Uploading an example test system
    test_case_5_acdc = "case5_acdc.m"
    s = Dict("output" => Dict("branch_flows" => true), "conv_losses_mp" => true)
    data_file_5_acdc = "./test_cases/$test_case_5_acdc"
    test_grid = _PM.parse_file(joinpath("./$(data_file_5_acdc)"))
    _PMACDC.process_additional_data!(test_grid)

    test_file = "./test_cases/DC_overlay_grid.xlsx"
    #test_grid = _PM.parse_file(test_file)

    # Calling the Excel file
    xf = XLSX.readxlsx(joinpath("./$(test_file)"))
    XLSX.sheetnames(xf)
    number_of_buses = 6

    # Creating a European grid dictionary in PowerModels format
    DC_overlay_grid = Dict{String,Any}()
    DC_overlay_grid["dcpol"] = 2
    DC_overlay_grid["name"] = "DC_overlay_grid"
    DC_overlay_grid["baseMVA"] = 100
    DC_overlay_grid["base_kv_AC"] = 380
    DC_overlay_grid["base_kv_DC"] = 525
    DC_overlay_grid["per_unit"] = true #adapt values from Excel file
    DC_overlay_grid["Z_base"] = (DC_overlay_grid["base_kv_AC"])^2/(DC_overlay_grid["baseMVA"]*1000)
    

    # Buses
    DC_overlay_grid["bus"] = Dict{String,Any}()
    for idx in 1:number_of_buses
        DC_overlay_grid["bus"]["$idx"] = Dict{String,Any}()
        DC_overlay_grid["bus"]["$idx"]["name"] = "Bus_$(idx)"
        DC_overlay_grid["bus"]["$idx"]["index"] = idx #i
        DC_overlay_grid["bus"]["$idx"]["bus_i"] = idx 
        if idx == 3
            DC_overlay_grid["bus"]["$idx"]["bus_type"] = 3 #slack bus
        else
            DC_overlay_grid["bus"]["$idx"]["bus_type"] = 1 
        end
        DC_overlay_grid["bus"]["$idx"]["pd"] = 0.0 
        DC_overlay_grid["bus"]["$idx"]["qd"] = 0.0 
        DC_overlay_grid["bus"]["$idx"]["gs"] = 0.0 
        DC_overlay_grid["bus"]["$idx"]["bs"] = 0.0 
        DC_overlay_grid["bus"]["$idx"]["area"] = 1 
        #DC_overlay_grid["bus"]["$idx"]["vm"] = xx
        #DC_overlay_grid["bus"]["$idx"]["va"] = xx
        DC_overlay_grid["bus"]["$idx"]["base_kv"] = 380 #kV
        DC_overlay_grid["bus"]["$idx"]["vmax"] = 1.1
        DC_overlay_grid["bus"]["$idx"]["vmin"] = 0.9
        DC_overlay_grid["bus"]["$idx"]["source_id"] = []
        push!(DC_overlay_grid["bus"]["$idx"]["source_id"],"bus")
        push!(DC_overlay_grid["bus"]["$idx"]["source_id"],idx)
    end

    # Bus dc -> it can be commented out if you want to keep only the AC system
    DC_overlay_grid["busdc"] = Dict{String,Any}()
    for idx in 1:number_of_buses
        DC_overlay_grid["busdc"]["$idx"] = Dict{String,Any}()
        DC_overlay_grid["busdc"]["$idx"]["name"] = "Bus_dc_$(idx)"
        DC_overlay_grid["busdc"]["$idx"]["index"] = idx
        DC_overlay_grid["busdc"]["$idx"]["busdc_i"] = idx
        DC_overlay_grid["busdc"]["$idx"]["bus_type"] = 1
        DC_overlay_grid["busdc"]["$idx"]["pd"] = 0.0
        DC_overlay_grid["busdc"]["$idx"]["qd"] = 0.0
        DC_overlay_grid["busdc"]["$idx"]["gs"] = 0.0
        DC_overlay_grid["busdc"]["$idx"]["bs"] = 0.0
        DC_overlay_grid["busdc"]["$idx"]["area"] = 2
        #DC_overlay_grid["busdc"]["$idx"]["vm"] = r[9] #buses_dc_dict[:,9][i]
        DC_overlay_grid["busdc"]["$idx"]["Vdcmax"] = 1.1
        DC_overlay_grid["busdc"]["$idx"]["Vdcmin"] = 0.9
        DC_overlay_grid["busdc"]["$idx"]["Vdc"] = 1
        DC_overlay_grid["busdc"]["$idx"]["Pdc"] = 0
        DC_overlay_grid["busdc"]["$idx"]["Cdc"] = 0
        DC_overlay_grid["busdc"]["$idx"]["source_id"] = []
        push!(DC_overlay_grid["busdc"]["$idx"]["source_id"],"busdc")
        push!(DC_overlay_grid["busdc"]["$idx"]["source_id"],idx)
    end

    # Branches
    DC_overlay_grid["branch"] = Dict{String,Any}()
    for r in XLSX.eachrow(xf["AC_grid_build"])
        i = XLSX.row_number(r)
        if i > 1 && i <= 9 
            # compensate for header and limit the number of rows to the existing branches
            idx = i - 1
            DC_overlay_grid["branch"]["$idx"] = Dict{String,Any}()
            DC_overlay_grid["branch"]["$idx"]["type"] = "AC_line"
            DC_overlay_grid["branch"]["$idx"]["index"] = idx
            DC_overlay_grid["branch"]["$idx"]["f_bus"] = parse(Int64,r[1][3])
            DC_overlay_grid["branch"]["$idx"]["t_bus"] = parse(Int64,r[1][4])
            DC_overlay_grid["branch"]["$idx"]["br_r"] = r[7]*r[3]/DC_overlay_grid["Z_base"]   
            DC_overlay_grid["branch"]["$idx"]["br_x"] = r[8]*r[3]/DC_overlay_grid["Z_base"]   
            DC_overlay_grid["branch"]["$idx"]["b_fr"] = 1/sqrt((DC_overlay_grid["branch"]["$idx"]["br_r"])^2+(DC_overlay_grid["branch"]["$idx"]["br_x"])^2) 
            DC_overlay_grid["branch"]["$idx"]["b_to"] = 1/sqrt((DC_overlay_grid["branch"]["$idx"]["br_r"])^2+(DC_overlay_grid["branch"]["$idx"]["br_x"])^2)
            DC_overlay_grid["branch"]["$idx"]["g_fr"] = 0.0
            DC_overlay_grid["branch"]["$idx"]["g_to"] = 0.0
            DC_overlay_grid["branch"]["$idx"]["rate_a"] = r[2]/DC_overlay_grid["baseMVA"]
            DC_overlay_grid["branch"]["$idx"]["rate_b"] = r[2]/DC_overlay_grid["baseMVA"] 
            DC_overlay_grid["branch"]["$idx"]["rate_c"] = r[2]/DC_overlay_grid["baseMVA"]
            DC_overlay_grid["branch"]["$idx"]["ratio"] = 1 #/ branches_dict[:,11][i]
            DC_overlay_grid["branch"]["$idx"]["angmin"] = -pi #deepcopy(test_grid["branch"]["1"]["angmin"])
            DC_overlay_grid["branch"]["$idx"]["angmax"] = pi #deepcopy(test_grid["branch"]["1"]["angmax"])
            DC_overlay_grid["branch"]["$idx"]["br_status"] = 1
            DC_overlay_grid["branch"]["$idx"]["tap"] = 1.0
            DC_overlay_grid["branch"]["$idx"]["transformer"] = false
            DC_overlay_grid["branch"]["$idx"]["shift"] = 0.0
            DC_overlay_grid["branch"]["$idx"]["source_id"] = []
            push!(DC_overlay_grid["branch"]["$idx"]["source_id"],"branch")
            push!(DC_overlay_grid["branch"]["$idx"]["source_id"], idx)
        end
    end

    # DC Branches
    DC_overlay_grid["branchdc"] = Dict{String,Any}()
    for r in XLSX.eachrow(xf["DC_grid_build"])
        i = XLSX.row_number(r)
        if i > 1 && i <= 9 # compensate for header and limit the number of rows to the existing branches
            idx = i - 1
            DC_overlay_grid["branchdc"]["$idx"] = Dict{String,Any}()
            DC_overlay_grid["branchdc"]["$idx"]["type"] = "DC_line" #branches_dc_dict[:,2][i]
            DC_overlay_grid["branchdc"]["$idx"]["index"] = idx
            DC_overlay_grid["branchdc"]["$idx"]["fbusdc"] = parse(Int64,r[1][3])
            DC_overlay_grid["branchdc"]["$idx"]["tbusdc"] = parse(Int64,r[1][4])
            DC_overlay_grid["branchdc"]["$idx"]["r"] = r[6]*r[3]/DC_overlay_grid["Z_base"]
            DC_overlay_grid["branchdc"]["$idx"]["c"] = 0.0
            DC_overlay_grid["branchdc"]["$idx"]["l"] = 0.0
            DC_overlay_grid["branchdc"]["$idx"]["status"] = 1
            DC_overlay_grid["branchdc"]["$idx"]["rateA"] = r[2]/DC_overlay_grid["baseMVA"] #
            DC_overlay_grid["branchdc"]["$idx"]["rateB"] = r[2]/DC_overlay_grid["baseMVA"] #
            DC_overlay_grid["branchdc"]["$idx"]["rateC"] = r[2]/DC_overlay_grid["baseMVA"] # 
            DC_overlay_grid["branchdc"]["$idx"]["ratio"] = 1
            DC_overlay_grid["branchdc"]["$idx"]["source_id"] = []
            push!(DC_overlay_grid["branchdc"]["$idx"]["source_id"],"branchdc")
            push!(DC_overlay_grid["branchdc"]["$idx"]["source_id"], idx)
        end
    end

    # DC Converters
    DC_overlay_grid["convdc"] = Dict{String,Any}()
    for idx in 1:number_of_buses
            DC_overlay_grid["convdc"]["$idx"] = Dict{String,Any}()
            DC_overlay_grid["convdc"]["$idx"] = deepcopy(test_grid["convdc"]["1"])  # To fill in default values......
            DC_overlay_grid["convdc"]["$idx"]["busdc_i"] = idx #conv_dc_dict[:,3][i]
            DC_overlay_grid["convdc"]["$idx"]["busac_i"] = idx #conv_dc_dict[:,2][i]
            DC_overlay_grid["convdc"]["$idx"]["index"] = idx
            DC_overlay_grid["convdc"]["$idx"]["status"] = 1
            DC_overlay_grid["convdc"]["$idx"]["Pacmax"] = 200.0 # Values to not having this power constraining the OPF
            DC_overlay_grid["convdc"]["$idx"]["Pacmin"] = 200.0
            DC_overlay_grid["convdc"]["$idx"]["Qacmin"] = 200.0
            DC_overlay_grid["convdc"]["$idx"]["Qacmax"] = 200.0
            #DC_overlay_grid["convdc"]["$idx"]["Imax"] = value of 1.11 taken from the test grid
            DC_overlay_grid["convdc"]["$idx"]["Pg"] = 0.0 #Adjusting with pu values
            DC_overlay_grid["convdc"]["$idx"]["ratio"] = 1
            DC_overlay_grid["convdc"]["$idx"]["transformer"] = 1
            DC_overlay_grid["convdc"]["$idx"]["reactor"] = 1
            DC_overlay_grid["convdc"]["$idx"]["source_id"] = []
            push!(DC_overlay_grid["convdc"]["$idx"]["source_id"],"convdc")
            push!(DC_overlay_grid["convdc"]["$idx"]["source_id"], idx)
    end

    #-> Worked till here in the morning, missing only gen and load 

    # Load
    DC_overlay_grid["load"] = Dict{String,Any}()
    for r in XLSX.eachrow(xf["DEMAND"])
        i = XLSX.row_number(r)
        if i > 1 # compensate for header
            idx = i - 1
            DC_overlay_grid["load"]["$idx"] = Dict{String,Any}()
            DC_overlay_grid["load"]["$idx"]["country"] = r[1] #load_dict[:,1][i]
            DC_overlay_grid["load"]["$idx"]["zone"] = r[2] #load_dict[:,2][i]
            DC_overlay_grid["load"]["$idx"]["load_bus"] = r[3] #load_dict[:,3][i]
            DC_overlay_grid["load"]["$idx"]["pmax"] = r[4] /DC_overlay_grid["baseMVA"] #load_dict[:,4][i]/DC_overlay_grid["baseMVA"] 
            if r[5] == "-"
                DC_overlay_grid["load"]["$idx"]["cosphi"] = 1
            else
                DC_overlay_grid["load"]["$idx"]["cosphi"] = r[5] #load_dict[:,5][i]
            end
            DC_overlay_grid["load"]["$idx"]["pd"] = r[4] / DC_overlay_grid["baseMVA"] # load_dict[:,4][i]/DC_overlay_grid["baseMVA"] 
            DC_overlay_grid["load"]["$idx"]["qd"] = r[4] / DC_overlay_grid["baseMVA"] * sqrt(1 - DC_overlay_grid["load"]["$idx"]["cosphi"]^2)
            DC_overlay_grid["load"]["$idx"]["index"] = idx
            DC_overlay_grid["load"]["$idx"]["status"] = 1
            DC_overlay_grid["load"]["$idx"]["source_id"] = []
            push!(DC_overlay_grid["load"]["$idx"]["source_id"],"bus")
            push!(DC_overlay_grid["load"]["$idx"]["source_id"], idx)
        end
    end

    # Including load
    for (l_id,l) in DC_overlay_grid["load"]
        for r in XLSX.eachrow(xf["DEMAND_OVERVIEW"])
            i = XLSX.row_number(r)
            if i > 1
                if l["zone"] == r[1]
                    l["country_peak_load"] = r[2] / DC_overlay_grid["baseMVA"] 
                end
            end
        end
        l["powerportion"] = l["pmax"]/l["country_peak_load"]
    end

    ####### READ IN GENERATION DATA #######################
    DC_overlay_grid["gen"] = Dict{String,Any}()
    for r in XLSX.eachrow(xf["GEN"])
        i = XLSX.row_number(r)
        if i > 1 # compensate for header
            idx = i - 1
            DC_overlay_grid["gen"]["$idx"] = Dict{String,Any}()
            DC_overlay_grid["gen"]["$idx"]["index"] = idx
            DC_overlay_grid["gen"]["$idx"]["country"] = r[4] #xf["GEN"]["D2:D1230"][idx]
            DC_overlay_grid["gen"]["$idx"]["zone"] = r[5] #xf["GEN"]["E2:E1230"][idx]
            DC_overlay_grid["gen"]["$idx"]["gen_bus"] = r[6] #xf["GEN"]["F2:F1230"][idx]
            DC_overlay_grid["gen"]["$idx"]["pmax"] = r[8] / DC_overlay_grid["baseMVA"]  #xf["GEN"]["H2:H1230"][idx]/DC_overlay_grid["baseMVA"] 
            DC_overlay_grid["gen"]["$idx"]["pmin"] = r[7] / DC_overlay_grid["baseMVA"] #xf["GEN"]["G2:G1230"][idx]/DC_overlay_grid["baseMVA"] 
            DC_overlay_grid["gen"]["$idx"]["qmax"] =  DC_overlay_grid["gen"]["$idx"]["pmax"] * 0.5
            DC_overlay_grid["gen"]["$idx"]["qmin"] = -DC_overlay_grid["gen"]["$idx"]["pmax"] * 0.5
            DC_overlay_grid["gen"]["$idx"]["cost"] = [r[13] * DC_overlay_grid["baseMVA"], 0.0]  # Assumption here, to be checked
            DC_overlay_grid["gen"]["$idx"]["marginal_cost"] = r[11] * DC_overlay_grid["baseMVA"]  # Assumption here, to be checked
            DC_overlay_grid["gen"]["$idx"]["co2_add_on"] = r[12] * DC_overlay_grid["baseMVA"]  # Assumption here, to be checked
            DC_overlay_grid["gen"]["$idx"]["ncost"] = 2
            DC_overlay_grid["gen"]["$idx"]["model"] = 2
            DC_overlay_grid["gen"]["$idx"]["gen_status"] = 1
            DC_overlay_grid["gen"]["$idx"]["vg"] = 1.0
            DC_overlay_grid["gen"]["$idx"]["source_id"] = []
            DC_overlay_grid["gen"]["$idx"]["name"] = r[1]  # Assumption here, to be checked
            push!(DC_overlay_grid["gen"]["$idx"]["source_id"],"gen")
            push!(DC_overlay_grid["gen"]["$idx"]["source_id"], idx)

            type = r[2]
            if type == "Gas"
                type_tyndp = "Gas CCGT new"
            elseif type == "Oil"
                type_tyndp = "Heavy oil old 1 Bio"
            elseif type == "Nuclear"
                type_tyndp = "Nuclear"
            elseif type == "Biomass"
                type_tyndp = "Other RES"
            elseif type == "Hard Coal"
                type_tyndp = "Hard coal old 2 Bio"
            elseif type == "Lignite"
                type_tyndp = "Lignite old 1"
            end
            DC_overlay_grid["gen"]["$idx"]["type"] = type
            DC_overlay_grid["gen"]["$idx"]["type_tyndp"] = type_tyndp
        end
    end

    # Run-off-river
    number_of_gens = maximum([gen["index"] for (g, gen) in DC_overlay_grid["gen"]])
    for r in XLSX.eachrow(xf["ROR"])
        i = XLSX.row_number(r)
        if i > 1 # compensate for header
            idx = number_of_gens + i - 1 # 
            DC_overlay_grid["gen"]["$idx"] = Dict{String,Any}()
            DC_overlay_grid["gen"]["$idx"]["index"] = idx 
            DC_overlay_grid["gen"]["$idx"]["country"] = r[3] #gen_hydro_ror_dict[:,3][i]
            DC_overlay_grid["gen"]["$idx"]["zone"] = r[4] #gen_hydro_ror_dict[:,4][i]
            DC_overlay_grid["gen"]["$idx"]["gen_bus"] = r[5] #gen_hydro_ror_dict[:,5][i]
            DC_overlay_grid["gen"]["$idx"]["type"] = r[1] #gen_hydro_ror_dict[:,1][i]
            DC_overlay_grid["gen"]["$idx"]["type_tyndp"] = "Run-of-River"
            DC_overlay_grid["gen"]["$idx"]["pmax"] = r[6] / DC_overlay_grid["baseMVA"] #gen_hydro_ror_dict[:,6][i]/DC_overlay_grid["baseMVA"] 
            DC_overlay_grid["gen"]["$idx"]["pmin"] = 0.0
            DC_overlay_grid["gen"]["$idx"]["qmax"] =  DC_overlay_grid["gen"]["$idx"]["pmax"] * 0.5
            DC_overlay_grid["gen"]["$idx"]["qmin"] = -DC_overlay_grid["gen"]["$idx"]["pmax"] * 0.5
            DC_overlay_grid["gen"]["$idx"]["cost"] = [25.0  * DC_overlay_grid["baseMVA"], 0.0] # Assumption here, to be checked
            DC_overlay_grid["gen"]["$idx"]["ncost"] = 2
            DC_overlay_grid["gen"]["$idx"]["model"] = 2
            DC_overlay_grid["gen"]["$idx"]["gen_status"] = 1
            DC_overlay_grid["gen"]["$idx"]["vg"] = 1.0
            DC_overlay_grid["gen"]["$idx"]["source_id"] = []
            push!(DC_overlay_grid["gen"]["$idx"]["source_id"],"gen")
            push!(DC_overlay_grid["gen"]["$idx"]["source_id"], idx)
        end
    end

    ##### ONSHORE WIND
    number_of_gens = maximum([gen["index"] for (g, gen) in DC_overlay_grid["gen"]])
    for r in XLSX.eachrow(xf["ONSHORE"])
        i = XLSX.row_number(r)
        if i > 1 # compensate for header
            idx = number_of_gens + i - 1 # 
            DC_overlay_grid["gen"]["$idx"] = Dict{String,Any}()
            DC_overlay_grid["gen"]["$idx"]["index"] = idx
            DC_overlay_grid["gen"]["$idx"]["country"] = r[3] #gen_onshore_wind_dict[:,3][i]
            DC_overlay_grid["gen"]["$idx"]["zone"] = r[4] #gen_onshore_wind_dict[:,4][i]
            DC_overlay_grid["gen"]["$idx"]["gen_bus"] = r[5] #gen_onshore_wind_dict[:,5][i]
            DC_overlay_grid["gen"]["$idx"]["type"] = r[1] #gen_onshore_wind_dict[:,1][i]
            DC_overlay_grid["gen"]["$idx"]["type_tyndp"] = "Onshore Wind"
            DC_overlay_grid["gen"]["$idx"]["pmax"] = r[6] /DC_overlay_grid["baseMVA"] #gen_onshore_wind_dict[:,6][i]/DC_overlay_grid["baseMVA"] 
            DC_overlay_grid["gen"]["$idx"]["pmin"] = 0.0
            DC_overlay_grid["gen"]["$idx"]["qmax"] =  DC_overlay_grid["gen"]["$idx"]["pmax"] * 0.5
            DC_overlay_grid["gen"]["$idx"]["qmin"] = -DC_overlay_grid["gen"]["$idx"]["pmax"] * 0.5
            DC_overlay_grid["gen"]["$idx"]["cost"] = [25.0  * DC_overlay_grid["baseMVA"] ,0.0] 
            DC_overlay_grid["gen"]["$idx"]["ncost"] = 2
            DC_overlay_grid["gen"]["$idx"]["model"] = 2
            DC_overlay_grid["gen"]["$idx"]["gen_status"] = 1
            DC_overlay_grid["gen"]["$idx"]["vg"] = 1.0
            DC_overlay_grid["gen"]["$idx"]["source_id"] = []
            push!(DC_overlay_grid["gen"]["$idx"]["source_id"],"gen")
            push!(DC_overlay_grid["gen"]["$idx"]["source_id"], idx)
        end
    end
    ####### OFFSHORE WIND
    number_of_gens = maximum([gen["index"] for (g, gen) in DC_overlay_grid["gen"]])
    for r in XLSX.eachrow(xf["OFFSHORE"])
        i = XLSX.row_number(r)
        if i > 1 # compensate for header
            idx = number_of_gens + i - 2 # Weird read in bug....
            DC_overlay_grid["gen"]["$idx"] = Dict{String,Any}()
            DC_overlay_grid["gen"]["$idx"]["index"] = idx
            DC_overlay_grid["gen"]["$idx"]["country"] = r[3] #gen_offshore_wind_dict[:,3][i]
            DC_overlay_grid["gen"]["$idx"]["zone"] = r[4] #gen_offshore_wind_dict[:,4][i]
            DC_overlay_grid["gen"]["$idx"]["gen_bus"] = r[5] #gen_offshore_wind_dict[:,5][i]
            DC_overlay_grid["gen"]["$idx"]["type"] = r[1] #gen_offshore_wind_dict[:,1][i]
            DC_overlay_grid["gen"]["$idx"]["type_tyndp"] = "Offshore Wind"
            DC_overlay_grid["gen"]["$idx"]["pmax"] = r[6] / DC_overlay_grid["baseMVA"]  #gen_offshore_wind_dict[:,6][i]/DC_overlay_grid["baseMVA"] 
            DC_overlay_grid["gen"]["$idx"]["pmin"] = 0.0
            DC_overlay_grid["gen"]["$idx"]["qmax"] =  DC_overlay_grid["gen"]["$idx"]["pmax"] * 0.5
            DC_overlay_grid["gen"]["$idx"]["qmin"] = -DC_overlay_grid["gen"]["$idx"]["pmax"] * 0.5
            DC_overlay_grid["gen"]["$idx"]["cost"] = [59.0 * DC_overlay_grid["baseMVA"],0.0] 
            DC_overlay_grid["gen"]["$idx"]["ncost"] = 2
            DC_overlay_grid["gen"]["$idx"]["model"] = 2
            DC_overlay_grid["gen"]["$idx"]["gen_status"] = 1
            DC_overlay_grid["gen"]["$idx"]["vg"] = 1.0
            DC_overlay_grid["gen"]["$idx"]["source_id"] = []
            push!(DC_overlay_grid["gen"]["$idx"]["source_id"],"gen")
            push!(DC_overlay_grid["gen"]["$idx"]["source_id"], idx)
        end
    end

    ##### SOLAR PV
    number_of_gens = maximum([gen["index"] for (g, gen) in DC_overlay_grid["gen"]])
    for r in XLSX.eachrow(xf["SOLAR"])
        i = XLSX.row_number(r)
        if i > 1 # compensate for header
            idx = number_of_gens + i - 1 # 
            DC_overlay_grid["gen"]["$idx"] = Dict{String,Any}()
            DC_overlay_grid["gen"]["$idx"]["index"] = idx
            DC_overlay_grid["gen"]["$idx"]["country"] = r[3] #gen_solar_dict[:,3][i]
            DC_overlay_grid["gen"]["$idx"]["zone"] = r[4] #gen_solar_dict[:,4][i]
            DC_overlay_grid["gen"]["$idx"]["gen_bus"] = r[5] #gen_solar_dict[:,5][i]
            DC_overlay_grid["gen"]["$idx"]["type"] = r[1] #gen_solar_dict[:,1][i]
            DC_overlay_grid["gen"]["$idx"]["type_tyndp"] = "Solar PV"
            DC_overlay_grid["gen"]["$idx"]["pmax"] = r[6] / DC_overlay_grid["baseMVA"]  #gen_solar_dict[:,6][i]/DC_overlay_grid["baseMVA"] 
            DC_overlay_grid["gen"]["$idx"]["pmin"] = 0.0
            DC_overlay_grid["gen"]["$idx"]["qmax"] =  DC_overlay_grid["gen"]["$idx"]["pmax"] * 0.5
            DC_overlay_grid["gen"]["$idx"]["qmin"] = -DC_overlay_grid["gen"]["$idx"]["pmax"] * 0.5
            DC_overlay_grid["gen"]["$idx"]["cost"] = [18.0  * DC_overlay_grid["baseMVA"],0.0] 
            DC_overlay_grid["gen"]["$idx"]["ncost"] = 2
            DC_overlay_grid["gen"]["$idx"]["model"] = 2
            DC_overlay_grid["gen"]["$idx"]["gen_status"] = 1
            DC_overlay_grid["gen"]["$idx"]["vg"] = 1.0
            DC_overlay_grid["gen"]["$idx"]["source_id"] = []
            push!(DC_overlay_grid["gen"]["$idx"]["source_id"],"gen")
            push!(DC_overlay_grid["gen"]["$idx"]["source_id"], idx)
        end
    end


    # ####### Hydro reservoir as storage
    DC_overlay_grid["storage"] = Dict{String, Any}()
    for r in XLSX.eachrow(xf["RES"])
        i = XLSX.row_number(r)
        if i > 1 # compensate for header
            idx = i - 1 # 
            DC_overlay_grid["storage"]["$idx"] = Dict{String,Any}()
            DC_overlay_grid["storage"]["$idx"]["index"] = idx
            DC_overlay_grid["storage"]["$idx"]["country"] = r[3] #hydro_res_dict[:,3][i]
            if r[4] == "DE-LU"   # FIX for Germany, weirdly zone is "DE-LU" in data model, altough country is DE
                DC_overlay_grid["storage"]["$idx"]["zone"] = "DE" #hydro_res_dict[:,4][i]
            else
                DC_overlay_grid["storage"]["$idx"]["zone"] = r[4]
            end
            DC_overlay_grid["storage"]["$idx"]["storage_bus"] = r[5] #hydro_res_dict[:,5][i]
            DC_overlay_grid["storage"]["$idx"]["type"] = r[1] #hydro_res_dict[:,1][i]
            DC_overlay_grid["storage"]["$idx"]["type_tyndp"] = "Reservoir"
            DC_overlay_grid["storage"]["$idx"]["ps"] = 0.0
            DC_overlay_grid["storage"]["$idx"]["qs"] = 0.0
            DC_overlay_grid["storage"]["$idx"]["energy"] = r[7] / DC_overlay_grid["baseMVA"] #hydro_res_dict[:,7][i]/DC_overlay_grid["baseMVA"]
            DC_overlay_grid["storage"]["$idx"]["energy_rating"] = r[7] / DC_overlay_grid["baseMVA"]
            DC_overlay_grid["storage"]["$idx"]["charge_rating"] = 0.0
            DC_overlay_grid["storage"]["$idx"]["discharge_rating"] = r[6] / DC_overlay_grid["baseMVA"] 
            DC_overlay_grid["storage"]["$idx"]["charge_efficiency"] = 1.0 
            DC_overlay_grid["storage"]["$idx"]["discharge_efficiency"] = 0.95 
            DC_overlay_grid["storage"]["$idx"]["thermal_rating"] = r[6] / DC_overlay_grid["baseMVA"]
            DC_overlay_grid["storage"]["$idx"]["qmax"] =  DC_overlay_grid["storage"]["$idx"]["thermal_rating"] * 0.5
            DC_overlay_grid["storage"]["$idx"]["qmin"] = -DC_overlay_grid["storage"]["$idx"]["thermal_rating"] * 0.5
            DC_overlay_grid["storage"]["$idx"]["r"] = 0.0
            DC_overlay_grid["storage"]["$idx"]["x"] = 0.0
            DC_overlay_grid["storage"]["$idx"]["p_loss"] = 0.0
            DC_overlay_grid["storage"]["$idx"]["q_loss"] = 0.0
            DC_overlay_grid["storage"]["$idx"]["status"] = 1
            DC_overlay_grid["storage"]["$idx"]["cost"] = [r[8] * DC_overlay_grid["baseMVA"] ,0.0] # Assumption here, to be checked
            DC_overlay_grid["storage"]["$idx"]["source_id"] = []
            push!(DC_overlay_grid["storage"]["$idx"]["source_id"],"storage")
            push!(DC_overlay_grid["storage"]["$idx"]["source_id"], idx)
        end
    end

    # ####### Pumped hydro stroage
    number_of_strg = maximum([storage["index"] for (s, storage) in DC_overlay_grid["storage"]])
    for r in XLSX.eachrow(xf["PHS"])
        i = XLSX.row_number(r)
        if i > 1 # compensate for header
            idx = number_of_strg + i - 1 # 
            DC_overlay_grid["storage"]["$idx"] = Dict{String,Any}()
            DC_overlay_grid["storage"]["$idx"]["index"] = idx
            DC_overlay_grid["storage"]["$idx"]["country"] = r[3] #hydro_phs_dict[:,3][i]
            if r[4] == "DE-LU"   # FIX for Germany, weirdly zone is "DE-LU" in data model, altough country is DE
                DC_overlay_grid["storage"]["$idx"]["zone"] = "DE" #hydro_res_dict[:,4][i]
            else
                DC_overlay_grid["storage"]["$idx"]["zone"] = r[4]
            end
            DC_overlay_grid["storage"]["$idx"]["storage_bus"] = r[5] #hydro_phs_dict[:,5][i]
            DC_overlay_grid["storage"]["$idx"]["type"] = r[1] #hydro_phs_dict[:,1][i]
            DC_overlay_grid["storage"]["$idx"]["type_tyndp"] = "Reservoir"
            DC_overlay_grid["storage"]["$idx"]["ps"] = 0.0
            DC_overlay_grid["storage"]["$idx"]["qs"] = 0.0
            DC_overlay_grid["storage"]["$idx"]["energy"] = r[8] / DC_overlay_grid["baseMVA"] / 2
            DC_overlay_grid["storage"]["$idx"]["energy_rating"] = r[8]/DC_overlay_grid["baseMVA"]
            DC_overlay_grid["storage"]["$idx"]["charge_rating"] =  -r[7]/DC_overlay_grid["baseMVA"]
            DC_overlay_grid["storage"]["$idx"]["discharge_rating"] = r[6]/DC_overlay_grid["baseMVA"] 
            DC_overlay_grid["storage"]["$idx"]["charge_efficiency"] = 1.0 
            DC_overlay_grid["storage"]["$idx"]["discharge_efficiency"] = 0.95 
            DC_overlay_grid["storage"]["$idx"]["thermal_rating"] = r[6]/DC_overlay_grid["baseMVA"]
            DC_overlay_grid["storage"]["$idx"]["qmax"] =  DC_overlay_grid["storage"]["$idx"]["thermal_rating"] * 0.5
            DC_overlay_grid["storage"]["$idx"]["qmin"] = -DC_overlay_grid["storage"]["$idx"]["thermal_rating"] * 0.5
            DC_overlay_grid["storage"]["$idx"]["r"] = 0.0
            DC_overlay_grid["storage"]["$idx"]["x"] = 0.0
            DC_overlay_grid["storage"]["$idx"]["p_loss"] = 0.0
            DC_overlay_grid["storage"]["$idx"]["q_loss"] = 0.0
            DC_overlay_grid["storage"]["$idx"]["status"] = 1
            DC_overlay_grid["storage"]["$idx"]["cost"] = [r[9] * DC_overlay_grid["baseMVA"] ,0.0] # Assumption here, to be checked
            DC_overlay_grid["storage"]["$idx"]["source_id"] = []
            push!(DC_overlay_grid["storage"]["$idx"]["source_id"],"storage")
            push!(DC_overlay_grid["storage"]["$idx"]["source_id"], idx)
        end
    end

    ############## OVERVIEW ###################
    # TO DO: Fix later with dynamic lenght of sheet.....
    zone_names = xf["BUS_OVERVIEW"]["F2:F43"]
    DC_overlay_grid["zonal_generation_capacity"] = Dict{String, Any}()
    DC_overlay_grid["zonal_peak_demand"] = Dict{String, Any}()

    for zone in zone_names
        idx = findfirst(zone .== zone_names)[1]
        # Generation
        DC_overlay_grid["zonal_generation_capacity"]["$idx"] = Dict{String, Any}()
        DC_overlay_grid["zonal_generation_capacity"]["$idx"]["zone"] = zone
            # Wind
            DC_overlay_grid["zonal_generation_capacity"]["$idx"]["Onshore Wind"] =  xf["WIND_OVERVIEW"]["B2:B43"][idx]/DC_overlay_grid["baseMVA"]
            DC_overlay_grid["zonal_generation_capacity"]["$idx"]["Offshore Wind"] =  xf["WIND_OVERVIEW"]["C2:C43"][idx]/DC_overlay_grid["baseMVA"]
            # PV
            DC_overlay_grid["zonal_generation_capacity"]["$idx"]["Solar PV"] =  xf["SOLAR_OVERVIEW"]["B2:B43"][idx]/DC_overlay_grid["baseMVA"]
            # Hydro
            DC_overlay_grid["zonal_generation_capacity"]["$idx"]["Run-of-River"] =  xf["HYDRO_OVERVIEW"]["B3:B44"][idx]/DC_overlay_grid["baseMVA"]
            DC_overlay_grid["zonal_generation_capacity"]["$idx"]["Reservoir"] =  xf["HYDRO_OVERVIEW"]["C3:C44"][idx]/DC_overlay_grid["baseMVA"]
            DC_overlay_grid["zonal_generation_capacity"]["$idx"]["Reservoir capacity"] =  xf["HYDRO_OVERVIEW"]["D3:D44"][idx]/DC_overlay_grid["baseMVA"]
            DC_overlay_grid["zonal_generation_capacity"]["$idx"]["PHS"] =  xf["HYDRO_OVERVIEW"]["E3:E44"][idx]/DC_overlay_grid["baseMVA"]
            DC_overlay_grid["zonal_generation_capacity"]["$idx"]["PHS capacity"] =  xf["HYDRO_OVERVIEW"]["F3:F44"][idx]/DC_overlay_grid["baseMVA"]
            # Thermal -> This may need to be updated, no nuclear in BE instead of 5.943 GW ...
            DC_overlay_grid["zonal_generation_capacity"]["$idx"]["Other RES"] =  xf["THERMAL_OVERVIEW"]["B2:B43"][idx]/DC_overlay_grid["baseMVA"]
            DC_overlay_grid["zonal_generation_capacity"]["$idx"]["Gas CCGT new"] =  xf["THERMAL_OVERVIEW"]["C2:C43"][idx]/DC_overlay_grid["baseMVA"]
            DC_overlay_grid["zonal_generation_capacity"]["$idx"]["Hard coal old 2 Bio"] =  xf["THERMAL_OVERVIEW"]["D2:D43"][idx]/DC_overlay_grid["baseMVA"]
            DC_overlay_grid["zonal_generation_capacity"]["$idx"]["Lignite old 1"] =  xf["THERMAL_OVERVIEW"]["E2:E43"][idx]/DC_overlay_grid["baseMVA"]
            DC_overlay_grid["zonal_generation_capacity"]["$idx"]["Nuclear"] =  xf["THERMAL_OVERVIEW"]["F2:F43"][idx]/DC_overlay_grid["baseMVA"]
            DC_overlay_grid["zonal_generation_capacity"]["$idx"]["Heavy oil old 1 Bio"] =  xf["THERMAL_OVERVIEW"]["G2:G43"][idx]/DC_overlay_grid["baseMVA"]
        # Demand
        DC_overlay_grid["zonal_peak_demand"]["$idx"] = xf["THERMAL_OVERVIEW"]["B2:B43"][idx]/DC_overlay_grid["baseMVA"]
    end
    ######
    # Making sure all the keys for PowerModels are there
    DC_overlay_grid["source_type"] = deepcopy(test_grid["source_type"])
    DC_overlay_grid["switch"] = deepcopy(test_grid["switch"])
    DC_overlay_grid["shunt"] = deepcopy(test_grid["shunt"])
    DC_overlay_grid["dcline"] = deepcopy(test_grid["dcline"])

    # Fixing NaN branches
    DC_overlay_grid["branch"]["6282"]["br_r"] = 0.001
    DC_overlay_grid["branch"]["8433"]["br_r"] = 0.001
    DC_overlay_grid["branch"]["8439"]["br_r"] = 0.001
    DC_overlay_grid["branch"]["8340"]["br_r"] = 0.001



    string_data = JSON.json(DC_overlay_grid)
    open(output_filename,"w" ) do f
        write(f,string_data)
    end

end