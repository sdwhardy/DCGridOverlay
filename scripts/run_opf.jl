-# Script to run the OPF simulations for the DC Grid overlay project using CbaOPF
# Script to create the DC grid overlay project's grid based on the develped function
# Refer to the excel file in the package
# 7th August 2023

using XLSX
using PowerModels; const _PM = PowerModels
using PowerModelsACDC; const _PMACDC = PowerModelsACDC
using JSON
using JuMP
using Ipopt, Gurobi
using PlotlyJS
using FileIO
using DataFrames, CSV 
include("create_grid_and_opf_functions.jl")


#########################################################################################################################
# One can choose the Pmax of the conv_power among 2.0, 4.0 and 8.0 GW
conv_power = 6.0
test_case_file = "DC_overlay_grid_$(conv_power)_GW_convdc.json"
test_case = _PM.parse_file("./test_cases/$test_case_file")

#remove DC grid
#=for i=1:1:8
    test_case["branchdc"][string(i)]["rateA"]=test_case["branchdc"][string(i)]["rateB"]=test_case["branchdc"][string(i)]["rateC"]=0.0
    end=#

#use real power only
#[test_case["load"][string(i)]["pd"]=0 for i in 1:1:6]
#=test_case["convdc"]["19"]["busac_i"]
test_case["convdc"]["19"]["busdc_i"]

test_case["bus"][string(1)]["bus_type"]=1
test_case["bus"][string(2)]["bus_type"]=1
test_case["bus"][string(4)]["bus_type"]=1
test_case["bus"][string(5)]["bus_type"]=1
test_case["bus"][string(6)]["bus_type"]=1
for i=5:1:8
    test_case["branch"][string(i)]["angmax"]=pi
    test_case["branch"][string(i)]["angmin"]=-pi
    end

    for i=5:1:8
        test_case["branch"][string(i)]["b_to"]=0.0
        test_case["branch"][string(i)]["b_fr"]=0.0
        end
        [delete!(test_case["branch"],s) for s in ["5","8","6","7"]]
        [delete!(test_case["branch"],s) for s in ["5"]]

        [delete!(test_case["branchdc"],s) for s in ["1","2","3","4","5","8","6","7"]]#Overlay

        [delete!(test_case["branchdc"],s) for s in ["9","10","11","12"]]#ptp
        [delete!(test_case["branchdc"],s) for s in ["12"]]#ptp
    #test_case["branch"]["5"]
=#
s = Dict("output" => Dict("branch_flows" => true), "conv_losses_mp" => true)
start_hour = 1
number_of_hours = 8760

RES_time_series_file = "DC_overlay_grid_RES_$(start_hour)_$(number_of_hours).json"
RES_time_series = JSON.parsefile("./test_cases/$RES_time_series_file")

load_time_series_file = "DC_overlay_grid_Demand_$(start_hour)_$(number_of_hours).json"
load_time_series = JSON.parsefile("./test_cases/$load_time_series_file")

##########################################################################
# Define solvers
##########################################################################
ipopt = JuMP.optimizer_with_attributes(Ipopt.Optimizer, "tol" => 1e-6, "print_level"=>1)
gurobi = JuMP.optimizer_with_attributes(Gurobi.Optimizer)

##########################################################################
# Creating dictionaries for RES and load time series
##########################################################################
selected_timesteps_RES_time_series = Dict{String,Any}()
selected_timesteps_load_time_series = Dict{String,Any}()
result_timesteps = Dict{String,Any}()

#MAX LOAD/RES: 475
#max LOAD ratio: 1124 
#max congestion: 1258
#Min RES: 6363

timesteps = ["475"]
#timesteps = ["475","6363"]
#timesteps = ["475","1124","1258","6363"]

#timesteps = collect(1:8760)

for l in timesteps
    if typeof(l) == String 
        selected_timesteps_RES_time_series["$l"] = Dict{String,Any}()
        for i in keys(RES_time_series)
            selected_timesteps_RES_time_series["$l"]["$i"] = Dict{String,Any}()
            selected_timesteps_RES_time_series["$l"]["$i"]["name"] = deepcopy(RES_time_series["$i"]["name"])
            selected_timesteps_RES_time_series["$l"]["$i"]["time_series"] = deepcopy(RES_time_series["$i"]["time_series"][parse(Int64,l)])
        end
    elseif typeof(l) == Int64 
        selected_timesteps_RES_time_series["$l"] = Dict{String,Any}()
        for i in keys(RES_time_series)
            selected_timesteps_RES_time_series["$l"]["$i"] = Dict{String,Any}()
            selected_timesteps_RES_time_series["$l"]["$i"]["name"] = deepcopy(RES_time_series["$i"]["name"])
            selected_timesteps_RES_time_series["$l"]["$i"]["time_series"] = deepcopy(RES_time_series["$i"]["time_series"][l])
        end
    end
end

for l in timesteps
    if typeof(l) == String 
        selected_timesteps_load_time_series["$l"] = Dict{String,Any}()
        for i in keys(load_time_series)
            selected_timesteps_load_time_series["$l"]["$i"] = Dict{String,Any}()
            selected_timesteps_load_time_series["$l"]["$i"]["time_series"] = deepcopy(load_time_series["$i"][parse(Int64,l)])
        end
    elseif typeof(l) == Int64 
        selected_timesteps_load_time_series["$l"] = Dict{String,Any}()
        for i in keys(load_time_series)
            selected_timesteps_load_time_series["$l"]["$i"] = Dict{String,Any}()
            selected_timesteps_load_time_series["$l"]["$i"]["time_series"] = deepcopy(load_time_series["$i"][l])
        end
    end
end


#####################################################################
#Only for testing purposes - to rapidly adjust infrastructure values#
#####################################################################
#=x=1
y=1
ps=[("19",220.0*x*y),("20",520.0*x*y),("21",1090.0*x*y),("22",860.0*x*y),("23",1520.0*x*y),("24",620.0*x*y)]

for p in ps
    test_case["gen"][first(p)]["pmax"]=last(p)

    test_case["gen"][first(p)]["pmin"]=0.0

    test_case["gen"][first(p)]["qmax"]=test_case["gen"][first(p)]["pmax"]/2

    test_case["gen"][first(p)]["qmin"]=-1*test_case["gen"][first(p)]["pmax"]/2
    #println(p)

    #println(test_case["gen"][first(p)]["pmax"])


end

for p in ["1","2","7","8","13","14","19","20"]

    test_case["gen"][p]["qmax"]=0.0

    test_case["gen"][p]["qmin"]=0.0

end=#

########################################################################
########################################################################
########################################################################

result_ac, demand_series = solve_opf_timestep(test_case,selected_timesteps_RES_time_series,selected_timesteps_load_time_series,timesteps,conv_power)

#save("C:\\Users\\shardy\\Documents\\julia\\times_series_input_large_files\\supernode\\HVDC_HVAC_AC_powerflows_6GW.jld2",result_ac)

#result_ac=load("C:\\Users\\shardy\\Documents\\julia\\times_series_input_large_files\\supernode\\HVDC_HVAC_AC_powerflows_6GW.jld2")


############################################################################################################
# printing the csv files for strathclyde
############################################################################################################

#Generator description
P_max=[test_case["gen"][string(g)]["pmax"] for g in sort(parse.(Int64,keys(last(first(result_ac))["solution"]["gen"])))]
Q_max=[test_case["gen"][string(g)]["qmax"] for g in sort(parse.(Int64,keys(last(first(result_ac))["solution"]["gen"])))]
name=[test_case["gen"][string(g)]["name"] for g in sort(parse.(Int64,keys(last(first(result_ac))["solution"]["gen"])))]
type=[test_case["gen"][string(g)]["type"] for g in sort(parse.(Int64,keys(last(first(result_ac))["solution"]["gen"])))]
bus=[test_case["gen"][string(g)]["gen_bus"] for g in sort(parse.(Int64,keys(last(first(result_ac))["solution"]["gen"])))]
Gen=[string(g) for g in sort(parse.(Int64,keys(last(first(result_ac))["solution"]["gen"])))]
gens=DataFrame(:Gen=>Gen,:bus=>bus,:type=>type,:name=>name,:P_max=>P_max,:Q_max=>Q_max)
CSV.write("results//AC_gen.csv", gens)

#DCtl description
P_max=[test_case["branchdc"][string(b)]["rateA"] for b in sort(parse.(Int64,keys(last(first(result_ac))["solution"]["branchdc"])))]
t_bus=[test_case["branchdc"][string(b)]["tbusdc"] for b in sort(parse.(Int64,keys(last(first(result_ac))["solution"]["branchdc"])))]
f_bus=[test_case["branchdc"][string(b)]["fbusdc"] for b in sort(parse.(Int64,keys(last(first(result_ac))["solution"]["branchdc"])))]
type=[test_case["branchdc"][string(b)]["type"] for b in sort(parse.(Int64,keys(last(first(result_ac))["solution"]["branchdc"])))]
Ldc=[b for b in sort(parse.(Int64,keys(last(first(result_ac))["solution"]["branchdc"])))]
br_r=[test_case["branchdc"][string(b)]["r"] for b in sort(parse.(Int64,keys(last(first(result_ac))["solution"]["branchdc"])))]
DCtl=DataFrame(:Ldc=>Ldc,:type=>type,:f_bus=>f_bus,:t_bus=>t_bus,:capacity=>P_max,:br_r=>br_r)
CSV.write("results//DC_tl.csv", DCtl)

#ACtl description  
P_max=[test_case["branch"][string(b)]["rate_a"] for b in sort(parse.(Int64,keys(last(first(result_ac))["solution"]["branch"])))]
t_bus=[test_case["branch"][string(b)]["t_bus"] for b in sort(parse.(Int64,keys(last(first(result_ac))["solution"]["branch"])))]
f_bus=[test_case["branch"][string(b)]["f_bus"] for b in sort(parse.(Int64,keys(last(first(result_ac))["solution"]["branch"])))]
type=[test_case["branch"][string(b)]["type"] for b in sort(parse.(Int64,keys(last(first(result_ac))["solution"]["branch"])))]
Lac=[b for b in sort(parse.(Int64,keys(last(first(result_ac))["solution"]["branch"])))]
br_r=[test_case["branch"][string(b)]["br_r"] for b in sort(parse.(Int64,keys(last(first(result_ac))["solution"]["branch"])))]
br_x=[test_case["branch"][string(b)]["br_x"] for b in sort(parse.(Int64,keys(last(first(result_ac))["solution"]["branch"])))]
b_to=[test_case["branch"][string(b)]["b_to"] for b in sort(parse.(Int64,keys(last(first(result_ac))["solution"]["branch"])))]
b_fr=[test_case["branch"][string(b)]["b_fr"] for b in sort(parse.(Int64,keys(last(first(result_ac))["solution"]["branch"])))]
ACtl=DataFrame(:Lac=>Lac,:type=>type,:f_bus=>f_bus,:t_bus=>t_bus,:capacity=>P_max,:br_r=>br_r,:br_x=>br_x,:b_to=>b_to,:b_fr=>b_fr)
CSV.write("results//AC_tl.csv", ACtl)

#DCconv description
P_max=[test_case["convdc"][string(c)]["Pacmax"] for c in sort(parse.(Int64,keys(last(first(result_ac))["solution"]["convdc"])))]
Q_max=[test_case["convdc"][string(c)]["Qacmax"] for c in sort(parse.(Int64,keys(last(first(result_ac))["solution"]["convdc"])))]
busdc_i=[test_case["convdc"][string(c)]["busdc_i"] for c in sort(parse.(Int64,keys(last(first(result_ac))["solution"]["convdc"])))]
busac_i=[test_case["convdc"][string(c)]["busac_i"] for c in sort(parse.(Int64,keys(last(first(result_ac))["solution"]["convdc"])))]
Conv=[c for c in sort(parse.(Int64,keys(last(first(result_ac))["solution"]["convdc"])))]
DCconv=DataFrame(:Conv=>Conv,:busdc_i=>busdc_i,:busac_i=>busac_i,:P_max=>P_max,:Q_max=>Q_max)
CSV.write("results//DC_conv.csv", DCconv)



#generation scenarios
gen_df=DataFrame(:time_step=>[string(i) for i=1:1:length(result_ac)])
TS=[i for i in sort(parse.(Int64,keys(result_ac)))]

for g in sort(parse.(Int64,keys(last(first(result_ac))["solution"]["gen"])))
    pg=[result_ac[string(ts)]["solution"]["gen"][string(g)]["pg"] for ts in TS]
    qg=[result_ac[string(ts)]["solution"]["gen"][string(g)]["qg"] for ts in TS]
    gen_df[Symbol("pg_"*string(g))]=deepcopy(pg)
    gen_df[Symbol("qg_"*string(g))]=deepcopy(qg)
end

for l in sort(parse.(Int64,keys(last(first(demand_series))["load"])))
    pd=[demand_series[string(ts)]["load"][string(l)]["pd"] for ts in TS]
    qd=[demand_series[string(ts)]["load"][string(l)]["qd"] for ts in TS]
    gen_df[Symbol("pd_"*string(l))]=deepcopy(pd)
    gen_df[Symbol("qd_"*string(l))]=deepcopy(qd)
end

CSV.write("results//scenarios_gen.csv", gen_df)

#converter values per scenario
convdc_df=DataFrame(:time_step=>[string(i) for i=1:1:length(result_ac)])
TS=[i for i in sort(parse.(Int64,keys(result_ac)))]

for c in sort(parse.(Int64,keys(last(first(result_ac))["solution"]["convdc"])))
    pac=[result_ac[string(ts)]["solution"]["convdc"][string(c)]["pgrid"] for ts in TS]
    pdc=[result_ac[string(ts)]["solution"]["convdc"][string(c)]["pdc"] for ts in TS]
    convdc_df[Symbol("pac_"*string(c))]=deepcopy(pac)
    convdc_df[Symbol("pdc_"*string(c))]=deepcopy(pdc)
end

CSV.write("results//scenarios_convdc.csv", convdc_df)

#cable values per scenario
cables_df=DataFrame(:time_step=>[string(i) for i=1:1:length(result_ac)])
TS=[i for i in sort(parse.(Int64,keys(result_ac)))]

#ac
for b in sort(parse.(Int64,keys(last(first(result_ac))["solution"]["branch"])))
    pf=[result_ac[string(ts)]["solution"]["branch"][string(b)]["pf"] for ts in TS]
    pt=[result_ac[string(ts)]["solution"]["branch"][string(b)]["pt"] for ts in TS]
    qf=[result_ac[string(ts)]["solution"]["branch"][string(b)]["qf"] for ts in TS]
    qt=[result_ac[string(ts)]["solution"]["branch"][string(b)]["qt"] for ts in TS]
    cables_df[Symbol("pf_ac_"*string(b-4))]=deepcopy(pf)
    cables_df[Symbol("pt_ac_"*string(b-4))]=deepcopy(pt)
    cables_df[Symbol("qf_ac_"*string(b-4))]=deepcopy(qf)
    cables_df[Symbol("qt_ac_"*string(b-4))]=deepcopy(qt)
end

#dc
for b in sort(parse.(Int64,keys(last(first(result_ac))["solution"]["branchdc"])))
    pf=[result_ac[string(ts)]["solution"]["branchdc"][string(b)]["pf"] for ts in TS]
    pt=[result_ac[string(ts)]["solution"]["branchdc"][string(b)]["pt"] for ts in TS]
    cables_df[Symbol("pf_dc_"*string(b))]=deepcopy(pf)
    cables_df[Symbol("pt_dc_"*string(b))]=deepcopy(pt)
end

CSV.write("results//scenarios_tls.csv", cables_df)

#voltage angles per scenario
for b in sort(parse.(Int64,keys(last(first(result_ac))["solution"]["bus"])))
    vm=[result_ac[string(ts)]["solution"]["bus"][string(b)]["vm"] for ts in TS]
    va=[result_ac[string(ts)]["solution"]["bus"][string(b)]["va"] for ts in TS]
    cables_df[Symbol("vm_bus_"*string(b))]=deepcopy(vm)
    cables_df[Symbol("va_bus_"*string(b))]=deepcopy(va)
end

CSV.write("results//scenarios_angles.csv", cables_df)