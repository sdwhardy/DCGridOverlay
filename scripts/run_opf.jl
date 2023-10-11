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

#Simple test
##########################################################################
# Define solvers 
##########################################################################
#=s = Dict("output" => Dict("branch_flows" => true), "conv_losses_mp" => true)

ipopt = JuMP.optimizer_with_attributes(Ipopt.Optimizer, "tol" => 1e-6)

gurobi = JuMP.optimizer_with_attributes(Gurobi.Optimizer)

result_timesteps_dc= deepcopy(_PMACDC.run_acdcopf("test_cases/case5_acdc.m", DCPPowerModel, gurobi, setting = s))

result_timesteps_ac = deepcopy(_PMACDC.run_acdcopf("test_cases/case5_acdc.m", ACPPowerModel, ipopt, setting = s))

test_case_5_acdc = "case5_acdc.m"

data_file_5_acdc = "./test_cases/$test_case_5_acdc"

test_grid = _PM.parse_file("test_cases/case5_acdc.m")

_PMACDC.process_additional_data!(test_grid)

result_timesteps_dc_2= deepcopy(_PMACDC.run_acdcopf(test_grid, DCPPowerModel, gurobi, setting = s))

result_timesteps_ac_2 = deepcopy(_PMACDC.run_acdcopf(test_grid, ACPPowerModel, ipopt, setting = s))

result_timesteps_dc["objective"]==result_timesteps_dc_2["objective"]

result_timesteps_ac["objective"]==result_timesteps_ac_2["objective"]

test_grid["branch"]["1"]
##########################################################################
# SINGLE HVDC LINE + 2 Converters, gen and load opposite sides ONLY
##########################################################################
conv_power = 6.0
test_case_file = "DC_overlay_grid_$(conv_power)_GW_convdc.json"
test_case = _PM.parse_file("./test_cases/$test_case_file")
[delete!(test_case["bus"],string(s)) for s in 1:1:2]
[delete!(test_case["bus"],string(s)) for s in 4:1:5]

#=g0=deepcopy(test_case["gen"]["21"])
g1=deepcopy(test_case["gen"]["24"])
test_case["gen"]=Dict{String,Any}()
g0["cost"][1]=10000.0
g1["cost"][1]=100.0
push!(test_case["gen"],"21"=>g0)
push!(test_case["gen"],"24"=>g1)=#

d0=deepcopy(test_case["load"]["3"])
d1=deepcopy(test_case["load"]["6"])
#d0["pd"]=d0["pd"]*0.2
#d0["qd"]=d0["qd"]*0.2
#d1["pd"]=d1["pd"]*0.2
#d1["qd"]=d1["qd"]*0.2
test_case["load"]=Dict{String,Any}()
push!(test_case["load"],"3"=>d0)
push!(test_case["load"],"6"=>d1)

#HVDC
[delete!(test_case["convdc"],string(s)) for s in 1:1:2]
[delete!(test_case["convdc"],string(s)) for s in 4:1:5]
[delete!(test_case["convdc"],string(s)) for s in 7:1:30]
#test_case["convdc"]["6"]["filter"]=0
#test_case["convdc"]["3"]["filter"]=0

[delete!(test_case["branchdc"],string(s)) for s in 1:1:5]
[delete!(test_case["branchdc"],string(s)) for s in 7:1:12]


[delete!(test_case["busdc"],string(s)) for s in 1:1:2]
[delete!(test_case["busdc"],string(s)) for s in 4:1:5]

test_case["branch"]=Dict{String,Any}()

#HVAC
b=deepcopy(test_case["branch"]["7"])
test_case["branch"]=Dict{String,Any}()
push!(test_case["branch"],"7"=>b)

test_case["branch"]["7"]["b_to"]=0
test_case["branch"]["7"]["b_fr"]=0

test_case["branch"]["7"]["angmin"]=-pi
test_case["branch"]["7"]["angmax"]=pi

test_case["branchdc"]=Dict{String,Any}()
test_case["convdc"]=Dict{String,Any}()
test_case["busdc"]=Dict{String,Any}()

#Bus types
#test_case["bus"]["3"]["bus_type"]=3
#test_case["bus"]["6"]["bus_type"]=2

#output
show_topo(test_case)
s = Dict("output" => Dict("branch_flows" => true), "conv_losses_mp" => true)

result_timesteps_dc_2= deepcopy(_PMACDC.run_acdcopf(test_case, DCPPowerModel, gurobi, setting = s))

result_timesteps_ac_2 = deepcopy(_PMACDC.run_acdcopf(test_case, ACPPowerModel, ipopt, setting = s))


result_timesteps_dc_2["solution"]["convdc"]
result_timesteps_dc_2["solution"]["branchdc"]
result_timesteps_dc_2["solution"]["branch"]
result_timesteps_dc_2["solution"]["gen"]

result_timesteps_ac_2["solution"]["convdc"]["3"]
result_timesteps_ac_2["solution"]["branchdc"]
result_timesteps_ac_2["solution"]["branch"]
result_timesteps_ac_2["solution"]["gen"]

test_case["load"]["6"]
test_case["load"]["3"]

test_case["convdc"]["3"]

test_case["branchdc"]["6"]
test_case["gen"]["6"]
test_case["bus"]["2"]

#single ac line
#####################################
conv_power = 8.0
test_case_file = "DC_overlay_grid_$(conv_power)_GW_convdc.json"
test_case = _PM.parse_file("./test_cases/$test_case_file")
[delete!(test_case["bus"],string(s)) for s in 3:1:6]

test_case["branchdc"]=Dict{String,Any}()
test_case["convdc"]=Dict{String,Any}()
test_case["busdc"]=Dict{String,Any}()

g=deepcopy(test_case["gen"]["19"])
test_case["gen"]=Dict{String,Any}()
push!(test_case["gen"],"19"=>g)

d=deepcopy(test_case["load"]["2"])
d["pd"]=d["pd"]/20
d["qd"]=d["qd"]/20
#d["pmax"]=60
#d["pavg"]=30
#d["qd"]=d["qd"]/10
test_case["load"]=Dict{String,Any}()
push!(test_case["load"],"2"=>d)



result_timesteps_dc_2= deepcopy(_PMACDC.run_acdcopf(test_case, DCPPowerModel, gurobi, setting = s))

result_timesteps_ac_2 = deepcopy(_PMACDC.run_acdcopf(test_case, ACPPowerModel, ipopt, setting = s))

result_timesteps_dc_2["solution"]["branch"]
result_timesteps_dc_2["solution"]["branchdc"]

result_timesteps_ac_2["solution"]["branch"]
result_timesteps_ac_2["solution"]["branchdc"]

###############################################################################################################################################################################################
###############################################################################################################################################################################################
###############################################################################################################################################################################################
##########################################################################
# Call and parse the grid, RES and load time series 
##########################################################################
conv_power = 8.0
test_case_file = "DC_overlay_grid_$(conv_power)_GW_convdc.json"
test_case = _PM.parse_file("./test_cases/$test_case_file")

#single line
#####################################
#test_case["branch"]["5"]#3->4
#test_case["load"]["3"]#3
#test_case["gen"]["22"]#4

[delete!(test_case["branch"],s) for s in ["6","7","8"]]

[delete!(test_case["load"],s) for s in ["1","2","4","5","6"]]

[delete!(test_case["gen"],string(s)) for s in 1:1:21]

[delete!(test_case["gen"],string(s)) for s in 23:1:24]

[delete!(test_case["bus"],string(s)) for s in 1:1:2]

[delete!(test_case["bus"],string(s)) for s in 5:1:6]

test_case["convdc"]=test_case["storage"]

test_case["branchdc"]=test_case["storage"]

test_case["busdc"]=test_case["storage"]

dc=deepcopy(test_case["branch"]["5"])

test_case["branch"]["5"]=test_grid["branch"]["1"]

test_case["branch"]["5"]["rate_a"]=test_case["branch"]["5"]["rate_b"]=test_case["branch"]["5"]["rate_c"]=100000

test_case["branch"]["5"]["f_bus"]=dc["f_bus"]

test_case["branch"]["5"]["t_bus"]=dc["t_bus"]

test_case["branch"]["5"]["source_id"]=dc["source_id"]

s = Dict("output" => Dict("branch_flows" => true), "conv_losses_mp" => true)

result_timesteps_dc_2= deepcopy(_PMACDC.run_acdcopf(test_case, DCPPowerModel, gurobi, setting = s))

result_timesteps_ac_2 = deepcopy(_PMACDC.run_acdcopf(test_case, ACPPowerModel, ipopt, setting = s))=#
#########################################################################################################################
#########################################################################################################################
######################################
#########################################################################################################################
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

#[test_case["gen"][string(i)]["cost"]=test_case["gen"][string(i)]["cost"]*10 for i=19:1:24]
#timesteps = ["476", "6541", "2511", "2723","6311", "1125"]
timesteps = ["475"]
#timesteps = ["475","6363"]

#timesteps = collect(1:8760)
#timesteps = ["6541"]
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


############################################################
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

result_ac, demand_series = solve_opf_timestep(test_case,selected_timesteps_RES_time_series,selected_timesteps_load_time_series,timesteps,conv_power)

save("C:\\Users\\shardy\\Documents\\julia\\times_series_input_large_files\\supernode\\HVDC_HVAC_AC_powerflows_6GW.jld2",result_ac)

result_ac=load("C:\\Users\\shardy\\Documents\\julia\\times_series_input_large_files\\supernode\\HVDC_HVAC_AC_powerflows_6GW.jld2")

result_ac["475"]["solution"]["branchdc"][]

congestion=[]
for ts in keys(result_ac)
    p=sum([maximum([sqrt(br["pf"]^2+br["qf"]^2),sqrt(br["pt"]^2+br["qt"]^2)])  for (b,br) in result_ac[ts]["solution"]["branch"]])+
    sum([maximum([br["pt"], br["pf"]]) for (b,br) in result_ac[ts]["solution"]["branchdc"]])
    push!(congestion, (ts,p))
end

congestion[argmax(last.(congestion))]
test_case["gen"]["1"]["pmax"]

    using DataFrames, CSV 

    #Generator description
    P_max=[test_case["gen"][string(g)]["pmax"] for g in sort(parse.(Int64,keys(result_ac[string("475")]["solution"]["gen"])))]
    Q_max=[test_case["gen"][string(g)]["qmax"] for g in sort(parse.(Int64,keys(result_ac[string("475")]["solution"]["gen"])))]
    name=[test_case["gen"][string(g)]["name"] for g in sort(parse.(Int64,keys(result_ac[string("475")]["solution"]["gen"])))]
    type=[test_case["gen"][string(g)]["type"] for g in sort(parse.(Int64,keys(result_ac[string("475")]["solution"]["gen"])))]
    bus=[test_case["gen"][string(g)]["gen_bus"] for g in sort(parse.(Int64,keys(result_ac[string("475")]["solution"]["gen"])))]
    Gen=[string(g) for g in sort(parse.(Int64,keys(result_ac[string("475")]["solution"]["gen"])))]
    gens=DataFrame(:Gen=>Gen,:bus=>bus,:type=>type,:name=>name,:P_max=>P_max,:Q_max=>Q_max)
    CSV.write("results//AC_gen.csv", gens)

    #DCtl description
    P_max=[test_case["branchdc"][string(b)]["rateA"] for b in sort(parse.(Int64,keys(result_ac[string("475")]["solution"]["branchdc"])))]
    t_bus=[test_case["branchdc"][string(b)]["tbusdc"] for b in sort(parse.(Int64,keys(result_ac[string("475")]["solution"]["branchdc"])))]
    f_bus=[test_case["branchdc"][string(b)]["fbusdc"] for b in sort(parse.(Int64,keys(result_ac[string("475")]["solution"]["branchdc"])))]
    type=[test_case["branchdc"][string(b)]["type"] for b in sort(parse.(Int64,keys(result_ac[string("475")]["solution"]["branchdc"])))]
    Ldc=[b for b in sort(parse.(Int64,keys(result_ac[string("475")]["solution"]["branchdc"])))]
    br_r=[test_case["branchdc"][string(b)]["r"] for b in sort(parse.(Int64,keys(result_ac[string("475")]["solution"]["branchdc"])))]
    DCtl=DataFrame(:Ldc=>Ldc,:type=>type,:f_bus=>f_bus,:t_bus=>t_bus,:capacity=>P_max,:br_r=>br_r)
    CSV.write("results//DC_tl.csv", DCtl)

    #ACtl description  
    P_max=[test_case["branch"][string(b)]["rate_a"] for b in sort(parse.(Int64,keys(result_ac[string("475")]["solution"]["branch"])))]
    t_bus=[test_case["branch"][string(b)]["t_bus"] for b in sort(parse.(Int64,keys(result_ac[string("475")]["solution"]["branch"])))]
    f_bus=[test_case["branch"][string(b)]["f_bus"] for b in sort(parse.(Int64,keys(result_ac[string("475")]["solution"]["branch"])))]
    type=[test_case["branch"][string(b)]["type"] for b in sort(parse.(Int64,keys(result_ac[string("475")]["solution"]["branch"])))]
    Lac=[b for b in sort(parse.(Int64,keys(result_ac[string("475")]["solution"]["branch"])))]
    br_r=[test_case["branch"][string(b)]["br_r"] for b in sort(parse.(Int64,keys(result_ac[string("475")]["solution"]["branch"])))]
    br_x=[test_case["branch"][string(b)]["br_x"] for b in sort(parse.(Int64,keys(result_ac[string("475")]["solution"]["branch"])))]
    b_to=[test_case["branch"][string(b)]["b_to"] for b in sort(parse.(Int64,keys(result_ac[string("475")]["solution"]["branch"])))]
    b_fr=[test_case["branch"][string(b)]["b_fr"] for b in sort(parse.(Int64,keys(result_ac[string("475")]["solution"]["branch"])))]
    ACtl=DataFrame(:Lac=>Lac,:type=>type,:f_bus=>f_bus,:t_bus=>t_bus,:capacity=>P_max,:br_r=>br_r,:br_x=>br_x,:b_to=>b_to,:b_fr=>b_fr)
    CSV.write("results//AC_tl.csv", ACtl)

    #DCconv description
    P_max=[test_case["convdc"][string(c)]["Pacmax"] for c in sort(parse.(Int64,keys(result_ac[string("475")]["solution"]["convdc"])))]
    Q_max=[test_case["convdc"][string(c)]["Qacmax"] for c in sort(parse.(Int64,keys(result_ac[string("475")]["solution"]["convdc"])))]
    busdc_i=[test_case["convdc"][string(c)]["busdc_i"] for c in sort(parse.(Int64,keys(result_ac[string("475")]["solution"]["convdc"])))]
    busac_i=[test_case["convdc"][string(c)]["busac_i"] for c in sort(parse.(Int64,keys(result_ac[string("475")]["solution"]["convdc"])))]
   #type=[test_case["convdc"][string(c)]["type"] for c in sort(parse.(Int64,keys(result_ac[string("475")]["solution"]["convdc"])))]
    Conv=[c for c in sort(parse.(Int64,keys(result_ac[string("475")]["solution"]["convdc"])))]
    DCconv=DataFrame(:Conv=>Conv,:busdc_i=>busdc_i,:busac_i=>busac_i,:P_max=>P_max,:Q_max=>Q_max)
    CSV.write("results//DC_conv.csv", DCconv)



    #generation scenarios
    #gen_df=DataFrame(:time_step=>["1","2","3","4"])
    gen_df=DataFrame(:time_step=>["1"])
    TS=[i for i in sort(parse.(Int64,keys(result_ac)))]

    for g in sort(parse.(Int64,keys(result_ac[string("475")]["solution"]["gen"])))
        pg=[result_ac[string(ts)]["solution"]["gen"][string(g)]["pg"] for ts in TS]
        qg=[result_ac[string(ts)]["solution"]["gen"][string(g)]["qg"] for ts in TS]
        gen_df[Symbol("pg_"*string(g))]=deepcopy(pg)
        gen_df[Symbol("qg_"*string(g))]=deepcopy(qg)
    end

    for l in sort(parse.(Int64,keys(demand_series[string("475")]["load"])))
        pd=[demand_series[string(ts)]["load"][string(l)]["pd"] for ts in TS]
        qd=[demand_series[string(ts)]["load"][string(l)]["qd"] for ts in TS]
        gen_df[Symbol("pd_"*string(l))]=deepcopy(pd)
        gen_df[Symbol("qd_"*string(l))]=deepcopy(qd)
    end

    CSV.write("results//scenarios_gen.csv", gen_df)

    #convdc_df=DataFrame(:time_step=>["1","2","3","4"])
    convdc_df=DataFrame(:time_step=>["1"])
    TS=[i for i in sort(parse.(Int64,keys(result_ac)))]

    for c in sort(parse.(Int64,keys(result_ac[string("475")]["solution"]["convdc"])))
        pac=[result_ac[string(ts)]["solution"]["convdc"][string(c)]["pgrid"] for ts in TS]
        pdc=[result_ac[string(ts)]["solution"]["convdc"][string(c)]["pdc"] for ts in TS]
        convdc_df[Symbol("pac_"*string(c))]=deepcopy(pac)
        convdc_df[Symbol("pdc_"*string(c))]=deepcopy(pdc)
    end

    CSV.write("results//scenarios_convdc.csv", convdc_df)

    #cables_df=DataFrame(:time_step=>["1","2","3","4"])
    cables_df=DataFrame(:time_step=>["1"])
    TS=[i for i in sort(parse.(Int64,keys(result_ac)))]

    for b in sort(parse.(Int64,keys(result_ac[string("475")]["solution"]["branch"])))
        pf=[result_ac[string(ts)]["solution"]["branch"][string(b)]["pf"] for ts in TS]
        pt=[result_ac[string(ts)]["solution"]["branch"][string(b)]["pt"] for ts in TS]
        qf=[result_ac[string(ts)]["solution"]["branch"][string(b)]["qf"] for ts in TS]
        qt=[result_ac[string(ts)]["solution"]["branch"][string(b)]["qt"] for ts in TS]
        cables_df[Symbol("pf_ac_"*string(b-4))]=deepcopy(pf)
        cables_df[Symbol("pt_ac_"*string(b-4))]=deepcopy(pt)
        cables_df[Symbol("qf_ac_"*string(b-4))]=deepcopy(qf)
        cables_df[Symbol("qt_ac_"*string(b-4))]=deepcopy(qt)
    end

    for b in sort(parse.(Int64,keys(result_ac[string("475")]["solution"]["branchdc"])))
        pf=[result_ac[string(ts)]["solution"]["branchdc"][string(b)]["pf"] for ts in TS]
        pt=[result_ac[string(ts)]["solution"]["branchdc"][string(b)]["pt"] for ts in TS]
        cables_df[Symbol("pf_dc_"*string(b))]=deepcopy(pf)
        cables_df[Symbol("pt_dc_"*string(b))]=deepcopy(pt)
    end

    CSV.write("results//scenarios_tls.csv", cables_df)


    for b in sort(parse.(Int64,keys(result_ac[string("475")]["solution"]["bus"])))
        vm=[result_ac[string(ts)]["solution"]["bus"][string(b)]["vm"] for ts in TS]
        va=[result_ac[string(ts)]["solution"]["bus"][string(b)]["va"] for ts in TS]
        cables_df[Symbol("vm_bus_"*string(b))]=deepcopy(vm)
        cables_df[Symbol("va_bus_"*string(b))]=deepcopy(va)
    end

    CSV.write("results//scenarios_angles.csv", cables_df)

test_case["gen"]["19"]