# Script to run the OPF simulations for the DC Grid overlay project using CbaOPF
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
s = Dict("output" => Dict("branch_flows" => true), "conv_losses_mp" => true)

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
conv_power = 8.0
test_case_file = "DC_overlay_grid_$(conv_power)_GW_convdc.json"
test_case = _PM.parse_file("./test_cases/$test_case_file")
[delete!(test_case["bus"],string(s)) for s in 1:1:2]
[delete!(test_case["bus"],string(s)) for s in 4:1:5]

#g0=deepcopy(test_case["gen"]["21"])
#g1=deepcopy(test_case["gen"]["24"])
#test_case["gen"]=Dict{String,Any}()
#g0["cost"][1]=10000.0
#g1["cost"][1]=10000.0
#push!(test_case["gen"],"21"=>g0)
#push!(test_case["gen"],"24"=>g1)

d0=deepcopy(test_case["load"]["3"])
d1=deepcopy(test_case["load"]["6"])
test_case["load"]=Dict{String,Any}()
push!(test_case["load"],"3"=>d0)
push!(test_case["load"],"6"=>d1)

#HVDC
[delete!(test_case["convdc"],string(s)) for s in 1:1:2]
[delete!(test_case["convdc"],string(s)) for s in 4:1:5]
[delete!(test_case["convdc"],string(s)) for s in 7:1:26]

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

result_timesteps_ac_2["solution"]["convdc"]
result_timesteps_ac_2["solution"]["branchdc"]
result_timesteps_ac_2["solution"]["branch"]
result_timesteps_ac_2["solution"]["gen"]

test_case["load"]["6"]
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

result_timesteps_ac_2 = deepcopy(_PMACDC.run_acdcopf(test_case, ACPPowerModel, ipopt, setting = s))
#########################################################################################################################
#########################################################################################################################
######################################
#########################################################################################################################
#########################################################################################################################
conv_power = 8.0
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
#timesteps = ["1","476", "2511", "2723","6311", "1125"]
timesteps = collect(1:8760)
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

result_ac = solve_opf_timestep(test_case,selected_timesteps_RES_time_series,selected_timesteps_load_time_series,timesteps,conv_power)
save("C:\\Users\\shardy\\Documents\\julia\\times_series_input_large_files\\supernode\\HVDC_HVAC_AC_powerflows3.jld2",result_ac)


br="8"
branchflows=[(k,maximum([sqrt(result_ac[k]["solution"]["branch"][br]["pf"]^2+
result_ac[k]["solution"]["branch"][br]["qf"]^2),
sqrt(result_ac[k]["solution"]["branch"][br]["pt"]^2+
result_ac[k]["solution"]["branch"][br]["qt"]^2)])) for k in keys(result_ac)]

#actual congestion
v=filter(x->x>57.9, last.(branchflows))
branchflows[argmax(last.(branchflows))]#5: 27 (>26.9) (2839)
branchflows[argmax(last.(branchflows))]#6: 27 (>26.3) (2672) 
branchflows[argmax(last.(branchflows))]#7: 58 (>57.9) (2508)
###################################################################################################
#branchflows[argmax(last.(branchflows))]#8: 7894, 20 - 19.9=>>57.9 (2153) -> (2306)/ 20 (>10) (1754) 5-6
##################################################################################################
br="1"
branchflows=[(k,maximum([result_ac[k]["solution"]["branchdc"][br]["pf"],
result_ac[k]["solution"]["branchdc"][br]["pt"]])) for k in keys(result_ac)]

v=filter(x->x>7.4, last.(branchflows))
####################################
#1: 30 7.5 (2046) 1-3 (7.5?)
#########################################
#2: 100 99.9 (3888)
#3: 80 79.9 (3776)
#4: 100 99.9 (2725)
#5: 80 79.9 (3103)
#6: 100 99.9 (3492)
#7: 100 99.9 (4065)
###########################################
#8: 20 5.9 (1724) 5-6 (5?)
#############################################
#9: 45 44.9 (3660)
#10: 40 39.9 (5594)
#11: 60 59.9 (2993)
#12: 60 59.9 (3939)

#congestion goal: 0.8*8760 (approx 7000)
#5 @
test_case["branchdc"]["1"]
timesteps = ["476", "2511", "2723","6311", "1125"]
##########################################################################
#result_dc["6541"]["solution"]["branchdc"]
result_ac["2511"]["solution"]["branchdc"]
#result_dc["476"]["solution"]["branch"]
result_ac["2511"]["solution"]["branch"]

result_dc["1125"]["solution"]["bus"]["7"]


result_dc["1125"]["solution"]["gen"]
result_ac["476"]["solution"]["gen"]["20"]

test_case["gen"]["20"]

result_ac["1125"]["solution"]["bus"]["6"]

result_ac["1125"]["solution"]["convdc"]["23"]
result_ac["1125"]["solution"]["convdc"]["25"]
test_case["convdc"]["19"]
test_case["convdc"]["20"]
test_case["convdc"]["21"]
test_case["convdc"]["22"]
test_case["convdc"]["23"]
test_case["convdc"]["24"]
test_case["convdc"]["25"]
test_case["convdc"]["26"]

test_case["branch"]["8"]
test_case["branchdc"]["10"]
test_case["branchdc"]["11"]
test_case["branchdc"]["12"]

test_case["gen"]["21"]

test_case["branch"]
result_dc["1125"]["solution"]["branchdc"]
result_ac["1125"]["solution"]["branchdc"]

result_timesteps_dc["1125"]["solution"]["branch"]
result_timesteps_ac["1125"]["solution"]["branch"]





function solve_opf_timestep(data,RES,load,timesteps,conv_power;output_filename::String = "/Users/giacomobastianel/Library/CloudStorage/OneDrive-KULeuven/DC_grid_overlay_results/results")
    #function solve_opf_timestep(data,RES,load,timesteps,conv_power;output_filename::String = "./results/OPF_results_selected_timesteps")
    result_timesteps_dc = Dict{String,Any}()
    result_timesteps_ac = Dict{String,Any}()
    #tcs=[]
    for t in timesteps
        println(t)
        test_case_timestep = deepcopy(data)
        res_total=Dict(1=>0.0,2=>0.0,3=>0.0,4=>0.0,5=>0.0,6=>0.0)
        load_total=Dict(1=>0.0,2=>0.0,3=>0.0,4=>0.0,5=>0.0,6=>0.0)
        for (g_id,g) in test_case_timestep["gen"]
            if (g["type"] != "Conventional" && g["type"] != "VOLL")
                res_total[g["gen_bus"]]=res_total[g["gen_bus"]]+g["pmax"]*RES["$t"][g_id]["time_series"]
                g["pmax"] = deepcopy(g["pmax"]*RES["$t"][g_id]["time_series"])
                g["qmax"] = deepcopy(g["qmax"]*RES["$t"][g_id]["time_series"]) 
            end
        end
        for (l_id,l) in test_case_timestep["load"]
            load_total[l["load_bus"]]=load_total[l["load_bus"]]+load["$t"]["Bus_"*l_id]["time_series"]*l["cosphi"]
            l["pd"] = deepcopy(load["$t"]["Bus_"*l_id]["time_series"]*l["cosphi"])
            l["qd"] = deepcopy(load["$t"]["Bus_"*l_id]["time_series"]*sqrt(1-(l["cosphi"])^2))
        end
        for (g_id,g) in test_case_timestep["gen"]
            if (g["type"] == "Conventional")
                ratio=load_total[g["gen_bus"]]/res_total[g["gen_bus"]]
                cost=100+ratio*10
                #println(cost)
                g["cost"]=[cost,0]
            end
        end
        #push!(tcs,test_case_timestep)
        #result_timesteps_dc["$t"] = deepcopy(_PMACDC.run_acdcopf(test_case_timestep, DCPPowerModel, gurobi; setting = s))
        result_timesteps_ac["$t"] = deepcopy(_PMACDC.run_acdcopf(test_case_timestep, ACPPowerModel, ipopt; setting = s))
    end

    #=string_data = JSON.json(result_timesteps_dc)
    open(output_filename*"_DCPPowerModel_$(length(timesteps))_timesteps_$(conv_power)_GW_convdc.json","w" ) do f
        write(f,string_data)
    end

    string_data = JSON.json(result_timesteps_ac)
    open(output_filename*"_ACPPowerModel_$(length(timesteps))_timesteps_$(conv_power)_GW_convdc.json","w" ) do f
        write(f,string_data)
    end=#
    return result_timesteps_ac#, tcs
end



string_data = JSON.json(result_timesteps_dc)
    open(output_filename*"_DCPPowerModel_$(length(timesteps))_timesteps_$(conv_power)_GW_convdc.json","w" ) do f
        write(f,string_data)
    end

    string_data = JSON.json(result_timesteps_ac)
    open(output_filename*"_ACPPowerModel_$(length(timesteps))_timesteps_$(conv_power)_GW_convdc.json","w" ) do f
        write(f,string_data)
    end

######################################### topology check


function show_topo(test_case)
    map=map4topo_display()


    every=Array{GenericTrace{Dict{Symbol,Any}},1}()
    ac_buses = scatter(;x=[first(v) for (k,v) in map["busac_i"]], y=[last(v) for (k,v) in map["busac_i"]],mode="marker+text", textfont_color="red",textposition="top right",text=[k for (k,v) in map["busac_i"]])
    push!(every, ac_buses)

    dc_busesl = scatter(;x=[first(v) for (k,v) in map["busdc_i"] if (parse(Int,k)<9)], y=[last(v) for (k,v) in map["busdc_i"] if (parse(Int,k)<9)], marker_color="black",mode="marker+text", textfont_color="black",textposition="bottom left", marker_size=40,text=[k for (k,v) in map["busdc_i"] if (parse(Int,k)<9)])
    push!(every, dc_busesl)
    dc_busesr = scatter(;x=[first(v) for (k,v) in map["busdc_i"]  if (parse(Int,k)>8)], y=[last(v) for (k,v) in map["busdc_i"]  if (parse(Int,k)>8)], marker_color="black",mode="marker+text", textfont_color="black",textposition="bottom right", marker_size=30,text=[k for (k,v) in map["busdc_i"] if (parse(Int,k)>8)])
    push!(every, dc_busesr)

    for (k,v) in test_case["branch"]
        acb=scatter(;x=[first(map["busac_i"][string(v["f_bus"])]),first(map["busac_i"][string(v["t_bus"])])], y=[last(map["busac_i"][string(v["f_bus"])]),last(map["busac_i"][string(v["t_bus"])])], line_color="red", mode="lines")  
        push!(every,acb)
    end

    for (k,v) in test_case["branchdc"]
        
            dcb=scatter(;x=[first(map["busdc_i"][string(v["fbusdc"])]),first(map["busdc_i"][string(v["tbusdc"])])], y=[last(map["busdc_i"][string(v["fbusdc"])]), last(map["busdc_i"][string(v["tbusdc"])])], line_color="black", mode="lines")

        push!(every,dcb)
    end

    for (k,v) in test_case["convdc"]
        acdc=scatter(;x=[first(map["busac_i"][string(v["busac_i"])]),first(map["busdc_i"][string(v["busdc_i"])])], y=[last(map["busac_i"][string(v["busac_i"])]),last(map["busdc_i"][string(v["busdc_i"])])], line_color="blue", mode="lines")  
        push!(every,acdc)
    end
        
    plot(every)
end

function map4topo_display()
    map=Dict("busac_i"=>Dict(),"busdc_i"=>Dict())
    map["busac_i"]["1"]=(1.1,2.1)
    map["busac_i"]["2"]=(2.3,3.1)
    map["busac_i"]["3"]=(2.2,2.1)
    map["busac_i"]["4"]=(2.1,1.1)
    map["busac_i"]["5"]=(4.1,1.1)
    map["busac_i"]["6"]=(4.2,2.1)
    map["busdc_i"]["1"]=(1,12)
    map["busdc_i"]["2"]=(2,13)
    map["busdc_i"]["3"]=(2,12)
    map["busdc_i"]["4"]=(2,11)
    map["busdc_i"]["5"]=(4,11)
    map["busdc_i"]["6"]=(4,12)

    map["busdc_i"]["7"]=(1,1.6)
    map["busdc_i"]["8"]=(2,1)
    map["busdc_i"]["9"]=(0.8,2)
    map["busdc_i"]["10"]=(2,2)
    map["busdc_i"]["11"]=(1,2.4)
    map["busdc_i"]["12"]=(2,3)
    map["busdc_i"]["13"]=(2,2.8)
    map["busdc_i"]["14"]=(2,2.2)
    return map
end    