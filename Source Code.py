import os
from re import sub
import tkinter
from tkinter import filedialog

import mhrc.automation
import numpy as np
import pandapower as pp
import pandas as pd
import xlsxwriter
from tktooltip import ToolTip
import networkx as nx


# initialization of all component lists
def find_components():
    global bus_list
    global wire_list
    global meter_list
    global pin_list
    global trafo_list
    global load_list
    global gen_list
    global tline_list
    global cable_list
    global cap_list

    bus_list = main.find_all("Bus")
    wire_list = main.find_all("WireOrthogonal")
    meter_list = main.find_all("master:multimeter")
    pin_list = main.find_all("master:pin")
    trafo_list = main.find_all("master:xfmr-3p2w")
    load_list = main.find_all("master:fixed_load")
    gen_list = main.find_all("master:source3") + main.find_all("master:source_3")
    tline_list = main.find_all("TLine")
    cable_list = main.find_all("Cable")
    cap_list = main.find_all("master:capacitor")


# get bus from node location
def get_bus(loc):
    try:
        return g.nodes[loc]["bus"]
    except:
        intersect(loc, loc)
        return g.nodes[loc]["bus"]


# sets bus attributes for a given node
def update_node_with_bus(node):
    connected_nodes = list(nx.node_connected_component(g, node))
    bus_name = ""
    # find one node that has a bus attribute, assign that to all other connected nodes
    for cn in connected_nodes:
        if g.nodes[cn]["bus"]:
            bus_name = g.nodes[cn]["bus"]
            break

    for cn in connected_nodes:
        g.nodes[cn]["bus"] = bus_name


# adds intersection point as  between two lines as a new node into graph and connect it to both edges
def intersect(p1, p2):
    for edge in list(g.edges):
        # vertical edge
        if edge[0][0] == edge[1][0]:
            if min(p1[0], p2[0]) <= edge[0][0] <= max(p1[0], p2[0]):
                if min(edge[0][1], edge[1][1]) <= p1[1] <= max(edge[0][1], edge[1][1]):
                    x = edge[0][0]
                    y = p1[1]
                    g.add_node((x, y), bus="")
                    if (x, y) != edge[0]: g.add_edge((x, y), edge[0])
                    if (x, y) != p1: g.add_edge((x, y), p1)

                    update_node_with_bus((x, y))

        # horizontal edge
        elif edge[0][1] == edge[1][1]:
            if min(p1[1], p2[1]) <= edge[1][1] <= max(p1[1], p2[1]):
                if min(edge[0][0], edge[1][0]) <= p1[0] <= max(edge[0][0], edge[1][0]):
                    x = p1[0]
                    y = edge[1][1]
                    g.add_node((x, y), bus="")
                    if (x, y) != edge[0]: g.add_edge((x, y), edge[0])
                    if (x, y) != p1: g.add_edge((x, y), p1)

                    update_node_with_bus((x, y))


# creates a graph from wires and certain components from PSCAD.
def create_network_graph():
    global g
    g = nx.Graph()

    # add wire vertices as nodes and wires as edges
    for wire in wire_list:
        for i in range(0, len(wire.vertices) - 1):
            node1 = tuple(np.add(wire.location, wire.vertices[i]))
            node2 = tuple(np.add(wire.location, wire.vertices[i + 1]))
            g.add_node(node1, bus="")
            g.add_node(node2, bus="")
            g.add_edge(node1, node2)

    # add nodes and edges for multimeters
    for meter in meter_list:
        node1, node2 = meter.get_port_location("A"), meter.get_port_location("B")
        g.add_node(node1, bus="")
        g.add_node(node2, bus="")
        g.add_edge(node1, node2)

    # add buses as nodes and edges. Search for connected nodes and assign them the proper bus attribute
    for bus in bus_list:
        # upper left end of bus
        bus_end1 = min(bus.location, tuple(np.add(bus.location, bus.vertices[1])))
        # lower right end of bus
        bus_end2 = max(bus.location, tuple(np.add(bus.location, bus.vertices[1])))

        bus_name = bus.get_parameters()["Name"]

        g.add_node(bus_end1, bus=bus_name)
        g.add_node(bus_end2, bus=bus_name)
        g.add_edge(bus_end1, bus_end2)

        intersect(bus_end1, bus_end2)

        # get all nodes that are also connected to it and set the bus tag
        connected_nodes = list(nx.node_connected_component(g, bus_end1))
        for cn in connected_nodes:
            g.nodes[cn]["bus"] = bus_name

    # check if any nodes are connected to an edge. If they are, add a new connection a one of the nodes of that edge
    for node in g.nodes:
        # node that's not connected directly to a bus
        if not g.nodes[node]["bus"]:
            intersect(node, node)

    # check if there are any master pins and add them as new nodes into the graph
    for pin in pin_list:
        x, y = pin.location
        g.add_node((x, y), bus="")
        for edge in g.edges:
            # vertical edge
            if edge[0][0] == edge[1][0] == x:
                if min(edge[0][1], edge[1][1]) <= y <= max(edge[0][1], edge[1][1]):
                    g.add_edge((x, y), edge[0])

            # horizontal edge
            elif edge[0][1] == edge[1][1] == y:
                if min(edge[0][0], edge[1][0]) <= x <= max(edge[0][0], edge[1][0]):
                    g.add_edge((x, y), edge[0])

        update_node_with_bus((x, y))


# gets pandapower bus index from name for easier referencing
def get_bus_index(name):
    return pp.get_element_index(net=net, element_type="bus", name=str(name))


# creates buses in pandapower with values from pscad
def create_buses_from_pscad():
    for bus in bus_list:
        vn_kv = float(bus.get_parameters()["BaseKV"].split("[")[0].replace(" ", ""))
        name = bus.get_parameters()["Name"]

        if sim_bus_var.get():
            index = int(sub("\D", "", name))
        else:
            index = None

        pp.create_bus(net=net, vn_kv=vn_kv, index=index, name=name)


# creates trafos in pandapower with values from pscad
def create_trafos_from_pscad():
    try:
        df = pd.read_excel(directory + "\\" + "man_input.xlsx", sheet_name="trafo")
    except FileNotFoundError:
        df = pd.DataFrame()

    for trafo in trafo_list:
        # check if a name exists, if it does use it, otherwise use cmp id
        if trafo.get_parameters()["Name"]:
            name = trafo.get_parameters()["Name"]
        else:
            name = str(trafo._id[0])

        # set winding types
        if trafo.get_parameters()["YD1"] == "0":
            winding_1 = "YN"
        else:
            winding_1 = "D"

        if trafo.get_parameters()["YD2"] == "0":
            winding_2 = "YN"
        else:
            winding_2 = "D"

        # get which winding the tap is on
        tap_winding = trafo.get_parameters()["Tap"]
        sn_mva = float(trafo.get_parameters()["Tmva"].split("[")[0].replace(" ", ""))
        v1 = float(trafo.get_parameters()["V1"].split("[")[0].replace(" ", ""))
        v2 = float(trafo.get_parameters()["V2"].split("[")[0].replace(" ", ""))

        # check which side is the low voltage side and set voltages and winding types for both sides
        if v1 < v2:
            if not df.empty:
                if not pd.isna(df.at[df[df["Name"] == name].index[0], "lv_bus"]):
                    lv_bus = get_bus_index(df.at[df[df["Name"] == name].index[0], "lv_bus"])
                else:
                    lv_bus = get_bus_index(get_bus(trafo.get_port_location("N1")))
                if not pd.isna(df.at[df[df["Name"] == name].index[0], "hv_bus"]):
                    lv_bus = get_bus_index(df.at[df[df["Name"] == name].index[0], "hv_bus"])
                else:
                    hv_bus = get_bus_index(get_bus(trafo.get_port_location("N2")))
            else:
                lv_bus = get_bus_index(get_bus(trafo.get_port_location("N1")))
                hv_bus = get_bus_index(get_bus(trafo.get_port_location("N2")))

            vn_hv_kv = v2
            vn_lv_kv = v1
            vector_group = winding_2.upper() + winding_1.lower()

            # tap
            if tap_winding == "1":
                tap_side = "lv"
            elif tap_winding == "2":
                tap_side = "hv"
            else:
                tap_side = ""

        else:
            if not df.empty:
                if not pd.isna(df.at[df[df["Name"] == name].index[0], "lv_bus"]):
                    lv_bus = get_bus_index(df.at[df[df["Name"] == name].index[0], "lv_bus"])
                else:
                    lv_bus = get_bus_index(get_bus(trafo.get_port_location("N2")))
                if not pd.isna(df.at[df[df["Name"] == name].index[0], "hv_bus"]):
                    lv_bus = get_bus_index(df.at[df[df["Name"] == name].index[0], "hv_bus"])
                else:
                    hv_bus = get_bus_index(get_bus(trafo.get_port_location("N1")))
            else:
                lv_bus = get_bus_index(get_bus(trafo.get_port_location("N2")))
                hv_bus = get_bus_index(get_bus(trafo.get_port_location("N1")))

            vn_hv_kv = v1
            vn_lv_kv = v2
            vector_group = winding_1.upper() + winding_2.lower()

            # tap
            if tap_winding == "1":
                tap_side = "hv"
            elif tap_winding == "2":
                tap_side = "lv"
            else:
                tap_side = ""

        # add hour index to vector group and set shift degree
        # YNd1
        if vector_group == "YNd" and trafo.get_parameters()["Lead"] == "1":
            vector_group += "1"
            shift_degree = 30
        # YNd11
        elif vector_group == "YNd" and trafo.get_parameters()["Lead"] == "2":
            vector_group += "11"
            shift_degree = -30
        # Dyn11
        elif vector_group == "Dyn" and trafo.get_parameters()["Lead"] == "1":
            vector_group += "11"
            shift_degree = -30
        # Dyn1
        elif vector_group == "Dyn" and trafo.get_parameters()["Lead"] == "2":
            vector_group += "1"
            shift_degree = 30
        # YNyn0 or Dd0
        else:
            vector_group += "0"
            shift_degree = 0

        vkr_percent = float(trafo.get_parameters()["CuL"].split("[")[0].replace(" ", "")) * 100
        vk_percent = float(trafo.get_parameters()["Xl"].split("[")[0].replace(" ", "")) * 100
        pfe_kw = float(trafo.get_parameters()["NLL"].split("[")[0].replace(" ", "")) * sn_mva / 1E3
        i0_percent = float(trafo.get_parameters()["Im1"].split("[")[0].replace(" ", ""))

        # use parameters from manual input sheet. If none are available, set default values

        if not df.empty:
            if not pd.isna(df.at[df[df["Name"] == name].index[0], "tap_step_percent"]):
                tap_step_percent = df.at[df[df["Name"] == name].index[0], "tap_step_percent"]
            else:
                tap_step_percent = np.nan

            if not pd.isna(df.at[df[df["Name"] == name].index[0], "tap_pos"]):
                tap_pos = df.at[df[df["Name"] == name].index[0], "tap_pos"]
            else:
                tap_pos = np.nan

            if not pd.isna(df.at[df[df["Name"] == name].index[0], "tap_neutral"]):
                tap_neutral = df.at[df[df["Name"] == name].index[0], "tap_neutral"]
            else:
                tap_neutral = np.nan

            if not pd.isna(df.at[df[df["Name"] == name].index[0], "tap_step_degree"]):
                tap_step_degree = df.at[df[df["Name"] == name].index[0], "tap_step_degree"]
            else:
                tap_step_degree = np.nan
        else:
            tap_step_percent = np.nan
            tap_pos = np.nan
            tap_neutral = np.nan
            tap_step_degree = np.nan

        pp.create_transformer_from_parameters(net=net, hv_bus=hv_bus, lv_bus=lv_bus, sn_mva=sn_mva, vn_hv_kv=vn_hv_kv,
                                              vn_lv_kv=vn_lv_kv, vkr_percent=vkr_percent, vk_percent=vk_percent,
                                              pfe_kw=pfe_kw, i0_percent=i0_percent, name=name,
                                              vector_group=vector_group, shift_degree=shift_degree, tap_side=tap_side,
                                              tap_step_percent=tap_step_percent, tap_pos=tap_pos,
                                              tap_neutral=tap_neutral, tap_step_degree=tap_step_degree)


# creates load in pandapower with parameters from pscad
def create_loads_from_pscad():
    try:
        df = pd.read_excel(directory + "\\" + "man_input.xlsx", sheet_name="load")
    except FileNotFoundError:
        df = pd.DataFrame()

    for load in load_list:
        name = int(load._id[0])
        p_mw = float(load.get_parameters()["PO"].split("[")[0].replace(" ", "")) * 3
        q_mvar = float(load.get_parameters()["QO"].split("[")[0].replace(" ", "")) * 3

        if not df.empty:
            if not pd.isna(df.at[df[df["Name"] == name].index[0], "Bus"]):
                bus = get_bus_index(df.at[df[df["Name"] == name].index[0], "Bus"])
            else:
                bus = get_bus_index(get_bus(load.get_port_location("IA")))
        else:
            bus = get_bus_index(get_bus(load.get_port_location("IA")))

        pp.create_load(net=net, bus=bus, p_mw=p_mw, q_mvar=q_mvar, name=name)


# creates generator in pandapower with parameters from pscad
def create_gens_from_pscad():
    try:
        df = pd.read_excel(directory + "\\" + "man_input.xlsx", sheet_name="gen")
    except FileNotFoundError:
        df = pd.DataFrame()

    for gen in gen_list:
        # check if a name exists, if it does use it, otherwise use cmp id
        if gen.get_parameters()["Name"]:
            name = gen.get_parameters()["Name"]
        else:
            name = str(gen._id[0])

        if not df.empty:
            # get q limits from manual input sheet
            if not pd.isna(df.at[df[df["Name"] == name].index[0], "max_q_mvar"]):
                max_q_mvar = df.at[df[df["Name"] == name].index[0], "max_q_mvar"]
            else:
                max_q_mvar = np.nan

            if not pd.isna(df.at[df[df["Name"] == name].index[0], "min_q_mvar"]):
                min_q_mvar = df.at[df[df["Name"] == name].index[0], "min_q_mvar"]
            else:
                min_q_mvar = np.nan
        else:
            max_q_mvar = np.nan
            min_q_mvar = np.nan

        # check which PSCAD definition the current generator has
        if str(gen.get_definition()) == "master:source3":
            # Base values for per unit quantities
            U_n = float(gen.get_parameters()["Vm"].split("[")[0].replace(" ", ""))
            S_n = float(gen.get_parameters()["MVA"].split("[")[0].replace(" ", ""))

            p_mw = float(gen.get_parameters()["Pinit"].split("[")[0].replace(" ", "")) * S_n
            vm_pu = float(gen.get_parameters()["Es"].split("[")[0].replace(" ", "")) / U_n
            # only used for slack bus
            va_degree = float(gen.get_parameters()["Ph"].split("[")[0].replace(" ", ""))

            if not df.empty:
                if not pd.isna(df.at[df[df["Name"] == name].index[0], "Bus"]):
                    bus = get_bus_index(df.at[df[df["Name"] == name].index[0], "Bus"])
                else:
                    bus = get_bus_index(get_bus(gen.get_port_location("N3")))
            else:
                bus = get_bus_index(get_bus(gen.get_port_location("N3")))

        elif str(gen.get_definition()) == "master:source_3":
            p_mw = float(gen.get_parameters()["Pinit"].split("[")[0].replace(" ", ""))
            vm_pu = float(gen.get_parameters()["Vpu"].split("[")[0].replace(" ", ""))
            # only used for slack bus
            va_degree = float(gen.get_parameters()["PhT"].split("[")[0].replace(" ", ""))

            if not df.empty:
                if not pd.isna(df.at[df[df["Name"] == name].index[0], "Bus"]):
                    bus = get_bus_index(df.at[df[df["Name"] == name].index[0], "Bus"])
                else:
                    bus = get_bus_index(get_bus(gen.get_port_location("N")))
            else:
                bus = get_bus_index(get_bus(gen.get_port_location("N")))

        # check if gen is connected to slack bus
        if bus == get_bus_index(slack_ent.get()):

            # get short circuit parameters for external grid from manual input sheet. Use default values if there are none
            pp.create_ext_grid(net=net, bus=bus, vm_pu=vm_pu, va_degree=va_degree, name=name, max_q_mvar=max_q_mvar,
                               min_q_mvar=min_q_mvar)

        else:
            pp.create_gen(net=net, bus=bus, p_mw=p_mw, vm_pu=vm_pu, name=name, max_q_mvar=max_q_mvar,
                          min_q_mvar=min_q_mvar)


# creates transmission lines in PandaPower with parameters from pscad; type ol = overhead line, cs = underground cable system
def create_lines_from_pscad():
    # setup directory for needed documents
    if fortran_version == "GFortran 4.2.1":
        folder = directory + "\\" + project_name + ".gf42"
    elif fortran_version == "GFortran 4.6.2":
        folder = directory + "\\" + project_name + ".gf46"

    try:
        df = pd.read_excel(directory + "\\" + "man_input.xlsx", sheet_name="line")
    except FileNotFoundError:
        df = pd.DataFrame()

    # get .dta file name
    dtafile = folder + "\\" + "main.dta"

    # defines area with lower and upper limits for node-bus allocation
    with open(folder + "\\" + "main.dta", "r") as fp:
        lines = fp.readlines()
        for line in lines:
            if line.find("! Local Node Voltages") != -1:
                lower_limit = lines.index(line)
            if line.find("! Local Branch Data") != -1:
                upper_limit = lines.index(line)

    # create lines in PandaPower with parameters from PSCAD
    for tline in tline_list:
        name = tline.get_parameters()["Name"]
        length_km = float(tline.get_parameters()["Length"].split("[")[0].replace(" ", ""))
        type = "ol"

        # get .out file name
        outfile = folder + "\\" + name + ".out"

        # read nodes from main.dta file
        with open(folder + "\\" + "main.dta", "r") as fp:
            lines = fp.readlines()
            for line in lines:
                if line.find(r"! " + name) != -1:
                    node_1 = lines[lines.index(line) + 2].split(" ")[1]
                    node_2 = lines[lines.index(line) + 3].split(" ")[1]

        # lookup nodes for buses
        with open(folder + "\\" + "main.dta", "r") as fp:
            lines = fp.readlines()
            for line in lines[lower_limit:upper_limit]:
                if line.split("0.0")[0].replace(" ", "") == node_1:
                    to_bus = get_bus_index(line.split(r"//")[1].lstrip().split("(")[0])

                if line.split("0.0")[0].replace(" ", "") == node_2:
                    from_bus = get_bus_index(line.split(r"//")[1].lstrip().split("(")[0])

        with open(outfile, "r") as fp:
            # If load_flow_data_exists is true, load flow data exists in rxb form and can be used. Otherwise it needs
            # to be extracted from the admittance/impedance matrices
            load_flow_data_exists = False
            lines = fp.readlines()

            # check if rxb values exist
            # assign values for base of per unit quantities
            for line in lines:
                if line.find("LOAD FLOW RXB FORMATTED DATA") != -1:
                    load_flow_data_exists = True
                    freq = float(line.replace(" ", "").replace("LOADFLOWRXBFORMATTEDDATA@", "").replace("Hz:", ""))
                    w = 2 * np.pi * freq
                    U_n = float(
                        lines[lines.index(line) + 3].replace(" ", "").split(",")[0].replace("BaseofPer-UnitQuantities:",
                                                                                            "").replace("kV(L-L)",
                                                                                                        "")) * 1E3
                    S_n = float(lines[lines.index(line) + 3].replace(" ", "").split(",")[1].replace("MVA", "")) * 1E6

            # read  output file from pscad to get zero and positive sequence values
        with open(outfile, "r") as fp:

            # search for zero/positive sequences in tline output files
            lines = fp.readlines()
            for line in lines:

                # find positive sequence data if rxb values exist and calculate absolute values in a format for pandapower from per unit values
                if line.strip() == "Positive Sequence" and load_flow_data_exists:
                    R_pu = float(lines[lines.index(line) + 3].replace(" ", "").replace(r"ResistanceRsq[pu]:", ""))
                    X_pu = float(lines[lines.index(line) + 4].replace(" ", "").replace(r"ReactanceXsq[pu]:", ""))
                    B_pu = float(lines[lines.index(line) + 5].replace(" ", "").replace(r"SusceptanceBsq[pu]:", ""))

                    r_ohm_per_km = R_pu * (U_n ** 2) / S_n / length_km
                    x_ohm_per_km = X_pu * (U_n ** 2) / S_n / length_km
                    c_nf_per_km = B_pu * S_n / (U_n ** 2) / length_km / w * 1E9

                # find pi component data from matrix form if rxb values do not exist
                # the admittance/impedance matrices come in the format "a,b" with a being the real part and b the imaginary part
                # split them up after the comma to assign the proper values for each parameter
                if line.strip() == r"SERIES IMPEDANCE MATRIX (Z) [ohms/m]:" and not load_flow_data_exists:
                    r_ohm_per_km = float(lines[lines.index(line) + 1].replace(" ", "").split(",")[0]) * 1E3
                    x_ohm_per_km = float(lines[lines.index(line) + 1].replace(" ", "").split(",")[1]) * 1E3
                    c_nf_per_km = float(lines[lines.index(line) + 4].replace(" ", "").split(",")[1]) / w * 1E3 * 1E9

        # read values for max_i_ka from manual input spreadsheet. If no value exist, set a default value
        if not df.empty:
            if not pd.isna(df.at[df[df["Name"] == name].index[0], "max_i_ka"]):
                max_i_ka = df.at[df[df["Name"] == name].index[0], "max_i_ka"]
            else:
                max_i_ka = 1E9
        else:
            max_i_ka = 1E9

        # create tline in pandapower
        pp.create_line_from_parameters(net=net, from_bus=from_bus, to_bus=to_bus, length_km=length_km,
                                       type=type, r_ohm_per_km=r_ohm_per_km, x_ohm_per_km=x_ohm_per_km,
                                       c_nf_per_km=c_nf_per_km, max_i_ka=max_i_ka, name=name)

    # create cables in PandaPower with parameters from PSCAD
    for cable in cable_list:
        # get definition name
        name = cable.get_parameters()["Name"]

        # get .out file name
        outfile = folder + "\\" + name + ".out"

        # assign default values
        type = "cs"
        length_km = float(cable.get_parameters()["Length"].split("[")[0].replace(" ", ""))

        # read nodes from main.dta file
        with open(dtafile, "r") as fp:
            lines = fp.readlines()
            for line in lines:
                if line.find(r"! " + name) != -1:
                    node_1 = lines[lines.index(line) + 2].split(" ")[1]
                    node_2 = lines[lines.index(line) + 3].split(" ")[1]

        # lookup nodes for buses in pre defined area
        with open(folder + "\\" + "main.dta", "r") as fp:
            lines = fp.readlines()

            for line in lines[lower_limit:upper_limit]:
                if line.split("0.0")[0].replace(" ", "") == node_1:
                    to_bus = get_bus_index(line.split(r"//")[1].lstrip().split("(")[0])

                if line.split("0.0")[0].replace(" ", "") == node_2:
                    from_bus = get_bus_index(line.split(r"//")[1].lstrip().split("(")[0])

        with open(outfile, "r") as fp:
            # If load_flow_data_exists is true, load flow data exists in rxb form and can be used. Otherwise it needs to be extracted from
            # the admittance/impedance matrices
            load_flow_data_exists = False
            lines = fp.readlines()

            # check if rxb values exist
            # assign values for base of per unit quantities
            for line in lines:
                if line.find("LOAD FLOW RXB FORMATTED DATA") != -1:
                    load_flow_data_exists = True
                    freq = float(line.replace(" ", "").replace("LOADFLOWRXBFORMATTEDDATA@", "").replace("Hz:", ""))
                    w = 2 * np.pi * freq
                    U_n = float(
                        lines[lines.index(line) + 3].replace(" ", "").split(",")[0].replace("BaseofPer-UnitQuantities:",
                                                                                            "").replace("kV(L-L)",
                                                                                                        "")) * 1E3
                    S_n = float(lines[lines.index(line) + 3].replace(" ", "").split(",")[1].replace("MVA", "")) * 1E6

                elif line.find(r"SEQUENCE COMPONENT DATA @") != -1 and not load_flow_data_exists:
                    freq = float(line.replace(" ", "").split(r"@")[1].replace("Hz:", ""))
                    w = 2 * np.pi * freq

        # read  output file from pscad to get positive sequence values
        with open(outfile, "r") as fp:

            # search for positive sequences in cable output files
            lines = fp.readlines()
            for line in lines:

                # find positive sequence data if rxb values exist and calculate absolute values in a format for pandapower from per unit values
                if line.strip() == "Positive Sequence" and load_flow_data_exists:
                    R_pu = float(lines[lines.index(line) + 3].replace(" ", "").replace(r"ResistanceRsq[pu]:", ""))
                    X_pu = float(lines[lines.index(line) + 4].replace(" ", "").replace(r"ReactanceXsq[pu]:", ""))
                    B_pu = float(lines[lines.index(line) + 5].replace(" ", "").replace(r"SusceptanceBsq[pu]:", ""))

                    r_ohm_per_km = R_pu * (U_n ** 2) / S_n / length_km
                    x_ohm_per_km = X_pu * (U_n ** 2) / S_n / length_km
                    c_nf_per_km = B_pu * S_n / (U_n ** 2) / length_km / w * 1E9

                # find pi component data from matrix form if rxb values do not exist
                # the admittance/impedance matrices come in the format "a,b" with a being the real part and b the imaginary part
                # split them up after the comma to assign the proper values for each parameter
                if line.strip() == r"SERIES IMPEDANCE MATRIX (Z) [ohms/m]:" and not load_flow_data_exists:
                    r_ohm_per_km = float(lines[lines.index(line) + 1].lstrip().split("   ")[0].split(",")[0]) * 1E3
                    x_ohm_per_km = float(lines[lines.index(line) + 1].lstrip().split("   ")[0].split(",")[1]) * 1E3

                if line.strip() == r"SHUNT ADMITTANCE MATRIX (Y) [mhos/m]:" and not load_flow_data_exists:
                    c_nf_per_km = float(
                        lines[lines.index(line) + 1].lstrip().split("   ")[0].split(",")[1]) / w * 1E3 * 1E9

        # if there are values for certain parameters, use those. If there are none, use the parameters from pscad
        if not df.empty:
            if not pd.isna(df.at[df[df["Name"] == name].index[0], "max_i_ka"]):
                max_i_ka = df.at[df[df["Name"] == name].index[0], "max_i_ka"]
            else:
                max_i_ka = 1E9
        else:
            max_i_ka = 1E9

        # create cable in pandapower
        pp.create_line_from_parameters(net=net, from_bus=from_bus, to_bus=to_bus, length_km=length_km, type=type,
                                       r_ohm_per_km=r_ohm_per_km, x_ohm_per_km=x_ohm_per_km, c_nf_per_km=c_nf_per_km,
                                       max_i_ka=max_i_ka, name=name)


# creates capacity banks in PandaPower with parameters from PSCAD
def create_cap_banks_from_pscad():
    try:
        df = pd.read_excel(directory + "\\" + "man_input.xlsx", sheet_name="cap_bank")
    except FileNotFoundError:
        df = pd.DataFrame()

    for cap in cap_list:
        name = int(cap._id[0])

        if not df.empty:
            if not pd.isna(df.at[df[df["Name"] == name].index[0], "Bus"]):
                bus = get_bus_index(df.at[df[df["Name"] == name].index[0], "Bus"])
            else:
                try:
                    bus = get_bus_index(get_bus(cap.get_port_location("A")))
                except KeyError:
                    bus = get_bus_index(get_bus(cap.get_port_location("B")))
        else:
            try:
                bus = get_bus_index(get_bus(cap.get_port_location("A")))
            except KeyError:
                bus = get_bus_index(get_bus(cap.get_port_location("B")))

        # get base voltage for reactive power calculation from connected bus
        vn_kv = net.bus["vn_kv"][bus]
        # calculate reactive power from capacity. q = wcu^2
        c = float(cap.get_parameters()["C"].split("[")[0].replace(" ", "")) * 1E-6
        w = 2 * np.pi * float(freq_ent.get())
        q_mvar = w * c * ((vn_kv * 1E3) ** 2) / 1E6

        pp.create_shunt_as_capacitor(net=net, bus=bus, q_mvar=q_mvar, vn_kv=vn_kv, loss_factor=0, name=name)


# updates generators in PSCAD with results from PandaPower load flow analysis
def update_gens_in_pscad():
    for gen in gen_list:
        # use PSCAD name, if none exist use the PSCAD id
        if gen.get_parameters()["Name"]:
            name = gen.get_parameters()["Name"]
        else:
            name = int(gen._id[0])

        # get PandaPower results for current gen, depending on whether it's represented as a generator or external grid in PandaPower
        try:
            this_gen_pp_index = pp.get_element_index(net=net, element="gen", name=name)
            pinit_mw = float(net.res_gen["p_mw"][this_gen_pp_index])
            qinit_mvar = float(net.res_gen["q_mvar"][this_gen_pp_index])
            v_pu = float(net.res_gen["vm_pu"][this_gen_pp_index])
            ph = float(net.res_gen["va_degree"][this_gen_pp_index])

        except UserWarning:
            this_ext_grid_index = pp.get_element_index(net=net, element="ext_grid", name=name)

            pinit_mw = float(net.res_ext_grid["p_mw"][this_ext_grid_index])
            qinit_mvar = float(net.res_ext_grid["q_mvar"][this_ext_grid_index])
            v_pu = float(net.ext_grid["vm_pu"][this_ext_grid_index])
            ph = float(net.ext_grid["va_degree"][this_ext_grid_index])

        if str(gen.get_definition()) == "master:source3":
            u_n_kv = float(gen.get_parameters()["Vm"].split("[")[0].replace(" ", ""))
            s_n_mva = float(gen.get_parameters()["MVA"].split("[")[0].replace(" ", ""))

            pinit_pu = pinit_mw / s_n_mva
            qinit_pu = qinit_mvar / s_n_mva
            es_kv = v_pu * u_n_kv

            gen.set_parameters(Pinit=pinit_pu, Qinit=qinit_pu, Ph=ph, Es=es_kv)

        elif str(gen.get_definition()) == "master:source_3":
            s_n_mva = float(gen.get_parameters()["Sbase"].split("[")[0].replace(" ", ""))

            pinit_pu = pinit_mw / s_n_mva
            qinit_pu = qinit_mvar / s_n_mva

            gen.set_parameters(Pinit=pinit_pu, Qinit=qinit_pu, PhT=ph, Vpu=v_pu)


def button_select_path():
    global filename
    global directory
    global project_name
    global path
    path = filedialog.askopenfilename()
    directory = os.path.dirname(path)
    filename = os.path.basename(path)
    project_name = os.path.splitext(filename)[0]


def button_run():
    global fortran_version
    global main
    global net
    fortran_version = fcomp_var.get()

    # launch pscad, silence: surpress dialogues, certificate: False = Legacy Licensing
    pscad = mhrc.automation.launch_pscad(silence=True, minimize=True)
    pscad.load(path)
    project = pscad.project(project_name)
    main = project.user_canvas("Main")

    # build project
    if build_var.get():
        pscad.settings(fortran_version=fortran_version)
        pscad.build_current()

    # create PandaPower network with given frequency
    net = pp.create_empty_network(f_hz=float(freq_ent.get()), add_stdtypes=False)

    # create component lists
    find_components()

    # create graph for electrical connections
    create_network_graph()

    # create PandaPower components
    create_buses_from_pscad()
    create_lines_from_pscad()
    create_cap_banks_from_pscad()
    create_loads_from_pscad()
    create_trafos_from_pscad()
    create_gens_from_pscad()

    # export pandapower results into excel sheet. Before the powerflow analysis so in case it gives an error you can
    # troubleshoot the inputs easier
    if pp_excel_var.get():
        pp.to_excel(net=net, filename=directory + "\\" + "pandapower_result.xlsx")

    # get max iterations for powerflow calculation from gui
    max_iteration = pp_it_ent.get()
    if max_iteration.isdigit():
        max_iteration = int(max_iteration)

    # set if q-limits should get taken into account
    enforce_q_lims = q_limit_var.get()

    # run powerflow
    pp.runpp(net=net, calculate_voltage_angles=True, init=pp_init_ent.get(), max_iteration=max_iteration,
             enforce_q_lims=enforce_q_lims)

    # export pandapower results into excel sheet. After the powerflow analysis to add results to it
    if pp_excel_var.get():
        pp.to_excel(net=net, filename=directory + "\\" + "pandapower_result.xlsx")

    # transfer powerflow results into PSCAD
    update_gens_in_pscad()
    project.save()
    print("done")


def button_create_man_inp():
    pscad = mhrc.automation.launch_pscad(silence=True, minimize=True)
    pscad.load(path)
    project = pscad.project(project_name)
    main = project.user_canvas("Main")

    workbook = xlsxwriter.Workbook(directory + "\\" + "man_input.xlsx")

    trafo_list = main.find_all("master:xfmr-3p2w")
    sheet_trafo = workbook.add_worksheet(name="trafo")
    sheet_trafo.write("A1", "Name")
    sheet_trafo.write("B1", "hv_bus")
    sheet_trafo.write("C1", "lv_bus")
    sheet_trafo.write("D1", "tap_step_percent")
    sheet_trafo.write("E1", "tap_step_degree")
    sheet_trafo.write("F1", "tap_pos")
    sheet_trafo.write("G1", "tap_neutral")

    gen_list = main.find_all("master:source3") + main.find_all("master:source_3")
    sheet_gen = workbook.add_worksheet(name="gen")
    sheet_gen.write("A1", "Name")
    sheet_gen.write("B1", "Bus")
    sheet_gen.write("C1", "max_q_mvar")
    sheet_gen.write("D1", "min_q_mvar")

    line_list = main.find_all("TLine") + main.find_all("Cable")
    sheet_line = workbook.add_worksheet(name="line")
    sheet_line.write("A1", "Name")
    sheet_line.write("B1", "max_i_ka")

    load_list = main.find_all("master:fixed_load")
    sheet_load = workbook.add_worksheet(name="load")
    sheet_load.write("A1", "Name")
    sheet_load.write("B1", "Bus")

    shunt_list = main.find_all("master:capacitor")
    sheet_cap_bank = workbook.add_worksheet(name="cap_bank")
    sheet_cap_bank.write("A1", "Name")
    sheet_cap_bank.write("B1", "Bus")

    # setup sheets with names or ids from PSCAD components
    for i, line in enumerate(line_list):
        name = line.get_parameters()["Name"]
        sheet_line.write(i + 1, 0, name)

    for i, trafo in enumerate(trafo_list):
        if trafo.get_parameters()["Name"]:
            name = trafo.get_parameters()["Name"]
        else:
            name = trafo._id[0]

        sheet_trafo.write(i + 1, 0, name)

    for i, gen in enumerate(gen_list):
        if gen.get_parameters()["Name"]:
            name = gen.get_parameters()["Name"]
        else:
            name = gen._id[0]

        sheet_gen.write(i + 1, 0, name)

    for i, load in enumerate(load_list):
        name = load._id[0]
        sheet_load.write(i + 1, 0, name)

    for i, cap_bank in enumerate(shunt_list):
        name = cap_bank._id[0]
        sheet_cap_bank.write(i + 1, 0, name)

    workbook.close()
    pscad.quit()
    print("manual input template created")


def main():
    global pp_excel_var
    global pp_it_ent
    global pp_init_ent
    global freq_ent
    global fcomp_var
    global slack_ent
    global sim_bus_var
    global build_var
    global q_limit_var

    root = tkinter.Tk()
    root.title("PSCAD Loadflow initializer")
    root.geometry("650x400")
    root.grid_rowconfigure((0, 1, 2, 3, 4, 5, 6, 7, 8, 9), weight=1)
    root.grid_columnconfigure((0, 1, 2, 3), weight=1)

    # get list of all installed fortran compilers and remove all of them except the supported ones 4.6.2 and 4.2.1
    fcomp_list = mhrc.automation.controller().get_paramlist_names("fortran")
    for fcomp in fcomp_list:
        if not fcomp == "GFortran 4.6.2" and not fcomp == "GFortran 4.2.1":
            del fcomp

    fcomp_var = tkinter.StringVar(value=fcomp_list[1])
    sim_bus_var = tkinter.BooleanVar()
    pp_excel_var = tkinter.BooleanVar()
    build_var = tkinter.BooleanVar()
    q_limit_var = tkinter.BooleanVar()

    # create button which starts the program to run a powerflow analysis
    run_bt = tkinter.Button(master=root, text="Run", command=button_run)
    run_bt.grid(row=7, column=1, sticky="ew")

    # create button which creates manual input excel template
    man_inp_bt = tkinter.Button(master=root, text="Create manual input template", command=button_create_man_inp)
    man_inp_bt.grid(row=7, column=2, sticky="ew")
    ToolTip(man_inp_bt, msg="Create blank excel file for required manual inputs. Select PSCAD file first.")

    # create checkbox for selecting whether you want similiar bus names or not
    sim_bus_cb = tkinter.Checkbutton(master=root, text="Similiar bus indices", variable=sim_bus_var, onvalue=True,
                                     offvalue=False)
    sim_bus_cb.grid(row=3, column=0, sticky="ew")
    sim_bus_cb.select()
    ToolTip(sim_bus_cb, msg="If checked, use the same bus indices for PandaPower that are used in PSCAD bus names")

    # create checkbox for making excel file from pandapower
    pp_excel_cb = tkinter.Checkbutton(master=root, text="Create Excel file", variable=pp_excel_var, onvalue=True,
                                      offvalue=False)
    pp_excel_cb.select()
    pp_excel_cb.grid(row=3, column=1, sticky="ew")
    ToolTip(pp_excel_cb, msg="Create Excel file from PandaPower data")

    # create checkbox for building project
    build_cb = tkinter.Checkbutton(master=root, text="Build project on launch", variable=build_var, onvalue=True,
                                   offvalue=False)
    build_cb.grid(row=3, column=2, sticky="ew")
    ToolTip(build_cb, msg="Build project before the powerflow calculation")

    # create checkbox for enforcing q limits
    q_lim_cb = tkinter.Checkbutton(master=root, text="Q limits", variable=q_limit_var, onvalue=True, offvalue=False)
    q_lim_cb.grid(row=3, column=3, sticky="ew")
    ToolTip(q_lim_cb, msg="Consider Q limits for generators in powerflow calculation")

    # create  entries for setting frequency for powerflow calculations
    freq_ent = tkinter.Entry(master=root, width=50)
    freq_ent.insert(0, "60")
    freq_ent.grid(row=1, column=0)
    ToolTip(freq_ent, msg="Enter frequency for powerflow calculation")
    freq_label = tkinter.Label(master=root, text="Frequency")
    freq_label.grid(row=0, column=0)

    # create entries for assigning slack bus
    slack_ent = tkinter.Entry(master=root, width=50)
    slack_ent.insert(0, "Bus1")
    slack_ent.grid(row=1, column=1)
    ToolTip(slack_ent, msg="Enter PSCAD bus name that should be treated as the slack bus")
    slack_label = tkinter.Label(master=root, text="Slack bus")
    slack_label.grid(row=0, column=1)

    # create  entries for amount of pandapower iterations
    pp_it_ent = tkinter.Entry(master=root, width=50)
    pp_it_ent.insert(0, "auto")
    pp_it_ent.grid(row=1, column=2)
    ToolTip(pp_it_ent,
            msg='Enter amount of maximal iterations for PandaPower powerflow calculation. Enter "auto" for default values')
    pp_it_label = tkinter.Label(master=root, text="LF iterations")
    pp_it_label.grid(row=0, column=2)

    # create entry for initialisation for pandapower powerflow calculation
    pp_init_ent = tkinter.Entry(master=root, width=50)
    pp_init_ent.insert(0, "auto")
    pp_init_ent.grid(row=1, column=3)
    ToolTip(pp_init_ent,
            msg='Enter initialisation type for PandaPower powerflow calculation. Options are: "auto", "flat", "dc", "results". View PandaPower documentation for further infomation')
    pp_init_label = tkinter.Label(master=root, text="LF initialization")
    pp_init_label.grid(row=0, column=3)

    # create option menu to select fortran compiler
    fcomp_om = tkinter.OptionMenu(root, fcomp_var, *fcomp_list)
    fcomp_om.grid(row=5, column=0)
    ToolTip(fcomp_om, msg="Select Fortran Compiler")
    fcomp_label = tkinter.Label(master=root, text="Compiler")
    fcomp_label.grid(row=4, column=0)

    # create button to select file path
    select_path_bt = tkinter.Button(master=root, text="Select path", command=button_select_path)
    select_path_bt.grid(row=7, column=0, sticky="ew")
    ToolTip(select_path_bt, msg="Select PSCAD file")

    root.mainloop()


if __name__ == "__main__":
    main()
