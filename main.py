import plotly.graph_objects as px
import matplotlib.pyplot as plt
import pandas as pd
import xlrd
import openpyxl
import tkinter as tk
import dash
from dash import Dash, dcc, html, Input, Output
import plotly.express as px
import numpy as np
from PIL import ImageTk, Image
import os
from threading import Timer
import webbrowser

# Create Window
root = tk.Tk()
# Set Window Size
root.geometry("1200x500")

# Set Window Title
root.title("CECS 450 - Project 2 (Visual Project)")
# Insert image into window
image = Image.open('eyetrackingImage.jpeg')
photo = ImageTk.PhotoImage(image)
label = tk.Label(root, image = photo)
label.image = photo
label.grid(column=2, row=0)

# Create a Tkinter variable
tkvar1 = tk.StringVar(root)
tkvar2 = tk.StringVar(root)

# Set up User Choices (These are the names of our Excel Sheets)
choiceList1 = ["p1","p2","p3","p4","p5","p6","p7","p10","p11","p12","p13","p14","p15","p16","p17","p18","p19","p20","p21",
           "p23","p24","p25","p27","p28","p30","p31","p32","p33","p34","p35","p36"]
# Set the default option to p1
tkvar1.set('p1')

# Set up User Choices (These are the different graphing options)
choiceList2 = ["Scatter Plot", "Heatmap", "Bubble Chart"]
# Set the default option to Scatter Plot
tkvar2.set('Scatter Plot')

# Function that creates global variable (user's choice of Sheet)
def selectSheet(value1):
    global userChoiceSheet
    userChoiceSheet = value1

# Function that creates global variable (user's choice of Graph) & destroys tkWindow when selection made
def selectGraph(value2):
    global userChoiceGraph
    userChoiceGraph = value2
    root.destroy()

# Populate Tkinter Window with both drop-down menus
popUpMenu1 = tk.OptionMenu(root, tkvar1, *choiceList1, command = selectSheet)
tk.Label(root, text="Choose a Person's Dataset to Use: ").grid(row=25, column=0)
popUpMenu1.grid(row=25, column=5)

popUpMenu2 = tk.OptionMenu(root, tkvar2, *choiceList2, command = selectGraph)
tk.Label(root, text="Choose a Graph: ").grid(row=75, column=0)
popUpMenu2.grid(row=75, column=5)

root.mainloop()

# Print User Choice (Helps for testing purposes)
# print('You have chosen %s' % userChoiceSheet)
# print('You have chosen %s' % userChoiceGraph)

# Import entire Excel Worksheet using user's choice from dropdown
wb = pd.read_excel('fxd_data_by_participants.xlsx', sheet_name=userChoiceSheet)

# Switch Case to determine which graph will be displayed
match userChoiceGraph:
    case "Heatmap":
        print("You chose Heatmap")
        import pandas as pd
        import matplotlib.pyplot as plt
        import numpy as np
        import seaborn as sns

        data_fxd = pd.read_excel('fxd_data_by_participants.xlsx')

        #Variables to pull success rate for each person
        val_tree = data_fxd.groupby(by="GraphOrTree")
        mean_val = val_tree.mean()["Success Rate"]
        mean_graph = mean_val["graph"]
        mean_tree = mean_val["tree"]
        mean_graph_and_tree = data_fxd.mean()
        mean_tot = mean_graph_and_tree["Success Rate"]

        # To show there are some fixation values out of the screen
        # print(max(data_fxd.ScreenX),max(data_fxd.ScreenY),min(data_fxd.ScreenX),min(data_fxd.ScreenY))

        # To add categories to the dataframe for each square
        def conditionsX(s):
            X = 1600
            N_square = 40
            if (s['ScreenX'] < 0 or s['ScreenX'] > X):
                return 0
            for i in range(0, N_square):
                if (s['ScreenX'] <= i * X // N_square):
                    return i


        def conditionsY(s):
            Y = 1200
            N_square = 40
            if (s['ScreenY'] < 0 or s['ScreenY'] > Y):
                return 0
            for i in range(0, N_square):
                if (s['ScreenY'] <= i * Y // N_square):
                    return i


        data_fxd["CaseX"] = data_fxd.apply(conditionsX, axis=1)
        data_fxd["CaseY"] = data_fxd.apply(conditionsY, axis=1)

        # Creating dataframes for trees, graphs and both
        data_heat_tree = data_fxd[data_fxd["GraphOrTree"] == "tree"]
        data_heat_graph = data_fxd[data_fxd["GraphOrTree"] == "graph"]

        data_heat = data_fxd[["Duration", "CaseX", "CaseY"]]
        data_heat_tree = data_heat_tree[["Duration", "CaseX", "CaseY"]]
        data_heat_graph = data_heat_graph[["Duration", "CaseX", "CaseY"]]

        data_heat = data_heat[(data_heat["CaseX"] != 0) & (data_heat["CaseY"] != 0)].groupby(by=["CaseX", "CaseY"])
        data_heat_tree = data_heat_tree[(data_heat_tree["CaseX"] != 0) & (data_heat_tree["CaseY"] != 0)].groupby(
            by=["CaseX", "CaseY"])
        data_heat_graph = data_heat_graph[(data_heat_graph["CaseX"] != 0) & (data_heat_graph["CaseY"] != 0)].groupby(
            by=["CaseX", "CaseY"])

        heat = data_heat.sum()
        heat.reset_index(inplace=True)

        heat_tree = data_heat_tree.sum()
        heat_tree.reset_index(inplace=True)

        heat_graph = data_heat_graph.sum()
        heat_graph.reset_index(inplace=True)

        x = np.full((40, 40), np.nan)
        print(x.shape)
        for index, row in heat.iterrows():
            # print(int(row['CaseX']),int(row['CaseY']))
            x[int(row['CaseX']) - 1][int(row['CaseY']) - 1] = row["Duration"]

        x_tree = np.full((40, 40), np.nan)
        for index, row in heat_tree.iterrows():
            x_tree[int(row['CaseX']) - 1][int(row['CaseY']) - 1] = row["Duration"]

        x_graph = np.full((40, 40), np.nan)
        for index, row in heat_graph.iterrows():
            x_graph[int(row['CaseX']) - 1][int(row['CaseY']) - 1] = row["Duration"]

        import dash
        from dash import Dash, dcc, html, Input, Output
        import plotly.express as px
        import numpy as np

        img = np.ones((10, 10))

        app = dash.Dash()
        app.layout = html.Div([
            dcc.Graph(id="graph1"),
            dcc.Checklist(
                id='mod',
                options=["graph", "tree"],
                value=["graph", "tree"],
            )
        ])

        @app.callback(
            Output("graph1", "figure"),
            Input("mod", "value"))
        def filter_heatmap(cols):
            if (cols == ["tree"]):
                fig = px.imshow(np.fliplr(x_tree).transpose(), color_continuous_scale='agsunset_r', title="Heatmap - " + userChoiceSheet + "  Success Rate: " + str(100 * round(mean_tree, 2)) + "%")
                fig.update_layout(width=1300, height=700)
            elif (cols == ["graph"]):
                fig = px.imshow(np.fliplr(x_graph).transpose(), color_continuous_scale='agsunset_r', title="Heatmap - " + userChoiceSheet + "  Success Rate: " + str(100 * round(mean_graph, 2)) + "%")
                fig.update_layout(width=1300, height=700)
            elif ("graph" in cols and "tree" in cols):
                fig = px.imshow(np.fliplr(x).transpose(), color_continuous_scale='agsunset_r', title="Heatmap - " + userChoiceSheet + "  Success Rate (Both Tree & Graph): " + str(100 * round(mean_tot, 2)) + "%")
                fig.update_layout(width=1300, height=700)
            else:
                fig = px.imshow(img, title="Heatmap - " + userChoiceSheet)
                fig.update_layout(width=1300, height=700)
            return fig


        def open_browser():
            webbrowser.open_new('http://127.0.0.1:8050/')

        Timer(1, open_browser).start()
        app.run_server(debug=True, use_reloader=False, port=8050)


    case "Scatter Plot":
        print("You chose Scatter Plot")
        scatterPlot = px.scatter(wb, x='ScreenX', y='ScreenY', title="Scatter Plot - " + userChoiceSheet, color='GraphOrTree', hover_data=['Success Rate', 'Participant'])
        scatterPlot.show()

    case "Bubble Chart":
        print("You chose Bubble Chart")
        bubbleChart = px.scatter(wb, x="ScreenX", y="ScreenY", title="Bubble Chart - " + userChoiceSheet, color="GraphOrTree", size="Duration", hover_data=['Success Rate', 'Participant'])
        bubbleChart.show()

    case _:
        print("Error: No matching graphing option exists!!!")
