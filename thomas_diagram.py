import os
import json
import sys
import matplotlib.pyplot as plt
import pandas as pd
import numpy as np
import xlwings as xw

base_path = ""
font_size = 10

color_table = ('red','blue','green','orange','purple')

def get_scaler(axis):
    if axis.get('scaler') != None:
        return axis['scaler']
    return 1

def get_inputs_x(curves, y_data):
    d = dict()
    path_book = base_path + curves['book']
    # read excel file
    try:
        if os.path.splitext(path_book)[-1] == '.csv':
            df = pd.read_csv(path_book)
        else:
            df = pd.read_excel(path_book,sheet_name = [curves['sheet']])[curves['sheet']]
    except:
        print ("inputs excel file or sheet not exist.")
        os._exit(0)
        
    try:
        x_data=df.loc[curves['x_axis']['row_start']:curves['x_axis']['row_end'],[df.columns[curves['x_axis']['column']]]].values
        for i in range(0,len(x_data)):
            x_data[i][0] = float(x_data[i][0])
        x_data = np.asarray(x_data) * get_scaler(curves['x_axis'])
        d['x'] = x_data
    except:
        print ("BOOK: \'" + curves['book'], "\',    SHEET: \'" + curves['sheet'] + "\'")
        print ("Invalid XData. They are not numeric")
        print (x_data)
        os._exit(0)

    d['y'] = y_data

    d['name'] = curves['name']
    print (d)
    return d

def get_inputs_xy(curves):
    d = dict()
    path_book = base_path + curves['book']
    # read excel file
    try:
        if os.path.splitext(path_book)[-1] == '.csv':
            df = pd.read_csv(path_book)
        else:
            df = pd.read_excel(path_book,sheet_name = [curves['sheet']])[curves['sheet']]
    except:
        print ("inputs excel file or sheet not exist.")
        os._exit(0)
        
    try:
        x_data=df.loc[curves['x_axis']['row_start']:curves['x_axis']['row_end'],[df.columns[curves['x_axis']['column']]]].values
        for i in range(0,len(x_data)):
            x_data[i][0] = float(x_data[i][0])
        x_data = np.asarray(x_data) * get_scaler(curves['x_axis'])
        d['x'] = x_data
    except:
        print ("BOOK: \'" + curves['book'], "\',    SHEET: \'" + curves['sheet'] + "\'")
        print ("Invalid XData. They are not numeric")
        print (x_data)
        os._exit(0)

    try:
        y_data=df.loc[curves['y_axis']['row_start']:curves['y_axis']['row_end'],[df.columns[curves['y_axis']['column']]]].values
        for i in range(0,len(y_data)):
            y_data[i][0] = float(y_data[i][0])
        y_data = np.asarray(y_data) * get_scaler(curves['y_axis'])
        d['y'] = y_data
    except:
        print ("BOOK: \'" + curves['book'], "\',    SHEET: \'" + curves['sheet'] + "\'")
        print ("Invalid YData. They are not numeric")
        print (y_data)
        os._exit(0)

    d['name'] = curves['name']
    print (d)
    return d

def add_to_xls(outputs, fig):
    app=xw.App(visible=False,add_book=True) #后台操作，可添加book
    try:
        bk = app.books.open(base_path + outputs['book'])
        #excel_file_exist = True
        try:
            sht = bk.sheets[outputs['sheet']]
        except:
            print ("Add sheet: '" + outputs['sheet'] + "'")
            sht = bk.sheets.add(outputs['sheet'])
    except:
        print ("Create book: \"" + base_path + outputs['book'] + "\"")
        print ("Add sheet: '" + outputs['sheet'] + "'")
        bk = app.books.add()
        sht = bk.sheets.add(outputs['sheet'])
    
    sht.pictures.add(fig,name='myplot',update=True, left=0, top=0, width=outputs['plot_width'] * outputs['layout']['ncols'], height=outputs['plot_height'] * outputs['layout']['nrows'])
    bk.save(base_path + outputs['book'])
    bk.close()
        
def get_axis(_curves):
    _min = 999999
    _max = -999999
    for _curve in _curves:
        for i in _curve['x']:
            if _min > i[0]:
                _min = i[0]
            if _max < i[0]:
                _max = i[0]
    _div = (_max - _min)/6.0
    _div = float(format(_div, '.2g'))#round(_div,2)
    _min = float(format(_min, '.2g'))#round(_div,2)
    _min = float(format(_min - _div, '.2g'))#round((_min - _div),2)
    _max = _min + 8*_div
    return np.arange(_min,_max,_div)

def draw_plots(outputs, inputs):
    i = 1
    _nrows = outputs['layout']['nrows']
    _ncols = outputs['layout']['ncols']
    #pixel dpi
    fig = plt.figure(figsize=(15,15),dpi=600)
    for _plots in inputs:
        axes = plt.subplot(_nrows, _ncols, i, frameon = False)
        plt.title(_plots['title'])
        plt.xlabel(_plots['xlabel'], fontsize=font_size)
        plt.ylabel(_plots['ylabel'], fontsize=font_size)
        plt.xticks(get_axis(_plots['curves']))
        #plt.xticks(np.arange(0.1,0.4,0.02))
        plt.tick_params(labelsize=font_size-2)
        j = 0
        lgnd = list()
        for _curve in _plots['curves']:
            #plt.xticks(get_axis(_curve['x']))
            if _plots['type'] == 'plot' or _plots['type'] == None:
                plt.plot(_curve['x'],_curve['y'],label=_curve['name'],color=color_table[j],linestyle='-',marker='o',markersize=2,linewidth=1)
            elif _plots['type'] == 'scatter':
                plt.scatter(_curve['x'],_curve['y'],label=_curve['name'],color=color_table[j],linestyle='-',marker='o',linewidth=1)
            plt.subplots_adjust(left=None,bottom=None,right=None,top=None,wspace=outputs['layout']['space'],hspace=outputs['layout']['space'])
            lgnd.append(_curve['name'])
            j += 1
        plt.grid()
        plt.legend(lgnd,prop = {'family':'Arial','weight': 'normal','size':font_size},bbox_to_anchor = (0.5, 0), loc = 'lower left')
        i += 1
    #plt.grid()
    #plt.show()
    add_to_xls(outputs,fig)



if __name__ == "__main__":
    try:
        arg1 = sys.argv[1]
    except:
        print ("no arg1 please input the json script for this app.")
        os._exit(0)
    print("JSON: " + arg1)
    with open(arg1, 'r') as load_fp:
        l_inputs = list()
        try:
            load_dict = json.load(load_fp)
            #print(load_dict)
            base_path = load_dict['Path']
            if base_path[-1] != '//' and base_path[-1] != '/':
                base_path += "/"
            font_size = load_dict['outputs']['layout']['font_size']
            j_inputs = load_dict['inputs']
            j_outputs = load_dict['outputs']
            for j_subplots in j_inputs:
                l_plots = dict()
                l_plots['title'] = j_subplots['title']
                l_plots['xlabel'] = j_subplots['xlabel']
                l_plots['ylabel'] = j_subplots['ylabel']
                if j_subplots.get('type') != None:
                    l_plots['type'] = j_subplots['type']
                else:
                    l_plots['type'] = None
                l_plots['curves'] = list()
                if j_subplots.get('curve_y') != None:
                    y_curve = get_inputs_xy(j_subplots['curve_y'])
                    l_plots['curves'].append(get_inputs_x(j_curves,y_curve))
                else:
                    for j_curves in j_subplots['curves']:
                        l_plots['curves'].append(get_inputs_xy(j_curves))
                l_inputs.append(l_plots)
        except:
            print ("invalid json script format, please check it")
            os._exit(0)
        draw_plots(j_outputs, l_inputs)