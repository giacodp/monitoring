# Libraries
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib import rc
from datetime import date
import argparse

def main():
    
    # Create command in order to pass the path file by command line
    parser = argparse.ArgumentParser()
    parser.add_argument('-p', '--path_file', type=str, default = 'report_prova.xlsx',
                        help='Path of the input file (default: report_prova.xlsx')
    parser.add_argument('-s', '--save_folder', type=str, default = '',
                        help='Save folder (default: your actual folder')
    args = parser.parse_args()

    # Data preparation
    dtf_bar, lst_pie = tempData(args)
    
    # Create a Graphic object and plot charts
    gr = Graphic(dtf_bar, lst_pie, args)
    gr.plotBarChart()
    gr.plotPieChart()
 
def tempData(args):
    # Data
    file_xlsx = args.path_file
    str_n_b_features = 'No bank features'
    str_y_b_features = 'Yes bank features'
    lst_sheets = ['Elaborazioni','Funnel']
    dtf_loaded = pd.read_excel(file_xlsx, sheetname = lst_sheets)
    lst_n_b_features = list([dtf_loaded[lst_sheets[0]].iloc[1, 1]]) + list(dtf_loaded[lst_sheets[1]].iloc[1, [1, 2]])
    lst_y_b_features = list([dtf_loaded[lst_sheets[0]].iloc[2, 1]]) + list(dtf_loaded[lst_sheets[1]].iloc[2, [1, 2]])
    dict_raw_data = {str_n_b_features: lst_n_b_features, str_y_b_features: lst_y_b_features}
    dtf_bar = pd.DataFrame(dict_raw_data)
    lst_pie = [100 * x / lst_y_b_features[-1] for x in list(dtf_loaded[lst_sheets[1]].iloc[3:10, 2])]
    del dtf_loaded
    return dtf_bar, lst_pie


class Graphic: 
    '''
    Create costomized plot
    '''
    
    
    def __init__(self, dtf_bar, lst_pie, args):
        '''
        Settings
        '''
        
        self.str_save_folder = args.save_folder
        self.dtf_bar = dtf_bar
        self.lst_pie = lst_pie
        self.str_n_b_features = dtf_bar.columns[0]
        self.str_y_b_features = dtf_bar.columns[1]
        
        # Data: from raw value to percentage
        self.lst_totals = [i + j for i, j in zip(self.dtf_bar[self.str_n_b_features], self.dtf_bar[self.str_y_b_features])]
        self.lst_n_bar = [i / j * 100 for i, j in zip(self.dtf_bar[self.str_n_b_features], self.lst_totals)]
        self.lst_y_bar = [i / j * 100 for i, j in zip(self.dtf_bar[self.str_y_b_features], self.lst_totals)]
 
        # Plot setting
        
            # General
        self.str_today = date.today().strftime('%Y-%m-%d')
        self.str_basecolor = '#120a77'
        rc('text', color = self.str_basecolor)
        rc('patch', edgecolor = self.str_basecolor)
        rc('axes', edgecolor = self.str_basecolor)
        rc('xtick', color = self.str_basecolor)
        rc('ytick', color = self.str_basecolor)
        self.str_legend_edgecolor = 'lightblue'
        self.int_dpi = 300
        
            # Bar chart
                # Axes
        self.lst_bar_ticks = [0, 1, 2]
        self.int_bar_right_xlim = 4
                # Bars
        self.flt_bar_width = 0.85
        self.str_n_bar_color = '#b5ffb9'
        self.str_y_bar_color = '#f9bc86'
        self.str_bar_edgecolor = 'white'
        self.lst_bar_names = ('All', 'Expiring', 'Contracts')
                # Title
        self.str_bar_title = '"Raccolta targhe"'
        self.lst_bar_title_pos = (0.5, 1.09)
                # Legend
        self.str_bar_legend_loc = 'lower left'
                # Conversion boxes
        self.str_conversion_boxstyle = 'circle'
        self.flt_conversion_boxalpha = 0.5
        self.str_y_conversion_boxfacecolor = 'wheat'
        self.str_n_conversion_boxfacecolor = '#c1fcab'
        self.flt_conversion_box_xcoord = .85
        self.int_conversion_box_fontsize = 14
                # Arrow boxes
        self.str_arrow_box = 'conversion'
        self.str_arrow_boxstyle = 'rarrow'
        self.flt_arrow_boxalpha = 0.5
        self.str_arrow_boxfacecolor = 'lightblue'
        self.flt_arrow_box_xcoord = .68
        self.int_arrow_box_fontsize = 8
        
            # Pie chart
                # Pie
        self.flt_pie_pct_limit = 10
        self.bln_pie_counterclock = False
        self.flt_pie_startangle = -33
        self.dict_pie_textprops = dict(color='w')
                # Title
        self.str_pie_title = 'Conversion drill-down'
                # Legend
        self.str_pie_legend = ('0 discount','5 discount','10 discount','15 discount','20 discount','25 discount','30 discount')
        self.str_pie_legend_loc = 'center right'
        self.lst_pie_legend_bboxtoanchor = (1.13, 0.5)

    
    def plotBarChart(self):
        '''
        Plot bar chart with addictional boxes
        '''
        
        fig, ax1 = plt.subplots(1, 1)

        # Plot "no" and "yes" bank features' bars
        plt.bar(self.lst_bar_ticks, self.lst_n_bar, color = self.str_n_bar_color, edgecolor = self.str_bar_edgecolor, width = self.flt_bar_width)
        plt.bar(self.lst_bar_ticks, self.lst_y_bar, bottom = self.lst_n_bar, color = self.str_y_bar_color, edgecolor = self.str_bar_edgecolor, width = self.flt_bar_width)
        # Create blue Bars
        #plt.bar(self.lst_bar_ticks, blueBars, bottom = [i + j for i, j in zip(self.lst_n_bar, self.lst_y_bar)], color = '#a3acff', edgecolor = self.str_bar_edgecolor, width = self.flt_bar_width)
 
        # Custom x axis
        ax1.set_xlim(right = self.int_bar_right_xlim)
        ax1.set_xticks(self.lst_bar_ticks)
        ax1.set_xticklabels(self.lst_bar_names)
        ax2 = ax1.twiny()
        plt.bar(self.lst_bar_ticks, self.lst_n_bar, color = self.str_n_bar_color, edgecolor = self.str_bar_edgecolor, width = self.flt_bar_width)
        plt.bar(self.lst_bar_ticks, self.lst_y_bar, bottom = self.lst_n_bar, color = self.str_y_bar_color, edgecolor = self.str_bar_edgecolor, width = self.flt_bar_width)
        ax2.set_xlim(right = self.int_bar_right_xlim)
        ax2.set_xticks(self.lst_bar_ticks)
        ax2.set_xticklabels(self.lst_totals)

        # Show title and legend
        plt.title(self.str_bar_title, position = self.lst_bar_title_pos)
        plt.legend((self.str_n_b_features, self.str_y_b_features), loc = self.str_bar_legend_loc, edgecolor = self.str_legend_edgecolor)

        # Set text for conversion boxes
        str_y_conversion_box = str(round(self.dtf_bar[self.str_y_b_features][2] / self.dtf_bar[self.str_y_b_features][1] * 100, 2)) + '%'
        str_n_conversion_box = str(round(self.dtf_bar[self.str_n_b_features][2] / self.dtf_bar[self.str_n_b_features][1] * 100, 2)) + '%'

        # These are matplotlib.patch.Patch properties
        dict_y_conversion_box = dict(boxstyle = self.str_conversion_boxstyle, facecolor = self.str_y_conversion_boxfacecolor, edgecolor = self.str_y_bar_color, alpha = self.flt_conversion_boxalpha)
        dict_n_conversion_box = dict(boxstyle = self.str_conversion_boxstyle, facecolor = self.str_n_conversion_boxfacecolor, edgecolor = self.str_n_bar_color, alpha = self.flt_conversion_boxalpha)
        dict_arrow_box = dict(boxstyle = self.str_arrow_boxstyle, facecolor = self.str_arrow_boxfacecolor, alpha = self.flt_arrow_boxalpha)

        # Set conversion box y-coordinates
        flt_y_lastbar_heigth = self.lst_n_bar[-1] + self.lst_y_bar[-1] / 2
        flt_n_lastbar_heigth = self.lst_n_bar[-1] / 2

        # Draw conversion boxes and arrow boxes
        ax2.text(self.flt_conversion_box_xcoord, flt_y_lastbar_heigth / 100, str_y_conversion_box, transform = ax2.transAxes, fontsize = self.int_conversion_box_fontsize,
                verticalalignment = 'top', bbox = dict_y_conversion_box)
        ax2.text(self.flt_conversion_box_xcoord, flt_n_lastbar_heigth / 100, str_n_conversion_box, transform = ax2.transAxes, fontsize = self.int_conversion_box_fontsize,
                verticalalignment = 'top', bbox = dict_n_conversion_box)
        ax2.text(self.flt_arrow_box_xcoord, flt_y_lastbar_heigth / 100, self.str_arrow_box, transform = ax2.transAxes, fontsize = self.int_arrow_box_fontsize,
                verticalalignment = 'top', bbox = dict_arrow_box)
        ax2.text(self.flt_arrow_box_xcoord, flt_n_lastbar_heigth / 100, self.str_arrow_box, transform = ax2.transAxes, fontsize = self.int_arrow_box_fontsize,
                verticalalignment = 'top', bbox = dict_arrow_box)

        # Draw simple arrows
    #arrowprop = dict(arrowstyle = 'simple', shrinkA = 5, shrinkB = 5, fc = self.str_basecolor, ec = self.str_basecolor, connectionstyle = "arc3, rad=-0.05")
        #ax2.annotate('', (3.2, flt_y_lastbar_heigth), (2.5, flt_y_lastbar_heigth), ha = "right", va = "top", size = 14, arrowprops = arrowprop)
        #ax2.annotate('', (3.2, flt_n_lastbar_heigth), (2.5, flt_n_lastbar_heigth), ha = "right", va = "top", size = 14, arrowprops = arrowprop)

        # Save image
        plt.savefig(self.str_save_folder + self.str_today + '_bar_chart.png', dpi = self.int_dpi)
        
        # Show graphic
        plt.show()


    def plotPieChart(self):
        '''
        Plot pie chart
        '''
        
        def autopctGenerator(flt_pie_pct_limit):
            '''
            Pie percentages are not showed if they are lower than flt_pie_pct_limit
            '''
            def innerAutopct(flt_pie_pct):
                return ('%.f%%' % flt_pie_pct) if flt_pie_pct > flt_pie_pct_limit else ''
            return innerAutopct
        
        # Plot pie
        plt.pie(self.lst_pie, counterclock = self.bln_pie_counterclock, startangle = self.flt_pie_startangle,
                autopct = autopctGenerator(self.flt_pie_pct_limit), textprops = self.dict_pie_textprops)
        
        # Show title and legend
        plt.title(self.str_pie_title)
        plt.legend(self.str_pie_legend, loc = self.str_pie_legend_loc, bbox_to_anchor = self.lst_pie_legend_bboxtoanchor, edgecolor = self.str_legend_edgecolor)
        plt.axis('equal')
                
        # Save image
        plt.savefig(self.str_save_folder + self.str_today + '_pie_chart.png', dpi = self.int_dpi)
        
        # Show graphic
        plt.show()
        

if __name__ == '__main__':
    main()
