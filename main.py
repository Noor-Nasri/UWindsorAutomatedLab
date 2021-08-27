from numpy.polynomial.polynomial import polyfit
import matplotlib.pyplot as plt
import xlrd

constant_ID = "SQRT(I_D)"
constant_VG = "V_G"
constant_dist = 0.5
#26063286.16

class Point:
    def __init__(self, gateV, drainI):
        self.gateV = float(gateV)
        self.sqDrainI = abs(float(drainI))**(1/2)
        self.artist = None

class Graph:
    def __init__(self, run, w, l, c):
        self.selected = [None, None]
        self.multiplier = 2*l/(c*w)
        self.sqRange = 0
        self.points = []
        self.run = run
    
    def sortKey(self, p): return p.gateV

    def addPoint(self, point):
        self.points.append(point)
        self.points = sorted(self.points, key=self.sortKey, reverse=True)
        if self.sqRange < point.sqDrainI:
            self.sqRange = point.sqDrainI
    
    def lineOfBestFit(self):
        ind0, ind1 = [self.points.index(self.selected[i]) for i in range(2)]
        xValues = [p.gateV for p in self.points[ind0:ind1 + 1]]
        yValues = [p.sqDrainI for p in self.points[ind0:ind1 + 1]]

        b, m = polyfit(xValues, yValues, 1)
        calculatedValues = [m * xValues[i] + b for i in range(len(xValues))]
        
        mobility = self.multiplier * (m**2)
        coeff_y = (1/m)*(1/b)
        coeff_x = -(1/b)
        const_c = (1/m)

        return (xValues, calculatedValues, coeff_x, coeff_y, const_c, mobility)

    def displayGraph(self):
        self.selected = [self.points[0], self.points[-1]]

        fig, ax = plt.subplots()
        fig.patch.set_facecolor((0.6, 0.6, 0.6))
        ax.set_facecolor((0.5, 0.5, 0.5))
        ax.set_xlabel(constant_VG)
        ax.set_ylabel(constant_ID)
        ax.set_title(self.run)
        self.lineArtist = None
        self.textArtist = None

        def on_pick(event=None):
            # swap selected and update color
            if event != None and not event.artist.obj in self.selected:
                p = event.artist.obj
                ind = event.mouseevent.button == 3 and 1 or 0

                p.artist.set_color("c")
                self.selected[ind].artist.set_color("b")
                self.selected[ind] = p

                if self.selected[0].gateV < self.selected[1].gateV:
                    self.selected = self.selected[::-1]

            # update line and text
            xValues, calculatedValues, x, y, c, m = self.lineOfBestFit()
            
            distext = "{0}x + {1}y = {2}\nMobility={3}".format(str(x), str(y), str(c), str(m))
            if (event == None):
                self.lineArtist = ax.plot(xValues, calculatedValues, "r-")[0]
                self.textArtist = ax.text(0.95, 0.95, distext, transform=ax.transAxes, fontsize=14, horizontalalignment='right',
                    verticalalignment='top', bbox=dict(boxstyle='round', facecolor='wheat', alpha=0.5))

            else:
                self.lineArtist.set_xdata(xValues)
                self.lineArtist.set_ydata(calculatedValues)
                self.textArtist.set_text(distext)
            
            # update display
            fig.canvas.draw()
            fig.canvas.flush_events()

        def customPicker(artist, mouseevent):
            return [
                abs(mouseevent.xdata - artist.obj.gateV) < constant_dist/2 and 
                abs(mouseevent.ydata - artist.obj.sqDrainI) < self.sqRange / 10, {}
            ]

        def createPlot():
            for p in self.points:
                artist = ax.plot(p.gateV, p.sqDrainI, p in self.selected and "c." or "b.", picker=customPicker)[0]
                artist.obj = p
                p.artist = artist

            fig.canvas.callbacks.connect('pick_event', on_pick)

        createPlot()
        on_pick()
        plt.show()

def main():
    length = float(input("Please input length: "))
    width = float(input("Please input width: "))
    capacitance = float(input("Please input capacitance: "))

    print("""=============================================
Perfect! You can now open .xls files to automate calculations
Simply input the path to the files one at a time, then for every Run[x] tab a graphic will show.
For better experience, try maximizing the screen. You can use Ctrl-C to force close
=============================================""")

    while True:
        path = input("Please input the path to an xls file: ")
        workbook = xlrd.open_workbook(path)

        for i, s in enumerate(workbook.sheets()):
            if ( s.name[:3] == "Run"):
                print("Opening", s.name, " - once done, simply close the graph!")
                graph = Graph(s.name, width, length, capacitance)
                last = 1000000

                for dr, gv in zip(s.col(0)[1:], s.col(4)[1:]):
                    if gv.value >= last:
                        break

                    graph.addPoint(Point(gv.value, dr.value))
                    last = gv.value

                graph.displayGraph()

if __name__ == '__main__':
    main()