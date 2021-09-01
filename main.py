from numpy.polynomial.polynomial import polyfit
from numpy.linalg import norm
from numpy import cross, abs as np_abs
import matplotlib.pyplot as plt
import xlsxwriter
import xlrd

constant_ID = "SQRT(I_D)"
constant_VG = "V_G"
constant_dist = 0.5

class Point:
    def __init__(self, gateV, drainI):
        self.gateV = float(gateV)
        self.sqDrainI = abs(float(drainI))**(1/2)
        self.artist = None
    
    def getVector(self):
        return [self.gateV, self.sqDrainI]

class Graph:
    def __init__(self, run, w, l, c):
        self.selected = [None, None]
        self.multiplier = 2*l/(c*w)
        self.lastMobility = None
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
        
        self.lastMobility = self.multiplier * (m**2)
        self.lastVThresh = -b/m

        coeff_y = (1/m)*(1/b)
        coeff_x = -(1/b)
        const_c = (1/m)

        return (xValues, calculatedValues, coeff_x, coeff_y, const_c)

    def displayGraph(self):
        
        # calculate best initial line
        best = [0, None, None]
        for p1 in range(len(self.points) - 4):
            originPoint = self.points[p1]
            bestIter = [0, -1]
            totalSlope = 0
            for p2 in range(p1 + 1, len(self.points)):
                secondaryPoint = self.points[p2]
                newSlope = (secondaryPoint.sqDrainI - originPoint.sqDrainI)/(secondaryPoint.gateV - originPoint.gateV)
                totalSlope += newSlope

                if (p2 > p1 + 3):
                    averageSlope = totalSlope / (p2 - p1)
                    ld = abs(newSlope - averageSlope)
                    p = (abs(averageSlope) - ld) / abs(averageSlope)

                    if p*1.05 > bestIter[0]:
                        bestIter = [p, p2]

            if bestIter[0] > best[0]:
                best = [bestIter[0], p1, bestIter[1]]

        self.selected = [self.points[best[1]], self.points[best[2]]]

        # create graph
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
                ind = event.mouseevent.button != 3 and 1 or 0

                p.artist.set_color("c")
                self.selected[ind].artist.set_color("b")
                self.selected[ind] = p

                if self.selected[0].gateV < self.selected[1].gateV:
                    self.selected = self.selected[::-1]

            # update line and text
            xValues, calculatedValues, x, y, c = self.lineOfBestFit()
            
            distext = "{0}x + {1}y = {2}\nMobility={3}\nVth={4}".format(str(x), str(y), str(c), str(self.lastMobility), str(self.lastVThresh))
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
    dc = float(input("Please input the dielectric constant: "))
    dlt = float(input("Please input the dielectric layer thickness: "))
    capacitance = dc * 8.8541878176E-14 / dlt

    print("""=============================================
Perfect! You can now open .xls files to automate calculations
Simply input the path to the files one at a time, then for every Run[x] tab a graphic will show.
For better experience, try maximizing the screen. You can use Ctrl-C to force close
=============================================""")

    while True:
        path = input("Please input the path to an xls file, or 0 to exit ")
        if path == "0":
            break
        
        workbook = xlrd.open_workbook(path)
        saved = []
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
                print("Mobility -", graph.lastMobility)
                print("Threshold voltage -", graph.lastVThresh)

                saved.append([s.name, graph.lastMobility, graph.lastVThresh])
        
        
        newPath = path.replace(".xls", "-results.xls")

        workbook = xlsxwriter.Workbook(newPath)
        worksheet = workbook.add_worksheet()
        worksheet.write('A1', 'Run')
        worksheet.write('B1', 'Mobility')
        worksheet.write('C1', 'Threshold Voltage (Vth)')

        for y in range(len(saved)):
            worksheet.write('A' + str(y + 2), str(saved[y][0]))
            worksheet.write('B' + str(y + 2), str(saved[y][1]))
            worksheet.write('C' + str(y + 2), str(saved[y][2]))

        workbook.close()
        print("Mobilities saved to", newPath)
        

if __name__ == '__main__':
    main()
