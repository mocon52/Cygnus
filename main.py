import matplotlib.pyplot as plt
import xlrd as xl
import keyboard as kb
from matplotlib import interactive

interactive(True)

year = 2020
ax = plt.axes(projection='3d')
ax.set_autoscaley_on(True)
ax.set_autoscalex_on(True)
ax.set_autoscalez_on(True)
n = True

name = []
ra_2020 = []
dec_2020 = []
parallax_2020 = []
rv = []
pm_dec = []
pm_ra = []
dist_2020 = []
text_points = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]

wb = xl.open_workbook('Cygnus.xlsx')
data = wb.sheet_by_index(0)

for row in range(1, 12):
    name.append(data.cell_value(row, 0))
    ra_2020.append(data.cell_value(row, 1))
    dec_2020.append(data.cell_value(row, 2))
    parallax_2020.append(data.cell_value(row, 3))
    rv.append(data.cell_value(row, 4))
    pm_dec.append(data.cell_value(row, 5))
    pm_ra.append(data.cell_value(row, 6))
    dist_2020.append(30860000000000 / (parallax_2020[row - 1] * .001))

print(dist_2020)
text_year = plt.gcf().text(0.02, 0.8, 'Year= ' + str(year) + ' AD', fontsize=14)
text_extra = plt.gcf().text(0.02, 0.2, "A: 360\nG: Recenter", fontsize=8)

points_2020 = ax.scatter3D(ra_2020, dist_2020, dec_2020, color='red')
line_2020_1 = ax.plot3D([ra_2020[0], ra_2020[3], ra_2020[7], ra_2020[1], ra_2020[2]],
                        [dist_2020[0], dist_2020[3], dist_2020[7], dist_2020[1], dist_2020[2]],
                        [dec_2020[0], dec_2020[3], dec_2020[7], dec_2020[1], dec_2020[2]], color='red', linestyle=":")
line_2020_2 = ax.plot3D([ra_2020[10], ra_2020[9], ra_2020[8], ra_2020[4], ra_2020[3], ra_2020[5], ra_2020[6]],
                        [dist_2020[10], dist_2020[9], dist_2020[8], dist_2020[4], dist_2020[3], dist_2020[5],
                         dist_2020[6]],
                        [dec_2020[10], dec_2020[9], dec_2020[8], dec_2020[4], dec_2020[3], dec_2020[5], dec_2020[6]],
                        color='red', linestyle=":", label="Year 2020")

ax.set_xlabel("Right Ascension (Degrees)")
ax.set_ylabel("Distance (km)")
ax.set_zlabel("Declination (Degrees)")


def Update(year):
    global name, ra_2020, dec_2020, dist_2020, rv, pm_dec, pm_ra, text_points

    key = True
    ra = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
    dec = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
    dist = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
    year = int(year)
    t_year = year - 2020

    for x in range(0, 11):
        ra[x] = ra_2020[x] + (t_year * (pm_ra[x] * .001) / 3600)
        dist[x] = dist_2020[x] + t_year * 60 * 60 * 24 * 365.25 * rv[x]
        dec[x] = dec_2020[x] + (t_year * (pm_dec[x] * .001) / 3600)
    points_new = ax.scatter3D(ra, dist, dec, color='blue')
    line_1 = ax.plot3D([ra[0], ra[3], ra[7], ra[1], ra[2]], [dist[0], dist[3], dist[7], dist[1], dist[2]],
                       [dec[0], dec[3], dec[7], dec[1], dec[2]], color='blue')
    line_2 = ax.plot3D([ra[10], ra[9], ra[8], ra[4], ra[3], ra[5], ra[6]],
                       [dist[10], dist[9], dist[8], dist[4], dist[3], dist[5], dist[6]],
                       [dec[10], dec[9], dec[8], dec[4], dec[3], dec[5], dec[6]], color='blue', label='Current Year')

    ax.legend()
    ax.view_init(elev=0, azim=90)
    if year >= 0:
        text_year.set_text("Year= " + str(year) + ' AD')
    else:
        text_year.set_text("Year= " + str(-1 * year) + " BCE")

    for x in range(0, 11):
        ax.text(ra[x], dist[x], dec[x], name[x])
    plt.draw()
    plt.pause(1)
    while key:
        if kb.is_pressed('enter'):
            key = False
        elif kb.is_pressed('a'):
            for angle in range(90, 360 + 90):
                ax.view_init(elev=0, azim=angle)
                plt.draw()
                plt.pause(.0001)
        elif kb.is_pressed('g'):
            ax.view_init(elev=0, azim=90)
        else:
            plt.draw()
            plt.pause(.001)
    for txt in ax.texts:
        txt.set_visible(False)
    ax.lines.pop(3)
    ax.lines.pop(2)
    ax.collections.pop(1)


Update(year)
while n:
    year = input("Time Travel to: ")
    Update(year)

plt.show()
