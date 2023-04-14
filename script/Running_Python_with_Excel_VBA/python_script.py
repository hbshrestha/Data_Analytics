import pandas as pd
import matplotlib.pyplot as plt
import os
import sys

#Enter os system to current working directory
os.chdir(sys.path[0])

file = "solar_irradiance.xlsm"

#read all rows and first 5 columns
df = pd.read_excel(file).iloc[:, :5]

df["Datetime"] = pd.to_datetime(df["Datetime"])
df.set_index(["Datetime"], inplace = True)

frequency = input("Enter the frequency you want to display? \n1. Original \n2. Monthly average\n3. Daily average \n4. Weekly average\n 5.Quarterly average \n 6.All of the above \n 7. Hourly average \n? ")

if frequency == "Original":
    print (df)
    df.plot()
    plt.title("Original solar irradiance in 2020")
    plt.ylabel("Wh/m$^2$")
    plt.legend()
    plt.show()

elif frequency == "Monthly average":
    print (df.resample(rule = "M").mean())
    df.resample(rule = "M").mean().plot()
    plt.title("Monthly average solar irradiance in 2022")
    plt.ylabel("Wh/m$^2$")
    plt.legend()
    plt.show()

elif frequency == "Daily average":
    print (df.resample(rule = "D").mean())
    df.resample(rule = "D").mean().plot()
    plt.title("Daily average solar irradiance in 2022")
    plt.ylabel("Wh/m$^2$")
    plt.show()

elif frequency == "Weekly average":
    print (df.resample(rule = "W").mean())
    df.resample(rule = "W").mean().plot()
    plt.title("Weekly average solar irradiance in 2022")
    plt.ylabel("Wh/m$^2$")
    plt.legend()
    plt.show()


elif frequency == "Quarterly average":
    print (df.resample(rule = "Q").mean())
    df.resample(rule = "Q").mean().plot()
    plt.title("Quarterly average solar irradiance in 2022")
    plt.ylabel("Wh/m$^2$")
    plt.legend()
    plt.show()
    
elif frequency == "All of the above":
    fig, axs = plt.subplots(2, 2, figsize = (20, 10), sharey = True, sharex = True)
    df.resample(rule = "D").mean().plot(ax = axs[0, 0])
    axs[0, 0].set_title("Daily mean")
    df.resample(rule = "W").mean().plot(ax = axs[0, 1])
    axs[0, 1].set_title("Weekly mean")
    df.resample(rule = "M").mean().plot(ax = axs[1, 0])
    axs[1, 0].set_title("Monthly mean")
    df.resample(rule = "Q").mean().plot(ax = axs[1, 1])
    axs[1, 1].set_title("Quarterly mean")
    fig.suptitle("Mean solar irradiance in four locations converted to different temporal frequencies")
    plt.show()

elif frequency == "Hourly average":
    #average value in each hour within 24 hours of a day
    print (df.groupby(df.index.hour).mean())
    df.groupby(df.index.hour).mean().plot()
    plt.title("Hourly average solar irradiance in 2022")
    plt.ylabel("Wh/m$^2$")
    plt.legend()
    plt.show()

else:
    print ("The frequency you entered is incorrect.")
