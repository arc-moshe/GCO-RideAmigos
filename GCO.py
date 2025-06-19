import streamlit as st
import pandas as pd
import geopandas as gpd 
from shapely.geometry import Point 
import datetime
import io 
import numpy as np
import zipfile 

st.title("RideAmigos Report Processing")

startDate = st.date_input("Select the first date of the reporting period", datetime.date.today())
endDate = st.date_input("Select the last date of the reporting period", datetime.date.today())


user_file = st.file_uploader("Choose the Users Report", type ="xlsx")
if user_file is not None:
    df_users=pd.read_excel(user_file)
    st.write("File uploaded!")

    # Get rid of junk/test records in user file
    df_users.drop(df_users[df_users["Networks"].isin(["RideAmigos Employees", "RideAmigos Test Network"])].index, inplace=True)
    df_users.drop(df_users[df_users["Email"].fillna("").str.contains("@rideamigos.com|@example.com|@test.com", regex=True)].index, inplace=True)
    df_users.drop(df_users[df_users["Email"].isin([
        "appreview2055@icloud.com", 
        "webteam@odonnellco.com", 
        "maureen.contestabile@odonnellco.com", 
        "kathryn.hagerman@gmail.com", 
        "bendalton+aminew@gmail.com", 
        "chancemagno@gmail.com", 
        "acuadrado@atlantaregional.com", 
        "acuadrado@gacommuteoptions.com", 
        "support@mygacommuteoptions.com", 
        "support@gacommuteoptions.com"
    ])].index, inplace=True)
    df_users.drop(df_users[df_users["Employer Name"].isin([
        "RideAmigos",
        "Test Employer"
    ])].index, inplace=True)

    st.subheader("Users Data Preview")
    st.write(df_users.head())
else: 
   st.write("Waiting for upload.")

trip_file = st.file_uploader("Choose the Trips Report", type ="xlsx")
if trip_file is not None:
    df_trips=pd.read_excel(trip_file)
    st.write("File uploaded!")

    # Get rid of junk/test records from trips file:

    df_trips.drop(df_trips[df_trips["Networks"].isin(["RideAmigos Employees", "RideAmigos Test Network"])].index, inplace=True)
    df_trips.drop(df_trips[df_trips["User Email"].fillna("").str.contains("@rideamigos.com|@example.com|@test.com", regex=True)].index, inplace=True)
    df_trips.drop(df_trips[df_trips["User Email"].isin([
        "appreview2055@icloud.com", 
        "webteam@odonnellco.com", 
        "maureen.contestabile@odonnellco.com", 
        "kathryn.hagerman@gmail.com", 
        "bendalton+aminew@gmail.com", 
        "chancemagno@gmail.com", 
        "acuadrado@atlantaregional.com", 
        "acuadrado@gacommuteoptions.com", 
        "support@mygacommuteoptions.com", 
        "support@gacommuteoptions.com"
    ])].index, inplace=True)
    df_trips.drop(df_trips[df_trips["User Name"].isin(["Network Log"])].index, inplace=True)

    st.subheader("Trips Data Preview")
    st.write(df_trips.head())
else: 
   st.write("Waiting for upload.")


if trip_file is not None and user_file is not None: 
    st.subheader("Process Records")
    st.write(f"Click the button below to process records for the period {startDate} through {endDate}.")
    if st.button("PROCESS RECORDS"):
        
        st.write("Working: may take a few minutes to process...")

        # The User ID is called "_id" in the Users table but "User ID" in the trip log, so we adjust the name in the users dataframe to match for joining purposes.
        df_users = df_users.rename(columns={'_id': 'User ID'})

        # Split coordinates into latitude and longitude
        df_users[['LonHome', 'LatHome']] = df_users['Home Location Coords'].str.split(',', expand=True).astype(float)
        df_users[['LonWork', 'LatWork']] = df_users['Work Location Coords'].str.split(',', expand=True).astype(float)

        # Convert to GeoDataFrame
        df_users['geometry_work'] = df_users.apply(lambda row: Point(row['LonWork'], row['LatWork']), axis=1)
        df_users['geometry_home'] = df_users.apply(lambda row: Point(row['LonHome'], row['LatHome']), axis=1)

        gdf_work = gpd.GeoDataFrame(df_users, geometry='geometry_work', crs="EPSG:4326")
        gdf_home = gpd.GeoDataFrame(df_users, geometry='geometry_home', crs="EPSG:4326")

        # Load ESO shapefile
        eso_gdf = gpd.read_file("data/Employer_Service_Organizations.shp")
        eso_gdf.rename(columns={'NAME': 'ESO'}, inplace=True)

        # Load Counties shapefile
        counties_gdf = gpd.read_file("data/tl_2020_13_county20.shp")

        # Load ZCTA shapefile
        zip_gdf = gpd.read_file("data/tl_2020_13_zcta520.shp")

        # Ensure all dataframes are in the coordinaate reference system
        eso_gdf = eso_gdf.to_crs(gdf_work.crs)
        zip_gdf = zip_gdf.to_crs(gdf_work.crs)
        counties_gdf = counties_gdf.to_crs(gdf_work.crs)

        # Perform spatial joins

        # Joins to ESO: (Note: we're determining Home ESO in order to check whether a home address is within region)
        gdf_work = gpd.sjoin(gdf_work, eso_gdf, how="left", predicate="within")
        gdf_work.drop(columns=['index_right'], inplace=True)

        gdf_home = gdf_home.reset_index(drop=True)
        eso_gdf = eso_gdf.reset_index(drop=True)

        gdf_home = gpd.sjoin(gdf_home, eso_gdf, how="left", predicate="within")
        gdf_home.drop(columns=['index_right'], inplace=True)
        gdf_home.rename(columns={'ESO': 'Home ESO'}, inplace=True)

        # Joins to ZCTA:
        gdf_home = gpd.sjoin(gdf_home, zip_gdf, how="left", predicate="within")
        gdf_home.rename(columns={'GEOID20': 'Home ZIP'}, inplace=True)
        gdf_home.drop(columns=['index_right'], inplace=True)

        gdf_work = gpd.sjoin(gdf_work, zip_gdf, how="left", predicate="within")
        gdf_work.rename(columns={'GEOID20': 'Work ZIP'}, inplace=True)
        gdf_work.drop(columns=['index_right'], inplace=True)

        # Joins to Counties:
        gdf_work = gpd.sjoin(gdf_work, counties_gdf, how="left", predicate="within")
        gdf_work.rename(columns={'NAME20': 'Work County Name', 'GEOID20': 'Work County FIPS'}, inplace=True)

        gdf_home = gpd.sjoin(gdf_home, counties_gdf, how="left", predicate="within")
        gdf_home.rename(columns={'NAME20': 'Home County Name', 'GEOID20': 'Home County FIPS'}, inplace=True)

        # Merge results of spatial joins into original users dataframe
        df_users = df_users.merge(gdf_work[['User ID', 'ESO', 'Work ZIP', 'Work County Name','Work County FIPS']], on="User ID", how="left")
        df_users = df_users.merge(gdf_home[['User ID', 'Home ESO', 'Home ZIP','Home County Name','Home County FIPS']], on="User ID", how="left")


        # Handle missing data: If lat/lon exists but ESO spatial join is empty, code as "Out of Region"; if lat/lon null, code as "Unknown"
        # Then we'll use the ESO results to determine unknown/out of region for the remaining fields

        df_users['ESO'] = df_users.apply(lambda row: "Out of Region" if pd.isna(row['ESO']) and pd.notna(row['LonWork']) and pd.notna(row['LatWork']) 
                                        else ("Unknown" if pd.isna(row['LonWork']) or pd.isna(row['LatWork']) else row['ESO']), axis=1)

        df_users['Home ESO'] = df_users.apply(lambda row: "Out of Region" if pd.isna(row['Home ESO']) and pd.notna(row['LonHome']) and pd.notna(row['LatHome']) 
                                        else ("Unknown" if pd.isna(row['LonHome']) or pd.isna(row['LatHome']) else row['Home ESO']), axis=1)

        df_users.loc[df_users['ESO'] == "Unknown", 'Work ZIP'] = "Unknown"
        df_users.loc[df_users['ESO'] == "Out of Region", 'Work ZIP'] = "Out of Region"

        df_users.loc[df_users['ESO'] == "Unknown", 'Work County Name'] = "Unknown"
        df_users.loc[df_users['ESO'] == "Out of Region", 'Work County Name'] = "Out of Region"

        df_users.loc[df_users['ESO'] == "Unknown", 'Work County FIPS'] = "Unknown"
        df_users.loc[df_users['ESO'] == "Out of Region", 'Work County FIPS'] = "Out of Region"

        df_users.loc[df_users['Home ESO'] == "Unknown", 'Home ZIP'] = "Unknown"
        df_users.loc[df_users['Home ESO'] == "Out of Region", 'Home ZIP'] = "Out of Region"

        df_users.loc[df_users['Home ESO'] == "Unknown", 'Home County Name'] = "Unknown"
        df_users.loc[df_users['Home ESO'] == "Out of Region", 'Home County Name'] = "Out of Region"

        df_users.loc[df_users['Home ESO'] == "Unknown", 'Home County FIPS'] = "Unknown"
        df_users.loc[df_users['Home ESO'] == "Out of Region", 'Home County FIPS'] = "Out of Region"


        # The GDOT report needs certain records marked as "GCO/State Fed". 
        # Add a new field containing the same value as 'ESO' if 'State/Fed' is blank, and "GCO State/Fed" if 'State/Fed' contains any value.
        df_users['ESO Adjust State/Fed'] = df_users['ESO']
        df_users.loc[df_users['State/Fed'].notna() & (df_users['State/Fed'] != ''), 'ESO Adjust State/Fed'] = "GCO State/Fed"

        # Create dummy variable to count New Users. First create datetime, then convert to date, then make comparison to date range previously specified by the user.
        # df_users['Registration Date'] = pd.to_datetime(df_users['Created'], format='%m/%d/%Y')
        df_users['Registration Date'] = pd.to_datetime(df_users['Created'], format='%m/%d/%y %I:%M %p')
        df_users['Registration Date'] = df_users['Registration Date'].dt.date
        df_users['New Users'] = ((df_users['Registration Date'] >= startDate) & (df_users['Registration Date'] <= endDate)).astype(int)

        # We can save intermediate results for diagnostic purposes if needed:
        # df_users.to_excel("Spatial Join Test.xlsx", index=False)



        # Next we work on the TRIPS dataframe

        # Add fields present in the Users Dataframe to the Trips dataframe: "ESO", "Home ZIP", and "ESO Adjust State/Fed".
        df_trips = df_trips.merge(df_users[['User ID', 'ESO', 'Home ZIP', 'ESO Adjust State/Fed']], on='User ID', how='left')

        # We can save intermediate results for diagnostic purposes if needed:
        #df_trips.to_excel("Merge Test.xlsx", index=False)

        # Rename column names to match desired output
        df_trips.rename(columns={'Mode': 'Method', 'CO2 Savings (grams)': 'CO2', 'Dollars Savings': 'Dollars', 'Vehicle Miles Reduced': 'VMR'}, inplace=True)

        # Change the Method values to match desired output, e.g., "cww" in raw data should come out "CWW"
        df_trips['Method'] = df_trips['Method'].replace({'bike': 'Bike', 'carpool': 'Carpool', 'cww': 'CWW', 'drive':'Drive', 'scooter': 'Scooter', 'telework': 'Telework', 'transit': 'Transit', 'vanpool': 'Vanpool', 'walk': 'Walk'})

        # Keep only relevant columns and reorder
        keep_columns = ['User ID', 'ESO', 'ESO Adjust State/Fed', 'Home ZIP', 'Method', 'Trips', 'Miles', 'VMR','CO2','Dollars']
        df_trips = df_trips[keep_columns]

        # We can save intermediate results for diagnostic purposes if needed:
        # df_trips.to_excel("DF Trips Test.xlsx", index=False)

        # Add a field called "Logs"-- summing this up will eventually give us trips logged x Method
        df_trips['Logs'] = 1

        # We can save intermediate results for diagnostic purposes if needed:
        # df_trips.to_excel("Merge Test.xlsx", index=False)

        # Collapse the trips dataframe to df_individual and df_individual_adjusted (ESO adjusted for State/Fed), each with with one record per person

        df_individual = df_trips.groupby(['User ID', 'Method', 'ESO', 'Home ZIP'], as_index=False).agg({'Trips': 'sum', 'Miles': 'sum', 'VMR': 'sum', 'CO2': 'sum', 'Dollars': 'sum'})

        df_individual_adjusted = df_trips.groupby(['User ID', 'Method', 'ESO Adjust State/Fed', 'Home ZIP'], as_index=False).agg({'Trips': 'sum', 'Miles': 'sum', 'VMR': 'sum', 'CO2': 'sum', 'Dollars': 'sum'})

        # Save for diagnostics
        # df_individual_adjusted.to_excel("DF Individual Test.xlsx", index=False)


        # Now we'll count loggers and clean loggers (used in GDOT report, but may be useful elsewhere)

        # Take a slice of the individual DF to work on separately 
        df_loggers = df_individual_adjusted[['User ID','Method']].copy()

        # Logger always equals 1 (sum to count)
        df_loggers['Logger'] = 1

        # Generate 'Clean Loggers' dummy varible (sum to count)
        df_loggers['Clean Loggers'] = 1
        df_loggers.loc[df_loggers['Method'] == 'Drive', 'Clean Loggers'] = 0

        # We can save intermediate results for diagnostic purposes if needed:
        # df_loggers.to_excel("DF Loggers PreAgg.xlsx", index=False)

        # Collapse to 1 record per user ID. Using "max" means even one clean trip logged makes you a clean logger for counting
        df_loggers = df_loggers.groupby(['User ID'], as_index=False).agg({'Clean Loggers': 'max', 'Logger': 'max'})

        # Join ESO back to df_loggers
        df_loggers = df_loggers.merge(df_users[['User ID', 'ESO', 'ESO Adjust State/Fed']], on='User ID', how='left')


        # Generate Territory field, which collapses Unknown and Out of region to a single category; also combines all GCO regions into one
        df_loggers['Territory'] = df_loggers['ESO Adjust State/Fed']
        df_loggers['Territory'] = df_loggers['Territory'].replace({'Unknown': 'Unknown/Out of Region', 'Out of Region': 'Unknown/Out of Region'})
        df_loggers.loc[df_loggers['Territory'].str.contains('GCO', na=False), 'Territory'] = 'GCO'

        # Then sum loggers and clean loggers dummies to 'Territory' in order to get counts
        df_loggers = df_loggers.groupby(['Territory']).agg({'Clean Loggers': 'sum', 'Logger': 'sum'})

        # We can save intermediate results for diagnostic purposes if needed:
        # df_loggers.to_excel("DF Loggers Test.xlsx", index=True)




        # OUTPUT FILE #1 Tableau Excel sheet
        # This report wants one record per O/D Pair by Method. We use the unadjusted ESO for this report.

        # Collapse the individual level data to one record per Method/ESO/Home ZIP triplet
        df_tableau = df_individual.groupby(['Method','ESO', 'Home ZIP']).sum().reset_index()
        df_tableau.drop(columns=['User ID'], inplace=True)


        # Add a date field, then keep only required columns in their desired order
        df_tableau['Date'] = startDate
        keep_columns = ['Date', 'Home ZIP','ESO', 'Method', 'Trips', 'Miles', 'VMR','CO2','Dollars']
        df_tableau = df_tableau[keep_columns]

        # We don't want drivers in the viz:
        df_tableau = df_tableau[df_tableau['Method'] != 'drive']

        # Sort by ZIP, ESO, and Method
        df_tableau = df_tableau.sort_values(by=['Home ZIP', 'ESO', 'Method'])

        # Export!
        # df_tableau.to_excel("c:\data\Tableau Test.xlsx", index=False)





        # OUTPUT FILE #2: GDOT Report
        # This report wants one line per ESO, called "Territory", and using the ESO Adjusted for State/Fed
        # Data Fields: "New Users", "Loggers",  "Clean Loggers", "Carpool Logs", "Vanpool Logs", "Transit Logs", "Telework Logs",
        #              "Walk Logs", "Bike Logs", "Scooter Logs", "CWW Logs", "Reduced VMT", "Reduced CO2 (pounds)" 



        # Start with the individual-level data:
        # Create a new field, Territory, which combines all GCO regions into one category and also combines unknown and out of region into one category.

        df_individual_adjusted['Territory'] = df_individual_adjusted['ESO Adjust State/Fed']
        df_individual_adjusted['Territory'] = df_individual_adjusted['Territory'].replace({'Unknown': 'Unknown/Out of Region', 'Out of Region': 'Unknown/Out of Region'})
        df_individual_adjusted.loc[df_individual_adjusted['Territory'].str.contains('GCO', na=False), 'Territory'] = 'GCO'

        # We can save intermediate results for diagnostic purposes if needed:
        # df_individual_adjusted.to_excel("Pre-Agg Test.xlsx", index=False)

        # Aggregate Individual-level data to ESO, then drop unneded fields
        df_gdot = df_individual_adjusted.groupby(['Territory']).sum().reset_index()

        # We can save intermediate results for diagnostic purposes if needed:
        # df_gdot.to_excel("Count Test.xlsx", index=False)


        # CO2 is in grams create new field with value converted to pounds:
        df_gdot['Reduced CO2 (pounds)'] = df_gdot['CO2'] * 0.00220462
        # df_gdot['Reduced CO2 (pounds)'] = df_gdot['CO2'] * 0.00220462262


        # Rename fields to desired output names for the GDOT report
        df_gdot = df_gdot.rename(columns={'VMR':'Reduced VMT', 'Dollars':'Money Saved'})


        # Now Transition to the Users data; count new users by territory

        df_users['Territory'] = df_users['ESO Adjust State/Fed']
        df_users['Territory'] = df_users['Territory'].replace({'Unknown': 'Unknown/Out of Region', 'Out of Region': 'Unknown/Out of Region'})
        df_users.loc[df_users['Territory'].str.contains('GCO', na=False), 'Territory'] = 'GCO'
        df_gdot_newusers = df_users.groupby('Territory').sum('New User').reset_index()

        # Add the new user field to the gdot
        df_gdot = df_gdot.merge(df_gdot_newusers, on='Territory', how='inner')


        # Now transition to the Trips data:

        # Again, create a new Territory field
        df_trips['Territory'] = df_trips['ESO Adjust State/Fed']
        df_trips['Territory'] = df_trips['Territory'].replace({'Unknown': 'Unknown/Out of Region', 'Out of Region': 'Unknown/Out of Region'})
        df_trips.loc[df_trips['Territory'].str.contains('GCO', na=False), 'Territory'] = 'GCO'

        # Collapse to Territory
        df_gdot_long = df_trips.groupby(['Territory', 'Method']).sum().reset_index()
        keep_columns = ['Territory','Method','Logs']
        df_gdot_long = df_gdot_long[keep_columns]

        # We can save intermediate results for diagnostic purposes if needed:
        # df_gdot_long.to_excel("Test GDOT Long.xlsx", index=False)


        # Reshaping from long to wide format (and dumping the "Drive" column because we don't report driving trips to GDOT)

        df_gdot_wide = df_gdot_long.pivot(index='Territory', columns='Method', values='Logs')
        df_gdot_wide.reset_index(inplace=True)

        df_gdot_wide.drop(columns=['Drive'], inplace=True)

        # Changing nulls to zeroes and renaming fields to desired output names

        columns_to_replace = ['Bike', 'Carpool', 'CWW', 'Scooter', 'Telework', 'Transit', 'Vanpool', 'Walk']
        df_gdot_wide[columns_to_replace] = df_gdot_wide[columns_to_replace].fillna(0)

        df_gdot_wide = df_gdot_wide.rename(columns={'Bike':'Bike Logs', 'Carpool': 'Carpool Logs', 'CWW':'CWW Logs', 'Scooter':'Scooter Logs', 
                                                    'Telework': 'Telework Logs', 'Transit': 'Transit Logs', 'Vanpool':'Vanpool Logs', 'Walk': 'Walk Logs'})


        # We can save intermediate results for diagnostic purposes if needed:
        # df_gdot_wide.to_excel("Test GDOT Wide.xlsx", index=False)

        # Add the fields from df_gdot_wide to the main df_gdot dataframe
        df_gdot = df_gdot.merge(df_gdot_wide, on='Territory', how='inner')

        # And now also the Loggers and Clean Loggers fields from the df_loggers dataframs
        df_gdot = df_gdot.merge(df_loggers, on='Territory', how='inner')

        # And rename "Logger" to "Loggers" to match desierd output
        df_gdot.rename(columns={'Logger':'Loggers'}, inplace = True)

        # Keep just what we need in the desired order
        keep_columns = ['Territory', 'New Users', 'Loggers', 'Clean Loggers', 'Carpool Logs', 'Vanpool Logs', 'Transit Logs','Telework Logs','Walk Logs',
                        'Bike Logs', 'Scooter Logs', 'CWW Logs', 'Reduced VMT', 'Reduced CO2 (pounds)', 'Money Saved']
        df_gdot = df_gdot[keep_columns]


        # Export!
        # df_gdot.to_excel("c:\data\GDOT Report Test.xlsx", index=False)



        # OUTPUT FILE #3: TDM VIZ

        # Start with the Users Dataframe: create a new df with just the fields we need
        # Using the unadjusted for now?

        df_tdm = df_users[['User ID', 'LonHome', 'LatHome', 'LonWork', 'LatWork', 'Active Account', 'New Users', 'Created', 
                        'Work County FIPS', 'Work County Name', 'Work ZIP', 'Home County FIPS', 'Home County Name', 'Home ZIP',
                        'Legacyid', 'ESO', 'Tmas']].copy()

        # Add a month field
        df_tdm['Month'] = startDate

        # Legacy is blank if there's no legacy ID, and takes a value of 1 if there is one:
        df_tdm['Legacy'] = np.nan 
        df_tdm.loc[df_tdm['Legacyid'] != "", 'Legacy'] = 1 

        # Rename fields to have the desired output names
        df_tdm.rename(columns={'User ID': 'User_ID', 'LonHome': 'Home_X', 'LatHome': 'Home_Y', 'LonWork': 'Work_X', 
                            'LatWork': 'Work_Y', 'Active Account': 'Active', 'New Users': 'New', 'Created': 'Created_Date', 
                            'Work County FIPS': 'County_ID_Work', 'Work County Name': 'County_Work', 'Work ZIP': 'Zip_Code_Work', 
                            'Home County FIPS': 'County_ID_Home', 'Home County Name': 'County_Home', 'Home ZIP': 'Zip_Code_Home'}, inplace = True)

        # Next, we move to the df_individual dataframe (contains one aggregated record per user); using unadjusted ESO

        # Convert grams to pounds
        df_individual['CO2_lbs'] = df_individual['CO2'] * 0.00220462


        # Reshape long to wide

        df_individual_wide = df_individual.pivot(index='User ID', columns='Method', values=['Trips', 'Miles', 'VMR', 'CO2_lbs', 'Dollars']).reset_index()
        df_individual_wide.columns = [f"{method}_{var}" if method else var for var, method in df_individual_wide.columns]
        df_individual_wide.rename(columns={'User ID': 'User_ID'}, inplace=True)


        # We can save intermediate results for diagnostic purposes if needed:
        # df_individual_wide.to_excel("Test Individual Wide.xlsx", index=True)

        # Add the df_individual_wide data to the main TDM dataframe
        df_tdm = df_tdm.merge(df_individual_wide, on='User_ID', how='left')


        # Replace missing values with zeroes
        patterns = ['Trips', 'Miles', 'VMR', 'CO2_lbs', 'Dollars']
        columns_to_replace = [col for col in df_tdm.columns if any(pattern in col for pattern in patterns)]
        df_tdm[columns_to_replace] = df_tdm[columns_to_replace].fillna(0)

        # Create Clean fields and Loggers fields
        # Initialize aggregated columns

        df_tdm["Clean_Trips"] = 0
        df_tdm["Clean_Miles"] = 0
        df_tdm["Clean_VMR"] = 0
        df_tdm["Clean_CO2_lbs"] = 0
        df_tdm["Clean_Dollars"] = 0

        # Create "Clean" totals for trips, miles, VMR, CO2 reduction, and dollars saved by summing across all of the different modes 
        modes = ["Bike", "Carpool", "CWW", "Scooter", "Telework", "Transit", "Vanpool", "Walk"]
        for mode in modes:
            df_tdm[f"{mode}_Logger"] = np.where(df_tdm[f"{mode}_Trips"].fillna(0) > 0, 1, np.nan)
            
            df_tdm["Clean_Trips"] += df_tdm[f"{mode}_Trips"].fillna(0)
            df_tdm["Clean_Miles"] += df_tdm[f"{mode}_Miles"].fillna(0)
            df_tdm["Clean_VMR"] += df_tdm[f"{mode}_VMR"].fillna(0)
            df_tdm["Clean_CO2_lbs"] += df_tdm[f"{mode}_CO2_lbs"].fillna(0)
            df_tdm["Clean_Dollars"] += df_tdm[f"{mode}_Dollars"].fillna(0)

        # Create Clean_Logger column
        df_tdm["Clean_Logger"] = np.where(df_tdm["Clean_Trips"] > 0, 1, np.nan)


        # Rename Tmas to TMA to match desired column name in output
        df_tdm.rename(columns={'Tmas': 'TMA'}, inplace=True)

        # Reorder to match desired output
        keep_columns = ['User_ID', 'Home_X', 'Home_Y', 'Work_X', 'Work_Y', 'TMA', 'Legacy', 'Active', 'New', 'Created_Date', 
                        'Bike_Logger', 'Bike_Trips', 'Bike_Miles', 'Bike_VMR', 'Bike_CO2_lbs', 'Bike_Dollars', 
                        'Carpool_Logger', 'Carpool_Trips', 'Carpool_Miles', 'Carpool_VMR', 'Carpool_CO2_lbs', 'Carpool_Dollars', 
                        'CWW_Logger', 'CWW_Trips', 'CWW_Miles', 'CWW_VMR', 'CWW_CO2_lbs', 'CWW_Dollars', 
                        'Scooter_Logger', 'Scooter_Trips', 'Scooter_Miles', 'Scooter_VMR', 'Scooter_CO2_lbs', 'Scooter_Dollars', 
                        'Telework_Logger', 'Telework_Trips', 'Telework_Miles', 'Telework_VMR', 'Telework_CO2_lbs', 'Telework_Dollars', 
                        'Transit_Logger', 'Transit_Trips', 'Transit_Miles', 'Transit_VMR', 'Transit_CO2_lbs', 'Transit_Dollars', 
                        'Vanpool_Logger', 'Vanpool_Trips', 'Vanpool_Miles', 'Vanpool_VMR', 'Vanpool_CO2_lbs', 'Vanpool_Dollars', 
                        'Walk_Logger', 'Walk_Trips', 'Walk_Miles', 'Walk_VMR', 'Walk_CO2_lbs', 'Walk_Dollars', 
                        'Clean_Logger', 'Clean_Trips', 'Clean_Miles', 'Clean_VMR', 'Clean_CO2_lbs', 'Clean_Dollars', 
                        'Month', 'County_ID_Work', 'County_Work', 'Zip_Code_Work', 'ESO', 'County_ID_Home', 'County_Home', 'Zip_Code_Home']
        df_tdm = df_tdm[keep_columns]

        # keep only active users
        df_tdm = df_tdm[df_tdm['Active'] == 1]

        # Export! 
        # df_tdm.to_excel("c:\data\TDM Report Test.xlsx", index=False)



        # OUTFILE #4: Data Audit: Flag records where ESO from spatial join does not match 'Tmas' in Ride Amigos  

        df_diff = df_users[['User ID', 'First Name', 'Last Name', 'Work Location', 'Tmas', 'ESO']].copy()
        df_diff = df_diff.rename(columns={'Tmas': 'TMA', 'ESO': 'ESO Geocoded'})

        # Remove colons from ESO Geocoded to match what is in TMA
        df_diff['ESO Geocoded'] = df_diff['ESO Geocoded'].str.replace(":", "", regex=True)

        # Replace specific values to match what is in TMA
        df_diff.loc[df_diff['ESO Geocoded'].isin(["Out of Region", "Unknown"]), 'ESO Geocoded'] = "Unknown/Out of Region"
        df_diff.loc[df_diff['ESO Geocoded'] == "Midtown Transportation", 'ESO Geocoded'] = "Midtown Alliance"
        df_diff.loc[df_diff['ESO Geocoded'] == "ASAP", 'ESO Geocoded'] = "Atlantic Station (ASAP)"

        # Change Null values of TMA to "Unknown/Out of Region"
        df_diff['TMA'] = df_diff['TMA'].fillna("Unknown/Out of Region")

        # Keep only rows where the contents of ESO Geocoded is different from what was in TMA
        df_diff = df_diff.loc[df_diff['ESO Geocoded'] != df_diff['TMA']]

        # Export!
        # df_diff.to_excel("c:\data\ESO Audit Test.xlsx", index=False)

        st.subheader("Processing Complete")


        # Create a BytesIO buffer for the ZIP file
        zip_buffer = io.BytesIO()

        with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED) as zip_file:
            # Create Excel files in memory and add them to the ZIP
            for df, name in zip([df_tableau, df_gdot, df_tdm, df_diff], ["Tableau.xlsx", "GDOT Report.xlsx", "TDM.xlsx", "ESO Audit.xlsx"]):
                excel_buffer = io.BytesIO()
                with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
                    df.to_excel(writer, index=False)
                zip_file.writestr(name, excel_buffer.getvalue())

        zip_buffer.seek(0)

        # Download button for the ZIP
        st.download_button(
            label="📦 Download All Excel Files",
            data=zip_buffer,
            file_name="excel_files_bundle.zip",
            mime="application/zip"
        )


        # Diagnostic if needed
        # print(df_tableau.columns.tolist())




