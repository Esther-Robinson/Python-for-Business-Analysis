
import pandas as pd
import numpy as np


df_raw = pd.read_excel('your_file.xlsx', header=None)


activity_names_row = df_raw.iloc[0] # Row 0 contains activity names. My activities are in the first row, but they are not the column headers.
#They are above the percentage/amount columns. I need to map them to the correct columns.
column_headers_row = df_raw.iloc[1] # Row 1 contains column headers (first_name, last_name, percentage, amount_1,.) 
#this are my column headers, but they are not unique. I need to make them unique by appending a number to the end of duplicate column names. they could have typos.


df = df_raw.iloc[2:].copy()  # Skip first two rows  Keep everything from Row 2 onwards, which contains the actual data. 
#We will set the column headers in the next step.

df.columns = column_headers_row # now set the column headers to the second row of the original data. 
#This will give us the correct column names, but they may not be unique. We will handle duplicates in the next step.

# Handle duplicate column names by making them unique. the file has excel column names that are not unique.
#  For example, there may be multiple "percentage" columns and multiple "amount" columns. 
# We need to make them unique by appending a number to the end of duplicate column names. For example, if there are two "percentage" columns,
#  we can rename them to "percentage_1" and "percentage_2". This will allow us to easily identify which percentage column corresponds to which amount column.
#  We will also do this for any other duplicate column names, such as "first_name" or "last_name" if they exist. 
# We will use a simple loop to check for duplicates and rename them
cols = pd.Series(df.columns) #pd.Series is like a list but with more functionality. It allows us to easily identify duplicates and modify them.
#doing pd.series behaves like a fancy list where you can filter, modify individual items. search easily and use boolean indexing. It also allows us to easily identify duplicates and modify them.
for dup in cols[cols.duplicated()].unique():  #cols.duplicated() returns a boolean series where True indicates a duplicate value. cols[cols.duplicated()] filters the columns to only those that are duplicates. .unique() gives us the unique duplicate column names. 
#Basically saying that 0 is my first time seeing "percentage", 1 is my second time seeing "percentage", etc. So I can append _1, _2, etc. to make them unique.
#putting cols[cols.duplicates ()] us saying this filters only the columns that are duplicates. Then .unique() gives us the unique duplicate column names. So if we have "percentage" duplicated 3 times, we will get "percentage" as a unique duplicate column name. Then we can loop through the indices of these duplicates and rename them accordingly.
#for dup in is saying the for loop will iterate through each unique duplicate column name.
#  So if we have "percentage" duplicated 3 times, we will get "percentage" as a unique duplicate column name. 
# Then we can loop through the indices of these duplicates and rename them accordingly. for dup in ['percentage_1', 'amount_1']:
    dup_indices = cols[cols == dup].index.values
# dup_indices will give us the indices of the columns that are duplicates. 
# For example, if "percentage" is duplicated in columns 3, 5, and 7, then dup_indices will be [3, 5, 7]. 
# We can then loop through these indices and rename the columns accordingly. this checks every position cols == 'percentage_1' and returns a boolean series where True indicates a match. Then .index.values gives us the indices of these matches. cols[cols == dup] is saying that we are filtering the columns to only those that match the current duplicate column name. Then .index.values gives us the indices of these matches. So if we have "percentage" duplicated in columns 3, 5, and 7, then dup_indices will be [3, 5, 7]. We can then loop through these indices and rename the columns accordingly.
#Index.values this extracts the positions where that name appears in the columns. So if "percentage" appears in columns 3, 5, and 7, then dup_indices will be [3, 5, 7]. We can then loop through these indices and rename the columns accordingly.
    for i, idx in enumerate(dup_indices[1:], 1):  # rename only the duplicates, not the first occurrence. So we start from dup_indices[1:] to skip the first occurrence. We also start the enumeration from 1 so that the first duplicate gets _1, the second duplicate gets _2, etc.
        cols[idx] = f"{dup}_dup{i}" # actually rename so it becomes like cols[2] = "percentage_dup1", cols[4] = "percentage_dup2", etc. This will make all duplicate column names unique by appending _dup1, _dup2, etc. to the end of the duplicate column name. So if we have "percentage" duplicated in columns 3, 5, and 7, then after this loop we will have "percentage" in column 3, "percentage_dup1" in column 5, and "percentage_dup2" in column 7. This will allow us to easily identify which percentage column corresponds to which amount column.
df.columns = cols # this replaces the old column names with the new column names that we just created. Now all duplicate column names are unique, which will make it easier to work with the data. We can now easily identify which percentage column corresponds to which amount column based on their names. For example, if we have "percentage" in column 3 and "amount" in column 4, we can assume that they correspond to the same activity. If we have "percentage_dup1" in column 5 and "amount_dup1" in column 6, we can assume that they correspond to the same activity as well. This will allow us to easily map the percentage and amount columns to their corresponding activities in the next steps of our analysis.

df.reset_index(drop=True, inplace=True) # this resets the index of the dataframe to be a simple range from 0 to n-1, where n is the number of rows in the dataframe. This is useful because we have skipped the first two rows of the original data, so the index may not be sequential. By resetting the index, we can ensure that it is clean and easy to work with. 
#The drop=True argument tells pandas to drop the old index instead of adding it as a new column. The inplace=True argument tells pandas to modify the dataframe in place instead of returning a new dataframe.

print("Column headers detected:")
print("=" * 80) # separator line for better readability
print(df.columns.tolist()) #it will show you the actual column names that we have in our dataframe after processing. 
#This is important to check because we had to handle duplicate column names and make them unique. 
# #By printing the column headers, we can verify that they are now unique and correctly reflect the structure of the data.
# # This will also help us in the next steps when we need to map the percentage and amount columns to their corresponding activities. 
# #If there are any issues with the column names, we can identify them at this stage and fix them before proceeding with the analysis.
#.tolist is used to convert the pandas Index object (which is what df.columns returns) into a regular Python list. 
# #This makes it easier to read and work with the column names.
print("=" * 80)

# create mapping of each column to its activity
# The activity name is above each percentage/amount pair
# row 0 is saring basically activity | black| another activity |black| another activity | black | etc.
#row 1 is the column headers like first_name, last_name, percentage, amount, percentage, amount, etc.
#row 2 + data is ana lopez | manager | 10% | 1000 | 20% | 2000 | etc.
#the issue is that activity name is not reapted above eacy column, it often it appears like activity and the the next columns uare blank until the next activity appears. So we need to map each percentage/amount column to the correct activity by looking at the row above and filling in the blanks with the last non-empty activity name. This way we can create a mapping of each percentage/amount column to its corresponding activity, which will be crucial for our analysis later on when we want to unpivot the data and analyze it by activity.
#so we need to map it like this {"percentage_1": "saring", "amount_1": "saring", "percentage_2": "another activity", "amount_2": "another activity", etc.}

activity_mapping = {} # create a empty dictionary 
# a dictionary is a data structure that stores key-value pairs. In this case, the key will be the column name (e.g., "percentage_1") and the value will be the corresponding activity name (e.g., "saring").

for i, col_name in enumerate(column_headers_row):
#now lopp through every column header, with its position.  Enumerate gives us both the index (i) and the column name (col_name) for each column in the column_headers_row. This will allow us to check if the column is a percentage or amount column, and then look up the corresponding activity name from the activity_names_row using the same index (i).
#i = the index position (0,1,2,3...) col_name the value at that position (first_name, last_name, percentage, amount, etc.)Example is i=0, col_name = "first_name", i=1, col_name = "last_name", i=2, col_name = "percentage", i=3, col_name = "amount", etc.
#This is crutial because we will use i to look up the activity name in the row above (activity_names_row) which has the same index positions. So if we are looking at column index 2 which is "percentage", we can look at activity_names_row[2] to get the activity name that corresponds to that percentage column. This way we can create a mapping of each percentage/amount column to its corresponding activity.
    # Get the activity name from the row above
    activity_name = activity_names_row.iloc[i]
#get the activity name from row 0 at the same column position activity_names_row is row 0 of the raw data, which contains the activity names. By using .iloc[i], we are getting the value at the same column index (i) as the current column header we are processing. This will allow us to map the percentage/amount columns to their corresponding activities based on their position in the original data. 
#.iloc[i] means give me the cell at position i so if you are at column header percentage_1 say its position 2 you grab whatever is above it in row 0 at position 2.  sometimes its a real activity and sometimes is blank / nan 
    
    # Only map percentage and amount columns
    if pd.notna(col_name) and ('percentage' in str(col_name).lower() or 'amount' in str(col_name).lower()):
# This is a filter. you do not want to map things like first_name or last_name or role and annual salary columns to activities. You only want to map the percentage and amount columns to activities. So this condition checks if the column name is not null and contains either "percentage" or "amount" in its name (case-insensitive). If this condition is true, then we proceed to map this column to its corresponding activity. If the column name does not meet this condition, we skip it and do not include it in the activity mapping.
# pd.notna(col_name) says "col_name is not missing (not NAN)". This ensures that we only consider columns that have a valid name. The second part of the condition checks if the column name contains "percentage" or "amount", which are the types of columns we want to map to activities. By combining these conditions, we ensure that we are only mapping relevant columns to their corresponding activities.
#str(col_name).lower() converts it into text makes it lowercase 
        # If activity name is NaN or empty, use previous non-empty activity name
        if pd.isna(activity_name) or str(activity_name).strip() == '':
            # Look backwards to find the activity name 
# This says is the activity name above is missing, search backward for the last valid activity. This checks tow kinds of missing values. pd.isna(activity_name) true if its NAN and str(activity_name).strip() == '' checks if its an empty string (after stripping any whitespace). If either of these conditions is true, it means we do not have a valid activity name at this position. In that case, we need to look backwards in the activity_names_row to find the last non-empty activity name. This is because in the original data, the activity name may only be listed once and then the next few columns may be blank until the next activity name appears. By looking backwards, we can fill in those blanks with the correct activity name for each percentage/amount column.

            for j in range(i-1, -1, -1):
# search lef/backwards to find the nearest activity label  this is a backwards loop starts at i-1 (the column to the left of the current column) and goes all the way to 0 (the first column). The step -1 means we are moving backwards through the columns. This loop will check each column to the left of the current column until it finds a non-empty activity name. Once it finds a valid activity name, it will break out of the loop and use that activity name for the current percentage/amount column.
                prev_activity = activity_names_row.iloc[j]
# get the previous activity name at position j. This is the same as what we did for the current column, but now we are checking the columns to the left (previous columns) to find a valid activity name. We will check if this prev_activity is valid (not missing and not empty) and if it is, we will use it as the activity name for the current column.
                if pd.notna(prev_activity) and str(prev_activity).strip() != '':
                    activity_name = prev_activity
                    break
# if the previous activity is not blank use it and stop searching. This checks that the activity name is real. 
# if it is then assign it as the activity for the current column, break stops the loop immediately (We have found the nearest valid activity name to the left, so we can stop searching further.)   
# If we do not find any valid activity name to the left, then activity_name will remain as it is (which may be NaN or empty), and we will still map the current column to that value. This means that if there are columns at the beginning of the row that do not have an activity name above them, they will be mapped to NaN or empty, which is fine because those columns are likely not percentage/amount columns anyway (e.g., first_name, last_name).    
        activity_mapping[col_name] = str(activity_name).strip()  #this saves the mapping of the current column name to its corresponding activity name in the activity_mapping dictionary. The key is the column name (e.g., "percentage_1") and the value is the activity name (e.g., "saring"). We also use .strip() to remove any leading or trailing whitespace from the activity name, just in case there are any extra spaces in the original data.

print("\nActivity Mapping:")
print("=" * 80)
for col, activity in activity_mapping.items():
    print(f"{col} -> {activity}")
print("=" * 80) # print show me each mapping items

# Identify static columns
static_cols = [] # Create an empty container (python list) to hold the names of the static columns. Static columns are those that contain information that does not change across the different activities, such as first name, last name, role, annual salary, etc. We want to identify these columns so that we can keep them as is when we unpivot the data later on. By checking for keywords like "first_name", "last_name", "role", "annual", and "sum_of" in the column names, we can determine which columns are likely to be static and should be included in this list.
for col in df.columns: #lloop through each column in the dataframe and check if it is a static column based on its name. We will convert the column name to lowercase and check if it contains any of the keywords that indicate it is a static column. If it does, we will add it to the static_cols list. This way we can easily identify which columns are static and should be kept as is when we unpivot the data later on.
    col_str = str(col).lower() 
    if 'first_name' in col_str or 'last_name' in col_str or 'role' in col_str or 'annual' in col_str or 'sum_of' in col_str: # if columns contains any of these words, treat as static. 
        static_cols.append(col) 

print(f"\nStatic columns: {static_cols}")
print(f"Data shape: {df.shape}")
print(f"\nFirst few rows:")
print(df.head(3))

# Unpivot the data this makes my table to go from wide format to long format. (one row per activity per person) This is called unpivoting or melting conceptually.  
unpivoted_data = [] # This creates an empty python list a list because we dont know how many rows we will wnd up with after unpivoting. We will append each new row of unpivoted data to this list as we process the original dataframe. Each item in this list will be a dictionary representing a single row of the unpivoted data, with keys corresponding to column names and values corresponding to the data for that row. After we have processed all the rows in the original dataframe, we will convert this list of dictionaries into a new dataframe that contains the unpivoted data.

for idx, row in df.iterrows(): # This loop thorugh each person each row. df.iterrows() doe it returns one at a time. idx = the row number and row is a pandas Series object that contains the data for that row. 
# example for ana's row behaves like a dictionary row['first_name'] = "ana" row['last_name'] = "lopez" row['role'] = "manager" row['percentage_1'] = 10% row['amount_1'] = 1000 row['
# this outer loop means for each person go find their percentage amount pairs and covert them into separate rows. 
    # Find all percentage and amount column pairs
    processed_pairs = set() # create a set to remeber which columns you already processed as pairs. This is important because we want to make sure that we are correctly matching each percentage column with its corresponding amount column, and we do not want to accidentally match a percentage column with the wrong amount column. By keeping track of which columns we have already processed as pairs, we can ensure that we are creating accurate rows of unpivoted data for each activity.
    
    for col in df.columns: # Loop thorugh evey column name inside that row. now you are scanning the columns one by one to find the activity columns  " for this person, inspect every column and look for percentage columsn / amount" 
        col_str = str(col).lower() # covert column name to lowercase string this standardizes it for matching 
        
        # Look for percentage columns
        if 'percentage' in col_str and col not in processed_pairs: # find percentages columns that havent been proceed yet.  
#col not in processed_pairs is saying if this column has not already been processed as part of a percentage/amount pair, then we will consider it for processing. This is important because we want to avoid processing the same column multiple times and creating duplicate rows of unpivoted data. By checking if the column is not in the processed_pairs set, we can ensure that we only process each percentage/amount pair once for each person.

            # Find corresponding amount column
            # Extract the number from percentage column if it exists
            percentage_col = col # this is the percentage column we are currently processing. We will use this to find the corresponding amount column. For example, if percentage_col is "percentage_1", we will look for an amount column that has the same number (e.g., "amount_1") to pair it with. This way we can ensure that we are matching the correct percentage column with its corresponding amount column for the same activity.
            
            amount_col = None # I havent found the matching amount column yet so I set it to None for now. We will try to find the correct amount column that corresponds to this percentage column. If we find it, we will update this variable with the name of the amount column. If we do not find it, it will remain as None and we will skip this percentage column because we cannot create a valid row of unpivoted data without both a percentage and an amount.
            
            if '_' in str(col): #if the column name contains an underscore, we will try to extract the number from the percentage column and look for an amount column with the same number. This is a common pattern in the data where percentage and amount columns are named with a suffix that indicates their pairing (e.g., "percentage_1" pairs with "amount_1"). By extracting the number from the percentage column, we can directly look for the corresponding amount column that has the same number, which increases our chances of correctly matching them.
                num = str(col).split('_')[-1] # extract the suffix ( the number after underscore) from the percentage column name. For example, if col is "percentage_1", then num will be "1". We can then use this num to look for an amount column that ends with the same number (e.g., "amount_1"). This is a straightforward way to find the corresponding amount column based on the naming convention in the data.
                for amt_col in df.columns: # search all cloumns to find an amount column ending with the same number "Look though all columns, If you find one that conatins 'amount' and ends with the same number, thats the match" then stop because we found it.  
                    if 'amount' in str(amt_col).lower() and str(amt_col).endswith(num):
                        amount_col = amt_col
                        break
            
#if we still haven't found amount column, use fallcak approach to look for the nearest amount column to the right of the percentage column. This is a backup method in case the naming convention is not consistent and we cannot find a direct match based on the suffix number. By looking to the right of the percentage column, we can try to find an amount column that is likely to be paired with it based on its position in the original data. This is not as reliable as the suffix matching, but it can help us find pairs that do not follow the naming convention.
            if amount_col is None: 
                col_list = df.columns.tolist() # turn columns into a list and get index position .
                col_index = col_list.index(col) # Because we want to search to the right of the current percentage column. col_list is normal python list. col_index gets the position od current percentage column. 
                for i in range(col_index + 1, len(col_list)): # llok to the right for the next amount column. 
                    if 'amount' in str(col_list[i]).lower(): # Staring from the column right after percentage column, scan forward until you find an amount. 
                        amount_col = col_list[i]
                        break
            
            if amount_col is None: # If no amount column found, skip this percentage column ( Skip this loop iteration and move to the next column) because we cannot create a valid row of unpivoted data without both a percentage and an amount. We need both pieces of information to accurately represent the activity for this person. If we only have a percentage or only have an amount, it would not make sense to include it in the unpivoted data, so we will skip it and move on to the next column.
                continue
# Add both columns to processed set. now the set loooks like percentage_1, amount_1. so later when the inner reaches amount_1 or the same percentage it wont reprocess it because we already know they are a pair. This ensures that we do not accidentally match the same percentage column with multiple amount columns or vice versa. By adding both the percentage and amount columns to the processed_pairs set, we can keep track of which columns have already been paired together and avoid creating duplicate rows of unpivoted data for the same activity.           
            processed_pairs.add(percentage_col) # grab the percentage value
            processed_pairs.add(amount_col) # grab the amount value 
            
            percentage_val = row[percentage_col]
            amount_val = row[amount_col]
            
            # If it's a Series, take the first value This is a safeguard in case the cell contains multiple values (e.g., due to merged cells in Excel) and is read as a Series. By taking the first value, we can still extract a single percentage and amount for that activity. This is not ideal, but it allows us to handle cases where the data may not be perfectly clean and still include as much information as possible in the unpivoted data.
            if isinstance(percentage_val, pd.Series):
                percentage_val = percentage_val.iloc[0]
            if isinstance(amount_val, pd.Series):
                amount_val = amount_val.iloc[0]
            
            # Skip if both are null " if both percentage and amount are blank dont create a long format row. This is important because if we have a percentage column and an amount column that are both empty for a particular activity, it means that this person did not contribute to that activity, and we do not need to include a row for it in the unpivoted data. By skipping rows where both percentage and amount are null, we can keep our unpivoted data cleaner and more focused on the activities that people actually contributed to."
            if pd.isna(percentage_val) and pd.isna(amount_val):
                continue
            
            # Skip if both are zero or empty use try/except because some cells might contain text, blanks and weird formatting this avoid the program to crash. 
            try:
                if float(percentage_val) == 0 and float(amount_val) == 0:
                    continue
            except (ValueError, TypeError):
                pass
            
            # Get activity name from mapping based on percentage column name. We use the percentage column name to look up the corresponding activity name from the activity_mapping dictionary that we created earlier. This way we can ensure that we are correctly associating each percentage/amount pair with the right activity based on the original structure of the data. If for some reason we cannot find a mapping for this percentage column, we will default to "Unknown Activity" to indicate that we were not able to determine the activity for this pair.
            activity_name = activity_mapping.get(percentage_col, 'Unknown Activity')
            
            # Get static column values
            new_row = {} 
            for static_col in static_cols: #for each static column copy the person info into this new row. This ensures that when we create a new row of unpivoted data for a particular activity, we also include all the relevant static information about the person (e.g., first name, last name, role, annual salary) in that same row. This way we can maintain the context of who contributed to each activity and have all the necessary information in one place for analysis.
                val = row[static_col]
                # If it's a Series, take the first value
                if isinstance(val, pd.Series):
                    val = val.iloc[0]
                new_row[static_col] = val
            
            new_row['activity_name'] = activity_name 
            new_row['percentage'] = percentage_val
            new_row['amount'] = amount_val
            
            unpivoted_data.append(new_row) # save the new row into the list. This adds the new row of unpivoted data to our list of unpivoted_data. Each item in this list is a dictionary that represents a single row of the unpivoted data, with keys corresponding to column names (including static columns, activity_name, percentage, and amount) and values corresponding to the data for that row. After we have processed all the rows in the original dataframe and created these dictionaries for each activity contribution, we will convert this list of dictionaries into a new dataframe that contains all the unpivoted data for analysis.

# Create the unpivoted dataframe
df_long = pd.DataFrame(unpivoted_data) # COnvert the list of dictionaries into a new dataframe. Each dictionary in the unpivoted_data list becomes a row in the new dataframe, and the keys of the dictionaries become the column names. This new dataframe (df_long) will be in long format, where each row represents a single activity contribution for a person, along with all the relevant static information about that person. We can then perform further cleaning and analysis on this unpivoted dataframe as needed.

print(f"\n{'='*80}")
print(f"Unpivoted {len(unpivoted_data)} rows")
print(f"{'='*80}")

if len(df_long) == 0: # if it has 0 rows something went wrong.
    print("\nERROR: No data was unpivoted!")
    print("Please check the structure of your file.")
else:
    # Clean column names for easier access
    df_long.columns = [str(col).strip().lower().replace(' ', '_') for col in df_long.columns] #otherwise give me that data and ensure the column names are clean and standardized. This will make it easier to work with the columns in the unpivoted dataframe. By stripping whitespace, converting to lowercase, and replacing spaces with underscores, we can create consistent column names that are easier to reference in our analysis. For example, if we had a column named "First Name", it would be converted to "first_name", which is more convenient for coding and analysis purposes.
    
    # Sort the data
    if 'last_name' in df_long.columns and 'first_name' in df_long.columns:
        df_long = df_long.sort_values(['last_name', 'first_name', 'activity_name'])
    df_long.reset_index(drop=True, inplace=True) # data neatly group per person and sorted by activity.   

    
    # Convert numeric columns
    if 'percentage' in df_long.columns:
        df_long['percentage'] = pd.to_numeric(df_long['percentage'], errors='coerce') #errors='coerce' means if a velue cannot be converted to a number pandas turns it into a nan instead of crashing.     
    if 'amount' in df_long.columns:
        df_long['amount'] = pd.to_numeric(df_long['amount'], errors='coerce') # turning our numbers into numbers so we can do math on them. This will convert the percentage and amount columns to numeric data types, which will allow us to perform calculations and analysis on these columns later on. The errors='coerce' argument means that if there are any values in these columns that cannot be converted to numbers (e.g., due to text or formatting issues), they will be set to NaN instead of causing an error. This helps to ensure that our data is clean and usable for analysis, even if there are some irregularities in the original data.
    
    # Save the cleaned data
    #df_long.to_excel('cleaned_activity_data.xlsx', index=False)
    #df_long.to_csv('cleaned_activity_data.csv', index=False)
    
    print("\n" + "=" * 80)
    print("FIRST 20 ROWS OF CLEANED DATA:")
    print("=" * 80)
    print(df_long.head(20).to_string())
    
    print("\n" + "=" * 80)
    print(f"UNIQUE ACTIVITIES ({df_long['activity_name'].nunique()} total):")
    print("=" * 80)
    for i, activity in enumerate(sorted(df_long['activity_name'].unique()), 1):
        print(f"{i}. {activity}")
    
    print("\n" + "=" * 80)
    print("TOTAL CONTRIBUTION BY ACTIVITY:")
    print("=" * 80)
    activity_summary = df_long.groupby('activity_name').agg({
        'amount': ['sum', 'count']
    }).round(2)
    activity_summary.columns = ['Total_Amount', 'Num_People']
    activity_summary = activity_summary.sort_values('Total_Amount', ascending=False)
    print(activity_summary.to_string())
    
    print("\n" + "=" * 80)
    print("Files saved successfully!")
    print("  - cleaned_activity_data.xlsx")
    print("  - cleaned_activity_data.csv")
    print("=" * 80)

 #df_long.to_excel('cleaned_activity_data.xlsx', index=False)