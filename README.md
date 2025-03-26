**1. Introduction**

**Project Description**

The purpose of this project is to randomly select phrases or attributes related to the real estate market, particularly about specific properties, using the MaxDiff method to randomly select phrases. MaxDiff is used here as a technique to determine which phrases or attributes are most relevant or preferred from a set of available options, providing a clearer view of preferences in contexts where multiple versions or descriptions are involved.

**Motivation**
My main motivation for creating this project is the constant search for learning. In my work environment, I often encounter situations where this type of activity is relevant, and I consider it essential to understand in a practical way how phrase randomization works and the concept underlying MaxDiff. This project seeks to solve problems related to random selections, offering a structured method to understand and apply preferences within the context of the real estate market. This practice not only improves theoretical understanding, but also provides insights into the practical application of these techniques in real-world scenarios.

**2. What is MaxDiff?**

**Definition**

MaxDiff, or Maximum Difference Scaling, is a statistical technique used to measure relative preferences between different items or attributes. This methodology is commonly used in market research to understand which attributes of a product or service are most valued by consumers. By presenting sets of options and asking participants to select the most and least preferred items, MaxDiff gathers data that reflects aggregate preferences more effectively than traditional methods.

**Project Application**
In this project, MaxDiff is applied to randomly select phrases or attributes related to specific properties. We use four main metrics to ensure that the selection process is robust and understandable:

**Total Number of Distinct Phrases:** Determines how many unique phrases or attributes will be included in the set of options analyzed. This metric ensures enough variety to provide a comprehensive assessment of preferences.

**Coverage (Phrase Repetition):** We define how many times each phrase should appear in the different sets (or screens). This ensures that all sentences have adequate coverage and are evaluated the same number of times, ensuring representativeness in the data collection.

**Number of Sentences per Set (Screen):** Specifies how many sentences will be shown simultaneously in each set presented to the interviewees. This number balances the cognitive complexity required for selection, without overloading the participants.

**Number of Interviewees:** Establishes how many individuals will participate in the study, contributing to the statistical robustness of the results. A larger number of interviewees can increase confidence in the conclusions drawn from the data.

These metrics form the basis for the effective implementation of MaxDiff in sentence selection, providing insights into which attributes are most valued in a real estate scenario. The rigorous methodological approach allows real estate companies to analyze preference data in an effective and informed manner.

**3. How MaxDiff Works**

**Algorithm Overview**

MaxDiff, in this project, is implemented to perform an optimized selection of phrases or attributes, ensuring that the respondents' preferences are well mapped and understood. Here is a detailed view of how the algorithm works and what its main parts are:

**Processes and Functions:**
**Calculating Optimized Sets**

- **Function: calculate_optimized_sets**
- This function calculates the minimum number of sets needed to present all phrases so that each one appears in a set a target number of times.

Parameters:
- **num_phrases:** The more different phrases we have, the higher the number of sets needed.
- **target_appearances:** How many times we want a phrase to appear, ensuring adequate coverage.
- **items_per_set:** The number of phrases that appear on each screen presented to the respondents.
- This approach ensures that variation and coverage are maintained at an acceptable level, distributing the sentences evenly.

**Generating MaxDiff Sets and Exporting**

- **Function: generate_excel_for_interviewees**
- Based on the calculated combinations, this function generates MaxDiff sets for each interviewee and exports the data to an Excel file, organizing the information in a clear and accessible way.
- **Process Details:**
- For each interviewee, sets of sentences are generated according to the previously calculated metrics.
- The sentences are randomly organized to avoid any bias.
- Various combinations ensure that all attributes are covered

**Validation of Input Data**

- **Function: validate_input**
- Before executing the draw, this function ensures that the number of sentences and the configuration of the sets are sufficient to cover the sample needs.
- **Included Checks:**
- Ensures that the number of sentences is not less than what is needed to fill the sets.
- Evaluates whether the configurations will allow the necessary repetition of the sentences.

**Validation Report**

- **Function: generate_validation_report**
- Generates a textual report that summarizes the results of the input validation, indicating whether there are any problems that need to be corrected before the algorithm can be fully executed.

**Expected Output**
- The final result is an Excel file containing the complete organization of the optimized MaxDiff sets, with clear summaries of how many sets were generated and the distribution of sentences per set.
- Other outputs include a validation report that helps contextualize whether the configured data was sufficient and correctly distributed. This general outline of how the MaxDiff algorithm works in the project provides a solid understanding of how the technique is applied to phrase sorting in the real estate context. The code is designed to maximize the effectiveness of respondents’ decisions, ensuring that preferences are reflected fairly and accurately.

**4. Phrase Sorting**

**Purpose of Phrase Sorting**
Phrase sorting is a crucial step in the project, designed to ensure that all phrases or attributes related to real estate are presented to respondents in a balanced manner. The idea is to explore respondents’ preferences without bias and in a comprehensive way, maximizing the efficiency of the data collection.

**Phrase Sorting Process**

**Preparation and Structuring**

First, phrases related to real estate attributes are listed and indexed in a dictionary, where each phrase is associated with a unique code. This makes it easier to track and identify phrases throughout the process.

**Sorting Algorithm**

- **Random Selection:** Using the np.random.choice function, the code randomly selects a specific number of sentences for each set (or screen) to be presented. This selection is done without replacement, ensuring that the sentences in the same set are distinct.

- **Number of Sets:** The total number of sets, calculated in the previous step, is generated for each interviewee so that all sentences are adequately covered, according to the target number of appearances established.

**Balanced Distribution**

The drawing is designed to ensure that all sentences appear in different combinations across the sets, minimizing repetition and maximizing the diversity presented to the interviewees. This balance is essential for the accuracy of the data collected, preventing certain words from being underrepresented.

**Data Creation and Storage**

Once the sentences are drawn for each set and interviewee, this information is organized in a DataFrame and exported to an Excel file. This allows the data to be easily analyzed and reviewed, as well as allowing a clear visualization of the presented sets.

**Benefits of the Random Selection**

- **Complete Coverage:** The use of a random selection ensures that all sentences have an equal opportunity to be selected, fulfilling the defined appearance requirements.

- **Bias Reduction:** The randomness in the selection provides a fairer evaluation field and, at the same time, deepens the behavioral analysis of the interviewees.

- **Ease of Implementation:** Automating the selection and export of data in organized files simplifies the process of data collection and analysis, allowing a greater focus on the interpretation of the results.

**5. About the Code**

**Python Code**

**Libraries Used:**

- **NumPy and Pandas:** Used to manipulate arrays and create DataFrames for organizing and exporting data.

- **OS:** For possible operations on the file system, although in the current example this is not explicitly used.

**Main Functions:**

- **calculate_optimized_sets:** Determines the minimum number of sets needed based on the total number of sentences, target number of appearances, and items per set.

- **generate_excel_for_interviewees:** Creates sets of sentences for each interviewee, randomly selecting them, and exports the results to an Excel file.

- **validate_input:** Checks whether the input conditions are valid, ensuring that the sentence selection can be performed without problems.

- **generate_validation_report:** Generates a text report with the validation of the inputs.

**Output:**

An Excel file with data from the MaxDiff sets and a summary containing statistics about the sets and sentences.

**Imported Libraries**

- **import numpy as np:** Imports the NumPy library, useful for fast and efficient operations on arrays and matrices.
- **import pandas as pd:** Imports the Pandas library, which facilitates the manipulation and analysis of tabular data through DataFrames.
- **import math:** Brings additional mathematical functions, such as rounding, exponentials, logarithms, etc.
- **import os:** Used for interactions with the operating system, such as file and directory manipulation.

**Specific Functions**

- **math.ceil(x):** Returns the smallest integer greater than or equal to x, used to ensure that the minimum number of sets is an integer value.
- **np.random.choice(arr, size, replace):** Randomly selects elements from an array.
- **arr:** Array from which to choose the elements.
- **size:** Number of elements to choose.
- **replace=False:** Ensures that elements will not be repeated in the selection.

**Data Manipulation with Pandas**

- **pd.DataFrame(dict):** Creates a DataFrame from a dictionary, where keys represent columns and values ​​represent column data.

- **df.to_excel(writer, index=False, sheet_name='Name'):** Exports a DataFrame to an Excel file, allowing you to specify the name of the sheet.

**Other Commands**

- with open(file_name, 'w') as file: Opens a file for writing, ensuring that it is closed correctly after I/O operations.

**R Code**

**Libraries Used:**

- **openxlsx:** Used to create and manipulate Excel files, allowing data export in a similar way to Python.

**Main Functions:**

- **calculate_optimized_sets:** Similar to the Python version, calculates the required sets.

- **generate_excel_for_interviewees:** Similar to Python, this function sorts sentences and organizes their data for export to Excel.

- **validate_input:** Checks the adequacy of the inputs in terms of the sentences and set requirements.

- **generate_validation_report:** Creates a text report summarizing the validation of the input conditions.

**Imported Libraries**

- **library(openxlsx):** Loads the openxlsx package, essential for manipulating Excel spreadsheets.

**Specific Functions**

- **ceiling(x):** Same as the Python function math.ceil, returns the smallest integer greater than or equal to x.

- **sample(arr, size, replace=FALSE):** Similar to np.random.choice, select elements randomly from a vector:
- **arr:** Vector from which the elements will be chosen.
- **size:** Number of elements to select.
- **replace=FALSE:** Does not allow duplication of elements in the same draw.

**Data Manipulation**

- **data.frame(...):** Creates a DataFrame to organize tabular data, similar to the DataFrame in Pandas.
- **rbind(df1, df2):** Combines DataFrames along the lines, incrementing the existing data.

**Writing to Files**

- **writeLines(...):** Writes text to an output file, used to save validation reports.
- **cat(...):** Displays messages on the console or writes them to a file, similar to print in Python.

**openxlsx Excel Functions**

- **createWorkbook():** Creates a new Excel workbook.

- **addWorksheet(wb, "Name"):** Adds a new tab to the workbook.

- **writeData(wb, "SheetName", data):** Writes data to a specified sheet in the workbook.

- **saveWorkbook(wb, file, overwrite=TRUE):** Saves the specified workbook, potentially overwriting existing files.

**Notable Differences:**

- While Python uses the np.random.choice function for random generation, R uses sample for the same purpose.

- In R, data is organized into data frames and exported using the features provided by the openxlsx package.

**Why Both Implementations?**

Both implementations offer flexibility and allow users to choose their preferred language based on the needs of the team or project. Both Python and R have unique strengths in data analysis, and this project balances those advantages.
