# Writing Speed Reporter 
### 1. Overview
This app’s purpose is to track and report on a user's writing speed. It uses command line UI for data collection and generates PDF reports including visualizations of CPM (characters per minute).
### 2. Features
* Start/Stop - after starting the session the data collection begins until it is stopped. Upon stopping data from the session is saved in .csv file and .pdf report for the session is generated
* Saved - connects selected data from previous session to be visualized on one report
* Options - allows for changing upper and lower threshold of CPM considered to weed out moments where nothing is written in data analysis, deciding if app should include all of the keys or only those which are used for writing (exclude backspace, alt, etc.)
* Tags - tags can be added or deleted in options and used to denote the type of writing session (e.g. lecture, coding, etc.). Later they can be compared in multi-session reports. Tags are kept in between uses of the app.

![writing1](https://github.com/user-attachments/assets/f79ff975-df18-4246-9f84-3ef7bc7b2b60)

#### 2.1. Data visualization
As of now the report includes two or three graphs: barplot showing the relationship between mean CPM and process used at the time of writing;

![barplot](https://github.com/user-attachments/assets/ba8797d6-3d62-4ae6-b89d-cbb1c75ce9f5)

lineplot showing CPM over time – in case of single-session report over minutes and hours, and while using multi-session data whole days are considered. 

![lineplot](https://github.com/user-attachments/assets/9deef5fb-2b61-4524-83f5-a6854410c713)

The third graph is generated only in a multi-session report and highlights the relationship between CPM and the tag used during a particular tracking session.

![barplot_2](https://github.com/user-attachments/assets/4cca90e2-5124-4923-8323-6f8c184ee814)

Report also includes date (of creating the report), peak speed, average speed and total characters written.

### 3. Running the project 
1. Clone repository locally
2. `cd` to the project directory
3. Install the dependencies if needed `pip install -r req.txt`
4. Run the project in commandline `py -m main`
5. In-app navigation - keyboard arrows to move around, enter to select/confirm, esc to go back

#### 3.1 Dependencies
* [py_cui](https://jwlodek.github.io/py_cui-docs/)
* [Seaborn](https://seaborn.pydata.org/)
* [ReportLab](https://www.reportlab.com/)
* [pynput](https://pynput.readthedocs.io/en/latest/)
* [NumPy](https://numpy.org/)
* [pandas](https://pandas.pydata.org/)
* [psutil](https://psutil.readthedocs.io/en/latest/)



