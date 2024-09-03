import re
import numpy as np
import pandas as pd

from titlecase import titlecase

MAX_NUM_ORGANIZERS = 9
code2day = {
    "Th": "Thursday, August 29",
    "Fr": "Friday, August 30",
    "Sa": "Saturday, August 31" 
}

columns_to_save = ["Session title", "Session time", "Session location", "SchCode"] + [f"Session organizer {i}" for i in range(1,10)]

special_session = pd.read_excel("../data/SpecialSession.xlsx")\
                    .filter(["Title", "Abstract (initial)"])

special_session_data = pd.read_excel("../data/Papers_2.xlsx")\
                         .query('(`Session type` == "Special Session") & (`Type` == "Special Session Abstract")')\
                         .drop_duplicates("Session title")\
                         .filter(columns_to_save)\
                         .sort_values("Session title")

with open("special-sessions.tex", "w") as f:
    for _, row in special_session.iterrows():
        title = row["Title"]
        abstract = row["Abstract (initial)"]
        parts = special_session_data.query('`Session title`.str.contains(@title)')
    
        # Some special sessions got no papers
        if parts.empty: 
            continue
        
        # Title 
        f.write(f"\\section{{{titlecase(title)}}}\n\n")

        # # Session parts
        # for j, (_, part) in enumerate(parts.iterrows()):
        #     part_num = part["Session title"].split()[-1]
        #     time = part["Session time"]
        #     location = part["Session location"]
        #     day_code = part["SchCode"][:2]
        #     # day = code2day[part["SchCode"]] 

        #     f.write(f"{part_num}: \\textit{{{code2day[day_code]}, {time}, {location}}}")

        #     if not j == parts.shape[0] - 1:
        #         f.write(" \\\\ \n")
        #     else: 
        #         f.write("\n")
     
        # Collect organizers
        f.write("\n\\large \\textbf{Organizers} \\normalsize \\vspace{2mm} \\\\\n")
        organizers = []
        for i in range(1,MAX_NUM_ORGANIZERS+1):
            organizer = parts.iloc[0][f"Session organizer {i}"]
            if organizer is not np.nan:
                organizers.append(organizer)

        for j,organizer in enumerate(organizers):
            surname, name, affiliation = re.split(r", |\(", organizer.strip(")"))
            f.write(f"\\textbf{{{name.title()} {surname.title()}}} \\\\ \n\\textit{{{titlecase(affiliation)}}}")
            
            if not j == len(organizers) - 1:
                f.write(" \\vspace{{2mm}} \\\\\n")
            else: 
                f.write("\n\n")

        f.write(f"{abstract.replace("&", r"\&")} \n\n")

    f.write("\\vfill\\null")
