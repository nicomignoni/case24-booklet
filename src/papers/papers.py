import pandas as pd

code2day = {
    "Th": "Thursday, August 29",
    "Fr": "Friday, August 30",
    "Sa": "Saturday, August 31" 
}

code2date = {
    "Th": "29-08-2024",
    "Fr": "30-08-2024",
    "Sa": "31-08-2024" 
}


def extract_day(code):
    return code if pd.isna(code) else code2day[code[:2]]

def extract_date(code):
    return code if pd.isna(code) else code2date[code[:2]]

data = pd.read_excel("../data/Papers.xlsx")\
          .dropna(subset="SchCode")\
          .query("(`Session type` == 'Regular Session') | (`Session type` == 'Special Session')")
data["Session day"] = data["SchCode"].apply(extract_day, 1)
data["Session date"] = data["SchCode"].apply(extract_date, 1)

groups_criteria = ["Session date", "Session time", "Session location", "Session title", "Paper time"]
data.sort_values(groups_criteria, inplace=True)

with open("papers.tex", "w") as file:
    file.write("\\begin{multicols*}{2}\n")

    for day in data["Session day"].unique():
        file.write(f"\\subsection*{{{day}}}\n\n")
        sessions = data.query(f"`Session day` == '{day}'")

        for session_title in sessions["Session title"].unique():
            papers = sessions.query(f"`Session title` == '{session_title.replace('&', 'and')}'")
            session_time, session_room = papers.iloc[0][["Session time", "Session location"]]
            file.write(f"\\normalsize \\textbf{{{session_title}}}\\\\\n\\small \\textit{{{session_room} @ {session_time}}}\n\n")

            for _, paper in papers.iterrows():
                authors = ", ".join(f"{' '.join(name.split(', ')[::-1])}" for name in paper['Authors'].split(";"))
                file.write(f"\\small {paper['Title']}\\\\ \n\\footnotesize \\textcolor{{darkgray}}{{\\textit{{{authors.title()}}}}}\n\n")
    
    file.write("\\end{multicols*}")


