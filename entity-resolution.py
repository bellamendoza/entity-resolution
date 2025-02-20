import streamlit as st
import pandas as pd
from io import BytesIO
from Levenshtein import ratio

# find top 3 closest matches between two lists with Levenshtein similarity scores
def find_top_matches_with_scores(source, target):
    matches = []
    scores = []
    for item in source:
        scored_candidates = [(candidate, ratio(item, candidate)) for candidate in target]
        scored_candidates = sorted(scored_candidates, key=lambda x: x[1], reverse=True)[:3]  # Top 3 matches

        top_matches = [scored_candidates[i][0] if i < len(scored_candidates) else None for i in range(3)]
        top_scores = [round(scored_candidates[i][1] * 100, 2) if i < len(scored_candidates) else None for i in range(3)]

        matches.append(top_matches)
        scores.append(top_scores)

    return matches, scores

# generate Excel file
def generate_excel(data):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        data.to_excel(writer, index=False, sheet_name='Sheet1')
    return output.getvalue()


def main():
    st.title("Entity Resolution Tool")

    # File upload section
    st.header("Upload Excel Files")
    file1 = st.file_uploader("Upload Original", type="xlsx")
    file2 = st.file_uploader("Upload Lookup Table", type="xlsx")

    if file1 and file2:
        # Read the uploaded Excel files into DataFrames
        df1 = pd.read_excel(file1)
        df2 = pd.read_excel(file2)

        # Select column for matching
        column_a = st.selectbox("Select column from File A", df1.columns, key="column_a")
        column_b = st.selectbox("Select column from File B", df2.columns, key="column_b")

        if st.button("Perform Entity Resolution"):
            # Perform closest match computation
            source = df1[column_a].astype(str).tolist()
            target = df2[column_b].astype(str).tolist()

            # Temporarily store matches and scores to avoid recomputation
            if 'matches_with_scores' not in st.session_state:
                matches, scores = find_top_matches_with_scores(source, target)
                st.session_state.matches_with_scores = (matches, scores)

            matches, scores = st.session_state.matches_with_scores

            # Add matches and scores to the DataFrame
            output_df = pd.DataFrame(data={'Original':df1[column_a]})
            output_df['1st Closest Match'] = [m[0] for m in matches]
            output_df['1st Similarity Score (%)'] = [s[0] for s in scores]
            output_df['2nd Closest Match'] = [m[1] for m in matches]
            output_df['2nd Similarity Score (%)'] = [s[1] for s in scores]
            output_df['3rd Closest Match'] = [m[2] for m in matches]
            output_df['3rd Similarity Score (%)'] = [s[2] for s in scores]

            st.write("### Results of Entity Resolution")
            st.dataframe(output_df)

            # Button to download the result
            excel_data = generate_excel(output_df)
            st.download_button(
                label="Download Results",
                data=excel_data,
                file_name="entity_resolution_results.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

if __name__ == "__main__":
    main()
