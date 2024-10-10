import streamlit as st
import pandas as pd
import numpy as np
from textblob import TextBlob
import matplotlib.pyplot as plt
from wordcloud import WordCloud
from io import BytesIO  # Import BytesIO

# Title for the App
st.title("Professor Review Sentiment Analysis")

# Data Collection
st.write("## Step 1: Data Collection")
st.write("Upload an Excel file containing professor reviews (or fetch data from RateMyProfessor, GitHub, etc.)")
uploaded_file = st.file_uploader("Upload an Excel file with professor reviews", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    
    # Display uploaded data
    st.subheader("Uploaded Data")
    st.dataframe(df)
    
    # Data Preprocessing
    st.write("## Step 2: Data Preprocessing")
    
    st.write("### Preprocessing Steps:")
    # Standardize responses for "take_again" column
    df['take again'] = df['take again'].str.capitalize()
    
    # Fill missing values
    df.fillna({'comments': 'No comment', 'difficulty': df['difficulty'].mean(), 'overall quality': df['overall quality'].mean()}, inplace=True)
    df['difficulty'] = df['difficulty'].astype(int)
    df['overall quality'] = df['overall quality'].astype(int)
    
    st.write("Data after preprocessing:")
    st.dataframe(df)

    # Sentiment Analysis
    st.write("## Step 3: Sentiment Analysis")
    
    # Function to analyze sentiment
    def analyze_sentiment(comment):
        analysis = TextBlob(comment)
        if analysis.sentiment.polarity > 0:
            return 'Positive'
        elif analysis.sentiment.polarity == 0:
            return 'Neutral'
        else:
            return 'Negative'

    # Apply sentiment analysis
    df['sentiment'] = df['comments'].apply(analyze_sentiment)

    st.write("Sentiment Analysis Results:")
    sentiment_summary = df['sentiment'].value_counts()
    st.write(sentiment_summary)

    # Visualization
    st.write("## Step 4: Visualization")

    # Average Ratings for Difficulty and Overall Quality
    st.write("### Average Ratings for Difficulty and Overall Quality")
    avg_ratings = df[['difficulty', 'overall quality']].mean()
    st.bar_chart(avg_ratings)

    # Pie chart for "Take Again" responses
    st.write("### Would Students Take the Course Again?")
    take_again_counts = df['take again'].value_counts()
    fig1, ax1 = plt.subplots()
    take_again_counts.plot(kind='pie', autopct='%1.1f%%', colors=['lightgreen', 'lightcoral'], ax=ax1)
    ax1.set_ylabel('')
    ax1.set_title('Take Again?')
    st.pyplot(fig1)

    # Sentiment Distribution
    st.write("### Sentiment Distribution")
    st.bar_chart(sentiment_summary)

    # Word cloud for comments
    st.write("### Word Cloud for Comments")
    wordcloud = WordCloud(width=800, height=400, background_color='white').generate(' '.join(df['comments']))
    fig2, ax2 = plt.subplots(figsize=(10, 5))
    ax2.imshow(wordcloud, interpolation='bilinear')
    ax2.axis('off')
    st.pyplot(fig2)

    # Automate Feedback Response
    st.write("## Step 5: Automated Feedback Response")

    # Function to generate feedback based on sentiment
    def generate_feedback(sentiment):
        if sentiment == 'Positive':
            return "Thank you for the positive feedback! We're glad you had a great experience."
        elif sentiment == 'Neutral':
            return "Thank you for your feedback! We will continue to work on improving the course."
        else:
            return "We're sorry you had a negative experience. Your feedback will help us improve."

    # Apply the feedback generation function
    df['automated_feedback'] = df['sentiment'].apply(generate_feedback)

    # Show feedback responses
    st.write("Generated Feedback Responses:")
    for idx, row in df.iterrows():
        st.write(f"**Comment:** {row['comments']}")
        st.write(f"**Automated Response:** {row['automated_feedback']}")
        st.write("---")

    # Saving Processed Data
    st.write("## Step 6: Download Processed Data")

    def convert_df_to_excel(df):
        # Use BytesIO to create a buffer
        output = BytesIO()
        # Write the DataFrame to the buffer
        df.to_excel(output, index=False, engine='openpyxl')
        # Move to the beginning of the stream
        output.seek(0)
        return output

    df_to_download = convert_df_to_excel(df)

    st.download_button(
        label="Download Processed Data as Excel",
        data=df_to_download,
        file_name='processed_professor_reviews.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
else:
    st.write("Please upload an Excel file to continue.")
